const axios = require("axios");

//Smartsheetの設定
const ACCESS_TOKEN = process.env.ACCESS_TOKEN;
const SHEET_ID = process.env.SHEET_ID;

const headers = {
    "Authorization":`Bearer ${ACCESS_TOKEN}`,
    "Content-Type":"application/json"
};

async function checkUnenteredWork(){
    try{
        //1.シート全体を取得
        const sheetResp = await axios.get(
            `https://api.smartsheet.com/2.0/sheets/${SHEET_ID}`,
            {headers}
        );
        const sheet = sheetResp.data;

        const rows = sheet.rows;
        const cols = sheet.columns;

        //列タイトルを確認
        console.log("列タイトル一覧:", cols.map(c => c.title));

        //必要な列を特定
        const flagCol = cols.find(c => c.title === "未入力");
        const statusCol = cols.find(c => c.title === "ステータス");
        const startCol = cols.find(c => c.title === "開始日（実績）");
        const endCol = cols.find(c => c.title === "終了日（実績）");

        if(!flagCol || !statusCol || !startCol || !endCol){
            throw new Error("必要な列が見つかりません。列名を確認してください。");
        }

        //全日付列を抽出
        const dateCols = cols.filter(c => /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(c.title));

        //昨日の日付を作成
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        const yTitle = `${yesterday.getFullYear()}/${yesterday.getMonth()+1}/${yesterday.getDate()}`;
        const yCol = cols.find(c => c.title.trim().startsWith(yTitle));

        //確認用ログ
        console.log("昨日の日付タイトル:", yTitle);
        console.log("見つかった列:", yCol);
        
        const updateRows = [];

        //2.各行をチェック
        for(const row of rows){
            const status = row.cells.find(c => c.columnId === statusCol.id)?.value;
            const startDate = row.cells.find(c => c.columnId === startCol.id)?.value;
            const endDate = row.cells.find(c => c.columnId === endCol.id)?.value;

            let flag = false;

            //ステータスごとの処理
            if(status === "進行中" && yCol){
                //機能が空欄なら未入力フラグTRUE
                const yCell = row.cells.find(c => c.columnId === yCol.id);         
                if(!yCell || yCell.value === null || yCell.value === undefined || yCell.value === ""){
                    flag = true;
                }
            }

            if(status === "完了" && startDate && endDate){
                //開始～終了日の範囲の日付列をチェック
                const s = new Date(startDate);
                const e = new Date(endDate);

                for(const col of dateCols){
                    const[yy, mm, dd] = col.title.split("/").map(Number);
                    const d = new Date(yy, mm-1, dd);

                    if(d >= s && d <= e){
                        const cell = row.cells.find(c => c.columnId === col.id);
                        if(!cell || cell.value === null || cell.value === undefined || cell.value === ""){
                            flag = true;
                            break; //空白が見つかったら未入力フラグTRUE
                        }
                    }
                }
            }

            //更新リストに追加 ログ有バージョン
            const update = {
                id: row.id,
                cells:[{columnId: flagCol.idr, value: flag}]
            };

            console.log("更新内容:", update);

            updateRows.push(update);
            
            /*
            //更新リストに追加
            updateRows.push({
                id: row.id,
                cells:[{columnId: flagCol.id, value: flag}]
            });    */     
        }

        //3.更新を反映
        if(updateRows.length > 0){
            await axios.put(
                `https://api.smartsheet.com/2.0/sheets/${SHEET_ID}/rows`,
                updateRows,
                {headers}
            );
        }

        console.log(`✅ 未入力チェック完了: ${updateRows.length} 行を更新しました`);

    }catch(err){
        console.error("❌ エラー:", err.response?.data || err.message);
    }
}

// 実行

checkUnenteredWork();





