const axios = require("axios");

//Smartsheetの設定
const ACCESS_TOKEN = process.env.ACCESS_TOKEN;
const SHEET_ID = process.env.SHEET_ID;

const headers = {
    "Authorization":`Bearer ${ACCESS_TOKEN}`,
    "Content-Type":"application/json"
};

//取り出した列名を本物のDateに変換したり同じ日かチェックしたり処理で日付として使える形に残す処理を用意する(最後にループで呼び出し実行する)
//日付列列名をDateに変換
const parseYmd = (title) => {
    const m = title.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
};

//日付を時刻をなくし日付のみにする
const normalizeDate = (d) => new Date(d.getFullYear(), d.getMonth(), d.getDate());

//昨日と列名が同じか、比較できるようにする
const sameDay = (a, b) => a && b && normalizeDate(a).getTime() === normalizeDate(b).getTime();

//メイン処理
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

        /*
        //列タイトルを確認
        console.log("列タイトル一覧:", cols.map(c => c.title));
        */

        //必要な列を特定
        const flagCol = cols.find(c => c.title === "未入力");
        const statusCol = cols.find(c => c.title === "ステータス");
        const startCol = cols.find(c => c.title === "開始日（実績）");
        const endCol = cols.find(c => c.title === "終了日（実績）");

        if(!flagCol || !statusCol || !startCol || !endCol){
            throw new Error("必要な列が見つかりません。列名を確認してください。");
        }

        /*
        //未入力列のID確認
        console.log("未入力列 flagCol:", flagCol);
        */

        //全日付列を抽出　形式が日付っぽい列を取り出すための処理
        const dateCols = cols.filter(c => /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(c.title));

        //昨日の日付を作成
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        const yTitle = `${yesterday.getFullYear()}/${yesterday.getMonth()+1}/${yesterday.getDate()}`; //JavaScriptのDateオブジェクトは月が０始まりだから+1
        const yCol = cols.find(c => c.title.trim().startsWith(yTitle)); //startsWithで指定した文字列で始まっているかを判定

        /*
        //確認用ログ
        console.log("昨日の日付タイトル:", yTitle);
        console.log("見つかった列:", yCol);
        */
        
        const updateRows = [];

        //2.各行をチェック
        //rows シート全体の全行のデータ配列　シートの全行を１行ずつ処理するループを作成。ステータス、開始日、終了日の値を取得
        for(const row of rows){
            const status = row.cells.find(c => c.columnId === statusCol.id)?.value;
            const startDate = row.cells.find(c => c.columnId === startCol.id)?.value;
            const endDate = row.cells.find(c => c.columnId === endCol.id)?.value;

            let flag = false; //フラッグをすべてOFFにする

            //ステータスごとの処理
            //進行中　昨日の日付の列が空の場合フラグON　昨日の日付に対応する列yColが存在する場合実施
            if(status === "進行中" && yCol){
                const yCell = row.cells.find(c => c.columnId === yCol.id);         
                if(!yCell || yCell.value === null || yCell.value === undefined || yCell.value === ""){
                    flag = true; //昨日の列が空だった場合フラッグON
                }
            }

            //完了　開始日～終了日の間で空欄があればフラッグON
            if(status === "完了" && startDate && endDate){
                const s = normalizeDate(new Date(startDate)); //新しく日付オブジェクトを作る前にAPIから取ってくると文字列になるのでDate型にする。時刻を消すためにnormalizeDateを挟む
                const e = normalizeDate(new Date(endDate));

                //全シートのうち全日付列の中でループの中でparseYmdが実行される(列が順番に回ってきてその列に対応する行のセルを一つずつチェックしていくイメージ)
                for(const col of dateCols){
                    const d = parseYmd(col.title);
                    if (!d) continue;
                    const dNorm = normalizeDate(d);

                    if(dNorm >= s && dNorm <= e){ //開始日から終了日の間を取り出す
                        const cell = row.cells.find(c => c.columnId === col.id); //この行にあるcol(今チェック中の日付列)に対応するセルを取得する
                        if(!cell || cell.value === null || cell.value === undefined || cell.value === ""){
                            flag = true; //空白が見つかったら未入力フラグTRUE
                            break; //空白を見つけた時点で残りの列はチェックしない
                        }
                    }
                }
            }

            /*

            //更新リストに追加 ログ有バージョン
            const update = {
                id: row.id,
                cells:[{columnId: flagCol.id, value: flag}]
            };

            
            console.log("更新内容:", update);
            updateRows.push(update);

            */
            
            //更新リストに追加
            updateRows.push({
                id: row.id,
                cells:[{columnId: flagCol.id, value: flag}]
            });    
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
