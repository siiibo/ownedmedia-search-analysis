import { endOfDay, format } from "date-fns";

class KeywordUrl {
    [key: string]: string[];
}

export const init = () => {
    const dataSpreadSheet = getDataSpreadSheet();
    // トリガーの追加
    ScriptApp.newTrigger(createOnOpen.name).forSpreadsheet(dataSpreadSheet).onOpen().create();
};

const getDataSpreadSheet = (): GoogleAppsScript.Spreadsheet.Spreadsheet => {
    const spreadSheetUrl = PropertiesService.getScriptProperties().getProperty("SPREAD_SHEET_URL");
    if (!spreadSheetUrl) throw new Error("SPREAD_SHEET_URL is not defined");
    return SpreadsheetApp.openByUrl(spreadSheetUrl);
};

export const createOnOpen = () => {
    SpreadsheetApp.getUi().createAddonMenu().addItem("検索の実行", askExecute.name).addToUi();
};

export const askExecute = () => {
    const question = Browser.msgBox("検索を実行しますか?", Browser.Buttons.YES_NO);
    if (question == "yes") {
        getKeywordData();
    }
};

export const getKeywordData = () => {
    //スプレッドシートから期間の取得
    const dataSpreadSheet = getDataSpreadSheet();
    const periodSheet = dataSpreadSheet.getSheetByName("期間指定");
    if (!periodSheet) throw new Error("SHEET is not defined");
    const startDate = periodSheet.getRange("B4").getValue();
    const endDate = endOfDay(periodSheet.getRange("C4").getValue());
    getSearchData(dataSpreadSheet, startDate, endDate);
};

const getSearchData = (dataSpreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet, startDate: Date, endDate: Date) => {
    // 結果記入用シートの追加
    const keywordResultSheet = dataSpreadSheet.insertSheet(dataSpreadSheet.getNumSheets());
    const keywordUrlReusltSheet = dataSpreadSheet.insertSheet(dataSpreadSheet.getNumSheets());

    // タイトルの設定
    const title = [["キーワード", "記事URL", "クリック数", "インプレッション", "平均順位", "平均CTR"]];
    console.log("title", title[0].length);
    keywordResultSheet.getRange(1, 1, 1, 1).setValues([["意図していない表示URL"]]);
    keywordResultSheet.getRange(2, 1, 1, title[0].length).setValues(title);
    keywordResultSheet.getRange(1, 1 + title[0].length + 1, 1, 1).setValues([["枝付きURL"]]);
    keywordResultSheet.getRange(2, 1 + title[0].length + 1, 1, title[0].length).setValues(title);
    keywordUrlReusltSheet.getRange(1, 1, 1, title[0].length).setValues(title);

    //サーチコンソールに登録しているサイトドメイン
    const siteDomain = "siiibo.com";

    //リクエストするAPIのURLを設定
    const apiURL =
        "https://www.googleapis.com/webmasters/v3/sites/sc-domain%3A" + siteDomain + "/searchAnalytics/query";
    //サーチコンソールから取得するキーワードの最大数を設定する
    const maxRecord = 1000;

    // KWの取得，KW URLリストの作成
    const keywordUrlSheet = dataSpreadSheet.getSheetByName("対キーワードURL週次検索結果");
    if (!keywordUrlSheet) throw new Error("SHEET is not defined");

    const keywordStartRow = 2;
    const keywordEndColumn = 1;
    const keywordColumnSize = 2;
    const rowEndData = keywordUrlSheet.getDataRange().getLastRow();
    const keywordRowSize = rowEndData - (keywordStartRow - 1);
    const keywordUrlData = keywordUrlSheet
        .getRange(keywordStartRow, keywordEndColumn, keywordRowSize, keywordColumnSize)
        .getValues();

    const keywordUrl = new KeywordUrl();

    for (let i = 0; i < keywordUrlData.length; i++) {
        const keyword: string = keywordUrlData[i][0];
        const url: string = keywordUrlData[i][1];
        keywordUrl[keyword] = [];
        keywordUrl[keyword].push(url);
    }

    // それぞれのKWについて結果を取得 → URLの条件で考える
    // dimensionの値によるグループ化
    //リクエスト送る時のペイロード指定(=出力条件)
    for (const keyword of Object.keys(keywordUrl)) {
        const keyword1 = keyword.replace(" ", "( |　)");
        const keyword2 = keyword1.replace("　", "( |　)");
        const keyword3 = "^" + keyword2 + "$";
        console.log(keyword);
        const payload = {
            startDate: format(startDate, "yyyy-MM-dd"),
            endDate: format(endDate, "yyyy-MM-dd"),
            dimensions: ["query", "page"], // このフィルターが適用されるディメンション。
            rowLimit: maxRecord, //取得するキーワードの最大数
            dimensionFilterGroups: [
                {
                    filters: [
                        {
                            dimension: "query", //指定されたクエリ文字列に対してフィルター処理します。
                            operator: "includingRegex", //指定した値が行のディメンション値とどのように一致する (または一致しない) 必要があるか
                            expression: keyword3, //演算子に応じて、一致または除外するフィルターの値。
                        },
                    ],
                },
            ],
        };
        //ヘッダーのオプション
        const options = {
            payload: JSON.stringify(payload),
            myamethod: "POST",
            muteHttpExceptions: true,
            headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
            contentType: "application/json",
        };

        //APIリクエス送信とJSONの分解
        const response = UrlFetchApp.fetch(apiURL, options);
        const json = JSON.parse(response.getContentText());

        const urlMatched = [];
        const urlNotMatched = [];
        const urlBranched = [];

        // filterを使うように書きなおしたい
        if (!(typeof json["rows"] === "undefined" || json["rows"].length === 0)) {
            //分解したデータを配列化して、入れ物の配列にpushでぶちこんでいく
            for (let i = 0; i < json["rows"].length; i++) {
                // URLが対策URLと一致するなら
                if (keywordUrl[keyword].includes(json["rows"][i]["keys"][1])) {
                    urlMatched.push([
                        json["rows"][i]["keys"][0],
                        json["rows"][i]["keys"][1],
                        json["rows"][i]["clicks"],
                        json["rows"][i]["impressions"],
                        json["rows"][i]["position"],
                        json["rows"][i]["ctr"],
                    ]);
                }
                // 対策URLと一致しないかつ枝付きじゃないかつクリック数が1以上
                else if (!json["rows"][i]["keys"][1].match("#") && json["rows"][i]["clicks"] >= 1) {
                    urlNotMatched.push([
                        json["rows"][i]["keys"][0],
                        json["rows"][i]["keys"][1],
                        json["rows"][i]["clicks"],
                        json["rows"][i]["impressions"],
                        json["rows"][i]["position"],
                        json["rows"][i]["ctr"],
                    ]);
                }
                // URLが枝付きかつクリックが1以上なら
                else if (json["rows"][i]["clicks"] >= 1) {
                    urlBranched.push([
                        json["rows"][i]["keys"][0],
                        json["rows"][i]["keys"][1],
                        json["rows"][i]["clicks"],
                        json["rows"][i]["impressions"],
                        json["rows"][i]["position"],
                        json["rows"][i]["ctr"],
                    ]);
                }
            }

            if (urlMatched.length >= 1) {
                const urlMatchedColumnBVals = keywordUrlReusltSheet.getRange("A:A").getValues();
                const urlMatchedLastRow = urlMatchedColumnBVals.filter(String).length;
                keywordUrlReusltSheet
                    .getRange(urlMatchedLastRow + 1, 1, urlMatched.length, urlMatched[0].length)
                    .setValues(urlMatched);
                keywordUrlReusltSheet.getRange(urlMatchedLastRow + 1, 6, urlMatched.length).setNumberFormat("0.00%"); //CTRの表示形式変更
            }
            console.log(urlNotMatched);
            if (urlNotMatched.length >= 1) {
                const urlNotMatchedColumnBVals = keywordResultSheet.getRange("A:A").getValues();
                const urlNotMatchedLastRow = urlNotMatchedColumnBVals.filter(String).length;
                keywordResultSheet
                    .getRange(urlNotMatchedLastRow + 1, 1, urlNotMatched.length, urlNotMatched[0].length)
                    .setValues(urlNotMatched);
                keywordResultSheet.getRange(urlNotMatchedLastRow + 1, 6, urlNotMatched.length).setNumberFormat("0.00%"); //CTRの表示形式変更
            }

            if (urlBranched.length >= 1) {
                console.log("branched", urlBranched);
                const urlBranchedColumnBVals = keywordResultSheet.getRange("H:H").getValues();
                const urlBranchedLastRow = urlBranchedColumnBVals.filter(String).length;
                keywordResultSheet
                    .getRange(urlBranchedLastRow + 1, 8, urlBranched.length, urlBranched[0].length)
                    .setValues(urlBranched);
                keywordResultSheet.getRange(urlBranchedLastRow + 1, 13, urlBranched.length).setNumberFormat("0.00%"); //CTRの表示形式変更
            }
        } else {
            console.log("該当するデータがありませんでした。");
        }
    }
    keywordResultSheet.setName(
        format(startDate, "yyyy-MM-dd") + "~" + format(endDate, "MM-dd") + "-" + "対キーワード週次検索結果"
    );
    keywordUrlReusltSheet.setName(
        format(startDate, "yyyy-MM-dd") + "~" + format(endDate, "MM-dd") + "-" + "対キーワードURL週次検索結果"
    );
};
