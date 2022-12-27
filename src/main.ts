import { endOfDay, format } from "date-fns";

type KeywordUrl = {
    [key: string]: string[];
};

// 検索を実行するまで: 対象スプシの指定→対象スプシをOpenしたらUI上にアドオンメニューを追加する→操作者がUIを操作して検索実行
export const init = () => {
    const spreadSheet = getSpreadSheet();
    // トリガーの追加
    ScriptApp.newTrigger(createOnOpen.name).forSpreadsheet(spreadSheet).onOpen().create();
};

const getSpreadSheet = (): GoogleAppsScript.Spreadsheet.Spreadsheet => {
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
        main();
    }
};

export const main = () => {
    //スプレッドシートから期間の取得
    const spreadSheet = getSpreadSheet();
    const startEndDate = getStartEndDate(spreadSheet);
    const startDate = startEndDate.start;
    const endDate = startEndDate.end;
    getSearchConsoleReusults(spreadSheet, startDate, endDate);
};
const getStartEndDate = (spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
    const periodSheet = spreadSheet.getSheetByName("期間指定");
    if (!periodSheet) throw new Error("periodSheet is not defined");
    const startDate = periodSheet.getRange("B4").getValue();
    const endDate = endOfDay(periodSheet.getRange("C4").getValue());
    return { start: startDate, end: endDate };
};

const setHeader = (
    keywordResultSheet: GoogleAppsScript.Spreadsheet.Sheet,
    keywordUrlReusltSheet: GoogleAppsScript.Spreadsheet.Sheet
) => {
    const header = [["キーワード", "記事URL", "クリック数", "インプレッション", "平均順位", "平均CTR"]];
    keywordResultSheet.getRange(1, 1, 1, 1).setValues([["意図していない表示URL"]]);
    keywordResultSheet.getRange(2, 1, 1, header[0].length).setValues(header);
    keywordResultSheet.getRange(1, 1 + header[0].length + 1, 1, 1).setValues([["枝付きURL"]]);
    keywordResultSheet.getRange(2, 1 + header[0].length + 1, 1, header[0].length).setValues(header);
    keywordUrlReusltSheet.getRange(1, 1, 1, header[0].length).setValues(header);
};

const getKeywordUrlClass = (spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet) => {
    // KWの取得，KW URLリストの作成
    const keywordUrlSheet = spreadSheet.getSheetByName("対キーワードURL週次検索結果");
    if (!keywordUrlSheet) throw new Error("SHEET is not defined");
    const keywordUrlData = keywordUrlSheet.getDataRange().getValues();
    const [headers, ...records] = keywordUrlData;

    const keywordUrl: KeywordUrl = {};

    for (const record of records) {
        const keyword: string = record[0];
        keywordUrl[keyword] = [];
    }

    for (const record of records) {
        const keyword: string = record[0];
        const url: string = record[1];
        keywordUrl[keyword].push(url);
    }
    console.log(keywordUrl);
    return keywordUrl;
};

const getSearchConsoleReusults = (
    spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
    startDate: Date,
    endDate: Date
) => {
    // 結果記入用シートの追加
    const keywordResultSheet = spreadSheet.insertSheet(3);
    const keywordUrlReusltSheet = spreadSheet.insertSheet(4);

    // タイトルの設定
    setHeader(keywordResultSheet, keywordUrlReusltSheet);

    //サーチコンソールに登録しているサイトドメイン
    const siteDomain = "siiibo.com";

    //リクエストするAPIのURLを設定
    const apiURL =
        "https://www.googleapis.com/webmasters/v3/sites/sc-domain%3A" + siteDomain + "/searchAnalytics/query";
    //サーチコンソールから取得するキーワードの最大数を設定する
    const maxRecord = 1000;

    const keywordUrl = getKeywordUrlClass(spreadSheet);

    for (const keyword of Object.keys(keywordUrl)) {
        // KWを半角全角許容する
        const keyword_ = "^" + keyword.replace(" ", "( |　)").replace("　", "( |　)") + "$";

        // ペイロードの設定 キーワードひとつずつにしか送れない?
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
                            expression: keyword_, //演算子に応じて、一致または除外するフィルターの値。
                        },
                    ],
                },
            ],
        };

        //ヘッダーのオプション指定
        const options = {
            payload: JSON.stringify(payload),
            myamethod: "POST",
            muteHttpExceptions: true,
            headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
            contentType: "application/json",
        };

        //APIにリクエスト送信→レスポンスをもらう
        const response = UrlFetchApp.fetch(apiURL, options);
        //レスポンスの内容をJSONファイルへ
        const json = JSON.parse(response.getContentText());

        const urlMatched = []; //URLに一致
        const urlNotMatched = []; //URLに不一致
        const urlBranched = []; //分岐

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
