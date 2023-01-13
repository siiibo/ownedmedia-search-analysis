import { endOfDay, format } from "date-fns";
import { group } from "radash";

type KeywordUrl = {
    keyword: string;
    url: string;
};

type SearchConsoleResponse = {
    responseAggregationType: string;
    rows: {
        clicks: number;
        ctr: number;
        impressions: number;
        keys: string[];
        position: number;
    }[];
};

export const init = () => {
    const spreadsheet = getSpreadsheet();
    ScriptApp.newTrigger(createOnOpen.name).forSpreadsheet(spreadsheet).onOpen().create();
};

const getSpreadsheet = (): GoogleAppsScript.Spreadsheet.Spreadsheet => {
    const spreadsheetUrl = PropertiesService.getScriptProperties().getProperty("SPREAD_SHEET_URL");
    if (!spreadsheetUrl) throw new Error("SPREAD_SHEET_URL is not defined");
    return SpreadsheetApp.openByUrl(spreadsheetUrl);
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
    const spreadsheet = getSpreadsheet();
    const periodSheet = spreadsheet.getSheetByName("期間指定");
    if (!periodSheet) throw new Error("periodSheet is not defined");
    const { startDate, endDate } = getStartEndDate(periodSheet);

    const keywordUrlSheet = spreadsheet.getSheetByName("対キーワードURL週次検索結果");
    if (!keywordUrlSheet) throw new Error("SHEET is not defined");

    const keywordUrl = getUrlsGroupedByKeyword(keywordUrlSheet);

    const keywordResultSheet = spreadsheet.insertSheet(3);
    const keywordUrlResultSheet = spreadsheet.insertSheet(4);

    setHeader(keywordResultSheet, keywordUrlResultSheet);

    for (const [keyword, values] of Object.entries(keywordUrl)) {
        if (values == undefined) {
            continue;
        }
        const urls = values.map((value) => {
            return value.url;
        });
        const responseData = getDataFromSearchConsole(keyword, startDate, endDate);

        if (!(typeof responseData["rows"] === "undefined" || responseData["rows"].length === 0)) {
            if (keywordUrl[keyword] != undefined) {
                const results = formatData(responseData, urls);
                const urlMatched = results.matched;
                const urlNotMatched = results.notMatched;
                const urlBranched = results.branched;

                writeInSpreadsheet(urlMatched, urlNotMatched, urlBranched, keywordUrlResultSheet, keywordResultSheet);
            }
        } else {
            console.log("該当するデータがありませんでした。");
        }
    }
    keywordResultSheet.setName(
        format(startDate, "yyyy-MM-dd") + "~" + format(endDate, "MM-dd") + "-" + "対キーワード週次検索結果"
    );
    keywordUrlResultSheet.setName(
        format(startDate, "yyyy-MM-dd") + "~" + format(endDate, "MM-dd") + "-" + "対キーワードURL週次検索結果"
    );
};
const getStartEndDate = (periodSheet: GoogleAppsScript.Spreadsheet.Sheet): { startDate: Date; endDate: Date } => {
    const startDate = periodSheet.getRange("B4").getValue();
    const endDate = endOfDay(periodSheet.getRange("C4").getValue());
    return { startDate, endDate };
};

const setHeader = (
    keywordResultSheet: GoogleAppsScript.Spreadsheet.Sheet,
    keywordUrlResultSheet: GoogleAppsScript.Spreadsheet.Sheet
) => {
    const header = [["キーワード", "記事URL", "クリック数", "インプレッション", "平均順位", "平均CTR"]];
    keywordResultSheet.getRange(1, 1, 1, 1).setValues([["意図していない表示URL"]]);
    keywordResultSheet.getRange(2, 1, 1, header[0].length).setValues(header);
    keywordResultSheet.getRange(1, 1 + header[0].length + 1, 1, 1).setValues([["枝付きURL"]]);
    keywordResultSheet.getRange(2, 1 + header[0].length + 1, 1, header[0].length).setValues(header);
    keywordUrlResultSheet.getRange(1, 1, 1, header[0].length).setValues(header);
};

function getUrlsGroupedByKeyword(keywordUrlSheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const [_sheetHeader, ...sheetValues] = keywordUrlSheet.getDataRange().getValues();
    const keywordUrls: KeywordUrl[] = sheetValues.map((row) => {
        return {
            keyword: row[0],
            url: row[1],
        };
    });
    const urlGroupedByKeyword = group(keywordUrls, (keywordUrl) => keywordUrl.keyword);
    return urlGroupedByKeyword;
}

/**
 * スペース（全角・半角）を、全角または半角スペースを受け入れる正規表現に変換する.
 *
 * キーワードに含まれるスペースの全角・半角の違いにより，検索結果が違いが生じるので，
 * 全角・半角のどちらでも同じ結果を得るため．
 */
const convertSpaceToZenkakuHankakuSpaceRegExp = (s: string) => {
    return `${s.replace(/ |　/g, "( |　)")}`;
};

const getDataFromSearchConsole = (keyword: string, startDate: Date, endDate: Date): SearchConsoleResponse => {
    const keywordRegExp = `^${convertSpaceToZenkakuHankakuSpaceRegExp(keyword)}$`;
    const maxRow = 1000;

    const siteDomain = "siiibo.com";

    const apiUrl =
        "https://www.googleapis.com/webmasters/v3/sites/sc-domain%3A" + siteDomain + "/searchAnalytics/query";

    const payload = {
        startDate: format(startDate, "yyyy-MM-dd"),
        endDate: format(endDate, "yyyy-MM-dd"),
        dimensions: ["query", "page"],
        rowLimit: maxRow, //取得するキーワードの最大数
        dimensionFilterGroups: [
            {
                filters: [
                    {
                        dimension: "query",
                        operator: "includingRegex",
                        expression: keywordRegExp,
                    },
                ],
            },
        ],
    };

    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        payload: JSON.stringify(payload),
        method: "post",
        muteHttpExceptions: true,
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        contentType: "application/json",
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData: SearchConsoleResponse = JSON.parse(response.getContentText());
    return responseData;
};

const formatData = (
    responseData: SearchConsoleResponse,
    urls: string[] | undefined
): { matched: (string | number)[][]; notMatched: (string | number)[][]; branched: (string | number)[][] } => {
    const urlMatched = [];
    const urlNotMatched = [];
    const urlBranched = [];

    for (let i = 0; i < responseData["rows"].length; i++) {
        // URLが対策URLと一致するなら
        if (urls?.includes(responseData["rows"][i]["keys"][1])) {
            urlMatched.push([
                responseData["rows"][i]["keys"][0],
                responseData["rows"][i]["keys"][1],
                responseData["rows"][i]["clicks"],
                responseData["rows"][i]["impressions"],
                responseData["rows"][i]["position"],
                responseData["rows"][i]["ctr"],
            ]);
        }
        // 対策URLと一致しないかつ枝付きじゃないかつクリック数が1以上
        else if (!responseData["rows"][i]["keys"][1].match("#") && responseData["rows"][i]["clicks"] >= 1) {
            urlNotMatched.push([
                responseData["rows"][i]["keys"][0],
                responseData["rows"][i]["keys"][1],
                responseData["rows"][i]["clicks"],
                responseData["rows"][i]["impressions"],
                responseData["rows"][i]["position"],
                responseData["rows"][i]["ctr"],
            ]);
        }
        // URLが枝付きかつクリックが1以上なら
        else if (responseData["rows"][i]["clicks"] >= 1) {
            urlBranched.push([
                responseData["rows"][i]["keys"][0],
                responseData["rows"][i]["keys"][1],
                responseData["rows"][i]["clicks"],
                responseData["rows"][i]["impressions"],
                responseData["rows"][i]["position"],
                responseData["rows"][i]["ctr"],
            ]);
        }
    }
    return { matched: urlMatched, notMatched: urlNotMatched, branched: urlBranched };
};

const writeInSpreadsheet = (
    urlMatched: (string | number)[][],
    urlNotMatched: (string | number)[][],
    urlBranched: (string | number)[][],
    keywordUrlResultSheet: GoogleAppsScript.Spreadsheet.Sheet,
    keywordResultSheet: GoogleAppsScript.Spreadsheet.Sheet
) => {
    if (urlMatched.length >= 1) {
        const urlMatchedColumnBVals = keywordUrlResultSheet.getRange("A:A").getValues();
        const urlMatchedLastRow = urlMatchedColumnBVals.filter(String).length;
        keywordUrlResultSheet
            .getRange(urlMatchedLastRow + 1, 1, urlMatched.length, urlMatched[0].length)
            .setValues(urlMatched);
        keywordUrlResultSheet.getRange(urlMatchedLastRow + 1, 6, urlMatched.length).setNumberFormat("0.00%");
    }
    console.log(urlNotMatched);
    if (urlNotMatched.length >= 1) {
        const urlNotMatchedColumnBVals = keywordResultSheet.getRange("A:A").getValues();
        const urlNotMatchedLastRow = urlNotMatchedColumnBVals.filter(String).length;
        keywordResultSheet
            .getRange(urlNotMatchedLastRow + 1, 1, urlNotMatched.length, urlNotMatched[0].length)
            .setValues(urlNotMatched);
        keywordResultSheet.getRange(urlNotMatchedLastRow + 1, 6, urlNotMatched.length).setNumberFormat("0.00%");
    }

    if (urlBranched.length >= 1) {
        console.log("branched", urlBranched);
        const urlBranchedColumnBVals = keywordResultSheet.getRange("H:H").getValues();
        const urlBranchedLastRow = urlBranchedColumnBVals.filter(String).length;
        keywordResultSheet
            .getRange(urlBranchedLastRow + 1, 8, urlBranched.length, urlBranched[0].length)
            .setValues(urlBranched);
        keywordResultSheet.getRange(urlBranchedLastRow + 1, 13, urlBranched.length).setNumberFormat("0.00%");
    }
};
