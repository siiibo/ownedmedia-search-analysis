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

    const keywordUrlSheet = spreadsheet.getSheetByName("キーワードURL指定");
    if (!keywordUrlSheet) throw new Error("SHEET is not defined");

    const keywordUrl = getUrlsGroupedByKeyword(keywordUrlSheet);

    const resultSheet = spreadsheet.insertSheet(
        `${format(startDate, "yyyy-MM-dd")}~${format(endDate, "MM-dd")}-掲載順位結果`,
        3
    );

    setHeader(resultSheet);

    const results: (string | number)[][][] = [];
    Object.entries(keywordUrl)
        .filter((kv): kv is [string, KeywordUrl[]] => kv[1] != undefined)
        .forEach(([keyword, values]) => {
            const urls = values.map((value) => {
                return value.url;
            });
            const response = getDataFromSearchConsole(keyword, startDate, endDate);
            const result: (string | number)[][] = formatData(response, urls);
            results.push(result);
        });

    results.forEach((result) => writeInSpreadsheet(result, resultSheet));
};
const getStartEndDate = (periodSheet: GoogleAppsScript.Spreadsheet.Sheet): { startDate: Date; endDate: Date } => {
    const startDate = periodSheet.getRange("B4").getValue();
    const endDate = endOfDay(periodSheet.getRange("C4").getValue());
    return { startDate, endDate };
};

const setHeader = (resultSheet: GoogleAppsScript.Spreadsheet.Sheet) => {
    const header = [["キーワード", "記事URL", "タイプ", "クリック数", "インプレッション", "平均順位", "平均CTR"]];
    resultSheet.getRange(1, 1, 1, header[0].length).setValues(header);
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

const formatData = (responseData: SearchConsoleResponse, urls: string[]): (string | number)[][] => {
    const withAnchor = responseData["rows"].filter((row) => row["keys"][1].match("#") && row["clicks"] >= 1);
    const withoutAnchor = responseData["rows"].filter((row) => !row["keys"][1].match("#"));
    const matchedWithoutAnchor = withoutAnchor.filter((row) => urls?.includes(row["keys"][1]));
    const notMatchedWithoutAnchor = withoutAnchor.filter(
        (row) => !urls?.includes(row["keys"][1]) && row["clicks"] >= 1
    );
    const resultWithAnchor = withAnchor.map((row) => [
        row["keys"][0],
        row["keys"][1],
        "アンカー付き",
        row["clicks"],
        row["impressions"],
        row["position"],
        row["ctr"],
    ]);

    const resultMatchedWithoutAnchor = matchedWithoutAnchor.map((row) => [
        row["keys"][0],
        row["keys"][1],
        "完全一致",
        row["clicks"],
        row["impressions"],
        row["position"],
        row["ctr"],
    ]);

    const resultNotMatchedWithoutAnchor = notMatchedWithoutAnchor.map((row) => [
        row["keys"][0],
        row["keys"][1],
        "不一致",
        row["clicks"],
        row["impressions"],
        row["position"],
        row["ctr"],
    ]);

    const result = resultMatchedWithoutAnchor.concat(resultNotMatchedWithoutAnchor).concat(resultWithAnchor);
    return result;
};

const writeInSpreadsheet = (result: (string | number)[][], resultSheet: GoogleAppsScript.Spreadsheet.Sheet) => {
    if (result.length >= 1) {
        const resultColumnBVals = resultSheet.getRange("A:A").getValues();
        const resultLastRow = resultColumnBVals.filter(String).length;
        resultSheet.getRange(resultLastRow + 1, 1, result.length, result[0].length).setValues(result);
        resultSheet.getRange(resultLastRow + 1, 7, result.length).setNumberFormat("0.00%");
    }
};
