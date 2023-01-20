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

type FormattedResponse = {
    clicks: number;
    ctr: number;
    impressions: number;
    position: number;
    query: string;
    page: string;
}[];

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

    const keywordUrlEntries = Object.entries(keywordUrl).filter(
        (kv): kv is [string, KeywordUrl[]] => kv[1] != undefined
    );

    const searchConsoleResponses = keywordUrlEntries.map(([keyword, keywordUrls]) => {
        const response = getDataFromSearchConsole(keyword, startDate, endDate);
        return { response, keyword, keywordUrls };
    });

    const formattedDataList = searchConsoleResponses.map(({ response, keywordUrls }) => {
        const urls = keywordUrls.map((keywordUrl) => {
            return keywordUrl.url;
        });
        return formatData(response, urls);
    });

    writeInSpreadsheet(formattedDataList, resultSheet);
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

    const httpResponse = UrlFetchApp.fetch(apiUrl, options);
    const response: SearchConsoleResponse = JSON.parse(httpResponse.getContentText());
    return response;
};

const formatData = (
    response: SearchConsoleResponse,
    urls: string[]
): {
    withAnchor: FormattedResponse;
    matchedWithoutAnchor: FormattedResponse;
    notMatchedWithoutAnchor: FormattedResponse;
} => {
    const formattedResponse = response["rows"].map(({ keys, ...rest }) => {
        return {
            query: keys[0],
            page: keys[1],
            ...rest,
        };
    });
    /**
     * アンカー付き, 不一致はさらに「クリック数1以上のみ」で絞り込みを行う. 完全一致は行わない.
     *
     * 参考: https://github.com/siiibo/ownedmedia-search-analysis/pull/4#discussion_r1080962946
     */
    const withAnchor = formattedResponse.filter((row) => row["page"].includes("#") && row["clicks"] >= 1);
    const withoutAnchor = formattedResponse.filter((row) => !row["page"].includes("#"));
    const matchedWithoutAnchor = withoutAnchor.filter((row) => urls.includes(row["page"]));
    const notMatchedWithoutAnchor = withoutAnchor.filter((row) => !urls.includes(row["page"]) && row["clicks"] >= 1);

    return { withAnchor, matchedWithoutAnchor, notMatchedWithoutAnchor };
};

const writeInSpreadsheet = (
    formattedDataList: {
        withAnchor: FormattedResponse;
        matchedWithoutAnchor: FormattedResponse;
        notMatchedWithoutAnchor: FormattedResponse;
    }[],
    resultSheet: GoogleAppsScript.Spreadsheet.Sheet
) => {
    const results = formattedDataList.flatMap((data) => {
        const resultWithAnchor = data.withAnchor.map((row) => [
            row["query"],
            row["page"],
            "アンカー付き",
            row["clicks"],
            row["impressions"],
            row["position"],
            row["ctr"],
        ]);

        const resultMatchedWithoutAnchor = data.matchedWithoutAnchor.map((row) => [
            row["query"],
            row["page"],
            "完全一致",
            row["clicks"],
            row["impressions"],
            row["position"],
            row["ctr"],
        ]);

        const resultNotMatchedWithoutAnchor = data.notMatchedWithoutAnchor.map((row) => [
            row["query"],
            row["page"],
            "不一致",
            row["clicks"],
            row["impressions"],
            row["position"],
            row["ctr"],
        ]);

        return resultMatchedWithoutAnchor.concat(resultNotMatchedWithoutAnchor).concat(resultWithAnchor);
    });

    const resultsRowNum = results.length;
    const resultsColumnNum = results[0].length;

    if (resultsRowNum >= 1) {
        const lastRow = resultSheet.getLastRow();
        resultSheet.getRange(lastRow + 1, 1, resultsRowNum, resultsColumnNum).setValues(results);
        resultSheet.getRange(lastRow + 1, 7, resultsRowNum, 1).setNumberFormat("0.00%");
    }
};
