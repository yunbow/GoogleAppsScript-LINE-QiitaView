/**
 * QiitaViewBOT
 */
const LINE_NOTIFY_TOKEN = '*****'; // LINE NOTIFYのアクセストークン
const SSID_QIITA_VIEW = '*****'; // QiitaのスプレッドシートのID
const SSN_QIITA_VIEW = 'QiitaVIEW'; // Qiitaのスプレッドシートのシート名
const QIITA_TOKEN = '*****'; // Qiitaトークン
const TITLE_MAX_LENGTH = 40; // 記事タイトル最大長
const WEEKDAY = ["日", "月", "火", "水", "木", "金", "土"];

let spreadsheet = SpreadsheetApp.openById(SSID_QIITA_VIEW);
let sheet = spreadsheet.getSheetByName(SSN_QIITA_VIEW);

/**
 * メイン処理
 */
function main() {
    try {
        addViewCount();

        let itemList = getItemList();

        for (let i in itemList) {
            let item = itemList[i];
            item.increase = 0;
            if (0 < item.list.length) {
                let list = [].concat(item.list);
                item.increase = list.pop().viewCount - list.pop().viewCount;
            }
        }

        itemList.sort((a, b) => {
            return (a.increase > b.increase) ? -1 : 1;
        });

        let nowDt = new Date();
        let dt = Utilities.formatDate(nowDt, 'Asia/Tokyo', `MM/dd(${WEEKDAY[nowDt.getDay()]})`);
        let message = `\n今日のQiitaViewだよ!!\n\n--- ${dt} ----\n\n`;
        for (let i = 0; i < 5; i++) {
            let item = itemList[i];
            let list = [].concat(item.list);
            let data = getQiitaItem(item.id);
            message += `${parseInt(i) + 1}: +${item.increase} (${list.pop().viewCount} views)\n`
            message += `${omit(data.title)}\n\n`;
        }

        let range = sheet.getRange(sheet.getLastRow(), 2, sheet.getLastRow(), sheet.getLastColumn());
        let chart = sheet.newChart().addRange(range).setChartType(Charts.ChartType.COLUMN).setOption('title', 'QiitaViews').build();
        let imageBlob = chart.getBlob().getAs('image/png').setName(`chart.png`);

        console.log(message);
        sendLineNotify(message, imageBlob);

    } catch (e) {
        console.error(e.stack);
    }
}

/**
 * Qiita View数を取得して、スプレッドシートに保存する
 */
function addViewCount() {
    let idList;
    let lastRow = sheet.getLastRow();
    if (0 < lastRow) {
        idList = sheet.getRange(1, 2, 1, sheet.getLastColumn()).getValues()[0];
    }

    let timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    let viewCountList = [timestamp];
    let itemList = getQiitaItemList();
    let addCount = 0;
    let lastColumn = 0;
    if (0 < lastRow) {
        lastColumn = sheet.getLastColumn() - 1;
    }

    for (let i in itemList) {
        let item = itemList[i];
        item = getQiitaItem(item.id);

        let index = 0;
        let isExist = false;

        if (0 < lastRow) {
            for (let j in idList) {
                let id = idList[j];
                if (id == item.id) {
                    isExist = true;
                    index = parseInt(j);
                }
            }
        }
        if (!isExist) {
            index = lastColumn + addCount;
            addCount++;
        }
        sheet.getRange(1, (index + 2)).setValue(item.id);
        viewCountList[index + 1] = item.page_views_count;
    }
    sheet.appendRow(viewCountList);
}

/**
 * スプレッドシートのデータを取得する
 */
function getItemList() {

    let itemList = [];
    let lastRow = sheet.getLastRow();
    let lastColumn = sheet.getLastColumn();
    let rowList = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

    for (let i = 1; i < lastColumn; i++) {
        let id = rowList[0][i];
        let viewList = [];

        for (let j = 1; j < lastRow; j++) {
            viewList.push({
                createDt: rowList[j][0],
                viewCount: rowList[j][i]
            });
        }
        itemList.push({
            id: id,
            list: viewList
        });
    }
    return itemList;
}

/**
 * 文字列を省略する
 * @param {String} str 
 */
function omit(str) {
    if (TITLE_MAX_LENGTH < str.length) {
        return str.slice(0, TITLE_MAX_LENGTH) + '...';
    } else {
        return str;
    }
}

/**
 * ユーザー記事一覧を取得する
 * @return JSON
 */
function getQiitaItemList() {
    let url = `https://qiita.com/api/v2/authenticated_user/items?page=1&per_page=100`;
    let options = {
        'method': 'get',
        'headers': {
            'Authorization': `Bearer ${QIITA_TOKEN}`
        },
    };
    let response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText('UTF-8'));
}

/**
 * ユーザー記事を取得する
 * @param {String} id
 * @return JSON
 */
function getQiitaItem(id) {
    let url = `https://qiita.com/api/v2/items/${id}`;
    let options = {
        'method': 'get',
        'headers': {
            'Authorization': `Bearer ${QIITA_TOKEN}`
        },
    };
    let response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText('UTF-8'));
}

/**
 * LINEにメッセージを送信する
 * @param {String} message メッセージ
 * @param {Object} blob 画像ファイル
 */
function sendLineNotify(message, blob) {
    let url = 'https://notify-api.line.me/api/notify';
    let options = {
        'method': 'post',
        'headers': {
            'Authorization': `Bearer ${LINE_NOTIFY_TOKEN}`
        },
        'payload': {
            'message': message,
            'imageFile': blob
        }
    };
    let response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText('UTF-8'));
}