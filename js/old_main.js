/**
 * 各シートのヘッダー要素
 */
const headerSheetTabDiv = $('#header-sheet-wrap');

/**
 * エビデンス添付テーブル要素
 */
const tableWrap = $('#table-wrap');

/**
 * シート用タブ要素
 */
const sheetTabDiv = $('#sheet-tab-wrap');

/**
 * ロード中のインジケータ
 */
const loader = $('<div>', { id: 'loader', class: 'loader' }).appendTo('body').hide();

/**
 * ロード中のオーバーレイ背景
 */
const overlay = $('<div>', { id: 'overlay', class: 'overlay' }).appendTo('body').hide();

/**
 * Excelファイル格納用
 */
let workbook2 = null;

/**
 * エビデンスフォルダ階層でのアラート表示エラーメッセージ
 */
const errorFolderStructureAlertMessage = "エビデンスフォルダジェネレーターを用いてフォルダ生成し、再度選択してください。";

/**
 * エビデンスファイル添付処理
 */
$('#attachment-evidencefile-btn').on('click', async (event) => {
    const btn = $(event.currentTarget);
    btn.prop('disabled', true);

    // ロード中インジケータとオーバーレイ表示
    loader.show();
    overlay.show();

    try {
        const directoryHandle = await window.showDirectoryPicker();
        isFolderStructureValid = false;
        const structure = await readDirectory(directoryHandle);
        if (structure.length == 0) {
            throw new Error(errorFolderStructureMessage);
        }
        createMyTable(structure);
        $('#dl-excel').attr('disabled', false);
    } catch (error) {
        if (error instanceof DOMException && error.name === 'AbortError') {
            console.log("ディレクトリ読み込みをキャンセルしました。");
            return;
        }

        console.error('Error opening directory:', error);
        alert(errorFolderStructureAlertMessage);
    } finally {
        btn.prop('disabled', false);
        // ロード中インジケータとオーバーレイ非表示
        loader.hide();
        overlay.hide();
    }
});

/**
 * フォルダ構造データを基にテーブルを動的に生成し、シートタブを生成
 * @param {Array} structure テーブルとシートの構造データ
 */
function createMyTable(structure) {

    // 既にテーブル要素が存在する場合削除
    headerSheetTabDiv.empty();
    tableWrap.empty();
    sheetTabDiv.empty();
    // テーブルIDと最大カラム数を保持
    let maxCols = new Object();

    structure.forEach((element, index) => {
        // シートの繰り返し
        if (element.children) {
            let tableId = `sheet-${index + 1}`;
            const table = $('<table>', { border: 1, id: tableId });
            if (index != 0) {
                table.hide();
            }
            const tableBody = $('<tbody>');
            let maxCol = 0;
            element.children.forEach(element1 => {
                // 各試験項目の繰り返し
                if (element1.children) {
                    let inputTrTag = $('<tr>');
                    let imgTrTag = $('<tr>');
                    element1.children.forEach((element2, index2) => {
                        if (index2 == 0) {
                            // テーブルの体裁を整えるために必要
                            // ダミー記入行
                            let inputTdTag = $('<td>');
                            inputTrTag.append(inputTdTag);

                            // 試験項目入力
                            let imgTdTag = $('<td>');
                            imgTdTag.text(element1.name);
                            imgTrTag.append(imgTdTag);
                        }

                        // 記入行
                        let inputTag = $('<input>');
                        inputTag.attr('type', 'text');
                        inputTag.attr('value', '');
                        let inputTdTag = $('<td>');
                        inputTdTag.append(inputTag);
                        inputTrTag.append(inputTdTag);

                        // 画像ファイルの繰り返し
                        let imgTag = $('<img>');
                        imgTag.attr('src', element2.img_base64);
                        let imgTdTag = $('<td>');
                        imgTdTag.append(imgTag);
                        imgTrTag.append(imgTdTag);

                    });
                    tableBody.append(inputTrTag);
                    tableBody.append(imgTrTag);
                    maxCol = maxCol < element1.children.length ? element1.children.length : maxCol;
                }
            });
            table.append(tableBody);
            tableWrap.append(table);
            tableWrap.css('height', '750px');
            createSheetTab(element.name, index);
            maxCols[tableId] = maxCol;
        }
    });
    // tableのwidhtを算出して設定
    Object.entries(maxCols).forEach(([key, value]) => {
        $(`#${key}`).css('width', 200 * (value + 1) + 140);
    });
}

/**
 * 新しいシートタブとヘッダーフォームを生成
 * @param {String} sheetName 生成するシートの名前
 * @param {Number} index シートのインデックス 
 */
function createSheetTab(sheetName, index) {
    let id = index + 1;
    // ヘッダーシートタブ生成
    let headerSheet = $('<div>');
    headerSheet.attr('id', `header-sheet-${id}`);
    if (index != 0) {
        headerSheet.hide();
    }

    function createInputField(labelText, inputId) {
        return $('<div>', { class: 'header-form' })
            .append($('<label>', { for: inputId, text: labelText }))
            .append($('<input>', { type: 'text', id: inputId }));
    }
    // 実施日時
    headerSheet.append(createInputField('実施日時', `implementation-date-input-${id}`));
    // OSバージョン
    headerSheet.append(createInputField('OSバージョン', `os-input-${id}`));
    // その他
    headerSheet.append(createInputField('その他', `others-input-${id}`));
    headerSheetTabDiv.append(headerSheet);

    // シートタブ生成
    let sheetBtn = $('<button>', { 'data-sheet': `sheet-${index + 1}` });
    sheetBtn.text(sheetName);
    // 最初のシートをアクティブに
    if (index == 0) {
        sheetBtn.attr('class', 'active');
    }
    sheetTabDiv.append(sheetBtn);
    sheetTabDiv.show();
}
