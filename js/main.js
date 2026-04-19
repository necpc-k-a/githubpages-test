/**
 * 各シートのヘッダー要素
 */
const headerSheetTabDiv = $('#headerSheetWrap');

/**
 * ロード中のインジケータ
 */
const loadingDiv = $('#loading');
const loader = $('<div>', { id: 'loader', class: 'text-light spinner-border position-absolute top-50 z-3' });
loader.css({ 'width': '10em', 'height': '10em', 'border-width': '16px', 'border-right': '16px solid #3498db' });
loader.append($('<spab>', { class: 'visually-hidden' })).appendTo(loadingDiv).hide();

/**
 * ロード中のオーバーレイ背景
 */
const overlay = $('<div>', { id: 'overlay', class: 'bg-black bg-opacity-50 w-100 h-100 start-0 top-0 z-1 position-fixed' });
overlay.appendTo('body').hide();

/**
 * エビデンスフォルダ階層でのアラート表示エラーメッセージ
 */
const errorFolderStructureAlertMessage = "エビデンスフォルダジェネレーターを用いてフォルダ生成し、再度選択してください。";

/**
 * エビデンスファイル添付処理
 */
$('#attachmentEvidencefileBtn').on('click', async (event) => {
    const btn = $(event.currentTarget);
    btn.prop('disabled', true);

    // ロード中インジケータとオーバーレイ表示
    // loader.show();
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
        $('#dlExcel').attr('disabled', false);

        // TODO: テーブルのセルを入れ替える
        enableTableSwap();
        // TODO: ダミーセルの追加処理
        addTrEvenPlusBtn();
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
 * エビデンス添付テーブル要素
 */
const sheetTableWrap = $('#sheetTableWrap');

/**
 * シート用タブ要素
 */
const sheetTabWrap = $('#sheetTabWrap');

/**
 * フォルダ構造データを基にテーブルを動的に生成し、シートタブを生成
 * @param {Array} structure テーブルとシートの構造データ
 */
function createMyTable(structure) {

    // 既にテーブル要素が存在する場合削除
    headerSheetTabDiv.empty();
    sheetTableWrap.empty();
    sheetTabWrap.empty();

    // テーブルIDと最大カラム数を保持
    let maxCols = new Object();

    structure.forEach((element, index) => {
        // シートの繰り返し
        if (element.children) {
            let tableId = `sheet-${index + 1}`;
            // const table = $('<table>', { id: tableId, class: "table table-bordered" });
            // TODO: tableクラスを削除
            const table = $('<table>', { id: tableId, class: "" , border: "1"});
            // TODO: 確認用で1テーブルだけ
            if (index != 0) {
                return;
            }

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
                            // TODO: inputタグを入れる予定
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
                        // imgTag.attr('src', element2.img_base64);
                        imgTag.attr('src', `./img/img${index2 + 1}.png`);
                        imgTag.css('max-width', '200px').css('max-height', '200px').css('margin', '10px');
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
            sheetTableWrap.append(table);
            sheetTableWrap.css('height', '750px').css('max-height', '750px');
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
 * 新しいシートタブとヘッダーフォームを作成
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
            .append($('<label>', { for: inputId, text: labelText, class: 'form-label' }))
            .append($('<input>', { type: 'text', id: inputId, class: 'form-control' })
            );
    }
    // 実施日時
    headerSheet.append(createInputField('実施日時', `implementationDateInput-${id}`));
    // OSバージョン
    headerSheet.append(createInputField('OSバージョン', `osInput-${id}`));
    // その他
    headerSheet.append(createInputField('その他', `othersInput-${id}`));
    headerSheetTabDiv.append(headerSheet);

    // シートタブ生成
    let sheetBtn = $('<button>', { 'data-sheet': `sheet-${index + 1}` });
    sheetBtn.text(sheetName);
    // 最初のシートをアクティブに
    if (index == 0) {
        sheetBtn.attr('class', 'active');
    }
    sheetBtn.addClass('btn w-25')
    sheetTabWrap.append(sheetBtn);
    sheetTabWrap.show();
}

/**
 * エビデンス添付テーブル要素
 */
const tableWrap = $('#sheetTableWrap');

/**
 * シートタブ押下時の切り替え処理
 */
$('#sheetTabWrap').on('click', 'button', function () {
    const sheetName = $(this).data('sheet');
    // アクティブ状態切り替え
    $('#sheetTabWrap button').removeClass('active');
    $(this).addClass('active');

    // ヘッダー、テーブルの表示・非表示
    headerSheetTabDiv.children().hide().filter(`#header-${sheetName}`).show();
    tableWrap.children().hide().filter(`#${sheetName}`).show();
});

/**
 * Excelファイル書き込みと保存処理
 */
$('#dlExcel').on('click', async (event) => {
    const btn = $(event.currentTarget);
    btn.prop('disabled', true);

    // ロード中インジケータとオーバーレイ表示
    loader.show();
    overlay.show();

    try {
        const excelBlob = await createExcel();
        const workbook = await loadExcel(excelBlob);

        // 必要シート数
        const sheetNum = sheetTabWrap.children().length;
        const sheetNameArray = sheetTabWrap.children().map((_, elem) => $(elem).text()).get();
        // シートのコピー
        copySheet(workbook, sheetNum, sheetNameArray);
        // テンプレートシート削除
        const delTempSheet = workbook.getWorksheet(TEMPLATE_SHEET_NAME);
        workbook.removeWorksheet(delTempSheet.id);
        // excelファイルへの書き込み
        exportExamSpecToExcel(workbook);
        const modifiedBuffer = await workbook.xlsx.writeBuffer({
        });
        const blob = new Blob([modifiedBuffer.buffer], {
            type: 'application/octet-stream'
        });
        saveAs(blob, 'example.xlsx');
    } catch (error) {
        console.error(error);
    } finally {
        btn.prop('disabled', false);
        // ロード中インジケータとオーバーレイ非表示
        loader.hide();
        overlay.hide();
    }
});

/**
 * 指定されたワークブックのテンプレートシートを指定回数コピーし、各コピーシートに新しい名前を割り当てる
 * セル値、スタイル、結合セル、列幅、行高、ウィンドウ固定設定をコピー
 *
 * @param {ExcelJS.Workbook} workbook - 操作対象のワークブック
 * @param {number} copyNum - コピーするシートの数
 * @param {string[]} sheetNameArray - 新しいシート名の配列。`copyNum` と同数必要です
 * @returns {void} 渡された `workbook` オブジェクトを直接変更します
 */
function copySheet(workbook, copyNum, sheetNameArray) {
    copyNum = [...Array(copyNum).keys()];
    copyNum.forEach((elem, i) => {
        // テンプレートシート読み込み
        const originalSheet = workbook.getWorksheet(TEMPLATE_SHEET_NAME);
        const copiedSheet = workbook.addWorksheet(`${sheetNameArray[i]}`);

        // テンプレートシートの全セルのスタイル設定
        originalSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const newRow = copiedSheet.getRow(rowNumber);
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                const newCell = newRow.getCell(colNumber);
                newCell.value = cell.value;
                newCell.style = { ...cell.style };
                if (cell.type === ExcelJS.ValueType.RichText) {
                    newCell.value = { richText: cell.value.richText.map(part => ({ ...part })) };
                }
                if (cell.isMerged && cell.address === cell.master.address) {
                    const startColAlpha = cell.address.substring(0, 1);
                    const startColNum = cell.address.substring(2, 1);
                    let endColAlpha = excelColumnToNumber(startColAlpha) + cell._mergeCount;
                    endColAlpha = numberToExcelColumn(endColAlpha);
                    copiedSheet.mergeCells(`${startColAlpha}${startColNum}:${endColAlpha}${startColNum}`);
                }
            });
        });

        // 列の幅とスタイルの設定
        originalSheet.columns.forEach((col, index) => {
            const newCol = copiedSheet.getColumn(index + 1);
            newCol.width = col.width;
            newCol.style = { ...col.style };
        });

        // 行の高さの設定
        originalSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const newRow = copiedSheet.getRow(rowNumber);
            newRow.height = row.height;
        });

        // シートのウィンドウ固定の設定
        const sheetViews = originalSheet.views;
        if (sheetViews.length != 0) {
            const freezePane = sheetViews[0];
            copiedSheet.views[0] = freezePane;
        }
    });
}

/**
 * 試験仕様書データをExcelファイルとしてエクスポートします
 * DOMから必要な要素の値を読み込み、新しいExcelワークブックを作成し、
 * 各シートに試験仕様書のヘッダー情報とテーブルデータを挿入します
 * 画像データも埋め込み、最終的にExcelファイルをBlobとして保存します
 *
 * @returns {Promise<void>} 処理の完了を示すPromise。ファイル保存が成功したかどうかにかかわらず解決されます
 * @throws {Error} Excelファイル生成または保存中にエラーが発生した場合
 */
function exportExamSpecToExcel(workbook) {
    // ヘッダー要素
    // 試験仕様書名
    const examSpecName = $('#examSpecificationNameInput').val();
    // 各シートの実施日時、OSバージョン、その他
    const implementationDateInput = 'implementationDateInput-';
    const osInput = 'osInput-';
    const othersInput = 'othersInput-';

    // テーブルタブの繰り返し
    sheetTabWrap.children().each((index, elem) => {

        // 作業シート
        worksheet = workbook.getWorksheet($(elem).text());
        console.log(worksheet);
        // 試験仕様書名の入力
        worksheet.getCell('C1').value = examSpecName;
        // 各シートの要素を抽出するID
        let dataId = $(elem).data('sheet');
        console.log(dataId);
        // ヘッダー要素の設定
        let headerForm = $(`#header-${dataId}`);
        let headerIndex = index + 1;

        let headerFormInput = headerForm.find('input');
        console.log(headerForm);
        console.log(headerFormInput);
        console.log(headerFormInput.filter(`#${implementationDateInput}${headerIndex}`));
        console.log(headerFormInput.filter(`#${implementationDateInput}${headerIndex}`).val());
        worksheet.getCell('F1').value = headerFormInput.filter(`#${implementationDateInput}${headerIndex}`).val();
        worksheet.getCell('A2').value = headerFormInput.filter(`#${osInput}${headerIndex}`).val();
        worksheet.getCell('C2').value = headerFormInput.filter(`#${othersInput}${headerIndex}`).val();

        // テーブル要素
        let table = $(`#${dataId}`);
        let tableBody = $($(`#${dataId}`).children()[0]);
        let insertRow = tableBody.children().length / 2 - 1;
        let insertCol = Math.max(...table.find('tr').map((index, e) => {
            return $(e).children('td').length;
        }).get()) - 1;
        copyColumnWithStyle(worksheet, insertCol);
        copyRowWithStyle(worksheet, insertCol, insertRow);
        populateWorksheetWithTableData(workbook, worksheet, tableBody);
    });
}

/**
 * 指定されたExcelワークシート内で、特定の列のスタイルをコピーし、新しい列として挿入します。
 * 元の列（srcColumnAlpha）のスタイル（罫線、幅など）をコピーし、
 * 指定された開始列（startColumnAlpha）に新しい列として挿入します。
 * この処理はdestEndColNumで指定された回数だけ繰り返されます。
 *
 * @param {object} worksheet - スタイルをコピーするExcelJSのワークシートオブジェクト
 * @param {number} destEndColNum - 挿入する列の最終的なアルファベットに対応する数値（例: Excelの列数が7（G）なら7）
 * @param {string} srcColumnAlpha - コピー元の列のアルファベット（例: 'G'）
 * @param {string} startColumnAlpha - 新しい列を挿入する開始列のアルファベット（例: 'H'）
 * @param {number} startRowNum - スタイルをコピーする開始行番号
 * @param {number} endRowNum - スタイルをコピーする終了行番号
 * この値から6を引いた数が挿入する列数になります
 */
function copyColumnWithStyle(worksheet, destEndColNum, srcColumnAlpha = 'G', startColumnAlpha = 'H', startRowNum = 1, endRowNum = 6) {

    // 追加列数がH列以下なら追加しない
    destEndColNum = destEndColNum - 6;
    if (destEndColNum < 1) {
        return;
    }
    const cols = [...Array(destEndColNum).keys()];
    const rows = [...Array(endRowNum - startRowNum + 1).keys()];
    cols.forEach((_) => {
        // H列に新しい列を挿入
        worksheet.spliceColumns(excelColumnToNumber(startColumnAlpha), 0, []);
        // 行のスタイル設定
        rows.forEach((num, _) => {
            cellNum = num + 1;
            const srcCell = worksheet.getCell(`${srcColumnAlpha}${cellNum}`);
            const destCell = worksheet.getCell(`${startColumnAlpha}${cellNum}`);
            destCell.style = { ...srcCell.style };
            if (num < 2) {
                delete destCell.style.border;
            }
        });
        // 列の幅を設定
        const srcColumn = worksheet.getColumn(excelColumnToNumber(srcColumnAlpha));
        const destColumn = worksheet.getColumn(excelColumnToNumber(startColumnAlpha));
        destColumn.width = srcColumn.width;
    });
}

/**
 * 指定されたExcelワークシート内で、既存の行のスタイルをコピーし、新しい行として追加します
 * 入力行と画像行の高さとセルのスタイルをコピーして、指定された回数分だけ新しい行を作成します
 *
 * @param {object} worksheet - スタイルをコピーするExcelJSのワークシートオブジェクト
 * @param {number} lastCol - コピー対象となる最終列の数値（例: 'G'なら7）
 * @param {number} destEndRow - 挿入する最終行の数（元のデータ行数に基づいて、何行複製するか）
 * @param {number} [startRow=5] - スタイルコピー元の開始行番号。デフォルトは5
 */
function copyRowWithStyle(worksheet, lastCol, destEndRow, startRow = 5) {
    // 指定した行数分の高さを設定
    const srcInputRow = worksheet.getRow(startRow);
    const srcImageRow = worksheet.getRow(startRow + 1);

    const destStartRow = startRow + 2;
    const rows = [...Array(destEndRow * 2).keys()];
    // 行の高さを設定
    rows.forEach((index) => {
        const height = index % 2 == 0 ? srcInputRow.height : srcImageRow.height;
        const destRow = worksheet.getRow(index + destStartRow);
        destRow.height = height;
    });

    // 8はA〜H列の数
    const cols = [...Array(lastCol + 8).keys()].map(i => i + 1);
    cols.forEach((i) => {
        rows.forEach((j) => {
            const srcCell = worksheet.getCell(`${numberToExcelColumn(i)}${j % 2 == 0 ? 5 : 6}`);
            const destCell = worksheet.getCell(`${numberToExcelColumn(i)}${j + 7}`);
            destCell.style = { ...srcCell.style };
        });
    });
}

/**
 *
 * @param {} workbook -
 * @param {object} worksheet - データを書き込むExcelJSのワークシートオブジェクト
 * @param {jQuery} tableBody - データを取得するHTMLのtbody要素のjQueryオブジェクト
 */
function populateWorksheetWithTableData(workbook, worksheet, tableBody) {

    // A列の開始行
    let aColStartRow = 6;
    // 入力行
    let inputStartRow = 5;
    // 画像行列
    let imageStartRow = 4;
    let imageStartCol = 1;

    tableBody.children().each((i, e) => {
        // trの繰り返し
        $(e).children().each((j, e1) => {
            // tdの繰り返し
            if (i % 2 == 0 && j == 0) {
                // 何も記入しない行のため処理しない
                return;
            }
            if (j == 0) {
                // A列の記入内容
                let cell = worksheet.getCell(`A${aColStartRow}`);
                cell.value = $(e1).text();
                cell.alignment = {
                    vertical: 'middle',
                    wrapText: true
                };
                aColStartRow += 2;
                return;
            }
            if (i % 2 == 0) {
                // 記入行
                let inputTableCell = $(e1);
                let value = $(inputTableCell.children()[0]).val();
                console.log(`inputTableCell: ${value}`);
                let cell = worksheet.getCell(`${numberToExcelColumn(j + 1)}${inputStartRow}`);
                if (value != "") {
                    cell.value = value;
                }
                return;
            }
            // 画像行
            let imageTableCell = $(e1);
            let imageTag = $(imageTableCell.children()[0]);

            let sheetImg = workbook.addImage({
                base64: imageTag.attr('src'),
                extension: 'png'
            });

            worksheet.addImage(sheetImg, {
                tl: {
                    col: Number(imageStartCol + 0.2),
                    row: Number(imageStartRow + 0.5)
                },
                ext: {
                    width: 245,
                    height: 520
                }
            });
            imageStartCol++;
        });

        inputStartRow++;
        imageStartRow++;
        imageStartCol = 1;
    });
}

/**
 * Excelテンプレートダウンロード
 *
 * @param {MouseEvent} event - クリックイベントオブジェクト
 * @returns {Promise<void>} 処理が完了したときに解決するPromise
 */
$('#excelDl').on('click', async (event) => {
    // テンプレートエクセルファイルの作成
    const excelBlob = await createExcel();
    const workbook = await loadExcel(excelBlob);

    copySheet(workbook, 2, ['1_シート', '2_シート']);
    let worksheet = workbook.worksheets[0];

    let colNum = 0;
    copyColumnWithStyle(worksheet, colNum);
    let rowNum = 0;
    copyRowWithStyle(worksheet, colNum, rowNum);

    const execlBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([execlBuffer.buffer], { type: 'application/octet-stream' });
    saveAs(blob, 'template_excel.xlsx');
});


/**
 * tdセルのドラッグアンドドロップを行う処理
 */
function enableTableSwap() {
    const $table = $('table[id^="sheet-"]');
    $table.each(function() {
        // 偶数行と最初の列以外にドラッグ移動を行うクラスを付与
        $(this).find('tr:odd td:not(:first-child)')
            .addClass('drag-col')
            .attr('draggable', 'true');
    });

    let dragged;

    $table.on('dragstart', '.drag-col', function(e) {
        // e.originalEventでネイティブのイベントオブジェクトにアクセス
        dragged = this;
        setTimeout(() => $(this).css('opacity', 0.5), 0);
    });

    $table.on('dragend', '.drag-col', function(e) {
        $(this).css('opacity', '');
    });

    $table.on('dragover', '.drag-col', function(e) {
        e.preventDefault();
    });

    $table.on('drop', '.drag-col', function(e) {
        e.preventDefault();
        const dropTarget = this;

        if (dragged !== dropTarget) {
            // 要素の入れ替え
            const temp = $('<span>').hide(); // 一時的な要素を作成
            $(dropTarget).before(temp);
            $(dragged).before(dropTarget);
            temp.before(dragged);
            temp.remove();
        }
    });
}

/**
 * 奇数のtdセルに+ボタンを付与
 */
function addTrEvenPlusBtn() {
    const $table = $('table[id^="sheet-"]');

    $table.each(function() {
        $(this).find('tr:even').each(function() {
            const $newTd = $('<td>');
            $newTd.append($('<span>').addClass('btn btn-primary container-fluid').text("ダミーセル追加"));
            // $(this).append('<td class="btn">+</td>');
            $(this).append($newTd);
        });
    });

    $table.on('click', '.btn', function() {
        const currentCol = $(this).closest('td');
        console.log(currentCol);
        const nextRow = currentCol.parent().next();
        console.log(nextRow);

        currentCol.before($('<td>').text("test"));
        nextRow.append($('<td>').text("img"));
    });
}