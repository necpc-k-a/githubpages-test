/**
 * 英字カウント
 */
const ALPHABET_COUNT = 26;

/**
 * Excelの列英字を数字に変換
 * @param {String} c 列を表す英字文字列 (例: "A", "AB", "ZZ")
 * @returns {Number} 列番号 (例: "A" -> 1, "AB" -> 28)
 */
function excelColumnToNumber(c) {
    return [...c].reduce((n, ch) => {
        return n * ALPHABET_COUNT + ch.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
    }, 0);
}

/**
 * 数字をExcelの列英字に変換
 * @param {Number} n 列番号 (1以上の整数)
 * @returns {String} Excelの列英字文字列 (例: 1 -> "A", 28 -> "AB")
 */
function numberToExcelColumn(n) {
    return n <= 0 ? '' : numberToExcelColumn(Math.floor((n - 1) / ALPHABET_COUNT)) + String.fromCharCode((n - 1) % ALPHABET_COUNT + 'A'.charCodeAt(0));
}

/**
 * 画像ファイルをbase64に変換
 * @param {File} file 変換する画像ファイル
 * @returns {Promise<string>} 変換されたbase64形式の画像データ
 */
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = (error) => reject(error);
        reader.readAsDataURL(file);
    });
}

/**
 * Excelファイルを読み込み、ワークブックを生成
 * @param {Blob} blob ExcelファイルのBlobオブジェクト 
 */
async function loadExcel2(blob) {
    const arraBuffer = await blobToArrayBuffer(blob);
    // ExcelJSのワークブックを生成してExcelファイルをロード
    workbook2 = new ExcelJS.Workbook();
    await workbook2.xlsx.load(arraBuffer);
}

/**
 * Excelファイルを読み込み、ワークブックを生成
 * @param {Blob} blob ExcelファイルのBlobオブジェクト
 * @returns {Promise<ExcelJS.Workbook>} ExcelJSのワークブックオブジェクト
 */
async function loadExcel(blob) {
    const arraBuffer = await blobToArrayBuffer(blob);
    // ExcelJSのワークブックを生成してExcelファイルをロード
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arraBuffer);
    return workbook;
}

/**
 * BlobオブジェクトをArrayBufferに変換
 * @param {Blob} blob 変換するBlobオブジェクト
 * @returns {Promise<ArrayBuffer>} 変換されたArrayBuffer
 */
function blobToArrayBuffer(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(reader.error);
        reader.readAsArrayBuffer(blob);
    });
}

/**
 * 適切なフォルダ階層かの判定用フラグ\
 * 有効ならtrue, 無効ならfalse
 */
let isFolderStructureValid = false;

/**
 * エビデンスフォルダ階層が適切ではない場合のエラーメッセージ
 */
const errorFolderStructureMessage = "適切なフォルダ階層ではない";


/**
 * ディレクトリからファイルとサブディレクトリの構造を再帰的に読み取る
 * @param {FileSystemDirectoryHandle} directoryHandle 読み取るディレクトリのハンドル
 * @param {String} path 現在のパス
 * @param {Number} count 再起呼び出しの深さカウント
 * @returns {Promise<Array} ディレクトリ/ファイル構造の配列
 * @throws {Error} 構造が不正な場合にエラーをスロー
 */
async function readDirectory(directoryHandle, path = '', count = 0) {
    count++;
    if (count > 3) {
        throw new Error(errorFolderStructureMessage);
    }
    // 選択したディレクトリ内のフォルダ/ファイルをソート
    const entries = [];
    for await (const [name, handle] of directoryHandle.entries()) {
        if (!name.startsWith('.')) {
            // 先頭が「.」で始まるファイルは飛ばす
            entries.push([name, handle]);
        }
    }

    // フォルダ/ファイル名で昇順ソート
    entries.sort((a, b) => a[0].localeCompare(b[0]));

    const directoryStructure = [];
    for (const [name, handle] of entries) {
        if (handle.kind === 'file') {
            if (count <= 2) {
                // ファイルがあってはならないフォルダなので処理終了
                throw new Error(errorFolderStructureMessage);
            }
            // ファイル処理
            const file = await handle.getFile();
            imgBase64 = await fileToBase64(file);
            directoryStructure.push({
                name: name,
                type: 'file',
                img_base64: imgBase64
            });
            // 処理がここまで来たらフォルダ階層は問題ないので
            isFolderStructureValid = true;
            continue;
        }

        if (handle.kind === 'directory') {
            if (count == 3) {
                // フォルダがあってはならないフォルダなので処理終了
                throw new Error(errorFolderStructureMessage);
            }
            // ディレクトリ処理
            const subDirectoryStructure = await readDirectory(handle, `${path}${name}/`, count);
            directoryStructure.push({
                name: name,
                type: 'directory',
                children: subDirectoryStructure
            });
            continue;
        }
    }
    return isFolderStructureValid ? directoryStructure : new Array();
}
