/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタムメニュー') // メニュー名を「カスタムメニュー」に設定
    .addItem('BRタグを改行に変換する', 'convertBrToNewlineInSheet') // メニュー項目を追加
    .addToUi();
}

/**
 * 現在アクティブなシート全体を対象に、<br>タグを改行文字(\n)に変換します。
 * この関数がメニューから呼び出されます。
 */
function convertBrToNewlineInSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange(); // データが存在するすべての範囲を取得
  
  // 処理するセルがない場合はここで終了
  if (range.getNumRows() === 0 || range.getNumColumns() === 0) {
    SpreadsheetApp.getUi().alert('シートにデータがありません。');
    return;
  }
  
  replaceBrTagsInRange(range);
  SpreadsheetApp.getUi().alert('シート全体の<br>タグを改行に変換しました。');
}

/**
 * 指定された範囲内のセルの<br>タグを改行に置換する共通関数
 * @param {GoogleAppsScript.Spreadsheet.Range} range 処理対象の範囲
 */
function replaceBrTagsInRange(range) {
  // 正規表現を使用して、<br>, <br/>, <BR> など大文字小文字やスペースの有無に関わらずマッチさせる
  const brRegex = /<br\s*\/?>/gi;
  
  const values = range.getValues();
  
  const newValues = values.map(row => {
    return row.map(cell => {
      if (typeof cell === 'string') {
        return cell.replace(brRegex, '\n');
      }
      return cell;
    });
  });
  
  range.setValues(newValues);
}