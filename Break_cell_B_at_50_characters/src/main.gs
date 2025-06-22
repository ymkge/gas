/**
 * 指定されたテキストを、指定された文字数ごとに改行して返します。
 * @param {string} text - 改行を適用する元のテキスト。
 * @param {number} wrapLength - 改行を挿入する文字数。
 * @returns {string} 改行が挿入されたテキスト。
 */
function wrapTextByLength(text, wrapLength) {
  if (!text) {
    return "";
  }
  let result = "";
  for (let i = 0; i < text.length; i += wrapLength) {
    result += text.substring(i, i + wrapLength) + "\n";
  }
  return result.trim(); // 最後の余分な改行を削除
}

/**
 * 指定されたシートのB列に対し、50文字ごとに改行を適用し、
 * セルの折り返し表示を有効にします。
 * @param {string} sheetName - 操作対象のシート名。
 */
function applyWrapToColumnB(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Browser.msgBox("エラー", "シート '" + sheetName + "' が見つかりません。", Browser.Buttons.OK);
    return;
  }
  // B列のデータを取得（ヘッダー行を考慮しない場合、1行目から取得）
  // 使用しているスプレッドシートに合わせて、開始行と終了行を調整してください。
  // 例: 1行目がヘッダーで2行目からデータが始まる場合、getRange("B2:B" + sheet.getLastRow())
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) { // 行がない場合は処理しない
    Logger.log("シートにデータがありません。");
    return;
  }

  const range = sheet.getRange("B1:B" + lastRow); // B1から最終行までのB列全体
  const values = range.getValues(); // 範囲内の全てのセルの値を取得

  const newValues = [];
  const wrapLength = 50; // ここで改行文字数を50に設定

  for (let i = 0; i < values.length; i++) {
    const originalText = values[i][0]; // B列の値は配列の0番目
    const wrappedText = wrapTextByLength(String(originalText), wrapLength);
    newValues.push([wrappedText]); // 2次元配列としてプッシュ
  }

  // 改行されたテキストをB列に一括で設定
  range.setValues(newValues);

  // B列の全てのセルの折り返し表示を有効にする
  // これは列全体に適用されるので、一度設定すればOKです。
  range.setWrap(true);

  Browser.msgBox(
    "処理完了",
    "シート '" + sheetName + "' のB列のテキストを50文字ごとに改行しました。",
    Browser.Buttons.OK
  );
}

// メニューに追加するための関数 (スプレッドシートを開いたときに実行される)
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("カスタムメニュー")
    .addItem("B列を50文字で改行", "showDialog")
    .addToUi();
}

// シート名を入力するためのダイアログを表示する関数
function showDialog() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "シート名の入力",
    "処理を実行するシートの名前を入力してください:",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const sheetName = result.getResponseText();
    applyWrapToColumnB(sheetName);
  }
}