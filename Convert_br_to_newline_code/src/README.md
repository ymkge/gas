# Googleスプレッドシート <br>タグ改行変換スクリプト

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 概要

このリポジトリは、Googleスプレッドシート上で、セル内に含まれるHTMLの改行タグ (`<br>`) をセル内改行に一括で変換するためのGoogle Apps Script（GAS）です。

生成AIの回答やウェブサイトからコピーしたテキストをスプレッドシートに貼り付けた際、改行が`<br>`タグとして表示されてしまうことがあります。このスクリプトを使えば、シートのメニューからワンクリックで、シート全体の`<br>`タグを適切なセル内改行に変換できます。

## ✨ 機能

-   **簡単な操作**: スプレッドシートのカスタムメニューからワンクリックで実行できます。
-   **シート全体を一括変換**: 現在アクティブなシートに含まれるすべての`<br>`タグを変換します。
-   **柔軟なタグ検知**: `<br>`, `<br/>`, `<br >`, `<BR>`など、大文字・小文字やスペースの有無、閉じタグの有無に関わらず、すべての改行タグを検知して変換します。
-   **高速処理**: `getValues()`と`setValues()`を利用して、大量のデータでも高速に処理を実行します。

## 🚀 使い方

1.  **スプレッドシートを開く**
    変換したいデータが入力されているGoogleスプレッドシートを開きます。

2.  **スクリプトエディタを開く**
    メニューバーから `拡張機能` > `Apps Script` を選択します。

3.  **コードを貼り付ける**
    以下の`Code.gs`の内容をすべてコピーし、スクリプトエディタに貼り付けます。もし最初から`function myFunction() { ... }`といったコードがあれば、それは削除してください。

4.  **プロジェクトを保存**
    エディタの上部にあるフロッピーディスクのアイコン（💾 プロジェクトを保存）をクリックして、スクリプトを保存します。

5.  **スプレッドシートを再読み込み**
    スクリプトを反映させるため、スプレッドシートを開いているブラウザのタブを再読み込み（リロード）してください。

6.  **スクリプトの実行**
    メニューバーに新しく `カスタムメニュー` が追加されています。
    -   `カスタムメニュー` > `BRタグを改行に変換する` をクリックします。
    -   処理が完了すると、確認のダイアログが表示されます。

7.  **初回実行時の承認**
    初めてスクリプトを実行する際には、Googleから承認を求めるポップアップが表示されます。画面の指示に従い、スクリプトがスプレッドシートのデータにアクセスすることを許可してください。（ご自身で作成したスクリプトのため、安全です。）

## 🔧 コード (`Code.gs`)

```javascript
/**
 * スプレッドシートを開いたときにカスタムメニューを追加します。
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
```

## ⚠️ 注意事項

-   このスクリプトによる変換処理は**元に戻すことができません**。実行する前に、必要であればシートのバックアップ（`ファイル` > `コピーを作成`）を取ることをお勧めします。

## 📜 ライセンス

このプロジェクトは[MITライセンス](LICENSE)の下で公開されています。

---