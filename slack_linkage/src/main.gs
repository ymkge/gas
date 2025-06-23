// ====== 設定項目 ======
// 監視対象のスプレッドシート名を設定してください
const SHEET_NAME = '案件管理台帳'; // 例: '案件依頼台帳'
// 通知ステータスを記録する列（A=1, B=2, C=3, D=4, E=5）
const STATUS_COLUMN_INDEX = 5; 
// 通知完了後にステータス列に書き込むテキスト
const NOTIFIED_TEXT = '通知済み';
// ====================

/**
 * スプレッドシートが編集されたときに実行される関数（トリガーで設定）
 * @param {Object} e - イベントオブジェクト
 */
function checkAndNotifySlack(e) {
  try {
    const range = e.range; // 編集されたセル範囲を取得
    const sheet = range.getSheet(); // 編集されたシートを取得

    // 指定したシート以外での編集は無視する
    if (sheet.getName() !== SHEET_NAME) {
      return;
    }

    const editedRow = range.getRow();
    const editedCol = range.getColumn();

    // ヘッダー行(1行目)や、通知ステータス列自身の編集は無視する
    if (editedRow === 1 || editedCol === STATUS_COLUMN_INDEX) {
      return;
    }

    // A列からD列（依頼内容）以外の編集は無視する
    if (editedCol > 4) {
      return;
    }

    // 編集された行のA列からステータス列までのデータを取得
    const rowData = sheet.getRange(editedRow, 1, 1, STATUS_COLUMN_INDEX).getValues()[0];
    
    // ステータス列の値を取得
    const notificationStatus = rowData[STATUS_COLUMN_INDEX - 1];

    // 既に通知済みの場合は何もしない
    if (notificationStatus === NOTIFIED_TEXT) {
      return;
    }

    // A, B, C, D列の値を取得
    const entryDateValue = rowData[0];
    const summary = rowData[1];
    const requester = rowData[2];
    const desiredDateValue = rowData[3];

    // A, B, C, D列のいずれかが空欄の場合は処理を中断
    if (entryDateValue === '' || summary === '' || requester === '' || desiredDateValue === '') {
      return;
    }

    // すべてのセルが埋まり、かつ未通知の場合のみ、以下の処理を実行
    
    // スクリプトプロパティからWebhook URLを取得
    const slackWebhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
    if (!slackWebhookUrl) {
      Logger.log('エラー: SlackのWebhook URLがスクリプトプロパティに設定されていません。');
      return;
    }
    
    // 日付をフォーマット
    const entryDate = formatDate(entryDateValue);
    const desiredDate = formatDate(desiredDateValue);

    // Slackに送信するメッセージを作成
    const message = `案件依頼台帳に以下のエントリーがされました
- エントリー日付: ${entryDate}
- 依頼概要: ${summary}
- 依頼者: ${requester}
- 対応希望日: ${desiredDate}`;

    // Slackに送信するためのペイロードを作成
    const payload = {
      text: message,
    };

    // SlackにPOSTリクエストを送信
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
    };
    
    UrlFetchApp.fetch(slackWebhookUrl, options);

    // 通知が成功したら、ステータス列に「通知済み」と書き込む
    sheet.getRange(editedRow, STATUS_COLUMN_INDEX).setValue(NOTIFIED_TEXT);

  } catch (error) {
    // エラーが発生した場合にログを出力
    Logger.log('エラーが発生しました: ' + error.toString());
  }
}

/**
 * 日付オブジェクトを 'yyyy/MM/dd' 形式の文字列に変換するヘルパー関数
 * @param {Date|string} dateValue - 日付データ
 * @return {string} - フォーマットされた日付文字列
 */
function formatDate(dateValue) {
  if (dateValue instanceof Date) {
    // タイムゾーンは日本の'JST'を指定
    return Utilities.formatDate(dateValue, 'JST', 'yyyy/MM/dd');
  }
  return dateValue; // 日付オブジェクトでない場合はそのまま返す
}