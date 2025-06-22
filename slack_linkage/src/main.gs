// ====== 設定項目 ======
// 監視対象のスプレッドシート名を設定してください
const SHEET_NAME = '案件管理台帳'; 
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

    // 編集された列がA列からD列（1-4）の範囲外なら無視する
    const editedCol = range.getColumn();
    if (editedCol > 4) {
      return;
    }

    // 編集された行番号を取得
    const editedRow = range.getRow();
    // ヘッダー行（1行目）の編集は無視する
    if (editedRow === 1) {
      return;
    }

    // 編集された行のA列からD列の値を取得
    const rowValues = sheet.getRange(editedRow, 1, 1, 4).getValues()[0];

    // A, B, C, D列のいずれかが空欄の場合は処理を中断
    if (rowValues.some(cell => cell === '')) {
      return;
    }

    // すべてのセルが埋まった場合のみ、以下の処理を実行
    
    // スクリプトプロパティからWebhook URLを取得
    const slackWebhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
    if (!slackWebhookUrl) {
      Logger.log('エラー: SlackのWebhook URLがスクリプトプロパティに設定されていません。');
      return;
    }

    // データを取得して整形
    const entryDate = formatDate(rowValues[0]); // A列: エントリー日付
    const summary = rowValues[1];             // B列: 依頼概要
    const requester = rowValues[2];           // C列: 依頼者
    const desiredDate = formatDate(rowValues[3]); // D列: 対応希望日

    // Slackに送信するメッセージを作成
    const message = `案件依頼台帳に以下のエントリーがされました
- エントリー日付: ${entryDate}
- 依頼概要: ${summary}
- 依頼者: ${requester}
- 対応希望日: ${desiredDate}`;

    // Slackに送信するためのペイロード（データ）を作成
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
    return Utilities.formatDate(dateValue, 'JST', 'yyyy/MM/dd');
  }
  return dateValue; // 日付オブジェクトでない場合はそのまま返す
}