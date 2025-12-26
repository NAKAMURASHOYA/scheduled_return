// 条件
// トリガー：毎朝8時～9時に実行
// 本日の日時から45日後迄の期間データをチェック
// 個品番号（YRL管理番号）を変数内に格納しスプレッドシートへ記録
// スプレッドシート内にある個品番号データと比較して重複したデータは追加しない
// スプレッドシートに新しいデータが追加されたらGoogle Chat（【通知用】ITヘルプデスク対応依頼）へAllで通知

function fetchAndWriteContractData() {
  // --- プロパティの取得 ---
  var props = PropertiesService.getScriptProperties();
  var SPREADSHEET_ID = props.getProperty('SPREADSHEET_ID');
  var API_KEY = props.getProperty('API_KEY');
  var API_SECRET_KEY = props.getProperty('API_SECRET_KEY');
  
  // プロパティチェック
  if (!SPREADSHEET_ID || !API_KEY || !API_SECRET_KEY) {
    Logger.log("❌ Error: スクリプトプロパティが設定されていません。プロジェクトの設定を確認してください。");
    return;
  }

  // --- スプレッドシート準備 ---
  var spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch(e) {
    Logger.log("❌ Error: スプレッドシートが開けません。IDが正しいか確認してください。 ID: " + SPREADSHEET_ID);
    return;
  }
  
  var sheet = spreadsheet.getSheetByName("PC等レンタル返却管理");
  if (!sheet) {
    Logger.log("❌ Error: シート「PC等レンタル返却管理」が見つかりません。");
    return;
  }

  // 既存データの読み込み（重複チェック用）
  var existingContracts = new Set();
  var lastRow = sheet.getLastRow();
  
  if (lastRow > 1) { // データがある場合のみ
    var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B列(YRL管理番号)のみ取得して高速化
    data.forEach(function(row) {
      existingContracts.add(String(row[0])); // 文字列として保存
    });
  }

  // --- Step 1: API SignatureとSIDを取得 ---
  Logger.log("Step 1: API認証を開始します...");
  var authData = getAPISignatureAndSID(API_KEY, API_SECRET_KEY);
  if (!authData.apiSignature || !authData.sid) {
    Logger.log("❌ Stop: API認証に失敗したため終了します。");
    return;
  }
  Logger.log("Step 1 OK: SID取得成功");

  // --- Step 2: レンタル契約情報の取得 ---
  Logger.log("Step 2: 契約データを取得します...");
  var contractList = getContractList(API_KEY, authData.apiSignature, authData.sid);
  if (!contractList) {
    Logger.log("❌ Stop: 契約データの取得に失敗したため終了します。");
    return;
  }
  Logger.log("Step 2 OK: 取得件数 " + contractList.length + "件");

  // --- 書き込み処理 ---
  var newContractsCount = 0;
  var isNewDataAdded = false;

  contractList.forEach(function(contract) {
    var checkKey = String(contract.khno); // 比較用に文字列化

    // 重複チェック
    if (!existingContracts.has(checkKey)) {
      
      // 最終行を再取得（ループ内でinsertRowするため）
      lastRow = sheet.getLastRow();
      
      // H列以降にデータがある場合のロジック（元コード維持）
      if (sheet.getRange(lastRow, 8).getValue() !== "") {
        lastRow++;
      }
      
      // 行を追加して書き込み
      sheet.insertRowAfter(lastRow);
      var targetRow = lastRow + 1;
      
      sheet.getRange(targetRow, 1).setValue(contract.jkno); // A: 契約番号
      sheet.getRange(targetRow, 2).setValue(contract.khno); // B: YRL管理番号
      sheet.getRange(targetRow, 3).setValue(contract.rtod); // C: レンタル終了予定日
      sheet.getRange(targetRow, 4).setValue(contract.kmrk); // D: メーカー略称
      sheet.getRange(targetRow, 5).setValue(contract.khnm); // E: 品名
      sheet.getRange(targetRow, 6).setValue(contract.srno); // F: シリアル番号
      sheet.getRange(targetRow, 7).setValue(contract.statics_name_s); // G: 分類
      
      existingContracts.add(checkKey); // 同一実行内での重複防止
      isNewDataAdded = true;
      newContractsCount++;
      
      Logger.log("新規追加: " + contract.khno);
    }
  });

  // --- 通知 ---
  if (isNewDataAdded) {
    sendNotification(newContractsCount, SPREADSHEET_ID);
    Logger.log("通知を送信しました。新規件数: " + newContractsCount);
  } else {
    Logger.log("新規データはありませんでした。");
  }
}

// ▼ 修正箇所: GETリクエストのパラメータをURL結合に変更
function getAPISignatureAndSID(apiKey, apiSecretKey) {
  var baseUrl = "http://wrt.simplit.jp/direct/member/generate_api_signature/";
  // クエリパラメータとして構築
  var step1Url = baseUrl + "?api_key=" + apiKey + "&api_secret_key=" + apiSecretKey;

  var step1Params = {
    method: "GET",
    muteHttpExceptions: true
    // payload は削除 (GETでは使えないため)
  };
  
  try {
    var step1Response = UrlFetchApp.fetch(step1Url, step1Params);
    var jsonText = step1Response.getContentText();
    var step1Data = JSON.parse(jsonText);
    
    if (step1Data.status != "1") {
      Logger.log("API Error (Step 1): Status " + step1Data.status + " / Msg: " + step1Data.message);
      return { apiSignature: null, sid: null };
    }
    return { apiSignature: step1Data.api_signature, sid: step1Data.sid };
  } catch (e) {
    Logger.log("Exception (Step 1): " + e);
    return { apiSignature: null, sid: null };
  }
}

function getContractList(apiKey, apiSignature, sid) {
  var step2Url = "https://wrt.simplit.jp/management/slm/slm_contract_list_api/";
  
  // 日付計算
  var now = new Date();
  var date45DaysAfter = new Date(now.getTime() + 45 * 24 * 60 * 60 * 1000);
  
  var step2Params = {
    method: "POST",
    muteHttpExceptions: true,
    payload: {
      "api_key": apiKey,
      "api_signature": apiSignature,
      "sid": sid,
      "pageID": 1,
      "search[rtod1]": Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      "search[rtod2]": Utilities.formatDate(date45DaysAfter, Session.getScriptTimeZone(), "yyyy-MM-dd")
    }
  };
  
  try {
    var step2Response = UrlFetchApp.fetch(step2Url, step2Params);
    var step2Data = JSON.parse(step2Response.getContentText());
    
    if (step2Data.status != 1) {
      Logger.log("API Error (Step 2): Status " + step2Data.status);
      return null;
    }
    return step2Data.contract_list;
  } catch (e) {
    Logger.log("Exception (Step 2): " + e);
    return null;
  }
}

function sendNotification(newContractsCount, spreadsheetId) {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
  
  if (!webhookUrl) {
    Logger.log("Warning: CHAT_WEBHOOK_URL が設定されていないため通知できません。");
    return;
  }

  var message = {
    text: "～レンタル返却管理～\n" +
          "<users/all>\n" +
          "新たに返却予定の情報が " + newContractsCount + " 件追加されました！\n\n" +
          "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit#gid=1906719251"
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message)
  };
  
  try {
    UrlFetchApp.fetch(webhookUrl, options);
  } catch(e) {
    Logger.log("Notification Error: " + e);
  }
}