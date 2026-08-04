// 条件
// トリガー：毎朝8時～9時に実行
// 本日の日時から45日後迄の期間データをチェック
// 個品番号（YRL管理番号）と返却予定日を変数内に格納しスプレッドシートへ記録
// スプレッドシート内にある「個品番号＋返却予定日」と比較して重複したデータは追加しない
// スプレッドシートに新しいデータが追加されたらGoogle ChatへAllで通知

function fetchAndWriteContractData() {
  // ▼▼ 設定項目 ▼▼
  var SEARCH_DAYS_RANGE = 45; // 本日から何日後まで検索するか
  // ▲▲ 設定項目 ▲▲

  // --- プロパティの取得 ---
  var props = PropertiesService.getScriptProperties();
  var SPREADSHEET_ID = props.getProperty('SPREADSHEET_ID');
  var API_KEY = props.getProperty('API_KEY');
  var API_SECRET_KEY = props.getProperty('API_SECRET_KEY');
  
  if (!SPREADSHEET_ID || !API_KEY || !API_SECRET_KEY) {
    Logger.log("❌ Error: スクリプトプロパティ(SPREADSHEET_ID / API_KEY / API_SECRET_KEY)が不足しています。");
    return;
  }

  // --- スプレッドシート準備 ---
  var spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch(e) {
    Logger.log("❌ Error: スプレッドシートが開けません。ID確認: " + SPREADSHEET_ID);
    return;
  }
  
  var sheet = spreadsheet.getSheetByName("PC等レンタル返却管理");
  if (!sheet) {
    Logger.log("❌ Error: シート「PC等レンタル返却管理」が見つかりません。");
    return;
  }

  // =========================================================================
  // 変更点①：既存データの読み込み（管理番号 ＋ 終了予定日で重複チェック）
  // =========================================================================
  var existingContracts = new Set();
  var lastRow = sheet.getLastRow();
  
  if (lastRow > 1) { 
    // B列(管理番号)からC列(終了予定日)までの2列分を取得
    var data = sheet.getRange(2, 2, lastRow - 1, 2).getValues(); 
    
    data.forEach(function(row) {
      if (row[0] !== "" && row[0] !== null) {
        var khnoStr = String(row[0]);
        var rtodRaw = row[1];
        var rtodStr = "";
        
        // スプレッドシートの日付を "YYYY-MM-DD" 形式の文字列に統一
        if (rtodRaw instanceof Date) {
          rtodStr = Utilities.formatDate(rtodRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          rtodStr = String(rtodRaw).replace(/\//g, '-'); // スラッシュがあればハイフンに変換
        }
        
        // 「管理番号_日付」という複合キーで保存（例: "123456_2024-05-01"）
        existingContracts.add(khnoStr + "_" + rtodStr); 
      }
    });
  }

  appendDebugLog(spreadsheet, "run-start", "SEARCH_DAYS_RANGE=" + SEARCH_DAYS_RANGE + ", sheet=" + sheet.getName());

  // --- Step 1: API SignatureとSIDを取得 ---
  Logger.log("Step 1: API認証を開始します...");
  var authData = getAPISignatureAndSID(API_KEY, API_SECRET_KEY);
  if (!authData.apiSignature || !authData.sid) {
    Logger.log("❌ Stop: API認証失敗");
    return;
  }
  
  // --- Step 2: 全ページの契約データを取得 ---
  Logger.log("Step 2: 契約データを全ページ分取得します（期間: " + SEARCH_DAYS_RANGE + "日後まで）");
  
  var allContracts = [];
  var page = 1;
  var hasNextPage = true;

  while (hasNextPage) {
    var pageList = getContractList(API_KEY, authData.apiSignature, authData.sid, page, SEARCH_DAYS_RANGE);
    
    if (pageList && pageList.length > 0) {
      Logger.log("Page " + page + ": " + pageList.length + "件取得");
      allContracts = allContracts.concat(pageList);
      page++;
      
      if (pageList.length < 5) { 
        hasNextPage = false; 
      }
      if (page > 50) {
        hasNextPage = false;
        Logger.log("⚠️ ページ数が多すぎるため50ページで中断します");
      }
    } else {
      hasNextPage = false;
    }
    
    Utilities.sleep(500); 
  }
  
  Logger.log("データ取得完了。合計件数: " + allContracts.length + "件");

  // --- 書き込み処理 ---
  var newContractsCount = 0;
  var isNewDataAdded = false;

  allContracts.forEach(function(contract) {
    // =========================================================================
    // 変更点②：APIから取得したデータも「管理番号 ＋ 終了予定日」でキーを作成
    // =========================================================================
    var apiRtodStr = String(contract.rtod).replace(/\//g, '-');
    var checkKey = String(contract.khno) + "_" + apiRtodStr; 

    // 既存リストに「管理番号_日付」が存在しない場合のみ追加
    if (!existingContracts.has(checkKey)) {
      lastRow = sheet.getLastRow();
      
      sheet.insertRowAfter(lastRow);
      var targetRow = lastRow + 1;
      
      sheet.getRange(targetRow, 1).setNumberFormat("@").setValue(contract.jkno);
      sheet.getRange(targetRow, 2).setValue(contract.khno); 
      sheet.getRange(targetRow, 3).setValue(contract.rtod); 
      sheet.getRange(targetRow, 4).setValue(contract.kmrk); 
      sheet.getRange(targetRow, 5).setValue(contract.khnm); 
      sheet.getRange(targetRow, 6).setValue(contract.srno); 
      sheet.getRange(targetRow, 7).setValue(contract.statics_name_s); 
      
      existingContracts.add(checkKey); 
      isNewDataAdded = true;
      newContractsCount++;
      
      Logger.log("新規追加: " + contract.khno + " / " + contract.rtod);
    }
  });

  // --- 通知 ---
  Logger.log("通知判定: 取得件数=" + allContracts.length + ", 新規追加件数=" + newContractsCount);
  appendDebugLog(spreadsheet, "run-summary", "fetched=" + allContracts.length + ", new=" + newContractsCount);
  if (isNewDataAdded) {
    sendNotification(newContractsCount, SPREADSHEET_ID);
  } else {
    Logger.log("ℹ️ 新規データはありませんでした（通知はスキップされました）");
  }
}

function appendDebugLog(spreadsheet, label, detail) {
  try {
    var debugSheet = spreadsheet.getSheetByName("debug_log");
    if (!debugSheet) {
      debugSheet = spreadsheet.insertSheet("debug_log");
    }
    debugSheet.appendRow([new Date(), label, String(detail)]);
  } catch (e) {
    Logger.log("Debug log write failed: " + e);
  }
}

// Step 1: 認証 (GET URL結合版)
function getAPISignatureAndSID(apiKey, apiSecretKey) {
  var baseUrl = "https://wrt.simplit.jp/direct/member/generate_api_signature/";
  var step1Url = baseUrl + "?api_key=" + encodeURIComponent(apiKey) + "&api_secret_key=" + encodeURIComponent(apiSecretKey);

  var step1Params = {
    method: "GET",
    muteHttpExceptions: true
  };
  
  try {
    var step1Response = UrlFetchApp.fetch(step1Url, step1Params);
    var step1Data = JSON.parse(step1Response.getContentText());
    if (step1Data.status != "1") {
      Logger.log("API Error (Step 1): " + step1Data.message);
      return { apiSignature: null, sid: null };
    }
    return { apiSignature: step1Data.api_signature, sid: step1Data.sid };
  } catch (e) {
    Logger.log("Exception (Step 1): " + e);
    return { apiSignature: null, sid: null };
  }
}

// Step 2: データ取得
function getContractList(apiKey, apiSignature, sid, pageID, searchDaysRange) {
  var step2Url = "https://wrt.simplit.jp/management/slm/slm_contract_list_api/";
  
  var now = new Date();
  var dateEnd = new Date(now.getTime() + searchDaysRange * 24 * 60 * 60 * 1000);
  
  var step2Params = {
    method: "POST",
    muteHttpExceptions: true,
    payload: {
      "api_key": apiKey,
      "api_signature": apiSignature,
      "sid": sid,
      "pageID": pageID,
      "search[rtod1]": Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      "search[rtod2]": Utilities.formatDate(dateEnd, Session.getScriptTimeZone(), "yyyy-MM-dd")
    }
  };
  
  try {
    var step2Response = UrlFetchApp.fetch(step2Url, step2Params);
    var step2Data = JSON.parse(step2Response.getContentText());
    var responseStatus = step2Data.status;
    var contractList = step2Data.contract_list || step2Data.contractList || step2Data.list || (step2Data.data ? step2Data.data.contract_list : null);

    Logger.log("Step 2 response status: " + responseStatus + ", keys=" + (step2Data ? Object.keys(step2Data).join(",") : "none"));
    if (responseStatus != 1) {
      Logger.log("API Error (Step 2) Page " + pageID + ": " + responseStatus + " / " + JSON.stringify(step2Data));
      return null;
    }
    if (!contractList) {
      Logger.log("Step 2 returned no contract list. Raw response: " + JSON.stringify(step2Data));
      return [];
    }
    return contractList;
  } catch (e) {
    Logger.log("Exception (Step 2): " + e);
    return null;
  }
}

// 通知関数（ログ出力・エラーチェックの強化版）
function sendNotification(newContractsCount, spreadsheetId) {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
  if (!webhookUrl) {
    Logger.log("❌ Error: CHAT_WEBHOOK_URL がスクリプトプロパティに設定されていません。");
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
    contentType: "application/json; charset=UTF-8",
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(webhookUrl, options);
    var code = response.getResponseCode();
    var body = response.getContentText();
    if (code === 200) {
      Logger.log("✅ Google Chat通知送信成功: " + newContractsCount + "件");
    } else {
      Logger.log("❌ Google Chat通知失敗 (HTTP " + code + "): " + body);
    }
  } catch(e) {
    Logger.log("❌ Google Chat通信エラー: " + e.toString());
  }
}

// 【テスト用関数】通知単体テスト
function testGoogleChatNotification() {
  var props = PropertiesService.getScriptProperties();
  var SPREADSHEET_ID = props.getProperty('SPREADSHEET_ID');
  Logger.log("--- Google Chat 通知テスト開始 ---");
  sendNotification(99, SPREADSHEET_ID);
}