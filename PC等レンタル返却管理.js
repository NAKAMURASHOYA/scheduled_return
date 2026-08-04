// 条件
// トリガー：毎朝8時～9時に実行
// 本日の日時から45日後迄の期間データをチェック
// 個品番号（YRL管理番号）と返却予定日を変数内に格納しスプレッドシートへ記録
// 新規：一番下の行に追加。延長（日付変更）：既存行を上書きし、M列に「延長分」と記載。
// 更新があった機器のYRL管理番号をGoogle Chatに一覧で通知する

function fetchAndWriteContractData() {
  // ▼▼ 設定項目 ▼▼
  var SEARCH_DAYS_RANGE = 45; // 本日から何日後まで検索するか
  // ▲▲ 設定項目 ▲▲

  var props = PropertiesService.getScriptProperties();
  var SPREADSHEET_ID = props.getProperty('SPREADSHEET_ID');
  var API_KEY = props.getProperty('API_KEY');
  var API_SECRET_KEY = props.getProperty('API_SECRET_KEY');
  
  if (!SPREADSHEET_ID || !API_KEY || !API_SECRET_KEY) {
    Logger.log("❌ Error: スクリプトプロパティが不足しています。");
    return;
  }

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
  // 変更点A：既存データの読み込み（行番号も一緒に記憶する）
  // =========================================================================
  var existingData = {}; // { "YRL番号": { row: 行番号, rtod: "日付" } }
  var lastRow = sheet.getLastRow();
  
  if (lastRow > 1) { 
    var data = sheet.getRange(2, 2, lastRow - 1, 2).getValues(); 
    
    for (var i = 0; i < data.length; i++) {
      var rowObj = data[i];
      if (rowObj[0] !== "" && rowObj[0] !== null) {
        var khnoStr = String(rowObj[0]);
        var rtodRaw = rowObj[1];
        var rtodStr = "";
        
        if (rtodRaw instanceof Date) {
          rtodStr = Utilities.formatDate(rtodRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          rtodStr = String(rtodRaw).replace(/\//g, '-');
        }
        
        // YRL番号をキーにして、そのデータが「何行目」にあるかと「日付」を保存
        existingData[khnoStr] = {
          row: i + 2, // 2行目から始まっているため +2
          rtod: rtodStr
        };
      }
    }
  }

  appendDebugLog(spreadsheet, "run-start", "SEARCH_DAYS_RANGE=" + SEARCH_DAYS_RANGE);

  // --- Step 1: API認証 ---
  Logger.log("Step 1: API認証を開始します...");
  var authData = getAPISignatureAndSID(API_KEY, API_SECRET_KEY);
  if (!authData.apiSignature || !authData.sid) {
    Logger.log("❌ Stop: API認証失敗");
    return;
  }
  
  // --- Step 2: データ取得 ---
  Logger.log("Step 2: 契約データを全ページ分取得します（期間: " + SEARCH_DAYS_RANGE + "日後まで）");
  var allContracts = [];
  var page = 1;
  var hasNextPage = true;

  while (hasNextPage) {
    var pageList = getContractList(API_KEY, authData.apiSignature, authData.sid, page, SEARCH_DAYS_RANGE);
    if (pageList && pageList.length > 0) {
      allContracts = allContracts.concat(pageList);
      page++;
      if (pageList.length < 5) hasNextPage = false; 
      if (page > 50) hasNextPage = false;
    } else {
      hasNextPage = false;
    }
    Utilities.sleep(500); 
  }
  
  // =========================================================================
  // 変更点B：書き込み処理 ＆ 通知用リストの作成
  // =========================================================================
  var newContractsCount = 0;
  var notifiedItems = []; // 通知に記載するYRL番号のリスト

  allContracts.forEach(function(contract) {
    var khnoStr = String(contract.khno);
    var apiRtodStr = String(contract.rtod).replace(/\//g, '-');

    if (!existingData[khnoStr]) {
      // パターン1：完全新規（スプレッドシートに番号が存在しない）
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
      
      // 次の重複を防ぐため記憶に追加
      existingData[khnoStr] = { row: targetRow, rtod: apiRtodStr };
      
      newContractsCount++;
      notifiedItems.push("・" + khnoStr + "（新規）");
      Logger.log("新規追加: " + khnoStr);

    } else if (existingData[khnoStr].rtod !== apiRtodStr) {
      // パターン2：延長（番号は存在するが、日付が変わっている）
      var updateRow = existingData[khnoStr].row;
      
      // A〜G列のみ最新データで上書き（I〜K列はそのまま残る）
      sheet.getRange(updateRow, 1).setNumberFormat("@").setValue(contract.jkno);
      sheet.getRange(updateRow, 2).setValue(contract.khno); 
      sheet.getRange(updateRow, 3).setValue(contract.rtod); 
      sheet.getRange(updateRow, 4).setValue(contract.kmrk); 
      sheet.getRange(updateRow, 5).setValue(contract.khnm); 
      sheet.getRange(updateRow, 6).setValue(contract.srno); 
      sheet.getRange(updateRow, 7).setValue(contract.statics_name_s); 
      
      // M列（13列目）に「延長分」と記載
      sheet.getRange(updateRow, 13).setValue("延長分");

      // 記憶の日付を更新
      existingData[khnoStr].rtod = apiRtodStr;

      newContractsCount++;
      notifiedItems.push("・" + khnoStr + "（延長分）");
      Logger.log("延長更新: " + khnoStr + " (行: " + updateRow + ")");
    }
  });

  // --- 通知 ---
  appendDebugLog(spreadsheet, "run-summary", "fetched=" + allContracts.length + ", new=" + newContractsCount);
  
  if (newContractsCount > 0) {
    sendNotification(newContractsCount, notifiedItems, SPREADSHEET_ID);
  } else {
    Logger.log("ℹ️ 新規・延長データはありませんでした（通知スキップ）");
  }
}

function appendDebugLog(spreadsheet, label, detail) {
  try {
    var debugSheet = spreadsheet.getSheetByName("debug_log");
    if (!debugSheet) debugSheet = spreadsheet.insertSheet("debug_log");
    debugSheet.appendRow([new Date(), label, String(detail)]);
  } catch (e) {
    Logger.log("Debug log write failed: " + e);
  }
}

// Step 1: 認証
function getAPISignatureAndSID(apiKey, apiSecretKey) {
  var baseUrl = "https://wrt.simplit.jp/direct/member/generate_api_signature/";
  var step1Url = baseUrl + "?api_key=" + encodeURIComponent(apiKey) + "&api_secret_key=" + encodeURIComponent(apiSecretKey);
  var step1Params = { method: "GET", muteHttpExceptions: true };
  
  try {
    var step1Response = UrlFetchApp.fetch(step1Url, step1Params);
    var step1Data = JSON.parse(step1Response.getContentText());
    if (step1Data.status != "1") return { apiSignature: null, sid: null };
    return { apiSignature: step1Data.api_signature, sid: step1Data.sid };
  } catch (e) {
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
    if (step2Data.status != 1) return null;
    return step2Data.contract_list || step2Data.contractList || step2Data.list || (step2Data.data ? step2Data.data.contract_list : null);
  } catch (e) {
    return null;
  }
}

// =========================================================================
// 変更点C：通知関数（YRL番号のリストを受け取り、本文に挿入する）
// =========================================================================
function sendNotification(newContractsCount, notifiedItems, spreadsheetId) {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
  if (!webhookUrl) return;

  // 配列を改行区切りのテキストに変換
  var itemsText = notifiedItems.join("\n");

  var message = {
    text: "～レンタル返却管理～\n" +
          "<users/all>\n" +
          "新たに返却予定の情報が " + newContractsCount + " 件追加（更新）されました！\n\n" +
          "【対象YRL管理番号】\n" +
          itemsText + "\n\n" +
          "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit#gid=1906719251"
  };

  var options = {
    method: "post",
    contentType: "application/json; charset=UTF-8",
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch(webhookUrl, options);
}