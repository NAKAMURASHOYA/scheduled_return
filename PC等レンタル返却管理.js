// 条件
// トリガー：毎朝8時～9時に実行
// 本日の日時から45日後迄の期間データをチェック
// 個品番号（YRL管理番号）を変数内に格納しスプレッドシートへ記録
// スプレッドシート内にある個品番号データと比較して重複したデータは追加しない
// スプレッドシートに新しいデータが追加されたらGoogle Chat（【通知用】ITヘルプデスク対応依頼）へAllで通知

function fetchAndWriteContractData() {
  // ▼▼ 設定項目 ▼▼
  var SEARCH_DAYS_RANGE = 60; // 本日から何日後まで検索するか（ここを変更してください）
  // ▲▲ 設定項目 ▲▲

  // --- プロパティの取得 ---
  var props = PropertiesService.getScriptProperties();
  var SPREADSHEET_ID = props.getProperty('SPREADSHEET_ID');
  var API_KEY = props.getProperty('API_KEY');
  var API_SECRET_KEY = props.getProperty('API_SECRET_KEY');
  
  if (!SPREADSHEET_ID || !API_KEY || !API_SECRET_KEY) {
    Logger.log("❌ Error: スクリプトプロパティが設定されていません。");
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
    Logger.log("❌ Error: シートが見つかりません。");
    return;
  }

  // 既存データの読み込み（重複チェック用）
  var existingContracts = new Set();
  var lastRow = sheet.getLastRow();
  
  if (lastRow > 1) { 
    // B列(YRL管理番号)のみ取得
    var data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); 
    data.forEach(function(row) {
      existingContracts.add(String(row[0])); 
    });
  }

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
    // ページ番号を指定して取得
    var pageList = getContractList(API_KEY, authData.apiSignature, authData.sid, page, SEARCH_DAYS_RANGE);
    
    if (pageList && pageList.length > 0) {
      Logger.log("Page " + page + ": " + pageList.length + "件取得");
      allContracts = allContracts.concat(pageList); // 配列を結合
      page++; // 次のページへ
      
      // 安全策: もし1回で25件未満なら、それが最後のページなので終了
      // (API仕様によりますが、通常満タンでなければ次はないため)
      if (pageList.length < 5) { 
        hasNextPage = false; 
      }
      // 無限ループ防止（念のため50ページで止める）
      if (page > 50) {
        hasNextPage = false;
        Logger.log("⚠️ ページ数が多すぎるため50ページで中断します");
      }
    } else {
      // データが取れなくなったら終了
      hasNextPage = false;
    }
    
    // APIサーバーへの負荷軽減のため少し待機
    Utilities.sleep(500); 
  }
  
  Logger.log("データ取得完了。合計件数: " + allContracts.length + "件");

  // --- 書き込み処理 ---
  var newContractsCount = 0;
  var isNewDataAdded = false;

  allContracts.forEach(function(contract) {
    var checkKey = String(contract.khno); 

    if (!existingContracts.has(checkKey)) {
      lastRow = sheet.getLastRow();
      
      // 空白行チェック（念のため）
      if (sheet.getRange(lastRow, 1).getValue() !== "") {
         // 通常はこちら
      }

      // insertRowAfter は処理が重くなることがあるため、単純追記に変更しても良いですが
      // 元のロジック（最終行の後ろに追加）を維持します
      sheet.insertRowAfter(lastRow);
      var targetRow = lastRow + 1;
      
      sheet.getRange(targetRow, 1).setValue(contract.jkno); 
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
  if (isNewDataAdded) {
    sendNotification(newContractsCount, SPREADSHEET_ID);
    Logger.log("通知送信完了。新規: " + newContractsCount + "件");
  } else {
    Logger.log("新規データはありませんでした（取得済みデータのみ）");
  }
}

// Step 1: 認証 (GET URL結合版)
function getAPISignatureAndSID(apiKey, apiSecretKey) {
  var baseUrl = "http://wrt.simplit.jp/direct/member/generate_api_signature/";
  var step1Url = baseUrl + "?api_key=" + apiKey + "&api_secret_key=" + apiSecretKey;

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

// Step 2: データ取得 (ページ番号・日数指定に対応)
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
      "pageID": pageID, // ここが動的に変わるようになりました
      "search[rtod1]": Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      "search[rtod2]": Utilities.formatDate(dateEnd, Session.getScriptTimeZone(), "yyyy-MM-dd")
    }
  };
  
  try {
    var step2Response = UrlFetchApp.fetch(step2Url, step2Params);
    var step2Data = JSON.parse(step2Response.getContentText());
    
    if (step2Data.status != 1) {
      Logger.log("API Error (Step 2) Page " + pageID + ": " + step2Data.status);
      return null;
    }
    return step2Data.contract_list;
  } catch (e) {
    Logger.log("Exception (Step 2): " + e);
    return null;
  }
}

// 通知関数
function sendNotification(newContractsCount, spreadsheetId) {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
  if (!webhookUrl) return;

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
  UrlFetchApp.fetch(webhookUrl, options);
}