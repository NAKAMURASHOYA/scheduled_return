// 条件
// トリガー：毎朝8時～9時に実行
// 本日の日時から40日後迄の期間データをチェック
// 個品番号（YRL管理番号）を変数内に格納しスプレッドシートへ記録
// スプレッドシート内にある個品番号データと比較して重複したデータは追加しない
// スプレッドシートに新しいデータが追加されたらGoogle Chat（【通知用】ITヘルプデスク対応依頼）へAllで通知

function fetchAndWriteContractData() {
  // スプレッドシートの特定のシートを取得
  var spreadsheet = SpreadsheetApp.openById("1lLD0DAdixWd4dmMY9LXZzmSiNbfb9XaTTmP70cchNBc"); // スプレッドシートのIDを指定
  var sheet = spreadsheet.getSheetByName("PC等レンタル返却管理"); // シート名を指定
  
  // 以前に取得した契約データを保存する Set オブジェクト
  var existingContracts = new Set();
  var data = sheet.getDataRange().getValues();
  
  // スプレッドシート内の契約データを existingContracts に追加
  data.forEach(function(row) {
    existingContracts.add(row[1]); // 個品番号（YRL管理番号）をセットに追加
  });

  // APIキーとシークレットキー
  var apiKey = '9s7y7w5v0mkock8884wkk40cg';
  var apiSecretKey = '537xdz0p2c08wwscggg4400cw';

  // Step 1: API SignatureとSIDを取得
  var { apiSignature, sid } = getAPISignatureAndSID(apiKey, apiSecretKey);
  if (!apiSignature || !sid) {
    Logger.log("Failed to get API signature and SID");
    return;
  }

  // レンタル契約情報の取得
  var contractList = getContractList(apiKey, apiSignature, sid);
  if (!contractList) {
    Logger.log("Failed to fetch contract data");
    return;
  }

  // 新たに追加された契約数をカウントする変数
  var newContractsCount = 0;

  // 契約データを書き込む
  var isNewDataAdded = false; // 新しいデータが追加されたかどうかを示すフラグ
  contractList.forEach(function(contract) {
    // 重複する契約データが存在しない場合のみ処理を実行
    if (!existingContracts.has(contract.khno)) {
      // 最終行を取得
      var lastRow = sheet.getLastRow();
      // H列以降にデータがある場合は、G列までの最終行に書き込む
      if (sheet.getRange(lastRow, 8).getValue() !== "") {
        lastRow++;
      }
      
      // 新しいデータを追加するためにシートの末尾に行を追加
      sheet.insertRowAfter(lastRow);
      
      // 指定された列にデータを書き込む
      sheet.getRange(lastRow + 1, 1).setValue(contract.jkno); // 契約番号 (A列)
      sheet.getRange(lastRow + 1, 2).setValue(contract.khno); // YRL管理番号 (B列)
      sheet.getRange(lastRow + 1, 3).setValue(contract.rtod); // レンタル終了予定日 (C列)
      sheet.getRange(lastRow + 1, 4).setValue(contract.kmrk); // メーカー略称 (D列)
      sheet.getRange(lastRow + 1, 5).setValue(contract.khnm); // 品名 (E列)
      sheet.getRange(lastRow + 1, 6).setValue(contract.srno); // シリアル番号 (F列)
      sheet.getRange(lastRow + 1, 7).setValue(contract.statics_name_s); // 商品小分類=製品カテゴリ (G列)
      
      existingContracts.add(contract.khno); // 重複をチェックするために、契約番号をexistingContractsに追加する
      
      isNewDataAdded = true; // 新しいデータが追加されたことをフラグで示す
      newContractsCount++;
    } else {
      Logger.log("Duplicate contract found: " + contract.khno); // 重複が見つかった場合にログに記録する
    }
  });

  // 新しいデータが追加された場合に通知を送信
  if (isNewDataAdded) {
    sendNotification(newContractsCount); // Google Chat へ通知を送信する関数を呼び出す
  }

  Logger.log("Contract data fetched and written to spreadsheet successfully");
}

// API SignatureとSIDを取得する関数
function getAPISignatureAndSID(apiKey, apiSecretKey) {
  var step1Url = "http://wrt.simplit.jp/direct/member/generate_api_signature/";
  var step1Params = {
    method: "GET",
    muteHttpExceptions: true,
    payload: {
      api_key: apiKey,
      api_secret_key: apiSecretKey
    }
  };
  var step1Response = UrlFetchApp.fetch(step1Url, step1Params);
  var step1Data = JSON.parse(step1Response.getContentText());
  if (step1Data.status != "1") {
    Logger.log("Failed to get API signature and SID");
    return { apiSignature: null, sid: null };
  }
  return { apiSignature: step1Data.api_signature, sid: step1Data.sid };
}

// レンタル契約情報を取得する関数
function getContractList(apiKey, apiSignature, sid) {
  var step2Url = "https://wrt.simplit.jp/management/slm/slm_contract_list_api/";
  var step2Params = {
    method: "POST",
    muteHttpExceptions: true,
    payload: {
      api_key: apiKey,
      api_signature: apiSignature,
      sid: sid,
      pageID: 1, // ページ番号は1から始める
      "search[rtod1]": Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"), // 本日の日付
      "search[rtod2]": Utilities.formatDate(new Date(Date.now() + 40 * 24 * 60 * 60 * 1000), Session.getScriptTimeZone(), "yyyy-MM-dd") // 40日後の日付（横河レンタ・リースの返却連絡イメージが3週間～1ヶ月の間との事だった為）
    }
  };
  var step2Response = UrlFetchApp.fetch(step2Url, step2Params);
  var step2Data = JSON.parse(step2Response.getContentText());
  if (step2Data.status != 1) {
    Logger.log("Failed to fetch contract data");
    return null;
  }
  return step2Data.contract_list;
}

// Google Chat へ通知を送信する関数
function sendNotification(newContractsCount) {
  // 通知の内容を設定
  var message = {
    text:"～レンタル返却管理～"
    + '\n' +
    "<users/all>"
    + '\n' +
    "新たに返却予定の情報が " + newContractsCount + " 件追加されました！"
    + '\n\n' +
    "https://docs.google.com/spreadsheets/d/1lLD0DAdixWd4dmMY9LXZzmSiNbfb9XaTTmP70cchNBc/edit#gid=1906719251"
  };

  // 送信先のチャットルームの Webhook URL を設定
  var webhookUrl = "https://chat.googleapis.com/v1/spaces/AAAAKofE5zM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=GfY9OxioFoI-wp6iU2zP9X5v33QHP0wJT6nL8FhXhCw"; // ここに自分のGoogle ChatのWebhook URLを設定する

  // HTTPリクエストを送信して通知を送信
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message)
  };
  UrlFetchApp.fetch(webhookUrl, options);
}