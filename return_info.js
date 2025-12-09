function return_count() {

  const ss = SpreadsheetApp.openById('1lLD0DAdixWd4dmMY9LXZzmSiNbfb9XaTTmP70cchNBc');
  const st = ss.getSheetByName("PC等レンタル返却管理");

  let today = new Date();
  let lastRow = st.getLastRow();
  let delidayrange = st.getRange(2,14,lastRow-1);
  let todayarr = delidayrange.getValues();
  let rows = [];

  //返却日が今日の行番号を取得して配列に格納
  for (let i = 0; i < todayarr.length; i++){
   
    if(todayarr[i][0] instanceof Date && todayarr[i][0].toDateString() === today.toDateString()){
      rows.push(i + 2);
    }

  } 
  
  //今日返却される機器がある場合処理実施
  if (rows !== 0) {

    let cols = [16,2,4,7];
    let dData = [];

    // 対象行のB~F列の値を配列に格納する
    for (let n = 0; n < rows.length; n++) {

      let rowData = [];

      for (let s = 0; s < cols.length; s++) {
        rowData.push(st.getRange(rows[n], cols[s]).getValue());
      }

      dData.push(rowData);

    };

    if (dData.length == 0){
      return
    }else{
      return_info(rows,dData); // Google Chat へ通知を送信する関数を呼び出す  
    }
  }

}

// Google Chat へ通知を送信する関数
function return_info(rows,dData) {

  // 送信先のチャットルームの Webhook URL を設定
  let webhookUrl = "https://chat.googleapis.com/v1/spaces/AAAAKofE5zM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=GfY9OxioFoI-wp6iU2zP9X5v33QHP0wJT6nL8FhXhCw"; 

  // 通知の内容を設定
  let message = {
    text: "～返却通知～\n本日返却の機器があります\n"
  };

  let previousLocation = null; 

  for (let i = 0; i < rows.length; i++) {
    
    if (dData[i][0] !== previousLocation) {
      message.text += `\n【${dData[i][0]}】\n`;
      previousLocation = dData[i][0];
    }

    let equipmentInfo = [];
    
    for (let j = 1; j < dData[i].length; j++) {
      equipmentInfo.push(dData[i][j]);
    }
    
    message.text += `★ ${equipmentInfo.join(" ")} \n`;
  
  }

  // HTTPリクエストを送信して通知を送信
  let options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(message)
  };

  UrlFetchApp.fetch(webhookUrl, options);

}



