function fill_mcv_cv(ss, valueHash, sheetString){
  const sheet  = ss.getSheetByName(sheetString);
  const tgRowFrom = 3;
  const tgColFrom = 7;
  const tgRowCount = 100;
  const tgColCount = 15;
  
  const targets = sheet.getSheetValues(tgRowFrom, tgColFrom, tgRowCount, tgColCount);
  
  var sums = [];
  targets.forEach(function(keys){
    var sum = [0, 0];
    keys.forEach(function(key){
      if (valueHash[key]) {
        sum[0]+=valueHash[key][0];
        sum[1]+=valueHash[key][1];
      }
    })
    sums.push(sum);
  })
  sheet.getRange("D3:E" + (3+sums.length-1).toString()).setValues(sums);
  return;
}
function mcv_cv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName("webantenna");
  const targetSheets = ["バーム紐づけ", ["ホワイト紐づけ"]];
  
  const vRowFrom = 2;
  const vColFrom = 1;
  const vRowCount = 100;
  const vColCount = 3;
  
  
  //セルの内容を2次元配列に格納
  const values = sheet1.getSheetValues(vRowFrom, vColFrom, vRowCount, vColCount);
  
  // NOTE: valueHash とは、媒体ごとの MCV/CV をまとめたもの
  // key: 媒体名
  // value: [<mcv>, <cv>]
  // (例)
  // {
  //    "DUO直定期Ap1": [10, 20],
  //    "DUO直定期Ap16Q": [2, 5]
  // }
  var valueHash = {};
  values.forEach(
    function(value) {
      valueHash[value[0]] = [value[1], value[2]];
    }
  )
  
  // 対象のシートそれぞれの MCV/CVの列を埋めていく
  for (var i in targetSheets) {
      fill_mcv_cv(ss, valueHash, targetSheets[i]);
  }
  
  return;
}


function cost_balm() {
  //シートオブジェクトを取得
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1  = ss.getSheetByName("コスト貼り付け");
  const sheet2  = ss.getSheetByName("バーム紐づけ");
  
  const vRowFrom = 3;
  const vColFrom = 2;
  const vRowCount = 25;
  const vColCount = 4;
  
  const tgRowFrom = 3;
  const tgColFrom = 3;
  const tgRowCount = 30;
  const tgColCount = 1;
  
  
  //セルの内容を2次元配列に格納
  //const values = sheet.getSheetValues(1, colFrom, sheet.getMaxRows(), colTo);
  const values = sheet1.getSheetValues(vRowFrom, vColFrom, vRowCount, vColCount);
  
  const targets = sheet2.getSheetValues(tgRowFrom, tgColFrom, tgRowCount, tgColCount);
  
  var valueHash = {};
  values.forEach(
    function(value) {
      valueHash[value[0]] = value[1];
    }
  )
  
  
  
  // targets
  //りんご	みかん
  //ばなな	みかん
  
  var sums = [];
  targets.forEach(function(keys){
    if (valueHash[keys]){
      sums.push(valueHash[keys]);
    } else {
      sums.push("");
    }
  })
  return sums;
}


