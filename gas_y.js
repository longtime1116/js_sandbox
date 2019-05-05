function fill_mcv_cv(ss, valueHash, sheet){
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

function mcv_cv(ss, targetSheets) {
  const sheet1 = ss.getSheetByName("webantenna");
  
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



function fill_cost_duo(ss, valueHash, sheet, i){
  const tgRowFrom = 3;
  const tgColFrom = 1;
  const tgRowCount = 100;
  const tgColCount = 1;
  
  const targets = sheet.getSheetValues(tgRowFrom, tgColFrom, tgRowCount, tgColCount);
  
  var sums = [];
  targets.forEach(function(key){
    if (valueHash[key]){
      sums.push([valueHash[key][i]]);
    } else {
      sums.push([""]);
    }
  })
  sheet.getRange("F3:F" + (3+sums.length-1).toString()).setValues(sums);
  return;
}




function cost_duo(ss, targetSheets) { 
  //シートオブジェクトを取得
  const sheet1 = ss.getSheetByName("コスト貼り付け");
  
  const vRowFrom = 3;
  const vColFrom = 2;
  const vRowCount = 50;
  const vColCount = 10;
  
  //セルの内容を2次元配列に格納
  const values = sheet1.getSheetValues(vRowFrom, vColFrom, vRowCount, vColCount);
  
  var valueHash = {};
  values.forEach(
    function(value) {
      valueHash[value[0]] = [value[1], value[2]];
    }
  )
  for (var i in targetSheets) {
    fill_cost_duo(ss, valueHash, targetSheets[i], i);
  }
}



function on_edit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetsStr = ["バーム紐づけ", "ホワイト紐づけ"];
  const targetSheets = [];
  for (var i in targetSheetsStr) {
      targetSheets.push(ss.getSheetByName(targetSheetsStr[i]));
  }
  mcv_cv(ss, targetSheets);
  cost_duo(ss, targetSheets);
}
