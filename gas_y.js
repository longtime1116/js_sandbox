function mcv_cv() {
  //シートオブジェクトを取得
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1  = ss.getSheetByName("webantenna");
  const sheet2  = ss.getSheetByName("バーム紐づけ");
  
  const vRowFrom = 2;
  const vColFrom = 1;
  const vRowCount = 100;
  //const vRowCount = sheet1.getMaxRows();
  const vColCount = 3;
  
  const tgRowFrom = 3;
  const tgColFrom = 7;
  const tgRowCount = 100;
  //const tgRowCount = sheet2.getMaxRows();
  const tgColCount = 15;
  //const tgColCount = sheet2.getMaxColumns();
  
  
  //セルの内容を2次元配列に格納
  const values = sheet1.getSheetValues(vRowFrom, vColFrom, vRowCount, vColCount);
  const targets = sheet2.getSheetValues(tgRowFrom, tgColFrom, tgRowCount, tgColCount);
  
  // NOTE: value_hash とは?
  // key: 媒体名
  // value: [<mcv>, <cv>]
  // (例)
  // {
  //    "DUO直定期Ap1": [10, 20],
  //    "DUO直定期Ap16Q": [2, 5]
  // }
  var value_hash = {};
  values.forEach(
    function(value) {
      value_hash[value[0]] = [value[1], value[2]];
    }
  )
  
  var sums = [];
  targets.forEach(function(keys){
    var sum = [0, 0];
    keys.forEach(function(key){
      if (value_hash[key]) {
        sum[0]+=value_hash[key][0];
        sum[1]+=value_hash[key][1];
      }
    })
    sums.push(sum);
  })
  sheet2.getRange("D3:E" + (3+sums.length-1).toString()).setValues(sums);
  return sums;
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
  Logger.log("<values>");
  Logger.log(values);
  
  const targets = sheet2.getSheetValues(tgRowFrom, tgColFrom, tgRowCount, tgColCount);
  Logger.log("<targets>");
  Logger.log(targets);
  
  var value_hash = {};
  values.forEach(
    function(value) {
      value_hash[value[0]] = value[1];
    }
  )
  Logger.log("<value_hash>");
  Logger.log(value_hash);
  
  
  
  // targets
  //りんご	みかん
  //ばなな	みかん
  
  var sums = [];
  targets.forEach(function(keys){
    if (value_hash[keys]){
      sums.push(value_hash[keys]);
    } else {
      sums.push("");
    }
    Logger.log(sums);
  })
  return sums;
}

