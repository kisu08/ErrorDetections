function checkDataS2023(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2023(日経S)");
  var data = sheet.getDataRange().getValues();
  var headers = data[5]; // 6行目が項目名
  var flagRow = 4; // 5行目にフラグを立てる

  // エラー検知条件(ヘッダー部)
  var conditions = {
    "出典種別": function(value, row) {
      var documentType = data[row][headers.indexOf("資料名称")];

      // 資料名称と出典種別の対応関係を定義
      var documentValueMap = {
        "有価証券報告書": "004",
        "コーポレートガバナンス報告書": "005",
        "企業HP": "006"
      };

      var disclosureYear = data[row][headers.indexOf("開示年度")];

      //開示年度2021年のときNULLであること
      if (disclosureYear === 2021) {
        return value === "";
      
      //開示年度2021年度以外は「004→有価証券報告書」「005→コーポレートガバナンス報告書」「006→企業HP」であること
      } else {
        return value === documentValueMap[documentType];
      }
    },
  
    "種別名": function(value, row) {
 
      var validValues = ["資料開示", "開示データ", "加工データ", "単位", "URL", "ページ数", "対象範囲"];
      var isValid = validValues.includes(value);
      var startCol = headers.indexOf("【社会】男性従業員数");
    
      //資料開示の場合、収録項目のデータ入力がないこと
      //有報・CG報告書の場合、URL、ページ数、対象範囲のデータ入力がないこと
      if (value === "資料開示" || 
      (["URL", "ページ数", "対象範囲"].includes(value) && 
      ["有価証券報告書", "コーポレートガバナンス報告書"].includes(data[row][headers.indexOf("資料名称")]))) {
        if (startCol !== -1) {
          for (var col = startCol; col < data[row].length; col++) {
            if (data[row][col] !== "") {
              for (var errorCol = startCol; errorCol < data[row].length; errorCol++) {
                sheet.getRange(row + 1, errorCol + 1).setBackground("yellow");
              }
              return false;
            }
          }
        }
      }

      //「資料開示」「開示データ」「加工データ」「単位」「URL」「ページ数」「対象範囲」のいずれかであること
      return isValid;
    },
    
    "コード": function(value) {
      //4桁の数字もしくは英数字であること
      return /^[0-9]{4}$/.test(value) || /^[A-Za-z0-9]{4}$/.test(value);
    },

    "資料名称": function(value) {
      if(["有価証券報告書","コーポレートガバナンス報告書"].includes(value) && ["URL", "ページ数", "対象範囲"].includes(data[row][headers.indexOf("種別名")])){
        return false
      }
      //「有価証券報告書」「コーポレートガバナンス報告書」「企業HP」のいずれかであること
      return ["企業HP", "有価証券報告書", "コーポレートガバナンス報告書"].includes(value);
    },
  
    "資料公表日": function(value,row) {
      //「有価証券報告書」または「コーポレートガバナンス報告書」の場合、空欄をエラーとする
      var documentType = data[row][headers.indexOf("資料名称")];
      if (["有価証券報告書","コーポレートガバナンス報告書"].includes(documentType) && value === ""){
        return false;
      }
      
      // 日付オブジェクトかどうかを確認
      if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
        return true;
      }
      
      // 文字列として日付フォーマットを検証
      var valueStr = value.toString().trim();
      
      // 日付フォーマットを検証
      var datePattern = /^(?:(19|20)?\d\d年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|(?:19|20)?\d\d[-\/.](0[1-9]|1[0-2])[-\/.](0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])[-\/.](0[1-9]|1[0-2])[-\/.](19|20)?\d\d|(?:0[1-9]|1[0-2])[-\/.](0[1-9]|[12][0-9]|3[01])[-\/.](19|20)?\d\d|令和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|平成[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|昭和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|大正[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|[一二三四五六七八九十百千万]{1,4}年[一二三四五六七八九十]{1,2}月[一二三四五六七八九十]{1,2}日|(?:19|20)?\d\d年(0[1-9]|1[0-2])月|令和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|平成[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|昭和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|大正[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|[一二三四五六七八九十百千万]{1,4}年[一二三四五六七八九十]{1,2}月)|$/

      return datePattern.test(valueStr);
    },
    
    "開示年度": function(value) {
      //値が「2023」あること
      return value === 2023;
    },
    
    "過年度：年": function(value, row) {
      //種別名が資料開示のときNULLであること
      if (data[row][headers.indexOf("種別名")] === "資料開示") {
        return value === "";
      }
      //4桁の数字であること
      return /^[0-9]{4}$/.test(value);
    },
    "過年度：年月／単位（加工値）": function(value, row) {
      //種別名が資料開示のときNULLであること
      if (data[row][headers.indexOf("種別名")] === "資料開示") {
        return value === "";
      }
      //6桁の数字で末尾が「00」～「12」であること
      var valueStr = value.toString(); // 数値を文字列に変換
      return /^[0-9]{6}$/.test(valueStr) && /^(00|01|02|03|04|05|06|07|08|09|10|11|12)$/.test(valueStr.slice(-2));
    }
  };

 // エラー検知とフラグ設定（属性データ）
   for (var row = 6; row < data.length; row++) {
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];
      if (conditions[header] && !conditions[header](value, row)) {
        sheet.getRange(row + 1, col + 1).setBackground("yellow");
        sheet.getRange(flagRow + 1, col + 1).setValue(1);
        }
    }
  };

  //各項目の条件指定
  var textdata = {
    "【社会】男性役員数":function(value,row){
      //有価証券報告書であること。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && !(data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】女性役員数":function(value,row){
      //有価証券報告書であること。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && !(data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】女性役員比率":function(value,row){
      //有価証券報告書であること。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && !(data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },
  };
  // エラー検知とフラグ設定（各項目の条件指定）
  for (var row = 6; row < data.length; row++) {
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];
      if (textdata[header] && !textdata[header](value, row)) {
        sheet.getRange(row + 1, col + 1).setBackground("yellow");
        sheet.getRange(flagRow + 1, col + 1).setValue(1);
        }
    }
  };

    //各項目の条件指定
    var textdata = {
    "【社会】課長以上部長未満の女性管理職比率":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】ジェンダーペイギャップ指数":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】正規従業員数":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】非正規従業員数":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】非正規従業員比率":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】男性従業員の育児休業取得期間":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】女性従業員の育児休業取得期間":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】男性従業員の育児休業取得率":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

    "【社会】女性従業員の育児休業取得率":function(value,row){
      //有価証券報告書からの収録は不要。
      if ((data[row][headers.indexOf("種別名")] === "開示データ" ||data[row][headers.indexOf("種別名")] === "加工データ") && (data[row][headers.indexOf("資料名称")] === "有価証券報告書")){
        return value === "";
      }
      return true;
    },

  };

  // エラー検知とフラグ設定（各項目の条件指定）
  for (var row = 6; row < data.length; row++) {
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];
      if (textdata[header] && !textdata[header](value, row)) {
        sheet.getRange(row + 1, col + 1).setBackground("yellow");
        sheet.getRange(flagRow + 1, col + 1).setValue(1);
        }
    }
  };


  //同一の開示年度で、資料名と資料公表日の組み合わせが一致していること。
  var indexMap = {};
  for (var i = 0; i < headers.length; i++) {
    indexMap[headers[i]] = i;
  }
  var yearIndex = indexMap["開示年度"];
  var nameIndex = indexMap["資料名称"];
  var dateIndex = indexMap["資料公表日"];
  var dataMap = {};
  for (var row = 6; row < data.length; row++) {
    var year = data[row][yearIndex];
    var name = data[row][nameIndex].trim();
    var date = data[row][dateIndex].toString().trim();
    var key = year + "_" + name;
  if (!dataMap[key]) {
    dataMap[key] = date;
    } else if (dataMap[key] !== date) {
      // 資料公表日が一致しない場合
      sheet.getRange(row + 1, dateIndex + 1).setBackground("yellow");
      sheet.getRange(flagRow + 1, dateIndex + 1).setValue(1);
      errorDetected = true;
    }
  }

  // 同一の開示年度では、1資料につき必ず資料開示のレコードは1つであること
  function checkDocumentDisclosure() {
    var disclosureCount = 0;
    var uniqueCombinations = new Set();
    var errorRows = [];

    for (var row = 6; row < data.length; row++) {
      var documentName = data[row][headers.indexOf("資料名称")];
      var typeName = data[row][headers.indexOf("種別名")];
      var disclosureYear = data[row][headers.indexOf("開示年度")];
      
      if (typeName === "資料開示") {
        disclosureCount++;
        };
    
      var combination = documentName + "_" + disclosureYear;
      if(!uniqueCombinations.has(combination)){
        uniqueCombinations.add(combination);
        }
    }
    
    if (uniqueCombinations.size !== disclosureCount) {
      // 修正箇所: エラーが発生した場合の処理
      var typeNameCol = headers.indexOf("種別名");
      sheet.getRange(flagRow + 1, typeNameCol + 1).setValue(1);
      return false;
    }
    return true;
  };
  // 資料名称をキーとして、種別名ごとに「過年度：年」「過年度：年月」の組み合わせが一致するかをチェック（行の入力漏れを検知）
  function checkCombinationConsistency() {
    var baseCombinations = {};
    var typeNameCol = headers.indexOf("種別名");
    var documentNameCol = headers.indexOf("資料名称");
    var pastYearCol = headers.indexOf("過年度：年");
    var pastYearMonthCol = headers.indexOf("過年度：年月／単位（加工値）");

    for (var row = 6; row < data.length; row++) {
      var typeName = data[row][typeNameCol];
      var documentName = data[row][documentNameCol];
      var pastYear = data[row][pastYearCol];
      var pastYearMonth = data[row][pastYearMonthCol];
    
      if (typeName === "資料開示") continue;
    
      var combination = pastYear + "_" + pastYearMonth;
      if (!baseCombinations[documentName]) {
        baseCombinations[documentName] = {};
      }
    
      if (!baseCombinations[documentName][typeName]) {
       baseCombinations[documentName][typeName] = new Set();
      }
      baseCombinations[documentName][typeName].add(combination);
    }
   // デバッグ用にbaseCombinationsの内容を出力
   console.log(JSON.stringify(baseCombinations, null, 2));

    for (var documentName in baseCombinations) {
      var typeNameSets = baseCombinations[documentName];
      var allSets = new Set();
      for (var typeName in typeNameSets) {
        if (allSets.size === 0) {
          allSets = new Set(typeNameSets[typeName]);
        } else {
          console.log(`Checking document: ${documentName}, type: ${typeName}`);
          console.log(`All sets: ${[...allSets].join(", ")}`);
          console.log(`Current set: ${[...typeNameSets[typeName]].join(", ")}`);
          if (allSets.size !== typeNameSets[typeName].size || 
          ![...allSets].every(value => typeNameSets[typeName].has(value))) {
            var mismatch = 'Mismatch found in document: ${documentName}, type: ${typeName}';
            console.log(mismatch);            
            sheet.getRange(flagRow + 1, typeNameCol + 1).setValue(1);
            return false;
          }
        }
      }
    }
    return true;
  }

  //エラー検知とフラグ設定（行ずれ・データ不備）
  if (!checkDocumentDisclosure() || !checkCombinationConsistency()) {
    var typeNameCol = headers.indexOf("種別名");
    sheet.getRange(flagRow + 1, typeNameCol + 1).setBackground("red");
  }

  // エラー検知条件(項目共通)
  for (var row = 6; row < data.length; row++) {
    var typeName = data[row][headers.indexOf("種別名")];
    var startCol = headers.indexOf("【社会】男性従業員数");
    
    //開示データにデータが収録されている場合、加工データにもデータが収録されていること
    if (startCol !== -1 && typeName === "開示データ") {
      for (var col = startCol; col < data[row].length; col++) {
        if (data[row][col] !== "") {
          var matchfound = false;
          for (var row2 = 6; row2 < data.length; row2++) {
            if (data[row][headers.indexOf("出典種別")] === data[row2][headers.indexOf("出典種別")] &&
            data[row][headers.indexOf("コード")] === data[row2][headers.indexOf("コード")] &&
            data[row][headers.indexOf("資料名称")] === data[row2][headers.indexOf("資料名称")] &&
            data[row][headers.indexOf("開示年度")] === data[row2][headers.indexOf("開示年度")] &&
            data[row][headers.indexOf("過年度：年")] === data[row2][headers.indexOf("過年度：年")] &&
            data[row][headers.indexOf("過年度：年月／単位（加工値）")] === data[row2][headers.indexOf("過年度：年月／単位（加工値）")] &&
            (data[row2][headers.indexOf("種別名")] === "加工データ" || data[row2][headers.indexOf("種別名")] === "単位") ){
              if (data[row2][col] === "") {
                sheet.getRange(row2 + 1, col + 1).setBackground("yellow");
                sheet.getRange(flagRow + 1, col + 1).setValue(1);
              }
              matchfound = true;
              if (data[row2][col] === ""){
                sheet.getRange(row2 + 1, col + 1).setBackground("yellow");
                sheet.getRange(flagRow + 1, col + 1).setValue(1);
              }
            }
          }
          if (!matchfound){
            console.log("No match found for row: " + row + ", col: " + col);
          }
        }
      }
    };
    
    // 各項目の閾値設定(過去データの最大値・最小値から決定　※仮決め)
    var thresholds = {
      "開示データ": {
        "【社会】男性従業員数": { "min": 0, "max": 273292 },
        "【社会】女性従業員数": { "min": 0, "max": 395000 },
        "【社会】女性従業員比率": { "min": 0, "max": 100 },
        "【社会】従業員の平均勤続年数": { "min": 0, "max": 25.2 },
        "【社会】男性従業員の平均勤続年数": { "min": 0, "max": 43.4 },
        "【社会】女性従業員の平均勤続年数": { "min": 0, "max": 33.2 },
        "【社会】課長以上部長未満の男性管理職数": { "min": 0, "max": 24770 },
        "【社会】課長以上部長未満の女性管理職数": { "min": 0, "max": 9484 },
        "【社会】課長以上部長未満の女性管理職比率": { "min": 0, "max": 100 },
        "【社会】部長以上の男性管理職数": { "min": 0, "max": 2969 },
        "【社会】部長以上の女性管理職数": { "min": 0, "max": 374},
        "【社会】部長以上の女性管理職比率": { "min": 0, "max": 100 },
        "【社会】男性役員数": { "min": 0, "max": 50 },
        "【社会】女性役員数": { "min": 0, "max": 50 },
        "【社会】女性役員比率": { "min": 0, "max": 100 },
        "【社会】従業員の障害者比率": { "min": 0, "max": 100 },
        "【社会】ジェンダーペイギャップ指数": { "min": 0, "max": 150 },
        "【社会】正規従業員数": { "min": 0, "max": 368247 },
        "【社会】正規従業員の入職者数": { "min": 0, "max": 29539 },
        "【社会】正規従業員の離職者数": { "min": 0, "max": 7330 },
        "【社会】正規従業員の入職超過率": { "min": -20, "max": 100 },
        "【社会】非正規従業員数": { "min": 0, "max": 405000 },
        "【社会】非正規従業員比率": { "min": 0, "max": 100 },
        "【社会】キャリア採用人数": { "min": 0, "max": 5120 },
        "【社会】再雇用人数": { "min": 0, "max": 1795 },
        "【社会】組合加入従業員数": { "min": 0, "max": 300000 },
        "【社会】組合加入従業員比率": { "min": 0, "max": 100 },
        "【社会】団体交渉権をもつ従業員の比率": { "min": 0, "max": 100 },
        "【社会】従業員１人あたりの年間平均研修時間": { "min": 0, "max": 825 },
        "【社会】労働災害度数率": { "min": 0, "max": 10 },
        "【社会】従業員における労働関連の傷害による死亡者数": { "min": 0, "max": 100 },
        "【社会】従業員における労働関連の傷害による死亡者比率": { "min": 0, "max": 100 },
        "【社会】労働災害関連の発生件数": { "min": 0, "max": 2113 },
        "【社会】休業災害件数": { "min": 0, "max": 621 },
        "【社会】不休業災害件数": { "min": 0, "max": 599 },
        "【社会】年次有給休暇取得率": { "min": 0, "max": 111.9 },
        "【社会】介護休職取得者数": { "min": 0, "max": 488 },
        "【社会】介護休職取得者日数": { "min": 0, "max": 339 },
        "【社会】男性従業員の育児休業取得期間": { "min": 0, "max": 460 },
        "【社会】女性従業員の育児休業取得期間": { "min": 0, "max": 2289 },
        "【社会】男性従業員の育児休業取得率": { "min": 0, "max": 100 },
        "【社会】女性従業員の育児休業取得率": { "min": 0, "max": 100 },
        "【社会】男性従業員の平均年収": { "min": 100, "max": 30000000 },
        "【社会】女性従業員の平均年収": { "min": 100, "max": 30000000 },
        "【社会】メンタルヘルス不調による休職者数": { "min": 0, "max": 2552 },
        "【社会】メンタルヘルス不調による休職者率": { "min": 0, "max": 100 },
        "【社会】従業員１人あたりの年間平均総労働時間": { "min": 0, "max": 2366 },
        "【社会】従業員１人あたりの年間研修費": { "min": 0, "max": 555403 }
      },
      "加工データ": {
        "【社会】男性従業員数": { "min": 0, "max": 273292 },
        "【社会】女性従業員数": { "min": 0, "max": 395000 },
        "【社会】女性従業員比率": { "min": 0, "max": 100 },
        "【社会】従業員の平均勤続年数": { "min": 0, "max": 25.2 },
        "【社会】男性従業員の平均勤続年数": { "min": 0, "max": 43.4 },
        "【社会】女性従業員の平均勤続年数": { "min": 0, "max": 33.2 },
        "【社会】課長以上部長未満の男性管理職数": { "min": 0, "max": 24770 },
        "【社会】課長以上部長未満の女性管理職数": { "min": 0, "max": 9484 },
        "【社会】課長以上部長未満の女性管理職比率": { "min": 0, "max": 100 },
        "【社会】部長以上の男性管理職数": { "min": 0, "max": 2969 },
        "【社会】部長以上の女性管理職数": { "min": 0, "max": 374},
        "【社会】部長以上の女性管理職比率": { "min": 0, "max": 100 },
        "【社会】男性役員数": { "min": 0, "max": 50 },
        "【社会】女性役員数": { "min": 0, "max": 50 },
        "【社会】女性役員比率": { "min": 0, "max": 100 },
        "【社会】従業員の障害者比率": { "min": 0, "max": 100 },
        "【社会】ジェンダーペイギャップ指数": { "min": 0, "max": 150 },
        "【社会】正規従業員数": { "min": 0, "max": 368247 },
        "【社会】正規従業員の入職者数": { "min": 0, "max": 29539 },
        "【社会】正規従業員の離職者数": { "min": 0, "max": 7330 },
        "【社会】正規従業員の入職超過率": { "min": -20, "max": 100 },
        "【社会】非正規従業員数": { "min": 0, "max": 405000 },
        "【社会】非正規従業員比率": { "min": 0, "max": 100 },
        "【社会】キャリア採用人数": { "min": 0, "max": 5120 },
        "【社会】再雇用人数": { "min": 0, "max": 1795 },
        "【社会】組合加入従業員数": { "min": 0, "max": 300000 },
        "【社会】組合加入従業員比率": { "min": 0, "max": 100 },
        "【社会】団体交渉権をもつ従業員の比率": { "min": 0, "max": 100 },
        "【社会】従業員１人あたりの年間平均研修時間": { "min": 0, "max": 825 },
        "【社会】労働災害度数率": { "min": 0, "max": 10 },
        "【社会】従業員における労働関連の傷害による死亡者数": { "min": 0, "max": 100 },
        "【社会】従業員における労働関連の傷害による死亡者比率": { "min": 0, "max": 100 },
        "【社会】労働災害関連の発生件数": { "min": 0, "max": 2113 },
        "【社会】休業災害件数": { "min": 0, "max": 621 },
        "【社会】不休業災害件数": { "min": 0, "max": 599 },
        "【社会】年次有給休暇取得率": { "min": 0, "max": 111.9 },
        "【社会】介護休職取得者数": { "min": 0, "max": 488 },
        "【社会】介護休職取得者日数": { "min": 0, "max": 339 },
        "【社会】男性従業員の育児休業取得期間": { "min": 0, "max": 460 },
        "【社会】女性従業員の育児休業取得期間": { "min": 0, "max": 2289 },
        "【社会】男性従業員の育児休業取得率": { "min": 0, "max": 100 },
        "【社会】女性従業員の育児休業取得率": { "min": 0, "max": 100 },
        "【社会】男性従業員の平均年収": { "min": 100, "max": 10952 },
        "【社会】女性従業員の平均年収": { "min": 100, "max": 7319.069 },
        "【社会】メンタルヘルス不調による休職者数": { "min": 0, "max": 2552 },
        "【社会】メンタルヘルス不調による休職者率": { "min": 0, "max": 100 },
        "【社会】従業員１人あたりの年間平均総労働時間": { "min": 0, "max": 2366 },
        "【社会】従業員１人あたりの年間研修費": { "min": 0, "max": 555403 }
      }
    };
    // 各データが閾値の範囲内であることをチェック
    function checkThresholds(typeName, data, headers, row, startCol, sheet, flagRow) {
      var thresholdData = thresholds[typeName];
      for (var col = startCol; col < data[row].length; col++) {
        if (data[row][col] !== "") {
          var header = headers[col];
          var thresholdRange = thresholdData[header];
          if (!thresholdRange) {
            continue;
          }
          var value = parseFloat(data[row][col]);
          if (isNaN(value) || value < thresholdRange.min || value > thresholdRange.max) {
            sheet.getRange(row + 1, col + 1).setBackground("yellow");
            sheet.getRange(flagRow + 1, col + 1).setValue(1);
          }
        }
      }
    }
    // 閾値チェックを実行
    if (typeName === "開示データ" || typeName === "加工データ") {
      checkThresholds(typeName, data, headers, row, startCol, sheet, flagRow);
    }

    //「https:～」の文字列を含むこと
    if (typeName === "URL"){
      for (var col = startCol; col < data[row].length; col++){
        var cellValue = String(data[row][col]);
        if(cellValue !== "" && !(cellValue.includes("https://")||cellValue.includes("http://"))){
          // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
          sheet.getRange(row + 1, col + 1).setBackground("yellow");
          sheet.getRange(flagRow + 1, col + 1).setValue(1);
        }
      }

    //反対に、URL以外では「https:～」の文字列を含んでいないこと
    }else if (["開示データ", "加工データ", "単位", "ページ数", "対象範囲"].includes(typeName)){
      for (var col = startCol; col < data[row].length; col++) {
        var cellValue = String(data[row][col]);
        if (cellValue !== "" && (cellValue.includes("https://")||cellValue.includes("http://"))) {
          // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
          sheet.getRange(row + 1, col + 1).setBackground("yellow");
          sheet.getRange(flagRow + 1, col + 1).setValue(1);
        }
      }
    };

    //数値データであること
    if (typeName == "開示データ" || typeName == "加工データ" ){
      if (startCol !== -1){
        for (var col = startCol; col < data[row].length; col++){
          if (data[row][col] !== "" && isNaN(data[row][col])){
            //エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
            sheet.getRange(row + 1, col + 1).setBackground("yellow");
            sheet.getRange(flagRow +1, col+1).setValue(1);
          }
        }
      }
    }

    //ページ数は特定のフォーマットであること（数値、カンマ、ハイフンで構成）
    if (typeName == "ページ数"){
      if (startCol !== -1){
        for (var col = startCol; col < data[row].length; col++){
          if(data[row][col] !== ""){
            var value = data[row][col];
            var isValid = !isNaN(value) || /^[0-9,.-]+$/.test(value);
            if (!isValid) {
            // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
            sheet.getRange(row + 1, col + 1).setBackground("yellow");
            sheet.getRange(flagRow +1, col + 1).setValue(1);
            }
          }
        }
      }
    }

    //文字列になっていること
    if (typeName == "単位" || typeName == "対象範囲"){
      if(startCol !== -1){
        for (var col = startCol; col < data[row].length; col++){
          if(data[row][col] !== "" && !isNaN(data[row][col])){
           // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
            sheet.getRange(row + 1, col + 1).setBackground("yellow");
            sheet.getRange(flagRow +1, col + 1).setValue(1); 
          }
        }
      }
    };
  };

  //過年度年月が「00」の時に、過年度年と過年度年月の西暦が一致していること
  for (var row = 6; row < data.length; row++) {
    var pastYearCol = headers.indexOf("過年度：年");
    var pastYearMonthCol = headers.indexOf("過年度：年月／単位（加工値）");
    var pastYearValue = data[row][pastYearCol];
    var pastYearMonthValue = data[row][pastYearMonthCol];

    if (pastYearMonthValue.toString().slice(-2) === "00"){
      if (pastYearValue.toString().slice(0,4) !== pastYearMonthValue.toString().slice(0,4)){
        sheet.getRange(row + 1, pastYearCol + 1).setBackground("yellow");
        sheet.getRange(row + 1, pastYearMonthCol + 1).setBackground("yellow");
        sheet.getRange(flagRow +1, pastYearCol + 1).setValue(1); 
        sheet.getRange(flagRow +1, pastYearMonthCol + 1).setValue(1);
      }
    }
  };

  
  // 「開示データ」または「加工データ」は必ず1行は存在すること。
  // キー項目のインデックスを取得
  var keyIndices = [
    headers.indexOf("出典種別"),
    headers.indexOf("コード"),
    headers.indexOf("開示年度"),
    headers.indexOf("過年度：年"),
    headers.indexOf("過年度：年月／単位（加工値）")
  ];

  // キー項目の値を連結してキーを作成する関数
  function createKey(row) {
    return keyIndices.map(function(index) {
      return data[row][index];
    }).join("-");
  }

  // キーごとに「開示データ」「加工データ」の存在を確認するためのオブジェクト
  var keyMap = {};

  // データを走査してキーごとに「開示データ」「加工データ」の存在を確認
  for (var row = 6; row < data.length; row++) {
    var key = createKey(row);
    var type = data[row][headers.indexOf("種別名")];
    if (!keyMap[key]) {
      keyMap[key] = {
        "開示データ": false,
        "加工データ": false
      };
    }
    if (type === "開示データ" || type === "加工データ") {
      keyMap[key][type] = true;
    }
  }

  // エラー検知とフラグ設定
  for (var row = 6; row < data.length; row++) {
    var key = createKey(row);
    var type = data[row][headers.indexOf("種別名")];

    if (type === "開示データ" || type === "加工データ") {
      if (!keyMap[key]["開示データ"] || !keyMap[key]["加工データ"]) {
        for (var col = 0; col < headers.length; col++) {
          sheet.getRange(row + 1, col + 1).setBackground("yellow");
        }
        var flagCol = headers.indexOf("種別名");
        sheet.getRange(flagRow + 1, flagCol + 1).setValue(1);
        sheet.getRange(flagRow + 1, flagCol + 1).setBackground("red");
      }
    }
  };

  //キー項目が同じレコードがないこと（重複していないこと）
  // 「出典種別」「種別名」「コード」「開示年度」「過年度：年」「過年度：年月／単位（加工値）」の列の値が完全に一致している行が複数ある場合
  var uniqueRows = {};
  var duplicateRows = [];

  for (var row = 6; row < data.length; row++) {
    var key = [
      data[row][headers.indexOf("出典種別")],
      data[row][headers.indexOf("種別名")],
      data[row][headers.indexOf("コード")],
      data[row][headers.indexOf("開示年度")],
      data[row][headers.indexOf("過年度：年")],
      data[row][headers.indexOf("過年度：年月／単位（加工値）")]
    ].join("|");

    if (uniqueRows[key]) {
      duplicateRows.push(row);
      duplicateRows.push(uniqueRows[key]);
    } else {
      uniqueRows[key] = row;
    }
  }

  // 重複行が見つかった場合、5行目5列目のセルの背景色を赤色にし、そのセルに1を入力し、該当する行の2列目のセルを赤色にする
  if (duplicateRows.length > 0) {
    sheet.getRange(5, 2).setBackground("red").setValue(1);
    duplicateRows.forEach(function(row) {
      sheet.getRange(row + 1, 2).setBackground("red");
    });
  }

  SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');
};