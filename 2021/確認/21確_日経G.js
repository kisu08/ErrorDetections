function checkDataG2021(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2021(日経G)");
  var data = sheet.getDataRange().getValues();
  var headers = data[5]; // 6行目が項目名
  var flagRow = 4; // 5行目にフラグを立てる

  // ヘッダーインデックスを先頭に定義
  var headerIndices = {
    documentNameCol: headers.indexOf("資料名称"),
    disclosureYearCol: headers.indexOf("開示年度"),
    typeNameCol: headers.indexOf("種別名"),
    sourceTypeCol: headers.indexOf("出典種別"),
    codeCol: headers.indexOf("コード"),
    pastYearCol: headers.indexOf("過年度：年"),
    pastYearMonthCol: headers.indexOf("過年度：年月／単位（加工値）"),
    startCol: headers.indexOf("【ガバナンス】取締役人数"),
    disclosureDateCol: headers.indexOf("資料公表日")
  };
  // 色がついているセルをカウントしてA5に表示する関数
function countColoredCells(sheet) {
  // A7以降のA列の範囲を取得
  var range = sheet.getRange('A7:A');
  var backgrounds = range.getBackgrounds();
  
  var count = 0;
  
  // 色がついているセルをカウント
  for (var i = 0; i < backgrounds.length; i++) {
    if (backgrounds[i][0] !== '#ffffff' && backgrounds[i][0] !== '') {
      count++;
    }
  }
  
  // A5A6セルを連結し、エラー行数を表示
    var cell = sheet.getRange('A5:A6');
    if (count > 0) {
      // エラーがある場合の表示
      cell.merge() // A5A6を連結
          .setValue("エラー行数: " + count + "行")
          .setFontColor('red')          // 赤文字
          .setFontWeight('bold')        // 太字
          .setHorizontalAlignment('center') // 水平方向の中央揃え
          .setVerticalAlignment('middle')   // 垂直方向の中央揃え
          .setFontSize(13)              // フォントサイズを13に設定
          .setBackground('#FFCCCC');    // 明るい赤色
    } else {
      // エラーがない場合の表示
      cell.merge() // A5A6を連結
          .setValue("データは正常です") // メッセージを「データは正常です」に変更
          .setFontColor('blue')         // 青文字
          .setFontWeight('bold')        // 太字
          .setHorizontalAlignment('center') // 水平方向の中央揃え
          .setVerticalAlignment('middle')   // 垂直方向の中央揃え
          .setFontSize(13)              // フォントサイズを13に設定
          .setBackground('#CCE5FF');    // 明るい青色
    }
  }
  // エラー検知してイエローに変更し、A列に「エラーが発生」を追加する関数
  function setErrorHighlight(sheet, row, col, flagRow) {
    sheet.getRange(row + 1, col + 1).setBackground("yellow");  // エラーセルの背景色を黄色に変更
    sheet.getRange(flagRow + 1, col + 1).setValue(1);  // フラグ行に1をセット
    sheet.getRange(row + 1, 1).setValue("入力に不備があります");  // A列にエラーメッセージをセット
    sheet.getRange(row + 1, 1).setBackground("orange");
  }
  // 有価証券報告書からの収録条件を共通化
  function checkForReport(value, row, data, headerIndices, requiredDocument = "有価証券報告書", exclude = false) {
    var typeName = data[row][headerIndices.typeNameCol];
    var documentName = data[row][headerIndices.documentNameCol];
    
    // 開示データまたは加工データの場合に、指定した報告書名で条件が合致するかチェック
    var conditionMet = (typeName === "開示データ" || typeName === "加工データ") && (documentName === requiredDocument);
    
    if (exclude) {
      // 指定した報告書で「ない場合」をチェック
      return !conditionMet || value === "";
    } else {
      // 指定した報告書で「ある場合」をチェック
      return conditionMet ? value === "" : true;
    }
  }
  // エラー検知条件(ヘッダー部)
  var conditions = {
    "出典種別": function(value, row) {
      var documentType = data[row][headerIndices.documentNameCol];//資料名称

      // 資料名称と出典種別の対応関係を定義
      var documentValueMap = {
        "有価証券報告書": "004",
        "コーポレートガバナンス報告書": "005",
        "企業HP": "006"
      };

      var disclosureYear = data[row][headerIndices.disclosureYearCol];//開示年度

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
      var startCol = headerIndices.startCol;//【ガバナンス】取締役人数
    
      //資料開示の場合、収録項目のデータ入力がないこと
      //有報・CG報告書の場合、URL、ページ数、対象範囲のデータ入力がないこと
      if (value === "資料開示" || 
      (["URL", "ページ数", "対象範囲"].includes(value) && 
      ["有価証券報告書", "コーポレートガバナンス報告書"].includes(data[row][headerIndices.documentNameCol]))) {
        if (startCol !== -1) {
          for (var col = startCol; col < data[row].length; col++) {
            if (data[row][col] !== "") {
              for (var errorCol = startCol; errorCol < data[row].length; errorCol++) {
                sheet.getRange(row + 1, errorCol + 1).setBackground("yellow");
                sheet.getRange(row + 1, 1).setValue("不正なデータです");  // A列にエラーメッセージをセット
                sheet.getRange(row + 1, 1).setBackground("orange");
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
      if(["有価証券報告書","コーポレートガバナンス報告書"].includes(value) && ["URL", "ページ数", "対象範囲"].includes(data[row][headerIndices.typeNameCol])){
        return false
      }
      //「有価証券報告書」「コーポレートガバナンス報告書」「企業HP」のいずれかであること
      return ["企業HP", "有価証券報告書", "コーポレートガバナンス報告書"].includes(value);
    },
  
    "資料公表日": function(value,row) {
      //「有価証券報告書」または「コーポレートガバナンス報告書」の場合、空欄をエラーとする
      var documentType = data[row][headerIndices.documentNameCol];
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
      //値が「2021」あること
      return value === 2021;
    },
    
    "過年度：年": function(value, row) {
      //種別名が資料開示のときNULLであること
      if (data[row][headerIndices.typeNameCol] === "資料開示") {
        return value === "";
      }
      //4桁の数字であること
      return /^[0-9]{4}$/.test(value);
    },
    "過年度：年月／単位（加工値）": function(value, row) {
      //種別名が資料開示のときNULLであること
      if (data[row][headerIndices.typeNameCol] === "資料開示") {
        return value === "";
      }
      //6桁の数字で末尾が「00」～「12」であること
      var valueStr = value.toString(); // 数値を文字列に変換
      return /^[0-9]{6}$/.test(valueStr) && /^(00|01|02|03|04|05|06|07|08|09|10|11|12)$/.test(valueStr.slice(-2));
    },
    
  };

  // エラー検知とフラグ設定（ヘッダー部）
   for (var row = 6; row < data.length; row++) {
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];
      if (conditions[header] && !conditions[header](value, row)) {
        setErrorHighlight(sheet, row, col, flagRow);
        }
    }
  };

  // 各項目の条件指定
var textdata = {
  "【ガバナンス】取締役人数": function(value, row) {
    return checkForReport(value, row, data, headerIndices, "有価証券報告書", true); // 否定条件
  },
  "【ガバナンス】社外取締役人数": function(value, row) {
    return checkForReport(value, row, data, headerIndices, "有価証券報告書", true); // 否定条件
  },
  "【ガバナンス】独立社外取締役人数": function(value, row) {
    return checkForReport(value, row, data, headerIndices, "有価証券報告書", true); // 否定条件
  },
  "【ガバナンス】独立社外取締役比率": function(value, row) {
    return checkForReport(value, row, data, headerIndices, "有価証券報告書", true); // 否定条件
  },
  "【ガバナンス】監査役人数": function(value, row) {
    return checkForReport(value, row, data, headerIndices, "有価証券報告書", true); // 否定条件
  },
  "【ガバナンス】独立社外監査役人数": function(value, row) {
    return checkForReport(value, row, data, headerIndices, "有価証券報告書", true); // 否定条件
  },
  "【ガバナンス】独立社外監査役比率": function(value, row) {
    return checkForReport(value, row, data, headerIndices, "有価証券報告書", true); // 否定条件
  },
  "【ガバナンス】役員の固定報酬": function(value, row) {
    return checkForReport(value, row, data, headerIndices); // 肯定条件（デフォルト）
  },
  "【ガバナンス】役員の変動報酬": function(value, row) {
    return checkForReport(value, row, data, headerIndices); // 肯定条件（デフォルト）
  },
  "【ガバナンス】役員の変動報酬比率": function(value, row) {
    return checkForReport(value, row, data, headerIndices); // 肯定条件（デフォルト）
  },
  "【ガバナンス】ガバナンス体系（組織体系）": function(value, row) {
    return checkForReport(value, row, data, headerIndices) &&
      ["監査役設置会社", "委員会設置会社", "監査等委員会設置会社", "指名委員会等設置会社"].includes(value); // 特定の値
  },
  "【ガバナンス】取締役会の議長": function(value, row) {
    return checkForReport(value, row, data, headerIndices) &&
      ["社長", "会長（社長を兼任している場合を除く）", "会長・社長以外の執行役を兼任する取締役", 
       "会長・社長以外の代表取締役", "社外取締役", "その他の取締役", "なし"].includes(value); // 特定の値
  },
  "【ガバナンス】外国人株式保有比率": function(value, row) {
    return checkForReport(value, row, data, headerIndices) &&
      ["10%未満", "10%以上20%未満", "20%以上30%未満", "30%以上"].includes(value); // 特定の値
  }
};

  // エラー検知とフラグ設定（各項目の条件指定）
  for (var row = 6; row < data.length; row++) {
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];
      if (textdata[header] && !textdata[header](value, row)) {
        setErrorHighlight(sheet, row, col, flagRow);
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
      sheet.getRange(row + 1, 1).setValue("公表日一致していません");  // A列にエラーメッセージをセット
      sheet.getRange(row + 1, 1).setBackground("tan");
      errorDetected = true;
    }
  }

  // 同一の開示年度では、1資料につき必ず資料開示のレコードは1つであること
  function checkDocumentDisclosure() {
    var disclosureCount = 0;
    var uniqueCombinations = new Set();

    for (var row = 6; row < data.length; row++) {
      var documentName = data[row][headerIndices.documentNameCol];//資料名称
    var typeName = data[row][headerIndices.typeNameCol];//種別名
    var disclosureYear = data[row][headerIndices.disclosureYearCol];//開示年度
      
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
      var typeNameCol = headerIndices.typeNameCol;//種別名
      sheet.getRange(flagRow + 1, typeNameCol + 1).setValue(1);
      return false;
    }
    return true;
  };
  // 資料名称をキーとして、種別名ごとに「過年度：年」「過年度：年月」の組み合わせが一致するかをチェック（行の入力漏れを検知）
  function checkCombinationConsistency() {
    var baseCombinations = {};
    var typeNameCol = headerIndices.typeNameCol;//種別名
    var documentNameCol = headerIndices.documentNameCol; //資料名称
    var pastYearCol = headerIndices.pastYearCol; //過年度：年
    var pastYearMonthCol = headerIndices.pastYearMonthCol; //過年度：年月／単位（加工値）

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

             // エラー行のA列にエラーメッセージを表示する
          var row = data.findIndex(r => r[documentNameCol] === documentName && r[typeNameCol] === typeName);
          if (row !== -1) {
            sheet.getRange(row + 1, 1).setValue("一致していません");
            sheet.getRange(row + 1, 1).setBackground("tan");
          }                 
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
    var typeNameCol = headerIndices.typeNameCol; //種別名
    sheet.getRange(flagRow + 1, typeNameCol + 1).setBackground("red");
  }

  // エラー検知条件(項目共通)
  for (var row = 6; row < data.length; row++) {
    var typeName = data[row][headerIndices.typeNameCol];//種類別
    var startCol = headerIndices.startCol;//【ガバナンス】取締役人数
    
    //開示データにデータが収録されている場合、加工データにもデータが収録されていること
    if (headerIndices.startCol !== -1 && data[row][headerIndices.typeNameCol] === "開示データ") {
      for (var col = headerIndices.startCol; col < data[row].length; col++) {
        if (data[row][col] !== "") {
          var matchfound = false;
          for (var row2 = 6; row2 < data.length; row2++) {
            if (data[row][headerIndices.sourceTypeCol] === data[row2][headerIndices.sourceTypeCol] &&
              data[row][headerIndices.codeCol] === data[row2][headerIndices.codeCol] &&
              data[row][headerIndices.documentNameCol] === data[row2][headerIndices.documentNameCol] &&
              data[row][headerIndices.disclosureYearCol] === data[row2][headerIndices.disclosureYearCol] &&
              data[row][headerIndices.pastYearCol] === data[row2][headerIndices.pastYearCol] &&
              data[row][headerIndices.pastYearMonthCol] === data[row2][headerIndices.pastYearMonthCol]){
              var exclusionData = (headers.indexOf("【ガバナンス】ガバナンス体系（組織体系）")||headers.indexOf("【ガバナンス】取締役会の議長"));
              if (exclusionData){
                //加工データのチェックは行う
                if (data[row2][headerIndices.typeNameCol] === "加工データ" && data[row2][col] === "") {
                  setErrorHighlight(sheet, row2, col,flagRow);
                }
              }else{
                if ((data[row2][headerIndices.typeNameCol] === "加工データ" ||data[row2][headerIndices.typeNameCol] === "単位") && data[row2][col] === ""){
                  setErrorHighlight(sheet, row2, col,flagRow);
                }
              }
              matchfound = true;
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
        "【ガバナンス】取締役人数": { "min": 0, "max": 30 },
        "【ガバナンス】社外取締役人数": { "min": 0, "max": 30 },
        "【ガバナンス】独立社外取締役人数": { "min": 0, "max": 30 },
        "【ガバナンス】独立社外取締役比率": { "min": 0, "max": 100 },
        "【ガバナンス】監査役人数": { "min": 0, "max": 30 },
        "【ガバナンス】独立社外監査役人数": { "min": 0, "max": 30 },
        "【ガバナンス】独立社外監査役比率": { "min": 0, "max": 100 },
        "【ガバナンス】役員の固定報酬": { "min": 0, "max": 100000 },
        "【ガバナンス】役員の変動報酬": { "min": 0, "max": 100000 },
        "【ガバナンス】役員の変動報酬比率": { "min": 0, "max": 100 },
      },
      "加工データ": {
        "【ガバナンス】取締役人数": { "min": 0, "max": 30 },
        "【ガバナンス】社外取締役人数": { "min": 0, "max": 30 },
        "【ガバナンス】独立社外取締役人数": { "min": 0, "max": 30 },
        "【ガバナンス】独立社外取締役比率": { "min": 0, "max": 100 },
        "【ガバナンス】監査役人数": { "min": 0, "max": 30 },
        "【ガバナンス】独立社外監査役人数": { "min": 0, "max": 30 },
        "【ガバナンス】独立社外監査役比率": { "min": 0, "max": 100 },
        "【ガバナンス】役員の固定報酬": { "min": 0, "max": 3514 },
        "【ガバナンス】役員の変動報酬": { "min": 0, "max": 4995 },
        "【ガバナンス】役員の変動報酬比率": { "min": 0, "max": 100 },
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
            setErrorHighlight(sheet, row, col, flagRow);
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
          setErrorHighlight(sheet, row, col, flagRow);
        }
      }

    //反対に、URL以外では「https:～」の文字列を含んでいないこと
    }else if (["開示データ", "加工データ", "単位", "ページ数", "対象範囲"].includes(typeName)){
      for (var col = startCol; col < data[row].length; col++) {
        var cellValue = String(data[row][col]);
        if (cellValue !== "" && (cellValue.includes("https://")||cellValue.includes("http://"))) {
          // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
          setErrorHighlight(sheet, row, col, flagRow);
        }
      }
    };

    //数値データであること
    if (typeName == "開示データ" || typeName == "加工データ" ){
      if (startCol !== -1){
        for (var col = startCol; col < data[row].length; col++){
          var header = headers[col];
          var value = data[row][col];

          //特定の項目名の場合は数値でなくても問題ない
          if(header === "【ガバナンス】ガバナンス体系（組織体系）" ||
          header === "【ガバナンス】取締役会の議長" || 
          header === "【ガバナンス】外国人株式保有比率") {
            continue;
          }
         
          if (data[row][col] !== "" && isNaN(data[row][col])){
            //エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
            setErrorHighlight(sheet, row, col, flagRow);
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
            setErrorHighlight(sheet, row, col, flagRow);
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
           setErrorHighlight(sheet, row, col, flagRow);
          }
        }
      }
    };
  };
  // エラー行のカウントを実行
  countColoredCells(sheet);
  SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');
};