function checkDataE2023(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2023(日経E)");
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
      startCol: headers.indexOf("【環境】温室効果ガス（GHG）排出量（Scope1）"),
      environmentalReserveCol: headers.indexOf("【環境】（予備）"),
      disclosureDateCol: headers.indexOf("資料公表日")
    };

  // エラー検知してイエローに変更する関数
  function setErrorHighlight(sheet, row, col, flagRow) {
    sheet.getRange(row + 1, col + 1).setBackground("yellow");
    sheet.getRange(flagRow + 1, col + 1).setValue(1);
  }
  function setErrorHighlight2(sheet, row2, col,flagRow) {
    sheet.getRange(row2 + 1, col + 1).setBackground("yellow");
    sheet.getRange(flagRow + 1, col + 1).setValue(1);
  }
  //行のループ処理を関数化
  function iterateRows(data, startRow, callback) {
    for (var row = startRow; row < data.length; row++) {
      callback(row);
    }
  }
  //列のループ処理を関数にして共通化
  function iterateCols(data, row, startCol, callback) {
    for (var col = startCol; col < data[row].length; col++) {
      callback(col);
    }
  }

  // 項目特有のエラー検知条件を設定する
  var conditions = {
    "出典種別": function(value, row) {
      var documentType = data[row][headerIndices.documentNameCol];
      var disclosureYear = data[row][headerIndices.disclosureYearCol];
      var documentValueMap = {"有価証券報告書": "004","コーポレートガバナンス報告書": "005","企業HP": "006"};

      //開示年度2021年のときNULLであること
      if (disclosureYear === 2021) {
        return value === "";     
      } else {
        //開示年度2021年度以外は「004→有価証券報告書」「005→コーポレートガバナンス報告書」「006→企業HP」であること
        return value === documentValueMap[documentType];
      }
    },
  
    "種別名": function(value, row) {
      var validValues = ["資料開示", "開示データ", "加工データ", "単位", "URL", "ページ数", "対象範囲"];
      var isValid = validValues.includes(value);
      var startCol = headerIndices.startCol; // インデックス再利用

      //資料開示の場合、収録項目のデータ入力がないこと
      //有報・CG報告書の場合、URL、ページ数、対象範囲のデータ入力がないこと
      if (value === "資料開示" || (["URL", "ページ数", "対象範囲"].includes(value) && ["有価証券報告書", "コーポレートガバナンス報告書"].includes(data[row][headerIndices.documentNameCol]))) {
        if (headerIndices.startCol !== -1) {
          iterateCols(data, row, headerIndices.startCol, function(col) {
            if (data[row][col] !== "") {
              for (var errorCol = headerIndices.startCol; errorCol < data[row].length; errorCol++) {
                sheet.getRange(row + 1, errorCol + 1).setBackground("yellow");
              }
              return false;
            }
          })
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
      //「企業HP」であること※23年度よりEデータは企業HPのみ対象
      return ["企業HP"].includes(value);
    },
  
    "資料公表日": function(value,row) {
      //「有価証券報告書」または「コーポレートガバナンス報告書」の場合、空欄をエラーとする
      var documentType = data[row][headerIndices.documentNameCol];
      if (["有価証券報告書","コーポレートガバナンス報告書"].includes(documentType) && value === ""){
        return false;
      }
      // 文字化けしていないこと
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
      //値が「2023」であること
      return value === 2023;
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
    }
  };

 // エラー検知とフラグ設定
 iterateRows(data, 6, function(row) {
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];
      if (conditions[header] && !conditions[header](value, row)) {
        setErrorHighlight(sheet, row, col, flagRow);
        }
    }
  });

  //同一の開示年度で、資料名と資料公表日の組み合わせが一致していること。
  var indexMap = {};
  for (var i = 0; i < headers.length; i++) {
    indexMap[headers[i]] = i;
  }
  var yearIndex = headerIndices.disclosureYearCol;
  var nameIndex = headerIndices.documentNameCol;
  var dateIndex = headerIndices.disclosureDateCol;
  var dataMap = {};
  iterateRows(data, 6, function(row) {
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
  })

  // 同一の開示年度では、1資料につき必ず資料開示のレコードは1つであること
  function checkDocumentDisclosure() {
    var disclosureCount = 0;
    var uniqueCombinations = new Set();

    iterateRows(data, 6, function(row) {
    var documentName = data[row][headerIndices.documentNameCol];
    var typeName = data[row][headerIndices.typeNameCol];
    var disclosureYear = data[row][headerIndices.disclosureYearCol];
      
      if (typeName === "資料開示") {
        disclosureCount++;
        };
    
      var combination = documentName + "_" + disclosureYear;
      if(!uniqueCombinations.has(combination)){
        uniqueCombinations.add(combination);
        }
    })
    
    if (uniqueCombinations.size !== disclosureCount) {
      // 修正箇所: エラーが発生した場合の処理
      var typeNameCol = headerIndices.typeNameCol;
      sheet.getRange(flagRow + 1, typeNameCol + 1).setValue(1);
      return false;
    }
    return true;
  };

  // 資料名称をキーとして、種別名ごとに「過年度：年」「過年度：年月」の組み合わせが一致するかをチェック（行の入力漏れを検知）
  function checkCombinationConsistency() {
    var baseCombinations = {};
    var typeNameCol = headerIndices.typeNameCol; // すでに定義されたインデックスを使用
    var documentNameCol = headerIndices.documentNameCol; // 資料名称のインデックス
    var pastYearCol = headerIndices.pastYearCol; // 過年度：年のインデックス
    var pastYearMonthCol = headerIndices.pastYearMonthCol; // 過年度：年月／単位（加工値）のインデックス

    iterateRows(data, 6, function(row) {
      var typeName = data[row][typeNameCol];
      var documentName = data[row][documentNameCol];
      var pastYear = data[row][pastYearCol];
      var pastYearMonth = data[row][pastYearMonthCol];
    
      if (typeName === "資料開示") return;
    
      var combination = pastYear + "_" + pastYearMonth;
      if (!baseCombinations[documentName]) {
        baseCombinations[documentName] = {};
      }
    
      if (!baseCombinations[documentName][typeName]) {
       baseCombinations[documentName][typeName] = new Set();
      }
      baseCombinations[documentName][typeName].add(combination);
    })
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
            console.log(`Mismatch found in document: ${documentName}, type: ${typeName}`);
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
  var typeNameCol = headerIndices.typeNameCol; // インデックス再利用
  sheet.getRange(flagRow + 1, typeNameCol + 1).setBackground("red");
}

  // エラー検知条件(項目共通)
iterateRows(data, 6, function(row) {
  var typeName = data[row][headerIndices.typeNameCol]; // インデックス再利用
  var startCol = headerIndices.startCol; // インデックス再利用
    
    //開示データにデータが収録されている場合、加工データにもデータが収録されていること
    if (startCol !== -1 && typeName === "開示データ") {
      iterateCols(data, row, startCol, function(col) {
        if (data[row][col] !== "") {
          var matchfound = false;
          for (var row2 = 6; row2 < data.length; row2++) {
            if (
              data[row][headerIndices.sourceTypeCol] === data[row2][headerIndices.sourceTypeCol] && // 出典種別
              data[row][headerIndices.codeCol] === data[row2][headerIndices.codeCol] && // コード
              data[row][headerIndices.documentNameCol] === data[row2][headerIndices.documentNameCol] && // 資料名称
              data[row][headerIndices.disclosureYearCol] === data[row2][headerIndices.disclosureYearCol] && // 開示年度
              data[row][headerIndices.pastYearCol] === data[row2][headerIndices.pastYearCol] && // 過年度：年
              data[row][headerIndices.pastYearMonthCol] === data[row2][headerIndices.pastYearMonthCol] && // 過年度：年月／単位（加工値）
              (data[row2][headerIndices.typeNameCol] === "加工データ" || data[row2][headerIndices.typeNameCol] === "単位") // 種別名
            ) {
              if (data[row2][col] === "") {
                setErrorHighlight2(sheet, row2, col, flagRow);
              }
              matchfound = true;
              if (data[row2][col] === ""){
                setErrorHighlight2(sheet, row2, col, flagRow);
              }
            }
          }
          if (!matchfound){
            console.log("No match found for row: " + row + ", col: " + col);
          }
        }
      })
    };
    
    // 各項目の閾値設定(過去データの最大値・最小値から決定　※仮決め)
    var thresholds = {
      "開示データ": {
        "【環境】温室効果ガス（GHG）排出量（Scope1）": { "min": 0, "max": 21828013 },
        "【環境】温室効果ガス（GHG）排出量（Scope2）": { "min": 0, "max": 4090000 },
        "【環境】温室効果ガス（GHG）排出量（Scope3）": { "min": 0, "max": 309820000 },
        "【環境】温室効果ガス（GHG）総排出量": { "min": 0, "max": 315290000 },
        "【環境】温室効果ガス（GHG）総排出量対売上高原単位": { "min": 0, "max": 1690 },
        "【環境】Scope3上流": { "min": 0, "max": 4745613.8 },
        "【環境】Scope3下流": { "min": 0, "max": 6195386 },
        "【環境】Scope3カテゴリ１：購入した製品・サービス": { "min": 0, "max": 15165310 },
        "【環境】Scope3カテゴリ２：資本財": { "min": 0, "max": 1767060 },
        "【環境】Scope3カテゴリ３：Scope1,2に含まれない燃料及びエネルギー関連活動": { "min": 0, "max": 5370828 },
        "【環境】Scope3カテゴリ４：輸送、配送（上流）": { "min": 0, "max": 1354349 },
        "【環境】Scope3カテゴリ５：事業活動から出る廃棄物": { "min": 0, "max": 369119 },
        "【環境】Scope3カテゴリ６：出張": { "min": 0, "max": 79417 },
        "【環境】Scope3カテゴリ７：雇用者の通勤": { "min": 0, "max": 1176132 },
        "【環境】Scope3カテゴリ８：リース資産（上流）": { "min": 0, "max": 428056 },
        "【環境】Scope3カテゴリ９：輸送、配送（下流）": { "min": 0, "max": 1038586 },
        "【環境】Scope3カテゴリ10：販売した製品の加工": { "min": 0, "max": 600000 },
        "【環境】Scope3カテゴリ11：販売した製品の使用": { "min": 0, "max": 225549245 },
        "【環境】Scope3カテゴリ12：販売した製品の廃棄": { "min": 0, "max": 2272581 },
        "【環境】Scope3カテゴリ13：リース資産（下流）": { "min": 0, "max": 798946 },
        "【環境】Scope3カテゴリ14：フランチャイズ": { "min": 0, "max": 1221525 },
        "【環境】Scope3カテゴリ15：投資": { "min": 0, "max": 28515955 },
        "【環境】水総使用量（消費量）": { "min": 0, "max": 60736000 },
        "【環境】再生水使用量（消費量）": { "min": 0, "max": 1497013 },
        "【環境】総排水量": { "min": 0, "max": 197186000 },
        "【環境】組織内エネルギー消費量": { "min": 0, "max": 11259535 },
        "【環境】組織外エネルギー消費量": { "min": 0, "max": 726027203 },
        "【環境】総エネルギー消費量": { "min": 0, "max": 2273483738 },
        "【環境】総エネルギー消費量対売上高原単位": { "min": 0, "max": 5890.9 },
        "【環境】廃棄物の発生量": { "min": 0, "max": 7884635 }
      },
      "加工データ": {
        "【環境】温室効果ガス（GHG）排出量（Scope1）": { "min": 0, "max": 89578000 },
        "【環境】温室効果ガス（GHG）排出量（Scope2）": { "min": 0, "max": 242827000 },
        "【環境】温室効果ガス（GHG）排出量（Scope3）": { "min": 0, "max": 1578348000 },
        "【環境】温室効果ガス（GHG）総排出量": { "min": 0, "max": 447970000 },
        "【環境】温室効果ガス（GHG）総排出量対売上高原単位": { "min": 0, "max": 17000 },
        "【環境】Scope3上流": { "min": 0, "max": 121840000 },
        "【環境】Scope3下流": { "min": 0, "max": 1575000000 },
        "【環境】Scope3カテゴリ１：購入した製品・サービス": { "min": 0, "max": 110490000 },
        "【環境】Scope3カテゴリ２：資本財": { "min": 0, "max": 6280000 },
        "【環境】Scope3カテゴリ３：Scope1,2に含まれない燃料及びエネルギー関連活動": { "min": 0, "max": 112535000 },
        "【環境】Scope3カテゴリ４：輸送、配送（上流）": { "min": 0, "max": 58006000 },
        "【環境】Scope3カテゴリ５：事業活動から出る廃棄物": { "min": 0, "max": 26276000 },
        "【環境】Scope3カテゴリ６：出張": { "min": 0, "max": 27312000 },
        "【環境】Scope3カテゴリ７：雇用者の通勤": { "min": 0, "max": 13931000 },
        "【環境】Scope3カテゴリ８：リース資産（上流）": { "min": 0, "max": 776000 },
        "【環境】Scope3カテゴリ９：輸送、配送（下流）": { "min": 0, "max": 26585000 },
        "【環境】Scope3カテゴリ10：販売した製品の加工": { "min": 0, "max": 60016000 },
        "【環境】Scope3カテゴリ11：販売した製品の使用": { "min": 0, "max": 1575000000 },
        "【環境】Scope3カテゴリ12：販売した製品の廃棄": { "min": 0, "max": 64507000 },
        "【環境】Scope3カテゴリ13：リース資産（下流）": { "min": 0, "max": 3861000 },
        "【環境】Scope3カテゴリ14：フランチャイズ": { "min": 0, "max": 4650000 },
        "【環境】Scope3カテゴリ15：投資": { "min": 0, "max": 36000000 },
        "【環境】水総使用量（消費量）": { "min": 0, "max": 68843000000 },
        "【環境】再生水使用量（消費量）": { "min": 0, "max": 5800000000 },
        "【環境】総排水量": { "min": 0, "max": 108379666000  },
        "【環境】組織内エネルギー消費量": { "min": 0, "max": 11259.535 },
        "【環境】組織外エネルギー消費量": { "min": 0, "max": 2613697.931 },
        "【環境】総エネルギー消費量": { "min": 0, "max": 5160000 },
        "【環境】総エネルギー消費量対売上高原単位": { "min": 0, "max": 119.62721 },
        "【環境】廃棄物の発生量": { "min": 0, "max": 299243000 }
      }
    };
    // 各データが閾値の範囲内であることをチェック
    function checkThresholds(typeName, data, headers, row, startCol, sheet, flagRow) {
      var thresholdData = thresholds[typeName];
      iterateCols(data, row, startCol, function(col) {
        if (data[row][col] !== "") {
          var header = headers[col];
          var thresholdRange = thresholdData[header];
          if (!thresholdRange) {
            return;
          }
          var value = parseFloat(data[row][col]);
          if (isNaN(value) || value < thresholdRange.min || value > thresholdRange.max) {
            setErrorHighlight(sheet, row, col, flagRow);
          }
        }
      })
    }
    // 閾値チェックを実行
    if (typeName === "開示データ" || typeName === "加工データ") {
      checkThresholds(typeName, data, headers, row, startCol, sheet, flagRow);
    }

    //「https:～」の文字列を含むこと
    if (typeName === "URL"){
      iterateCols(data, row, startCol, function(col) {
        var cellValue = String(data[row][col]);
        if(cellValue !== "" && !(cellValue.includes("https://")||cellValue.includes("http://"))){
          // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
          setErrorHighlight(sheet, row, col, flagRow);
        }
      })

    //反対に、URL以外では「https:～」の文字列を含んでいないこと
    }else if (["開示データ", "加工データ", "単位", "ページ数", "対象範囲"].includes(typeName)){
      iterateCols(data, row, startCol, function(col) {
        var cellValue = String(data[row][col]);
        if (cellValue !== "" && (cellValue.includes("https://")||cellValue.includes("http://"))) {
          // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
          setErrorHighlight(sheet, row, col, flagRow);
        }
      })
    };

    //数値データであること
    if (typeName == "開示データ" || typeName == "加工データ" ){
      if (startCol !== -1){
        iterateCols(data, row, startCol, function(col) {
          if (data[row][col] !== "" && isNaN(data[row][col])){
            //エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
            setErrorHighlight(sheet, row, col, flagRow);
          }
        })
      }
    }

    //ページ数は特定のフォーマットであること（数値、カンマ、ハイフンで構成）
    if (typeName == "ページ数"){
      if (startCol !== -1){
        iterateCols(data, row, startCol, function(col) {
          if(data[row][col] !== ""){
            var value = data[row][col];
            var isValid = !isNaN(value) || /^[0-9,.-]+$/.test(value);
            if (!isValid) {
            // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
            setErrorHighlight(sheet, row, col, flagRow);
            }
          }
        })
      }
    }

    //文字列になっていること
    if (typeName == "単位" || typeName == "対象範囲"){
      if(startCol !== -1){
        iterateCols(data, row, startCol, function(col) {
          if(data[row][col] !== "" && !isNaN(data[row][col])){
           // エラー検知時に該当するセルの背景色を色塗りし、列の5行目に1を入力
           setErrorHighlight(sheet, row, col, flagRow);
          }
        })
      }
    };

    //予備項目にデータが収録されていないことを確認
    var environmentalReserveCol = headerIndices.environmentalReserveCol;  // 修正：インデックスを再利用
    if (environmentalReserveCol !== -1 && data[row][environmentalReserveCol] !== "") {
      sheet.getRange(row + 1, environmentalReserveCol + 1).setBackground("yellow");
      sheet.getRange(flagRow + 1, environmentalReserveCol + 1).setValue(1);
    };
  });
  
  // 過年度年月が「00」の時に、過年度年と過年度年月の西暦が一致していること
iterateRows(data, 6, function(row) {
  var pastYearCol = headerIndices.pastYearCol;  // 修正：インデックスを再利用
  var pastYearMonthCol = headerIndices.pastYearMonthCol;  // 修正：インデックスを再利用
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
  });

  // 「開示データ」または「加工データ」は必ず1行は存在すること。
// キー項目のインデックスを取得
var keyIndices = [
  headerIndices.sourceTypeCol,  // 修正：インデックスを再利用
  headerIndices.codeCol,  // 修正：インデックスを再利用
  headerIndices.disclosureYearCol,  // 修正：インデックスを再利用
  headerIndices.pastYearCol,  // 修正：インデックスを再利用
  headerIndices.pastYearMonthCol  // 修正：インデックスを再利用
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
  iterateRows(data, 6, function(row) {
    var key = createKey(row);
    var type = data[row][headerIndices.typeNameCol];  // 修正：インデックスを再利用
    if (!keyMap[key]) {
      keyMap[key] = {
        "開示データ": false,
        "加工データ": false
      };
    }
    if (type === "開示データ" || type === "加工データ") {
      keyMap[key][type] = true;
    }
  })

  // エラー検知とフラグ設定
  iterateRows(data, 6, function(row) {
    var key = createKey(row);
    var type = data[row][headerIndices.typeNameCol];  // 修正：インデックスを再利用

    if (type === "開示データ" || type === "加工データ") {
      if (!keyMap[key]["開示データ"] || !keyMap[key]["加工データ"]) {
        for (var col = 0; col < headers.length; col++) {
          sheet.getRange(row + 1, col + 1).setBackground("yellow");
        }
        var flagCol = headerIndices.typeNameCol;  // 修正：インデックスを再利用
        sheet.getRange(flagRow + 1, flagCol + 1).setValue(1);
        sheet.getRange(flagRow + 1, flagCol + 1).setBackground("red");
      }
    }
  });

  //キー項目が同じレコードがないこと（重複していないこと）
  // 「出典種別」「種別名」「コード」「開示年度」「過年度：年」「過年度：年月／単位（加工値）」の列の値が完全に一致している行が複数ある場合
  var uniqueRows = {};
  var duplicateRows = [];

  iterateRows(data, 6, function(row) {
    var key = [
      data[row][headerIndices.sourceTypeCol],  // 出典種別
      data[row][headerIndices.typeNameCol],    // 種別名
      data[row][headerIndices.codeCol],        // コード
      data[row][headerIndices.disclosureYearCol], // 開示年度
      data[row][headerIndices.pastYearCol],    // 過年度：年
      data[row][headerIndices.pastYearMonthCol], // 過年度：年月／単位（加工値）
      data[row][headerIndices.disclosureDateCol]  // 修正: 資料公表日をインデックスで使用
    ].join("|");

    if (uniqueRows[key]) {
      duplicateRows.push(row);
      duplicateRows.push(uniqueRows[key]);
    } else {
      uniqueRows[key] = row;
    }
  })

  // 重複行が見つかった場合、5行目5列目のセルの背景色を赤色にし、そのセルに1を入力し、該当する行の2列目のセルを赤色にする
  if (duplicateRows.length > 0) {
    sheet.getRange(5, 2).setBackground("red").setValue(1);
    duplicateRows.forEach(function(row) {
      sheet.getRange(row + 1, 2).setBackground("red");
    });
  }
  
  SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');
};