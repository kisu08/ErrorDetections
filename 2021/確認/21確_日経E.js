function checkDataE2021() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2021(日経E)");
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

  // エラー検知してイエローに変更する関数
  function setErrorHighlight(sheet, row, col, flagRow) {
    sheet.getRange(row + 1, col + 1).setBackground("yellow");  // エラーセルの背景色を黄色に変更
    sheet.getRange(flagRow + 1, col + 1).setValue(1);  // フラグ行に1をセット
    sheet.getRange(row + 1, 1).setValue("入力に不備があります");  // A列にエラーメッセージをセット
    sheet.getRange(row + 1, 1).setBackground("orange");
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
      var documentType = data[row][headerIndices.documentNameCol];//資料名称
      var disclosureYear = data[row][headerIndices.disclosureYearCol];//開示年度
      var documentValueMap = {"有価証券報告書": "004","コーポレートガバナンス報告書": "005","企業HP": "006"};

      //開示年度2021年のときNULLであること
      if (disclosureYear === 2021) {
        return value === "";
      } else {
        return value === documentValueMap[documentType];
      }
    },

    "種別名": function(value, row) {
      var validValues = ["資料開示", "開示データ", "加工データ", "単位", "URL", "ページ数", "対象範囲"];
      var isValid = validValues.includes(value);
      var startCol = headerIndices.startCol; // 【環境】温室効果ガス（GHG）排出量（Scope1）

      //資料開示の場合、収録項目のデータ入力がないこと
      //有報・CG報告書の場合、URL、ページ数、対象範囲のデータ入力がないこと
      if (value === "資料開示" || (["URL", "ページ数", "対象範囲"].includes(value) && ["有価証券報告書", "コーポレートガバナンス報告書"].includes(data[row][headerIndices.documentNameCol]))) {
        if (headerIndices.startCol !== -1) {
          iterateCols(data, row, headerIndices.startCol, function(col) {
            if (data[row][col] !== "") {
              for (var errorCol = headerIndices.startCol; errorCol < data[row].length; errorCol++) {
                sheet.getRange(row + 1, errorCol + 1).setBackground("yellow");
                sheet.getRange(row + 1, 1).setValue("不正なデータです");  // A列にエラーメッセージをセット
                sheet.getRange(row + 1, 1).setBackground("orange");
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
      //「有価証券報告書」「コーポレートガバナンス報告書」「企業HP」のいずれかであること
      return ["企業HP", "有価証券報告書", "コーポレートガバナンス報告書"].includes(value);
    },
  
    "資料公表日": function(value,row) {
      //「有価証券報告書」または「コーポレートガバナンス報告書」の場合、空欄をエラーとする
      var documentType = data[row][headerIndices.documentNameCol];//資料名称
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
      return value === 2021;
    },

    "過年度：年": function(value, row) {
      // 種別名が資料開示のときNULLであること
      if (data[row][headerIndices.typeNameCol] === "資料開示") {
        return value === "";
      }
      // 4桁の数字であること
      if (!/^[0-9]{4}$/.test(value)) {
        return false;
      }

      // 数字の打ち間違いで異常に大きいまたは小さい年を検出（2000～2022年の範囲内）
      var year = parseInt(value, 10);
      return year >= 2000 && year <= 2022;
      },
    
    "過年度：年月／単位（加工値）": function(value, row) {
      // 種別名が資料開示のときNULLであること
      if (data[row][headerIndices.typeNameCol] === "資料開示") {
        return value === "";
      }
      
      // 6桁の数字で末尾が「00」～「12」であること
      var valueStr = value.toString();
      if (!/^[0-9]{6}$/.test(valueStr) || !/^(00|01|02|03|04|05|06|07|08|09|10|11|12)$/.test(valueStr.slice(-2))) {
        return false;
      }

      // 年月の年部分が2000～2022年の範囲内であること
      var yearMonth = parseInt(valueStr.slice(0, 4), 10);
      if (yearMonth < 2000 || yearMonth > 2022) {
        return false;
      }

      // 過年度と過年度年月の整合性チェック
      var pastYearValue = data[row][headerIndices.pastYearCol];
      var pastYear = parseInt(pastYearValue, 10);
      var month = parseInt(valueStr.slice(-2), 10); // 月部分を取り出す

      // 過年度年月の年が過年度年よりも下回る場合はエラー
      if (yearMonth < pastYear) {
        return false;
      }

      // 月が12以外の場合、過年度と過年度年月の年が1年違っていても許容する
      if (month !== 12 && Math.abs(pastYear - yearMonth) > 1) {
        return false;
      }

      // 月が12の場合、過年度と過年度年月の年が一致することを確認
      if (month === 12 && pastYear !== yearMonth) {
        return false;
      }

      return true;
    }, 
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
  var yearIndex = headerIndices.disclosureYearCol;//開示年度
  var nameIndex = headerIndices.documentNameCol;//資料名称
  var dateIndex = headerIndices.disclosureDateCol;//資料公表日
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
      sheet.getRange(row + 1, 1).setValue("公表日一致していません");  // A列にエラーメッセージをセット
      sheet.getRange(row + 1, 1).setBackground("tan");
      errorDetected = true;
    }
  })

  // 同一の開示年度では、1資料につき必ず資料開示のレコードは1つであること
  function checkDocumentDisclosure() {
    var disclosureCount = 0;
    var uniqueCombinations = new Set();

    iterateRows(data, 6, function(row) {
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
    })
    
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
    var typeNameCol = headerIndices.typeNameCol; //種別名
    var documentNameCol = headerIndices.documentNameCol; //資料名称
    var pastYearCol = headerIndices.pastYearCol; //過年度：年
    var pastYearMonthCol = headerIndices.pastYearMonthCol; //過年度：年月／単位（加工値）

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
iterateRows(data, 6, function(row) {
  var typeName = data[row][headerIndices.typeNameCol]; //種別名
  var startCol = headerIndices.startCol; //【環境】温室効果ガス（GHG）排出量（Scope1）
    
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
                setErrorHighlight(sheet, row2, col, flagRow);
              }
              matchfound = true;
              if (data[row2][col] === ""){
                setErrorHighlight(sheet, row2, col, flagRow);
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
    var environmentalReserveCol = headerIndices.environmentalReserveCol;  //【環境】（予備）
    if (environmentalReserveCol !== -1 && data[row][environmentalReserveCol] !== "") {
      sheet.getRange(row + 1, environmentalReserveCol + 1).setBackground("yellow");
      sheet.getRange(flagRow + 1, environmentalReserveCol + 1).setValue(1);
      sheet.getRange(row + 1, 1).setValue("予備項目にデータがあります");  // A列にエラーメッセージをセット
      sheet.getRange(row + 1, 1).setBackground("orange");
    };
  });
  // エラー行のカウントを実行
  countColoredCells(sheet);
  SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');
};