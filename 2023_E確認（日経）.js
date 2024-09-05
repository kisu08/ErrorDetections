function checkDataE2023() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2023(日経E)");
  var data = sheet.getDataRange().getValues();
  var headers = data[5]; // 6行目が項目名
  var flagRow = 4; // 5行目にフラグを立てる

  // セルの色を塗る関数
  function setError(row, col, color, flagValue) {
    sheet.getRange(row + 1, col + 1).setBackground(color);
    sheet.getRange(flagRow + 1, col + 1).setValue(flagValue);
  }

  // エラー条件をチェックする汎用関数
  function checkCondition(row, col, condition, color = "yellow") {
    if (!condition) {
      setError(row, col, color, 1);
    }
  }

  // 開示年度2021年のときNULLであること
  // 開示年度2021年度以外は「004→有価証券報告書」「005→コーポレートガバナンス報告書」「006→企業HP」であること
  function checkDisclosureYear(value, row) {
    var documentType = data[row][headers.indexOf("資料名称")];
    var disclosureYear = data[row][headers.indexOf("開示年度")];
    var documentValueMap = {
      "有価証券報告書": "004",
      "コーポレートガバナンス報告書": "005",
      "企業HP": "006"
    };

    if (disclosureYear === 2021) {
      return value === "";
    } else {
      return value === documentValueMap[documentType];
    }
  }

  // 閾値をチェックする関数
  function checkThresholds(typeName, data, headers, row, startCol) {
    var thresholdData = thresholds[typeName];
    for (var col = startCol; col < data[row].length; col++) {
      var value = parseFloat(data[row][col]);
      if (
        data[row][col] !== "" &&
        thresholdData[headers[col]] &&
        (isNaN(value) || value < thresholdData[headers[col]].min || value > thresholdData[headers[col]].max)
      ) {
        setError(row, col, "yellow", 1);
      }
    }
  }

  // キーを生成する関数
  function generateKey(row) {
    var keyCols = ["出典種別", "コード", "開示年度", "過年度：年", "過年度：年月／単位（加工値）"];
    return keyCols.map(function (col) {
      return data[row][headers.indexOf(col)];
    }).join("|");
  }

  // 閾値設定
  var thresholds = {
    "開示データ": {
      "【環境】温室効果ガス（GHG）排出量（Scope1）": { "min": 0, "max": 21828013 },
      "【環境】温室効果ガス（GHG）排出量（Scope2）": { "min": 0, "max": 4090000 },
      // 省略 (他の項目も同様に定義)
    },
    "加工データ": {
      "【環境】温室効果ガス（GHG）排出量（Scope1）": { "min": 0, "max": 89578000 },
      // 省略 (他の項目も同様に定義)
    }
  };

  // keyMapを初期化
  var keyMap = {};

  // デバッグ用にbaseCombinationsの内容を出力
  function debugCombinations(baseCombinations) {
    console.log(JSON.stringify(baseCombinations, null, 2));
  }

  // エラー検出ループ
  for (var row = 6; row < data.length; row++) {
    var typeName = data[row][headers.indexOf("種別名")];
    var startCol = headers.indexOf("【環境】温室効果ガス（GHG）排出量（Scope1）");

    // 種別名によるエラーチェック
    if (["開示データ", "加工データ"].includes(typeName)) {
      checkThresholds(typeName, data, headers, row, startCol);
    }

    // 資料開示の場合、収録項目のデータ入力がないこと
    // 有報・CG報告書の場合、URL、ページ数、対象範囲のデータ入力がないこと
    if (value === "資料開示" || (["URL", "ページ数", "対象範囲"].includes(value) && ["有価証券報告書", "コーポレートガバナンス報告書"].includes(data[row][headers.indexOf("資料名称")]))) {
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
  }

    // URL のチェック
    if (typeName === "URL") {
      for (var col = startCol; col < data[row].length; col++) {
        var cellValue = String(data[row][col]);
        checkCondition(row, col, cellValue.includes("https://") || cellValue.includes("http://"));
      }
    } else if (["開示データ", "加工データ", "単位", "ページ数", "対象範囲"].includes(typeName)) {
      for (var col = startCol; col < data[row].length; col++) {
        var cellValue = String(data[row][col]);
        checkCondition(row, col, !(cellValue.includes("https://") || cellValue.includes("http://")));
      }
    }

    //ページ数は特定のフォーマットであること（数値、カンマ、ハイフンで構成）
    if (typeName == "ページ数"){
      if (startCol !== -1){
        for (var col = startCol; col < data[row].length; col++){
          if(data[row][col] !== ""){
            var value = data[row][col];
            var isValid = !isNaN(value) || /^[0-9,.-]+$/.test(value);//ピリオドを消す
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

    // キー項目が同じレコードの重複チェック
    var uniqueRows = {};
    var duplicateRows = [];
    var key = generateKey(row);

    if (keyMap[key]) {
      duplicateRows.push(row);
      duplicateRows.push(keyMap[key]);
    } else {
      keyMap[key] = row; // keyMapに登録
    }

    // 重複が見つかった場合、エラーをフラグする
    if (duplicateRows.length > 0) {
      sheet.getRange(5, 2).setBackground("red").setValue(1);
      duplicateRows.forEach(function (row) {
        sheet.getRange(row + 1, 2).setBackground("red");
      });
    }

    // エラー検知とフラグ設定（行ずれ・データ不備）
    // 修正箇所: エラーが発生した場合の処理

    // デバッグ用ログ
    console.log(`Checking document: ${documentName}, type: ${typeName}`);
    console.log(`All sets: ${[...allSets].join(", ")}`);
    console.log(`Current set: ${[...typeNameSets[typeName]].join(", ")}`);
    console.log(`Mismatch found in document: ${documentName}, type: ${typeName}`);
    console.log(`No match found for row: ${row}, col: ${col}`);
  }
