function checkDataE2021() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2021(日経E)");
  var data = sheet.getDataRange().getValues();//mitsuikeis
  var headers = data[5]; // 6行目が項目名
  var flagRow = 4; // 5行目にフラグを立てる

  // エラーチェック条件の定義
  var conditions = {
    "出典種別": function(value, row) {
      var documentType = data[row][headers.indexOf("資料名称")];
      var documentValueMap = {
        "有価証券報告書": "004",
        "コーポレートガバナンス報告書": "005",
        "企業HP": "006"
      };
      var disclosureYear = data[row][headers.indexOf("開示年度")];
      if (disclosureYear === 2021) {
        return value === "";
      } else {
        return value === documentValueMap[documentType];
      }
    },

    "種別名": function(value, row) {
      var validValues = ["資料開示", "開示データ", "加工データ", "単位", "URL", "ページ数", "対象範囲"];
      var startCol = headers.indexOf("【環境】温室効果ガス（GHG）排出量（Scope1）");

      // 資料開示のとき、収録項目のデータが空であることを確認
      if (value === "資料開示" || 
          (["URL", "ページ数", "対象範囲"].includes(value) && 
           ["有価証券報告書", "コーポレートガバナンス報告書"].includes(data[row][headers.indexOf("資料名称")]))) {
        if (startCol !== -1) {
          for (var col = startCol; col < data[row].length; col++) {
            if (data[row][col] !== "") {
              markErrorCells(row, startCol, data[row].length);
              return false;
            }
          }
        }
      }

      return validValues.includes(value);
    },

    "コード": function(value) {
      return /^[0-9]{4}$/.test(value) || /^[A-Za-z0-9]{4}$/.test(value);
    },

    "資料名称": function(value, row) {
      var typeName = data[row][headers.indexOf("種別名")];
      if (["有価証券報告書", "コーポレートガバナンス報告書"].includes(value) && 
          ["URL", "ページ数", "対象範囲"].includes(typeName)) {
        return false;
      }
      return ["企業HP", "有価証券報告書", "コーポレートガバナンス報告書"].includes(value);
    },

    "資料公表日": function(value, row) {
      var documentType = data[row][headers.indexOf("資料名称")];
      if (["有価証券報告書", "コーポレートガバナンス報告書"].includes(documentType) && value === "") {
        return false;
      }
      var datePattern = /^(?:(19|20)?\d\d年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|(?:19|20)?\d\d[-\/.](0[1-9]|1[0-2])[-\/.](0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])[-\/.](0[1-9]|1[0-2])[-\/.](19|20)?\d\d|(?:0[1-9]|1[0-2])[-\/.](0[1-9]|[12][0-9]|3[01])[-\/.](19|20)?\d\d|令和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|平成[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|昭和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|大正[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|[一二三四五六七八九十百千万]{1,4}年[一二三四五六七八九十]{1,2}月[一二三四五六七八九十]{1,2}日|(?:19|20)?\d\d年(0[1-9]|1[0-2])月|令和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|平成[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|昭和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|大正[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|[一二三四五六七八九十百千万]{1,4}年[一二三四五六七八九十]{1,2}月)|$/;
      return datePattern.test(value.toString().trim());
    },

    "開示年度": function(value) {
      return value === 2021;
    },

    "過年度：年": function(value, row) {
      if (data[row][headers.indexOf("種別名")] === "資料開示") {
        return value === "";
      }
      return /^[0-9]{4}$/.test(value);
    },

    "過年度：年月／単位（加工値）": function(value, row) {
      if (data[row][headers.indexOf("種別名")] === "資料開示") {
        return value === "";
      }
      var valueStr = value.toString();
      return /^[0-9]{6}$/.test(valueStr) && /^(00|01|02|03|04|05|06|07|08|09|10|11|12)$/.test(valueStr.slice(-2));
    }
  };

  // エラーチェックとフラグ設定
  for (var row = 6; row < data.length; row++) {
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];
      if (conditions[header] && !conditions[header](value, row)) {
        markErrorCell(row, col);
      }
    }
  }

  // 開示年度と資料名称の一致チェック
  checkYearNameConsistency();

  // 資料開示の一貫性チェック
  if (!checkDocumentDisclosure() || !checkCombinationConsistency()) {
    markErrorColumn("種別名", "red");
  }

  // エラーチェック完了メッセージ
  SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');

  // エラーセルのマーク付け関数
  function markErrorCell(row, col) {
    sheet.getRange(row + 1, col + 1).setBackground("yellow");
    sheet.getRange(flagRow + 1, col + 1).setValue(1);
  }

  // エラーレンジのマーク付け関数
  function markErrorCells(row, startCol, endCol) {
    for (var col = startCol; col < endCol; col++) {
      markErrorCell(row, col);
    }
  }

  // エラー列のマーク付け関数
  function markErrorColumn(headerName, color) {
    var colIndex = headers.indexOf(headerName);
    sheet.getRange(flagRow + 1, colIndex + 1).setBackground(color);
  }

  // 開示年度と資料名称の一致チェック関数
  function checkYearNameConsistency() {
    var dataMap = {};
    var yearIndex = headers.indexOf("開示年度");
    var nameIndex = headers.indexOf("資料名称");
    var dateIndex = headers.indexOf("資料公表日");

    for (var row = 6; row < data.length; row++) {
      var year = data[row][yearIndex];
      var name = data[row][nameIndex].trim();
      var date = data[row][dateIndex].toString().trim();
      var key = year + "_" + name;
      if (!dataMap[key]) {
        dataMap[key] = date;
      } else if (dataMap[key] !== date) {
        markErrorCell(row, dateIndex);
      }
    }
  }

  // 資料開示の一貫性チェック関数
  function checkDocumentDisclosure() {
    var disclosureCount = 0;
    var uniqueCombinations = new Set();

    for (var row = 6; row < data.length; row++) {
      var documentName = data[row][headers.indexOf("資料名称")];
      var typeName = data[row][headers.indexOf("種別名")];
      var disclosureYear = data[row][headers.indexOf("開示年度")];
      
      if (typeName === "資料開示") {
        disclosureCount++;
      }

      var combination = documentName + "_" + disclosureYear;
      uniqueCombinations.add(combination);
    }

    return uniqueCombinations.size === disclosureCount;
  }

  // 種別名ごとの一貫性チェック関数
  function checkCombinationConsistency() {
    var baseCombinations = {};

    for (var row = 6; row < data.length; row++) {
      var typeName = data[row][headers.indexOf("種別名")];
      var documentName = data[row][headers.indexOf("資料名称")];
      var pastYear = data[row][headers.indexOf("過年度：年")];
      var pastYearMonth = data[row][headers.indexOf("過年度：年月／単位（加工値）")];

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

    for (var documentName in baseCombinations) {
      var typeNameSets = baseCombinations[documentName];
      var allSets = new Set();
      for (var typeName in typeNameSets) {
        if (allSets.size === 0) {
          allSets = new Set(typeNameSets[typeName]);
        } else {
          if (allSets.size !== typeNameSets[typeName].size || 
              ![...allSets].every(value => typeNameSets[typeName].has(value))) {
            markErrorColumn("種別名", "red");
            return false;
          }
        }
      }
    }
    return true;
  }
}
