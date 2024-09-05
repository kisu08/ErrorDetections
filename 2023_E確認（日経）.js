function checkDataE2023() {
  try {
    // 1. データ取得
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2023(日経E)");
    var data = sheet.getDataRange().getValues();  // 全データを取得
    var headers = data[5];  // 6行目にヘッダーがある
    var flagRow = 4;  // 5行目にフラグを立てる

    // 2. 条件チェック（バリデーション条件の定義）
    var conditions = {
      "出典種別": function(value, row) {
        var documentType = data[row][headers.indexOf("資料名称")];
        var disclosureYear = data[row][headers.indexOf("開示年度")];
        var documentValueMap = {"有価証券報告書": "004", "コーポレートガバナンス報告書": "005", "企業HP": "006"};
        return disclosureYear === 2021 ? value === "" : value === documentValueMap[documentType];
      },
      "種別名": function(value, row) {
  var validValues = ["資料開示", "開示データ", "加工データ", "単位", "URL", "ページ数", "対象範囲"];
  if (!validValues.includes(value)) {
    // エラーが発生した場合、該当するセルの背景色を赤色に設定
    sheet.getRange(row + 1, headers.indexOf("種別名") + 1).setBackground("red");
    return false;
  }

  var startCol = headers.indexOf("【環境】温室効果ガス（GHG）排出量（Scope1）");
  if (value === "資料開示" || (["URL", "ページ数", "対象範囲"].includes(value) &&
    ["有価証券報告書", "コーポレートガバナンス報告書"].includes(data[row][headers.indexOf("資料名称")]))) {
    for (var col = startCol; col < data[row].length; col++) {
      if (data[row][col] !== "") {
        sheet.getRange(row + 1, col + 1).setBackground("yellow");
        return false;
      }
    }
  }
  return true;
},

      "コード": value => /^[0-9]{4}$/.test(value) || /^[A-Za-z0-9]{4}$/.test(value),
      "資料名称": function(value, row) {
        if (["有価証券報告書", "コーポレートガバナンス報告書"].includes(value) && 
            ["URL", "ページ数", "対象範囲"].includes(data[row][headers.indexOf("種別名")])) return false;
        return value === "企業HP";
      },
       "資料公表日": function(value,row) {
      //「有価証券報告書」または「コーポレートガバナンス報告書」の場合、空欄をエラーとする
      var documentType = data[row][headers.indexOf("資料名称")];
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

      "開示年度": value => value === 2023,
      "過年度：年": function(value, row) {
        return data[row][headers.indexOf("種別名")] === "資料開示" ? value === "" : /^[0-9]{4}$/.test(value);
      },
      "過年度：年月／単位（加工値）": function(value, row) {
        var valueStr = value.toString();
        return data[row][headers.indexOf("種別名")] === "資料開示" ? value === "" : /^[0-9]{6}$/.test(valueStr) && /^(00|01|02|03|04|05|06|07|08|09|10|11|12)$/.test(valueStr.slice(-2));
      }
    };

    // 3. エラー検知とフラグ設定
    data.slice(6).forEach((rowData, rowIndex) => {
      rowData.forEach((value, colIndex) => {
        var header = headers[colIndex];
        if (conditions[header] && !conditions[header](value, rowIndex + 6)) {
          sheet.getRange(rowIndex + 7, colIndex + 1).setBackground("yellow");
          sheet.getRange(flagRow + 1, colIndex + 1).setValue(1);
        }
      });
    });

    // 4. 重複チェック
    var uniqueRows = {}, duplicateRows = [];
    data.slice(6).forEach((rowData, rowIndex) => {
      var key = ["出典種別", "種別名", "コード", "開示年度", "過年度：年", "過年度：年月／単位（加工値）"]
                .map(header => rowData[headers.indexOf(header)]).join("|");
      if (uniqueRows[key]) {
        duplicateRows.push(rowIndex + 6, uniqueRows[key]);
      } else {
        uniqueRows[key] = rowIndex + 6;
      }
    });

    // 5. 重複行のハイライト
    if (duplicateRows.length > 0) {
      sheet.getRange(5, 2).setBackground("red").setValue(1);
      duplicateRows.forEach(row => {
        sheet.getRange(row + 1, 2).setBackground("red");
      });
    }

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

    function checkThresholds(typeName, data, headers, row, startCol, sheet, flagRow) {
      var thresholdData = thresholds[typeName];
      for (var col = startCol; col < data[row].length; col++) {
        if (data[row][col] !== "") {
          var header = headers[col];
          var thresholdRange = thresholdData[header];
          if (!thresholdRange) continue;
          var value = parseFloat(data[row][col]);
          if (isNaN(value) || value < thresholdRange.min || value > thresholdRange.max) {
            sheet.getRange(row + 1, col + 1).setBackground("yellow");
            sheet.getRange(flagRow + 1, col + 1).setValue(1);
          }
        }
      }
    }

    // 7. 閾値チェック実行
    data.slice(6).forEach((rowData, rowIndex) => {
      var typeName = rowData[headers.indexOf("種別名")];
      var startCol = headers.indexOf("【環境】温室効果ガス（GHG）排出量（Scope1）");
      if (typeName === "開示データ" || typeName === "加工データ") {
        checkThresholds(typeName, data, headers, rowIndex + 6, startCol, sheet, flagRow);
      }
    });

    // 8. 処理完了通知
    SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');

  } catch (error) {
    // エラーハンドリング
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
  }
}
