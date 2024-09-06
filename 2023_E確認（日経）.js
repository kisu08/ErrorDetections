function checkDataE2023() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2023(日経E)"); // 確認用シートを取得
  var data = sheet.getDataRange().getValues(); // シート全体のデータを取得
  var headers = data[5]; // 6行目に項目名があると仮定して取得
  var flagRow = 4; // フラグを立てる行番号（5行目）

  // 環境データのスタート列を取得
  var startCol = headers.indexOf("【環境】温室効果ガス（GHG）排出量（Scope1）");

  // 重複チェックやデータペア確認用のマップを初期化
  var keyMap = {};  
  var dataMap = {};

  /**
   * セルにエラーを設定する関数
   * @param {number} row - 行番号
   * @param {number} col - 列番号
   * @param {string} color - 背景色
   * @param {number} flagValue - フラグ値
   */
  function setError(row, col, color, flagValue) {
    sheet.getRange(row + 1, col + 1).setBackground(color); // 指定セルに色を設定
    sheet.getRange(flagRow + 1, col + 1).setValue(flagValue); // フラグ行にフラグ値を設定
  }

  /**
   * 条件を満たさない場合にエラーを設定する関数
   * @param {number} row - 行番号
   * @param {number} col - 列番号
   * @param {boolean} condition - チェック条件
   * @param {string} color - 背景色 (デフォルト: "yellow")
   */
  function checkCondition(row, col, condition, color = "yellow") {
    if (!condition) {
      setError(row, col, color, 1); // 条件がfalseの場合、エラー設定
    }
  }

  /**
   * 出典種別のチェック関数
   * @param {string} value - 出典種別の値
   * @param {number} row - 行番号
   * @returns {boolean} - 出典種別が正しいかどうか
   */
  function checkSourceType(value, row) {
    var documentType = data[row][headers.indexOf("資料名称")]; // 資料名称を取得
    var disclosureYear = data[row][headers.indexOf("開示年度")]; // 開示年度を取得
    var documentValueMap = { "有価証券報告書": "004", "コーポレートガバナンス報告書": "005", "企業HP": "006" }; // 出典種別と資料名称の対応表

    if (disclosureYear === 2021) {
      return value === ""; // 2021年は空欄であるべき
    } else {
      return value === documentValueMap[documentType]; // 他の年度ではマップされた値と一致するかチェック
    }
  }

  /**
   * 日付フォーマットのチェック関数
   * @param {number} row - 行番号
   * @param {number} col - 列番号
   */
  function checkDateFormat(row, col) {
    var dateValue = data[row][col]; // 日付の値を取得
    checkCondition(row, col, !isNaN(new Date(dateValue).getTime())); // 有効な日付形式かどうかをチェック
  }

  /**
   * 資料公表日のチェック関数
   * @param {string} value - 資料公表日の値
   * @param {number} row - 行番号
   * @returns {boolean} - 資料公表日が正しいかどうか
   */
  function checkPublicationDate(value, row) {
    var documentType = data[row][headers.indexOf("資料名称")]; // 資料名称を取得
    if (["有価証券報告書", "コーポレートガバナンス報告書"].includes(documentType) && value === "") {
      return false; // 有報やCG報告書は空欄不可
    }
    var valueStr = value.toString().trim(); // 日付を文字列としてトリム
    var datePattern = /^(?:(19|20)?\d\d年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|(?:19|20)?\d\d[-\/.](0[1-9]|1[0-2])[-\/.](0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])[-\/.](0[1-9]|1[0-2])[-\/.](19|20)?\d\d|(?:0[1-9]|1[0-2])[-\/.](0[1-9]|[12][0-9]|3[01])[-\/.](19|20)?\d\d|令和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|平成[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|昭和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|大正[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|[一二三四五六七八九十百千万]{1,4}年[一二三四五六七八九十]{1,2}月[一二三四五六七八九十]{1,2}日|(?:19|20)?\d\d年(0[1-9]|1[0-2])月|令和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|平成[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|昭和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|大正[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|[一二三四五六七八九十百千万]{1,4}年[一二三四五六七八九十]{1,2}月)|$/
    return datePattern.test(valueStr); // 正しいフォーマットかチェック
  }

  /**
   * 過年度：年のチェック関数
   * @param {string} value - 過年度：年の値
   * @param {number} row - 行番号
   * @returns {boolean} - 過年度：年が正しいかどうか
   */
  function checkPastYear(value, row) {
    if (data[row][headers.indexOf("種別名")] === "資料開示") {
      return value === ""; // 資料開示の場合は空欄であるべき
    }
    return /^[0-9]{4}$/.test(value); // 4桁の数字かどうかをチェック
  }

  /**
   * キーを生成する関数
   * @param {number} row - 行番号
   * @returns {string} - 行に対応するユニークなキー
   */
  function generateKey(row) {
    var keyCols = ["出典種別", "コード", "開示年度", "過年度：年", "過年度：年月／単位（加工値）"]; // キー項目
    return keyCols.map(function (col) {
      return data[row][headers.indexOf(col)]; // 各列の値を連結してキーを作成
    }).join("|");
  }

  /**
   * 閾値をチェックする関数
   * @param {string} typeName - データの種別名
   * @param {Array} data - データ配列
   * @param {Array} headers - ヘッダー行
   * @param {number} row - 行番号
   * @param {number} startCol - データが始まる列
   */
  function checkThresholds(typeName, data, headers, row, startCol) {
    var thresholdData = thresholds[typeName]; // 種別名に対応する閾値データを取得
    for (var col = startCol; col < data[row].length; col++) {
      if (data[row][col] !== "") {
        var header = headers[col]; // 該当するヘッダーを取得
        var thresholdRange = thresholdData[header]; // 閾値範囲を取得
        if (!thresholdRange) {
          continue; // 閾値が定義されていない場合はスキップ
        }
        var value = parseFloat(data[row][col]); // データを数値に変換
        if (isNaN(value) || value < thresholdRange.min || value > thresholdRange.max) {
          setError(row, col, "yellow", 1); // 範囲外の場合はエラー設定
        }
      }
    }
  }

  /**
   * 開示データと加工データのペアをチェックする関数
   * @param {number} baseRow - 基準となる行
   * @param {string} key - チェックする行のキー
   * @param {object} dataMap - データマップ
   */
  function checkPairData(baseRow, key, dataMap) {
    if (startCol !== -1 && data[baseRow][headers.indexOf("種別名")] === "開示データ") {
      for (var col = startCol; col < data[baseRow].length; col++) {
        if (data[baseRow][col] !== "") { // 開示データにデータが収録されている場合
          var matchFound = false; // 加工データが見つかったかどうかのフラグ

          // データ全体をスキャンして、開示データに対応する加工データがあるか確認
          for (var row2 = 6; row2 < data.length; row2++) {
            if (
              data[baseRow][headers.indexOf("出典種別")] === data[row2][headers.indexOf("出典種別")] &&
              data[baseRow][headers.indexOf("コード")] === data[row2][headers.indexOf("コード")] &&
              data[baseRow][headers.indexOf("資料名称")] === data[row2][headers.indexOf("資料名称")] &&
              data[baseRow][headers.indexOf("開示年度")] === data[row2][headers.indexOf("開示年度")] &&
              data[baseRow][headers.indexOf("過年度：年")] === data[row2][headers.indexOf("過年度：年")] &&
              data[baseRow][headers.indexOf("過年度：年月／単位（加工値）")] === data[row2][headers.indexOf("過年度：年月／単位（加工値）")] &&
              (data[row2][headers.indexOf("種別名")] === "加工データ" || data[row2][headers.indexOf("種別名")] === "単位")
            ) {
              if (data[row2][col] === "") { // 加工データに対応するデータが存在しない場合
                sheet.getRange(row2 + 1, col + 1).setBackground("yellow"); // エラー表示
                sheet.getRange(flagRow + 1, col + 1).setValue(1); // フラグ設定
              }
              matchFound = true; // 対応する加工データが見つかった
              break; // 対応データが見つかった時点でループを終了
            }
          }

          // 対応する加工データが見つからなかった場合のエラーログ
          if (!matchFound) {
            console.log("No match found for baseRow: " + baseRow + ", col: " + col);
          }
        }
      }
    }
  }

  // 閾値設定
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

   // メイン処理ループ：データの行ごとにエラーチェックを実施
  for (var row = 6; row < data.length; row++) {
    var typeName = data[row][headers.indexOf("種別名")]; // 種別名を取得

    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];

      // 各項目ごとにエラーチェック
      if (header === "出典種別") {
        checkCondition(row, col, checkSourceType(value, row));
      }

      if (header === "資料公表日") {
        checkCondition(row, col, checkPublicationDate(value, row));
      }

      if (header === "過年度：年") {
        checkCondition(row, col, checkPastYear(value, row));
      }

      // 重複チェックのためのキー生成
      var key = generateKey(row);
      if (!dataMap[key]) {
        dataMap[key] = [];
      }
      dataMap[key].push(typeName);

      // 重複行がある場合はエラーを設定
      if (keyMap[key]) {
        setError(row, 1, "red", 1); // 現在の行
        setError(keyMap[key], 1, "red", 1); // 重複した行
      } else {
        keyMap[key] = row; // 初めての行ならマップに追加
      }
    }

    // 閾値チェック（例：開示データの範囲チェック）
    if (["開示データ", "加工データ"].includes(typeName)) {
      checkThresholds(typeName, data, headers, row, startCol);
    }

    // 資料公表日が日付フォーマットかどうか確認
    if (headers.includes("資料公表日")) {
      checkDateFormat(row, headers.indexOf("資料公表日"));
    }

    // URLのフォーマットをチェック
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

    // ページ数のフォーマットチェック
    if (typeName == "ページ数") {
      if (startCol !== -1) {
        for (var col = startCol; col < data[row].length; col++) {
          if (data[row][col] !== "") {
            var value = data[row][col];
            var isValid = !isNaN(value) || /^[0-9,-]+$/.test(value);
            checkCondition(row, col, isValid);
          }
        }
      }
    }

    // 数値であるべき項目のチェック
    if (typeName == "単位" || typeName == "対象範囲") {
      if (startCol !== -1) {
        for (var col = startCol; col < data[row].length; col++) {
          checkCondition(row, col, data[row][col] !== "" && !isNaN(data[row][col]));
        }
      }
    }

    // 過年度年月が「00」の時に西暦が一致するか確認
    var pastYearCol = headers.indexOf("過年度：年");
    var pastYearMonthCol = headers.indexOf("過年度：年月／単位（加工値）");
    var pastYearValue = data[row][pastYearCol];
    var pastYearMonthValue = data[row][pastYearMonthCol];

    if (pastYearMonthValue.toString().slice(-2) === "00") {
      checkCondition(row, pastYearCol, pastYearValue.toString().slice(0, 4) === pastYearMonthValue.toString().slice(0, 4));
      checkCondition(row, pastYearMonthCol, pastYearValue.toString().slice(0, 4) === pastYearMonthValue.toString().slice(0, 4));
    }
  }

  // 開示データと加工データのペアチェック
  for (var key in dataMap) {
    checkPairData(keyMap[key], key, dataMap);
  }
}