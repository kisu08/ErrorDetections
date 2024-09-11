function checkDataEdetail2023(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2023(詳細E)");
  var data = sheet.getDataRange().getValues();
  var headers = data[5]; // 6行目が項目名
  var flagRow = 4; // 5行目にフラグを立てる

   // ヘッダーインデックスを先頭で定義
   var headerIndices = {
    documentNameCol: headers.indexOf("資料名称"),
    pastYearCol: headers.indexOf("過年度：年"),
    pastYearMonthCol: headers.indexOf("過年度：年月"),
    parentItemCol: headers.indexOf("親項目"),
    rangeNoCol: headers.indexOf("対象範囲No."),
    rangeCol: headers.indexOf("対象範囲"),
    mergeFlagCol: headers.indexOf("合算フラグ"),
    value1Col: headers.indexOf("数値1"),
    value2Col: headers.indexOf("数値2"),
    itemNo1Col: headers.indexOf("項番1"),
    itemNo2Col: headers.indexOf("項番2"),
    itemNameCol: headers.indexOf("項目名1")
  };

  // エラー検知してイエローに変更する関数
  function setErrorHighlight(sheet, row, col, flagRow) {
    sheet.getRange(row + 1, col + 1).setBackground("yellow");
    sheet.getRange(flagRow + 1, col + 1).setValue(1);
  }

  // 項目特有のエラー検知条件を設定する
  var conditions = {  
    
    "コード": function(value) {
      //4桁の数字もしくは英数字であること
      //NULLでないこと
      return /^[0-9]{4}$/.test(value) || /^[A-Za-z0-9]{4}$/.test(value);
    },

    "開示年度": function(value) {
      //値が「2023」あること
      //NULLでないこと
      return value === 2023;
    },

    "過年度：年": function(value) {
      //4桁の数字であること
      //NULLでないこと
      return /^[0-9]{4}$/.test(value);
    },
    
    "過年度：年月": function(value) {
      //6桁の数字で末尾が「00」～「12」であること
      //NULLでないこと
      var valueStr = value.toString(); // 数値を文字列に変換
      return /^[0-9]{6}$/.test(valueStr) && /^(00|01|02|03|04|05|06|07|08|09|10|11|12)$/.test(valueStr.slice(-2));
    },

    "親項目": function(value){
      //親項目は詳細データ収録対象の項目名であること
      //NULLでないこと
      var basecodeJp = ["【環境】温室効果ガス（GHG）排出量（Scope1）","【環境】温室効果ガス（GHG）排出量（Scope2）","【環境】温室効果ガス（GHG）総排出量","【環境】水総使用量（消費量）","【環境】総排水量","【環境】総エネルギー消費量","【環境】有害廃棄物量","【環境】リサイクル廃棄物量","【環境】非リサイクル廃棄物量"]
      return basecodeJp.includes(value);
    },

    "親項目コード": function(value,row){
      //親項目に紐づく特定の値であること
      //NULLでないこと
      var basecode = data[row][headerIndices.parentItemCol];//親項目

      //親項目と親項目コードの対応関係を定義
      var basecodeMap = {
       "【環境】温室効果ガス（GHG）排出量（Scope1）":"E001",
        "【環境】温室効果ガス（GHG）排出量（Scope2）":"E002",
        "【環境】温室効果ガス（GHG）総排出量":"E004",
        "【環境】水総使用量（消費量）":"E024",
        "【環境】総排水量":"E026",
        "【環境】総エネルギー消費量":"E029",
        "【環境】有害廃棄物量":"Q_E001",
        "【環境】リサイクル廃棄物量":"Q_E002",
        "【環境】非リサイクル廃棄物量":"Q_E003"
      };
      return value === basecodeMap[basecode];
    },

    "資料名称": function(value) {
      //「企業HP」であること ※23年度よりEデータは企業HPのみ対象
      //NULLでないこと
      return ["企業HP"].includes(value);
    },

    "URL": function(value, row){
      //有報はEDINET、CG報告書はコーポレート・ガバナンス情報サービスから参照していること
      var disclosureName = data[row][headerIndices.documentNameCol];//資料名称
      if (disclosureName === "有価証券報告書"){
        return value.includes("https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf");
      }else if (disclosureName === "コーポレートガバナンス報告書"){
        return value.includes("https://www2.jpx.co.jp/disc/");
      }
      //「https:～」または「http:～」の文字列を含むこと
      //NULLでないこと
      return (value.includes("https:") || value.includes("http:"))
    } ,

    "ページ数": function(value){
      //数値データになっていること。「-」「,」との組み合わせであれば許容。「-」「,」単体は不可。
      if (value !== ""){
        if ((value === "," || value === "-")){
          return false
        }
        return /^[0-9,.-]+$/.test(value)
      }
      return true
    },
    "対象範囲No." : function(value){
      //数値データであること
      //NULLでないこと
      if ((value === "" || isNaN(value))){
        return false
      }
      else{
        //自然数であること
        return Number.isInteger(value) && value >= 1;
      }
    },

    "対象範囲" : function(value){
      //文字列であること
      return (isNaN(value)|| value === "")
    },

    "合算フラグ":function(value){
      //0または1であること
      //NULLでないこと
      return (value === 0 || value === 1)
    },

    "項目名1" : function(value){
      //項目名1（共通項目）が特定の値であること
      var commonItems = ["国・地域別","機能別","事業別","バウンダリ別","個人別","男女別","雇用管理区分別","役職別","算定基準別","施設別","種類別","合計","その他"]
      return commonItems.includes(value)
    }
  };

 // エラー検知とフラグ設定
 var seenKeys= {};  //「対象範囲No.」のキー項目を保持するオブジェクト
 var seenKeysCov = {}; //「対象範囲」のキー項目を保持するオブジェクト
   for (var row = 6; row < data.length; row++) {
    for (var col = 0; col < headers.length; col++) {
      var header = headers[col];
      var value = data[row][col];
      
      //conditionsで定義した項目特有のエラーを検知したかをチェック
      if (conditions[header] && !conditions[header](value, row)) {
        setErrorHighlight(sheet, row, col, flagRow);
      }
      
      //「対象範囲No.」と「対象範囲」の組み合わせが適切かをチェック
      // 同じ「コード」「開示年度」「過年度：年」「過年度：年月」「親項目コード」「資料名称」「対象範囲No.」の組み合わせがあれば、「対象範囲No.」と「対象範囲」の値が一致しているかチェック
      var keyColsCovNo = ["コード", "開示年度", "過年度：年", "過年度：年月", "親項目コード", "資料名称", "対象範囲No."];
      var keyValuesCovNo = keyColsCovNo.map(function(colName) {
        return data[row][headerIndices[colName]];
      }).join("_");
      
      var rangeNoCol = headerIndices.rangeNoCol;//対象範囲No.
      var rangeCol = headerIndices.rangeCol;//対象範囲
      if (seenKeys[keyValuesCovNo]) {
        // 最初に見つけた行と比較
        var firstRow = seenKeys[keyValuesCovNo];
        if (data[row][rangeCol] !== data[firstRow][rangeCol]) {
          // 一致していない場合、対象範囲No.と対象範囲のセルを赤色に設定
          sheet.getRange(row + 1, rangeNoCol + 1).setBackground("red");
          sheet.getRange(row + 1, rangeCol + 1).setBackground("red");
          sheet.getRange(flagRow + 1, rangeNoCol + 1).setValue(1);
          sheet.getRange(flagRow + 1, rangeCol + 1).setValue(1);
        }
      } else {
        // 新しいキーの組み合わせを記録
        seenKeys[keyValuesCovNo] = row;
      }
      
      // 同じ「コード」「開示年度」「過年度：年」「過年度：年月」「親項目コード」「資料名称」「対象範囲」の組み合わせがあれば、「対象範囲No.」と「対象範囲」の値が一致しているかチェック
      var keyColsCov = ["コード", "開示年度", "過年度：年", "過年度：年月", "親項目コード", "資料名称", "対象範囲"];
      var keyValuesCov = keyColsCov.map(function(colName) {
        return data[row][headerIndices[colName]];
      }).join("_");

      if (seenKeysCov[keyValuesCov]) {
        // 最初に見つけた行と比較
        var firstRow = seenKeysCov[keyValuesCov];
        if (data[row][rangeNoCol] !== data[firstRow][rangeNoCol]) {
          // 一致していない場合、対象範囲No.と対象範囲のセルを赤色に設定
          sheet.getRange(row + 1, rangeNoCol + 1).setBackground("red");
          sheet.getRange(row + 1, rangeCol + 1).setBackground("red");
          sheet.getRange(flagRow + 1, rangeNoCol + 1).setValue(1);
          sheet.getRange(flagRow + 1, rangeCol + 1).setValue(1);
        }
      } else {
        // 新しいキーの組み合わせを記録
        seenKeysCov[keyValuesCov] = row;
      }

      //ヘッダーが「項番」で始まる場合、それの列の値が自然数であることをチェック
      if (header.startsWith("項番") && value !== "" && !(Number.isInteger(value) && value >=1)){
        setErrorHighlight(sheet, row, col, flagRow);
      }

      //ヘッダーが「数値」で始まる場合、その列の値が数値であることをチェック
      if (header.startsWith("数値") && isNaN(value)){
        setErrorHighlight(sheet, row, col, flagRow);
      }

      //ヘッダーが「項目名」「単位」で始まる場合、その列の値が文字列であることをチェック
      if ((header.startsWith("項目名") || header.startsWith("単位")) && value!=="" &&!isNaN(value)){
        setErrorHighlight(sheet, row, col, flagRow);
      }

      //項番と項目名のNULLチェック
      for (var i = 1; i <= 8; i++){
        var itemHeader = "項目名" + i;
        var numHeader = "項番" + i;
        var itemCol = headers.indexOf(itemHeader);
        var numCol = headers.indexOf(numHeader);

        if (itemCol !== -1 && numCol !== -1){
          var itemValue = data[row][itemCol];
          var numValue = data[row][numCol];
          
          //項目名に値が収録されているにも関わらず、項番がNULLであればエラー検知
          if (header === numHeader && numValue === "" && itemValue !== ""){
            setErrorHighlight(sheet, row, col, flagRow);
          }

          //項番に値が収録されているにも関わらず、項目名がNULLであればエラー検知
          if (header === itemHeader && itemValue === "" && numValue !== ""){
            setErrorHighlight(sheet, row, col, flagRow);
          }
        }
      }
    }
  };

  //合算フラグ1の場合に、正しく合計値が算出されているかをチェック
  // 「数値2」の合算と「数値1」の比較
  for (var row = 6; row < data.length; row++) {
    var sumFlag = data[row][headerIndices.mergeFlagCol];//合算フラグ
    if (sumFlag === 1) {
      var keyColsSumFlag = ["コード", "開示年度", "過年度：年", "過年度：年月", "親項目コード", "資料名称", "対象範囲No.", "項番1", "項目名1"];
      var keyValuesSumFlag = keyColsSumFlag.map(function(colName) {
        return data[row][headerIndices[colName]];
      }).join("_");

      // 同じキー項目を持つ行を集めて「数値2」を合算
      var uniqueValues = new Set();
      var sumValue2 = 0;
      for (var i = 6; i < data.length; i++) {
        var currentKey = keyColsSumFlag.map(function(colName) {
          return data[i][headerIndices[colName]];
        }).join("_");

        if (currentKey === keyValuesSumFlag) {
          var value2 = data[i][headerIndices.value2Col];//数値2
          var itemNo2 = data[i][headerIndices.itemNo2Col];//項番2

          // 数値2の値がNULLである場合は0として扱う
          if (value2 == null || value2 === ""){
            value2 = 0;
          }

          if (!uniqueValues.has(itemNo2)) {
            uniqueValues.add(itemNo2);
            sumValue2 += value2;
          }
        }
      }

      var value1 = data[row][headerIndices.value1Col];//数値1

      // 「数値2」の合算値と「数値1」が一致するかをチェック
      if (sumValue2 !== value1) {
        var cell = sheet.getRange(row + 1, headerIndices.value1Col + 1);
        cell.setBackground("red");
        sheet.getRange(flagRow + 1, headerIndices.value1Col + 1).setValue(1);
      }
    }
  }

  //過年度年月が「00」の時に、過年度年と過年度年月の西暦が一致していること
  for (var row = 6; row < data.length; row++) {
    var pastYearCol = headerIndices.pastYearCol;//過年度：年
    var pastYearMonthCol = headerIndices.pastYearMonthCol;//過年度：年月
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

  //同じレコードがないこと（重複していないこと）
  var uniqueRows = {};
  for (var row = 6; row < data.length; row++) {
    var rowData = data[row].slice(2).join(",");
    if (uniqueRows[rowData]) {
      uniqueRows[rowData].push(row);
    } else {
      uniqueRows[rowData] = [row];
    }
  }

  for (var key in uniqueRows) {
    if (uniqueRows[key].length > 1) {
      // 一致する行が複数ある場合の処理
      sheet.getRange(flagRow + 1, 2).setBackground("red").setValue(1);
      uniqueRows[key].forEach(function(row) {
        sheet.getRange(row + 1, 2).setBackground("red");
      });
    }
  }

  SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');
};