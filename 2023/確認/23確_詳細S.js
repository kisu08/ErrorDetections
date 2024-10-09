function checkDataSdetail2023(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2023(詳細S)");
  var data = sheet.getDataRange().getValues();
  var headers = data[5]; // 6行目が項目名
  var flagRow = 4; // 5行目にフラグを立てる

   // ヘッダーインデックスを先頭で定義
   var headerIndices = {
    codeCol: headers.indexOf("コード"), // コード列のインデックスを追加
    disclosureYearCol: headers.indexOf("開示年度"), // 開示年度列のインデックスを追加
    documentNameCol: headers.indexOf("資料名称"),
    pastYearCol: headers.indexOf("過年度：年"),
    pastYearMonthCol: headers.indexOf("過年度：年月"),
    parentItemCol: headers.indexOf("親項目"),
    parentItemCodeCol: headers.indexOf("親項目コード"), // 親項目コード列のインデックスを追加
    rangeNoCol: headers.indexOf("対象範囲No."),
    rangeCol: headers.indexOf("対象範囲"),
    mergeFlagCol: headers.indexOf("合算フラグ"),
    value1Col: headers.indexOf("数値1"), // 数値1のインデックス
    value2Col: headers.indexOf("数値2"), // 数値2のインデックス
    value3Col: headers.indexOf("数値3"), // 数値3のインデックス
    value4Col: headers.indexOf("数値4"), // 数値4のインデックス
    value5Col: headers.indexOf("数値5"), // 数値5のインデックス
    value6Col: headers.indexOf("数値6"), // 数値6のインデックス
    value7Col: headers.indexOf("数値7"), // 数値7のインデックス
    value8Col: headers.indexOf("数値8"), // 数値8のインデックス
    itemNo1Col: headers.indexOf("項番1"), // 項番1のインデックス
    itemNo2Col: headers.indexOf("項番2"), // 項番2のインデックス
    itemNameCol: headers.indexOf("項目名1"), // 項目名1のインデックス
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

    "過年度：年": function(value, row) {
      // 種別名が資料開示のときNULLであること
      if (data[row][headerIndices.typeNameCol] === "資料開示") {
        return value === "";
      }
      // 4桁の数字であること
      if (!/^[0-9]{4}$/.test(value)) {
        return false;
      }

      // 数字の打ち間違いで異常に大きいまたは小さい年を検出（2000～2024年の範囲内）
      var year = parseInt(value, 10);
      return year >= 2000 && year <= 2024;
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

      // 年月の年部分が2000～2024年の範囲内であること
      var yearMonth = parseInt(valueStr.slice(0, 4), 10);
      if (yearMonth < 2000 || yearMonth > 2024) {
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

    "親項目": function(value){
      //親項目は詳細データ収録対象の項目名であること
      //NULLでないこと
      var basecodeJp = ["【社会】従業員１人あたりの年間平均研修時間","【社会】出産・育児休暇後の復職率"]
      return basecodeJp.includes(value);
    },

    "親項目コード": function(value,row){
      //親項目に紐づく特定の値であること
      //NULLでないこと
      var basecode = data[row][headerIndices.parentItemCol];//親項目

      //親項目と親項目コードの対応関係を定義
      var basecodeMap = {
        "【社会】従業員１人あたりの年間平均研修時間": "S029",
        "【社会】出産・育児休暇後の復職率": "Q_S001"
      };
      return value === basecodeMap[basecode];
    },

    "資料名称": function(value) {
      //「企業HP」「有報」「CG報告書」であること
      //NULLでないこと
      return ["企業HP","有価証券報告書","コーポレートガバナンス報告書"].includes(value);
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
// 対象範囲No.の組み合わせをチェックするためのキー列
      var keyColsCovNo = [headerIndices.codeCol, headerIndices.disclosureYearCol, headerIndices.pastYearCol, headerIndices.pastYearMonthCol, headerIndices.parentItemCodeCol, headerIndices.documentNameCol, headerIndices.rangeNoCol];
      var keyValuesCovNoArray = keyColsCovNo.map(function(colIndex) {
        var value = data[row][colIndex]; // colIndexを使ってデータを取得
        return value;
      });
      var keyValuesCovNo = keyValuesCovNoArray.join("_");
      var rangeNoCol = headerIndices.rangeNoCol; //対象範囲No.
      var rangeCol = headerIndices.rangeCol;     //対象範囲
      if (seenKeys[keyValuesCovNo]) {
        // 最初に見つけた行と比較
        var firstRow = seenKeys[keyValuesCovNo];
        if (data[row][rangeCol] !== data[firstRow][rangeCol]) {
          // 一致していない場合、対象範囲No.と対象範囲のセルを赤色に設定
          sheet.getRange(row + 1, rangeNoCol + 1).setBackground("red");
          sheet.getRange(row + 1, rangeCol + 1).setBackground("red");
          sheet.getRange(flagRow + 1, rangeNoCol + 1).setValue(1);
          sheet.getRange(flagRow + 1, rangeCol + 1).setValue(1);
          sheet.getRange(row + 1, 1).setValue("一致していません");  // A列にエラーメッセージをセット
          sheet.getRange(row + 1, 1).setBackground("tan");
        }
      } else {
        // 新しいキーの組み合わせを記録
        seenKeys[keyValuesCovNo] = row;
      }
      
      // 同じ「コード」「開示年度」「過年度：年」「過年度：年月」「親項目コード」「資料名称」「対象範囲」の組み合わせがあれば、「対象範囲No.」と「対象範囲」の値が一致しているかチェック
      var keyColsCov = [headerIndices.codeCol, headerIndices.disclosureYearCol, headerIndices.pastYearCol, headerIndices.pastYearMonthCol, headerIndices.parentItemCodeCol, headerIndices.documentNameCol, headerIndices.rangeCol];
      // keyValuesCovの値を取得する際もインデックスを使う
      var keyValuesCovArray = keyColsCov.map(function(colIndex) {
        var value = data[row][colIndex]; // colIndexを使ってデータを取得
        return value;
      });
      var keyValuesCov = keyValuesCovArray.join("_");

      if (seenKeysCov[keyValuesCov]) {
        // 最初に見つけた行と比較
        var firstRow = seenKeysCov[keyValuesCov];
        if (data[row][rangeNoCol] !== data[firstRow][rangeNoCol]) {
          // 一致していない場合、対象範囲No.と対象範囲のセルを赤色に設定
          sheet.getRange(row + 1, rangeNoCol + 1).setBackground("red");
          sheet.getRange(row + 1, rangeCol + 1).setBackground("red");
          sheet.getRange(flagRow + 1, rangeNoCol + 1).setValue(1);
          sheet.getRange(flagRow + 1, rangeCol + 1).setValue(1);
          sheet.getRange(row + 1, 1).setValue("一致していません");  // A列にエラーメッセージをセット
          sheet.getRange(row + 1, 1).setBackground("tan");
        } else {
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
  // 数値2がNULLなら数値3を、それもNULLなら数値4を...数値8まで合算して数値1と比較
  for (var row = 6; row < data.length; row++) {
    var sumFlag = data[row][headerIndices.mergeFlagCol];//合算フラグ
    if (sumFlag === 1) {
      var keyColsSumFlag = [headerIndices.codeCol,headerIndices.disclosureYearCol,headerIndices.pastYearCol,headerIndices.pastYearMonthCol,headerIndices.parentItemCodeCol,headerIndices.documentNameCol,headerIndices.rangeNoCol,headerIndices.itemNo1Col,headerIndices.itemNameCol];
      var keyValuesSumFlag = keyColsSumFlag.map(function(colIndex) {
        return data[row][colIndex];
      }).join("_");

      // 合算のための変数を初期化
    var uniqueValues = new Set();
    var sumValue = 0; // ここで sumValue を初期化

     // 同じキー項目を持つ行を探して数値2を合算
      for (var i = 6; i < data.length; i++) {
        var currentKey = keyColsSumFlag.map(function(colIndex) {
          return data[i][colIndex];
        }).join("_");

        if (currentKey === keyValuesSumFlag) {
          var value2 = data[i][headerIndices.value2Col];//数値2
          var value2 = data[i][headerIndices.value2Col] || 0;
          var value3 = data[i][headerIndices.value3Col] || 0;
          var value4 = data[i][headerIndices.value4Col] || 0;
          var value5 = data[i][headerIndices.value5Col] || 0;
          var value6 = data[i][headerIndices.value6Col] || 0;
          var value7 = data[i][headerIndices.value7Col] || 0;
          var value8 = data[i][headerIndices.value8Col] || 0;
          var itemNo2 = data[i][headerIndices.itemNo2Col];//項番2

          // 数値2の値がNULLである場合は0として扱う
          var actualValue = parseFloat(value2) || parseFloat(value3) || parseFloat(value4) ||
                          parseFloat(value5) || parseFloat(value6) || parseFloat(value7) ||
                          parseFloat(value8) || 0;
          if (value2 == null || value2 === ""){
            value2 = 0;
          }

          if (!uniqueValues.has(itemNo2)) {
            uniqueValues.add(itemNo2);
            sumValue += actualValue; // 数値2〜数値8のいずれかの値を合算
          }
        }
      }

      var value1 = parseFloat(data[row][headerIndices.value1Col]); // 数値1を明示的に数値に変換

      // 合算した値と数値1を比較
      if (sumValue.toFixed(2) !== value1.toFixed(2)) {
        var cell = sheet.getRange(row + 1, headerIndices.value1Col + 1);
        cell.setBackground("red");
        sheet.getRange(flagRow + 1, headerIndices.value1Col + 1).setValue(1);
        sheet.getRange(row + 1, 1).setValue("合計値にミスがあります");  // A列にエラーメッセージをセット
        sheet.getRange(row + 1, 1).setBackground("orange");
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
        sheet.getRange(row + 1, 1).setValue("一致していません");  // A列にエラーメッセージをセット
        sheet.getRange(row + 1, 1).setBackground("tan");
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
        sheet.getRange(row + 1, 1).setValue("複数存在します");  // A列にエラーメッセージをセット
        sheet.getRange(row + 1, 1).setBackground("orange");
      });
    }
  }
  // エラー行のカウントを実行
  countColoredCells(sheet);
  SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');
};