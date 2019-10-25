function myFunction(){
  //明細元情報
  var payslipFile = SpreadsheetApp.openById('hogehoge');
  var payslipSheet = payslipFile.getSheetByName('シート名');
  var payslipListCount = payslipSheet.getLastRow();  
  
  // PDFの保存先となるフォルダID 確認方法は後述
  var folderid = "フォルダID";
  
  // 現在開いているスプレッドシートを取得
  var pdfSheet = SpreadsheetApp.getActiveSpreadsheet();
  // スプレッドシートのIDを取得
  var pdfSheetId = pdfSheet.getId();
  // スプレッドシートのシートIDを取得
  var sheetid = pdfSheet.getActiveSheet().getSheetId();
  // getActiveSheetの後の()を忘れると、TypeError: オブジェクト function getActiveSheet() {/* */} で関数 getSheetId が見つかりません。

  // ファイル名に使用する名前を取得
  var targetName = pdfSheet.getRange("O3").getValue();
  
  // ファイル名に使用する名前を取得
  var fileName = pdfSheet.getRange("B1").getValue();
  
  //一覧シートの初期化
  pdfSheet.getRange("B6:O22").clear();
  
  //明細データ取得(low,col,low,col)→二次元配列
  var payslipData = payslipSheet.getRange(1, 1, payslipListCount, 4).getValues();
    
  //↓PDFの中身を出力する 
  //出力先の項目場所情報
    var pdfCount = 6;
    var dateVal = "";
    var costNameVal = "";
    var nameVal = "";
    
    for(var j=1; j<payslipListCount; j++){
      var payslipTarget = payslipData[j][3];
      
        if(targetName == payslipTarget){
          //日付設定
          dateVal = payslipData[j][0];
          pdfSheet.getRange("B" + pdfCount).setValue(dateVal);
          //支払い名称設定
          costNameVal = payslipData[j][1];
          pdfSheet.getRange("F" + pdfCount).setValue(costNameVal);
          //金額設定
          nameVal = payslipData[j][2];
          pdfSheet.getRange("O" + pdfCount).setValue(nameVal);
          pdfCount = pdfCount + 1;
        
        }
    } 
    //セルのレイアウト作成
    //セルの結合
    for(var i=6; i <= 22; i++){
        var dateMarge = pdfSheet.getRange("B"+i+":E"+i).merge();
        var payslipMarge = pdfSheet.getRange("F"+i+":N"+i).merge();
        var costMarge = pdfSheet.getRange("O"+i+":R"+i).merge();
    }
    //セルの範囲指定と罫線作成
    var line = pdfSheet.getRange("B6"+":R22");
    line.setBorder(true, true, true, true, true, true);
    //名前設定
    //pdfSheet.getRange("O" + 3).setValue(targetName);
    
    // PDF作成関数
    //createPDF( folderid, pdfSheetId, sheetid, fileName + "_" + targetName );
}


function makePDF(){
  // PDFの保存先となるフォルダID 確認方法は後述
  var folderid = "フォルダID";
  // 現在開いているスプレッドシートを取得
  var pdfSheet = SpreadsheetApp.getActiveSpreadsheet();
  // スプレッドシートのIDを取得
  var pdfSheetId = pdfSheet.getId();
  // スプレッドシートのシートIDを取得
  var sheetid = pdfSheet.getActiveSheet().getSheetId();
  // getActiveSheetの後の()を忘れると、TypeError: オブジェクト function getActiveSheet() {/* */} で関数 getSheetId が見つかりません。

  // ファイル名に使用する名前を取得
  var targetName = pdfSheet.getRange("O3").getValue();
  
  // ファイル名に使用する名前を取得
  var fileName = pdfSheet.getRange("B1").getValue();
  var nowMonth = Utilities.formatDate(new Date(), "JST", "YYYY/MM");
  
  createPDF( folderid, pdfSheetId, sheetid, nowMonth + "_" + targetName + "_" + fileName );
}

// PDF作成関数 引数は（folderid:保存先フォルダID, ssid:PDF化するスプレッドシートID, sheetid:PDF化するシートID, filename:PDFの名前）
function createPDF(folderid, ssid, sheetid, filename){

  // PDFファイルの保存先となるフォルダをフォルダIDで指定
  var folder = DriveApp.getFolderById(folderid);

  // スプレッドシートをPDFにエクスポートするためのURL。このURLに色々なオプションを付けてPDFを作成
  var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);

  // PDF作成のオプションを指定
  var opts = {
    exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         "A4",     // 用紙サイズの指定 legal / letter / A4
    portrait:     "true",   // true → 縦向き、false → 横向き
    fitw:         "true",   // 幅を用紙に合わせるか
    sheetnames:   "false",  // シート名をPDF上部に表示するか
    printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
    pagenumbers:  "false",  // ページ番号の有無
    gridlines:    "false",  // グリッドラインの表示有無
    fzr:          "false",  // 固定行の表示有無
    validateHttpsCertificates : "false",
    gid:          sheetid   // シートIDを指定 sheetidは引数で取得
  };
  
  var url_ext = [];
  
  // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
  for( optName in opts ){
    url_ext.push( optName + "=" + opts[optName] );
  }

  // url_extの各要素を「&」で繋げる
  var options = url_ext.join("&");

  // optionsは以下のように作成しても同じです。
  // var ptions = 'exportFormat=pdf&format=pdf'
  // + '&size=A4'                       
  // + '&portrait=true'                    
  // + '&sheetnames=false&printtitle=false' 
  // + '&pagenumbers=false&gridlines=false' 
  // + '&fzr=false'                         
  // + '&gid=' + sheetid;

  // API使用のためのOAuth認証
  var token = ScriptApp.getOAuthToken();

    // PDF作成
    var response = UrlFetchApp.fetch(url + options, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });

    // 
    var blob = response.getBlob().setName(filename + '.pdf');

  //}

  // PDFを指定したフォルダに保存
  folder.createFile(blob);

}

// スプレッドシートのメニューからPDF作成用の関数を実行出来るように、「スクリプト」というメニューを追加。
function onOpen() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [       
        {
            name : "明細出力",
            functionName : "myFunction"
        } ,   
        {
            name : "Create PDF",
            functionName : "makePDF"
        }
        ];
    sheet.addMenu("スクリプト", entries);
};
