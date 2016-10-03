// Sidebar.htmlをテンプレートにHTMLServiceを生成
// Sidebar.htmlが必要
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle("シート集計");
  SpreadsheetApp.getUi().showSidebar(ui);
}

// スプレッドシートを開いたときに呼ばれる
function onOpen() {
 showSidebar();
}

// 集計対象フォルダ以下のファイルを集計
function aggregateData() {
  var cfg = getConfig();
  var targetFolder = DriveApp.getFoldersByName(cfg['集計対象フォルダ']).next();
  var files = targetFolder.getFiles();
  Logger.log(files);

  var outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg['書き込み先シート']);
  if (! outputSheet) {
    var sheetName = String(cfg['書き込み先シート']);
    Logger.log(sheetName);
    outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }
  outputSheet.clear();

  var curRow = 1;
  var fileList = [];
  while(files.hasNext()){
    var file = files.next();

    var spreadsheet = SpreadsheetApp.open(file);
    var sheet = spreadsheet.getSheets()[0];
    var id = sheet.getRange(cfg['識別子']).getValue();
    fileList.push( { "fileName": file.getName(), "id": id } );

    var srcRange = sheet.getRange(cfg['読み込み領域']);
    outputSheet.getRange(curRow, 2, srcRange.getHeight(), srcRange.getWidth())
               .setValues(srcRange.getValues());

    var lastRow = outputSheet.getLastRow()+1;
    outputSheet.getRange(curRow, 1, lastRow-curRow, 1)
               .setValue(id);      // 識別子の書き込み
    curRow = lastRow;
  }

  return fileList;
}

// 集計設定という名前のシートを読み込んで key-valueに落とし込む
function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = ss.getSheetByName("集計設定");
  var cfgValues = cfgSheet.getDataRange().getValues();
  var cfg = {};

  for (var i=0; i<cfgValues.length; i++) {
    cfg[cfgValues[i][0]] = String(cfgValues[i][1]);
  }
  return cfg;
}
