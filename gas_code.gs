var SHEET_ID = '10EUU_Cb8yy1-CSCVw2jOeSl_OCR-l4K0MeMQxcGEub8';

function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var raw = ss.getSheetByName('raw_orders');
    if (!raw) return ContentService.createTextOutput('{"orders":[]}').setMimeType(ContentService.MimeType.JSON);
    var json = raw.getRange(1, 1).getValue();
    var orders = json ? JSON.parse(json) : [];
    return ContentService.createTextOutput(JSON.stringify({orders: orders})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput('{"orders":[]}').setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SHEET_ID);

    if (data.action === 'sync_orders') {
      var rows = data.data.rows || [];

      // 顧客管理シートに全件書き込み
      var s1 = ss.getSheetByName('顧客管理');
      if (!s1) s1 = ss.insertSheet('顧客管理');
      s1.clearContents();
      s1.appendRow(['日付', '顧客名', '修理品', '金額', '完了']);
      s1.getRange(1, 1, 1, 5).setFontWeight('bold');
      if (rows.length > 0) s1.getRange(2, 1, rows.length, 5).setValues(rows);

      // 全受付データをJSONで保存（データ取得用）
      var fullOrders = data.data.orders || [];
      if (fullOrders.length > 0) {
        var raw = ss.getSheetByName('raw_orders');
        if (!raw) raw = ss.insertSheet('raw_orders');
        raw.clearContents();
        raw.getRange(1, 1).setValue(JSON.stringify(fullOrders));
      }

      // 月次集計をGAS側で直接計算して書き込み
      setupMonthlySheet(ss, rows);
      return ok();
    }

    return ok();
  } catch(err) {
    return ContentService.createTextOutput('{"status":"error"}').setMimeType(ContentService.MimeType.JSON);
  }
}

function setupMonthlySheet(ss, rows) {
  var ms = ss.getSheetByName('月次集計');
  if (!ms) ms = ss.insertSheet('月次集計');
  ms.clearContents();

  ms.getRange(1, 1, 1, 4).setValues([['年月', '月間件数', '月間売上', '月間完了']]);
  ms.getRange(1, 1, 1, 4).setFontWeight('bold');

  var now = new Date();
  var results = [];
  for (var i = 11; i >= 0; i--) {
    var d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    var ym = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
    var count = 0, sales = 0, done = 0;
    for (var j = 0; j < rows.length; j++) {
      var row = rows[j];
      if (String(row[0]).slice(0, 7) === ym) {
        count++;
        if (row[4] === true) {
          sales += row[3] || 0;
          done++;
        }
      }
    }
    results.push([ym, count, sales, done]);
  }
  if (results.length > 0) ms.getRange(2, 1, results.length, 4).setValues(results);
}

function ok() {
  return ContentService.createTextOutput('{"status":"ok"}').setMimeType(ContentService.MimeType.JSON);
}
