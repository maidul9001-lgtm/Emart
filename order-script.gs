var KEY = 'CHANGE-ME'; // এখানে নিজের গোপন কী বসান (এডমিন প্যানেলে এটাই দিতে হবে)

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Orders') || ss.insertSheet('Orders');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['অর্ডার নং', 'তারিখ', 'নাম', 'মোবাইল', 'ঠিকানা', 'এলাকা', 'নোট', 'পণ্য', 'পণ্যের দাম', 'ডেলিভারি', 'মোট', 'স্ট্যাটাস']);
  }
  return sh;
}

// এডমিন প্যানেল অর্ডার পড়ার জন্য
function doGet(e) {
  if (!e.parameter || e.parameter.key !== KEY) return out({ ok: false });
  var rows = getSheet().getDataRange().getValues();
  var orders = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    orders.push({ no: r[0], date: r[1], name: r[2], phone: String(r[3]), addr: r[4], area: r[5], note: r[6], items: r[7], sub: r[8], del: r[9], tot: r[10], status: r[11] });
  }
  return out({ ok: true, orders: orders });
}

// নতুন অর্ডার জমা + স্ট্যাটাস আপডেট
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var sh = getSheet();
  if (d.action === 'status') {
    if (d.key !== KEY) return out({ ok: false });
    var rows = sh.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(d.no)) { sh.getRange(i + 1, 12).setValue(d.status); break; }
    }
    return out({ ok: true });
  }
  sh.appendRow([d.no, d.date, d.name, "'" + d.phone, d.addr, d.area, d.note, d.items, d.sub, d.del, d.tot, 'পেন্ডিং']);
  MailApp.sendEmail(
    Session.getEffectiveUser().getEmail(),
    'নতুন অর্ডার #' + d.no + ' - ' + d.tot + '৳',
    'নাম: ' + d.name + '\nমোবাইল: ' + d.phone + '\nঠিকানা: ' + d.addr +
    '\nপণ্য: ' + d.items + '\nমোট: ' + d.tot + '৳\nনোট: ' + d.note
  );
  return out({ ok: true });
}
