// 報名表後端（Google Apps Script 網頁應用程式）。前端 form-demo.html 放在 GitHub Pages，用 fetch 送 JSON 進來。
var SHEET = '報名資料';
var FOLDER = '衛教報名_上傳檔案';
var CODE = 'TYH2026';           // 報名代碼：只印在海報與課堂上，防止網址外流後亂送
var MAX = 5 * 1024 * 1024;      // 附件上限 5 MB
var TYPES = ['image/jpeg', 'image/png', 'application/pdf'];

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.code !== CODE) return out({ ok: false, error: '報名代碼錯誤' });
    if (!d.name || !/^09\d{8}$/.test(d.phone || '')) return out({ ok: false, error: '姓名或手機格式不對' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
    if (sh.getLastRow() === 0) sh.appendRow(['送出時間', '姓名', '手機', 'Email', '場次', '家屬人數', '檔案連結']);
    var link = '';
    if (d.file && d.file.data) {
      var bytes = Utilities.base64Decode(d.file.data);
      if (bytes.length > MAX) return out({ ok: false, error: '檔案超過 5 MB' });
      if (TYPES.indexOf(d.file.type) < 0) return out({ ok: false, error: '只收 jpg、png、pdf' });
      var it = DriveApp.getFoldersByName(FOLDER);
      var folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
      var stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd');
      link = folder.createFile(Utilities.newBlob(bytes, d.file.type, d.name + '_' + stamp + '_' + d.file.name)).getUrl();
    }
    sh.appendRow([new Date(), d.name, d.phone, d.email || '', d.session || '', d.family || 0, link]);
    sh.getRange(sh.getLastRow(), 3).setNumberFormat('@').setValue(d.phone); // 手機保留開頭的 0
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet() { return out({ ok: true, msg: '報名表後端運作中' }); }

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
