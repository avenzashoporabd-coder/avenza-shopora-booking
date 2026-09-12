const SHEET_NAME = 'Bookings';
const INITIAL_STOCK = 23;

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ok:true,service:'Avenza Free Booking API'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    if (!d.name || !d.phone || !d.address || !d.receiveAddress || !d.quantity || d.quantity < 1)
      throw new Error('Required fields missing');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(['Booking ID','Time','Name','Phone','Address','Receive Address','Emergency','Selected Products','Quantity','Note','Remaining Stock']);
      sh.getRange('K2').setValue(INITIAL_STOCK);
    }

    const stockCell = sh.getRange('K2');
    let stock = Number(stockCell.getValue() || INITIAL_STOCK);
    if (Number(d.quantity) > stock) throw new Error('Only ' + stock + ' pcs are available');

    stock -= Number(d.quantity);
    stockCell.setValue(stock);

    const id = 'AVZ-BK-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    const items = (d.items || []).map(x => x.name + ' × ' + x.qty).join(', ');
    sh.appendRow([id,new Date(),d.name,d.phone,d.address,d.receiveAddress,d.emergency || '',items,d.quantity,d.note || '',stock]);

    return json({ok:true,bookingId:id,remainingStock:stock});
  } catch (err) {
    return json({ok:false,error:String(err.message || err)});
  } finally {
    lock.releaseLock();
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
