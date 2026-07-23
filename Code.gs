/** HemoStock GAS backend — paste this file into the Apps Script project bound to the supplied spreadsheet. */
const SHEET_ID = '1Y7MnJLzmzeD10QMmVXnwSxBtXbE6lEmz8DTFsChYp2A';
const HEADERS = ['ItemCode','ชื่อรายการ','ประเภท','หน่วยนับ','วันที่รับเข้า','วันหมดอายุ','คงเหลือ','ผู้รับเข้า','หมายเหตุ','UpdatedAt'];
const CATEGORY_EN = {'น้ำยาฟอกไต':'Dialysis','น้ำยา Disinfectant':'Disinfect','ตัวกรองเลือด':'Filter','สายส่งเลือด':'Line','Needle':'Needle'};

function doGet(e) {
  const action = (e.parameter.action || 'inventory').toLowerCase();
  if (action === 'inventory') return json_({ ok:true, items: getInventory_() });
  if (action === 'catalog') return json_({ ok:true, items: getCatalog_() });
  return json_({ ok:false, message:'Unknown action' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'stockIn')    return json_(stockIn_(body.item, body.note || ''));
    if (body.action === 'stockOut')   return json_(stockOut_(body));
    if (body.action === 'deleteItem') return json_(deleteItem_(body.code));
    if (body.action === 'editItem')   return json_(editItem_(body.item));
    return json_({ ok:false, message:'Unknown action' });
  } catch (err) { return json_({ ok:false, message:String(err) }); }
}

function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureSheet_(ss, 'Inventory', HEADERS);
  ensureSheet_(ss, 'Transactions', ['TransactionId','วันที่','ประเภท','ItemCode','ชื่อรายการ','จำนวน','หน่วยนับ','การตัดจ่าย','หมายเหตุ','ผู้บันทึก']);
}

function getCatalog_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Data');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,4).getValues().filter(r=>r[0]).map(r=>({name:r[0],category:r[1],unit:r[2],note:r[3]}));
}
function getInventory_() {
  const sh = sheet_('Inventory', HEADERS), values=sh.getDataRange().getValues();
  return values.slice(1).filter(r=>r[0]).map(r=>({code:r[0],name:r[1],category:r[2],unit:r[3],received:date_(r[4]),expiry:date_(r[5]),qty:Number(r[6])||0,receivedBy:r[7]||'',note:r[8]||''}));
}
function stockIn_(item,note) {
  if (!item || !item.name || !item.qty) throw new Error('กรอกข้อมูลรับเข้าไม่ครบ');
  const sh=sheet_('Inventory',HEADERS), rows=sh.getDataRange().getValues();
  const ymd=String(item.received||'').replace(/-/g,'') || Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd');
  const prefix=CATEGORY_EN[item.category] || 'Item';
  const code=item.code || nextCode_(rows,prefix,ymd);
  const index=rows.findIndex((r,n)=>n>0 && r[0]===code);
  if (index > 0) {
    const row=index+1, old=Number(rows[index][6])||0;
    sh.getRange(row,5,1,6).setValues([[item.received,item.expiry,old+Number(item.qty),item.receivedBy||'',note,new Date()]]);
  } else sh.appendRow([code,item.name,item.category,item.unit,item.received,item.expiry,Number(item.qty),item.receivedBy||'',note,new Date()]);
  transaction_('รับเข้า',item,Number(item.qty),'',note,undefined,item.receivedBy);
  return {ok:true,code:code};
}
function stockOut_(body) {
  const sh=sheet_('Inventory',HEADERS), rows=sh.getDataRange().getValues();
  const index=rows.findIndex((r,n)=>n>0 && r[0]===body.code);
  if(index<1) throw new Error('ไม่พบ ItemCode');
  const item=rows[index], qty=Number(body.quantity);
  if(!qty || qty<1 || qty>Number(item[6])) throw new Error('จำนวนจ่ายไม่ถูกต้องหรือคงเหลือไม่พอ');
  sh.getRange(index+1,7).setValue(Number(item[6])-qty);
  sh.getRange(index+1,10).setValue(new Date());
  transaction_('จ่ายออก',{code:item[0],name:item[1],unit:item[3]},qty,body.type||'',body.note||'',body.date,body.staff);
  return {ok:true,balance:Number(item[6])-qty};
}
function deleteItem_(code) {
  if (!code) throw new Error('ไม่ระบุ ItemCode');
  const sh = sheet_('Inventory', HEADERS), rows = sh.getDataRange().getValues();
  const index = rows.findIndex((r, n) => n > 0 && r[0] === code);
  if (index < 1) throw new Error('ไม่พบ ItemCode: ' + code);
  sh.deleteRow(index + 1);
  return { ok: true, code: code };
}
function editItem_(item) {
  if (!item || !item.code) throw new Error('ข้อมูลไม่ครบ');
  const sh = sheet_('Inventory', HEADERS), rows = sh.getDataRange().getValues();
  const index = rows.findIndex((r, n) => n > 0 && r[0] === item.code);
  if (index < 1) throw new Error('ไม่พบ ItemCode: ' + item.code);
  const row = index + 1;
  sh.getRange(row, 1, 1, 10).setValues([[item.code, item.name, item.category, item.unit, item.received, item.expiry, Number(item.qty) || 0, item.receivedBy || '', item.note || '', new Date()]]);
  return { ok: true, code: item.code };
}
function transaction_(kind,item,qty,outType,note,date,staff) {
  sheet_('Transactions',['TransactionId','วันที่','ประเภท','ItemCode','ชื่อรายการ','จำนวน','หน่วยนับ','การตัดจ่าย','หมายเหตุ','ผู้บันทึก'])
    .appendRow(['TX-'+Utilities.getUuid().slice(0,8).toUpperCase(),date||new Date(),kind,item.code,item.name,qty,item.unit,outType,note,staff||Session.getActiveUser().getEmail()]);
}
function sheet_(name,headers){return ensureSheet_(SpreadsheetApp.openById(SHEET_ID),name,headers)}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh){sh=ss.insertSheet(name);sh.appendRow(headers);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#d9eef3')}return sh}
function nextCode_(rows,prefix,ymd){const lead=prefix+'-'+ymd+'-';const max=rows.slice(1).reduce((m,r)=>{const code=String(r[0]||'');return code.indexOf(lead)===0?Math.max(m,Number(code.slice(lead.length))||0):m},0);return lead+String(max+1).padStart(3,'0')}
function date_(value){return value instanceof Date?Utilities.formatDate(value,Session.getScriptTimeZone(),'yyyy-MM-dd'):String(value||'')}
function json_(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON)}
