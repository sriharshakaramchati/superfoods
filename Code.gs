/** 360 Super Food - orders backend (Google Apps Script web app).
 * Script properties: CLIENT_ID, ADMIN_EMAILS (comma list), SHEET_ID (set by setup()), NOTIFY_EMAILS (comma list, optional)
 */
var P = PropertiesService.getScriptProperties();
var HEAD = ['Order ID','Placed at','Status','Name','Phone','Email','Address','City','Pincode','Items','Total items','Notes','Admin note'];

function setup() {
  var id = P.getProperty('SHEET_ID'), ss;
  if (id) { ss = SpreadsheetApp.openById(id); }
  else { ss = SpreadsheetApp.create('360 Super Food - Orders'); P.setProperty('SHEET_ID', ss.getId()); }
  var sh = ss.getSheetByName('Orders') || ss.insertSheet('Orders');
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.getRange(1,1,1,HEAD.length).setFontWeight('bold').setBackground('#FFD400');
    sh.setFrozenRows(1);
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['New','Confirmed','Packed','Dispatched','Delivered','Cancelled'], true).build();
    sh.getRange('C2:C2000').setDataValidation(rule);
    sh.setColumnWidth(10, 420); sh.setColumnWidth(7, 280);
  }
  var d = ss.getSheetByName('Sheet1'); if (d) ss.deleteSheet(d);
  (P.getProperty('ADMIN_EMAILS')||'').split(',').forEach(function(e){ e=e.trim(); if(e && e!==Session.getEffectiveUser().getEmail()) { try{ ss.addEditor(e);}catch(x){} } });
  Logger.log(ss.getUrl());
  return ss.getUrl();
}

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function verify(token) {
  if (!token) throw new Error('Please sign in with Google first.');
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), {muteHttpExceptions:true});
  if (r.getResponseCode() !== 200) throw new Error('Sign-in expired. Please sign in again.');
  var t = JSON.parse(r.getContentText());
  if (t.aud !== P.getProperty('CLIENT_ID')) throw new Error('Sign-in not valid for this site.');
  if (String(t.email_verified) !== 'true') throw new Error('Google email not verified.');
  return {email: String(t.email).toLowerCase(), name: t.name || ''};
}
function isAdmin(email){ return (P.getProperty('ADMIN_EMAILS')||'').toLowerCase().split(',').map(function(s){return s.trim();}).indexOf(email) >= 0; }
function sheet(){ return SpreadsheetApp.openById(P.getProperty('SHEET_ID')).getSheetByName('Orders'); }
function clean(s, n){ return String(s==null?'':s).replace(/[\u0000-\u001f]/g,' ').trim().slice(0, n||200); }
function safeCell(s){ return /^[=+\-@]/.test(s) ? "'" + s : s; } // block formula injection

function doGet(e) {
  var a = (e.parameter||{}).action;
  if (a === 'ping' || !a) return out({ok:true, service:'360-superfood-orders'});
  return out({ok:false, error:'Use POST'});
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    var u = verify(b.token);
    if (b.action === 'order') return out(placeOrder(b, u));
    if (b.action === 'mine') return out(listOrders(u, false));
    if (b.action === 'list') { if(!isAdmin(u.email)) throw new Error('Not allowed.'); return out(listOrders(u, true)); }
    if (b.action === 'setStatus') { if(!isAdmin(u.email)) throw new Error('Not allowed.'); return out(setStatus(b)); }
    if (b.action === 'whoami') return out({ok:true, email:u.email, admin:isAdmin(u.email)});
    throw new Error('Unknown action');
  } catch (err) { return out({ok:false, error:String(err.message||err)}); }
}

function placeOrder(b, u) {
  var items = (b.items||[]).slice(0, 80).map(function(i){ return {name:clean(i.name,80), qty:Math.max(1, Math.min(999, parseInt(i.qty,10)||1)), note:clean(i.note,80)}; });
  if (!items.length) throw new Error('Your cart is empty.');
  var c = b.contact||{};
  var phone = clean(c.phone,20).replace(/[^\d+]/g,'');
  if (phone.replace(/\D/g,'').length < 10) throw new Error('Please enter a valid phone number.');
  var name = clean(c.name,80), addr = clean(c.address,300);
  if (!name || !addr) throw new Error('Name and address are required.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet();
    var n = sh.getLastRow(); // header counts as 1
    var id = 'SF-' + Utilities.formatDate(new Date(),'Asia/Kolkata','yyMMdd') + '-' + ('000'+n).slice(-3);
    var text = items.map(function(i){ return i.name + ' x ' + i.qty + (i.note ? ' (' + i.note + ')' : ''); }).join('\n');
    var total = items.reduce(function(s,i){return s+i.qty;},0);
    sh.appendRow([id, new Date(), 'New', safeCell(name), safeCell(phone), u.email, safeCell(addr), safeCell(clean(c.city,60)), safeCell(clean(c.pincode,10)), text, total, safeCell(clean(b.notes,300)), '']);
  } finally { lock.releaseLock(); }
  var body = 'New order ' + id + '\n\nName: ' + name + '\nPhone: ' + phone + '\nEmail: ' + u.email + '\nAddress: ' + addr + ', ' + clean(c.city,60) + ' ' + clean(c.pincode,10) + '\n\nItems:\n' + text + '\n\nNotes: ' + clean(b.notes,300) + '\n\nOpen orders: ' + SpreadsheetApp.openById(P.getProperty('SHEET_ID')).getUrl();
  var notify = (P.getProperty('NOTIFY_EMAILS')||P.getProperty('ADMIN_EMAILS')||'');
  try { if (notify) MailApp.sendEmail({to:notify, subject:'New order ' + id + ' - ' + name, body:body, replyTo:u.email}); } catch(x){}
  try { MailApp.sendEmail({to:u.email, subject:'We got your order ' + id + ' - 360 Super Food', body:'Hi ' + name + ',\n\nThank you. We received your order and will call or WhatsApp you on ' + phone + ' to confirm the price and delivery.\n\n' + text + '\n\n360 Super Food Products Company\nCustomer care: 9705555360'}); } catch(x){}
  return {ok:true, orderId:id};
}

function rows(){ var sh=sheet(); var n=sh.getLastRow(); if(n<2) return []; return sh.getRange(2,1,n-1,HEAD.length).getValues(); }
function listOrders(u, admin) {
  var out_ = rows().map(function(r,i){ return {row:i+2,id:r[0],at:r[1] instanceof Date?r[1].toISOString():String(r[1]),status:r[2],name:r[3],phone:r[4],email:r[5],address:r[6],city:r[7],pincode:r[8],items:r[9],total:r[10],notes:r[11],adminNote:r[12]}; });
  if (!admin) out_ = out_.filter(function(o){ return o.email === u.email; }).map(function(o){ return {id:o.id,at:o.at,status:o.status,items:o.items}; });
  return {ok:true, orders:out_.reverse()};
}
function setStatus(b) {
  var ok = ['New','Confirmed','Packed','Dispatched','Delivered','Cancelled'];
  if (ok.indexOf(b.status) < 0) throw new Error('Bad status');
  var sh = sheet(), r = parseInt(b.row,10);
  if (sh.getRange(r,1).getValue() !== b.id) throw new Error('Order moved, refresh.');
  sh.getRange(r,3).setValue(b.status);
  if (b.adminNote != null) sh.getRange(r,13).setValue(safeCell(clean(b.adminNote,200)));
  return {ok:true};
}
