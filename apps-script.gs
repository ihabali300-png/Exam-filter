// كود شيت فلتر عملاء الفحص — يُلصق في Extensions > Apps Script داخل الشيت
const TOKEN = 'tabukpark';
const C_SHEET = 'العملاء', B_SHEET = 'المواعيد', S_SHEET = 'الملخص';
const C_HEAD = ['التاريخ','الوقت','العمر','سبب الزيارة','الأعراض','النسبة','الاحتمالية','المشكلة المتوقعة','التأمين','الاهتمام بالشراء','حجز موعد'];
const B_HEAD = ['المعرف','التاريخ','وقت الحجز','الاسم','الجوال','الموعد','الحالة','النسبة','الاحتمالية','آخر تحديث'];

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(){ setup(); return out({ok:true, msg:'ready'}); }

function sh(name, head){
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s){ s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); s.getRange(1,1,1,head.length).setFontWeight('bold').setBackground('#13202E').setFontColor('#ffffff'); s.setRightToLeft(true); }
  return s;
}

function setup(){
  sh(C_SHEET, C_HEAD); sh(B_SHEET, B_HEAD);
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(S_SHEET)) return;
  const s = ss.insertSheet(S_SHEET, 0); s.setRightToLeft(true);
  const C = "'" + C_SHEET + "'!", B = "'" + B_SHEET + "'!";
  const d = 'TEXT($B$1,"yyyy-mm-dd")';
  const rows = [
    ['اليوم', '=TODAY()', 'كل الأيام'],
    ['عدد من فحص أو تم تقييمه', '=COUNTIF(' + C + 'A:A,' + d + ')', '=COUNTA(' + C + 'A:A)-1'],
    ['احتمالية عالية', '=COUNTIFS(' + C + 'A:A,' + d + ',' + C + 'G:G,"احتمالية عالية")', '=COUNTIF(' + C + 'G:G,"احتمالية عالية")'],
    ['احتمالية متوسطة', '=COUNTIFS(' + C + 'A:A,' + d + ',' + C + 'G:G,"احتمالية متوسطة")', '=COUNTIF(' + C + 'G:G,"احتمالية متوسطة")'],
    ['احتمالية ضعيفة', '=COUNTIFS(' + C + 'A:A,' + d + ',' + C + 'G:G,"احتمالية ضعيفة")', '=COUNTIF(' + C + 'G:G,"احتمالية ضعيفة")'],
    ['احتمالية ضعيفة جدًا', '=COUNTIFS(' + C + 'A:A,' + d + ',' + C + 'G:G,"احتمالية ضعيفة جدًا")', '=COUNTIF(' + C + 'G:G,"احتمالية ضعيفة جدًا")'],
    ['مواعيد محجوزة', '=COUNTIF(' + B + 'B:B,' + d + ')', '=COUNTA(' + B + 'A:A)-1'],
    ['حضر في موعده', '=COUNTIFS(' + B + 'B:B,' + d + ',' + B + 'G:G,"حضر")', '=COUNTIF(' + B + 'G:G,"حضر")'],
    ['لم يحضر', '=COUNTIFS(' + B + 'B:B,' + d + ',' + B + 'G:G,"لم يحضر")', '=COUNTIF(' + B + 'G:G,"لم يحضر")'],
    ['ملغي', '=COUNTIFS(' + B + 'B:B,' + d + ',' + B + 'G:G,"ملغي")', '=COUNTIF(' + B + 'G:G,"ملغي")']
  ];
  s.getRange(1, 1, rows.length, 3).setValues(rows);
  s.getRange('B1').setNumberFormat('yyyy-mm-dd');
  s.getRange('A1:C1').setFontWeight('bold').setBackground('#13202E').setFontColor('#ffffff');
  s.getRange('B1').setBackground('#2457D6');
  s.setColumnWidth(1, 220);
  s.getRange('A13').setValue('غيّر تاريخ الخانة B1 لتشوف أرقام أي يوم');
}

function doPost(e){
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.token !== TOKEN) return out({ok:false, error:'bad token'});
    setup();
    const t = v => "'" + (v === undefined || v === null ? '' : String(v));
    if (d.action === 'log'){
      sh(C_SHEET, C_HEAD).appendRow([t(d.date), t(d.time), d.age, t(d.why), t(d.sym), d.score, t(d.tier), t(d.dx), t(d.ins), t(d.buy), t(d.booked)]);
      return out({ok:true});
    }
    if (d.action === 'book'){
      const s = sh(B_SHEET, B_HEAD);
      const ids = s.getRange(1, 1, s.getLastRow(), 1).getDisplayValues().flat();
      if (!ids.includes(d.id)) s.appendRow([t(d.id), t(d.date), t(d.made), t(d.name), t(d.phone), t(d.time), t(d.status || 'قادم'), d.score, t(d.tier), new Date()]);
      return out({ok:true});
    }
    if (d.action === 'status'){
      const s = sh(B_SHEET, B_HEAD);
      const ids = s.getRange(1, 1, s.getLastRow(), 1).getDisplayValues().flat();
      const i = ids.indexOf(d.id);
      if (i > 0){ s.getRange(i + 1, 7).setValue(d.status); s.getRange(i + 1, 10).setValue(new Date()); }
      return out({ok:true});
    }
    if (d.action === 'bookings'){
      const v = sh(B_SHEET, B_HEAD).getDataRange().getDisplayValues().slice(1);
      const rows = v.filter(r => r[1] === d.date).map(r => ({id:r[0], date:r[1], made:r[2], name:r[3], phone:r[4], time:r[5], status:r[6], score:r[7], tier:r[8]}));
      return out({ok:true, rows});
    }
    if (d.action === 'search'){
      const v = sh(B_SHEET, B_HEAD).getDataRange().getDisplayValues().slice(1);
      const rows = v.filter(r => String(r[4]).includes(d.q)).slice(-30).map(r => ({id:r[0], date:r[1], made:r[2], name:r[3], phone:r[4], time:r[5], status:r[6], score:r[7], tier:r[8]}));
      return out({ok:true, rows});
    }
    if (d.action === 'edit' || d.action === 'delete'){
      const s = sh(B_SHEET, B_HEAD);
      const ids = s.getRange(1, 1, s.getLastRow(), 1).getDisplayValues().flat();
      const i = ids.indexOf(d.id);
      if (i > 0){
        if (d.action === 'delete') s.deleteRow(i + 1);
        else {
          s.getRange(i + 1, 2).setValue(t(d.date)); s.getRange(i + 1, 4).setValue(t(d.name)); s.getRange(i + 1, 5).setValue(t(d.phone));
          s.getRange(i + 1, 6).setValue(t(d.time)); s.getRange(i + 1, 7).setValue(d.status || 'قادم'); s.getRange(i + 1, 10).setValue(new Date());
        }
      }
      return out({ok:true});
    }
    return out({ok:false, error:'unknown action'});
  } catch(err){
    return out({ok:false, error:String(err)});
  } finally { lock.releaseLock(); }
}
