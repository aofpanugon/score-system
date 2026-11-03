/**
 * ระบบเก็บคะแนนและส่งงานออนไลน์ — ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
 *
 * ทำหน้าที่เป็น API ให้หน้าเว็บที่โฮสต์บน GitHub Pages (หรือเปิดผ่าน Apps Script ก็ได้)
 * ข้อมูลเก็บใน Google Sheet, ไฟล์งานเก็บใน Google Drive ของเจ้าของสคริปต์
 *
 * วิธีติดตั้งแบบย่อ (รายละเอียดใน README.md):
 *   1) สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script แล้ววางไฟล์นี้ทั้งหมดแทน Code.gs
 *   2) รันฟังก์ชัน setup() หนึ่งครั้ง (อนุญาตสิทธิ์) — สร้างชีต โฟลเดอร์ Drive และรหัสครูตัวอย่าง T001
 *   3) (ไม่บังคับ) รัน seedDemo() เพื่อสร้างข้อมูลทดลอง
 *   4) Deploy > New deployment > Web app: Execute as = Me, Who has access = Anyone แล้วคัดลอก URL (/exec)
 *
 * PIN เริ่มต้นของแอดมินและครูทุกคนคือ 123456 (6 หลัก) ให้เปลี่ยนทันทีหลังเข้าใช้งานครั้งแรก
 */

const DEFAULT_PIN = '123456';
const MAX_FAILS = 5;          // กรอก PIN ผิดครบกี่ครั้งจึงล็อก
const LOCK_SECONDS = 900;     // ล็อก 15 นาที
const MAX_FILES = 5, MAX_BYTES = 25 * 1024 * 1024;
const SPECIAL = ['ร.', '0', 'มส.', 'มผ.', 'ผ.'];
const BLOCKED_EXT = /\.(exe|bat|cmd|com|msi|scr|vbs|js|jar|ps1|sh|apk|dll|lnk|hta|reg)$/i;

/** นิยามชีต: h = หัวคอลัมน์, txt = คอลัมน์ที่ต้องเก็บเป็นข้อความ (กันเลขศูนย์นำหน้า เช่น 02900 หาย) */
const T = {
  Teachers:    { h: ['id', 'prefix', 'name', 'surname', 'pinHash'], txt: [1] },
  Students:    { h: ['id', 'prefix', 'name', 'surname', 'classRoom', 'number'], txt: [1] },
  Subjects:    { h: ['code', 'name', 'gradeLevel', 'semester', 'teacherId'], txt: [1] },
  Enrollments: { h: ['studentId', 'subjectCode'], txt: [1, 2] },
  Assignments: { h: ['id', 'subjectCode', 'title', 'category', 'maxScore', 'deadline', 'status', 'description'], txt: [1, 2, 6] },
  Submissions: { h: ['submissionId', 'studentId', 'assignmentId', 'timestamp', 'status', 'score', 'feedback', 'files'], txt: [2, 3, 4] },
  Scores:      { h: ['studentId', 'subjectCode', 'midterm', 'finalExam', 'special'], txt: [1, 2] }
};

class UErr extends Error {}
const fail_ = m => { throw new UErr(m); };

/* ====================== ตัวช่วยพื้นฐาน ====================== */
function sheet_(n) {
  const s = SpreadsheetApp.getActive().getSheetByName(n);
  if (!s) fail_('ยังไม่ได้ติดตั้งระบบ: ให้รันฟังก์ชัน setup() ก่อน (ไม่พบชีต ' + n + ')');
  return s;
}
function all_(n) { const v = sheet_(n).getDataRange().getValues(); v.shift(); return v.filter(r => r.some(c => c !== '')); }
function put_(n, pred, values) { // เพิ่มหรือแก้แถวแรกที่ตรงเงื่อนไข
  const sh = sheet_(n), d = sh.getDataRange().getValues();
  const i = d.findIndex((r, k) => k > 0 && pred(r)), row = i > 0 ? i + 1 : sh.getLastRow() + 1;
  T[n].txt.forEach(c => sh.getRange(row, c).setNumberFormat('@'));
  sh.getRange(row, 1, 1, values.length).setValues([values]);
}
function del_(n, pred) {
  const sh = sheet_(n), d = sh.getDataRange().getValues();
  let k = 0;
  for (let i = d.length - 1; i > 0; i--) if (pred(d[i])) { sh.deleteRow(i + 1); k++; }
  return k;
}
const prop_ = k => PropertiesService.getScriptProperties().getProperty(k);
const setProp_ = (k, v) => PropertiesService.getScriptProperties().setProperty(k, v);
const id5_ = v => String(v).trim().padStart(5, '0');
const name_ = r => (r[1] || '') + r[2] + ' ' + (r[3] || '');
const toNum_ = v => { const n = Number(v); return isNaN(n) ? 0 : n; };
const round_ = x => Math.round(x * 100) / 100;
const tz_ = () => Session.getScriptTimeZone();
const fmt_ = (d, f) => Utilities.formatDate(d, tz_(), f);
const fmtDate_ = d => d instanceof Date ? fmt_(d, "yyyy-MM-dd'T'HH:mm") : String(d || '');
const fmtTs_ = d => d instanceof Date ? fmt_(d, 'yyyy-MM-dd HH:mm') : String(d || '');
const gradeOf_ = t => t >= 80 ? '4' : t >= 75 ? '3.5' : t >= 70 ? '3' : t >= 65 ? '2.5' : t >= 60 ? '2' : t >= 55 ? '1.5' : t >= 50 ? '1' : '0';
const safeName_ = n => String(n || 'file').replace(/[\\\/:*?"<>|\u0000-\u001f]/g, '_').slice(0, 100);
const isPin_ = p => /^\d{6}$/.test(String(p));

function hash_(pin) {
  const salt = prop_('SALT');
  if (!salt) fail_('ยังไม่ได้ติดตั้งระบบ: ให้รันฟังก์ชัน setup() ก่อน');
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + String(pin)));
}

/* ----- ล็อกกันเดา PIN ----- */
function guard_(key) {
  if (Number(CacheService.getScriptCache().get('f:' + key) || 0) >= MAX_FAILS) fail_('กรอกผิดหลายครั้ง กรุณารอ 15 นาทีแล้วลองใหม่');
}
function bad_(key, msg) {
  const c = CacheService.getScriptCache();
  c.put('f:' + key, String(Number(c.get('f:' + key) || 0) + 1), LOCK_SECONDS);
  fail_(msg);
}
const good_ = key => CacheService.getScriptCache().remove('f:' + key);

function teacherAuth_(cred) { // cred = "รหัสครู:PIN"
  const s = String(cred || ''), i = s.indexOf(':');
  if (i < 1) fail_('รหัสครูหรือ PIN ไม่ถูกต้อง');
  const id = s.slice(0, i), pin = s.slice(i + 1), key = 't:' + id;
  guard_(key);
  const r = all_('Teachers').find(x => String(x[0]) === id);
  if (!r || !isPin_(pin) || (r[4] || hash_(DEFAULT_PIN)) !== hash_(pin)) bad_(key, 'รหัสครูหรือ PIN ไม่ถูกต้อง');
  good_(key);
  return { id: String(r[0]), name: name_(r) };
}
function adminAuth_(pin) {
  guard_('admin');
  if (!isPin_(pin) || hash_(pin) !== (prop_('ADMIN_PIN_HASH') || hash_(DEFAULT_PIN))) bad_('admin', 'PIN แอดมินไม่ถูกต้อง');
  good_('admin');
}
function ownedSubject_(t, code) {
  const r = all_('Subjects').find(x => String(x[0]) === String(code));
  if (!r || String(r[4]) !== t.id) fail_('ไม่ใช่รายวิชาของคุณ');
  return r;
}
function ownedAssignment_(t, id) {
  const a = all_('Assignments').find(x => String(x[0]) === String(id));
  if (!a) fail_('ไม่พบชิ้นงาน');
  ownedSubject_(t, a[1]);
  return a;
}
const enrolledIds_ = code => all_('Enrollments').filter(r => String(r[1]) === String(code)).map(r => id5_(r[0]));

/* ---- คำนวณคะแนน ---- */
function totals_(sid, code, asg, subs, scores) {
  const sc = scores.find(x => id5_(x[0]) === sid && String(x[1]) === String(code));
  const midterm = sc ? toNum_(sc[2]) : 0, finalExam = sc ? toNum_(sc[3]) : 0, special = sc && sc[4] ? String(sc[4]) : '';
  const sum = c => asg.filter(a => a[3] === c).reduce((n, a) => {
    const s = subs.find(x => id5_(x[1]) === sid && x[2] === a[0]);
    return n + (s && s[5] !== '' ? toNum_(s[5]) : 0);
  }, 0);
  const work = round_(sum('work')), collect = round_(sum('collect')), test = round_(sum('test'));
  const total = round_(work + collect + test + midterm + finalExam);
  return { work, collect, test, midterm, finalExam, total, special, grade: special || gradeOf_(total) };
}

/* ====================== ส่วนสาธารณะ (นักเรียน) ====================== */
function getPublicConfig() {
  return {
    success: true,
    title: prop_('TITLE') || 'ระบบเก็บคะแนนและส่งงานออนไลน์', tagline: prop_('TAGLINE') || '',
    school: prop_('SCHOOL') || '', semester: prop_('SEMESTER') || '', logo: prop_('LOGO_URL') || '',
    banners: JSON.parse(prop_('BANNERS') || '[]')
  };
}

function getStudentScorecard(code, subjectCode) {
  code = String(code || '').trim();
  if (!/^\d{5}$/.test(code)) fail_('กรอกรหัสนักเรียน 5 หลัก');
  const st = all_('Students').find(r => id5_(r[0]) === code);
  if (!st) fail_('ไม่พบรหัสนักเรียนนี้');
  const enr = all_('Enrollments').filter(r => id5_(r[0]) === code).map(r => String(r[1]));
  const subjects = all_('Subjects').filter(r => enr.indexOf(String(r[0])) >= 0);
  if (!subjects.length) fail_('นักเรียนยังไม่ได้ลงทะเบียนรายวิชาใด');
  const cur = subjects.some(r => String(r[0]) === String(subjectCode)) ? String(subjectCode) : String(subjects[0][0]);
  const asg = all_('Assignments').filter(a => String(a[1]) === cur);
  const ids = asg.map(a => a[0]);
  const subs = all_('Submissions').filter(s => id5_(s[1]) === code && ids.indexOf(s[2]) >= 0);
  const card = totals_(code, cur, asg, subs, all_('Scores'));
  card.items = asg.map(a => {
    const s = subs.find(x => x[2] === a[0]), has = s && s[5] !== '';
    let files = []; try { files = s && s[7] ? JSON.parse(s[7]) : []; } catch (e) {}
    return { title: a[2], max: toNum_(a[4]), score: has ? toNum_(s[5]) : null,
      state: a[3] === 'collect' ? (has ? 'graded' : 'unrec') : s ? (has ? 'graded' : 'pending') : 'none',
      late: s && s[4] === 'ส่งช้า' ? 'ส่งช้า' : '', files, feedback: s ? String(s[6] || '') : '' };
  });
  return { success: true, student: { id: code, name: name_(st), classRoom: st[4], number: st[5] },
    subjects: subjects.map(r => ({ code: String(r[0]), name: r[1] })), current: cur, card };
}

function findPendingWork(q) {
  q = String(q || '').trim();
  if (!q) fail_('กรอกรหัสหรือชื่อนักเรียน');
  const sts = all_('Students');
  let st;
  if (/^\d{5}$/.test(q)) st = sts.filter(r => id5_(r[0]) === q);
  else { if (q.length < 2) fail_('พิมพ์ชื่ออย่างน้อย 2 ตัวอักษร'); st = sts.filter(r => name_(r).indexOf(q) >= 0).slice(0, 10); }
  if (!st.length) fail_('ไม่พบนักเรียน');
  const matches = st.map(r => ({ id: id5_(r[0]), name: name_(r), classRoom: r[4], number: r[5] }));
  if (st.length > 1) return { success: true, matches, pending: null };
  const sid = matches[0].id;
  const enr = all_('Enrollments').filter(r => id5_(r[0]) === sid).map(r => String(r[1]));
  const subs = all_('Submissions').filter(s => id5_(s[1]) === sid), subj = all_('Subjects');
  const pending = all_('Assignments')
    .filter(a => a[3] === 'work' && a[6] === 'Active' && enr.indexOf(String(a[1])) >= 0 && !subs.some(s => s[2] === a[0]))
    .map(a => { const sj = subj.find(s => String(s[0]) === String(a[1]));
      return { id: a[0], subjectCode: String(a[1]), subjectName: sj ? sj[1] : '', title: a[2], maxScore: toNum_(a[4]), deadline: fmtDate_(a[5]), description: a[7] || '' }; })
    .sort((x, y) => String(x.deadline).localeCompare(String(y.deadline)));
  return { success: true, matches, pending };
}

function getRootFolder_() {
  const id = prop_('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('ไฟล์ส่งงาน - ' + SpreadsheetApp.getActive().getName());
  setProp_('FOLDER_ID', f.getId());
  return f;
}
function subFolder_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function fileIdFromUrl_(u) { const m = /[-\w]{25,}/.exec(String(u || '')); return m ? m[0] : ''; }

function submitAssignment(p) { // ไม่ล็อกทั้งฟังก์ชัน เพื่อไม่ให้การอัปโหลดไฟล์ใหญ่บล็อกคนอื่น
  p = p || {};
  const sid = String(p.studentId || '').trim();
  if (!/^\d{5}$/.test(sid)) fail_('รหัสนักเรียนต้องเป็นตัวเลข 5 หลัก');
  if (!all_('Students').some(r => id5_(r[0]) === sid)) fail_('ไม่พบรหัสนักเรียนนี้');
  const a = all_('Assignments').find(x => x[0] === p.assignmentId);
  if (!a || a[3] !== 'work') fail_('ไม่พบชิ้นงานที่ต้องส่ง');
  if (a[6] !== 'Active') fail_('ชิ้นงานนี้ปิดรับแล้ว');
  if (!all_('Enrollments').some(r => id5_(r[0]) === sid && String(r[1]) === String(a[1]))) fail_('นักเรียนไม่ได้ลงทะเบียนในรายวิชานี้');
  const files = Array.isArray(p.files) ? p.files : [];
  if (!files.length || files.length > MAX_FILES) fail_('แนบไฟล์ 1-' + MAX_FILES + ' ไฟล์');
  const decoded = files.map(f => {
    if (BLOCKED_EXT.test(String(f.fileName || ''))) fail_('ไม่รับไฟล์ชนิดนี้: ' + f.fileName);
    return { name: safeName_(f.fileName), mime: f.mimeType || 'application/octet-stream', bytes: Utilities.base64Decode(String(f.fileData || '')) };
  });
  if (decoded.reduce((n, f) => n + f.bytes.length, 0) > MAX_BYTES) fail_('ไฟล์รวมเกิน 25 MB');
  const old = all_('Submissions').find(s => id5_(s[1]) === sid && s[2] === a[0]);
  if (old && old[5] !== '') fail_('งานนี้ตรวจและให้คะแนนแล้ว ไม่สามารถส่งซ้ำได้');

  const folder = subFolder_(subFolder_(getRootFolder_(), String(a[1])), String(a[0]));
  const saved = decoded.map(f => {
    const file = folder.createFile(Utilities.newBlob(f.bytes, f.mime, sid + '_' + Date.now() + '_' + f.name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { name: f.name, url: file.getUrl() };
  });
  const dl = a[5] ? new Date(String(a[5])) : null, now = new Date();
  const status = dl && !isNaN(dl) && now > dl ? 'ส่งช้า' : 'ตรงเวลา';
  const subId = old ? old[0] : 'SUB-' + fmt_(now, 'yyMMdd') + '-' + Math.floor(Math.random() * 1e6).toString().padStart(6, '0');
  withLock_(() => put_('Submissions', r => id5_(r[1]) === sid && r[2] === a[0], [subId, sid, a[0], fmtTs_(now), status, '', '', JSON.stringify(saved)]));
  if (old) { try { JSON.parse(old[7] || '[]').forEach(f => { const id = fileIdFromUrl_(f.url); if (id) DriveApp.getFileById(id).setTrashed(true); }); } catch (e) {} }
  return { success: true, message: old ? 'ส่งงานใหม่แทนที่ฉบับเดิมแล้ว' : 'ส่งงานสำเร็จ', submissionId: subId, subjectCode: String(a[1]), timestamp: fmtTs_(now), status, files: saved };
}

/* ====================== ครูผู้สอน ====================== */
function teacherDashboard(cred) {
  const t = teacherAuth_(cred);
  const enr = all_('Enrollments'), subs = all_('Submissions');
  const mine = all_('Subjects').filter(r => String(r[4]) === t.id), codes = mine.map(r => String(r[0]));
  const asg = all_('Assignments').filter(r => codes.indexOf(String(r[1])) >= 0);
  const folder = prop_('FOLDER_ID'), myStu = {};
  enr.filter(r => codes.indexOf(String(r[1])) >= 0).forEach(r => myStu[id5_(r[0])] = 1);
  const cnt = code => enr.filter(r => String(r[1]) === code).length;
  return {
    success: true, teacher: t, sheetUrl: SpreadsheetApp.getActive().getUrl(), folderUrl: folder ? 'https://drive.google.com/drive/folders/' + folder : '',
    counts: { subjects: mine.length, students: Object.keys(myStu).length, assignments: asg.length, submissions: subs.filter(s => asg.some(a => a[0] === s[2])).length },
    settings: { semester: prop_('SEMESTER') || '' },
    subjects: mine.map(r => ({ code: String(r[0]), name: r[1], gradeLevel: r[2], semester: r[3], studentCount: cnt(String(r[0])) })),
    assignments: asg.map(a => ({ id: a[0], subjectCode: String(a[1]), title: a[2], category: a[3], maxScore: toNum_(a[4]), deadline: fmtDate_(a[5]), status: a[6], description: a[7] || '',
      submittedCount: subs.filter(s => s[2] === a[0] && (a[3] !== 'collect' || s[5] !== '')).length, totalEnrolled: cnt(String(a[1])) }))
  };
}

function teacherChangePin(cred, newPin) {
  const t = teacherAuth_(cred);
  if (!isPin_(newPin)) fail_('PIN ต้องเป็นตัวเลข 6 หลัก');
  const r = all_('Teachers').find(x => String(x[0]) === t.id);
  put_('Teachers', x => String(x[0]) === t.id, [r[0], r[1], r[2], r[3], hash_(newPin)]);
  return { success: true, message: 'เปลี่ยน PIN เรียบร้อยแล้ว' };
}

function saveSubject(cred, s) {
  const t = teacherAuth_(cred);
  s = s || {};
  const code = String(s.code || '').trim(), name = String(s.name || '').trim();
  if (!code || !name) fail_('กรอกรหัสวิชาและชื่อวิชา');
  if (code.length > 30 || name.length > 120) fail_('รหัสหรือชื่อวิชายาวเกินไป');
  const old = all_('Subjects').find(r => String(r[0]) === code);
  if (old && String(old[4]) !== t.id) fail_('รหัสวิชานี้เป็นของครูท่านอื่น');
  put_('Subjects', r => String(r[0]) === code, [code, name, String(s.gradeLevel || '').trim(), String(s.semester || '').trim(), t.id]);
  return { success: true, message: 'บันทึกรายวิชาแล้ว' };
}

function deleteSubject(cred, code) {
  const t = teacherAuth_(cred); ownedSubject_(t, code);
  const ids = all_('Assignments').filter(a => String(a[1]) === String(code)).map(a => a[0]);
  del_('Submissions', r => ids.indexOf(r[2]) >= 0);
  del_('Assignments', r => String(r[1]) === String(code));
  del_('Enrollments', r => String(r[1]) === String(code));
  del_('Scores', r => String(r[1]) === String(code));
  del_('Subjects', r => String(r[0]) === String(code));
  return { success: true, message: 'ลบรายวิชาแล้ว' };
}

function saveAssignment(cred, a) {
  const t = teacherAuth_(cred);
  a = a || {};
  const id = String(a.id || '').trim();
  if (!/^[A-Za-z0-9ก-๙_.-]{1,40}$/.test(id)) fail_('รหัสชิ้นงานใช้ตัวอักษร ตัวเลข - _ . ไม่เกิน 40 ตัว');
  ownedSubject_(t, a.subjectCode);
  const old = all_('Assignments').find(r => String(r[0]) === id);
  if (old) ownedSubject_(t, old[1]);
  if (['work', 'test', 'collect'].indexOf(a.category) < 0) fail_('หมวดคะแนนไม่ถูกต้อง');
  const max = Number(a.maxScore);
  if (!(max > 0 && max <= 100)) fail_('คะแนนเต็มต้องอยู่ระหว่าง 1-100');
  if (['Active', 'Closed'].indexOf(a.status) < 0) fail_('สถานะไม่ถูกต้อง');
  if (!String(a.title || '').trim()) fail_('กรอกชื่อชิ้นงาน');
  if (a.category !== 'collect' && !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(String(a.deadline || ''))) fail_('กำหนดวันและเวลาส่งงาน');
  put_('Assignments', r => String(r[0]) === id, [id, String(a.subjectCode), String(a.title).trim(), a.category, max, a.category === 'collect' ? '' : a.deadline, a.status, String(a.description || '').slice(0, 300)]);
  return { success: true, message: 'บันทึกชิ้นงานแล้ว' };
}

function deleteAssignment(cred, id) {
  const t = teacherAuth_(cred); ownedAssignment_(t, id);
  del_('Submissions', r => r[2] === id);
  del_('Assignments', r => String(r[0]) === String(id));
  return { success: true, message: 'ลบชิ้นงานแล้ว' };
}

function getAssignmentSubmissions(cred, id) {
  const t = teacherAuth_(cred), a = ownedAssignment_(t, id);
  const subs = all_('Submissions').filter(s => s[2] === id), sts = all_('Students'), en = enrolledIds_(a[1]);
  return {
    success: true, assignment: { id: a[0], title: a[2], category: a[3], maxScore: toNum_(a[4]) },
    students: en.map(sid => { const r = sts.find(x => id5_(x[0]) === sid), s = subs.find(x => id5_(x[1]) === sid);
      let files = []; try { files = s && s[7] ? JSON.parse(s[7]) : []; } catch (e) {}
      return { id: sid, name: r ? name_(r) : sid, submitted: !!s, score: s && s[5] !== '' ? toNum_(s[5]) : '', feedback: s ? String(s[6] || '') : '', at: s ? fmtTs_(s[3]) : '', status: s ? String(s[4]) : '', files }; })
  };
}

function saveAssignmentGrades(cred, id, list) {
  const t = teacherAuth_(cred), a = ownedAssignment_(t, id), max = toNum_(a[4]);
  list = Array.isArray(list) ? list : [];
  const en = enrolledIds_(a[1]), subs = all_('Submissions').filter(s => s[2] === id);
  list.forEach(g => {
    const v = Number(g.score);
    if (String(g.score) === '' || isNaN(v) || v < 0 || v > max) fail_('คะแนนของรหัส ' + g.studentId + ' ต้องอยู่ระหว่าง 0 ถึง ' + max);
    if (en.indexOf(id5_(g.studentId)) < 0) fail_('รหัส ' + g.studentId + ' ไม่ได้ลงทะเบียนในรายวิชานี้');
  });
  let n = 0;
  list.forEach(g => {
    const sid = id5_(g.studentId), s = subs.find(x => id5_(x[1]) === sid), fb = String(g.feedback || '').slice(0, 300);
    if (s) put_('Submissions', r => id5_(r[1]) === sid && r[2] === id, [s[0], sid, id, fmtTs_(s[3]), s[4], Number(g.score), fb, s[7]]), n++;
    else if (a[3] === 'collect') put_('Submissions', r => id5_(r[1]) === sid && r[2] === id, ['SC-' + Date.now() + '-' + sid, sid, id, fmtTs_(new Date()), 'ลงคะแนนโดยครู', Number(g.score), fb, '[]']), n++;
  });
  return { success: true, message: 'บันทึกคะแนนแล้ว ' + n + ' รายการ' };
}

function getGradebook(cred, code) {
  const t = teacherAuth_(cred); ownedSubject_(t, code);
  const asg = all_('Assignments').filter(a => String(a[1]) === String(code)), ids = asg.map(a => a[0]);
  const subs = all_('Submissions').filter(s => ids.indexOf(s[2]) >= 0), scores = all_('Scores'), sts = all_('Students');
  const rows = enrolledIds_(code).map(sid => {
    const r = sts.find(x => id5_(x[0]) === sid) || [sid, '', sid, '', '', ''];
    const missing = asg.filter(a => a[3] === 'work' && a[6] === 'Active' && !subs.some(s => id5_(s[1]) === sid && s[2] === a[0])).map(a => a[0]);
    return Object.assign({ id: sid, name: name_(r), classRoom: r[4], number: r[5], missing }, totals_(sid, code, asg, subs, scores));
  }).sort((x, y) => String(x.classRoom).localeCompare(String(y.classRoom)) || toNum_(x.number) - toNum_(y.number) || x.id.localeCompare(y.id));
  return { success: true, rows };
}

function saveScoresBatch(cred, list) {
  const t = teacherAuth_(cred);
  list = Array.isArray(list) ? list : [];
  const mine = {};
  list.forEach(r => {
    const sid = id5_(r.studentId), code = String(r.subjectCode);
    if (!mine[code]) { ownedSubject_(t, code); mine[code] = enrolledIds_(code); }
    if (mine[code].indexOf(sid) < 0) fail_('รหัส ' + sid + ' ไม่ได้ลงทะเบียนในรายวิชา ' + code);
    ['midterm', 'finalExam'].forEach(k => { const v = r[k] === '' || r[k] == null ? 0 : Number(r[k]); if (isNaN(v) || v < 0 || v > 100) fail_('คะแนนกลางภาค/ปลายภาคของ ' + sid + ' ต้องอยู่ระหว่าง 0-100'); });
    if (r.special && SPECIAL.indexOf(r.special) < 0) fail_('เกรดพิเศษไม่ถูกต้อง');
  });
  list.forEach(r => {
    const sid = id5_(r.studentId), code = String(r.subjectCode);
    put_('Scores', x => id5_(x[0]) === sid && String(x[1]) === code, [sid, code, toNum_(r.midterm), toNum_(r.finalExam), r.special || '']);
  });
  return { success: true, message: 'บันทึกสมุดคะแนนแล้ว' };
}

function addStudentIfMissing_(id, prefix, name, surname, cls, no) {
  if (all_('Students').some(r => id5_(r[0]) === id)) return false;
  put_('Students', r => id5_(r[0]) === id, [id, prefix || '', name, surname || '', cls || '', no || '']);
  return true;
}
function enroll_(sid, code) {
  if (all_('Enrollments').some(r => id5_(r[0]) === sid && String(r[1]) === String(code))) return false;
  put_('Enrollments', r => id5_(r[0]) === sid && String(r[1]) === String(code), [sid, String(code)]);
  return true;
}

function importStudentsFromCSV(cred, text, code) { // รหัส,คำนำหน้า,ชื่อ,นามสกุล,ระดับชั้น,เลขที่
  const t = teacherAuth_(cred); ownedSubject_(t, code);
  const lines = String(text || '').replace(/^﻿/, '').split(/\r?\n/).filter(l => l.trim()).slice(0, 500);
  let created = 0, enrolled = 0, skipped = 0;
  lines.forEach(l => {
    const c = l.split(/[,\t]/).map(x => x.trim());
    if (!/^\d{1,5}$/.test(c[0]) || !c[2]) { skipped++; return; } // ข้ามหัวตารางและแถวที่ไม่ครบ
    const id = id5_(c[0]);
    if (addStudentIfMissing_(id, c[1], c[2], c[3], c[4], c[5])) created++;
    if (enroll_(id, code)) enrolled++;
  });
  return { success: true, message: 'นำเข้าแล้ว: นักเรียนใหม่ ' + created + ' คน, เพิ่มเข้าวิชา ' + enrolled + ' คน, ข้าม ' + skipped + ' แถว' };
}

function enrollStudents(cred, code, idsText) {
  const t = teacherAuth_(cred); ownedSubject_(t, code);
  const ids = String(idsText || '').split(/[\s,;]+/).filter(Boolean).slice(0, 200);
  if (!ids.length) fail_('กรอกรหัสนักเรียนอย่างน้อย 1 รหัส');
  const sts = all_('Students'), unknown = [];
  let n = 0;
  ids.forEach(x => { const id = id5_(x); if (!/^\d{5}$/.test(id) || !sts.some(r => id5_(r[0]) === id)) { unknown.push(x); return; } if (enroll_(id, code)) n++; });
  if (unknown.length && !n) fail_('ไม่พบรหัสนักเรียน: ' + unknown.join(', ') + ' (แอดมินต้องเพิ่มนักเรียนในระบบก่อน หรือใช้นำเข้า CSV)');
  return { success: true, message: 'เพิ่มเข้าวิชา ' + n + ' คน' + (unknown.length ? ' (ไม่พบรหัส: ' + unknown.join(', ') + ')' : '') };
}

function deleteStudent(cred, id, code) { // ลบออกจากวิชานี้ (ไม่ลบข้อมูลนักเรียนในระบบ)
  const t = teacherAuth_(cred); ownedSubject_(t, code);
  const sid = id5_(id);
  del_('Enrollments', r => id5_(r[0]) === sid && String(r[1]) === String(code));
  del_('Scores', r => id5_(r[0]) === sid && String(r[1]) === String(code));
  return { success: true, message: 'นำออกจากรายวิชาแล้ว' };
}

/* ====================== แอดมิน ====================== */
function adminLogin(pin) {
  adminAuth_(pin);
  return {
    success: true,
    users: {
      teachers: all_('Teachers').map(r => ({ id: String(r[0]), prefix: r[1], name: r[2], surname: r[3] })),
      students: all_('Students').map(r => ({ id: id5_(r[0]), prefix: r[1], name: r[2], surname: r[3], classRoom: r[4], number: r[5] }))
    },
    settings: { title: prop_('TITLE') || '', tagline: prop_('TAGLINE') || '', school: prop_('SCHOOL') || '', semester: prop_('SEMESTER') || '', logo: prop_('LOGO_URL') || '' },
    banners: JSON.parse(prop_('BANNERS') || '[]'),
    counts: { subjects: all_('Subjects').length, submissions: all_('Submissions').length }
  };
}

function adminSaveUser(pin, u) {
  adminAuth_(pin);
  u = u || {};
  const id = String(u.id || '').trim(), teacher = u.role === 'teacher';
  if (!String(u.name || '').trim()) fail_('กรอกชื่อ');
  if (u.pin && !isPin_(u.pin)) fail_('PIN ต้องเป็นตัวเลข 6 หลัก');
  if (teacher) {
    if (!/^[A-Za-z0-9_-]{2,20}$/.test(id)) fail_('รหัสครูใช้ตัวอักษรอังกฤษ/ตัวเลข 2-20 ตัว');
    const old = all_('Teachers').find(r => String(r[0]) === id);
    put_('Teachers', r => String(r[0]) === id, [id, u.prefix || '', u.name, u.surname || '', u.pin ? hash_(u.pin) : (old && old[4]) || hash_(DEFAULT_PIN)]);
  } else {
    if (!/^\d{5}$/.test(id)) fail_('รหัสนักเรียนต้องเป็นตัวเลข 5 หลัก');
    put_('Students', r => id5_(r[0]) === id, [id, u.prefix || '', u.name, u.surname || '', u.classRoom || '', u.number || '']);
  }
  return { success: true, message: 'บันทึกสมาชิกแล้ว' };
}

function adminDeleteUser(pin, role, id) {
  adminAuth_(pin);
  if (role === 'teacher') {
    if (all_('Subjects').some(r => String(r[4]) === String(id))) fail_('ครูท่านนี้ยังเป็นเจ้าของรายวิชา ให้ลบหรือย้ายรายวิชาก่อน');
    del_('Teachers', r => String(r[0]) === String(id));
  } else {
    const sid = id5_(id);
    del_('Enrollments', r => id5_(r[0]) === sid); del_('Scores', r => id5_(r[0]) === sid);
    del_('Submissions', r => id5_(r[1]) === sid); del_('Students', r => id5_(r[0]) === sid);
  }
  return { success: true, message: 'ลบสมาชิกแล้ว' };
}

function adminSaveSettings(pin, s) {
  adminAuth_(pin);
  s = s || {};
  if (s.clearLogo) { PropertiesService.getScriptProperties().deleteProperty('LOGO_URL'); return { success: true, message: 'ลบโลโก้แล้ว' }; }
  if (s.newPin && !isPin_(s.newPin)) fail_('PIN ต้องเป็นตัวเลข 6 หลัก');
  PropertiesService.getScriptProperties().setProperties({ TITLE: String(s.title || '').slice(0, 100), TAGLINE: String(s.tagline || '').slice(0, 80), SCHOOL: String(s.school || '').slice(0, 100), SEMESTER: String(s.semester || '').slice(0, 20) });
  if (s.newPin) setProp_('ADMIN_PIN_HASH', hash_(s.newPin));
  return { success: true, message: 'บันทึกการตั้งค่าแล้ว' };
}

function saveImage_(f, label, maxBytes) {
  f = f || {};
  if (!/^image\/(png|jpeg|webp)$/.test(f.mimeType)) fail_('รองรับเฉพาะ PNG/JPG/WebP');
  const bytes = Utilities.base64Decode(String(f.fileData || ''));
  if (bytes.length > maxBytes) fail_('ไฟล์รูปใหญ่เกินกำหนด');
  const file = subFolder_(getRootFolder_(), '_images').createFile(Utilities.newBlob(bytes, f.mimeType, label + '-' + Date.now()));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}
function adminUploadLogo(pin, f) { adminAuth_(pin); setProp_('LOGO_URL', saveImage_(f, 'logo', 2 * 1048576)); return { success: true }; }
function adminUploadBanner(pin, f) { adminAuth_(pin); return { success: true, url: saveImage_(f, 'banner', 3 * 1048576) }; } // หน้าเว็บย่อรูปเหลือกว้าง ≤1600px ก่อนส่ง
function adminSaveBanners(pin, list) {
  adminAuth_(pin);
  if (!Array.isArray(list) || list.length > 12) fail_('ใส่ภาพหน้าปกได้ไม่เกิน 12 ภาพ');
  const clean = list.map(b => ({ id: String(b.id || '').slice(0, 30), url: String(b.url || ''), title: String(b.title || '').slice(0, 80), subtitle: String(b.subtitle || '').slice(0, 120) }));
  if (clean.some(b => !/^https:\/\//.test(b.url))) fail_('ภาพหน้าปกต้องเป็นลิงก์ https:// (ระบบอัปโหลดจะสร้างลิงก์ให้เอง)');
  const json = JSON.stringify(clean);
  if (json.length > 8500) fail_('ข้อมูลภาพหน้าปกมากเกินไป ลดจำนวนภาพหรือตัดข้อความ');
  setProp_('BANNERS', json);
  return { success: true, message: 'บันทึกภาพหน้าปกแล้ว' };
}

/* ====================== API (doGet / doPost) ====================== */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}
// รายการฟังก์ชันที่เรียกจากหน้าเว็บได้ (w = เขียนข้อมูล จึงต้องล็อก)
const API = {
  getPublicConfig: { f: getPublicConfig }, getStudentScorecard: { f: getStudentScorecard }, findPendingWork: { f: findPendingWork },
  submitAssignment: { f: submitAssignment },
  teacherDashboard: { f: teacherDashboard }, getAssignmentSubmissions: { f: getAssignmentSubmissions }, getGradebook: { f: getGradebook },
  teacherChangePin: { f: teacherChangePin, w: 1 }, saveSubject: { f: saveSubject, w: 1 }, deleteSubject: { f: deleteSubject, w: 1 },
  saveAssignment: { f: saveAssignment, w: 1 }, deleteAssignment: { f: deleteAssignment, w: 1 }, saveAssignmentGrades: { f: saveAssignmentGrades, w: 1 },
  saveScoresBatch: { f: saveScoresBatch, w: 1 }, importStudentsFromCSV: { f: importStudentsFromCSV, w: 1 }, enrollStudents: { f: enrollStudents, w: 1 }, deleteStudent: { f: deleteStudent, w: 1 },
  adminLogin: { f: adminLogin }, adminSaveUser: { f: adminSaveUser, w: 1 }, adminDeleteUser: { f: adminDeleteUser, w: 1 }, adminSaveSettings: { f: adminSaveSettings, w: 1 },
  adminUploadLogo: { f: adminUploadLogo, w: 1 }, adminUploadBanner: { f: adminUploadBanner, w: 1 }, adminSaveBanners: { f: adminSaveBanners, w: 1 }
};
function dispatch_(fn, args) {
  const spec = Object.prototype.hasOwnProperty.call(API, fn) ? API[fn] : null;
  if (!spec) fail_('ไม่พบฟังก์ชันที่เรียก');
  args = Array.isArray(args) ? args : [];
  return spec.w ? withLock_(() => spec.f.apply(null, args)) : spec.f.apply(null, args);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return json_({ ok: true, name: 'score-system-api', version: '1.0.0' }); }
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    out = dispatch_(String(req.fn), req.args);
  } catch (err) {
    if (!(err instanceof UErr)) console.error(err && err.stack || err);
    out = { success: false, message: err instanceof UErr ? err.message : 'เกิดข้อผิดพลาดในระบบ กรุณาลองใหม่อีกครั้ง' };
  }
  return json_(out);
}
// รองรับกรณีเปิดหน้าเว็บผ่าน Apps Script เอง (google.script.run) — เรียกชื่อฟังก์ชันที่อนุญาตเท่านั้น
function apiCall(fn, args) { try { return dispatch_(String(fn), args); } catch (err) { return { success: false, message: err instanceof UErr ? err.message : 'เกิดข้อผิดพลาดในระบบ' }; } }

/* ====================== ติดตั้งครั้งแรก ====================== */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(T).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, T[n].h.length).setValues([T[n].h]); sh.setFrozenRows(1); }
    T[n].txt.forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
  });
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (def && def.getLastRow() <= 1 && ss.getSheets().length > 1) ss.deleteSheet(def);
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SALT')) p.setProperty('SALT', Utilities.getUuid() + Utilities.getUuid());
  if (!p.getProperty('ADMIN_PIN_HASH')) p.setProperty('ADMIN_PIN_HASH', hash_(DEFAULT_PIN));
  const defaults = { TITLE: 'ระบบเก็บคะแนนและส่งงานออนไลน์', TAGLINE: 'เรียนรู้ ส่งงาน ติดตามผล ได้ทุกที่ทุกเวลา', SCHOOL: 'ชื่อสถานศึกษา', SEMESTER: '1/2569' };
  Object.keys(defaults).forEach(k => { if (!p.getProperty(k)) p.setProperty(k, defaults[k]); });
  getRootFolder_();
  if (!all_('Teachers').length) put_('Teachers', r => false, ['T001', 'นาย', 'ครูตัวอย่าง', 'ผู้สอน', hash_(DEFAULT_PIN)]);
  Logger.log('ติดตั้งเสร็จแล้ว • แอดมิน PIN ' + DEFAULT_PIN + ' • ครูตัวอย่าง T001 PIN ' + DEFAULT_PIN + ' (ให้เปลี่ยนทันที) • โฟลเดอร์ไฟล์งาน ' + getRootFolder_().getUrl());
}

/** ข้อมูลทดลอง (ไม่บังคับ): นักเรียน 02900-02902, วิชา DEMO-101 ของครู T001, ชิ้นงาน 3 รายการ */
function seedDemo() {
  setup();
  [['02900', 'นาย', 'สมชาย', 'ใจดี'], ['02901', 'นางสาว', 'สมหญิง', 'รักเรียน'], ['02902', 'นาย', 'สมชาย', 'มั่นคง']].forEach((s, i) => {
    addStudentIfMissing_(s[0], s[1], s[2], s[3], 'ปวช.2/1', String(i + 5)); enroll_(s[0], 'DEMO-101');
  });
  put_('Subjects', r => String(r[0]) === 'DEMO-101', ['DEMO-101', 'วิชาตัวอย่าง', 'ปวช.2', prop_('SEMESTER') || '1/2569', 'T001']);
  const dl = fmt_(new Date(Date.now() + 7 * 864e5), "yyyy-MM-dd'T'23:59");
  put_('Assignments', r => r[0] === 'DEMO-HW01', ['DEMO-HW01', 'DEMO-101', 'ใบงานที่ 1 (ตัวอย่าง)', 'work', 10, dl, 'Active', 'ส่งไฟล์ใดก็ได้เพื่อทดลองระบบ']);
  put_('Assignments', r => r[0] === 'DEMO-T01', ['DEMO-T01', 'DEMO-101', 'ทดสอบย่อยครั้งที่ 1', 'test', 10, dl, 'Active', '']);
  put_('Assignments', r => r[0] === 'DEMO-C01', ['DEMO-C01', 'DEMO-101', 'คะแนนเก็บ: กิจกรรมในชั้นเรียน', 'collect', 10, '', 'Active', 'ครูลงคะแนนได้เลย']);
  Logger.log('สร้างข้อมูลทดลองแล้ว: ลองกรอกรหัสนักเรียน 02900');
}
