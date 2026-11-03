const { A, callGas, sheets, drive } = require('./gas-fake.js');
const call = (fn, ...args) => callGas(JSON.stringify({ fn, args }));
let fails = 0;
const t = (name, cond, extra) => { if (!cond) { fails++; console.log('FAIL', name, extra !== undefined ? JSON.stringify(extra) : ''); } else console.log('ok  ', name); };
A.setup(); A.seedDemo();
let r;
// สาธารณะ
r = call('getPublicConfig'); t('config', r.success && r.title && Array.isArray(r.banners), r);
r = call('getStudentScorecard', '02900'); t('scorecard (zero-padded id)', r.success && r.student.name.includes('สมชาย') && r.subjects.length === 1, r);
r = call('getStudentScorecard', '99999'); t('unknown student', !r.success);
r = call('getStudentScorecard', 'abc'); t('bad id rejected', !r.success);
r = call('findPendingWork', '02900'); t('pending by id', r.success && r.pending.length === 1 && r.pending[0].id === 'DEMO-HW01', r);
r = call('findPendingWork', 'สมชาย'); t('pending by name → 2 matches', r.success && r.matches.length === 2 && r.pending === null, r);
r = call('findPendingWork', 'ส'); t('1-char name rejected', !r.success);
// ส่งงาน
const png = Buffer.from('hello').toString('base64');
r = call('submitAssignment', { studentId: '02900', assignmentId: 'DEMO-HW01', files: [{ fileName: 'a.pdf', mimeType: 'application/pdf', fileData: png }] });
t('submit ok', r.success && r.status === 'ตรงเวลา' && r.files.length === 1, r);
r = call('findPendingWork', '02900'); t('pending now empty', r.success && r.pending.length === 0, r);
r = call('submitAssignment', { studentId: '02900', assignmentId: 'DEMO-HW01', files: [{ fileName: 'b.pdf', mimeType: 'application/pdf', fileData: png }] });
t('resubmit replaces (ungraded)', r.success && /แทนที่/.test(r.message), r);
r = call('submitAssignment', { studentId: '02900', assignmentId: 'DEMO-HW01', files: [{ fileName: 'x.exe', mimeType: 'a/b', fileData: png }] });
t('exe blocked', !r.success, r);
r = call('submitAssignment', { studentId: '02900', assignmentId: 'DEMO-C01', files: [{ fileName: 'a.pdf', fileData: png }] });
t('cannot submit collect', !r.success, r);
r = call('submitAssignment', { studentId: '02999', assignmentId: 'DEMO-HW01', files: [{ fileName: 'a.pdf', fileData: png }] });
t('unknown student cannot submit', !r.success, r);
// ครู: auth + throttle
r = call('teacherDashboard', 'T001:000000'); t('wrong pin', !r.success);
r = call('teacherDashboard', 'T001:123456'); t('teacher login default pin', r.success && r.subjects.length === 1 && r.counts.students === 3, r);
r = call('teacherDashboard', 'T001:12345678'); t('8-digit pin rejected', !r.success);
for (let i = 0; i < 5; i++) call('teacherDashboard', 'T001:111111');
r = call('teacherDashboard', 'T001:123456'); t('locked out after repeated fails', !r.success && /หลายครั้ง/.test(r.message), r);
require('./gas-fake.js').cache && Object.keys(require('./gas-fake.js').cache).forEach(k => delete require('./gas-fake.js').cache[k]);
const C = 'T001:123456';
r = call('teacherDashboard', C); t('unlock after cache clear', r.success);
// รายวิชา: เพิ่ม/แก้/ลบของตนเอง, ของผู้อื่นห้าม
r = call('saveSubject', C, { code: 'WEB-1', name: 'เว็บ', gradeLevel: 'ปวช.2', semester: '1/2569' }); t('add subject', r.success, r);
r = call('teacherDashboard', C); t('2 subjects', r.subjects.length === 2);
// ครูคนที่ 2
r = call('adminSaveUser', '123456', { role: 'teacher', id: 'T002', prefix: 'นาย', name: 'สอง', surname: 'ครู' }); t('admin adds teacher (default pin)', r.success, r);
const C2 = 'T002:123456';
r = call('teacherDashboard', C2); t('new teacher default pin works, sees none', r.success && r.subjects.length === 0, r);
r = call('saveSubject', C2, { code: 'WEB-1', name: 'แย่ง' }); t('cannot overwrite other teacher subject', !r.success, r);
r = call('deleteSubject', C2, 'WEB-1'); t('cannot delete other teacher subject', !r.success, r);
r = call('saveAssignment', C2, { id: 'X1', subjectCode: 'DEMO-101', title: 'x', category: 'collect', maxScore: 5, status: 'Active' }); t('cannot add assignment to other subject', !r.success, r);
// ลงทะเบียน
r = call('importStudentsFromCSV', C, 'รหัส,คำนำหน้า,ชื่อ,นามสกุล,ชั้น,เลขที่\n2903,นาย,สี่,ทดสอบ,ปวช.2/1,8\n02904,นางสาว,ห้า,ทดสอบ,ปวช.2/1,9\nabc,x,y', 'DEMO-101'); t('csv import (pads 2903→02903)', r.success && /ใหม่ 2/.test(r.message), r);
r = call('enrollStudents', C, 'WEB-1', '02900, 02901 99999'); t('enroll by ids (partial)', r.success && /2 คน/.test(r.message), r);
r = call('enrollStudents', C, 'WEB-1', '99999'); t('enroll unknown fails', !r.success, r);
// คะแนนเก็บ ลงได้เลย
r = call('saveAssignment', C, { id: 'DEMO-C02', subjectCode: 'DEMO-101', title: 'คะแนนเก็บ 2', category: 'collect', maxScore: 10, deadline: '', status: 'Active', description: '' }); t('add collect assignment (no deadline)', r.success, r);
r = call('saveAssignment', C, { id: 'DEMO-HW02', subjectCode: 'DEMO-101', title: 'ใบงาน 2', category: 'work', maxScore: 10, deadline: '', status: 'Active' }); t('work without deadline rejected', !r.success, r);
r = call('getAssignmentSubmissions', C, 'DEMO-C02'); t('collect roster lists all enrolled (5)', r.success && r.students.length === 5, r.students && r.students.length);
r = call('saveAssignmentGrades', C, 'DEMO-C02', [{ studentId: '02902', score: 7, feedback: 'ดี' }, { studentId: '02900', score: 10 }]); t('grade collect directly', r.success, r);
r = call('saveAssignmentGrades', C, 'DEMO-C02', [{ studentId: '02902', score: 11 }]); t('score over max rejected', !r.success, r);
r = call('saveAssignmentGrades', C, 'DEMO-HW01', [{ studentId: '02901', score: 9 }]); t('work without submission: silently skipped (0 saved)', r.success && /0 รายการ/.test(r.message), r);
r = call('saveAssignmentGrades', C, 'DEMO-HW01', [{ studentId: '02900', score: 8, feedback: 'ok' }]); t('grade submitted work', r.success && /1 รายการ/.test(r.message), r);
r = call('submitAssignment', { studentId: '02900', assignmentId: 'DEMO-HW01', files: [{ fileName: 'c.pdf', fileData: png }] }); t('cannot resubmit graded work', !r.success, r);
// สมุดคะแนน กลางภาค/ปลายภาค
r = call('saveScoresBatch', C, [{ studentId: '02900', subjectCode: 'DEMO-101', midterm: 20, finalExam: 30, special: '' }]); t('save midterm/final', r.success, r);
r = call('getGradebook', C, 'DEMO-101'); const row = r.rows.find(x => x.id === '02900');
t('gradebook total = work8 + collect10 + test0 + mid20 + final30 = 68 → grade 2.5', r.success && row.work === 8 && row.collect === 10 && row.midterm === 20 && row.finalExam === 30 && row.total === 68 && row.grade === '2.5', row);
t('gradebook has no legacy affective', !('affective' in row));
r = call('saveScoresBatch', C, [{ studentId: '02900', subjectCode: 'DEMO-101', midterm: 150 }]); t('midterm >100 rejected', !r.success, r);
r = call('saveScoresBatch', C, [{ studentId: '02900', subjectCode: 'DEMO-101', midterm: 20, finalExam: 30, special: 'ขร.' }]); t('special grade saved', r.success);
r = call('getStudentScorecard', '02900', 'DEMO-101'); t('scorecard: special overrides grade, items typed', r.card.grade === 'ขร.' && r.card.midterm === 20 && r.card.items.some(i => i.state === 'unrec' || i.state === 'graded'), r.card);
r = call('saveScoresBatch', C, [{ studentId: '02900', subjectCode: 'DEMO-101', midterm: 20, finalExam: 30, special: '' }]);
r = call('deleteStudent', C, '02904', 'DEMO-101'); t('remove from subject', r.success);
r = call('getGradebook', C, 'DEMO-101'); t('roster shrinks', r.rows.length === 4, r.rows.length);
// เปลี่ยน PIN ครู
r = call('teacherChangePin', C, '12345'); t('5-digit pin rejected', !r.success);
r = call('teacherChangePin', C, '654321'); t('change own pin', r.success);
r = call('teacherDashboard', C); t('old pin no longer works', !r.success);
r = call('teacherDashboard', 'T001:654321'); t('new pin works', r.success);
// แอดมิน
r = call('adminLogin', '000000'); t('admin wrong pin', !r.success);
r = call('adminLogin', '123456'); t('admin login', r.success && r.users.teachers.length === 2 && !JSON.stringify(r).includes('pinHash'), r.users && r.users.teachers.length);
r = call('adminSaveUser', '123456', { role: 'student', id: '12', name: 'x' }); t('bad student id rejected', !r.success);
r = call('adminSaveUser', '123456', { role: 'student', id: '02910', prefix: 'นาย', name: 'ใหม่', surname: 'ทดลอง', classRoom: 'ปวช.1', number: '1' }); t('admin adds student', r.success);
r = call('adminSaveBanners', '123456', [{ id: 'b1', url: 'https://example.com/a.jpg', title: 'ทดสอบ', subtitle: '' }]); t('banners saved', r.success);
r = call('adminSaveBanners', '123456', [{ id: 'b1', url: 'javascript:alert(1)', title: '' }]); t('non-https banner rejected', !r.success);
r = call('getPublicConfig'); t('public config returns banners', r.banners.length === 1);
r = call('adminUploadBanner', '123456', { mimeType: 'image/jpeg', fileData: png }); t('banner upload returns https url', r.success && /^https:\/\//.test(r.url), r);
r = call('adminUploadBanner', '123456', { mimeType: 'text/html', fileData: png }); t('non-image upload rejected', !r.success);
r = call('adminSaveSettings', '123456', { title: 'ระบบทดสอบ', tagline: 'คำโปรย', school: 'รร.', semester: '2/2569', newPin: '246810' }); t('admin settings + new 6-digit pin', r.success);
r = call('adminLogin', '123456'); t('old admin pin invalid', !r.success);
r = call('adminLogin', '246810'); t('new admin pin valid', r.success);
r = call('adminSaveSettings', '246810', { newPin: '12345678', title: 'x' }); t('8-digit admin pin rejected', !r.success);
r = call('adminDeleteUser', '246810', 'teacher', 'T001'); t('cannot delete teacher who owns subjects', !r.success, r);
r = call('adminDeleteUser', '246810', 'teacher', 'T002'); t('delete teacher without subjects', r.success);
// ความปลอดภัยของ dispatcher
r = call('setup'); t('setup() not callable via API', !r.success);
r = call('constructor'); t('prototype names not callable', !r.success);
r = callGas('not json'); t('garbage body handled', r.success === false);
// ครูลบวิชา → ข้อมูลที่เกี่ยวข้องหาย
r = call('deleteSubject', 'T001:654321', 'WEB-1'); t('owner deletes subject', r.success);
console.log(fails ? '\n' + fails + ' FAILED' : '\nALL PASSED');
process.exit(fails ? 1 : 0);
