// app.js - منطق كل الصفحات + Excel + Preview Modal + شهادات (نسخة كاملة نهائية)

// ============================ Utils ============================
const $ = (id) => document.getElementById(id);

function esc(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

function termLabel(term) {
  const map = {
    'term1': 'الترم الأول',
    'term2': 'الترم الثاني',
    'final': 'النهائي',
    'term-1': 'الترم الأول',
    'term-2': 'الترم الثاني',
    'final-term': 'النهائي'
  };
  return map[term] || term || '';
}

// ⭐⭐⭐ دالة إخفاء الأزرار غير المسموحة
function hideForNonAdmin() {
  // روابط لوحة التحكم والشهادات
  document.querySelectorAll('a[href="admin.html"], a[href="index.html"]').forEach(el => {
    el.style.display = 'none';
  });
  // زر طباعة الكل (في dashboard)
  const printBtn = $('print-btn');
  if (printBtn) printBtn.style.display = 'none';
}

// ============================ SweetAlert2 ============================
const Swal2 = window.Swal;

const Toast = Swal2.mixin({
  toast: true,
  position: 'top-start',
  showConfirmButton: false,
  timer: 3200,
  timerProgressBar: true,
  customClass: { popup: 'swal-rtl' },
  didOpen: (t) => {
    t.addEventListener('mouseenter', Swal2.stopTimer);
    t.addEventListener('mouseleave', Swal2.resumeTimer);
  }
});

function toast(msg, type = 'info', ms = 3200) {
  const icon = ['success','error','warning','info','question'].includes(type) ? type : 'info';
  Toast.fire({ icon, title: msg, timer: ms });
}

async function confirmDialog(title, text = 'لا يمكن التراجع عن هذا الإجراء!', confirmText = 'نعم، احذف') {
  const r = await Swal2.fire({
    title, text, icon: 'warning',
    showCancelButton: true,
    confirmButtonText: confirmText,
    cancelButtonText: 'إلغاء',
    confirmButtonColor: '#ef4444',
    cancelButtonColor: '#64748b',
    reverseButtons: true,
    focusCancel: true,
    customClass: { popup: 'swal-rtl' }
  });
  return r.isConfirmed;
}

function loadingDialog(title = 'جاري المعالجة...', text = '') {
  if (Swal2.isVisible()) Swal2.close();
  Swal2.fire({
    title, html: text,
    allowOutsideClick: false,
    allowEscapeKey: false,
    showConfirmButton: false,
    didOpen: () => Swal2.showLoading()
  });
}

function closeDialog() {
  return new Promise(resolve => {
    if (!Swal2.isVisible()) { resolve(); return; }
    Swal2.close();
    setTimeout(resolve, 80);
  });
}

async function showImportResult(inserted, skipped, errors, entityName) {
  let html = `<div style="text-align:right; font-size:15px; line-height:1.9;">
    <p>✅ تم إضافة <b style="color:#16a34a;">${inserted}</b> ${entityName} جديد</p>`;
  if (skipped) html += `<p>⚠️ تم تخطي <b style="color:#f59e0b;">${skipped}</b> صف</p>`;
  if (errors && errors.length) {
    html += `<hr style="margin:12px 0;">
      <p><b>تفاصيل التحذيرات:</b></p>
      <div style="max-height:220px; overflow:auto; text-align:right; font-size:13px;
                  background:#fef3c7; padding:10px; border-radius:8px;">
        ${errors.slice(0, 40).map(e => '• ' + esc(e)).join('<br>')}
        ${errors.length > 40 ? `<br><i>...و ${errors.length - 40} تحذير آخر</i>` : ''}
      </div>`;
  }
  html += '</div>';

  return Swal2.fire({
    title: inserted ? '🎉 نتيجة الاستيراد' : 'لم يتم استيراد أي صف',
    html,
    icon: errors?.length ? 'warning' : (inserted ? 'success' : 'info'),
    confirmButtonText: 'تم',
    confirmButtonColor: '#1e5fa8',
    customClass: { popup: 'swal-rtl' }
  });
}

// ============================ Safe Print ============================
function safePrint(htmlContent) {
  if (!window._isAdmin) {
    return toast('غير مصرح بالطباعة', 'error');
  }
  if (Swal2.isVisible()) Swal2.close();
  const area = $('print-area');
  if (!area) return;
  area.innerHTML = htmlContent;
  setTimeout(() => {
    try { window.print(); }
    catch (e) { console.error('Print error:', e); toast('خطأ أثناء الطباعة', 'error'); }
    setTimeout(() => { area.innerHTML = ''; }, 500);
  }, 150);
}
window.addEventListener('afterprint', () => {
  const area = $('print-area');
  if (area) area.innerHTML = '';
});

// ============================ fillSelect ============================
function fillSelect(sel, items, { valueKey='id', labelKey='name', placeholder='اختر...', keep=false } = {}) {
  if (!sel) return;
  const cur = keep ? sel.value : null;
  sel.innerHTML = `<option value="">${esc(placeholder)}</option>`;
  items.forEach(it => {
    const o = document.createElement('option');
    o.value = it[valueKey];
    o.textContent = it[labelKey];
    sel.appendChild(o);
  });
  if (cur) sel.value = cur;
}

async function getMySubjects() {
  const p = await getCurrentProfile();
  if (!p) return [];
  if (p.role === 'admin' || p.role === 'employee') {
    const { data } = await _supabase.from('subjects').select('*').order('name');
    return data || [];
  }
  const { data } = await _supabase
    .from('teacher_subjects')
    .select('subject_id, subjects(id, name, max_score, absence_max_score)')
    .eq('teacher_id', p.id);
  return (data || []).map(r => r.subjects).filter(Boolean);
}

// ============================ Excel Helpers ============================
function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
        resolve(json);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function downloadTemplate(filename, headers, sampleRow) {
  const ws = XLSX.utils.aoa_to_sheet([headers, sampleRow]);
  ws['!cols'] = headers.map(() => ({ wch: 22 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Template');
  XLSX.writeFile(wb, filename);
  toast('تم تحميل القالب ✅', 'success');
}

function downloadIdReference(filename, headers, rows) {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols'] = headers.map(() => ({ wch: 25 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'IDs');
  XLSX.writeFile(wb, filename);
  toast('تم تحميل مرجع الأرقام ✅', 'success');
}

// ============================ IMPORT PREVIEW ============================
let _importPreview = { type: null, rows: [] };

function openImportPreview(type, rows) {
  _importPreview.type = type;
  _importPreview.rows = rows.map(r => ({ data: { ...r } }));
  renderImportPreview();
  const modal = $('import-modal');
  if (modal) modal.style.display = 'flex';
}

function closeImportModal() {
  const modal = $('import-modal');
  if (modal) modal.style.display = 'none';
  _importPreview = { type: null, rows: [] };
}
window.closeImportModal = closeImportModal;

const IMPORT_LABELS = {
  stages: 'المراحل', grades: 'الصفوف', classes: 'الفصول',
  subjects: 'المواد', teachers: 'المدرسين', students: 'الطلاب'
};
const IMPORT_ENTITY = {
  stages: 'مرحلة', grades: 'صف', classes: 'فصل',
  subjects: 'مادة', teachers: 'مدرس', students: 'طالب'
};

function getImportColumns(type) {
  switch (type) {
    case 'stages':
      return [
        { key: 'name', label: 'اسم المرحلة', type: 'text' },
        { key: 'sort_order', label: 'الترتيب', type: 'number' }
      ];
    case 'grades':
      return [
        { key: 'stage_id', label: 'المرحلة', type: 'select', source: 'stages',
          getLabel: (s) => `${s.name} (#${s.id})` },
        { key: 'name', label: 'اسم الصف', type: 'text' },
        { key: 'sort_order', label: 'الترتيب', type: 'number' }
      ];
    case 'classes':
      return [
        { key: 'grade_level_id', label: 'الصف', type: 'select', source: 'grades',
          getLabel: (g) => {
            const s = adminState.stages.find(x => String(x.id) === String(g.stage_id));
            return `${s ? s.name + ' / ' : ''}${g.name} (#${g.id})`;
          }},
        { key: 'name', label: 'اسم الفصل', type: 'text' }
      ];
    case 'subjects':
      return [
        { key: 'name', label: 'اسم المادة', type: 'text' },
        { key: 'max_score', label: 'الدرجة القصوى للاختبار', type: 'number' },
        { key: 'absence_max_score', label: 'الدرجة القصوى للغياب', type: 'number' }
      ];
    case 'teachers':
      return [
        { key: 'full_name', label: 'الاسم', type: 'text' },
        { key: 'email', label: 'البريد', type: 'text' },
        { key: 'password', label: 'كلمة المرور', type: 'text' },
        { key: 'role', label: 'الدور', type: 'select-fixed',
          options: [
            { value: 'teacher', label: 'مدرس' },
            { value: 'employee', label: 'موظف' },
            { value: 'admin', label: 'مدير' }
          ]}
      ];
    case 'students':
      return [
        { key: 'full_name', label: 'اسم الطالب', type: 'text' },
        { key: 'class_id', label: 'الفصل', type: 'select', source: 'classes',
          getLabel: (c) => {
            const g = adminState.grades.find(x => String(x.id) === String(c.grade_level_id));
            const s = g ? adminState.stages.find(x => String(x.id) === String(g.stage_id)) : null;
            return `${s ? s.name + ' / ' : ''}${g ? g.name + ' / ' : ''}${c.name} (#${c.id})`;
          }},
        { key: 'national_id', label: 'رقم الهوية', type: 'text' }
      ];
  }
  return [];
}

function renderImportPreview() {
  const t = _importPreview.type;
  const rows = _importPreview.rows;
  if (!t) return;

  $('import-modal-title').textContent = `معاينة استيراد ${IMPORT_LABELS[t]} (${rows.length} صف)`;

  const cols = getImportColumns(t);
  const thead = $('import-modal-thead');
  const tbody = $('import-modal-tbody');

  thead.innerHTML = '<tr>' + cols.map(c => `<th>${c.label}</th>`).join('') + '<th>حذف</th></tr>';
  tbody.innerHTML = rows.map((row, i) => renderImportRow(t, row, i, cols)).join('');

  bindPreviewEvents();
}

function renderImportRow(type, row, i, cols) {
  const cells = cols.map(col => {
    const val = row.data[col.key] ?? '';
    if (col.type === 'text') {
      return `<td><input type="text" class="form-control input-sm" data-row="${i}" data-key="${col.key}" value="${esc(val)}"></td>`;
    }
    if (col.type === 'number') {
      return `<td><input type="number" class="form-control input-sm" data-row="${i}" data-key="${col.key}" value="${esc(val)}"></td>`;
    }
    if (col.type === 'select-fixed') {
      const opts = col.options.map(o =>
        `<option value="${esc(o.value)}" ${String(val) === String(o.value) ? 'selected' : ''}>${esc(o.label)}</option>`
      ).join('');
      return `<td><select class="form-control input-sm" data-row="${i}" data-key="${col.key}">
        <option value="">— اختر —</option>${opts}</select></td>`;
    }
    if (col.type === 'select') {
      const source = adminState[col.source] || [];
      const hasMatch = source.some(s => String(s.id) === String(val));
      const opts = source.map(s =>
        `<option value="${s.id}" ${String(val) === String(s.id) ? 'selected' : ''}>${esc(col.getLabel(s))}</option>`
      ).join('');
      const invalidMark = (val && !hasMatch) ? `⚠️ #${val} غير موجود — ` : '';
      return `<td><select class="form-control input-sm ${val && !hasMatch ? 'input-error' : ''}" data-row="${i}" data-key="${col.key}">
        <option value="">${invalidMark}— اختر —</option>${opts}</select></td>`;
    }
    return '<td></td>';
  }).join('');

  return `<tr data-row="${i}">${cells}<td>
    <button class="btn btn-sm btn-danger" onclick="window.__removeImportRow(${i})" title="حذف هذا الصف">×</button>
  </td></tr>`;
}

window.__removeImportRow = (i) => {
  _importPreview.rows.splice(i, 1);
  renderImportPreview();
};

function bindPreviewEvents() {
  document.querySelectorAll('#import-modal-tbody input, #import-modal-tbody select').forEach(el => {
    el.addEventListener('input', (e) => {
      const rowIdx = parseInt(e.target.dataset.row);
      const key = e.target.dataset.key;
      _importPreview.rows[rowIdx].data[key] = e.target.value;
    });
    el.addEventListener('change', (e) => {
      const rowIdx = parseInt(e.target.dataset.row);
      const key = e.target.dataset.key;
      _importPreview.rows[rowIdx].data[key] = e.target.value;
      if (e.target.tagName === 'SELECT' && e.target.value) {
        e.target.classList.remove('input-error');
      }
    });
  });
}

async function confirmImport() {
  const type = _importPreview.type;
  const rows = _importPreview.rows.map(r => r.data);
  if (!rows.length) return toast('لا توجد صفوف للحفظ', 'warning');

  loadingDialog('جاري حفظ البيانات...');
  try {
    const res = await performImport(type, rows);
    await closeDialog();
    closeImportModal();
    await showImportResult(res.inserted, res.skipped, res.errors, IMPORT_ENTITY[type]);
    await loadAllAdminData();
  } catch (e) {
    await closeDialog();
    Swal2.fire({ icon: 'error', title: 'فشل الحفظ', text: e.message, customClass: { popup: 'swal-rtl' } });
  }
}
window.confirmImport = confirmImport;

async function performImport(type, rows) {
  switch (type) {
    case 'stages':   return await importStages(rows);
    case 'grades':   return await importGradeLevels(rows);
    case 'classes':  return await importClasses(rows);
    case 'subjects': return await importSubjects(rows);
    case 'teachers': return await importTeachers(rows);
    case 'students': return await importStudents(rows);
  }
  return { inserted: 0, skipped: 0, errors: ['نوع غير معروف'] };
}

// ---- Import functions ----
async function importStages(rows) {
  const existingNames = new Set(adminState.stages.map(s => s.name.trim()));
  const toInsert = []; const errors = [];
  rows.forEach((r, i) => {
    const name = String(r.name || '').trim();
    if (!name) { errors.push(`صف ${i + 1}: الاسم فارغ`); return; }
    if (existingNames.has(name)) return;
    toInsert.push({ name, sort_order: parseInt(r.sort_order) || 0 });
  });
  if (!toInsert.length) return { inserted: 0, skipped: rows.length, errors };
  const { error } = await _supabase.from('stages').insert(toInsert);
  if (error) throw error;
  return { inserted: toInsert.length, skipped: rows.length - toInsert.length, errors };
}

async function importGradeLevels(rows) {
  const validStageIds = new Set(adminState.stages.map(s => String(s.id)));
  const existingKeys = new Set(adminState.grades.map(g => `${g.stage_id}::${g.name.trim()}`));
  const toInsert = []; const errors = [];
  rows.forEach((r, i) => {
    const stage_id = String(r.stage_id || '').trim();
    const name = String(r.name || '').trim();
    if (!stage_id || !name) { errors.push(`صف ${i + 1}: بيانات ناقصة`); return; }
    if (!validStageIds.has(stage_id)) { errors.push(`صف ${i + 1}: المرحلة #${stage_id} غير موجودة`); return; }
    if (existingKeys.has(`${stage_id}::${name}`)) return;
    toInsert.push({ stage_id: parseInt(stage_id), name, sort_order: parseInt(r.sort_order) || 0 });
  });
  if (!toInsert.length) return { inserted: 0, skipped: rows.length, errors };
  const { error } = await _supabase.from('grade_levels').insert(toInsert);
  if (error) throw error;
  return { inserted: toInsert.length, skipped: rows.length - toInsert.length, errors };
}

async function importClasses(rows) {
  const validGradeIds = new Set(adminState.grades.map(g => String(g.id)));
  const existingKeys = new Set(adminState.classes.map(c => `${c.grade_level_id}::${c.name.trim()}`));
  const toInsert = []; const errors = [];
  rows.forEach((r, i) => {
    const grade_level_id = String(r.grade_level_id || '').trim();
    const name = String(r.name || '').trim();
    if (!grade_level_id || !name) { errors.push(`صف ${i + 1}: بيانات ناقصة`); return; }
    if (!validGradeIds.has(grade_level_id)) { errors.push(`صف ${i + 1}: الصف #${grade_level_id} غير موجود`); return; }
    if (existingKeys.has(`${grade_level_id}::${name}`)) return;
    toInsert.push({ grade_level_id: parseInt(grade_level_id), name });
  });
  if (!toInsert.length) return { inserted: 0, skipped: rows.length, errors };
  const { error } = await _supabase.from('classes').insert(toInsert);
  if (error) throw error;
  return { inserted: toInsert.length, skipped: rows.length - toInsert.length, errors };
}

async function importSubjects(rows) {
  const existingNames = new Set(adminState.subjects.map(s => s.name.trim()));
  const toInsert = []; const errors = [];
  rows.forEach((r, i) => {
    const name = String(r.name || '').trim();
    if (!name) { errors.push(`صف ${i + 1}: الاسم فارغ`); return; }
    if (existingNames.has(name)) return;
    toInsert.push({
      name,
      max_score: parseFloat(r.max_score) || 100,
      absence_max_score: parseFloat(r.absence_max_score) || 0
    });
  });
  if (!toInsert.length) return { inserted: 0, skipped: rows.length, errors };
  const { error } = await _supabase.from('subjects').insert(toInsert);
  if (error) throw error;
  return { inserted: toInsert.length, skipped: rows.length - toInsert.length, errors };
}

async function importTeachers(rows) {
  let inserted = 0, skipped = 0;
  const errors = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const full_name = String(r.full_name || '').trim();
    const email = String(r.email || '').trim();
    const password = String(r.password || '').trim() || 'ChangeMe@123';
    const role = (String(r.role || '').trim() || 'teacher').toLowerCase();
    if (!full_name || !email) { errors.push(`صف ${i + 1}: اسم أو بريد ناقص`); skipped++; continue; }
    if (!['admin','employee','teacher'].includes(role)) { errors.push(`صف ${i + 1}: دور غير صحيح`); skipped++; continue; }

    const { error } = await _signupClient.auth.signUp({
      email, password,
      options: { data: { full_name, role } }
    });
    if (error) { errors.push(`صف ${i + 1}: ${error.message}`); skipped++; continue; }
    inserted++;
    await new Promise(r => setTimeout(r, 250));
  }
  return { inserted, skipped, errors };
}

async function importStudents(rows) {
  const validClassIds = new Set(adminState.classes.map(c => String(c.id)));
  const prof = await getCurrentProfile();
  const toInsert = []; const errors = [];
  rows.forEach((r, i) => {
    const full_name = String(r.full_name || '').trim();
    const class_id = String(r.class_id || '').trim();
    if (!full_name) { errors.push(`صف ${i + 1}: الاسم ناقص`); return; }
    if (!class_id) { errors.push(`صف ${i + 1}: رقم الفصل ناقص`); return; }
    if (!validClassIds.has(class_id)) { errors.push(`صف ${i + 1}: الفصل #${class_id} غير موجود`); return; }
    toInsert.push({
      full_name,
      class_id: parseInt(class_id),
      national_id: String(r.national_id || '').trim() || null,
      created_by: prof?.id
    });
  });
  if (!toInsert.length) return { inserted: 0, skipped: rows.length, errors };
  const { error } = await _supabase.from('students').insert(toInsert);
  if (error) throw error;
  return { inserted: toInsert.length, skipped: rows.length - toInsert.length, errors };
}

// ============================ LOGIN PAGE ============================
async function initLoginPage() {
  const form = $('login-form');
  if (!form) return;
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const email = $('email').value.trim();
    const password = $('password').value;
    const err = $('error-msg');
    err.textContent = '';

    loadingDialog('جاري تسجيل الدخول...', '');
    const { error } = await _supabase.auth.signInWithPassword({ email, password });
    await closeDialog();

    if (error) {
      err.textContent = 'خطأ: ' + error.message;
      Swal2.fire({ icon: 'error', title: 'فشل تسجيل الدخول', text: error.message, confirmButtonText: 'حسناً', customClass: { popup: 'swal-rtl' } });
      return;
    }
    const p = await getCurrentProfile();
    if (!p) { err.textContent = 'لم يتم العثور على الملف الشخصي'; return; }

    toast(`مرحباً ${p.full_name} 👋`, 'success', 2000);
    setTimeout(() => {
      window.location.href = (p.role === 'teacher') ? 'dashboard.html' : 'admin.html';
    }, 600);
  });
}

// ============================ ADMIN PAGE ============================
const adminState = { stages: [], grades: [], classes: [], subjects: [], teachers: [], students: [] };

async function initAdminPage() {
  const profile = await requireRole(['admin','employee']);
  if (!profile) return;
  $('user-name').textContent = profile.full_name || '';

  window._isAdmin = profile.role === 'admin';

  // ⭐ الموظف: إخفاء زر "الشهادات"
  if (!window._isAdmin) {
    document.querySelectorAll('a[href="index.html"]').forEach(el => el.style.display = 'none');
  }

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  switchTab('classes');
  await loadAllAdminData();
  bindAdminForms();
  bindTeacherCreation();
  bindTeacherAssignments();
  bindStudentForm();
  bindGradesView();
  bindExcelImports();
}

function switchTab(name) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.dataset.panel === name));
  if (name === 'grades') loadGradesView();
}

async function loadAllAdminData() {
  const [stages, grades, classes, subjects, profiles, students] = await Promise.all([
    _supabase.from('stages').select('*').order('sort_order'),
    _supabase.from('grade_levels').select('*').order('sort_order'),
    _supabase.from('classes').select('*').order('name'),
    _supabase.from('subjects').select('*').order('name'),
    _supabase.from('profiles').select('*').order('full_name'),
    _supabase.from('students').select('*').order('full_name'),
  ]);
  adminState.stages   = stages.data   || [];
  adminState.grades   = grades.data   || [];
  adminState.classes  = classes.data  || [];
  adminState.subjects = subjects.data || [];
  adminState.teachers = (profiles.data || []).filter(p => p.role === 'teacher' || p.role === 'employee');
  adminState.students = students.data || [];

  renderStagesTree();
  renderSubjectsList();
  renderTeachersList();
  renderStudentsList();
  fillAdminSelects();
}

function fillAdminSelects() {
  const gradesWithPath = adminState.grades.map(g => {
    const s = adminState.stages.find(x => String(x.id) === String(g.stage_id));
    return { id: g.id, name: `${s ? s.name + ' / ' : ''}${g.name} (#${g.id})` };
  });
  fillSelect($('gl-stage'), gradesWithPath, { placeholder: 'اختر الصف' });
  fillSelect($('stage-select-for-grade'), adminState.stages, { placeholder: 'اختر المرحلة' });

  const classesWithPath = adminState.classes.map(c => {
    const g = adminState.grades.find(x => String(x.id) === String(c.grade_level_id));
    const s = g ? adminState.stages.find(x => String(x.id) === String(g.stage_id)) : null;
    return {
      id: c.id,
      name: `${s ? s.name + ' / ' : ''}${g ? g.name + ' / ' : ''}${c.name} (#${c.id})`
    };
  });
  fillSelect($('student-class'), classesWithPath, { placeholder: 'اختر الفصل' });

  fillSelect($('assign-subject'), adminState.subjects, { placeholder: 'اختر المادة' });
  fillSelect($('assign-teacher'),
    adminState.teachers.map(t => ({ id: t.id, name: t.full_name || t.email })),
    { placeholder: 'اختر المدرس' });

  fillSelect($('gv-class'), classesWithPath, { placeholder: 'كل الفصول' });
  fillSelect($('gv-subject'), adminState.subjects, { placeholder: 'كل المواد' });
}

// -------- Stages / Grades / Classes Tree --------
function renderStagesTree() {
  const root = $('stages-tree');
  if (!root) return;
  root.innerHTML = '';
  if (!adminState.stages.length) {
    root.innerHTML = '<p class="muted">لا توجد مراحل بعد.</p>'; return;
  }
  adminState.stages.forEach(st => {
    const grades = adminState.grades.filter(g => String(g.stage_id) === String(st.id));
    const div = document.createElement('div');
    div.className = 'tree-node';
    div.innerHTML = `
      <div class="tree-head">
        <strong>${esc(st.name)} <small class="muted">#${st.id}</small></strong>
        <div style="display:flex; gap:4px;">
          <button class="btn btn-sm btn-warning" onclick="window.__editStage('${st.id}')">تعديل</button>
          <button class="btn btn-sm btn-danger" data-del-stage="${st.id}">حذف</button>
        </div>
      </div>
      <div class="tree-children">
        ${grades.map(g => {
          const cls = adminState.classes.filter(c => String(c.grade_level_id) === String(g.id));
          return `
            <div class="tree-sub">
              <div class="tree-head">
                <span>${esc(g.name)} <small class="muted">#${g.id}</small></span>
                <div style="display:flex; gap:4px;">
                  <button class="btn btn-sm btn-warning" onclick="window.__editGrade('${g.id}')">تعديل</button>
                  <button class="btn btn-sm btn-danger" data-del-grade="${g.id}">حذف</button>
                </div>
              </div>
              <div class="chips">
                ${cls.map(c => `
                  <span class="chip">
                    ${esc(c.name)} <small style="opacity:.7;">#${c.id}</small>
                    <button class="chip-x" onclick="window.__editClass('${c.id}')" title="تعديل" style="color:#f59e0b;">✎</button>
                    <button class="chip-x" data-del-class="${c.id}" title="حذف">×</button>
                  </span>
                `).join('') || '<span class="muted">لا فصول</span>'}
              </div>
            </div>
          `;
        }).join('') || '<p class="muted">لا صفوف</p>'}
      </div>
    `;
    root.appendChild(div);
  });
  root.querySelectorAll('[data-del-stage]').forEach(b => b.onclick = () => delRow('stages', b.dataset.delStage, 'المرحلة'));
  root.querySelectorAll('[data-del-grade]').forEach(b => b.onclick = () => delRow('grade_levels', b.dataset.delGrade, 'الصف'));
  root.querySelectorAll('[data-del-class]').forEach(b => b.onclick = () => delRow('classes', b.dataset.delClass, 'الفصل'));
}

async function delRow(table, id, label = 'العنصر') {
  const ok = await confirmDialog(`حذف ${label}؟`, `سيتم حذف هذا الـ${label} وكل ما يتبعه. لا يمكن التراجع!`);
  if (!ok) return;
  loadingDialog('جاري الحذف...', '');
  const { error } = await _supabase.from(table).delete().eq('id', id);
  await closeDialog();
  if (error) {
    Swal2.fire({ icon: 'error', title: 'فشل الحذف', text: error.message, customClass: { popup: 'swal-rtl' } });
    return;
  }
  toast('تم الحذف بنجاح ✅', 'success');
  await loadAllAdminData();
}

// -------- Admin Forms --------
function bindAdminForms() {
  $('stage-form')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = $('stage-name').value.trim();
    if (!name) return toast('اكتب اسم المرحلة', 'warning');
    const { error } = await _supabase.from('stages').insert({ name });
    if (error) return toast(error.message, 'error');
    $('stage-name').value = '';
    toast('تمت إضافة المرحلة ✅', 'success');
    await loadAllAdminData();
  });

  $('grade-form')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const stage_id = $('stage-select-for-grade').value;
    const name = $('grade-name').value.trim();
    if (!stage_id || !name) return toast('اختر المرحلة واكتب اسم الصف', 'warning');
    const { error } = await _supabase.from('grade_levels').insert({ stage_id, name });
    if (error) return toast(error.message, 'error');
    $('grade-name').value = '';
    toast('تمت إضافة الصف ✅', 'success');
    await loadAllAdminData();
  });

  $('class-form')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const grade_level_id = $('gl-stage').value;
    const name = $('class-name').value.trim();
    if (!grade_level_id || !name) return toast('اختر الصف واكتب اسم الفصل', 'warning');
    const { error } = await _supabase.from('classes').insert({ grade_level_id, name });
    if (error) return toast(error.message, 'error');
    $('class-name').value = '';
    toast('تمت إضافة الفصل ✅', 'success');
    await loadAllAdminData();
  });

  $('subject-form')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const name = $('subject-name').value.trim();
    const max_score = parseFloat($('subject-max').value) || 100;
    const absence_max_score = parseFloat($('subject-absence-max').value) || 0;
    if (!name) return toast('اكتب اسم المادة', 'warning');
    const { error } = await _supabase.from('subjects').insert({ name, max_score, absence_max_score });
    if (error) return toast(error.message, 'error');
    $('subject-name').value = '';
    $('subject-max').value = 100;
    $('subject-absence-max').value = 5;
    toast('تمت إضافة المادة ✅', 'success');
    await loadAllAdminData();
  });
}

// -------- Subjects List --------
function renderSubjectsList() {
  const root = $('subjects-list');
  if (!root) return;
  if (!adminState.subjects.length) { root.innerHTML = '<p class="muted">لا مواد بعد.</p>'; return; }
  root.innerHTML = adminState.subjects.map(s => `
    <div class="row-item">
      <span>${esc(s.name)} <small class="muted">(اختبار: ${s.max_score} / غياب: ${s.absence_max_score ?? 0}) #${s.id}</small></span>
      <div style="display:flex; gap:4px;">
        <button class="btn btn-sm btn-warning" onclick="window.__editSubject('${s.id}')">تعديل</button>
        <button class="btn btn-sm btn-danger" onclick="window.__delSubject('${s.id}')">حذف</button>
      </div>
    </div>
  `).join('');
}
window.__delSubject = async (id) => { await delRow('subjects', id, 'المادة'); };

// -------- Teachers List --------
function renderTeachersList() {
  const root = $('teachers-list');
  if (!root) return;
  if (!adminState.teachers.length) { root.innerHTML = '<p class="muted">لا مدرسين بعد.</p>'; return; }
  root.innerHTML = adminState.teachers.map(t => `
    <div class="row-item">
      <span>${esc(t.full_name)} <small class="muted">${esc(t.email || '')} — ${esc(t.role)}</small></span>
      <div style="display:flex; gap:4px;">
        <button class="btn btn-sm btn-warning" onclick="window.__editTeacher('${t.id}')">تعديل</button>
        <button class="btn btn-sm btn-danger" onclick="window.__delTeacher('${t.id}')">حذف</button>
      </div>
    </div>
  `).join('');
}
window.__delTeacher = async (id) => {
  const ok = await confirmDialog('حذف المدرس؟', 'سيتم حذفه من النظام (لن يُحذف من Auth). متأكد؟', 'نعم، احذف');
  if (!ok) return;
  loadingDialog('جاري الحذف...', '');
  await _supabase.from('profiles').delete().eq('id', id);
  await closeDialog();
  toast('تم الحذف ✅', 'success');
  await loadAllAdminData();
};

function bindTeacherCreation() {
  const form = $('teacher-form');
  if (!form) return;
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const full_name = $('teacher-name').value.trim();
    const email = $('teacher-email').value.trim();
    const password = $('teacher-password').value || 'ChangeMe@123';
    const role = $('teacher-role').value || 'teacher';
    if (!full_name || !email) return toast('اكتب الاسم والإيميل', 'warning');

    loadingDialog('جاري إنشاء الحساب...', 'قد يأخذ ثانية');
    const { data, error } = await _signupClient.auth.signUp({
      email, password,
      options: { data: { full_name, role } }
    });
    await closeDialog();

    if (error) {
      Swal2.fire({ icon: 'error', title: 'فشل الإنشاء', text: error.message, customClass: { popup: 'swal-rtl' } });
      return;
    }
    if (!data.user) return toast('فشل إنشاء الحساب', 'error');

    Swal2.fire({
      icon: 'success',
      title: 'تم إنشاء الحساب 🎉',
      html: `<div style="text-align:right;">
        <p><b>الاسم:</b> ${esc(full_name)}</p>
        <p><b>الإيميل:</b> ${esc(email)}</p>
        <p><b>كلمة المرور:</b> <code>${esc(password)}</code></p>
      </div>`,
      confirmButtonText: 'تم',
      customClass: { popup: 'swal-rtl' }
    });
    form.reset();
    $('teacher-password').value = '';
    await loadAllAdminData();
  });
}

// -------- Teacher Assignments --------
function bindTeacherAssignments() {
  const form = $('assign-form');
  if (!form) return;
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const teacher_id = $('assign-teacher').value;
    const subject_id = $('assign-subject').value;
    if (!teacher_id || !subject_id) return toast('اختر المدرس والمادة', 'warning');
    const { error } = await _supabase.from('teacher_subjects').insert({ teacher_id, subject_id });
    if (error) return toast(error.message, 'error');
    toast('تم الإسناد ✅', 'success');
    renderAssignments();
  });
  renderAssignments();
}

async function renderAssignments() {
  const root = $('assign-list');
  if (!root) return;
  const { data } = await _supabase.from('teacher_subjects')
    .select('teacher_id, subject_id, profiles(full_name, email), subjects(name)');
  if (!data || !data.length) { root.innerHTML = '<p class="muted">لا إسنادات بعد.</p>'; return; }
  root.innerHTML = data.map(r => `
    <div class="row-item">
      <span>${esc(r.profiles?.full_name || r.profiles?.email || '')} → ${esc(r.subjects?.name || '')}</span>
      <button class="btn btn-sm btn-danger" onclick="window.__delAssign('${r.teacher_id}','${r.subject_id}')">فك</button>
    </div>
  `).join('');
}
window.__delAssign = async (t, s) => {
  const ok = await confirmDialog('فك الإسناد؟', 'سيتم فصل المدرس عن المادة.', 'نعم، افصل');
  if (!ok) return;
  await _supabase.from('teacher_subjects').delete().eq('teacher_id', t).eq('subject_id', s);
  toast('تم فك الإسناد ✅', 'success');
  renderAssignments();
};

// -------- Students List --------
function renderStudentsList() {
  const root = $('students-list');
  if (!root) return;
  if (!adminState.students.length) { root.innerHTML = '<p class="muted">لا طلاب بعد.</p>'; return; }
  const classMap = Object.fromEntries(adminState.classes.map(c => [String(c.id), c]));
  root.innerHTML = `
    <table class="data-table">
      <thead><tr><th>#</th><th>الاسم</th><th>الفصل</th><th>إجراء</th></tr></thead>
      <tbody>
        ${adminState.students.map(s => `
          <tr>
            <td>${s.id}</td>
            <td>${esc(s.full_name)}</td>
            <td>${esc(classMap[String(s.class_id)]?.name || '')} <small class="muted">#${s.class_id}</small></td>
            <td style="white-space:nowrap;">
              <button class="btn btn-sm btn-warning" onclick="window.__editStudent('${s.id}')">تعديل</button>
              <button class="btn btn-sm btn-danger" onclick="window.__delStudent('${s.id}')">حذف</button>
            </td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}
window.__delStudent = async (id) => { await delRow('students', id, 'الطالب'); };

function bindStudentForm() {
  const form = $('student-form');
  if (!form) return;
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const full_name = $('student-name').value.trim();
    const class_id = $('student-class').value;
    const national_id = $('student-nid').value.trim() || null;
    if (!full_name || !class_id) return toast('اكتب الاسم واختر الفصل', 'warning');
    const prof = await getCurrentProfile();
    const { error } = await _supabase.from('students').insert({
      full_name, class_id, national_id, created_by: prof?.id
    });
    if (error) return toast(error.message, 'error');
    $('student-name').value = ''; $('student-nid').value = '';
    toast('تمت إضافة الطالب ✅', 'success');
    await loadAllAdminData();
  });
}

// -------- Grades View --------
function bindGradesView() {
  $('gv-load')?.addEventListener('click', loadGradesView);
}

async function loadGradesView() {
  const tbody = $('gv-tbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="7">جاري التحميل...</td></tr>';

  let q = _supabase.from('grades').select(`
    id, score, absence_score, term, updated_at,
    students(full_name, class_id),
    subjects(name),
    classes(name),
    profiles(full_name, email)
  `).order('updated_at', { ascending: false }).limit(500);

  const cls = $('gv-class').value;
  const sub = $('gv-subject').value;
  if (cls) q = q.eq('class_id', cls);
  if (sub) q = q.eq('subject_id', sub);

  const { data, error } = await q;
  if (error) { tbody.innerHTML = `<tr><td colspan="7">خطأ: ${esc(error.message)}</td></tr>`; return; }
  if (!data.length) { tbody.innerHTML = '<tr><td colspan="7">لا بيانات</td></tr>'; return; }

  tbody.innerHTML = data.map(g => `
    <tr>
      <td>${esc(g.students?.full_name || '')}</td>
      <td>${esc(g.classes?.name || '')}</td>
      <td>${esc(g.subjects?.name || '')}</td>
      <td>${esc(termLabel(g.term))}</td>
      <td>${g.score ?? ''}</td>
      <td>${g.absence_score ?? ''}</td>
      <td>${esc(g.profiles?.full_name || g.profiles?.email || '')}</td>
    </tr>
  `).join('');
}

// ============================ EXCEL IMPORTS UI ============================
function bindExcelImports() {
  $('stages-template-btn')?.addEventListener('click', () => {
    downloadTemplate('قالب_المراحل.xlsx', ['name', 'sort_order'], ['المرحلة المتوسطة', 1]);
  });
  $('stages-import-btn')?.addEventListener('click', async () => {
    const f = $('stages-excel').files[0];
    if (!f) return toast('اختر ملف Excel أولاً', 'warning');
    try {
      const rows = await readExcelFile(f);
      if (!rows.length) return toast('الملف فارغ', 'warning');
      openImportPreview('stages', rows);
    } catch (e) { toast('خطأ: ' + e.message, 'error'); }
  });

  $('grades-template-btn')?.addEventListener('click', () => {
    downloadTemplate('قالب_الصفوف.xlsx', ['stage_id', 'name', 'sort_order'], [1, 'الصف الأول المتوسط', 1]);
  });
  $('grades-idref-btn')?.addEventListener('click', () => {
    const data = [['stage_id', 'stage_name']];
    adminState.stages.forEach(s => data.push([s.id, s.name]));
    downloadIdReference('مرجع_ارقام_المراحل.xlsx', data[0], data.slice(1));
  });
  $('grades-import-btn')?.addEventListener('click', async () => {
    const f = $('grades-excel').files[0];
    if (!f) return toast('اختر ملف Excel أولاً', 'warning');
    try {
      const rows = await readExcelFile(f);
      if (!rows.length) return toast('الملف فارغ', 'warning');
      openImportPreview('grades', rows);
    } catch (e) { toast('خطأ: ' + e.message, 'error'); }
  });

  $('classes-template-btn')?.addEventListener('click', () => {
    downloadTemplate('قالب_الفصول.xlsx', ['grade_level_id', 'name'], [1, 'فصل 1']);
  });
  $('classes-idref-btn')?.addEventListener('click', () => {
    const data = [['grade_level_id', 'stage_name', 'grade_name']];
    adminState.grades.forEach(g => {
      const s = adminState.stages.find(x => String(x.id) === String(g.stage_id));
      data.push([g.id, s?.name || '', g.name]);
    });
    downloadIdReference('مرجع_ارقام_الصفوف.xlsx', data[0], data.slice(1));
  });
  $('classes-import-btn')?.addEventListener('click', async () => {
    const f = $('classes-excel').files[0];
    if (!f) return toast('اختر ملف Excel أولاً', 'warning');
    try {
      const rows = await readExcelFile(f);
      if (!rows.length) return toast('الملف فارغ', 'warning');
      openImportPreview('classes', rows);
    } catch (e) { toast('خطأ: ' + e.message, 'error'); }
  });

  $('subjects-template-btn')?.addEventListener('click', () => {
    downloadTemplate('قالب_المواد.xlsx',
      ['name', 'max_score', 'absence_max_score'],
      ['رياضيات', 100, 5]);
  });
  $('subjects-import-btn')?.addEventListener('click', async () => {
    const f = $('subjects-excel').files[0];
    if (!f) return toast('اختر ملف Excel أولاً', 'warning');
    try {
      const rows = await readExcelFile(f);
      if (!rows.length) return toast('الملف فارغ', 'warning');
      openImportPreview('subjects', rows);
    } catch (e) { toast('خطأ: ' + e.message, 'error'); }
  });

  $('teachers-template-btn')?.addEventListener('click', () => {
    downloadTemplate('قالب_المدرسين.xlsx', ['full_name', 'email', 'password', 'role'],
      ['أ. محمد أحمد', 'mohammed@school.com', 'Teacher@123', 'teacher']);
  });
  $('teachers-import-btn')?.addEventListener('click', async () => {
    const f = $('teachers-excel').files[0];
    if (!f) return toast('اختر ملف Excel أولاً', 'warning');
    try {
      const rows = await readExcelFile(f);
      if (!rows.length) return toast('الملف فارغ', 'warning');
      openImportPreview('teachers', rows);
    } catch (e) { toast('خطأ: ' + e.message, 'error'); }
  });

  $('students-template-btn')?.addEventListener('click', () => {
    downloadTemplate('قالب_الطلاب.xlsx', ['class_id', 'full_name', 'national_id'],
      [1, 'عبدالله سعد الغامدي', '1234567890']);
  });
  $('students-idref-btn')?.addEventListener('click', () => {
    const data = [['class_id', 'stage_name', 'grade_name', 'class_name']];
    adminState.classes.forEach(c => {
      const g = adminState.grades.find(x => String(x.id) === String(c.grade_level_id));
      const s = g ? adminState.stages.find(x => String(x.id) === String(g.stage_id)) : null;
      data.push([c.id, s?.name || '', g?.name || '', c.name]);
    });
    downloadIdReference('مرجع_ارقام_الفصول.xlsx', data[0], data.slice(1));
  });
  $('students-import-btn')?.addEventListener('click', async () => {
    const f = $('students-excel').files[0];
    if (!f) return toast('اختر ملف Excel أولاً', 'warning');
    try {
      const rows = await readExcelFile(f);
      if (!rows.length) return toast('الملف فارغ', 'warning');
      openImportPreview('students', rows);
    } catch (e) { toast('خطأ: ' + e.message, 'error'); }
  });
}

// ============================ TEACHER DASHBOARD ============================
async function initTeacherDashboard() {
  const profile = await requireAuth();
  if (!profile) return;
  const p = await getCurrentProfile();
  const isAdmin = p?.role === 'admin';
  window._isAdmin = isAdmin;

  $('user-name').textContent = p?.full_name || '';
  $('role-badge').textContent = p?.role === 'teacher' ? 'مدرس' : (p?.role === 'admin' ? 'مدير' : 'موظف');

  // ⭐ إخفاء الأزرار غير المسموحة لغير الأدمن
  if (!isAdmin) {
    hideForNonAdmin();
  }

  const subjects = await getMySubjects();
  fillSelect($('subject'), subjects, { placeholder: 'اختر المادة' });

  const { data: stages } = await _supabase.from('stages').select('*').order('sort_order');
  fillSelect($('stage'), stages || [], { placeholder: 'اختر المرحلة' });

  $('stage').onchange = onStageChange;
  $('grade').onchange = onGradeChange;
  $('class').onchange = onClassChange;
  $('subject').onchange = onSubjectChange;
  $('term').onchange = loadStudentsAndGrades;
  $('save-all-btn').onclick = saveAllGrades;
  if (isAdmin && $('print-btn')) $('print-btn').onclick = printAllStudentsCertificates;
  if ($('bulk-fill-btn')) $('bulk-fill-btn').onclick = bulkFillFields;
}

async function onStageChange() {
  const stageId = $('stage').value;
  $('grade').innerHTML = '<option value="">اختر الصف</option>';
  $('class').innerHTML = '<option value="">اختر الفصل</option>';
  $('students-area').innerHTML = '';
  if (!stageId) return;
  const { data } = await _supabase.from('grade_levels').select('*').eq('stage_id', stageId).order('sort_order');
  fillSelect($('grade'), data || [], { placeholder: 'اختر الصف' });
}

async function onGradeChange() {
  const gradeId = $('grade').value;
  $('class').innerHTML = '<option value="">اختر الفصل</option>';
  $('students-area').innerHTML = '';
  if (!gradeId) return;
  const { data } = await _supabase.from('classes').select('*').eq('grade_level_id', gradeId).order('name');
  fillSelect($('class'), data || [], { placeholder: 'اختر الفصل' });
}

async function onClassChange() { await loadStudentsAndGrades(); }
async function onSubjectChange() { await loadStudentsAndGrades(); }

async function loadStudentsAndGrades() {
  const classId = $('class').value;
  const subjectId = $('subject').value;
  const term = $('term').value || 'term1';
  const area = $('students-area');
  const isAdmin = window._isAdmin === true;

  if (!classId) { area.innerHTML = ''; return; }
  area.innerHTML = '<p class="muted">جاري التحميل...</p>';

  let maxScore = 100, maxAbsence = 5;
  if (subjectId) {
    const { data: subj } = await _supabase
      .from('subjects').select('max_score, absence_max_score')
      .eq('id', subjectId).maybeSingle();
    if (subj) {
      maxScore = Number(subj.max_score) || 100;
      maxAbsence = Number(subj.absence_max_score) || 0;
    }
  }

  const { data: students, error: sErr } = await _supabase
    .from('students').select('*').eq('class_id', classId).eq('is_active', true).order('full_name');
  if (sErr) { area.innerHTML = `خطأ: ${esc(sErr.message)}`; return; }
  if (!students.length) { area.innerHTML = '<p class="muted">لا يوجد طلاب في هذا الفصل.</p>'; return; }

  let gradesMap = {}, absenceMap = {};
  if (subjectId) {
    const { data: gr } = await _supabase.from('grades')
      .select('student_id, score, absence_score')
      .eq('class_id', classId).eq('subject_id', subjectId).eq('term', term);
    (gr || []).forEach(g => {
      gradesMap[g.student_id] = g.score;
      absenceMap[g.student_id] = g.absence_score;
    });
  }

  area.innerHTML = `
    <table class="data-table">
      <thead>
        <tr>
          <th style="width:40px;">#</th>
          <th>اسم الطالب</th>
          <th style="width:130px;">الدرجة (من ${maxScore})</th>
          <th style="width:130px;">الغياب (من ${maxAbsence})</th>
          ${isAdmin ? '<th style="width:150px;">طباعة</th>' : ''}
        </tr>
      </thead>
      <tbody>
        ${students.map((s, i) => `
          <tr>
            <td>${i + 1}</td>
            <td>${esc(s.full_name)}</td>
            <td>
              <input type="number" class="form-control grade-input"
                     data-student-id="${s.id}"
                     value="${gradesMap[s.id] ?? ''}"
                     min="0" max="${maxScore}" step="0.5">
            </td>
            <td>
              <input type="number" class="form-control absence-input"
                     data-student-id="${s.id}"
                     value="${absenceMap[s.id] ?? ''}"
                     min="0" max="${maxAbsence}" step="0.5">
            </td>
            ${isAdmin ? `<td>
              <button class="btn btn-sm" onclick="window.__printOneStudent('${s.id}')">
                🖨️ شهادة الطالب
              </button>
            </td>` : ''}
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;

  bindValidationInputs();
}

function bindValidationInputs() {
  document.querySelectorAll('.grade-input, .absence-input').forEach(inp => {
    inp.addEventListener('input', (e) => {
      const v = e.target.value;
      if (v === '' || v === '-') return;
      if (!/^-?\d*\.?\d*$/.test(v)) {
        e.target.value = '';
        toast('يُسمح بالأرقام فقط', 'warning', 1800);
        return;
      }
      const num = parseFloat(v);
      const max = parseFloat(e.target.max);
      if (isNaN(num)) return;
      if (num < 0) {
        e.target.value = '';
        toast('لا يُسمح بالقيم السالبة', 'warning', 1800);
        return;
      }
      if (num > max) {
        e.target.value = '';
        toast(`الحد الأقصى ${max}`, 'warning', 1800);
      }
    });
  });
}

function bulkFillFields() {
  const scoreVal = $('bulk-score')?.value ?? '';
  const absenceVal = $('bulk-absence')?.value ?? '';

  if (scoreVal === '' && absenceVal === '') {
    return toast('اكتب قيمة في حقل الدرجة أو الغياب أولاً', 'warning');
  }

  let count = 0;
  if (scoreVal !== '') {
    document.querySelectorAll('.grade-input').forEach(inp => {
      const max = parseFloat(inp.max);
      const num = parseFloat(scoreVal);
      if (!isNaN(num) && num >= 0 && num <= max) {
        inp.value = scoreVal;
        count++;
      }
    });
  }
  if (absenceVal !== '') {
    document.querySelectorAll('.absence-input').forEach(inp => {
      const max = parseFloat(inp.max);
      const num = parseFloat(absenceVal);
      if (!isNaN(num) && num >= 0 && num <= max) {
        inp.value = absenceVal;
        count++;
      }
    });
  }

  toast('تم ملء الحقول ✅', 'success');
  if ($('bulk-score')) $('bulk-score').value = '';
  if ($('bulk-absence')) $('bulk-absence').value = '';
}
window.bulkFillFields = bulkFillFields;

window.__printOneStudent = async (studentId) => {
  if (!window._isAdmin) return toast('غير مصرح بالطباعة', 'error');
  const subjectId = $('subject').value;
  const term = $('term').value || 'term1';
  if (!subjectId) return toast('اختر المادة أولاً', 'warning');

  const { data, error } = await _supabase.from('grades').select(`
    id, score, absence_score, term,
    students(full_name, classes(name, grade_level_id, grade_levels(name, stages(name)))),
    subjects(name, max_score, absence_max_score),
    classes(name),
    profiles(full_name)
  `)
  .eq('student_id', studentId)
  .eq('subject_id', subjectId)
  .eq('term', term)
  .maybeSingle();

  if (error) return toast(error.message, 'error');
  if (!data) return toast('لا توجد درجة محفوظة لهذا الطالب', 'warning');

  safePrint(certHtml(data));
};

async function saveAllGrades() {
  const classId = $('class').value;
  const subjectId = $('subject').value;
  const term = $('term').value || 'term1';
  if (!classId || !subjectId) return toast('اختر الفصل والمادة', 'warning');

  const rows = [];
  document.querySelectorAll('.grade-input').forEach(inp => {
    const sid = inp.dataset.studentId;
    const absenceInp = document.querySelector(`.absence-input[data-student-id="${sid}"]`);
    rows.push({
      student_id: sid,
      score: inp.value,
      absence_score: absenceInp ? absenceInp.value : ''
    });
  });
  if (!rows.length) return toast('لا يوجد طلاب', 'warning');

  loadingDialog('جاري حفظ الدرجات...', '');
  const { error } = await _supabase.rpc('save_grades_bulk', {
    p_class_id: classId, p_subject_id: subjectId, p_term: term, p_rows: rows
  });
  await closeDialog();

  if (error) {
    Swal2.fire({ icon: 'error', title: 'فشل الحفظ', text: error.message, customClass: { popup: 'swal-rtl' } });
    return;
  }
  Swal2.fire({
    icon: 'success',
    title: '🎉 تم الحفظ بنجاح',
    text: `تم حفظ ${rows.length} درجة + غياب`,
    timer: 2000,
    showConfirmButton: false,
    customClass: { popup: 'swal-rtl' }
  });
}

async function printAllStudentsCertificates() {
  if (!window._isAdmin) return toast('غير مصرح بالطباعة', 'error');
  const classId = $('class').value;
  const subjectId = $('subject').value;
  const term = $('term').value || 'term1';
  if (!classId || !subjectId) return toast('اختر الفصل والمادة', 'warning');

  const { data, error } = await _supabase.from('grades').select(`
    id, score, absence_score, term,
    students(full_name, classes(name, grade_level_id, grade_levels(name, stages(name)))),
    subjects(name, max_score, absence_max_score),
    classes(name),
    profiles(full_name)
  `).eq('class_id', classId).eq('subject_id', subjectId).eq('term', term);

  if (error) return toast(error.message, 'error');
  if (!data?.length) return toast('لا توجد درجات محفوظة لهذا الفصل', 'warning');

  safePrint(data.map(g => certHtml(g)).join(''));
}

// ============================ CERTIFICATES PAGE ============================
async function initCertificatesPage() {
  // ⭐ فقط الأدمن يقدر يدخل
  const profile = await requireRole(['admin']);
  if (!profile) return;
  window._isAdmin = true;

  const p = await getCurrentProfile();
  $('user-name').textContent = p?.full_name || '';

  const { data: stages } = await _supabase.from('stages').select('*').order('sort_order');
  fillSelect($('f-stage'), stages || [], { placeholder: 'الكل' });

  const subjects = await getMySubjects();
  fillSelect($('f-subject'), subjects, { placeholder: 'الكل' });

  $('f-stage').onchange = async () => {
    const sid = $('f-stage').value;
    $('f-grade').innerHTML = '<option value="">الكل</option>';
    $('f-class').innerHTML = '<option value="">الكل</option>';
    if (!sid) return;
    const { data } = await _supabase.from('grade_levels').select('*').eq('stage_id', sid).order('sort_order');
    fillSelect($('f-grade'), data || [], { placeholder: 'الكل' });
  };
  $('f-grade').onchange = async () => {
    const gid = $('f-grade').value;
    $('f-class').innerHTML = '<option value="">الكل</option>';
    if (!gid) return;
    const { data } = await _supabase.from('classes').select('*').eq('grade_level_id', gid).order('name');
    fillSelect($('f-class'), data || [], { placeholder: 'الكل' });
  };
  $('f-load').onclick = loadCertificates;
  $('f-print-all').onclick = printAllCerts;
  if ($('f-print-full')) $('f-print-full').onclick = printClassFullReports;

  await loadCertificates();
}

let _currentCerts = [];

async function loadCertificates() {
  const tbody = $('cert-tbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="7">جاري التحميل...</td></tr>';

  let q = _supabase.from('grades').select(`
    id, score, absence_score, term,
    students(id, full_name, class_id, classes(name, grade_level_id, grade_levels(name, stages(name)))),
    subjects(name, max_score, absence_max_score),
    classes(name),
    profiles(full_name, email)
  `);

  const classId = $('f-class').value;
  const subjectId = $('f-subject').value;
  const term = $('f-term').value;
  if (classId) q = q.eq('class_id', classId);
  if (subjectId) q = q.eq('subject_id', subjectId);
  if (term) q = q.eq('term', term);

  const { data, error } = await q;
  if (error) { tbody.innerHTML = `<tr><td colspan="7">خطأ: ${esc(error.message)}</td></tr>`; return; }

  _currentCerts = data || [];
  if (!_currentCerts.length) { tbody.innerHTML = '<tr><td colspan="7">لا نتائج</td></tr>'; return; }

  tbody.innerHTML = _currentCerts.map(g => {
    const stage = g.students?.classes?.grade_levels?.stages?.name || '';
    const grade = g.students?.classes?.grade_levels?.name || '';
    return `
      <tr>
        <td>${esc(g.students?.full_name || '')}</td>
        <td>${esc(stage)}</td>
        <td>${esc(grade)}</td>
        <td>${esc(g.classes?.name || '')}</td>
        <td>${esc(g.subjects?.name || '')}</td>
        <td>${g.score ?? ''}</td>
        <td style="white-space:nowrap;">
          <button class="btn btn-sm" onclick="window.__printOne('${g.id}')">📄 المادة</button>
          <button class="btn btn-sm btn-secondary" onclick="window.printStudentFullReport('${g.id}')">📋 تقرير كامل</button>
        </td>
      </tr>
    `;
  }).join('');
}

// ============================ CERTIFICATE HTML ============================
function certHtml(g, allGrades = null) {
  const student = g.students || {};
  const cls = student.classes || g.classes || {};
  const gradeLevel = cls.grade_levels || {};
  const stage = gradeLevel.stages?.name || '';
  const grade = gradeLevel.name || '';
  const className = cls.name || '';

  const isMultiSubject = Array.isArray(allGrades) && allGrades.length > 1;
  const gradesList = (Array.isArray(allGrades) && allGrades.length) ? allGrades : [{
    subject_name: g.subjects?.name || '',
    score: g.score,
    absence_score: g.absence_score,
    max_score: g.subjects?.max_score ?? 100,
    absence_max_score: g.subjects?.absence_max_score ?? 5,
    term: g.term || ''
  }];

  const firstMaxScore = Number(gradesList[0]?.max_score) || 100;
  const firstMaxAbsence = Number(gradesList[0]?.absence_max_score) || 0;
  const totalMaxAll = firstMaxScore + firstMaxAbsence;

  const uniqueTerms = [...new Set(gradesList.map(x => x.term).filter(Boolean))];
  const singleTerm = uniqueTerms.length === 1 ? uniqueTerms[0] : null;

  const gradeRows = gradesList.map((x, i) => {
    const score = Number(x.score) || 0;
    const absence = Number(x.absence_score) || 0;
    const rowTotal = score + absence;
    return `<tr>
      <td>${i + 1}</td>
      <td>${esc(x.subject_name || '')}</td>
      <td>${score}</td>
      <td>${absence}</td>
      <td class="cert-col-total">${rowTotal}</td>
    </tr>`;
  }).join('');

  const totalRow = isMultiSubject ? (() => {
    let scoreSum = 0, absenceSum = 0;
    gradesList.forEach(x => {
      scoreSum += Number(x.score) || 0;
      absenceSum += Number(x.absence_score) || 0;
    });
    return `<tr class="cert-total-row">
      <td colspan="2">المجموع الكلي</td>
      <td>${scoreSum}</td>
      <td>${absenceSum}</td>
      <td>${scoreSum + absenceSum}</td>
    </tr>`;
  })() : '';

  let total = 0, totalMax = 0;
  gradesList.forEach(x => {
    total += Number(x.score) || 0;
    totalMax += Number(x.max_score) || 100;
  });
  const pct = totalMax ? Math.round((total / totalMax) * 1000) / 10 : 0;

  let achievementMsg = '';
  if (isMultiSubject) {
    if (pct >= 90) achievementMsg = `ممتاز — حقق الطالب مجموعاً كلياً <strong>${total}/${totalMax}</strong> بنسبة <strong>${pct}%</strong>.`;
    else if (pct >= 75) achievementMsg = `جيد جداً — حقق الطالب مجموعاً كلياً <strong>${total}/${totalMax}</strong> بنسبة <strong>${pct}%</strong>.`;
    else if (pct >= 60) achievementMsg = `جيد — حقق الطالب مجموعاً كلياً <strong>${total}/${totalMax}</strong> بنسبة <strong>${pct}%</strong>.`;
    else achievementMsg = `الأداء ضعيف في المستوى التحصيلي — يحتاج الطالب لمزيد من الجهد.`;
  } else {
    const x = gradesList[0];
    const score = Number(x.score) || 0;
    const max = Number(x.max_score) || 100;
    const sp = max ? Math.round((score / max) * 1000) / 10 : 0;
    if (sp >= 90) achievementMsg = `ممتاز — حقق الطالب <strong>${score}/${max}</strong> بنسبة <strong>${sp}%</strong> في مادة ${esc(x.subject_name || '')}.`;
    else if (sp >= 75) achievementMsg = `جيد جداً — حقق الطالب <strong>${score}/${max}</strong> بنسبة <strong>${sp}%</strong> في مادة ${esc(x.subject_name || '')}.`;
    else if (sp >= 60) achievementMsg = `جيد — حقق الطالب <strong>${score}/${max}</strong> بنسبة <strong>${sp}%</strong> في مادة ${esc(x.subject_name || '')}.`;
    else achievementMsg = `الأداء ضعيف في مادة ${esc(x.subject_name || '')} — يحتاج لمزيد من الجهد.`;
  }
  const disciplineMsg = 'نثمّن التزامك بالحضور والمواظبة، ونشجعك على الاستمرار في هذا السلوك المتميز.';

  const today = new Date();
  const hijri = today.toLocaleDateString('ar-SA-u-ca-islamic');
  const refNumber = `GES-${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;

  return `
    <div class="certificate-card">

      <!-- ⭐⭐ الهيدر: صورتين الخلفية + الشعارات فوقهم -->
      <div class="cert-header">
        <!-- الذهبي في الخلف -->
        <img class="bg-gold" src="Images/Image_Header_Certificate.png" alt="">
        <!-- التركواز فوق شمال -->
        <!--<img class="bg-turq" src="Images/cert-bottom-bg.png" alt="">-->
        <!-- الشعارات -->
        <div class="cert-logos">
          <img src="Images/Logo_Minister_Education_Sudia_With_Vision_3.png" alt="رؤية سعودية 2030">
          <img src="Images/logo-Geel_Ebdaa_3.png" alt="شعار المدرسة">
          <img src="Images/Logo_cognia_1.png" alt="Cognia">
        </div>
      </div>

      <!-- ⭐⭐ شريط العنوان التركواز -->
      <div class="cert-title-bar">
        <span class="title-line-left"></span>
        <h2>إشعار مستوى</h2>
        <span class="title-line-right"></span>
      </div>

      <!-- ⭐⭐ باقي المحتوى -->
      <div class="cert-body">
        <p class="cert-intro">
          حرصاً منا على إطلاع أولياء الأمور على المستوى الدراسي لأبنائهم الطلاب،
          نضع بين أيديكم هذا الإشعار بنتيجة تحصيل الطالب الدراسي خلال الفترة السابقة،
          حيث تم الاختبار منذ أسابيع وقبل نهاية الفترة. نأمل من أولياء الأمور الكرام
          متابعة أبنائهم الطلاب والاطلاع على مستوى أبنائهم الدراسي والاهتمام به،
          لتحقيق الأهداف المرجوة وتحقيق التقدم المنشود. نسأل الله العلي القدير
          أن يجعل أبناءنا من المتفوقين الناجحين.
        </p>

        <table class="cert-info-table">
          <tbody>
            <tr>
              <th>المدرسة</th>
              <td colspan="3">مدارس جيل الابداع المتوسطة بنين بمكة المكرمة</td>
            </tr>
            <tr>
              <th>اسم الطالب</th>
              <td colspan="3"><strong>${esc(student.full_name || '')}</strong></td>
            </tr>
            <tr>
              <th>الصف الدراسي</th>
              <td colspan="3">${esc(grade || '-')} ${stage ? ' - ' + esc(stage) : ''}</td>
            </tr>
            <tr>
              <th>الفصل</th>
              <td>${esc(className || '-')}</td>
              <th class="th-split">الفترة</th>
              <td>${singleTerm ? esc(termLabel(singleTerm)) : (uniqueTerms.length ? 'جميع الترمات' : '-')}</td>
            </tr>
          </tbody>
        </table>

        <table class="cert-grades-table">
          <thead>
            <tr>
              <th rowspan="2" style="width:35px;">م</th>
              <th rowspan="2">المواد الدراسية</th>
              <th>درجة الاختبار</th>
              <th>درجة الغياب</th>
              <th>مجموع الدرجة</th>
            </tr>
            <tr class="cert-max-row">
              <th>${firstMaxScore}</th>
              <th>${firstMaxAbsence}</th>
              <th>${totalMaxAll}</th>
            </tr>
          </thead>
          <tbody>
            ${gradeRows}
            ${totalRow}
          </tbody>
        </table>

        <div class="cert-section">
          <h4>رسالة التحصيل الدراسي</h4>
          <p>${achievementMsg}</p>
        </div>

        <div class="cert-section">
          <h4>رسالة الانضباط المدرسي (السلوك والمواظبة)</h4>
          <p>${disciplineMsg}</p>
        </div>

        <p class="cert-alert">
          كل الشكر والتقدير لأسرته لغرسها قيمة الانضباط المدرسي، والشكر موصول للابن البار لمحافظته على الانضباط.
        </p>

        <div class="cert-footer-block">
          <div class="cert-note">
            <div class="cert-note-title">📌 تنويه</div>
            <div class="cert-note-body">
              هذا الإشعار لتحديد مستوى الابن خلال الفترة الماضية.
              <div style="margin-top:2px;"><strong>التاريخ:</strong> ${hijri} هـ</div>
              <div><strong>المرجع:</strong> ${refNumber}</div>
            </div>
          </div>

          <div class="cert-stamp-box">
            <img src="Images/cert-stamp.png" alt="ختم المدرسة" class="stamp-img">
          </div>

          <div class="cert-signature">
            <div class="signature-label">مدير المدرسة</div>
            <div class="signature-name">مهنّد بن محمد بستناق</div>
            <img src="Images/cert-signature.png" alt="إمضاء المدير" class="signature-img">
          </div>
        </div>
      </div>

    </div>
  `;
}
// ⭐ طباعة شهادة مادة واحدة
window.__printOne = (id) => {
  if (!window._isAdmin) return toast('غير مصرح بالطباعة', 'error');
  const g = _currentCerts.find(x => String(x.id) === String(id));
  if (!g) return toast('لم يتم العثور على السجل', 'error');
  safePrint(certHtml(g));
};

// ⭐⭐ طباعة تقرير كامل (كل المواد للطالب)
async function printStudentFullReport(gradeId) {
  if (!window._isAdmin) return toast('غير مصرح بالطباعة', 'error');
  if (!gradeId || gradeId === 'undefined' || gradeId === 'null') {
    return toast('لم يتم تحديد الطالب', 'warning');
  }

  loadingDialog('جاري تجهيز التقرير...', '');

  const { data: gradeRow, error: gErr } = await _supabase
    .from('grades')
    .select('student_id')
    .eq('id', gradeId)
    .maybeSingle();

  if (gErr || !gradeRow?.student_id) {
    await closeDialog();
    return toast('لم يتم العثور على بيانات الطالب', 'error');
  }

  const studentId = gradeRow.student_id;

  const { data, error } = await _supabase
    .from('grades')
    .select(`
      id, score, term, absence_score,
      students(id, full_name,
        classes(name, grade_level_id,
          grade_levels(name, stages(name))
        )
      ),
      subjects(name, max_score, absence_max_score),
      classes(name),
      profiles(full_name)
    `)
    .eq('student_id', studentId);

  await closeDialog();

  if (error) return toast(error.message, 'error');
  if (!data?.length) return toast('لا توجد درجات لهذا الطالب', 'warning');

  const first = data[0];
  const allGrades = data.map(d => ({
    subject_name: d.subjects?.name || '',
    score: d.score,
    absence_score: d.absence_score,
    max_score: d.subjects?.max_score ?? 100,
    absence_max_score: d.subjects?.absence_max_score ?? 5,
    term: d.term
  }));

  setTimeout(() => {
    safePrint(certHtml(first, allGrades));
  }, 200);
}
window.printStudentFullReport = printStudentFullReport;

// ⭐⭐ طباعة تقارير الفصل الكاملة
async function printClassFullReports() {
  if (!window._isAdmin) return toast('غير مصرح بالطباعة', 'error');
  const classId = $('f-class').value;
  if (!classId) return toast('اختر الفصل أولاً', 'warning');

  loadingDialog('جاري تجهيز التقارير...', 'قد يأخذ لحظات حسب عدد الطلاب');

  try {
    const { data, error } = await _supabase
      .from('grades')
      .select(`
        id, score, term, absence_score, student_id,
        students(id, full_name,
          classes(name, grade_level_id,
            grade_levels(name, stages(name))
          )
        ),
        subjects(name, max_score, absence_max_score),
        classes(name),
        profiles(full_name)
      `)
      .eq('class_id', classId);

    await closeDialog();

    if (error) { toast(error.message, 'error'); return; }
    if (!data?.length) { toast('لا توجد درجات مسجلة لهذا الفصل', 'warning'); return; }

    const byStudent = {};
    data.forEach(g => {
      const sid = String(g.student_id);
      if (!byStudent[sid]) {
        byStudent[sid] = { first: g, grades: [] };
      }
      byStudent[sid].grades.push({
        subject_name: g.subjects?.name || '',
        score: g.score,
        absence_score: g.absence_score,
        max_score: g.subjects?.max_score ?? 100,
        absence_max_score: g.subjects?.absence_max_score ?? 5,
        term: g.term
      });
    });

    const studentsList = Object.values(byStudent).sort((a, b) =>
      (a.first.students?.full_name || '').localeCompare(b.first.students?.full_name || '', 'ar')
    );

    const htmls = studentsList.map(({ first, grades }) =>
      certHtml(first, grades)
    ).join('');

    safePrint(htmls);
    toast(`تم تجهيز ${studentsList.length} تقرير`, 'success');

  } catch (e) {
    await closeDialog();
    toast('خطأ: ' + e.message, 'error');
  }
}
window.printClassFullReports = printClassFullReports;

// ⭐ طباعة الكل
function printAllCerts() {
  if (!window._isAdmin) return toast('غير مصرح بالطباعة', 'error');
  if (!_currentCerts.length) return toast('لا بيانات للطباعة', 'warning');
  safePrint(_currentCerts.map(g => certHtml(g)).join(''));
}

// ============================ EDIT ITEMS ============================
async function editItem({ table, id, title, fields }) {
  const html = fields.map((f, i) => `
    <div style="text-align:right; margin-bottom:10px;">
      <label style="display:block; font-size:12px; font-weight:700; color:#1e5fa8; margin-bottom:4px;">${f.label}</label>
      ${f.type === 'select'
        ? `<select id="swal-f-${i}" class="swal2-input" style="width:100%; margin:0;">
            ${f.options.map(o => `<option value="${o.value}" ${String(f.value) === String(o.value) ? 'selected' : ''}>${o.label}</option>`).join('')}
          </select>`
        : `<input id="swal-f-${i}" class="swal2-input" type="${f.type || 'text'}"
            value="${esc(f.value ?? '')}" min="${f.min ?? ''}" max="${f.max ?? ''}"
            style="width:100%; margin:0;">`}
    </div>
  `).join('');

  const r = await Swal2.fire({
    title,
    html,
    showCancelButton: true,
    confirmButtonText: '💾 حفظ',
    cancelButtonText: 'إلغاء',
    confirmButtonColor: '#1e5fa8',
    cancelButtonColor: '#64748b',
    customClass: { popup: 'swal-rtl' },
    preConfirm: () => {
      const values = {};
      fields.forEach((f, i) => {
        values[f.key] = document.getElementById(`swal-f-${i}`).value.trim();
      });
      return values;
    }
  });

  if (!r.isConfirmed) return;

  loadingDialog('جاري الحفظ...', '');
  const { error } = await _supabase.from(table).update(r.value).eq('id', id);
  await closeDialog();

  if (error) {
    Swal2.fire({ icon: 'error', title: 'فشل', text: error.message, customClass: { popup: 'swal-rtl' } });
    return;
  }
  toast('تم التعديل ✅', 'success');
  await loadAllAdminData();
}

window.__editStage = async (id) => {
  const s = adminState.stages.find(x => String(x.id) === String(id));
  if (!s) return;
  await editItem({
    table: 'stages', id, title: 'تعديل المرحلة',
    fields: [
      { key: 'name', label: 'اسم المرحلة', value: s.name },
      { key: 'sort_order', label: 'الترتيب', type: 'number', value: s.sort_order || 0 }
    ]
  });
};

window.__editGrade = async (id) => {
  const g = adminState.grades.find(x => String(x.id) === String(id));
  if (!g) return;
  await editItem({
    table: 'grade_levels', id, title: 'تعديل الصف',
    fields: [
      { key: 'name', label: 'اسم الصف', value: g.name },
      { key: 'sort_order', label: 'الترتيب', type: 'number', value: g.sort_order || 0 }
    ]
  });
};

window.__editClass = async (id) => {
  const c = adminState.classes.find(x => String(x.id) === String(id));
  if (!c) return;
  await editItem({
    table: 'classes', id, title: 'تعديل الفصل',
    fields: [{ key: 'name', label: 'اسم الفصل', value: c.name }]
  });
};

window.__editSubject = async (id) => {
  const s = adminState.subjects.find(x => String(x.id) === String(id));
  if (!s) return;
  await editItem({
    table: 'subjects', id, title: 'تعديل المادة',
    fields: [
      { key: 'name', label: 'اسم المادة', value: s.name },
      { key: 'max_score', label: 'الدرجة القصوى للاختبار', type: 'number', value: s.max_score || 100 },
      { key: 'absence_max_score', label: 'الدرجة القصوى للغياب', type: 'number', value: s.absence_max_score || 5 }
    ]
  });
};

window.__editTeacher = async (id) => {
  const t = adminState.teachers.find(x => String(x.id) === String(id));
  if (!t) return;
  await editItem({
    table: 'profiles', id, title: 'تعديل المدرس',
    fields: [
      { key: 'full_name', label: 'الاسم', value: t.full_name },
      { key: 'role', label: 'الدور', type: 'select', value: t.role,
        options: [
          { value: 'teacher', label: 'مدرس' },
          { value: 'employee', label: 'موظف' },
          { value: 'admin', label: 'مدير' }
        ]}
    ]
  });
};

window.__editStudent = async (id) => {
  const s = adminState.students.find(x => String(x.id) === String(id));
  if (!s) return;
  const classesOptions = adminState.classes.map(c => {
    const g = adminState.grades.find(x => String(x.id) === String(c.grade_level_id));
    const st = g ? adminState.stages.find(x => String(x.id) === String(g.stage_id)) : null;
    return { value: c.id, label: `${st?.name || ''} / ${g?.name || ''} / ${c.name} (#${c.id})` };
  });
  await editItem({
    table: 'students', id, title: 'تعديل الطالب',
    fields: [
      { key: 'full_name', label: 'الاسم', value: s.full_name },
      { key: 'class_id', label: 'الفصل', type: 'select', value: s.class_id, options: classesOptions },
      { key: 'national_id', label: 'رقم الهوية', value: s.national_id || '' }
    ]
  });
};

// ============================ Router ============================
document.addEventListener('DOMContentLoaded', async () => {
  const page = document.body.dataset.page;
  try {
    if (page === 'login')        await initLoginPage();
    if (page === 'admin')        await initAdminPage();
    if (page === 'dashboard')    await initTeacherDashboard();
    if (page === 'certificates') await initCertificatesPage();
  } catch (e) {
    console.error(e);
    Swal2.fire({ icon: 'error', title: 'خطأ غير متوقع', text: e.message, customClass: { popup: 'swal-rtl' } });
  }
});