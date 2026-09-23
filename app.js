// app.js - منطق كل الصفحات + Excel + Preview Modal

// ============================ Utils ============================
const $ = (id) => document.getElementById(id);

function esc(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

// ============================ SweetAlert2 Helpers ============================
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

function loadingDialog(title = 'جاري المعالجة...', text = 'يرجى الانتظار قليلاً') {
  Swal2.fire({
    title, html: text,
    allowOutsideClick: false,
    allowEscapeKey: false,
    showConfirmButton: false,
    didOpen: () => Swal2.showLoading()
  });
}

function closeDialog() { Swal2.close(); }

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
    .select('subject_id, subjects(id, name, max_score)')
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

// ============================ IMPORT PREVIEW MODAL ============================
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
        { key: 'max_score', label: 'الدرجة القصوى', type: 'number' }
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
    closeDialog();
    closeImportModal();
    await showImportResult(res.inserted, res.skipped, res.errors, IMPORT_ENTITY[type]);
    await loadAllAdminData();
  } catch (e) {
    closeDialog();
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

// ============================ Login page ============================
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
    closeDialog();

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

// -------- Stages / Grades / Classes --------
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
        <button class="btn btn-sm btn-danger" data-del-stage="${st.id}">حذف</button>
      </div>
      <div class="tree-children">
        ${grades.map(g => {
          const cls = adminState.classes.filter(c => String(c.grade_level_id) === String(g.id));
          return `
            <div class="tree-sub">
              <div class="tree-head">
                <span>${esc(g.name)} <small class="muted">#${g.id}</small></span>
                <button class="btn btn-sm btn-danger" data-del-grade="${g.id}">حذف</button>
              </div>
              <div class="chips">
                ${cls.map(c => `
                  <span class="chip">
                    ${esc(c.name)} <small style="opacity:.7;">#${c.id}</small>
                    <button class="chip-x" data-del-class="${c.id}">×</button>
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
  closeDialog();
  if (error) {
    Swal2.fire({ icon: 'error', title: 'فشل الحذف', text: error.message, customClass: { popup: 'swal-rtl' } });
    return;
  }
  toast('تم الحذف بنجاح ✅', 'success');
  await loadAllAdminData();
}

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
    const max_score = parseInt($('subject-max').value) || 100;
    if (!name) return toast('اكتب اسم المادة', 'warning');
    const { error } = await _supabase.from('subjects').insert({ name, max_score });
    if (error) return toast(error.message, 'error');
    $('subject-name').value = ''; $('subject-max').value = 100;
    toast('تمت إضافة المادة ✅', 'success');
    await loadAllAdminData();
  });
}

// -------- Subjects --------
function renderSubjectsList() {
  const root = $('subjects-list');
  if (!root) return;
  if (!adminState.subjects.length) { root.innerHTML = '<p class="muted">لا مواد بعد.</p>'; return; }
  root.innerHTML = adminState.subjects.map(s => `
    <div class="row-item">
      <span>${esc(s.name)} <small class="muted">(الدرجة القصوى: ${s.max_score}) #${s.id}</small></span>
      <button class="btn btn-sm btn-danger" onclick="window.__delSubject('${s.id}')">حذف</button>
    </div>
  `).join('');
}
window.__delSubject = async (id) => { await delRow('subjects', id, 'المادة'); };

// -------- Teachers --------
function renderTeachersList() {
  const root = $('teachers-list');
  if (!root) return;
  if (!adminState.teachers.length) { root.innerHTML = '<p class="muted">لا مدرسين بعد.</p>'; return; }
  root.innerHTML = adminState.teachers.map(t => `
    <div class="row-item">
      <span>${esc(t.full_name)} <small class="muted">${esc(t.email || '')} — ${esc(t.role)}</small></span>
      <button class="btn btn-sm btn-danger" onclick="window.__delTeacher('${t.id}')">حذف</button>
    </div>
  `).join('');
}
window.__delTeacher = async (id) => {
  const ok = await confirmDialog('حذف المدرس؟', 'سيتم حذفه من النظام (لن يُحذف من Auth). متأكد؟', 'نعم، احذف');
  if (!ok) return;
  loadingDialog('جاري الحذف...', '');
  await _supabase.from('profiles').delete().eq('id', id);
  closeDialog();
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
    closeDialog();

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

// -------- Teacher assignments --------
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

// -------- Students --------
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
            <td><button class="btn btn-sm btn-danger" onclick="window.__delStudent('${s.id}')">حذف</button></td>
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

// -------- Grades view --------
function bindGradesView() {
  $('gv-load')?.addEventListener('click', loadGradesView);
}

async function loadGradesView() {
  const tbody = $('gv-tbody');
  if (!tbody) return;
  tbody.innerHTML = '<tr><td colspan="6">جاري التحميل...</td></tr>';

  let q = _supabase.from('grades').select(`
    id, score, term, updated_at,
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
  if (error) { tbody.innerHTML = `<tr><td colspan="6">خطأ: ${esc(error.message)}</td></tr>`; return; }
  if (!data.length) { tbody.innerHTML = '<tr><td colspan="6">لا بيانات</td></tr>'; return; }

  tbody.innerHTML = data.map(g => `
    <tr>
      <td>${esc(g.students?.full_name || '')}</td>
      <td>${esc(g.classes?.name || '')}</td>
      <td>${esc(g.subjects?.name || '')}</td>
      <td>${esc(g.term)}</td>
      <td>${g.score ?? ''}</td>
      <td>${esc(g.profiles?.full_name || g.profiles?.email || '')}</td>
    </tr>
  `).join('');
}

// ============================ EXCEL IMPORTS ============================
function bindExcelImports() {
  // --- Stages ---
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
    } catch (e) { toast('خطأ في قراءة الملف: ' + e.message, 'error'); }
  });

  // --- Grade Levels ---
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
    } catch (e) { toast('خطأ في قراءة الملف: ' + e.message, 'error'); }
  });

  // --- Classes ---
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
    } catch (e) { toast('خطأ في قراءة الملف: ' + e.message, 'error'); }
  });

  // --- Subjects ---
  $('subjects-template-btn')?.addEventListener('click', () => {
    downloadTemplate('قالب_المواد.xlsx', ['name', 'max_score'], ['رياضيات', 100]);
  });
  $('subjects-import-btn')?.addEventListener('click', async () => {
    const f = $('subjects-excel').files[0];
    if (!f) return toast('اختر ملف Excel أولاً', 'warning');
    try {
      const rows = await readExcelFile(f);
      if (!rows.length) return toast('الملف فارغ', 'warning');
      openImportPreview('subjects', rows);
    } catch (e) { toast('خطأ في قراءة الملف: ' + e.message, 'error'); }
  });

  // --- Teachers ---
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
    } catch (e) { toast('خطأ في قراءة الملف: ' + e.message, 'error'); }
  });

  // --- Students ---
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
    } catch (e) { toast('خطأ في قراءة الملف: ' + e.message, 'error'); }
  });
}

// ---- Import functions (تعمل بالـ IDs) ----
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
    toInsert.push({ name, max_score: parseInt(r.max_score) || 100 });
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

// ============================ TEACHER DASHBOARD ============================
async function initTeacherDashboard() {
  const profile = await requireAuth();
  if (!profile) return;
  const p = await getCurrentProfile();
  $('user-name').textContent = p?.full_name || '';
  $('role-badge').textContent = p?.role === 'teacher' ? 'مدرس' : (p?.role === 'admin' ? 'مدير' : 'موظف');

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
  $('print-btn').onclick = printAllStudentsCertificates;
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

  if (!classId) { area.innerHTML = ''; return; }
  area.innerHTML = '<p class="muted">جاري التحميل...</p>';

  const { data: students, error: sErr } = await _supabase
    .from('students').select('*').eq('class_id', classId).eq('is_active', true).order('full_name');
  if (sErr) { area.innerHTML = `خطأ: ${esc(sErr.message)}`; return; }
  if (!students.length) { area.innerHTML = '<p class="muted">لا يوجد طلاب في هذا الفصل.</p>'; return; }

  let gradesMap = {};
  if (subjectId) {
    const { data: gr } = await _supabase.from('grades')
      .select('student_id, score').eq('class_id', classId).eq('subject_id', subjectId).eq('term', term);
    (gr || []).forEach(g => gradesMap[g.student_id] = g.score);
  }

  area.innerHTML = `
    <table class="data-table">
      <thead>
        <tr>
          <th style="width:40px;">#</th>
          <th>اسم الطالب</th>
          <th style="width:140px;">الدرجة</th>
          <th style="width:140px;">طباعة</th>
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
                     min="0" max="100" step="0.5">
            </td>
            <td>
              <button class="btn btn-sm" onclick="window.__printOneStudent('${s.id}')">
                🖨️ شهادة الطالب
              </button>
            </td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}

window.__printOneStudent = async (studentId) => {
  const subjectId = $('subject').value;
  const term = $('term').value || 'term1';
  if (!subjectId) return toast('اختر المادة أولاً', 'warning');

  const { data, error } = await _supabase.from('grades').select(`
    id, score, term,
    students(full_name, classes(name, grade_level_id, grade_levels(name, stages(name)))),
    subjects(name, max_score),
    classes(name),
    profiles(full_name)
  `)
  .eq('student_id', studentId)
  .eq('subject_id', subjectId)
  .eq('term', term)
  .maybeSingle();

  if (error) return toast(error.message, 'error');
  if (!data) return toast('لا توجد درجة محفوظة لهذا الطالب', 'warning');

  $('print-area').innerHTML = certHtml(data);
  window.print();
};

async function saveAllGrades() {
  const classId = $('class').value;
  const subjectId = $('subject').value;
  const term = $('term').value || 'term1';
  if (!classId || !subjectId) return toast('اختر الفصل والمادة', 'warning');

  const rows = [];
  document.querySelectorAll('.grade-input').forEach(inp => {
    rows.push({ student_id: inp.dataset.studentId, score: inp.value });
  });
  if (!rows.length) return toast('لا يوجد طلاب', 'warning');

  loadingDialog('جاري حفظ الدرجات...', '');
  const { error } = await _supabase.rpc('save_grades_bulk', {
    p_class_id: classId, p_subject_id: subjectId, p_term: term, p_rows: rows
  });
  closeDialog();

  if (error) {
    Swal2.fire({ icon: 'error', title: 'فشل الحفظ', text: error.message, customClass: { popup: 'swal-rtl' } });
    return;
  }
  Swal2.fire({
    icon: 'success',
    title: '🎉 تم الحفظ بنجاح',
    text: `تم حفظ ${rows.length} درجة`,
    timer: 2000,
    showConfirmButton: false,
    customClass: { popup: 'swal-rtl' }
  });
}

async function printAllStudentsCertificates() {
  const classId = $('class').value;
  const subjectId = $('subject').value;
  const term = $('term').value || 'term1';
  if (!classId || !subjectId) return toast('اختر الفصل والمادة', 'warning');

  const { data, error } = await _supabase.from('grades').select(`
    id, score, term,
    students(full_name, classes(name, grade_level_id, grade_levels(name, stages(name)))),
    subjects(name, max_score),
    classes(name),
    profiles(full_name)
  `).eq('class_id', classId).eq('subject_id', subjectId).eq('term', term);

  if (error) return toast(error.message, 'error');
  if (!data?.length) return toast('لا توجد درجات محفوظة لهذا الفصل', 'warning');

  $('print-area').innerHTML = data.map(certHtml).join('');
  window.print();
}

// ============================ CERTIFICATES PAGE ============================
async function initCertificatesPage() {
  const profile = await requireAuth();
  if (!profile) return;
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

  await loadCertificates();
}

let _currentCerts = [];

async function loadCertificates() {
  const tbody = $('cert-tbody');
  tbody.innerHTML = '<tr><td colspan="7">جاري التحميل...</td></tr>';

  let q = _supabase.from('grades').select(`
    id, score, term,
    students(full_name, class_id, classes(name, grade_level_id, grade_levels(name, stages(name)))),
    subjects(name, max_score),
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
        <td><button class="btn btn-sm" onclick="window.__printOne('${g.id}')">طباعة</button></td>
      </tr>
    `;
  }).join('');
}

function certHtml(g) {
  const stage = g.students?.classes?.grade_levels?.stages?.name || '';
  const grade = g.students?.classes?.grade_levels?.name || '';
  const className = g.classes?.name || '';
  const max = g.subjects?.max_score ?? 100;
  return `
    <div class="certificate-card">
      <div class="cert-logos">
        <img src="Images/Sudia_Vision.png" alt="رؤية سعودية">
        <img src="Images/Logo_For_GeelEbdaa_School.png" alt="شعار المدرسة">
        <img src="Images/Logo_Cognia.png" alt="Cognia">
      </div>
      <div class="cert-header">
        <h2>مدارس جيل الإبداع بنين بمكة المكرمة</h2>
        <h4>شهادة إنجاز وتفوّق</h4>
      </div>
      <div class="cert-body">
        <p>تشهد إدارة المدرسة بأن الطالب: <strong>${esc(g.students?.full_name || '')}</strong></p>
        <p>المقيد بـ: <strong>${esc(stage)} - ${esc(grade)} - ${esc(className)}</strong></p>
        <p>قد حصل في مادة <strong>${esc(g.subjects?.name || '')}</strong>
           على درجة <strong>${g.score ?? '-'} / ${max}</strong> (${esc(g.term)})</p>
      </div>
      <div class="cert-footer">
        <div>معلم المادة<br>${esc(g.profiles?.full_name || '')}</div>
        <div>مدير المدرسة</div>
      </div>
    </div>
  `;
}

window.__printOne = (id) => {
  const g = _currentCerts.find(x => String(x.id) === String(id));
  if (!g) return;
  $('print-area').innerHTML = certHtml(g);
  window.print();
};

function printAllCerts() {
  if (!_currentCerts.length) return toast('لا بيانات للطباعة', 'warning');
  $('print-area').innerHTML = _currentCerts.map(certHtml).join('');
  window.print();
}

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