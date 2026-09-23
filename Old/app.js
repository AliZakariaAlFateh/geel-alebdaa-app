// app.js - إدارة المنطق وتحديث المراحل والفصول والشهادات

// هيكلة الفصول وفق المدارس
const structure = {
  "المرحلة المتوسطة": {
    "الصف الأول المتوسط": ["فصل 1", "فصل 2", "فصل 3"],
    "الصف الثاني المتوسط": ["فصل 1", "فصل 2", "فصل 3"],
    "الصف الثالث المتوسط": ["فصل 1", "فصل 2", "فصل 3"]
  },
  "المرحلة الثانوية": {
    "الصف الأول الثانوي": ["فصل 1", "فصل 2", "فصل 3"],
    "الصف الثاني الثانوي": ["فصل 1", "فصل 2", "فصل 3", "فصل 4"], // 4 فصول للصف الثاني الثانوي
    "الصف الثالث الثانوي": ["فصل 1", "فصل 2", "فصل 3"]
  }
};

// تحديث القوائم في Dashboard
function updateGradesAndClasses() {
  const stage = document.getElementById('stage').value;
  const gradeSelect = document.getElementById('grade_level');
  gradeSelect.innerHTML = '<option value="">اختر الصف</option>';
  
  if (stage && structure[stage]) {
    Object.keys(structure[stage]).forEach(grade => {
      gradeSelect.innerHTML += `<option value="${grade}">${grade}</option>`;
    });
  }
  updateClasses();
}

function updateClasses() {
  const stage = document.getElementById('stage').value;
  const grade = document.getElementById('grade_level').value;
  const classSelect = document.getElementById('class_name');
  classSelect.innerHTML = '<option value="">اختر الفصل</option>';

  if (stage && grade && structure[stage][grade]) {
    structure[stage][grade].forEach(c => {
      classSelect.innerHTML += `<option value="${c}">${c}</option>`;
    });
  }
}

// تحديث قوائم الفرز في Index
function updateFilterGradesAndClasses() {
  const stage = document.getElementById('filter-stage').value;
  const gradeSelect = document.getElementById('filter-grade');
  gradeSelect.innerHTML = '<option value="">الكل</option>';

  if (stage && structure[stage]) {
    Object.keys(structure[stage]).forEach(grade => {
      gradeSelect.innerHTML += `<option value="${grade}">${grade}</option>`;
    });
  }
  updateFilterClasses();
}

function updateFilterClasses() {
  const stage = document.getElementById('filter-stage').value;
  const grade = document.getElementById('filter-grade').value;
  const classSelect = document.getElementById('filter-class');
  classSelect.innerHTML = '<option value="">الكل</option>';

  if (stage && grade && structure[stage][grade]) {
    structure[stage][grade].forEach(c => {
      classSelect.innerHTML += `<option value="${c}">${c}</option>`;
    });
  }
}

// حفظ الطالب والدرجة
const addGradeForm = document.getElementById('add-grade-form');
if (addGradeForm) {
  addGradeForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const user = await getCurrentUser();

    const studentData = {
      stage: document.getElementById('stage').value,
      grade_level: document.getElementById('grade_level').value,
      class_name: document.getElementById('class_name').value,
      student_name: document.getElementById('student_name').value,
      subject: document.getElementById('subject').value,
      score: parseFloat(document.getElementById('score').value),
      created_by: user ? user.id : null
    };

    const { error } = await _supabase.from('grades').insert([studentData]);

    if (error) {
      alert('حدث خطأ أثناء التخزين: ' + error.message);
    } else {
      alert('تم حفظ البيانات بنجاح!');
      addGradeForm.reset();
    }
  });
}

// جلب وعرض بيانات الطلاب للطباعة
let currentStudentsData = [];

async function fetchAndRenderGrades() {
  const tbody = document.getElementById('students-table-body');
  if (!tbody) return;

  tbody.innerHTML = '<tr><td colspan="7">جاري تحميل البيانات...</td></tr>';

  let query = _supabase.from('grades').select('*');

  const stage = document.getElementById('filter-stage')?.value;
  const grade = document.getElementById('filter-grade')?.value;
  const className = document.getElementById('filter-class')?.value;
  const subject = document.getElementById('filter-subject')?.value;

  if (stage) query = query.eq('stage', stage);
  if (grade) query = query.eq('grade_level', grade);
  if (className) query = query.eq('class_name', className);
  if (subject) query = query.ilike('subject', `%${subject}%`);

  const { data, error } = await query;

  if (error) {
    tbody.innerHTML = '<tr><td colspan="7">حدث خطأ في جلب البيانات</td></tr>';
    return;
  }

  currentStudentsData = data;
  tbody.innerHTML = '';

  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="7">لا توجد بيانات مطابقة</td></tr>';
    return;
  }

  data.forEach((st) => {
    tbody.innerHTML += `
      <tr>
        <td>${st.student_name}</td>
        <td>${st.stage}</td>
        <td>${st.grade_level}</td>
        <td>${st.class_name}</td>
        <td>${st.subject}</td>
        <td>${st.score}</td>
        <td>
          <button class="btn" onclick="printSingleCertificate('${st.id}')">طباعة الشهادة</button>
        </td>
      </tr>
    `;
  });
}

// إنشاء هيكل الشهادة للطباعة
function generateCertHtml(st) {
  return `
    <div class="certificate-card">
      <div class="cert-header">
        <h2>مدارس جيل الإبداع بنين بمكة المكرمة</h2>
        <h4>شهادة إنجاز وتفوق في المادة</h4>
      </div>
      <div class="cert-body">
        <p>تشهد إدارة المدرسة بأن الطالب: <strong>${st.student_name}</strong></p>
        <p>المقيد بالصف: <strong>${st.grade_level} (${st.stage})</strong> - <strong>${st.class_name}</strong></p>
        <p>قد حصل في مادة: <strong>${st.subject}</strong> على درجة: <strong>${st.score}</strong></p>
      </div>
      <div class="cert-footer">
        <div>معلم المادة</div>
        <div>قائد المدرسة</div>
      </div>
    </div>
  `;
}

// طباعة شهادة طالب واحد
function printSingleCertificate(id) {
  const student = currentStudentsData.find(s => s.id == id);
  if (!student) return;

  const printArea = document.getElementById('certificate-print-area');
  printArea.innerHTML = generateCertHtml(student);
  window.print();
}

// طباعة شهادات الفصل كاملاً
function printAllFilteredCertificates() {
  if (currentStudentsData.length === 0) {
    alert('لا يوجد طلاب محددون للطباعة!');
    return;
  }

  const printArea = document.getElementById('certificate-print-area');
  printArea.innerHTML = '';

  currentStudentsData.forEach(student => {
    printArea.innerHTML += generateCertHtml(student);
  });

  window.print();
}