const STORAGE_KEY = "discipline-pulse-records";
const RECORDER_KEY = "discipline-pulse-recorder";
const STUDENTS_KEY = "discipline-pulse-students";

const form = document.querySelector("#disciplineForm");
const studentIdInput = document.querySelector("#studentId");
const studentNameInput = document.querySelector("#studentName");
const recorderNameInput = document.querySelector("#recorderName");
const otherRuleField = document.querySelector("#otherRuleField");
const otherRuleInput = document.querySelector("#otherRule");
const recordsList = document.querySelector("#recordsList");
const emptyState = document.querySelector("#emptyState");
const searchInput = document.querySelector("#searchInput");
const ruleFilter = document.querySelector("#ruleFilter");
const clearAllButton = document.querySelector("#clearAllButton");
const todayCount = document.querySelector("#todayCount");
const totalCount = document.querySelector("#totalCount");
const latestTime = document.querySelector("#latestTime");
const liveClock = document.querySelector("#liveClock");
const toast = document.querySelector("#toast");
const excelFile = document.querySelector("#excelFile");
const importStatus = document.querySelector("#importStatus");
const studentNameList = document.querySelector("#studentNameList");
const studentDetailsPanel = document.querySelector("#studentDetailsPanel");
const detailGrade = document.querySelector("#detailGrade");
const detailFullName = document.querySelector("#detailFullName");

let records = loadRecords();
let students = loadStudents();

const accentByRule = {
  "ถุงเท้าผิดระเบียบ": "#d6ff45",
  "กางเกงผิดระเบียบ": "#52e6c6",
  "เสื้อผิดระเบียบ": "#ff6b57",
  "เข็มขัดผิดระเบียบ": "#8c7cff",
  "รองเท้าผิดระเบียบ": "#4aa8ff",
  "อื่นๆ": "#ffd166",
};

recorderNameInput.value = localStorage.getItem(RECORDER_KEY) || "";

function loadStudents() {
  try {
    return JSON.parse(localStorage.getItem(STUDENTS_KEY)) || {};
  } catch {
    return {};
  }
}

function saveStudents() {
  localStorage.setItem(STUDENTS_KEY, JSON.stringify(students));
}

function updateStudentNameList() {
  studentNameList.innerHTML = "";
  Object.values(students).forEach((student) => {
    const option = document.createElement("option");
    option.value = student.fullName;
    studentNameList.append(option);
  });
}

function showStudentDetails(studentId) {
  const student = students[studentId];
  if (!student) {
    studentDetailsPanel.classList.add("is-hidden");
    showToast("❌ ไม่พบข้อมูลนักเรียน");
    return;
  }
  
  // แสดงข้อมูลในช่อง input
  studentNameInput.value = student.fullName;
  
  // แสดงข้อมูลในแผง
  detailGrade.textContent = `${student.grade} - ${student.sheet}`;
  detailFullName.textContent = student.fullName;
  studentDetailsPanel.classList.remove("is-hidden");
  
  console.log("📋 ข้อมูลนักเรียน:", student);
}

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      
      // อ่านจากทุกแท็บ (sheets)
      students = {};
      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Skip header row (index 0)
        rows.slice(1).forEach((row) => {
          if (row.length >= 5) {
            const studentId = String(row[1]).padStart(5, "0"); // เลขประจำตัว (column 1)
            if (studentId && studentId !== "00000") {
              const prefix = row[3] || ""; // คำนำหน้า (column 3)
              const firstName = row[4] || ""; // ชื่อ (column 4)
              const lastName = row[5] || ""; // สกุล (column 5)
              const fullName = `${prefix} ${firstName} ${lastName}`.trim();
              
              // แยกชั้นและหอง จากชื่อแท็บ (เช่น "ม.4-1" → ม.4 หอง 1)
              const [grade, room] = sheetName.split("-");
              
              students[studentId] = {
                studentId,
                fullName,
                prefix,
                firstName,
                lastName,
                grade: grade || "N/A",
                room: room || "N/A",
                sheet: sheetName,
              };
            }
          }
        });
      });

      saveStudents();
      updateStudentNameList();
      importStatus.innerHTML = `✅ นำเข้าสำเร็จ (${Object.keys(students).length} นักเรียน)`;
      importStatus.classList.add("is-success");
      setTimeout(() => {
        importStatus.innerHTML = "";
        importStatus.classList.remove("is-success");
        excelFile.value = "";
      }, 3000);
      showToast(`นำเข้าข้อมูล ${Object.keys(students).length} นักเรียน จาก ${workbook.SheetNames.length} ชั้น`);
    } catch (error) {
      importStatus.innerHTML = `❌ เกิดข้อผิดพลาด: ${error.message}`;
      importStatus.classList.add("is-error");
      console.error(error);
    }
  };
  reader.readAsArrayBuffer(file);
}

function saveRecords() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
}

function formatDateTime(dateValue) {
  return new Intl.DateTimeFormat("th-TH", {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(new Date(dateValue));
}

function formatTime(dateValue) {
  return new Intl.DateTimeFormat("th-TH", {
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(dateValue));
}

function getSelectedRule() {
  const selected = form.querySelector('input[name="rule"]:checked');
  if (!selected) return "";
  if (selected.value === "อื่นๆ") {
    return otherRuleInput.value.trim() || "อื่นๆ";
  }
  return selected.value;
}

function getRuleGroup(rule) {
  return Object.keys(accentByRule).includes(rule) ? rule : "อื่นๆ";
}

function renderRecords() {
  const query = searchInput.value.trim().toLowerCase();
  const selectedRule = ruleFilter.value;

  const visibleRecords = records.filter((record) => {
    const searchable = `${record.studentId} ${record.studentName} ${record.rule} ${record.recorderName}`.toLowerCase();
    const matchesSearch = searchable.includes(query);
    const matchesRule = selectedRule === "all" || getRuleGroup(record.rule) === selectedRule;
    return matchesSearch && matchesRule;
  });

  recordsList.innerHTML = "";
  emptyState.style.display = visibleRecords.length ? "none" : "grid";

  visibleRecords.forEach((record) => {
    const article = document.createElement("article");
    article.className = "record-card";
    article.style.setProperty("--accent", accentByRule[getRuleGroup(record.rule)]);
    article.innerHTML = `
      <div class="record-main">
        <span class="student-code">${escapeHtml(record.studentId)}</span>
        <div>
          <h3>${escapeHtml(record.studentName)}</h3>
          <p>ผู้บันทึก: ${escapeHtml(record.recorderName)}</p>
        </div>
        <span class="rule-badge">${escapeHtml(record.rule)}</span>
      </div>
      <div class="record-meta">
        <span>บันทึกเมื่อ ${formatDateTime(record.createdAt)}</span>
      </div>
      <button class="delete-record" type="button" data-id="${record.id}">ลบ</button>
    `;
    recordsList.append(article);
  });

  updateSummary();
}

function updateSummary() {
  const todayKey = new Date().toDateString();
  const countToday = records.filter((record) => new Date(record.createdAt).toDateString() === todayKey).length;

  todayCount.textContent = countToday;
  totalCount.textContent = records.length;
  latestTime.textContent = records[0] ? formatTime(records[0].createdAt) : "-";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("is-visible");
  window.setTimeout(() => toast.classList.remove("is-visible"), 2200);
}

function updateClock() {
  liveClock.textContent = formatDateTime(new Date());
}

form.addEventListener("change", (event) => {
  if (event.target.name !== "rule") return;

  const isOther = event.target.value === "อื่นๆ";
  otherRuleField.classList.toggle("is-hidden", !isOther);
  otherRuleInput.required = isOther;
  if (isOther) otherRuleInput.focus();
});

studentIdInput.addEventListener("input", () => {
  studentIdInput.value = studentIdInput.value.replace(/\D/g, "").slice(0, 5);
});

recorderNameInput.addEventListener("input", () => {
  localStorage.setItem(RECORDER_KEY, recorderNameInput.value.trim());
});

studentIdInput.addEventListener("blur", () => {
  const studentId = studentIdInput.value.trim();
  if (studentId && students[studentId]) {
    showStudentDetails(studentId);
  }
});

excelFile.addEventListener("change", handleExcelUpload);

form.addEventListener("submit", (event) => {
  event.preventDefault();

  const studentId = studentIdInput.value.trim();
  if (!/^\d{5}$/.test(studentId)) {
    studentIdInput.focus();
    showToast("กรุณากรอกเลขประจำตัว 5 หลัก");
    return;
  }

  const record = {
    id: crypto.randomUUID(),
    studentId,
    studentName: studentNameInput.value.trim(),
    rule: getSelectedRule(),
    recorderName: recorderNameInput.value.trim(),
    createdAt: new Date().toISOString(),
  };

  records = [record, ...records];
  saveRecords();
  renderRecords();
  showToast("บันทึกข้อมูลเรียบร้อยแล้ว");

  const recorderName = recorderNameInput.value;
  form.reset();
  recorderNameInput.value = recorderName;
  otherRuleField.classList.add("is-hidden");
  otherRuleInput.required = false;
  studentDetailsPanel.classList.add("is-hidden");
  studentIdInput.focus();
});

recordsList.addEventListener("click", (event) => {
  const button = event.target.closest(".delete-record");
  if (!button) return;

  records = records.filter((record) => record.id !== button.dataset.id);
  saveRecords();
  renderRecords();
  showToast("ลบรายการแล้ว");
});

clearAllButton.addEventListener("click", () => {
  if (!records.length) return;
  const confirmed = window.confirm("ต้องการล้างข้อมูลทั้งหมดใช่ไหม");
  if (!confirmed) return;

  records = [];
  saveRecords();
  renderRecords();
  showToast("ล้างข้อมูลทั้งหมดแล้ว");
});

searchInput.addEventListener("input", renderRecords);
ruleFilter.addEventListener("change", renderRecords);

updateClock();
window.setInterval(updateClock, 1000);
renderRecords();
updateStudentNameList();
