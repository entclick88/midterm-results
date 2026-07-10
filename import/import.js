// แปลงไฟล์ CSV คะแนน → import.sql สำหรับนำเข้า D1
// วิธีใช้:  node import.js grades.csv
// จากนั้น: cd ../worker && npx wrangler d1 execute midterm-results-db --remote --file=../import/import.sql
//
// คอลัมน์ CSV (บรรทัดแรกเป็นหัวตาราง, บันทึกจาก Excel เป็น "CSV UTF-8"):
// citizen_id,student_id,prefix,first_name,last_name,class,class_no,subject_code,subject_name,credits,midterm_score,max_score,teacher
// (1 แถว = นักเรียน 1 คน x 1 วิชา — นักเรียนคนเดิมใส่ซ้ำหลายแถวตามจำนวนวิชา)

const fs = require("fs");
const path = require("path");

const file = process.argv[2];
if (!file) {
  console.error("วิธีใช้: node import.js <ไฟล์.csv>");
  process.exit(1);
}

let text = fs.readFileSync(file, "utf8");
if (text.charCodeAt(0) === 0xfeff) text = text.slice(1); // ตัด BOM จาก Excel

// ตัวแยก CSV รองรับค่าที่ครอบด้วยเครื่องหมายคำพูด
function parseLine(line) {
  const out = [];
  let cur = "", inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQ) {
      if (c === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') inQ = false;
      else cur += c;
    } else if (c === '"') inQ = true;
    else if (c === ",") { out.push(cur); cur = ""; }
    else cur += c;
  }
  out.push(cur);
  return out.map((s) => s.trim());
}

const lines = text.split(/\r?\n/).filter((l) => l.trim() !== "");
const header = parseLine(lines[0]).map((h) => h.toLowerCase());
const required = ["citizen_id", "first_name", "last_name", "subject_name", "midterm_score"];
for (const r of required) {
  if (!header.includes(r)) {
    console.error(`ไม่พบคอลัมน์ "${r}" ในไฟล์ CSV — ตรวจสอบหัวตารางอีกครั้ง`);
    process.exit(1);
  }
}

const q = (v) => (v === undefined || v === "" ? "NULL" : `'${String(v).replace(/'/g, "''")}'`);
const n = (v) => (v === undefined || v === "" || isNaN(Number(v)) ? "NULL" : Number(v));

const students = new Map();
const gradeRows = [];
const errors = [];

for (let i = 1; i < lines.length; i++) {
  const cols = parseLine(lines[i]);
  const row = Object.fromEntries(header.map((h, j) => [h, cols[j]]));
  const cid = (row.citizen_id || "").replace(/\D/g, "");

  if (!/^\d{13}$/.test(cid)) {
    errors.push(`แถวที่ ${i + 1}: เลขบัตรประชาชนไม่ครบ 13 หลัก (${row.citizen_id})`);
    continue;
  }

  students.set(cid, `(${q(cid)}, ${q(row.student_id)}, ${q(row.prefix)}, ${q(row.first_name)}, ${q(row.last_name)}, ${q(row.class)}, ${n(row.class_no)})`);
  gradeRows.push(`(${q(cid)}, ${q(row.subject_code)}, ${q(row.subject_name)}, ${n(row.credits)}, ${n(row.midterm_score)}, ${n(row.max_score) === "NULL" ? 100 : n(row.max_score)}, ${q(row.teacher)})`);
}

if (errors.length) {
  console.error("พบข้อผิดพลาด:\n" + errors.join("\n"));
  process.exit(1);
}

const sql = [
  "-- สร้างโดย import.js เมื่อ " + new Date().toISOString(),
  "DELETE FROM grades;",
  "DELETE FROM students;",
  "INSERT INTO students (citizen_id, student_id, prefix, first_name, last_name, class, class_no) VALUES",
  [...students.values()].join(",\n") + ";",
  "INSERT INTO grades (citizen_id, subject_code, subject_name, credits, midterm_score, max_score, teacher) VALUES",
  gradeRows.join(",\n") + ";",
].join("\n");

const out = path.join(__dirname, "import.sql");
fs.writeFileSync(out, sql, "utf8");
console.log(`สำเร็จ: นักเรียน ${students.size} คน, ${gradeRows.length} รายการวิชา`);
console.log(`สร้างไฟล์แล้ว: ${out}`);
console.log("นำเข้าจริงด้วย: cd ../worker && npx wrangler d1 execute midterm-results-db --remote --file=../import/import.sql");
