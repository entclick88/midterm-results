-- โครงสร้างฐานข้อมูล midterm-results-db (รันบน remote ไปแล้วผ่าน MCP — ไฟล์นี้ไว้สำหรับ local dev / สร้างใหม่)
CREATE TABLE IF NOT EXISTS students (
  citizen_id TEXT PRIMARY KEY,
  student_id TEXT,
  prefix TEXT,
  first_name TEXT NOT NULL,
  last_name TEXT NOT NULL,
  class TEXT,
  class_no INTEGER
);

CREATE TABLE IF NOT EXISTS grades (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  citizen_id TEXT NOT NULL,
  subject_code TEXT,
  subject_name TEXT NOT NULL,
  credits REAL,
  midterm_score REAL,
  max_score REAL DEFAULT 100,
  teacher TEXT
);

CREATE INDEX IF NOT EXISTS idx_grades_cid ON grades(citizen_id);

CREATE TABLE IF NOT EXISTS settings (
  key TEXT PRIMARY KEY,
  value TEXT
);

INSERT OR REPLACE INTO settings (key, value) VALUES
  ('term_label', 'ผลการเรียนกลางภาค ภาคเรียนที่ 1 ปีการศึกษา 2569'),
  ('announce_open', '1');

-- ข้อมูลตัวอย่างสำหรับทดสอบ
INSERT OR REPLACE INTO students (citizen_id, student_id, prefix, first_name, last_name, class, class_no) VALUES
  ('1234567890121', '12345', 'เด็กชาย', 'สมชาย', 'ใจดี', 'ม.1/1', 5),
  ('1100700256030', '12346', 'เด็กหญิง', 'สมหญิง', 'เรียนเก่ง', 'ม.1/1', 6);

INSERT INTO grades (citizen_id, subject_code, subject_name, credits, midterm_score, max_score, teacher) VALUES
  ('1234567890121', 'ท21101', 'ภาษาไทย', 1.5, 42, 50, 'ครูสมศรี'),
  ('1234567890121', 'ค21101', 'คณิตศาสตร์', 1.5, 35, 50, 'ครูสมปอง'),
  ('1234567890121', 'ว21101', 'วิทยาศาสตร์', 1.5, 38, 50, 'ครูวิทยา'),
  ('1234567890121', 'ส21101', 'สังคมศึกษา', 1.5, 44, 50, 'ครูสังคม'),
  ('1234567890121', 'อ21101', 'ภาษาอังกฤษ', 1.5, 30, 50, 'ครูแอน'),
  ('1100700256030', 'ท21101', 'ภาษาไทย', 1.5, 48, 50, 'ครูสมศรี'),
  ('1100700256030', 'ค21101', 'คณิตศาสตร์', 1.5, 46, 50, 'ครูสมปอง'),
  ('1100700256030', 'ว21101', 'วิทยาศาสตร์', 1.5, 45, 50, 'ครูวิทยา'),
  ('1100700256030', 'ส21101', 'สังคมศึกษา', 1.5, 49, 50, 'ครูสังคม'),
  ('1100700256030', 'อ21101', 'ภาษาอังกฤษ', 1.5, 47, 50, 'ครูแอน');
