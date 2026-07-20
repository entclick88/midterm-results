-- โครงสร้างฐานข้อมูล midterm-results-db — รันไฟล์นี้ = ล้างและสร้างใหม่ทั้งหมด (พร้อมข้อมูลตัวอย่าง)
DROP TABLE IF EXISTS grades;
DROP TABLE IF EXISTS students;
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
  midterm_score REAL,        -- NULL = ครูยังไม่ส่งคะแนน (แสดง "รอคะแนน")
  max_score REAL DEFAULT 100,
  pending_work REAL,         -- งานค้าง (ชิ้น)
  attendance REAL,           -- สัดส่วนการเข้าเรียน 0-1
  teacher TEXT
);

CREATE INDEX IF NOT EXISTS idx_grades_cid ON grades(citizen_id);

CREATE TABLE IF NOT EXISTS settings (
  key TEXT PRIMARY KEY,
  value TEXT
);

INSERT OR REPLACE INTO settings (key, value) VALUES
  ('term_label', 'ผลการเรียนกลางภาค ภาคเรียนที่ 2 ปีการศึกษา 2568'),
  ('announce_open', '1');

-- ข้อมูลตัวอย่างสำหรับทดสอบ
INSERT OR REPLACE INTO students (citizen_id, student_id, prefix, first_name, last_name, class, class_no) VALUES
  ('1234567890121', '12345', 'เด็กชาย', 'สมชาย', 'ใจดี', 'ม.1/1', 5),
  ('1100700256030', '12346', 'เด็กหญิง', 'สมหญิง', 'เรียนเก่ง', 'ม.1/1', 6);

INSERT INTO grades (citizen_id, subject_code, subject_name, credits, midterm_score, max_score, pending_work, attendance, teacher) VALUES
  ('1234567890121', 'ท21102', 'ภาษาไทย 2', 1.5, 24.5, 25, 0, 0.8, 'ครูสมศรี'),
  ('1234567890121', 'ค21102', 'คณิตศาสตร์ 2', 1.5, 15.3, 20, 2, 0.9, 'ครูสมปอง'),
  ('1234567890121', 'ว21102', 'วิทยาศาสตร์ 2', 1.5, 20.5, 25, 0, 1, 'ครูวิทยา'),
  ('1234567890121', 'ส21103', 'สังคมศึกษา 2', 1.5, 22, 25, 0, 0.93, 'ครูสังคม'),
  ('1234567890121', 'อ21102', 'ภาษาอังกฤษ 2', 1.5, NULL, 30, 1, 0.87, 'ครูแอน'),
  ('1100700256030', 'ท21102', 'ภาษาไทย 2', 1.5, 24, 25, 0, 1, 'ครูสมศรี'),
  ('1100700256030', 'ค21102', 'คณิตศาสตร์ 2', 1.5, 18.5, 20, 0, 1, 'ครูสมปอง'),
  ('1100700256030', 'ว21102', 'วิทยาศาสตร์ 2', 1.5, 23, 25, 0, 1, 'ครูวิทยา'),
  ('1100700256030', 'ส21103', 'สังคมศึกษา 2', 1.5, 24.5, 25, 0, 1, 'ครูสังคม'),
  ('1100700256030', 'อ21102', 'ภาษาอังกฤษ 2', 1.5, 28, 30, 0, 0.97, 'ครูแอน');
