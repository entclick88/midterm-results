# -*- coding: utf-8 -*-
"""
แปลงไฟล์คะแนนกลางภาค (โครงสร้าง "ไฟล์นำส่งข้อมูลเข้าเว็บ") → import.sql สำหรับ D1

วิธีใช้:
    python convert_xlsx.py "ไฟล์คะแนน.xlsx"

โครงสร้างที่คาดหวังต่อชีต (M1..M6):
    แถว 1: ลงชื่อหัวหน้าหมวด | แถว 2: ลงชื่อผู้สอน | แถว 3: ห้อง เลขที่ ชื่อ-สกุล [เลขบัตร] + รหัสวิชา
    แถว 4: ชื่อวิชา | แถว 5: หน่วยกิต | แถว 6-7: หัวคอลัมน์ย่อย (เต็ม/ได้/ชิ้น/%)
    แถว 8+: ข้อมูลนักเรียน — วิชาละ 4 คอลัมน์: เต็ม, ได้, งานค้าง, %เข้าเรียน

คอลัมน์เลขบัตรประชาชน: ตรวจหาอัตโนมัติจากหัวตาราง (มีคำว่า บัตร/ปชช/ประชาชน)
หรือจากข้อมูลที่เป็นตัวเลข 13 หลัก — แทรกไว้คอลัมน์ไหนก็ได้ก่อนรหัสวิชาแรก
ถ้าไม่พบ จะรันแบบตรวจสอบ (dry-run) รายงานสรุปอย่างเดียว ไม่สร้าง SQL
"""
import openpyxl
import re
import sys
import os
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

TERM_LABEL = "ผลการเรียนกลางภาค ภาคเรียนที่ 2 ปีการศึกษา 2568"
PREFIXES = ["เด็กชาย", "เด็กหญิง", "นางสาว", "นาย", "นาง", "ด.ช.", "ด.ญ.", "น.ส."]
SKIP_NAME_WORDS = ["คะแนน", "เกณฑ์", "ลงชื่อ", "หมายเหตุ", "รวม"]


def q(v):
    if v is None or v == "":
        return "NULL"
    return "'" + str(v).replace("'", "''") + "'"


def num(v):
    if v is None or v == "":
        return "NULL"
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else str(round(f, 4))
    except (TypeError, ValueError):
        return "NULL"


def clean_teacher(v):
    """ตัดวันที่/วงเล็บท้ายชื่อครูออก เช่น 'ศิวพงศ์ สุชาติ (09/01/2026)' → 'ศิวพงศ์ สุชาติ'
    ช่องลงชื่อที่เป็นสถานะ (เช่น 'ยังไม่เรียบร้อย') ไม่ใช่ชื่อครู → คืน None"""
    if not v:
        return None
    s = re.sub(r"\(.*?\)", " ", str(v))
    s = re.sub(r"[\d/.\-]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    if not s or re.search(r"ยังไม่|ไม่เรียบร้อย|รอ(ส่ง|ตรวจ)|ค้าง", s):
        return None
    return s


def split_prefix(fullname):
    name = re.sub(r"\s+", " ", str(fullname)).strip()
    for p in PREFIXES:
        if name.startswith(p):
            rest = name[len(p):].strip()
            parts = rest.split(" ", 1)
            return p, parts[0], (parts[1] if len(parts) > 1 else "")
    parts = name.split(" ", 1)
    return "", parts[0], (parts[1] if len(parts) > 1 else "")


def cid_of(v):
    """คืนเลข 13 หลักถ้าค่านี้เป็นเลขบัตรประชาชน ไม่ใช่คืน None (รองรับ float จาก Excel)"""
    if v is None:
        return None
    if isinstance(v, float) and v == int(v):
        v = int(v)
    s = re.sub(r"[\s\-]", "", str(v))
    return s if re.fullmatch(r"\d{13}", s) else None


def parse_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    grade_level = "ม." + ws.title.lstrip("Mm")  # M1 → ม.1
    code_row, name_row, cred_row, teacher_row = rows[2], rows[3], rows[4], rows[1]

    # หาคอลัมน์รหัสวิชา (แถว 3 มีค่า และไม่ใช่ 3 คอลัมน์แรกๆ ที่เป็น ห้อง/เลขที่/ชื่อ/เลขบัตร)
    base_headers = {"ห้อง", "เลขที่", "ชื่อ - สกุล", "ชื่อ-สกุล", "ชื่อ – สกุล"}
    subj_cols, id_col_by_header = [], None
    for c, v in enumerate(code_row):
        if v is None or str(v).strip() == "":
            continue
        h = str(v).strip()
        if h in base_headers:
            continue
        if re.search(r"บัตร|ปชช|ประชาชน", h):
            id_col_by_header = c
            continue
        subj_cols.append(c)
    first_subj = subj_cols[0]

    subjects = []
    for c in subj_cols:
        subjects.append({
            "col": c,
            "code": str(code_row[c]).strip(),
            "name": str(name_row[c]).strip() if name_row[c] else str(code_row[c]).strip(),
            "credits": cred_row[c],
            "teacher": clean_teacher(teacher_row[c] if c < len(teacher_row) else None),
        })

    # หา/ยืนยันคอลัมน์เลขบัตร: จากหัวตาราง หรือจากข้อมูลจริง (คอลัมน์ก่อนวิชาแรกที่เป็นเลข 13 หลัก)
    data_rows = []
    for i in range(7, len(rows)):
        r = rows[i]
        if r[0] is None or str(r[0]).strip() in ("", "ห้อง"):
            continue
        # หาชื่อ: คอลัมน์ 2 หรือ 3 (เผื่อแทรกเลขบัตรก่อนชื่อ)
        data_rows.append((i + 1, r))

    id_col = id_col_by_header
    if id_col is None:
        for c in range(0, first_subj):
            hits = sum(1 for _, r in data_rows[:30] if cid_of(r[c]))
            if hits >= max(3, len(data_rows[:30]) // 2):
                id_col = c
                break

    students, grades, problems = [], [], []
    for rowno, r in data_rows:
        # ชื่ออยู่คอลัมน์ไหน: คอลัมน์แรกในช่วง 2..first_subj ที่เป็นข้อความไทยยาวๆ และไม่ใช่เลขบัตร
        name_val = None
        for c in range(2, first_subj):
            if c == id_col:
                continue
            v = r[c]
            if v is not None and not cid_of(v) and re.search(r"[ก-๙]", str(v)):
                name_val = str(v).strip()
                break
        if not name_val:
            continue
        if any(w in name_val for w in SKIP_NAME_WORDS):
            continue  # แถวสรุป/หมายเหตุ ไม่ใช่นักเรียน

        room = str(r[0]).strip()
        room = room[:-2] if room.endswith(".0") else room
        class_no = num(r[1])
        cid = cid_of(r[id_col]) if id_col is not None else None
        if id_col is not None and cid is None:
            problems.append(f"{ws.title} แถว {rowno}: {name_val} — เลขบัตรไม่ถูกต้อง ({r[id_col]})")

        prefix, first, last = split_prefix(name_val)
        students.append({
            "cid": cid, "prefix": prefix, "first": first, "last": last,
            "class": f"{grade_level}/{room}", "class_no": class_no,
            "sheet": ws.title, "row": rowno, "name": name_val,
        })

        for s in subjects:
            c = s["col"]
            nb = lambda v: None if (v is None or (isinstance(v, str) and v.strip() == "")) else v
            full, got, pend, att = nb(r[c]), nb(r[c + 1]), nb(r[c + 2]), nb(r[c + 3])
            # ช่องคะแนน (ได้) ว่าง = นักเรียนไม่ได้ลงเรียนวิชานี้ → ตัดออก (0 คือคะแนนจริง ไม่ตัด)
            if got is None or full is None:
                continue
            grades.append({
                "cid": cid, "code": s["code"], "name": s["name"], "credits": s["credits"],
                "score": got, "max": full, "pend": pend, "att": att, "teacher": s["teacher"],
            })

    return subjects, students, grades, problems, id_col


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    path = sys.argv[1]
    wb = openpyxl.load_workbook(path, data_only=True)

    all_students, all_grades, all_problems = [], [], []
    print(f"อ่านไฟล์: {os.path.basename(path)}")
    for ws in wb.worksheets:
        subjects, students, grades, problems, id_col = parse_sheet(ws)
        all_students += students
        all_grades += grades
        all_problems += problems
        id_msg = f"เลขบัตรคอลัมน์ {chr(65 + id_col)}" if id_col is not None else "*** ไม่พบคอลัมน์เลขบัตร ***"
        print(f"  {ws.title}: นักเรียน {len(students)} คน, {len(subjects)} วิชา, {len(grades)} รายการคะแนน | {id_msg}")

    # ตรวจเลขบัตรซ้ำ
    seen = {}
    for st in all_students:
        if st["cid"]:
            if st["cid"] in seen:
                all_problems.append(f"เลขบัตรซ้ำ: {st['cid']} — {seen[st['cid']]} และ {st['sheet']} แถว {st['row']} ({st['name']})")
            else:
                seen[st["cid"]] = f"{st['sheet']} แถว {st['row']} ({st['name']})"

    missing = [st for st in all_students if not st["cid"]]
    print(f"\nรวม: นักเรียน {len(all_students)} คน, คะแนน {len(all_grades)} รายการ, ไม่มีเลขบัตร {len(missing)} คน")

    if all_problems:
        print("\n⚠ ปัญหาที่พบ:")
        for p in all_problems[:30]:
            print("  -", p)
        if len(all_problems) > 30:
            print(f"  ... และอีก {len(all_problems) - 30} รายการ")

    if missing or all_problems:
        print("\n*** ยังไม่สร้าง import.sql — แก้ไขไฟล์ Excel ให้ครบก่อน (dry-run) ***")
        sys.exit(2)

    # สร้าง SQL
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "import.sql")
    lines = [
        f"-- สร้างโดย convert_xlsx.py เมื่อ {datetime.now().isoformat()} จาก {os.path.basename(path)}",
        "DROP TABLE IF EXISTS grades;",
        "DROP TABLE IF EXISTS students;",
        """CREATE TABLE students (
  citizen_id TEXT PRIMARY KEY, student_id TEXT, prefix TEXT,
  first_name TEXT NOT NULL, last_name TEXT NOT NULL, class TEXT, class_no INTEGER);""",
        """CREATE TABLE grades (
  id INTEGER PRIMARY KEY AUTOINCREMENT, citizen_id TEXT NOT NULL,
  subject_code TEXT, subject_name TEXT NOT NULL, credits REAL,
  midterm_score REAL, max_score REAL, pending_work REAL, attendance REAL, teacher TEXT);""",
        "CREATE INDEX idx_grades_cid ON grades(citizen_id);",
        f"INSERT OR REPLACE INTO settings (key, value) VALUES ('term_label', {q(TERM_LABEL)}), ('announce_open', '1');",
    ]

    CHUNK = 200
    st_vals = [f"({q(s['cid'])}, NULL, {q(s['prefix'])}, {q(s['first'])}, {q(s['last'])}, {q(s['class'])}, {s['class_no']})" for s in all_students]
    for i in range(0, len(st_vals), CHUNK):
        lines.append("INSERT INTO students (citizen_id, student_id, prefix, first_name, last_name, class, class_no) VALUES\n" + ",\n".join(st_vals[i:i + CHUNK]) + ";")

    gr_vals = [f"({q(g['cid'])}, {q(g['code'])}, {q(g['name'])}, {num(g['credits'])}, {num(g['score'])}, {num(g['max'])}, {num(g['pend'])}, {num(g['att'])}, {q(g['teacher'])})" for g in all_grades]
    for i in range(0, len(gr_vals), CHUNK):
        lines.append("INSERT INTO grades (citizen_id, subject_code, subject_name, credits, midterm_score, max_score, pending_work, attendance, teacher) VALUES\n" + ",\n".join(gr_vals[i:i + CHUNK]) + ";")

    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"\n✔ สร้าง {out} แล้ว")
    print("นำเข้าจริง: cd ../worker && npx wrangler d1 execute midterm-results-db --remote --file=../import/import.sql")


if __name__ == "__main__":
    main()
