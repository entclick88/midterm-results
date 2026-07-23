# -*- coding: utf-8 -*-
"""
แปลงไฟล์คะแนน (โครงสร้าง "ไฟล์นำส่งข้อมูลเข้าเว็บ") → import.sql สำหรับ D1

วิธีใช้:
    python convert_xlsx.py "ไฟล์คะแนน.xlsx"

รองรับ 2 โครงสร้าง (ตรวจหัวตารางอัตโนมัติ):
  แบบใหม่ (1-69): แถว 1 = รหัสนักเรียน | เลขประจำตัวประชาชน | ห้อง | เลขที่ | ชื่อ-สกุล | รหัสวิชา...
  แบบเก่า (2-68): แถว 3 = ห้อง | เลขที่ | เลขประจำตัวประชาชน | ชื่อ-สกุล | รหัสวิชา...
แต่ละวิชากว้าง 4 คอลัมน์: คะแนนเต็ม, คะแนนที่ได้, งานค้าง(ชิ้น), %เข้าเรียน

เลขบัตร ปชช. รองรับรูปแบบมีขีด (1-8499-02472-20-7) — ต้องกรอกครบทุกคน
กติกา: ไม่มี %เวลาเรียน = นักเรียนไม่ได้ลงเรียนวิชานั้น → ตัดวิชาออก
ถ้าเลขบัตรไม่ครบ/ซ้ำ จะรันแบบตรวจสอบ (dry-run) รายงานอย่างเดียว ไม่สร้าง SQL
"""
import openpyxl
import re
import sys
import os
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")

TERM_LABEL = "ภาคเรียนที่ 1 ปีการศึกษา 2569"
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
    """คืนเลข 13 หลักถ้าค่านี้เป็นเลขบัตรประชาชน ไม่ใช่คืน None (รองรับ float + ขีดคั่น)"""
    if v is None:
        return None
    if isinstance(v, float) and v == int(v):
        v = int(v)
    s = re.sub(r"[\s\-]", "", str(v))
    return s if re.fullmatch(r"\d{13}", s) else None


def id_str(v):
    """1.0 → '1', '2955.0' → '2955', เว้นว่าง → None"""
    if v is None:
        return None
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s or None


def parse_sheet(ws):
    rows = list(ws.iter_rows(values_only=True))
    grade_level = "ม." + ws.title.lstrip("Mm")  # M1 → ม.1

    # หาแถวหัวตาราง (มี 'ห้อง' หรือ 'ชื่อ...สกุล') — รองรับไฟล์เก่า (แถว 3) และใหม่ (แถว 1)
    header_idx = None
    for i, r in enumerate(rows[:8]):
        cells = [str(c).strip() if c is not None else "" for c in r]
        if "ห้อง" in cells or any("ชื่อ" in c and "สกุล" in c for c in cells):
            header_idx = i
            break
    if header_idx is None:
        header_idx = 2
    code_row = rows[header_idx]
    name_row = rows[header_idx + 1]
    cred_row = rows[header_idx + 2]

    # ระบุคอลัมน์พื้นฐานจากป้ายหัวตาราง + คอลัมน์วิชาแรก
    col_sid = col_cid = col_room = col_no = col_name = None
    first_subj = None
    for c, v in enumerate(code_row):
        if v is None or str(v).strip() == "":
            continue
        h = str(v).strip()
        if re.search(r"รหัสนักเรียน|รหัสนร", h):
            col_sid = c
        elif re.search(r"บัตร|ปชช|ประชาชน", h):
            col_cid = c
        elif h == "ห้อง":
            col_room = c
        elif h == "เลขที่":
            col_no = c
        elif "ชื่อ" in h and "สกุล" in h:
            col_name = c
        elif first_subj is None:
            first_subj = c

    subjects = []
    for c in range(first_subj, len(code_row)):
        code = code_row[c]
        if code is not None and str(code).strip():
            subjects.append({
                "col": c,
                "code": str(code).strip(),
                "name": str(name_row[c]).strip() if name_row[c] else str(code).strip(),
                "credits": cred_row[c],
            })

    students, grades, problems = [], [], []
    for i in range(header_idx + 3, len(rows)):
        r = rows[i]
        rowno = i + 1
        name_val = r[col_name] if col_name is not None else None
        if name_val is None or not re.search(r"[ก-๙]", str(name_val)):
            continue  # ข้ามแถวหัวย่อย (คะแนนเก็บ/เต็ม/ได้) และแถวว่าง
        name_val = re.sub(r"\s+", " ", str(name_val)).strip()
        if any(w in name_val for w in SKIP_NAME_WORDS):
            continue  # แถวสรุป/หมายเหตุ ไม่ใช่นักเรียน

        room = id_str(r[col_room]) if col_room is not None else ""
        class_no = num(r[col_no]) if col_no is not None else "NULL"
        cid = cid_of(r[col_cid]) if col_cid is not None else None
        student_id = id_str(r[col_sid]) if col_sid is not None else None
        if col_cid is not None and cid is None:
            problems.append(f"{ws.title} แถว {rowno}: {name_val} — เลขบัตรไม่ถูกต้อง ({r[col_cid]})")

        prefix, first, last = split_prefix(name_val)
        students.append({
            "cid": cid, "student_id": student_id,
            "prefix": prefix, "first": first, "last": last,
            "class": f"{grade_level}/{room}", "class_no": class_no,
            "sheet": ws.title, "row": rowno, "name": name_val,
        })

        for s in subjects:
            c = s["col"]
            # "-" (ขีด) และช่องว่าง = ไม่มีข้อมูล → ถือเป็น None
            nb = lambda v: None if (v is None or (isinstance(v, str) and v.strip() in ("", "-"))) else v
            full, got, pend, att = nb(r[c]), nb(r[c + 1]), nb(r[c + 2]), nb(r[c + 3])
            # ไม่มี %เวลาเรียน = นักเรียนไม่ได้ลงเรียนวิชานี้ → ตัดออก (ต่อให้มีคะแนน 0 ค้างอยู่)
            if att is None or got is None or full is None:
                continue
            grades.append({
                "cid": cid, "code": s["code"], "name": s["name"], "credits": s["credits"],
                "score": got, "max": full, "pend": pend, "att": att, "seq": s["col"],
            })

    id_msg = chr(65 + col_cid) if col_cid is not None else None
    return subjects, students, grades, problems, id_msg


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    path = sys.argv[1]
    wb = openpyxl.load_workbook(path, data_only=True)

    # --only M5,M6  = นำเข้าเฉพาะชั้นที่เลือก (ไม่แตะชั้นอื่น) — โหมดแก้บางส่วน
    only = None
    if "--only" in sys.argv:
        idx = sys.argv.index("--only")
        only = {s.strip().upper() for s in sys.argv[idx + 1].split(",") if s.strip()}
    partial = only is not None

    all_students, all_grades, all_problems = [], [], []
    print(f"อ่านไฟล์: {os.path.basename(path)}" + (f"  [โหมดบางส่วน: {', '.join(sorted(only))}]" if partial else ""))
    for ws in wb.worksheets:
        if only is not None and ws.title.upper() not in only:
            continue
        subjects, students, grades, problems, id_col = parse_sheet(ws)
        all_students += students
        all_grades += grades
        all_problems += problems
        id_msg = f"เลขบัตรคอลัมน์ {id_col}" if id_col else "*** ไม่พบคอลัมน์เลขบัตร ***"
        has_sid = "มีรหัสนักเรียน" if students and students[0]["student_id"] else "ไม่มีรหัสนักเรียน"
        print(f"  {ws.title}: นักเรียน {len(students)} คน, {len(subjects)} วิชา, {len(grades)} รายการคะแนน | {id_msg} | {has_sid}")

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

    if not all_students:
        print("\n*** ไม่พบข้อมูลในชั้นที่เลือก — ตรวจชื่อชีต (M1..M6) อีกครั้ง ***")
        sys.exit(2)

    # สร้าง SQL
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "import.sql")
    lines = [f"-- สร้างโดย convert_xlsx.py เมื่อ {datetime.now().isoformat()} จาก {os.path.basename(path)}"]

    if partial:
        # โหมดบางส่วน: ลบเฉพาะชั้นที่นำเข้าใหม่ แล้วใส่ทับ — ไม่แตะชั้นอื่นและตาราง settings
        levels = sorted({s["class"].split("/")[0] for s in all_students})
        lines.append(f"-- โหมดแก้บางส่วน: แทนที่เฉพาะ {', '.join(levels)}")
        for lv in levels:
            pat = q(lv + "/%")
            lines.append(f"DELETE FROM grades WHERE citizen_id IN (SELECT citizen_id FROM students WHERE class LIKE {pat});")
            lines.append(f"DELETE FROM students WHERE class LIKE {pat};")
    else:
        # โหมดเต็ม: ล้างและสร้างใหม่ทั้งหมด
        lines += [
            "DROP TABLE IF EXISTS grades;",
            "DROP TABLE IF EXISTS students;",
            """CREATE TABLE students (
  citizen_id TEXT PRIMARY KEY, student_id TEXT, prefix TEXT,
  first_name TEXT NOT NULL, last_name TEXT NOT NULL, class TEXT, class_no INTEGER);""",
            """CREATE TABLE grades (
  id INTEGER PRIMARY KEY AUTOINCREMENT, citizen_id TEXT NOT NULL,
  subject_code TEXT, subject_name TEXT NOT NULL, credits REAL,
  midterm_score REAL, max_score REAL, pending_work REAL, attendance REAL, teacher TEXT, seq INTEGER);""",
            "CREATE INDEX idx_grades_cid ON grades(citizen_id);",
            f"INSERT OR REPLACE INTO settings (key, value) VALUES ('term_label', {q(TERM_LABEL)}), ('announce_open', '1');",
        ]

    CHUNK = 200
    st_vals = [f"({q(s['cid'])}, {q(s['student_id'])}, {q(s['prefix'])}, {q(s['first'])}, {q(s['last'])}, {q(s['class'])}, {s['class_no']})" for s in all_students]
    for i in range(0, len(st_vals), CHUNK):
        lines.append("INSERT INTO students (citizen_id, student_id, prefix, first_name, last_name, class, class_no) VALUES\n" + ",\n".join(st_vals[i:i + CHUNK]) + ";")

    gr_vals = [f"({q(g['cid'])}, {q(g['code'])}, {q(g['name'])}, {num(g['credits'])}, {num(g['score'])}, {num(g['max'])}, {num(g['pend'])}, {num(g['att'])}, NULL, {num(g['seq'])})" for g in all_grades]
    for i in range(0, len(gr_vals), CHUNK):
        lines.append("INSERT INTO grades (citizen_id, subject_code, subject_name, credits, midterm_score, max_score, pending_work, attendance, teacher, seq) VALUES\n" + ",\n".join(gr_vals[i:i + CHUNK]) + ";")

    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"\n✔ สร้าง {out} แล้ว")
    print("นำเข้าจริง: cd ../worker && npx wrangler d1 execute midterm-results-db --remote --file=../import/import.sql")


if __name__ == "__main__":
    main()
