// midterm-results-api — API ค้นหาผลการเรียนกลางภาคด้วยเลขบัตรประชาชน 13 หลัก

function corsHeaders(env) {
  return {
    "Access-Control-Allow-Origin": env.ALLOWED_ORIGIN || "*",
    "Access-Control-Allow-Methods": "GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

function json(data, status, env) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", ...corsHeaders(env) },
  });
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders(env) });
    }

    const url = new URL(request.url);

    if (request.method === "GET" && url.pathname === "/api/results") {
      const cid = (url.searchParams.get("cid") || "").trim();

      if (!/^\d{13}$/.test(cid)) {
        return json({ error: "กรุณากรอกเลขบัตรประชาชน 13 หลัก" }, 400, env);
      }

      const settings = await env.DB.prepare(
        "SELECT key, value FROM settings WHERE key IN ('term_label','announce_open')"
      ).all();
      const cfg = Object.fromEntries(settings.results.map((r) => [r.key, r.value]));

      if (cfg.announce_open !== "1") {
        return json({ error: "ขณะนี้ปิดการประกาศผลชั่วคราว" }, 403, env);
      }

      const student = await env.DB.prepare(
        "SELECT student_id, prefix, first_name, last_name, class, class_no FROM students WHERE citizen_id = ?"
      ).bind(cid).first();

      if (!student) {
        return json({ error: "ไม่พบข้อมูล กรุณาตรวจสอบเลขบัตรประชาชนอีกครั้ง" }, 404, env);
      }

      const grades = await env.DB.prepare(
        `SELECT subject_code, subject_name, credits, midterm_score, max_score, teacher
         FROM grades WHERE citizen_id = ? ORDER BY subject_code`
      ).bind(cid).all();

      return json({
        term_label: cfg.term_label || "ผลการเรียนกลางภาค",
        student,
        grades: grades.results,
      }, 200, env);
    }

    return json({ error: "Not found" }, 404, env);
  },
};
