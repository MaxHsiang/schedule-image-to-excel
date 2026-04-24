# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import tempfile
from pathlib import Path
from urllib.parse import quote

import uvicorn
from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from schedule_core import default_output_path, run_conversion


app = FastAPI(title="班表圖片轉 Excel")


HTML_PAGE = """<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>班表圖片轉 Excel</title>
  <style>
    :root {
      --bg1: #f3fbff;
      --bg2: #fff7ea;
      --card: rgba(255,255,255,.92);
      --line: #d9e7e1;
      --text: #17322a;
      --muted: #567168;
      --accent: #2f6b53;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Microsoft JhengHei", "PingFang TC", sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at top left, #dff4ff 0, transparent 28%),
        radial-gradient(circle at bottom right, #ffe7b8 0, transparent 24%),
        linear-gradient(135deg, var(--bg1), var(--bg2));
      min-height: 100vh;
      display: grid;
      place-items: center;
      padding: 24px;
    }
    .card {
      width: min(720px, 100%);
      background: var(--card);
      border: 1px solid rgba(255,255,255,.7);
      box-shadow: 0 18px 50px rgba(34, 71, 58, .12);
      border-radius: 24px;
      padding: 28px;
    }
    h1 { margin: 0 0 10px; font-size: 32px; }
    p { color: var(--muted); margin: 0 0 20px; line-height: 1.6; }
    .grid { display: grid; gap: 16px; }
    label { display: grid; gap: 8px; font-weight: 700; }
    input {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px 16px;
      font: inherit;
      background: rgba(255,255,255,.96);
    }
    .actions {
      display: flex;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
      margin-top: 10px;
    }
    button {
      border: 0;
      border-radius: 999px;
      padding: 14px 24px;
      font: inherit;
      font-weight: 700;
      color: #fff;
      background: linear-gradient(135deg, #2f6b53, #3e8a6c);
      cursor: pointer;
    }
    .hint {
      color: var(--muted);
      font-size: 14px;
    }
    .status {
      min-height: 24px;
      margin-top: 14px;
      font-weight: 700;
    }
    .error { color: #b42318; }
    .ok { color: #2f6b53; }
  </style>
</head>
<body>
  <main class="card">
    <h1>班表圖片轉 Excel</h1>
    <p>上傳班表圖片後，系統會自動抓出指定員工的個人班表，並下載成 Excel。</p>

    <div class="grid">
      <label>
        上傳班表圖片
        <input id="imageInput" type="file" accept=".png,.jpg,.jpeg,.bmp">
      </label>

      <label>
        姓名
        <input id="nameInput" type="text" value="張盈慧" placeholder="請輸入要擷取的姓名">
      </label>
    </div>

    <div class="actions">
      <button id="submitBtn" type="button">生成 Excel</button>
      <span class="hint">輸出欄位：年 / 月 / 日 / 周幾 / 班別 / 時數 / 時薪 / 單量 / 總薪水</span>
    </div>

    <div id="status" class="status"></div>
  </main>

  <script>
    const submitBtn = document.getElementById("submitBtn");
    const imageInput = document.getElementById("imageInput");
    const nameInput = document.getElementById("nameInput");
    const status = document.getElementById("status");

    submitBtn.addEventListener("click", async () => {
      const file = imageInput.files[0];
      const employeeName = (nameInput.value || "張盈慧").trim();

      if (!file) {
        status.className = "status error";
        status.textContent = "請先選擇班表圖片。";
        return;
      }

      status.className = "status";
      status.textContent = "正在辨識並生成 Excel，請稍候...";

      try {
        const response = await fetch(`/convert?name=${encodeURIComponent(employeeName)}`, {
          method: "POST",
          headers: {
            "Content-Type": file.type || "application/octet-stream",
            "X-Filename": file.name,
          },
          body: await file.arrayBuffer(),
        });

        if (!response.ok) {
          const data = await response.json().catch(() => ({ detail: "轉換失敗" }));
          throw new Error(data.detail || "轉換失敗");
        }

        const blob = await response.blob();
        const disposition = response.headers.get("Content-Disposition") || "";
        const match = disposition.match(/filename\\*=UTF-8''([^;]+)/);
        const fileName = match ? decodeURIComponent(match[1]) : "personal_schedule.xlsx";

        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);

        status.className = "status ok";
        status.textContent = "Excel 已生成並開始下載。";
      } catch (error) {
        status.className = "status error";
        status.textContent = error.message || "轉換失敗。";
      }
    });
  </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    return HTML_PAGE


@app.post("/convert")
async def convert(request: Request, name: str = Query(default="張盈慧")):
    raw = await request.body()
    if not raw:
        raise HTTPException(status_code=400, detail="沒有收到圖片內容。")

    filename_header = request.headers.get("X-Filename", "schedule.jpg")
    safe_original_name = Path(filename_header).name or "schedule.jpg"
    suffix = Path(safe_original_name).suffix or ".jpg"

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / f"upload{suffix}"
            output_path = default_output_path(input_path, name)
            input_path.write_bytes(raw)
            saved_path, _ = run_conversion(input_path, name, output_path)
            excel_bytes = saved_path.read_bytes()
            download_name = quote(saved_path.name)
            return Response(
                content=excel_bytes,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={
                    "Content-Disposition": f"attachment; filename*=UTF-8''{download_name}",
                },
            )
    except HTTPException:
        raise
    except Exception as exc:
        return JSONResponse(status_code=400, content={"detail": str(exc)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run("schedule_web:app", host="0.0.0.0", port=port, reload=False)
