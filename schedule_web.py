# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import tempfile
import traceback
from pathlib import Path
from urllib.parse import quote

import uvicorn
from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import HTMLResponse, JSONResponse, Response

from schedule_core import records_to_dicts, run_conversion_debug
from schedule_excel_parser import run_excel_conversion_debug
from schedule_text_parser import run_text_conversion_debug


APP_VERSION = "excel-text-v2-image-beta"


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
      --card: rgba(255, 255, 255, .92);
      --line: #d9e7e1;
      --text: #17322a;
      --muted: #567168;
      --accent: #2f6b53;
      --accent2: #f2a63b;
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
      width: min(980px, 100%);
      background: var(--card);
      border: 1px solid rgba(255, 255, 255, .7);
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
      background: rgba(255, 255, 255, .96);
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
    .secondary {
      background: linear-gradient(135deg, #f5b14e, #eb9730);
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
    .preview {
      margin-top: 24px;
      padding: 18px;
      border: 1px solid var(--line);
      border-radius: 18px;
      background: rgba(255, 255, 255, .78);
    }
    .preview[hidden] { display: none; }
    .preview h2 {
      margin: 0 0 8px;
      font-size: 20px;
    }
    .count {
      margin: 0 0 14px;
      color: var(--muted);
      font-weight: 700;
    }
    .preview-list {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 10px;
    }
    .preview-item {
      border-radius: 14px;
      padding: 10px 12px;
      font-size: 15px;
      font-weight: 700;
      background: #f6fbf8;
      border: 1px solid #d9e7e1;
    }
    .download-row {
      margin-top: 16px;
      display: flex;
      gap: 12px;
      flex-wrap: wrap;
    }
    .debug {
      margin-top: 12px;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.7;
    }
    textarea {
      width: 100%;
      min-height: 220px;
      border: 1px solid var(--line);
      border-radius: 14px;
      padding: 14px 16px;
      font: inherit;
      line-height: 1.55;
      background: rgba(255, 255, 255, .96);
      resize: vertical;
    }
    .table-preview {
      margin-top: 16px;
      overflow-x: auto;
      border: 1px solid var(--line);
      border-radius: 14px;
      background: rgba(255,255,255,.86);
    }
    .table-preview table {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }
    .table-preview td {
      border: 1px solid #d9e7e1;
      padding: 8px 10px;
      text-align: center;
      white-space: nowrap;
    }
  </style>
</head>
<body>
  <main class="card">
    <h1>班表圖片轉 Excel</h1>
    <p>目前最穩定的方式是上傳 Excel，或直接貼上從 Excel 複製的班表文字。圖片分析功能仍為測試版，筆數可能不完全準確。</p>

    <div class="grid">
      <label>
        上傳班表檔案
        <span class="hint">支援圖片或 Excel。圖片分析目前為測試版，建議優先使用 Excel。</span>
        <input id="imageInput" type="file" accept=".png,.jpg,.jpeg,.bmp,.xlsx">
      </label>

      <label>
        姓名
        <input id="nameInput" type="text" value="張盈慧" placeholder="請輸入員工姓名">
      </label>

      <label>
        或直接貼上班表文字
        <textarea id="textInput" placeholder="可直接貼上從 Excel 複製出來的班表文字，系統會先編排成表格，再抓出指定員工班表。"></textarea>
      </label>
    </div>

    <div class="actions">
      <button id="previewBtn" type="button">先預覽結果</button>
      <button id="submitBtn" type="button" class="secondary">直接下載 Excel</button>
      <span class="hint">輸出欄位：年 / 月 / 日 / 周幾 / 班別 / 時數 / 時薪 / 單量 / 總薪水</span>
    </div>

    <div id="status" class="status"></div>

    <section id="previewBox" class="preview" hidden>
      <h2>辨識結果預覽</h2>
      <p id="previewCount" class="count"></p>
      <div id="previewList" class="preview-list"></div>
      <div id="debugInfo" class="debug"></div>
      <div id="tablePreview" class="table-preview" hidden></div>
      <div class="download-row">
        <button id="downloadBtn" type="button" class="secondary">確認並下載 Excel</button>
      </div>
    </section>
  </main>

  <script>
    const imageInput = document.getElementById("imageInput");
    const nameInput = document.getElementById("nameInput");
    const status = document.getElementById("status");
    const previewBtn = document.getElementById("previewBtn");
    const submitBtn = document.getElementById("submitBtn");
    const downloadBtn = document.getElementById("downloadBtn");
    const previewBox = document.getElementById("previewBox");
    const previewCount = document.getElementById("previewCount");
    const previewList = document.getElementById("previewList");
    const debugInfo = document.getElementById("debugInfo");
    const textInput = document.getElementById("textInput");
    const tablePreview = document.getElementById("tablePreview");

    async function sendRequest(mode) {
      const file = imageInput.files[0];
      const employeeName = (nameInput.value || "張盈慧").trim();
      const pastedText = (textInput.value || "").trim();

      if (!file && !pastedText) {
        status.className = "status error";
        status.textContent = "請先選擇班表檔案，或貼上班表文字。";
        return null;
      }

      status.className = "status";
      status.textContent = mode === "preview" ? "正在預覽辨識結果，請稍候..." : "正在產生 Excel，請稍候...";

      const headers = {};
      let body;

      if (pastedText) {
        headers["Content-Type"] = "text/plain;charset=utf-8";
        headers["X-Input-Mode"] = "pasted-text";
        headers["X-Filename"] = "pasted_schedule.txt";
        body = pastedText;
      } else {
        headers["Content-Type"] = file.type || "application/octet-stream";
        headers["X-Filename"] = encodeURIComponent(file.name);
        body = await file.arrayBuffer();
      }

      const response = await fetch(`/convert?name=${encodeURIComponent(employeeName)}&mode=${mode}`, {
        method: "POST",
        headers,
        body,
      });

      if (!response.ok) {
        const data = await response.json().catch(() => ({ detail: "轉換失敗" }));
        throw new Error(data.detail || "轉換失敗");
      }

      return response;
    }

    async function downloadExcel() {
      try {
        const response = await sendRequest("download");
        if (!response) return;

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
        status.textContent = "Excel 已下載完成。";
      } catch (error) {
        status.className = "status error";
        status.textContent = error.message || "轉換失敗";
      }
    }

    previewBtn.addEventListener("click", async () => {
      try {
        const response = await sendRequest("preview");
        if (!response) return;
        const data = await response.json();

        previewCount.textContent = `共抓到 ${data.count} 筆班次`;
        previewList.innerHTML = "";
        for (const row of data.records) {
          const item = document.createElement("div");
          item.className = "preview-item";
          item.textContent = `${row.month}/${row.day}（${row.weekday}）${row.shift}班 ${row.hours}小時`;
          previewList.appendChild(item);
        }

        const debug = data.debug || {};
        debugInfo.textContent = `版本：${data.version} ｜ OCR先抓到：${debug.ocr_direct_matches ?? "-"} ｜ 顏色補抓：${debug.color_fallback_matches ?? "-"} ｜ 總筆數：${debug.deduped_total ?? data.count}`;

        if (Array.isArray(data.table_rows) && data.table_rows.length) {
          const tableHtml = [
            "<table><tbody>",
            ...data.table_rows.map(
              (row) =>
                "<tr>" +
                row.map((cell) => `<td>${String(cell ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")}</td>`).join("") +
                "</tr>"
            ),
            "</tbody></table>",
          ].join("");
          tablePreview.innerHTML = tableHtml;
          tablePreview.hidden = false;
        } else {
          tablePreview.innerHTML = "";
          tablePreview.hidden = true;
        }

        previewBox.hidden = false;
        status.className = "status ok";
        status.textContent = "已完成預覽，請確認班次內容。";
      } catch (error) {
        previewBox.hidden = true;
        status.className = "status error";
        status.textContent = error.message || "預覽失敗";
      }
    });

    submitBtn.addEventListener("click", downloadExcel);
    downloadBtn.addEventListener("click", downloadExcel);
  </script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    return HTML_PAGE


@app.post("/convert")
async def convert(
    request: Request,
    name: str = Query(default="張盈慧"),
    mode: str = Query(default="download"),
):
    raw = await request.body()
    if not raw:
        raise HTTPException(status_code=400, detail="沒有收到圖片內容。")

    input_mode = request.headers.get("X-Input-Mode", "file")
    filename_header = request.headers.get("X-Filename", "schedule.jpg")
    safe_original_name = Path(request.headers.get("X-Filename", "schedule.jpg")).name or "schedule.jpg"
    try:
        from urllib.parse import unquote

        safe_original_name = Path(unquote(filename_header)).name or safe_original_name
    except Exception:
        pass
    suffix = Path(safe_original_name).suffix or ".jpg"

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            input_path = temp_path / f"upload{suffix}"
            if input_mode == "pasted-text":
                text = raw.decode("utf-8")
                saved_path, records, debug = run_text_conversion_debug(text, name, None)
            else:
                input_path.write_bytes(raw)
                if suffix.lower() == ".xlsx":
                    saved_path, records, debug = run_excel_conversion_debug(input_path, name, None)
                else:
                    saved_path, records, debug = run_conversion_debug(input_path, name, None)

            if mode == "preview":
                payload = {
                    "count": len(records),
                    "records": records_to_dicts(records),
                    "debug": debug,
                    "version": APP_VERSION,
                }
                if isinstance(debug, dict) and "table_rows" in debug:
                    payload["table_rows"] = debug["table_rows"]
                return JSONResponse(payload)

            excel_bytes = saved_path.read_bytes()
            download_name = quote(saved_path.name)
            return Response(
                content=excel_bytes,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": f"attachment; filename*=UTF-8''{download_name}"},
            )
    except HTTPException:
        raise
    except Exception as exc:
        print("convert failed:", repr(exc), flush=True)
        traceback.print_exc()
        return JSONResponse(status_code=400, content={"detail": str(exc)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run("schedule_web:app", host="0.0.0.0", port=port, reload=False)
