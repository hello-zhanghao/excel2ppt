import os
import sys
import json
import shutil
import tempfile
import zipfile
import threading
import time
import io
import queue
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string, Response

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from main import _run_pivot_mode, _run_ppt_mode, _detect_mode, __VERSION__

app = Flask(__name__, static_folder="static", static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

TEMP_ROOT = os.path.join(tempfile.gettempdir(), "excel2ppt_web")
if not os.path.exists(TEMP_ROOT):
    os.makedirs(TEMP_ROOT)

sessions = {}
sessions_lock = threading.Lock()


def _get_or_create_session(sid):
    with sessions_lock:
        if sid not in sessions:
            sessions[sid] = {
                "work_dir": tempfile.mkdtemp(prefix="job_", dir=TEMP_ROOT),
                "files": {},
                "log_queue": queue.Queue(),
                "status": "idle",
                "output": {},
            }
        return sessions[sid]


def _push_log(sid, message):
    s = _get_or_create_session(sid)
    s["log_queue"].put(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")


# ==================== 页面 ====================

@app.route("/")
def index():
    html_path = os.path.join(app.static_folder, "web", "index.html")
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            return render_template_string(f.read(), version=__VERSION__)
    return f"<h1>Excel 统一分析工具 v{__VERSION__}</h1><p>Web 界面文件不存在</p>"


# ==================== API ====================

@app.route("/api/upload", methods=["POST"])
def api_upload():
    sid = request.form.get("sid", "default")
    s = _get_or_create_session(sid)
    uploaded = []

    for key in request.files:
        file = request.files[key]
        if file.filename == "":
            continue
        safe_name = os.path.basename(file.filename)
        # 过滤掉无关文件（只保留 xlsx）
        if not safe_name.endswith(".xlsx"):
            continue
        # 跳过 Excel 临时文件
        if safe_name.startswith("~$"):
            continue
        dst = os.path.join(s["work_dir"], safe_name)
        file.save(dst)
        s["files"][safe_name] = dst
        uploaded.append(safe_name)

    _push_log(sid, f"上传完成: {', '.join(uploaded)}")

    # 检测配置和数据文件
    configs = [f for f in uploaded if "配置" in f or "config" in f.lower()]
    data_files = [f for f in uploaded if f not in configs]

    return jsonify({
        "ok": True,
        "work_dir": s["work_dir"],
        "configs": configs,
        "data_files": data_files,
        "uploaded": uploaded,
    })


@app.route("/api/analyze", methods=["POST"])
def api_analyze():
    sid = request.json.get("sid", "default")
    s = _get_or_create_session(sid)

    if s["status"] == "running":
        return jsonify({"ok": False, "error": "已有任务在运行"})

    config_name = request.json.get("config", "")
    config_path = s["files"].get(config_name) or _find_config(s["work_dir"])
    if not config_path or not os.path.exists(config_path):
        return jsonify({"ok": False, "error": "未找到配置文件，请先上传"})

    s["status"] = "running"
    threading.Thread(target=_run_analysis_web, args=(sid, config_path), daemon=True).start()
    return jsonify({"ok": True, "message": "分析已启动"})


def _find_config(work_dir):
    """自动查找工作目录下的配置文件"""
    candidates = []
    for f in os.listdir(work_dir):
        if f.endswith(".xlsx") and not f.startswith("~$"):
            fp = os.path.join(work_dir, f)
            candidates.append((fp, f))
    # 优先选文件名含"配置"的
    for fp, name in candidates:
        if "配置" in name or "config" in name.lower():
            return fp
    return candidates[0][0] if candidates else None


def _run_analysis_web(sid, config_path):
    s = _get_or_create_session(sid)

    class WebLogRedirector:
        def write(self, msg):
            if msg and msg.strip():
                s["log_queue"].put(msg.rstrip())
        def flush(self):
            pass

    old_stdout = sys.stdout
    sys.stdout = WebLogRedirector()

    try:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.basename(config_path).rsplit(".", 1)[0]

        # 检测配置类型
        detected = _detect_mode(config_path)
        _push_log(sid, f"配置类型: {detected}")

        if detected in ("pivot", "all"):
            pivot_out = os.path.join(s["work_dir"], f"{base}_分析_{ts}.xlsx")
            _run_pivot_mode(config_path, pivot_out)
            s["output"]["pivot"] = pivot_out

        if detected in ("ppt", "all"):
            ppt_out = os.path.join(s["work_dir"], f"{base}_报告_{ts}.pptx")
            pivot_src = s["output"].get("pivot")
            _run_ppt_mode(config_path, ppt_out, pivot_data_file=pivot_src)
            s["output"]["ppt"] = ppt_out

        _push_log(sid, "✅ 分析完成！")
    except Exception as e:
        _push_log(sid, f"❌ 错误: {e}")
        import traceback
        for line in traceback.format_exc().split("\n"):
            if line.strip():
                _push_log(sid, line)
    finally:
        sys.stdout = old_stdout
        s["status"] = "done"


@app.route("/api/logs")
def api_logs():
    sid = request.args.get("sid", "default")
    s = _get_or_create_session(sid)

    def generate():
        q = s["log_queue"]
        while True:
            try:
                msg = q.get(timeout=0.3)
                yield f"data: {json.dumps({'msg': msg})}\n\n"
            except queue.Empty:
                yield f"data: {json.dumps({'heartbeat': True})}\n\n"
                if s["status"] in ("done", "error"):
                    yield f"data: {json.dumps({'status': s['status']})}\n\n"
                    break

    return Response(generate(), mimetype="text/event-stream")


@app.route("/api/status")
def api_status():
    sid = request.args.get("sid", "default")
    s = _get_or_create_session(sid)
    return jsonify({
        "status": s["status"],
        "output": {k: os.path.basename(v) for k, v in s["output"].items()},
    })


@app.route("/api/preview/excel")
def api_preview_excel():
    sid = request.args.get("sid", "default")
    s = _get_or_create_session(sid)
    file_key = request.args.get("file", "pivot")
    filepath = s["output"].get(file_key)

    if not filepath or not os.path.exists(filepath):
        return "<p>文件尚未生成</p>"

    return _excel_to_html(filepath)


def _excel_to_html(filepath):
    import openpyxl
    wb = openpyxl.load_workbook(filepath, data_only=True)

    sheets_html = []
    for sname in wb.sheetnames:
        ws = wb[sname]
        rows = list(ws.iter_rows())
        if not rows:
            continue

        max_col = max(len(row) for row in rows) if rows else 1

        html = f'<div class="sheet-name">{sname}</div>\n<table>'
        for ri, row in enumerate(rows):
            html += "<tr>"
            is_header = (ri == 0)
            tag = "th" if is_header else "td"
            for ci in range(max_col):
                cell = row[ci] if ci < len(row) else None
                val = cell.value if cell else ""
                if val is None:
                    val = ""
                # 百分比格式
                if isinstance(val, float):
                    if cell and cell.number_format and "%" in str(cell.number_format):
                        val = f"{val:.1%}"
                    elif abs(val) < 1 and val != 0:
                        val = round(val, 4)
                html += f"<{tag}>{val}</{tag}>"
            html += "</tr>\n"
        html += "</table>"
        sheets_html.append(html)

    wb.close()
    css = """
    <style>
    body { font-family: 'Microsoft YaHei',sans-serif; margin:0; padding:12px; background:#f7f8fa; }
    .sheet-name { font-size:14px; font-weight:bold; color:#182B49; margin:12px 0 6px; }
    table { border-collapse:collapse; width:100%; font-size:12px; background:#fff; border:1px solid #ddd; }
    th { background:#182B49; color:#fff; padding:6px 10px; text-align:left; }
    td { padding:4px 10px; border-bottom:1px solid #eee; }
    tr:hover td { background:#f0f4ff; }
    </style>
    """
    return css + "<br>".join(sheets_html)


@app.route("/api/preview/ppt/<int:page>")
def api_preview_ppt_page(page):
    sid = request.args.get("sid", "default")
    s = _get_or_create_session(sid)
    filepath = s["output"].get("ppt")
    if not filepath or not os.path.exists(filepath):
        return "<p>PPT 文件尚未生成</p>", 404

    from pptx import Presentation
    from PIL import Image, ImageDraw, ImageFont
    import io

    prs = Presentation(filepath)
    slides = list(prs.slides)
    if page < 1 or page > len(slides):
        return "<p>页码超出范围</p>", 404

    slide = slides[page - 1]
    width = int(prs.slide_width / 12700)
    height = int(prs.slide_height / 12700)

    # 用 pptx 的 slide export 把每张 slide 的 shapes 绘制到 Pillow
    try:
        from pptx.util import Inches, Pt, Emu
        img = _render_slide_to_png(slide, width, height, prs)
    except Exception:
        img = _render_slide_simple(slide, width, height)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return send_file(buf, mimetype="image/png")


def _render_slide_to_png(slide, width, height, prs):
    """将 PPT 单页渲染为 Pillow Image（基于文本+形状绘制）"""
    from PIL import Image, ImageDraw, ImageFont

    img = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    try:
        font_large = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 24)
        font_small = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 14)
        font_mid = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 18)
    except Exception:
        font_large = ImageFont.load_default()
        font_small = ImageFont.load_default()
        font_mid = ImageFont.load_default()

    for shape in slide.shapes:
        left = int(shape.left / 12700)
        top = int(shape.top / 12700)
        w = int(shape.width / 12700)
        h = int(shape.height / 12700)

        if hasattr(shape, "fill") and shape.fill.type is not None:
            try:
                rgb = str(shape.fill.fore_color.rgb)
                if len(rgb) == 6:
                    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
                    draw.rectangle([left, top, left + w, top + h], fill=(r, g, b))
            except Exception:
                pass

        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if not text:
                    continue
                try:
                    sz = para.font.size
                    if sz:
                        pt = int(sz / 12700)
                        ft = font_mid if pt > 16 else font_small
                    else:
                        ft = font_small
                    draw.text((left + 4, top + 4), text, fill="#182B49", font=ft)
                    top += int(ft.size * 1.3)
                except Exception:
                    draw.text((left + 4, top + 4), text, fill="#182B49")

    return img


def _render_slide_simple(slide, width, height):
    """简化版：只提取文字"""
    from PIL import Image, ImageDraw

    img = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if text:
                texts.append(text)

    try:
        font = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 16)
    except Exception:
        font = ImageFont.load_default()

    y = 20
    for t in texts:
        draw.text((20, y), t, fill="#182B49", font=font)
        y += int(font.size * 1.5)

    return img


@app.route("/api/preview/ppt/info")
def api_preview_ppt_info():
    sid = request.args.get("sid", "default")
    s = _get_or_create_session(sid)
    filepath = s["output"].get("ppt")
    if not filepath or not os.path.exists(filepath):
        return jsonify({"pages": 0})

    from pptx import Presentation
    prs = Presentation(filepath)
    return jsonify({"pages": len(list(prs.slides))})


@app.route("/api/download/<filetype>")
def api_download(filetype):
    sid = request.args.get("sid", "default")
    s = _get_or_create_session(sid)
    filepath = s["output"].get(filetype)
    if not filepath or not os.path.exists(filepath):
        return "文件不存在", 404
    return send_file(filepath, as_attachment=True, download_name=os.path.basename(filepath))


# ==================== 启动 ====================

def start_server(port=8899, open_browser=True):
    if open_browser:
        import webbrowser
        threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    print(f"\n🌐 Excel 统一分析工具 Web 模式 v{__VERSION__}")
    print(f"   访问地址: http://localhost:{port}")
    print(f"   按 Ctrl+C 停止服务\n")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8899
    start_server(port)
