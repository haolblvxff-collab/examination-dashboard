#!/usr/bin/env python3
"""运城培优成绩追踪看板 - 轻量本地 API 服务。"""

import io
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote

import openpyxl
from fastapi import FastAPI, File, Query, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


SUBJECTS = ["语文", "数学", "英语", "物理", "化学", "生物"]
EXAM_ORDER = ["一模", "二模", "三模", "四模", "五模"]
BASE = Path(__file__).parent
DATA: dict[str, Any] = {"students": {}, "exams": [], "schools": [], "filename": ""}

app = FastAPI(title="运城培优成绩追踪")


def _resource_path(relative_path: str) -> Path:
    base = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else BASE
    return base / relative_path


app.mount("/static", StaticFiles(directory=str(_resource_path("static"))), name="static")


def _numeric_score(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.strip())
        except ValueError:
            return None
    return None


def _ordered_exams(values: list[str]) -> list[str]:
    unique = list(dict.fromkeys(values))
    position = {value: i for i, value in enumerate(unique)}
    return sorted(
        unique,
        key=lambda exam: (
            EXAM_ORDER.index(exam) if exam in EXAM_ORDER else len(EXAM_ORDER),
            position[exam],
        ),
    )


def _student_id(name: str, school: str) -> str:
    return f"{school}\u241f{name}"


def parse_excel(file_bytes: bytes) -> dict[str, Any]:
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception as exc:
        raise ValueError("无法读取文件，请上传有效的 .xlsx 成绩表") from exc
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        raise ValueError("文件至少需要表头和 1 行数据")

    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    columns = {header: index for index, header in enumerate(headers) if header}

    def locate(words: list[str]) -> str | None:
        return next((header for header in headers if any(word in header for word in words)), None)

    name_col = locate(["姓名", "名字", "学生"])
    school_col = locate(["学校"])
    exam_col = locate(["考试", "场次"])
    if not name_col:
        raise ValueError("找不到「姓名」列，请使用模板格式")

    students: dict[str, dict[str, Any]] = {}
    exams_seen: list[str] = []

    def ensure_student(name: str, school: str) -> dict[str, Any]:
        key = _student_id(name, school)
        if key not in students:
            students[key] = {"name": name, "school": school, "scores": {}, "totals": {}}
        return students[key]

    if exam_col:
        available_subjects = [subject for subject in SUBJECTS if subject in columns]
        for values in rows[1:]:
            if not any(value is not None for value in values):
                continue
            name = str(values[columns[name_col]]).strip() if values[columns[name_col]] is not None else ""
            if not name:
                continue
            school = (
                str(values[columns[school_col]]).strip()
                if school_col and values[columns[school_col]] is not None
                else "-"
            )
            raw_exam = values[columns[exam_col]]
            if raw_exam is None or not str(raw_exam).strip():
                continue
            exam = str(raw_exam).strip()
            exams_seen.append(exam)
            info = ensure_student(name, school)
            total = 0.0
            has_score = False
            for subject in available_subjects:
                score = _numeric_score(values[columns[subject]])
                if score is not None:
                    info["scores"].setdefault(subject, {})[exam] = score
                    total += score
                    has_score = True
            if has_score:
                info["totals"][exam] = round(total, 1)
    else:
        exam_map: dict[str, dict[str, int]] = {}
        for header, index in columns.items():
            parts = header.split("_", 1)
            if len(parts) == 2 and parts[1] in SUBJECTS:
                exam_map.setdefault(parts[0], {})[parts[1]] = index
        exams_seen = list(exam_map)
        for values in rows[1:]:
            name = str(values[columns[name_col]]).strip() if values[columns[name_col]] is not None else ""
            if not name:
                continue
            school = (
                str(values[columns[school_col]]).strip()
                if school_col and values[columns[school_col]] is not None
                else "-"
            )
            info = ensure_student(name, school)
            for exam, subject_map in exam_map.items():
                total = 0.0
                has_score = False
                for subject, index in subject_map.items():
                    score = _numeric_score(values[index])
                    if score is not None:
                        info["scores"].setdefault(subject, {})[exam] = score
                        total += score
                        has_score = True
                if has_score:
                    info["totals"][exam] = round(total, 1)

    if not students:
        raise ValueError("未找到有效成绩数据")
    return {
        "students": students,
        "exams": _ordered_exams(exams_seen),
        "schools": sorted({student["school"] for student in students.values()}),
    }


def _mean(values: list[float]) -> float:
    return round(sum(values) / len(values), 1) if values else 0


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    filename = file.filename or ""
    if not filename.lower().endswith(".xlsx"):
        return JSONResponse({"ok": False, "error": "目前仅支持 .xlsx 文件"}, status_code=400)
    try:
        result = parse_excel(await file.read())
    except ValueError as exc:
        return JSONResponse({"ok": False, "error": str(exc)}, status_code=400)
    DATA.update(result)
    DATA["filename"] = filename
    return {
        "ok": True,
        "students": len(result["students"]),
        "exams": len(result["exams"]),
        "schools": len(result["schools"]),
        "filename": filename,
    }


@app.get("/api/data")
def get_data():
    if not DATA["students"]:
        return {"ok": False, "error": "请先上传成绩文件"}
    return {
        "ok": True,
        "students": DATA["students"],
        "exams": DATA["exams"],
        "schools": DATA["schools"],
        "filename": DATA["filename"],
        "subjects": SUBJECTS,
    }


@app.get("/api/stats")
def get_stats():
    if not DATA["students"]:
        return JSONResponse({"error": "请先上传成绩文件"}, status_code=400)
    students = DATA["students"]
    exams = DATA["exams"]
    last_exam = exams[-1] if exams else None
    active = [student for student in students.values() if last_exam in student["totals"]]
    totals = [student["totals"][last_exam] for student in active] if last_exam else []
    top = max(active, key=lambda item: item["totals"][last_exam], default=None) if last_exam else None

    best_change = 0.0
    best_info = ""
    for student in students.values():
        for previous, current in zip(exams, exams[1:]):
            if previous in student["totals"] and current in student["totals"]:
                change = student["totals"][current] - student["totals"][previous]
                if change > best_change:
                    best_change = round(change, 1)
                    best_info = f'{student["name"]} ({previous}->{current})'

    subject_avg = {}
    for subject in SUBJECTS:
        scores = [
            student["scores"].get(subject, {}).get(last_exam)
            for student in active
            if student["scores"].get(subject, {}).get(last_exam) is not None
        ]
        subject_avg[subject] = _mean(scores)
    return {
        "n_students": len(students),
        "n_exams": len(exams),
        "n_schools": len(DATA["schools"]),
        "max_total": top["totals"][last_exam] if top and last_exam else 0,
        "max_name": top["name"] if top else "",
        "avg_total": _mean(totals),
        "best_improve": best_change,
        "best_improve_info": best_info,
        "subject_avg": subject_avg,
        "exams": exams,
    }


@app.get("/api/ranking")
def ranking(exam: str = Query(...), school: str = Query("")):
    if not DATA["students"]:
        return JSONResponse({"error": "无数据"}, status_code=400)
    items = []
    for info in DATA["students"].values():
        if school and info["school"] != school:
            continue
        if exam not in info["totals"]:
            continue
        row = {"name": info["name"], "school": info["school"], "total": info["totals"][exam]}
        for subject in SUBJECTS:
            row[subject] = info["scores"].get(subject, {}).get(exam)
        items.append(row)
    items.sort(key=lambda item: item["total"], reverse=True)
    return {"exam": exam, "ranking": items}


@app.get("/api/template")
def download_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "成绩录入"
    header_font = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )
    headers = ["姓名", "学校", "考试"] + SUBJECTS + ["总分(自动计算)", "备注"]
    for column, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=column, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    rows = [
        ["张三", "运城培优", "一模", 108, 120, 115, 85, 78, 82, None, "示例-可删除"],
        ["张三", "运城培优", "二模", 112, 125, 118, 88, 82, 85, None, ""],
        ["李四", "运城二中", "一模", 98, 110, 105, 78, 72, 76, None, ""],
    ]
    for row_index, values in enumerate(rows, 2):
        for column, value in enumerate(values, 1):
            cell = ws.cell(row=row_index, column=column, value=value)
            cell.alignment = center
            cell.border = border
    for column, width in enumerate([12, 14, 8, 8, 8, 8, 8, 8, 8, 16, 20], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(column)].width = width
    notes = wb.create_sheet("使用说明")
    lines = [
        "成绩录入模板使用说明",
        "",
        "1. 请在「成绩录入」sheet 中填写数据",
        "2. 必填列：姓名、学校、考试，三者共同标识一条记录",
        "3. 分数填写数字即可，未参加科目留空",
        "4. 总分列无需手动填写，系统自动计算",
        "5. 当前上传格式：.xlsx",
    ]
    for row_index, value in enumerate(lines, 1):
        notes.cell(row=row_index, column=1, value=value)
    notes.column_dimensions["A"].width = 70
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    filename = quote("成绩录入模板.xlsx")
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"},
    )


_HTML_TEMPLATE: str | None = None


@app.get("/", response_class=HTMLResponse)
def index():
    global _HTML_TEMPLATE
    if _HTML_TEMPLATE is None:
        _HTML_TEMPLATE = _resource_path("templates/index.html").read_text(encoding="utf-8")
    return _HTML_TEMPLATE


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8899)
