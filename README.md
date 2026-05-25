# 运城培优 · 成绩追踪看板

> Examination Dashboard — 本地成绩分析工具

[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://python.org)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.136+-green.svg)](https://fastapi.tiangolo.com)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## 📖 简介

运城培优成绩追踪看板是一款本地运行的成绩分析工具。上传学生 Excel 成绩表，自动生成：

- **6 项统计卡片** — 人数、最高分、均分、学校数、考试场次、进步最大
- **分析图表** — 走势图、分布图、箱线图、雷达图、热力图、学校对比、进步追踪
- **可筛选排名表** — 按考试/学校筛选，总分排序

所有数据仅在本机 `127.0.0.1` 服务中处理，不上传到外部服务器。

---

## 🚀 快速开始

### Windows

```batch
# 双击运行
安装依赖.bat
启动看板.bat
```

### macOS / Linux

```bash
pip install -r requirements-win.txt
python app.py
```

浏览器访问 `http://localhost:8899`

或使用桌面启动器：

```bash
python launcher.py    # 自动打开浏览器
```

---

## 📊 数据格式

上传的 `.xlsx` Excel 文件需包含以下列：

| 字段 | 说明 |
|------|------|
| 学生姓名 | 学生姓名 |
| 学校 | 所属学校 |
| 考试名称 | 考试场次标识 |
| 总分 | 总成绩 |
| 各科成绩 | 语文、数学、英语等 |

---

## 🏗️ 技术栈

- **FastAPI** — 本机 Web 服务 + REST API
- **OpenPyXL** — Excel 读写与数据解析
- **Apache ECharts** — 浏览器端离线图表渲染

打包版携带位于 `static/echarts.min.js` 的离线图表库，采用 Apache License 2.0。
源码运行时如未提供该大文件，将回退加载 Apache ECharts 的 CDN 发行文件。

## 📦 构建

```bash
pip install -r requirements-build.txt
pyinstaller --onefile --name "运城培优成绩追踪" --add-data "templates;templates" --add-data "static;static" launcher.py
```

运行依赖与构建依赖已分开；普通使用者无需安装 `pyinstaller`。
