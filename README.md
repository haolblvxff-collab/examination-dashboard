# 运城培优 · 成绩追踪看板

> Examination Dashboard — 本地成绩分析工具

[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.57+-red.svg)](https://streamlit.io)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## 📖 简介

运城培优成绩追踪看板是一款本地运行的成绩分析工具。上传学生 Excel 成绩表，自动生成：

- **6 项统计卡片** — 人数、最高分、均分、学校数、考试场次、进步最大
- **8 种分析图表** — 走势图、分布图、箱线图、雷达图、热力图、学校对比、进步追踪
- **可筛选排名表** — 按考试/学校筛选，总分排序

所有数据在本地处理，不上传任何服务器。

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
streamlit run app.py
```

浏览器访问 `http://localhost:8501`

---

## 📊 数据格式

上传的 Excel 文件需包含以下列：

| 字段 | 说明 |
|------|------|
| 学生姓名 | 学生姓名 |
| 学校 | 所属学校 |
| 考试名称 | 考试场次标识 |
| 总分 | 总成绩 |
| 各科成绩 | 语文、数学、英语等 |

---

## 🏗️ 技术栈

- **Streamlit** — Web UI 框架
- **Pandas** — 数据处理
- **Plotly** — 交互式图表

---

## 📦 构建

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "运城培优成绩追踪" app.py
```
