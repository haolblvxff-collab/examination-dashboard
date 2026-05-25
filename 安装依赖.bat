@echo off
chcp 65001 >nul
echo ================================
echo   运城培优成绩追踪看板 - 安装依赖
echo ================================
echo.
echo 正在安装所需组件，请稍候...
pip install -r requirements-win.txt
echo.
echo 安装完成！双击「启动看板.bat」即可使用。
pause
