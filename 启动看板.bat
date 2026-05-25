@echo off
chcp 65001 >nul
echo ================================
echo   运城培优成绩追踪看板
echo ================================
echo.
echo 正在启动服务，浏览器将自动打开...
echo 如未自动打开，请访问 http://localhost:8899
echo.
echo 关闭此窗口即可停止服务。
echo ================================
python launcher.py
pause
