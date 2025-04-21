@echo off
echo ========================================
echo    PPT专业翻译工具 - DeepSeek版
echo ========================================
echo.
echo 正在检查环境...

python check_environment.py

if errorlevel 1 (
  echo.
  echo 环境检查未通过，请解决上述问题后再启动翻译工具。
  echo.
  pause
  exit /b 1
)

echo.
echo 正在启动PPT翻译工具...
echo.

python gui.py

if errorlevel 1 (
  echo.
  echo 程序启动失败！
  echo 请确认Python环境及依赖已正确安装
  echo.
) else (
  echo 程序已关闭，感谢使用！
)

pause 