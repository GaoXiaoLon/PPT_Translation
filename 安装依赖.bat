@echo off
echo ========================================
echo    PPT专业翻译工具 - 依赖安装
echo ========================================
echo.
echo 正在安装必要的依赖包...
echo.

pip install -r requirements.txt

if errorlevel 1 (
  echo.
  echo 安装失败！请检查您的网络连接或Python环境。
  echo 您可以尝试手动安装依赖：
  echo   pip install python-pptx requests python-dotenv tqdm
  echo.
) else (
  echo.
  echo 依赖安装成功！现在您可以使用翻译工具了。
  echo 请运行 start_translator.bat 启动图形界面。
  echo.
)

pause 