@echo off
chcp 65001 >nul
echo ======================================
echo   Markdown Editor 打包工具
echo ======================================
echo.

:: 检查依赖
echo [1/4] 检查 Python...
python --version || goto :error

echo [2/4] 检查依赖包...
python -c "import webview, markdown, docx, markdown_it, pygments, tkinter" 2>nul || (
    echo 安装缺失的依赖...
    pip install pywebview markdown python-docx markdown-it-py pygments
)

echo [3/4] 清理旧构建...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

echo [4/4] 打包中（这可能需要几分钟）...
pyinstaller ^
    --name="MarkdownEditor" ^
    --windowed ^
    --onefile ^
    --icon=NONE ^
    --add-data="app.py;." ^
    --hidden-import=webview ^
    --hidden-import=markdown ^
    --hidden-import=markdown_it ^
    --hidden-import=markdown_it.common.utils ^
    --hidden-import=markdown_it.rules_core.state_block ^
    --hidden-import=markdown_it.rules_core.fence ^
    --hidden-import=markdown_it.rules_core.reference ^
    --hidden-import=markdown_it.rules_core.softbreak ^
    --hidden-import=markdown_it.token ^
    --hidden-import=markdown_it.node ^
    --hidden-import=markdown_it.renderer ^
    --hidden-import=markdown_it.renderer_core ^
    --hidden-import=markdown_it.parser ^
    --hidden-import=markdown_it.presets ^
    --hidden-import=markdown_it.presets.commonmark ^
    --hidden-import=markdown_it.presets.zero ^
    --hidden-import=markdown_it.utils ^
    --hidden-import=mdurl ^
    --hidden-import=mdurl.parse ^
    --hidden-import=pygments ^
    --hidden-import=pygments.lexers ^
    --hidden-import=pygments.formatters ^
    --hidden-import=docx ^
    --hidden-import=docx.api ^
    --hidden-import=docx.document ^
    --hidden-import=docx.oxml ^
    --hidden-import=docx.oxml.text ^
    --hidden-import=docx.oxml.table ^
    --hidden-import=docx.oxml.styles ^
    --hidden-import=lxml ^
    --hidden-import=PIL ^
    --hidden-import=tkinter ^
    app.py

if exist dist\MarkdownEditor.exe (
    echo.
    echo ======================================
    echo   打包成功！
    echo   exe 位置: dist\MarkdownEditor.exe
    echo ======================================
    echo.
    echo 即将打开文件夹...
    start explorer dist
) else (
    echo.
    echo 打包失败，请检查上方错误信息
    pause
    goto :eof
)

goto :eof

:error
echo Python 未找到，请先安装 Python 3.8+
pause
