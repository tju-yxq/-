@echo off
chcp 65001 >nul
title 一键生成所有分类 Word 文档
echo ================================================
echo       正在为您开始批量生成所有分类的文档......
echo ================================================
echo.

echo [1/5] 正在生成【美食】文档...
py 脚本/excel_to_docx_美食.py

echo.
echo [2/5] 正在生成【景点】文档...
py 脚本/excel_to_docx_景点.py

echo.
echo [3/5] 正在生成【购物】文档...
py 脚本/excel_to_docx_购物.py

echo.
echo [4/5] 正在生成【民宿】文档...
py 脚本/excel_to_docx_民宿.py

echo.
echo [5/5] 正在生成【特产】文档...
py 脚本/excel_to_docx_特产.py

echo.
echo ================================================
echo        所有类别文档均已处理完毕！
echo        请前往“生成结果”文件夹下验收成果。
echo ================================================
pause
