@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
echo Running STRIX RAG Search Test with UTF-8 encoding...
echo.
py test_rag_search.py