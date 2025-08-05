@echo off
echo Installing STRIX dependencies...
echo.

echo Upgrading pip...
venv\Scripts\python.exe -m pip install --upgrade pip

echo.
echo Installing core dependencies...
venv\Scripts\pip.exe install langchain langchain-openai langchain-community langgraph supabase

echo.
echo Installing document processing libraries...
venv\Scripts\pip.exe install pypdf docx2txt python-pptx beautifulsoup4

echo.
echo Installing other dependencies...
venv\Scripts\pip.exe install pandas numpy streamlit python-dotenv pydantic

echo.
echo Installation complete!
pause