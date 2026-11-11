@echo off
REM Atalho para o Agendador de Tarefas do Windows.
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python main.py %*
