@echo off
REM Usado pelo Agendador de Tarefas do Windows (ver agendar_tarefa.ps1).
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Ambiente .venv nao encontrado. Siga a instalacao do README.
  exit /b 2
)
".venv\Scripts\python.exe" main.py %*
exit /b %ERRORLEVEL%
