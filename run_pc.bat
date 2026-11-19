@echo off
rem PC auto run for Task Scheduler. Log: output\pc_run.log
cd /d C:\dev\daily-site
if not exist output mkdir output
set PYTHONIOENCODING=utf-8
.venv\Scripts\python.exe pc_run.py >> output\pc_run.log 2>&1
