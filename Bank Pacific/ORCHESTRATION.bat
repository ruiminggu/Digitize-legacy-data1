@echo off
REM Navigate to the directory
cd "C:\Users\User\Desktop\Alysson\bank of hawaii"

REM Set the execution policy to allow running scripts in this session
PowerShell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process"

REM Activate the virtual environment
call .\venv\Scripts\activate

REM Install required packages
pip install -r requirements.txt

REM Run the main script
python setup_and_run.py

pause
