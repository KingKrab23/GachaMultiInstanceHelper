@echo off
python -m venv venv
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
echo Virtual environment is ready! Use 'venv\Scripts\activate.bat' to activate it in the future.
