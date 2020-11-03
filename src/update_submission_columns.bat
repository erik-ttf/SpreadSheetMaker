set PYTHONPATH=X:\src\core\python\sitecustomize
set ttfDevMode=0
C:\Python27\python.exe S:\Pilgrim\custom\src\core\python\update_submission_columns\update_submission_columns.py %*
IF ERRORLEVEL 1 pause
pause