if not exist "%CD%\FROZEN" mkdir "%CD%\FROZEN"
pyinstaller Main.py -y --distpath "%CD%\FROZEN\dist" --workpath "%CD%\FROZEN\build" --specpath "%CD%\FROZEN"
pyinstaller FROZEN\Main.spec -y --distpath "%CD%\FROZEN\dist" --workpath "%CD%\FROZEN\build"
xcopy "%CD%\*.ui" "%CD%\FROZEN\dist" /y