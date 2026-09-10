@echo off
chcp 65001 >nul
setlocal
if not exist "%~dp0学生综合管理系统.exe" (
 echo 请把本文件放到学生综合管理系统.exe所在的文件夹，再双击运行。
 pause
 exit /b 1
)
pushd "%~dp0"
if errorlevel 1 exit /b 1
start "" "%~dp0学生综合管理系统.exe" --weekend
popd
endlocal
