@echo off
chcp 65001 > nul
pushd "%~dp0"
python inventory_check.py
pause
popd
