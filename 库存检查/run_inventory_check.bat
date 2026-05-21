@echo off
pushd "%~dp0"
python inventory_check.py >> inventory_check.log 2>&1
popd
