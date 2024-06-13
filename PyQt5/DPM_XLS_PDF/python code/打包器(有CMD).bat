@echo off
chcp 65001
cls
:start
@SET /P name=檔名:
pyinstaller -D %name%.py
goto start