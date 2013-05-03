@echo off
IF EXIST ship RMDIR /S /Q ship
IF EXIST x86 RMDIR /S /Q x86
IF EXIST x64 RMDIR /S /Q x64
IF EXIST Setup\x64 RMDIR /S /Q Setup\x64
IF EXIST Setup\x86 RMDIR /S /Q Setup\x86
IF EXIST Setup\bin RMDIR /S /Q Setup\bin

IF EXIST *.ncb DEL /S /Q *.ncb
IF EXIST *.cache DEL /S /Q *.cache
