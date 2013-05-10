@echo off
IF EXIST Template\bin RMDIR /S /Q Template\bin
IF EXIST Template\obj RMDIR /S /Q Template\obj
IF EXIST Template\*.user DEL /S /Q Template\*.user

IF EXIST VSIX\bin RMDIR /S /Q VSIX\bin
IF EXIST VSIX\obj RMDIR /S /Q VSIX\obj
IF EXIST VSIX\*.user DEL /S /Q VSIX\*.user
