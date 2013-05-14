@echo off
IF EXIST ipch RMDIR /S /Q ipch
IF EXIST Debug RMDIR /S /Q Debug
IF EXIST Release RMDIR /S /Q Release

IF EXIST Template\bin RMDIR /S /Q Template\bin
IF EXIST Template\obj RMDIR /S /Q Template\obj

IF EXIST VSIX\bin RMDIR /S /Q VSIX\bin
IF EXIST VSIX\obj RMDIR /S /Q VSIX\obj

IF EXIST VisioWixExtension\wixext\bin RMDIR /S /Q VisioWixExtension\wixext\bin
IF EXIST VisioWixExtension\wixext\obj RMDIR /S /Q VisioWixExtension\wixext\obj
IF EXIST VisioWixExtension\wixext\Data\Messages.cs DEL /Q VisioWixExtension\wixext\Data\Messages.cs
IF EXIST VisioWixExtension\wixext\Data\Messages.resources DEL /Q VisioWixExtension\wixext\Data\Messages.resources

IF EXIST VisioWixExtension\wixlib\Debug RMDIR /S /Q VisioWixExtension\wixlib\Debug
IF EXIST VisioWixExtension\wixlib\Release RMDIR /S /Q VisioWixExtension\wixlib\Release

IF EXIST VisioWixExtension\ca\Release RMDIR /S /Q VisioWixExtension\ca\Release
IF EXIST VisioWixExtension\ca\Debug RMDIR /S /Q VisioWixExtension\ca\Debug
