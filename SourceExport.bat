@echo off
cd %~dp0
set baseDir=%~dp0

rd /s /q "%baseDir%Source\έv_DataBaseyVXez.xlsm"
rd /s /q "%baseDir%Source\έv_IFyVXez.xlsm"
rd /s /q "%baseDir%Source\έv_ξ{yVXez.xlsm"

cd ..
cscript //nologo vbac.wsf decombine /binary:Template /source:Template/Source
