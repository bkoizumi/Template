@echo off
cd %~dp0
set baseDir=%~dp0

rd /s /q "%baseDir%Source\01.©ΟΰθHZop.xlsm"
rd /s /q "%baseDir%Source\10.έv_ξ{yVXez.xlsm"
rd /s /q "%baseDir%Source\11.έv_DataBaseyVXez.xlsm"
rd /s /q "%baseDir%Source\12.έv_IFyVXez.xlsm"
rd /s /q "%baseDir%Source\41.eXgΚρyz.xlsm"

cd ..
cscript //nologo vbac.wsf decombine /binary:Template /source:Template/Source
