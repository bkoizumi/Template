@echo off
cd %~dp0
set baseDir=%~dp0

rd /s /q "%baseDir%Source\01.©ÏàèHZop.xlsm"
rd /s /q "%baseDir%Source\10.Ýv_î{yVXez.xlsm"
rd /s /q "%baseDir%Source\11.Ýv_DByVXez.xlsm"
rd /s /q "%baseDir%Source\12.Ýv_IFyVXez.xlsm"
rd /s /q "%baseDir%Source\41.eXgÊñyz.xlsm"

cd ..
cscript //nologo vbac.wsf decombine /binary:Template /source:Template/Source
