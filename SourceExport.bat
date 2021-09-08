@echo off
cd %~dp0
set baseDir=%~dp0

rd /s /q "%baseDir%Source\設計書_DataBase【＊＊＊システム】.xlsm"
rd /s /q "%baseDir%Source\設計書_IF【＊＊＊システム】.xlsm"
rd /s /q "%baseDir%Source\設計書_基本【＊＊＊システム】.xlsm"

cd ..
cscript //nologo vbac.wsf decombine /binary:Template /source:Template/Source
