@echo off
cls
COLOR 1F
cls
@echo.
@echo TODAY DATE    :    %DATE%
@echo CURRENT TIME  :    %TIME% 
@echo.
@echo      Hello ! Everybody  ..~~^^
@echo.
@echo.
@echo. ...  This is GDS Clinic Routine Check Management Program  ....
@echo.
@echo.
 
@echo off

python C:/GDSRC/file_list.py
@echo.
Pause
EXIT()