@echo off
cls
COLOR 1F
cls
@echo.
@echo TODAY DATE    :    %DATE%
@echo CURRENT TIME  :    %TIME% 
@echo.
@echo      Hello ! Everybody  ..~~
@echo.
@echo.
@echo. ...  This is GDS Clinic Routine Check Management Program  ....
@echo.
@echo.
 
@echo off
python C:\GDSRC\A1_BP_EKG_PFT_0504.py
python C:\GDSRC\A2_US_DEXA_0510.py
python C:\GDSRC\A3_Lab_0627.py
python C:\GDSRC\A4_dist_files.py
@echo.

