@echo off

echo Running SSIS packages...

dtexec /f "C:\Users\ITDS.Pratap\Documents\LTI_Data.dtsx"
dtexec /f "C:\Users\ITDS.Pratap\Documents\PQQ.dtsx"
dtexec /f "C:\Users\ITDS.Pratap\Documents\PTW.dtsx"
dtexec /f "C:\Users\ITDS.Pratap\Documents\Trainings.dtsx"
dtexec /f "C:\Users\ITDS.Pratap\Documents\Audit.dtsx"
dtexec /f "C:\Users\ITDS.Pratap\Documents\Inspection.dtsx"
dtexec /f "C:\Users\ITDS.Pratap\Documents\Incidents.dtsx"

pause
