-Locate a machine that has Microsoft Office installed, and not connected to the network via wifi


-The server should not be hosting any business critical applications (like the file share server itself)

-Unzip contents to directory like c:\scripts


-Open a powershell prompt as administrator, change working directory to c:\scripts


-The first time you run the script, type the following: Set-ExecutionPolicy Unrestricted

-To run the script type: .\pre_migration_master.ps1 -startDirectory "c:\test -generateReports" where c:\test is the directory you want to scan
-Note: the start directory must be an actual directory, not the fileshare server itself
-You can also import a csv list of directories to scan with column name "HomeDirectoy" that contain the directories to scan.  
-Type: .\pre_migration_master.ps1 -directoryList ".\UserExport.csv"
-You can choose to generate reports or not by adding or removing the "-generateReports" flag from the end of the command


-During the second half of the process the script will test Office documents and may freeze up. It is important to log in from time to time and check to see if there are any instances of word/excel/powerpoint open. If so, close them down. 


-Once complete, the script will generate a crawl log if there are errors and a report log csv files. Additional reports can be ran from FileToOneDrive.db so don't delete it!
-The script can be run once the prior scan is complete with a different directory, and will generate additional reports

-If you receive an error like 'Install-Module is not recognized as a cmdlet' the server should have installed Windows Management Framework 5.1