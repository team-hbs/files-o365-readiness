-Locate a machine that has Microsoft Office installed, and not connected to the network via wifi


-The server should not be hosting any business critical applications (like the file share server itself)

-Unzip contents to directory like c:\scripts


-Open a powershell prompt as administrator, change working directory to c:\scripts


-The first time you run the script, type the following: Set-ExecutionPolicy Unrestricted

-USAGE
-    .\pre_migration_master.ps1 -mode "XXXX" -source "XXXX" -report "XXXX"

-PARAMETER OPTIONS
-mode "single" = runs the crawl on a single directory located at the path passed into "-source"
		 or if source is not specified user is prompted to select directory location
-mode "import" = runs the crawl on a list of directories by importing a csv file at the location
		 passed into "-source" with "HomeDirectory" as the directory list column heading
-mode "report" = does not run the crawl and only runs reports on the database

------------------------------------------------------

-source "XXXX" = XXXX is the path of a directory or csv file to import

------------------------------------------------------

-report "overall" = generates 1 report with all the info on one sheet and the errors report
		    for each direcory combined

-report "single" = generates a report for each directory and the errors specific to each directory

-EXAMPLE USAGES
Run crawl on a single directory and generate a report:
.\pre_migration_master.ps1 -mode "single" -source "c:\test" -report "single"

Run crawl on a list of directories without generating reports:
.\pre_migration_master.ps1 -mode "import" -source ".\UserExport.csv"

Run crawl on a list of directories and generate a report for each directory:
.\pre_migration_master.ps1 -mode "import" -source ".\UserExport.csv" -report "single"

Generate report on an already populated database for each directory:
.\pre_migration_master.ps1 -mode "report" -report "single"

Generate report on an already populated database for every directory combined:
.\pre_migration_master.ps1 -mode "report" -report "overall"


-During the second half of the process the script will test Office documents and may freeze up. It is important 
 to log in from time to time and check to see if there are any instances of word/excel/powerpoint open. If so, close them down. 


-Once complete, the script will generate a crawl log if there are errors and a report log csv files. 
 Additional reports can be ran from FileToOneDrive.db so don't delete it!
-The script can be run once the prior scan is complete with a different directory, and will generate additional reports

-If you receive an error like 'Install-Module is not recognized as a cmdlet' the server should have installed Windows Management Framework 5.1