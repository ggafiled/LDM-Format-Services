#LDM-Format-Services
###Windows service for read text file and clear text into mssql database. 

##Usage:
###Config: Open Registry editor > HKEY_LOCAL_MACHINE > SOFTWARE > LDM and create new key following below.
- BACKUP_DIRECTORY (Backup folder after already processing file)
- BACKUP_INCORECTFORMAT_DIRECTORY (Backup folder after already processing file if file is not format that process want.)
- INPUT_DIRECTORY (Input folder for service auto detech file if has file in folder)
- OUTPUT_DIRECTORY (Output folder hold output file that service already cleaning)

###Installation: Open power shell (Admin) and type below commands for install.
1. New-Service -Name "YourServiceName" -BinaryPathName <yourprojectpath>.exe
2. Start-Service -Name "YourServiceName"
