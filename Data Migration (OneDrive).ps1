# Modification must be made, here it goes
# Will figure out how to make variables for the following: excluded list, robocopy Operations, robocopy log file location, Dynamic Drive Location/Declara

# Declaring Variables
$Lotus = 'C:\Lotus'
$src = 'C:'
# $dst = 'D'
$xcopyOP= '/Q','/E','/V','/C','/I','/Y','/J'
$roboOP= '/E', '/R:2','/W:0','/IM','/NFL','/NDL','/V','/TS','/FP', '/XF', 'desktop.ini'
$cExempt= "/XD","C:\Program Files (x86)","C:\Program Files","C:\Oracle11G","C:\OneDriveTemp",
          "C:\PerfLogs","C:\Users","C:\Win-builds","C:\Windows","C:\boot","C:\Intel",
          "C:\ProgramData","C:\CheckPoint","C:\Lotus",'C:\$Windows.~WS','C:\$Recycle.Bin',
          "C:\Recovery","C:\Documents and Settings","C:\ESD","C:\System Volume Information",
          'C:\$WINDOWS.~BT'  
# This allows us to detect the logged in Employee
# Saves the employee number in $USERID
#$temp = $env:USERPROFILE
#$USERID = $temp -replace '\D+(\d+)', '$1'
# Or even better
# Because of this modification, we must now fix all paths in copy commands
$USERID = (Get-ChildItem Env:\USERNAME).Value
$profileExempt = '/XD', "C:\Users\$USERID\Desktop", "C:\Users\$USERID\Documents", 
                 "C:\Users\$USERID\Favorites", "C:\Users\$USERID\Downloads", 
                 "C:\Users\$USERID\Pictures", "C:\Users\$USERID\Videos", 
                 "C:\Users\$USERID\Music", "C:\Users\$USERID\Contacts", 
                 "C:\Users\$USERID\Saved Games", "C:\Users\$USERID\Links", 
                 "C:\Users\$USERID\OneDrive", "C:\Users\$USERID\OneDrive - Underwriters Laboratories", 
                 "C:\Users\$USERID\Underwriters Laboratories", "C:\Users\$USERID\AppData", 
                 "C:\Users\$USERID\Application Data", "C:\Users\$USERID\Cookies", 
                 "C:\Users\$USERID\IntelGraphicsProfiles", "C:\Users\$USERID\Local Settings", 
                 "C:\Users\$USERID\My Documents", "C:\Users\$USERID\NetHood", "C:\Users\$USERID\PrintHood", 
                 "C:\Users\$USERID\Recent", "C:\Users\$USERID\SendTo", "C:\Users\$USERID\Start Menu", 
                 "C:\Users\$USERID\Templates"      

# Numbering Prompts
$totalSteps = '22'
$newLine = "`n"
$step0 = $newLine + 'Step #0 of ' + $totalSteps + $newLine
$step1 = $newLine + 'Step #1 of ' + $totalSteps + $newLine
$step2 = $newLine + 'Step #2 of ' + $totalSteps + $newLine
$step3 = $newLine + 'Step #3 of ' + $totalSteps + $newLine 
$step4 = $newLine + 'Step #4 of ' + $totalSteps + $newLine
$step5 = $newLine + 'Step #5 of ' + $totalSteps + $newLine
$step6 = $newLine + 'Step #6 of ' + $totalSteps + $newLine
$step7 = $newLine + 'Step #7 of ' + $totalSteps + $newLine
$step8 = $newLine + 'Step #8 of ' + $totalSteps + $newLine
$step9 = $newLine + 'Step #9 of ' + $totalSteps + $newLine
$step10 = $newLine + 'Step #10 of ' + $totalSteps + $newLine
$step11 = $newLine + 'Step #11 of ' + $totalSteps + $newLine
$step12 = $newLine + 'Step #12 of ' + $totalSteps + $newLine
$step13 = $newLine + 'Step #13 of ' + $totalSteps + $newLine
$step14 = $newLine + 'Step #14 of ' + $totalSteps + $newLine
$step15 = $newLine + 'Step #15 of ' + $totalSteps + $newLine
$step16 = $newLine + 'Step #16 of ' + $totalSteps + $newLine
$step17 = $newLine + 'Step #17 of ' + $totalSteps + $newLine
$step18 = $newLine + 'Step #18 of ' + $totalSteps + $newLine
$step19 = $newLine + 'Step #19 of ' + $totalSteps + $newLine
$step20 = $newLine + 'Step #20 of ' + $totalSteps + $newLine
$step21 = $newLine + 'Step #21 of ' + $totalSteps + $newLine
$step22 = $newLine + 'Step #22 of ' + $totalSteps + $newLine

# Data Backup Prompts
# Declaring prompts here so I dont have to re-write them and have an ease of modification
$bkLotus = $step0 + 'Copying Lotus Notes'
$bkDesktop = $step1 + 'Copying Desktop'
$bkDocuments = $step2 + 'Copying Documents'
$bkFavorites = $step3 + 'Copying Favorites'
$bkDownloads = $step4 + 'Copying Downloads'
$bkPictures = $step5 + 'Copying Pictures'
$bkVideos = $step6 + 'Copying Videos'
$bkMusic = $step7 + 'Copying Music'
$bkGoogle = $step8 + 'Copying Google Bookmarks'
$bkEdge = $step9 + 'Copying Microsoft Edge Bookmarks'
$bkQuickLaunch = $step10 + 'Copying Pinned Applications'
# AppData Backup Prompts
$bkOneNote = $step11 + 'Copying OneNote Roaming and Local AppData'
$bkExcel = $step12 + 'Copying Excel AppData'
$bkOutlook = $step13 + 'Copying Outlook AppData'
$bkPowerPoint = $step14 + 'Copying PowerPoint AppData'
$bkSignatures = $step15 + 'Copying Outlook Signatures'
$bkSpelling = $step16 + 'Copying Spelling AppData'
$bkTeams = $step17 + 'Copying Teams Backgrounds AppData'
$bkWord = $step18 + 'Copying Word AppData'
$bkThemes = $step19 + 'Copying Windows Themes AppData'
# Other
$bkRegistry = $step20 + 'Exporting TaskBand, OneNote, and Mapped Network Drives Registry Values'
$bkProfile = $step21 + 'Copying Unique Profile Folders'
$bkCDrive = $step22 + 'Copying Unique C Drive Folders'

# Data Restore Prompts
$rtLotus = $step0 + 'Restoring Lotus Notes'
$rtDesktop = $step1 + 'Restoring Desktop'
$rtDocuments = $step2 + 'Restoring Documents'
$rtFavorites = $step3 + 'Restoring Favorites'
$rtDownloads = $step4 + 'Restoring Downloads'
$rtPictures = $step5 + 'Restoring Pictures'
$rtVideos = $step6 + 'Restoring Videos'
$rtMusic = $step7 + 'Restoring Music'
$rtGoogle = $step8 + 'Restoring Google Bookmarks'
$rtEdge = $step9 + 'Restoring Microsoft Edge Bookmarks'
$rtQuickLaunch = $step10 + 'Restoring Pinned Applications'
# AppData Restore Prompts
$rtOneNote = $step11 + 'Restoring OneNote Roaming and Local AppData'
$rtExcel = $step12 + 'Restoring Excel AppData'
$rtOutlook = $step13 + 'Restoring Outlook AppData'
$rtPowerPoint = $step14 + 'Restoring PowerPoint AppData'
$rtSignatures = $step15 + 'Restoring Outlook Signatures'
$rtSpelling = $step16 + 'Restoring Spelling AppData'
$rtTeams = $step17 + 'Restoring Teams Background AppData'
$rtWord = $step18 + 'Restoring Word AppData'
$rtThemes = $step19 + 'Restoring Windows Themes AppData'
# Other
$rtRegistry = $step20 + 'Importing TaskBand, OneNote, and Mapped Network Drives Registry Values'
$rtProfile = $step21 + 'Restoring Unique Profile Folders'
$rtCDrive = $step22 + 'Restoring Unique C Drive Folders'


# Implement GUI for Program
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
Add-Type -AssemblyName PresentationFramework
$msg = New-Object -ComObject Wscript.Shell

$objForm1 = New-Object System.Windows.Forms.Form
$objForm1.Backcolor="white"
$objForm1.StartPosition = "CenterScreen"
$objForm1.Size = New-Object System.Drawing.Size(500,500)
$objForm1.Text = "Data Migration v10"

$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(30,30)
$objLabel1.Size = New-Object System.Drawing.Size(200,40)
$objLabel1.Text = "Targeted Employee Number:" + $USERID
$objForm1.Controls.Add($objLabel1)

#$objTextBox1 = New-Object System.Windows.Forms.TextBox
#$objTextBox1.Location = New-Object System.Drawing.Size(30,70)
#$objTextBox1.Size = New-Object System.Drawing.Size(300,30)
#$objTextBox1.Text
#$objForm1.Controls.Add($objTextBox1)

$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(30,110)
$objLabel1.Size = New-Object System.Drawing.Size(200,40)
$objLabel1.Text = "Select an option: "
$objForm1.Controls.Add($objLabel1)

$objCombobox1 = New-Object System.Windows.Forms.Combobox
$objCombobox1.Location = New-Object System.Drawing.Size(30,150)
$objCombobox1.Size = New-Object System.Drawing.Size(300,30)
$objCombobox1.Height = 70
$objForm1.Controls.Add($objCombobox1)
# Removing OneDrive Options since they are never used
# [void] $objCombobox1.Items.Add("1) Backup User Data to OneDrive")
# [void] $objCombobox1.Items.Add("2) Restore User Data from OneDrive")
[void] $objCombobox1.Items.Add("Backup User Data to External Drive")
[void] $objCombobox1.Items.Add("Restore User Data from External Drive")
$objCombobox1.SelectedItem
$objCombobox1.Add_SelectedIndexChanged({ })

$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(30,190)
$objLabel2.Size = New-Object System.Drawing.Size(200,40)
$objLabel2.Text = "Drive Destination:"
$objForm1.Controls.Add($objLabel2)

$objCombobox2 = New-Object System.Windows.Forms.Combobox
$objCombobox2.Location = New-Object System.Drawing.Size(30,230)
$objCombobox2.Size = New-Object System.Drawing.Size(300,20)
$objCombobox2.Height = 70
$objForm1.Controls.Add($objCombobox2)
# Check the available existing paths and list them as an option
if (Test-Path -Path D:\) {
	[void] $objCombobox2.Items.Add("D")
}
if (Test-Path -Path E:\) {
	[void] $objCombobox2.Items.Add("E")
}
if (Test-Path -Path F:\) {
	[void] $objCombobox2.Items.Add("F")
}
$objCombobox2.SelectedItem
$objCombobox2.Add_SelectedIndexChanged({ })

$GoButton1 = New-Object System.Windows.Forms.Button
$GoButton1.Location = New-Object System.Drawing.Size(120,405)
$GoButton1.Size = New-Object System.Drawing.Size(100,30)
$GoButton1.Text = "Start"
$GoButton1.Name = "Start"
$GoButton1.DialogResult = "None"
# Implement the operation for all Options here
# Program features are displayed below:
# 1) Backup User Data (Onedrive)
# 2) Restore User Data (Onedrive)
# 3) Backup User Data (External (D:) Drive)
# 4) Restore User Data (External (D:) Drive)
# This function is responsible for the options that is selected
# We will use the null value to assume that the drive is on letter D
$GoButton1.Add_Click({
  # Close Form
  $objForm1.Close()
  $DRIVE = $null
  $CHOICE = $null
  # Not needed as we removed the drop down selection and automatically detected USERID
  #[string]$USERID = $objTextBox1.Text
  # Declaring variables by User selected options 
  $DRIVE = $objCombobox2.SelectedItem
  $CHOICE = $objCombobox1.SelectedItem
  $iserror = "False"

  $USERID = $USERID -as [string]
   if($USERID -like $null) {
	 [System.Windows.Forms.MessageBox]::Show("User ID is null, please do NOT run this program using admin credentials... Restart the program and try again") 
     $iserror = "True"
     $result = 2
   }  

  $DRIVE = $DRIVE -as [string]
   if($DRIVE -like $null) {
     $iserror = "True"
     $result = 3
  }

  if ($iserror -eq "True"){
     [System.Windows.Forms.MessageBox]::Show("Error") 
	 write-host $result
  }

  # Thanks to Dynamically Declared Variables, we can declare all variables here
  # setting dst drive variable, adding semicolon for formatting issues
  $dst = $DRIVE + ':'

  # srcPath variables
  $srcPath = $src + '\Users\' + $USERID + '\'
  $srcAppDataRoaming = $srcPath + 'AppData\Roaming\Microsoft\'
  # Setting src path variables
  $srcLotusDataPath = $src + "\Lotus\Notes\Data"
  $srcLotusIniPath = $src + "\Lotus\Notes\notes.ini"
  $srcDesktopPath = $srcPath + 'Desktop'
  $srcDocumentsPath = $srcPath + 'Documents'
  $srcFavoritesPath = $srcPath + 'Favorites'
  $srcDownloadsPath = $srcPath + 'Downloads'
  $srcPicturesPath = $srcPath + 'Pictures'
  $srcVideosPath = $srcPath + 'Videos'
  $srcMusicPath = $srcPath + 'Music'
  $srcGooglePath = $srcPath + 'AppData\Local\Google\Chrome\User Data\Default\Bookmarks'
  $srcEdgePath = $srcPath + 'AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks'
  $srcQuickLaunchPath = $srcPath + 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch'
  # src AppData variables
  $srcOneNoteLocal = $srcPath + 'AppData\Local\Microsoft\OneNote'
  $srcOneNoteRoaming = $srcAppDataRoaming + 'OneNote'
  $srcExcel = $srcAppDataRoaming + 'Excel'
  $srcOutlook = $srcAppDataRoaming + 'Outlook'
  $srcPowerPoint = $srcAppDataRoaming + 'PowerPoint'
  $srcSignatures = $srcAppDataRoaming + 'Signatures'
  $srcSpelling = $srcAppDataRoaming + 'Spelling'
  $srcTeamsBackgrounds = $srcAppDataRoaming + 'Teams\Backgrounds'
  $srcWord = $srcAppDataRoaming + 'Word'
  $srcThemes = $srcAppDataRoaming + 'Windows\Themes'
  
  # main dstPath 
  $dstPath = $dst + '\DataBackup\' + $USERID + '\'
  $dstAppDataPath = $dstPath + 'AppData\'
  # Setting dst path variables
  $dstLotusDataPath = $dstPath + "Lotus\Notes\Data"
  $dstLotusIniPath = $dstPath + "Lotus\Notes"
  $dstDesktopPath = $dstPath + 'Desktop'
  $dstDocumentsPath = $dstPath + 'Documents'
  $dstFavoritesPath = $dstPath + 'Favorites'
  $dstDownloadsPath = $dstPath + 'Downloads'
  $dstPicturesPath = $dstPath + 'Pictures'
  $dstVideosPath = $dstPath + 'Videos'
  $dstMusicPath = $dstPath + 'Music'
  $dstGooglePath = $dstPath + 'Google Bookmarks'
  $dstEdgePath = $dstPath + 'Edge Bookmarks'
  $dstQuickLaunchPath = $dstPath + 'Quick Launch'
  $dstRegistryPath = $dstPath + 'Registry'
  $dstCDrive = $dstPath + 'C_Drive'
  $dstProfileFolders = $dstPath + 'Profile Folders'
  # dst AppData variables
  $dstOneNoteLocal = $dstAppDataPath + 'OneNote Local'
  $dstOneNoteRoaming = $dstAppDataPath + 'OneNote Roaming'
  $dstExcel = $dstAppDataPath + 'Excel'
  $dstOutlook = $dstAppDataPath + 'Outlook'
  $dstPowerPoint = $dstAppDataPath + 'PowerPoint'
  $dstSignatures = $dstAppDataPath + 'Signatures'
  $dstSpelling = $dstAppDataPath + 'Spelling'
  $dstTeamsBackgrounds = $dstAppDataPath + 'Teams Backgrounds'
  $dstWord = $dstAppDataPath + 'Word'
  $dstThemes = $dstAppDataPath + 'Themes'   
  # Log Variables
  $startBackupLog = $dstPath + 'Backup_Log.txt'
  $startRestoreLog = $dstPath + 'Restore_Log.txt'
  $externalBackupLog = '/log+:' + $dst + '\DataBackup\' + $USERID + '\Backup_Log.txt'
  $externalRestoreLog = '/log+:' + $dst + '\DataBackup\' + $USERID + '\Restore_Log.txt'
  $appendBackupLog = "/log:$dst\DataBackup\$USERID\Backup_Log.txt"
  $appendRestoreLog =  "/log:$dstDataBackup\$USERID\Restore_Log.txt"
  # Registry Paths
  $regTaskband = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband'
  $regOneNote = 'HKCU\SOFTWARE\Microsoft\Office\16.0\OneNote\OpenNotebooks'
  $regNetwork = 'HKCU\Network'
  $taskbarFile = "$dst\DataBackup\$USERID\Registry\Taskbar_BackUp.reg"
  $onenoteFile = "$dst\DataBackup\$USERID\Registry\OneNote.reg"
  $networkdriveFile = "$dst\DataBackup\$USERID\Registry\Network_Drives.reg"

  # Calculating the amount of data that will be copied
  # Factor in OneDrive Potential Conflicts such as Desktop, Documents, and Pictures Folders being on the Cloud and not Local so they 
  # do not exist, this makes these lines of codes crash the program, do an if statement and see if it exists FIRST, then get the size
  # Implement Later once the rest of the program works
  # Total amount to be copied
  # $szTotal = 0;
  # if (Test-Path -Path $srcDesktopPath) {
  # $szDesktop = [Math]::Round(((Get-ChildItem $srcDesktopPath -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB),2)
  #}else {
  #$szDesktop = 0;
  #}
  #if (Test-Path -Path $srcDocumentsPath) {
  #$szDocuments = [Math]::Round(((Get-ChildItem $srcDocumentsPath -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB),2)
  #}else {
  #$szDocuments = 0;
  #}
  #if (Test-Path -Path $srcPicturesPath) {
  #$szPictures = [Math]::Round(((Get-ChildItem $srcPicturesPath -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB),2)
  #}else {
  #$szPictures = 0;
  #}
  

  # Enter if statements here since switch statments aint that good to 
  # handle a lot of commands 
  # Case 1 - Data Backup to OneDrive
  #if( "1) Backup User Data to OneDrive" -eq $Choice ) {
	  # [System.Windows.MessageBox]::Show('Do you want to proceed?', 'Confirm', 'YesNoCancel','Error')
	  #New-BurntToastNotification -AppLogo "$dst\UL Icon.PNG" -Text "Initializing Data Backup Procedure"
	  #$result = 1
  #}
  
  # Case 2 - Data Restore from OneDrive
   #if( "2) Restore User Data from OneDrive" -eq $Choice ) {
	  #$result = 1
  #}
  # Case 3 - Data Backup to External Drive
  # Find a way around the timing for the prompts
   if( "Backup User Data to External Drive" -eq $Choice ) {
	   $result = 1
	   
	   write-host "Starting Data Backup Program"
       write-host 'Creating Necessary SubDirectories'
	   # PowerShell runs into an error when the Directory already exists, going to delete the Targeted User Directory
       # Delete takes too long as it removes all files, going to check each directory if it exists
       # Changing to just skipped Path Creation, must do for all New-Item actions
       $pathTest = Test-Path -Path ($dst + '\DataBackup')
       $databackupFolder = $dst + '\DataBackup'
	   if ($pathTest -eq $false) {
		   # Make Directory DataBackup Folder
           New-Item -Path $databackupFolder -ItemType Directory -ea 0
		   # If true then no need to make Directory
	   }
       $pathTest = Test-Path -Path ($dstPath)
       if ($pathTest -eq $false) {
           # Make Directory for User ID in DataBackup Folder
           New-Item -Path $dstPath -ItemType Directory -ea 0
       } 
	   # Check if Lotus exists
	   if (Test-Path -Path $Lotus) {
	       # Create Directories for Lotus in External Drive
	       New-Item -Path "$dst\DataBackup\$USERID\Lotus" -ItemType Directory -ea 0
	       New-Item -Path "$dst\DataBackup\$USERID\Lotus\Notes" -ItemType Directory -ea 0
	       New-Item -Path "$dst\DataBackup\$USERID\Lotus\Notes\Data" -ItemType Directory -ea 0
		   # Start Copying Lotus Files
		   write-host $bkLotus
		   robocopy $srcLotusDataPath $dstLotusDataPath $roboOP /log:"$dst\DataBackup\$USERID\Backup_Log.txt"
           Xcopy $srcLotusIniPath $dstLotusIniPath /Y
	    } else {
				write-host 'Lotus does not exist, moving to Desktop'
			}
		# Continue to desktop
		# Make all Directories
		New-Item -Path $dstDesktopPath -ItemType Directory -ea 0
		New-Item -Path $dstDocumentsPath -ItemType Directory -ea 0
		New-Item -Path $dstFavoritesPath -ItemType Directory -ea 0
		New-Item -Path $dstDownloadsPath -ItemType Directory -ea 0
		New-Item -Path $dstPicturesPath -ItemType Directory -ea 0
		New-Item -Path $dstVideosPath -ItemType Directory -ea 0
		New-Item -Path $dstMusicPath -ItemType Directory -ea 0
		New-Item -Path $dstGooglePath -ItemType Directory -ea 0
        New-Item -Path $dstEdgePath -ItemType Directory -ea 0
		New-Item -Path $dstQuickLaunchPath -ItemType Directory -ea 0
        # Need to Delete this specific folder as it already exists and causes an issue
        # with registry keys
        if (Test-Path -Path $dstRegistryPath)
        {
            Remove-Item -Force -Recurse $dstRegistryPath
        }
		New-Item -Path $dstRegistryPath -ItemType Directory -ea 0
		New-Item -Path $dstCDrive -ItemType Directory -ea 0
		New-Item -Path $dstProfileFolders -ItemType Directory -ea 0
		# AppData Directories
		New-Item -Path $dstAppDataPath -ItemType Directory -ea 0
		New-Item -Path $dstOneNoteLocal -ItemType Directory -ea 0
		New-Item -Path $dstOneNoteRoaming -ItemType Directory -ea 0
		New-Item -Path $dstExcel -ItemType Directory -ea 0
		New-Item -Path $dstOutlook -ItemType Directory -ea 0
		New-Item -Path $dstPowerPoint -ItemType Directory -ea 0
		New-Item -Path $dstSignatures -ItemType Directory -ea 0
		New-Item -Path $dstTeamsBackgrounds -ItemType Directory -ea 0
		New-Item -Path $dstWord -ItemType Directory -ea 0
		New-Item -Path $dstThemes -ItemType Directory -ea 0
		# Begin starting copy process
		# If lotus is detected, APPEND the log instead of overwriting it
		# If statement below should resolve this issue
		write-host $bkDesktop
		if (Test-Path -Path $Lotus)
		{
			robocopy $srcDesktopPath $dstDesktopPath $roboOP $externalBackupLog
		}
		robocopy $srcDesktopPath $dstDesktopPath $roboOP $star
		write-host $bkDocuments
		robocopy $srcDocumentsPath $dstDocumentsPath $roboOP $externalBackupLog
		write-host $bkFavorites
		robocopy $srcFavoritesPath $dstFavoritesPath $roboOP $externalBackupLog
		write-host $bkDownloads
		robocopy $srcDownloadsPath $dstDownloadsPath $roboOP $externalBackupLog
		write-host $bkPictures
		robocopy $srcPicturesPath $dstPicturesPath $roboOP $externalBackupLog
		write-host $bkVideos
		robocopy $srcVideosPath $dstVideosPath $roboOP $externalBackupLog
		write-host $bkMusic
		robocopy $srcMusicPath $dstMusicPath $roboOP $externalBackupLog
		write-host $bkGoogle
		Xcopy $srcGooglePath $dstGooglePath /Y /Q /I
        write-host $bkEdge
        Xcopy $srcEdgePath $dstEdgePath /Y /Q /I
		write-host $bkQuickLaunch
		robocopy $srcQuickLaunchPath $dstQuickLaunchPath $roboOP $externalBackupLog
		# AppData Copy Process
		write-host $bkOneNote
		robocopy $srcOneNoteLocal $dstOneNoteLocal $roboOP $externalBackupLog
		robocopy $srcOneNoteRoaming $dstOneNoteRoaming $roboOP $externalBackupLog
		write-host $bkExcel
		robocopy $srcExcel $dstExcel $roboOP $externalBackupLog
		write-host $bkOutlook
		robocopy $srcOutlook $dstOutlook
		write-host $bkPowerPoint
		robocopy $srcPowerPoint $dstPowerPoint $roboOP $externalBackupLog
		write-host $bkSignatures
		robocopy $srcSignatures $dstSignatures $roboOP $externalBackupLog
		write-host $bkSpelling
		robocopy $srcSpelling $dstSpelling $roboOP $externalBackupLog
		write-host $bkTeams
		robocopy $srcTeamsBackgrounds $dstTeamsBackgrounds $roboOP $externalBackupLog
		write-host $bkWord
		robocopy $srcWord $dstWord $roboOP $externalBackupLog
		write-host $bkThemes
		robocopy $srcThemes $dstThemes $roboOP $externalBackupLog
		# Come back to manually add registry values
        # Exporting 
		write-host $bkRegistry
		reg export $regTaskband $taskbarFile
		reg export $regOneNote $onenoteFile
		reg export $regNetwork $networkdriveFile
		# Copying Unique Drive Folders
		write-host $bkProfile
		robocopy $srcPath $dstProfileFolders $roboOP /A-:SH $profileExempt $externalBackupLog
        write-host $bkCDrive
		robocopy $src\ $dstCDrive $roboOP /A-:SH $cExempt $externalBackupLog
		write-host "`nProgram Sequence has been completed"
		# Open Backup Log
		$msg.Popup("Program has been completed, please check log file for any issues", 100000, "End of Program", 64)
		Start-Process -FilePath $startBackupLog
	
  }
  # Case 4 - Data Restore from External Drive
   if( "Restore User Data from External Drive" -eq $Choice ) {
	  #$result = 1
      write-host "Starting Data Restoration"
      # No need to create any directories so skip checking
      # 1st check if Lotus exists so we can restore it
      # Creating a temp variable to resolve data being corrupted from copying over because it would try to 
      # path to a file instead of a directory, this notesPath solves this issue
      $notesPath = $src + "Lotus\Notes"
      if (Test-Path -Path $Lotus) {
      write-host $rtLotus
      Xcopy $dstLotusDataPath $srcLotusDataPath $xcopyOP
      Xcopy $dstLotusIniPath $notesPath $xcopyOP
      } else {
      write-host 'Lotus Does not Exist, moving on to Desktop'
      }
      write-host $rtDesktop
      robocopy $dstDesktopPath $srcDesktopPath $roboOP /log:"$dst\DataBackup\$USERID\Restore_Log.txt"
      write-host $rtDocuments
      robocopy $dstDocumentsPath $srcDocumentsPath $roboOP $externalRestoreLog
      write-host $rtFavorites
      robocopy $dstFavoritesPath $srcFavoritesPath $roboOP $externalRestoreLog
      write-host $rtDownloads
      robocopy $dstDownloadsPath $srcDownloadsPath $roboOP $externalRestoreLog
      write-host $rtPictures
      robocopy $dstPicturesPath $srcPicturesPath $roboOP $externalRestoreLog
      write-host $rtVideos
      robocopy $dstVideosPath $srcVideosPath $roboOP $externalRestoreLog
      write-host $rtMusic
      robocopy $dstMusicPath $srcMusicPath $roboOP $externalRestoreLog
      write-host $rtGoogle
      $defaultGooglePath = $srcPath + 'AppData\Local\Google\Chrome\User Data\Default'
      robocopy $dstGooglePath $defaultGooglePath $roboOP $externalRestoreLog
      write-host $rtEdge
      $defaultEdgePath = $srcPath + 'AppData\Local\Microsoft\Edge\User Data\Default'
      robocopy $dstEdgePath $defaultEdgePath $roboOP $externalRestoreLog
      write-host $rtQuickLaunch
      robocopy $dstQuickLaunchPath $srcQuickLaunchPath $roboOP $externalRestoreLog
      # AppData Restoration Process
      write-host $rtOneNote
      xcopy $dstOneNoteLocal $srcOneNoteLocal $xcopyOP
      xcopy $dstOneNoteRoaming $srcOneNoteRoaming $xcopyOP
      write-host $rtExcel
      xcopy $dstExcel $srcExcel $xcopyOP
      write-host $rtOutlook
      xcopy $dstOutlook $srcOutlook $xcopyOP
      write-host $rtPowerPoint
      xcopy $dstPowerPoint $srcPowerPoint $xcopyOP
      write-host $rtSignatures
      xcopy $dstSignatures $srcSignatures $xcopyOP
      write-host $rtSpelling
      xcopy $dstSpelling $srcSpelling $xcopyOP
      write-host $rtTeams
      xcopy $dstTeamsBackgrounds $srcTeamsBackgrounds $xcopyOP
      write-host $rtWord
      xcopy $dstWord $srcWord $xcopyOP
      write-host $rtThemes
      xcopy $dstThemes $srcThemes $xcopyOP
      # Import Registry Keys/Values
      write-host $rtRegistry
      reg import "$dst\DataBackup\$USERID\Registry\Taskbar_BackUp.reg"
      reg import "$dst\DataBackup\$USERID\Registry\OneNote.reg"
      reg import "$dst\DataBackup\$USERID\Registry\Network_Drives.reg"

      # Attempt to Restore Unique Folders in C and User Profile
      write-host $rtProfile
      robocopy $dstProfileFolders $srcPath $roboOP $externalRestoreLog
      # write-host $rtCDrive
      # robocopy $dstCDrive "$src\" $externalRestoreLog

      # Restart Taskbar
      write-host "`nRestarting Taskbar"
      Stop-Process -Name explorer -Force
      # Open Restore Log
      $msg.Popup("Program has been completed, please manually restore C_Drive Folder to finish Restoration, check log file for any issues", 100000, "End of Program", 64)
      write-host "`nPlease remember to restore C_Drive Folder"
	  Start-Process -FilePath $startRestoreLog
      
  }
  # Not needed but left just in case
  #if ($iserror -eq "False"){
     #[System.Windows.Forms.MessageBox]::Show("$result","Result:") 
  #}
})

$objForm1.Controls.Add($GoButton1)

$CancelButton1 = New-Object System.Windows.Forms.Button
$CancelButton1.Location = New-Object System.Drawing.Size(30,405)
$CancelButton1.Size = New-Object System.Drawing.Size(100,30)
$CancelButton1.Text = "Exit"
$CancelButton1.Name = "Exit"
$CancelButton1.DialogResult = "Cancel"
$CancelButton1.Add_Click({$objForm1.Close()})
$objForm1.Controls.Add($CancelButton1)

[void] $objForm1.ShowDialog()

