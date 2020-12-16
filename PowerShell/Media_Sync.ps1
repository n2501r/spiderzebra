###############################################################################
### Variables
###############################################################################
$Time = Get-Date -Format "MM-dd-yyyy-HHmmss" ### Getting the date and formatting it to be used in the transcript file name
$Transcript_File = $PSScriptRoot + "\Transcripts\Media_Sync_$Time.txt" ### Location for the transcript file
### Zeroing out the count variables to ensure the numbers are accurate every time
$Copied_Count = 0
$Moved_Count = 0 
$Duplicate_Count = 0
$Skipped_Count = 0
$Created_Dir_Count = 0
$Copied_Hash_Count = 0
$Moved_Hash_Count = 0
$Error_Count = 0
$Not_Photo_Video = 0

$Final_Results = @()

### Cleaning up the old transcripts, only keeping the last 300 files
Get-ChildItem $PSScriptRoot -Recurse -Include Media_Sync*.txt | Where-Object { -not $_.PsIsContainer } | Sort-Object LastWriteTime -desc | Select-Object -Skip 300 | Remove-Item -Force

### Starting the transcript
Start-Transcript -Path $Transcript_File -Force

###############################################################################
### Function to display a folder dialog menu, the message parameter is used to give a brief description of what folder to select
###############################################################################
Function Folder_Dialog ($message) {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Folder_Browser = New-Object System.Windows.Forms.FolderBrowserDialog
    $Folder_Browser.Description = "$message" ### Adds a custom message to the folder dialog menu
    $Folder_Browser.UseDescriptionForTitle = $true ### Puts the custom message in the title

    ### if the OK button is clicked the selected folder path will be returned
    if ($Folder_Browser.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true })) -eq "OK") {
        Return $Folder_Browser.SelectedPath
    }
    ### Otherwise the script will exit and send an error to screen
    else {
        Write-Host "No folder selected, script will now exit..." -BackgroundColor Red -ForegroundColor White
        Exit
    }
}
###############################################################################
### Function to gather file metadata and pass all photo and video files to be copied or moved
###############################################################################
Function Get-FileMetaData ($folder) {
    ### Zeroing out the count variables to ensure the numbers are accurate every time
    $Photos_Detected = 0
    $Videos_Detected = 0
    $Folders_Detected = 0
    $Others_Detected = 0
    ### creating an empty array to hold file metadata facts
    $customobj = @()
    #########################################################################################
    ### We use the Shell.Application COM object to gather file metada
    $ShellCOM = New-Object -ComObject Shell.Application
    $FolderCOM = $ShellCOM.namespace($folder)
    ### Going through each file in a given folder
    foreach ($File in $FolderCOM.items()) {
        #########################################################################################
        ### if the file type is Picture it will be processed accordingly
        if ($FolderCOM.getDetailsOf($File, 11) -eq "Picture") {
            $Photos_Detected++ ### Counter for files that are photos
            $FileSize_String = [string]$($File.Size) ### Converting file size from an integer to string
            $Last5_FileSize = $FileSize_String.Substring( [math]::Max( 0, $FileSize_String.Length - 5 ) ) ### Pulling out the last 5 digits of the file size string, this will be added to the file name for uniqueness
            ### The date taken field is the 12th object in the metadata, if that exists the script will convert it to a properly formatted date
            if ($FolderCOM.getDetailsOf($File, 12)) {
                $temp_date = $($FolderCOM.getDetailsOf($File, 12)) -replace '\p{Cf}' ### Removing invisible control characters from the date taken
                $DateTime = [DateTime]::ParseExact($temp_date, "M/d/yyyy h:mm tt", [System.Globalization.DateTimeFormatInfo]::InvariantInfo, [System.Globalization.DateTimeStyles]::None) ### Converts to proper datetime
                $DateFormated = $DateTime.ToString('ddMMMyyyy_HHmm') ### Sets the format for the file name (ie 01Jun2003_1134)
                ### Putting together the destination folder structure, using the 4 digit year and 3 digit month name (ie ..\2019\Oct)
                $Directory = $DestinationDir + "\" + $DateTime.ToString('yyyy') + "\" + $DateTime.ToString('MMM')
            }
            else {
                ### if the more accurate date taken field is blank then the script will use the Modify date instead
                $DateFormated = $File.ModifyDate.ToString('ddMMMyyyy_HHmm')
                ### Putting together the destination folder structure, using the 4 digit year and 3 digit month name (ie ..\2019\Oct)
                $Directory = $DestinationDir + "\" + $File.ModifyDate.ToString('yyyy') + "\" + $File.ModifyDate.ToString('MMM')
            }
            ### Updating custom PS object to hold file facts
            $customobj += [PSCustomObject] @{
                Destination_Filename = $DateFormated + $Last5_FileSize
                Destination_Path     = $Directory
                Extension            = $($FolderCOM.getDetailsOf($File, 164))
                Source_Path          = $File.Path
            }
            ### Clearing variables to keep the script clean  
            Clear-Variable FileSize_String, Last5_FileSize, temp_date, DateTime, DateFormated, Directory -ErrorAction SilentlyContinue
        }
        #########################################################################################
        ### if the file type is Video it will be processed accordingly
        elseif ($FolderCOM.getDetailsOf($File, 11) -eq "Video") {
            $Videos_Detected++ ### Counter for files that are Videos
            $FileSize_String = [string]$($File.Size) ### Converting file size from an integer to string
            $Last5_FileSize = $FileSize_String.Substring( [math]::Max( 0, $FileSize_String.Length - 5 ) ) ### Pulling out the last 5 digits of the file size string, this will be added to the file name for uniqueness
            ### The media created field is the 208th object in the metadata, if that exists the script will convert it to a properly formatted date
            if ($FolderCOM.getDetailsOf($File, 208)) {
                $temp_date = $($FolderCOM.getDetailsOf($File, 208)) -replace '\p{Cf}' ### Removing invisible control characters from the date taken
                $DateTime = [DateTime]::ParseExact($temp_date, "M/d/yyyy h:mm tt", [System.Globalization.DateTimeFormatInfo]::InvariantInfo, [System.Globalization.DateTimeStyles]::None) ### Converts to proper datetime
                $DateFormated = $DateTime.ToString('ddMMMyyyy_HHmm') ### Sets the format for the file name (ie 01Jun2003_1134)
                ### Putting together the destination folder structure, using the 4 digit year and 3 digit month name (ie ..\2019\Oct)
                $Directory = $DestinationDir + "\" + $DateTime.ToString('yyyy') + "\" + $DateTime.ToString('MMM')
            }
            else {
                ### if the more accurate media created field is blank then the script will use the Modify date instead
                $DateFormated = $File.ModifyDate.ToString('ddMMMyyyy_HHmm')
                ### Putting together the destination folder structure, using the 4 digit year and 3 digit month name (ie ..\2019\Oct)
                $Directory = $DestinationDir + "\" + $File.ModifyDate.ToString('yyyy') + "\" + $File.ModifyDate.ToString('MMM')
            }
            ### Updating custom PS object to hold file facts
            $customobj += [PSCustomObject] @{
                Destination_Filename = $DateFormated + $Last5_FileSize
                Destination_Path     = $Directory
                Extension            = $($FolderCOM.getDetailsOf($File, 164))
                Source_Path          = $File.Path
            }
            ### Clearing variables to keep the script clean  
            Clear-Variable FileSize_String, Last5_FileSize, temp_date, DateTime, DateFormated, Directory -ErrorAction SilentlyContinue
        }
        #########################################################################################
        ### if the type is folder than it will be ignored
        elseif ($FolderCOM.getDetailsOf($File, 11) -eq "Folder") {
            $Folders_Detected++ ### Counter for folders
        }
        #########################################################################################
        ### if the type is not folder, picture or video than it will be skipped
        else {
            $Others_Detected++ ### counter for other file types detected
            $Typeofitem = $($FolderCOM.getDetailsOf($File, 11)) ### Grabbing they type of file
            ### if the type is null or empty than a message will be displayed to the terminal
            if ($null -eq $Typeofitem -or $Typeofitem -eq "") {
                write-host "[Type of file could not be identified] ($($File.Path)) will NOT be included in the copy or move process" -ForegroundColor Yellow
            }
            ### if the type is populated it will be displayed to the terminal
            else {
                write-host "[$Typeofitem Detected] ($($File.Path)) will NOT be included in the copy or move process" -ForegroundColor Yellow
            }
            $Script:Not_Photo_Video++  ### Counter for files that are not a photo or video, this is used for the final rollup at the end
            ### Updating custom PS object to hold final results
            $Script:Final_Results += [PSCustomObject] @{
                Source      = $File.Path
                Destination = "N/A"
                Results     = "Not a photo or video, skipped"
            }
            ### Clearing variables to keep the script clean  
            Clear-Variable Typeofitem -ErrorAction SilentlyContinue
        }
        #########################################################################################
    } 
    #########################################################################################
    ### Sending the results to the terminal and transcript for each folder processed
    Write-host "Photos detected: $Photos_Detected" -ForegroundColor Cyan
    Write-host "Videos detected: $Videos_Detected" -ForegroundColor Cyan
    Write-host "Folders detected: $Folders_Detected" -ForegroundColor Cyan
    Write-host "Other files detected: $Others_Detected" -ForegroundColor Cyan
    ### Returning the file metadata results
    Return $customobj
} 
###############################################################################
### Function to copy or move photos / videos, the two supported parameters are copy and move
###############################################################################
Function Photos_Videos ($Mode) {
    ### creating empty arrays
    $Dirs_to_search = @()
    $Metadata = @()
    ### Calling the folder_dialog function for the source directory
    Write-Host "Select the SOURCE directory in the folder dialog window" -ForegroundColor Yellow
    $SourceDir = Folder_Dialog -Message "Select the SOURCE directory"
    ### Calling the folder_dialog function for the destination directory
    Write-Host "Select the DESTINATION directory in the folder dialog window" -ForegroundColor Yellow
    $DestinationDir = Folder_Dialog -Message "Select the DESTINATION directory"
    ### if the mode given was move then an additional folder dialog will be executed for the duplicates directory
    if ($Mode -eq "move") {
        ### if the user selects the same directory for both the source and destination the script will exit as this will cause issues
        if ($SourceDir -eq $DestinationDir) {
            Write-Host "Source and Destination directories are the same, script is exiting..." -ForegroundColor Red
            Exit
        }
        Write-Host "In the folder dialog window select the directory you want to move DUPLICATES to" -ForegroundColor Yellow
        $DuplicateDir = Folder_Dialog -Message "Select the DUPLICATE files directory"
        ### if the user selects the same directory for both the source and duplicate the script will exit as this will cause issues
        if ($SourceDir -eq $DuplicateDir) {
            Write-Host "Source and Duplicate directories are the same, script is exiting..." -ForegroundColor Red
            Exit
        }
        ### if the user selects the same directory for both the duplicate and destination the script will exit as this will cause issues
        if ($DestinationDir -eq $DuplicateDir) {
            Write-Host "Destination and Duplicate directories are the same, script is exiting..." -ForegroundColor Red
            Exit
        }
    }
    #########################################################################################
    ### Confirmation that the directories selected are correct
    Write-Host "`nConfirm that the following variables are correct" -ForegroundColor DarkYellow
    Write-host "Source Directory: $SourceDir" -ForegroundColor Cyan
    Write-host "Destination Directory: $DestinationDir" -ForegroundColor Cyan
    if ($Mode -eq "move") {
        Write-host "Duplicate Files Directory: $DuplicateDir" -ForegroundColor Cyan
    }
    $Confirm = Read-Host "Do you want to continue? (Y/N)"
    #########################################################################################
    ## if the user hit y or Y then the script will continue
    if ($Confirm.ToUpper() -eq "Y") {
        Write-Host "`n ---- Please wait while the script gathers the folder(s) to be searched ---- `n" -ForegroundColor Cyan
        ### Creating a custom powershell object to store the folders to be searched
        $Dirs_to_search += [PSCustomObject] @{Path = $SourceDir } ### Adding the selected directory
        Get-Childitem -Path $SourceDir -Recurse -Directory | ForEach-Object { ### Recursively finding all directories in the selected directory
            $Dirs_to_search += [PSCustomObject] @{Path = $_.FullName } ### Adding any directories found
        }
        $Count_dirs = $Dirs_to_search | Measure-Object | ForEach-Object { $_.Count } ### Counting how many directories found
        Write-Host "[$Count_dirs folder(s) identified]" -ForegroundColor DarkGreen
        Write-Host "`n ---- Please wait while the script gathers file metadata from the folder(s) identified ---- `n" -ForegroundColor Cyan
        ### Calling the Get-FileMetaData function for each directory in the PS object in order to gather file metadata
        $Dirs_to_search | ForEach-Object {
            Write-Host "[$($_.Path)]" -ForegroundColor DarkGreen
            $Metadata += Get-FileMetaData -folder $_.Path
        }
        Write-Host "`n ---- Please wait while the script begins the $Mode process ---- `n" -ForegroundColor Cyan
        #########################################################################################
        ### Going through each file returned from Get-FileMetaData and either copying or moving based on the given option
        $Metadata | ForEach-Object {
            #########################################################################################
            ### if the month or year directory isn't there it will create one
            if (-NOT(Test-Path $_.Destination_Path)) {
                try {
                    New-Item $_.Destination_Path -Type directory | Out-Null
                    $Created_Dir_Count++  ### Counter for directories created
                }
                Catch [Exception] {
                    ### if the command inside the try statement fails the error will be outputted
                    $errormessage = $_.Exception.Message
                    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                    $Error_Count++ ### Counter for errors handled
                }
            }
            #########################################################################################
            ### Putting the full path of the new file name and then checking to see if the file already exists
            $Destination_File = $_.Destination_Path + "\" + $_.Destination_Filename + $_.Extension
            if (Test-Path -Path $Destination_File) {
                #########################################################################################
                ### if the destination file matches the source file the script will not take any action as the file already exists
                if ($Destination_File -eq $_.Source_Path) {
                    Write-Host "[$($_.Source_Path)] The source and destination are the same file, skipping this file..." -ForegroundColor Gray
                    $Skipped_Count++ ### Counter for skipped files
                }
                else {
                    #########################################################################################
                    ### if the file does already exist then it will check the file hash to see if its the same file or not
                    try {
                        $OriginalHash = (Get-FileHash -Path $Destination_File -Algorithm MD5).hash ### The file hash of the file that was already in the destination
                        $NewFileHash = (Get-FileHash -Path $_.Source_Path -Algorithm MD5).hash ### The file hash of the new file that is set to be copied or moved over
                    }
                    Catch [Exception] {
                        ### if the command inside the try statement fails the error will be outputted
                        $errormessage = $_.Exception.Message
                        Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                        $Error_Count++ ### Counter for errors handled
                    }
                    #########################################################################################
                    if ($OriginalHash -ne $NewFileHash) {
                        ### if the two file hashes don't match then the script will keep both files
                        ### Putting together the new file name and path to include the first 4 characters of the hash to ensure the file name is unique         
                        $NewHashPath = $_.Destination_Path + "\" + $_.Destination_Filename + "~" + ($NewFileHash.subString(0, [System.Math]::Min(4, $NewFileHash.Length))) + $_.Extension
                        #########################################################################################
                        if ($Mode -eq "copy") {
                            ### if the mode selected was copy it will execute the copy-item command
                            if (Test-Path -Path $NewHashPath) {
                                ### if the file already exists, the script will skip this file
                                Write-Host "[$($_.Source_Path)] Already copied, nothing to do here..." -ForegroundColor Gray
                                $Skipped_Count++ ### Counter for skipped files
                                ### Updating custom PS object to hold final results
                                $Final_Results += [PSCustomObject] @{
                                    Source      = $_.Source_Path 
                                    Destination = $NewHashPath
                                    Results     = "Skipped, already copied"
                                }
                            }
                            else {
                                ### if the file doesn't exist already, the script will copy it
                                try {
                                    Write-Host "[$($_.Source_Path)] File name already exists in destination but hash values don't match, copying to $NewHashPath" -ForegroundColor DarkGreen
                                    Copy-Item -Path $_.Source_Path -Destination $NewHashPath -ErrorAction Stop
                                    $Copied_Hash_Count++ ### Counter for copied files (same name different hash)
                                    ### Updating custom PS object to hold final results
                                    $Final_Results += [PSCustomObject] @{
                                        Source      = $_.Source_Path 
                                        Destination = $NewHashPath
                                        Results     = "Copied (hash added)"
                                    }
                                }
                                Catch [Exception] {
                                    ### if the command inside the try statement fails the error will be outputted
                                    $errormessage = $_.Exception.Message
                                    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                                    $Error_Count++ ### Counter for errors handled
                                    ### Updating custom PS object to hold final results
                                    $Final_Results += [PSCustomObject] @{
                                        Source      = $_.Source_Path 
                                        Destination = $NewHashPath
                                        Results     = "ERROR, $errormessage"
                                    }
                                }
                            }
                        }
                        #########################################################################################
                        if ($Mode -eq "move") {
                            ### if the mode selected was move it will execute the move-item command
                            if (Test-Path -Path $NewHashPath) {
                                ### if the file already exists, the script will move the file to the duplicates directory
                                try {
                                    Write-host "[$($_.Source_Path)] Already present, moving to duplicates directory"
                                    Move-Item -Path $_.Source_Path -Destination $DuplicateDir -Force -ErrorAction Stop
                                    $Duplicate_Count++ ### Counter for duplicate files
                                    ### Updating custom PS object to hold final results
                                    $Final_Results += [PSCustomObject] @{
                                        Source      = $_.Source_Path 
                                        Destination = $DuplicateDir
                                        Results     = "Moved to duplicates directory"
                                    }
                                }
                                Catch [Exception] {
                                    ### if the command inside the try statement fails the error will be outputted
                                    $errormessage = $_.Exception.Message
                                    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                                    $Error_Count++ ### Counter for errors handled
                                    ### Updating custom PS object to hold final results
                                    $Final_Results += [PSCustomObject] @{
                                        Source      = $_.Source_Path 
                                        Destination = $DuplicateDir
                                        Results     = "ERROR, $errormessage"
                                    }
                                }
                            }
                            else {
                                ### if the file doesn't exist already, the script will move it to the destination directory
                                try {
                                    Write-Host "[$($_.Source_Path)] File name already exists in destination but hash values don't match, moving to $NewHashPath" -ForegroundColor DarkGreen
                                    Move-Item -Path $_.Source_Path -Destination $NewHashPath -Force -ErrorAction Stop
                                    $Moved_Hash_Count++ ### Counter for moved files
                                    $Final_Results += [PSCustomObject] @{
                                        Source      = $_.Source_Path 
                                        Destination = $NewHashPath
                                        Results     = "Moved (hash added)"
                                    }
                                }
                                Catch [Exception] {
                                    ### if the command inside the try statement fails the error will be outputted
                                    $errormessage = $_.Exception.Message
                                    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                                    $Error_Count++ ### Counter for errors handled
                                    ### Updating custom PS object to hold final results
                                    $Final_Results += [PSCustomObject] @{
                                        Source      = $_.Source_Path 
                                        Destination = $NewHashPath
                                        Results     = "ERROR, $errormessage"
                                    }
                                }
                            }
                        }
                        #########################################################################################
                    }
                    ######################################################################################### 
                    else {
                        ### if the two file hashes DO match then the script will skip the file or move it to the duplicates directory
                        if ($Mode -eq "copy") {
                            ### if the mode selected was copy then the script will skip this file
                            Write-Host "[$($_.Source_Path)] Already copied, nothing to do here..." -ForegroundColor Gray
                            $Skipped_Count++ ### Counter for skipped files
                            ### Updating custom PS object to hold final results
                            $Final_Results += [PSCustomObject] @{
                                Source      = $_.Source_Path 
                                Destination = $Destination_File
                                Results     = "Skipped, already copied"
                            }
                        }
                        if ($Mode -eq "move") {
                            ### if the mode selected was move then the script will move the file to the duplicates directory
                            try {
                                Write-host "[$($_.Source_Path)] Already present, moving to duplicates directory"
                                Move-Item -Path $_.Source_Path -Destination $DuplicateDir -Force
                                $Duplicate_Count++ ### Counter for duplicate files
                                ### Updating custom PS object to hold final results
                                $Final_Results += [PSCustomObject] @{
                                    Source      = $_.Source_Path 
                                    Destination = $DuplicateDir
                                    Results     = "Moved to duplicates directory"
                                }
                            }
                            Catch [Exception] {
                                ### if the command inside the try statement fails the error will be outputted
                                $errormessage = $_.Exception.Message
                                Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                                $Error_Count++ ### Counter for errors handled
                                ### Updating custom PS object to hold final results
                                $Final_Results += [PSCustomObject] @{
                                    Source      = $_.Source_Path 
                                    Destination = $DuplicateDir
                                    Results     = "ERROR, $errormessage"
                                }
                            }
                        }
                    }
                    #########################################################################################
                }
            }
            #########################################################################################
            else {
                ### if the file doesn't already exist then the script will either copy or move the file
                if ($Mode -eq "copy") {
                    ### if the mode selected was copy then the script will use the copy-item command to copy the file
                    try {
                        Write-Host "[$($_.Source_Path)] Copied to: $Destination_File"  -ForegroundColor DarkGreen
                        Copy-Item -Path $_.Source_Path -Destination $Destination_File
                        $Copied_Count++ ### Counter for copied files
                        ### Updating custom PS object to hold final results
                        $Final_Results += [PSCustomObject] @{
                            Source      = $_.Source_Path 
                            Destination = $Destination_File
                            Results     = "Copied"
                        }
                    }
                    Catch [Exception] {
                        ### if the command inside the try statement fails the error will be outputted
                        $errormessage = $_.Exception.Message
                        Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                        $Error_Count++ ### Counter for errors handled
                        ### Updating custom PS object to hold final results
                        $Final_Results += [PSCustomObject] @{
                            Source      = $_.Source_Path 
                            Destination = $Destination_File
                            Results     = "ERROR, $errormessage"
                        }
                    }
                }
                if ($Mode -eq "move") {
                    ### if the mode selected was move then the script will use the move-item command to move the file
                    try {
                        Write-Host "[$($_.Source_Path)] Moved to: $Destination_File" -ForegroundColor DarkGreen
                        Move-Item -Path $_.Source_Path -Destination $Destination_File -Force
                        $Moved_Count++ ### Counter for moved files
                        $Final_Results += [PSCustomObject] @{
                            Source      = $_.Source_Path 
                            Destination = $Destination_File
                            Results     = "Moved"
                        }
                    }
                    Catch [Exception] {
                        ### if the command inside the try statement fails the error will be outputted
                        $errormessage = $_.Exception.Message
                        Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                        $Error_Count++ ### Counter for errors handled
                        ### Updating custom PS object to hold final results
                        $Final_Results += [PSCustomObject] @{
                            Source      = $_.Source_Path 
                            Destination = $Destination_File
                            Results     = "ERROR, $errormessage"
                        }
                    }
                }
            }
            #########################################################################################
            ### Clearing variables to keep the script clean  
            Clear-Variable OriginalHash, NewFileHash, NewHashPath, Destination_File -ErrorAction SilentlyContinue
        }
        ### Sending the results to the terminal and transcript
        Write-Host "`nResults of the Media Sync Script" -ForegroundColor DarkYellow
        Write-host "Files copied: $Copied_Count" -ForegroundColor Cyan
        Write-host "Files moved: $Moved_Count" -ForegroundColor Cyan
        Write-host "Duplicate files: $Duplicate_Count" -ForegroundColor Cyan
        Write-host "Files skipped: $Skipped_Count" -ForegroundColor Cyan
        Write-host "Files copied (same name different hash): $Copied_Hash_Count" -ForegroundColor Cyan
        Write-host "Files moved (same name different hash): $Moved_Hash_Count" -ForegroundColor Cyan
        Write-host "Directories created: $Created_Dir_Count" -ForegroundColor Cyan
        Write-host "Files that are not Pictures or Videos: $Not_Photo_Video" -ForegroundColor Cyan
        Write-host "Errors: $Error_Count" -ForegroundColor Cyan

        $Final_Results | Out-GridView -Title "Final Results"
    }
    else { Write-Host "User Canceled Script" -ForegroundColor Red }
}
###############################################################################
### Function to remove files based on a user inputted file extension
###############################################################################
Function Remove_Extension {
    $Bad_Extensions = Read-Host "What file extension do you want to delete? (ie pdf)"
    Write-Host "Select the directory you want to remove files based on a given file extension in the folder dialog window" -ForegroundColor Yellow
    $Bad_Exts_Dir = Folder_Dialog -Message "Select the directory where the files reside"
    Write-Host "Please wait while the script gathers all the files to be deleted..." -ForegroundColor Yellow
    $Files_to_delete = Get-Childitem -path "$Bad_Exts_Dir\*" -Include @("*." + $Bad_Extensions) -Recurse  ### Getting all the files with given file extension
    $Final_files_to_delete = $Files_to_delete | Select-Object Name, FullName, LastWriteTime, Mode | Out-GridView -Title "Select the files that you want to delete." -OutputMode Multiple
    $Count_to_remove = $Final_files_to_delete | Measure-Object | ForEach-Object { $_.Count } ### Counting how many files are to be deleted

    ### Confirmation that the directories selected are correct
    Write-Host "`nConfirm that the following variables are correct" -ForegroundColor DarkYellow
    Write-host "Directory to remove unwanted files: $Bad_Exts_Dir" -ForegroundColor Cyan
    Write-host "Number of files to remove: $Count_to_remove"  -ForegroundColor Cyan
    Write-host "Files to be removed with extension: $Bad_Extensions" -ForegroundColor Cyan
    
    $Confirm = Read-Host "Do you want to continue? (Y/N)"
    #########################################################################################
    if ($Confirm.ToUpper() -eq "Y") {
        $Final_files_to_delete | ForEach-Object {
            try {
                Remove-Item $_.FullName -Force -Verbose -ErrorAction Stop
            }
            Catch [Exception] {
                ### if the command inside the try statement fails the error will be outputted
                $errormessage = $_.Exception.Message
                if ($errormessage -like "*does not exist*") {} ### Do not send an error if the item has already been deleted
                else {
                    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                }
            }
        }
    }
    else { Write-Host "User Canceled Script" -ForegroundColor Red }
}
###############################################################################
### Function to delete empty directories
###############################################################################
Function Delete_Empty_Dirs {
    Write-Host "Select the directory you want to remove the empty directories from" -ForegroundColor Yellow
    $Empty_Dir = Folder_Dialog -Message "Select the directory you want to remove the empty directories from"
    Write-Host "Please wait while the script gathers all the empty directories..." -ForegroundColor Yellow
    $Dirs_to_delete = Get-ChildItem $Empty_Dir -Directory -recurse -Force | Where-Object { -NOT $_.GetFiles("*", "AllDirectories") } ### Getting all the directories with no files
    $Final_Dirs_to_delete = $Dirs_to_delete | Select-Object Name, FullName, LastWriteTime, Mode | Out-GridView -Title "Select the directories that you want to delete." -OutputMode Multiple
    $Count_to_delete = $Final_Dirs_to_delete | Measure-Object | ForEach-Object { $_.Count } ### Counting how many folders are to be deleted

    ### Confirmation that the directories selected are correct
    Write-Host "`nConfirm that the following variables are correct" -ForegroundColor DarkYellow
    Write-host "Directory to remove unwanted files: $Empty_Dir" -ForegroundColor Cyan
    Write-host "Number of folders to remove: $Count_to_delete"  -ForegroundColor Cyan
        
    $Confirm = Read-Host "Do you want to continue? (Y/N)"
    #########################################################################################
    if ($Confirm.ToUpper() -eq "Y") {
        $Final_Dirs_to_delete | ForEach-Object {
            try {
                Remove-Item $_.FullName -Recurse -Force -Verbose -ErrorAction Stop
            }
            Catch [Exception] {
                ### if the command inside the try statement fails the error will be outputted
                $errormessage = $_.Exception.Message
                if ($errormessage -like "*does not exist*") {} ### Do not send an error if the item has already been deleted
                else {
                    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                }
            }
        }
        Write-host "Process complete!"  -ForegroundColor Cyan
    }
    else { Write-Host "User Canceled Script" -ForegroundColor Red }
}
###############################################################################
### Function to utilize OutGrid-View to select the files and folders to delete
###############################################################################
Function OutGrid_Delete {
    Write-Host "Select the directory you want to delete files and folders from" -ForegroundColor Yellow
    $Delete_Dir = Folder_Dialog -Message "Select the directory you want to delete files and folders from"
    Write-Host "Please wait while the script gathers all the files and folders in given directory..." -ForegroundColor Yellow
    $outgrid_delete = Get-ChildItem $Delete_Dir -Recurse | Select-Object Name, FullName, Length, LastWriteTime, Mode | Out-GridView -Title "Select the files and folders you want to delete from the $Delete_Dir directory" -OutputMode Multiple
    $Count_outgrid_delete = $outgrid_delete | Measure-Object | ForEach-Object { $_.Count } ### Counting how many items are to be deleted

    ### Confirmation that the directories selected are correct
    Write-Host "`nConfirm that the following variables are correct" -ForegroundColor DarkYellow
    Write-host "Directory to remove unwanted files: $Delete_Dir" -ForegroundColor Cyan
    Write-host "Number of objects to remove: $Count_outgrid_delete"  -ForegroundColor Cyan

    $Confirm = Read-Host "Do you want to continue? (Y/N)"
    #########################################################################################
    if ($Confirm.ToUpper() -eq "Y") {
        foreach ($Object in $outgrid_delete) {
            try {
                Remove-Item $Object.FullName -Recurse -Force -Verbose -ErrorAction Stop
            }
            Catch [Exception] {
                ### if the command inside the try statement fails the error will be outputted
                $errormessage = $_.Exception.Message
                if ($errormessage -like "*does not exist*") {} ### Do not send an error if the item has already been deleted
                else {
                    Write-host $errormessage -BackgroundColor Red -ForegroundColor White
                }
            }
        }
        Write-host "Process complete!"  -ForegroundColor Cyan
    }
    else { Write-Host "User Canceled Script" -ForegroundColor Red }
}
###############################################################################
### Function to find duplicate files
###############################################################################
Function Duplicate_File_Finder {
    Write-Host "Select the directory where you want to find duplicate files" -ForegroundColor Yellow
    $Duplicate_Dir = Folder_Dialog -Message "Select the directory where you want to find duplicate files"
    Write-Host "`n ---- Please wait while the script gathers all the files that have identical sizes ---- `n" -ForegroundColor Cyan
    $Files_Same_Size = Get-ChildItem -File -Path $Duplicate_Dir -Recurse | Group-Object Length | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty group
    Write-Host $Files_Same_Size.Length "Potential duplicate files found" -ForegroundColor Yellow
    Write-Host "`n ---- Please wait while the script gathers all the hash values ---- `n" -ForegroundColor Cyan
    $i = 0 # setting variable to 0 for the start of the progress bar
    $File_Hashes = $Files_Same_Size  | ForEach-Object {
        $i++ ### incrementing the progress bar by 1 each loop through
        Write-Progress -Activity "Gathering hash values" -Status ("Checking : {0}" -f $_.FullName) -PercentComplete ($i / $Files_Same_Size.count * 100) -Id 0 ### Progress bar
        Get-FileHash -Algorithm SHA1 $_.fullname ### Getting the SHA1 file hash
    } 
    
    $Groups_of_dupes = $File_Hashes | Group-Object -Property Hash | Where-Object { $_.count -gt 1 } ### Finding any file hashes that exist more than once, which would indicate a duplicate file
    Write-Host $Groups_of_dupes.Length "Groups of duplicate files found" -ForegroundColor Yellow
    #########################################################################################
    Write-Host "`n ----------------- Media Sync Menu -----------------" -ForegroundColor Cyan
    Write-Host "1: Enter 1 to view all of the duplicate files detected"
    Write-Host "2: Enter 2 to move files for ONE duplicate group at a time"
    Write-Host "3: Enter 3 to move files for ALL duplicate groups"
    Write-Host "4: Enter 4 to auto move files"
    Write-Host "Q: Enter Q to quit. `n"
    #########################################################################################
    $selection = (Read-Host "Please make a selection").ToUpper()
    
    switch ($selection) {
        #########################################################################################
        '1' {
            ### View only
            $Groups_of_dupes | Select-Object -ExpandProperty Group | Out-GridView -Title "List of all the duplicate files detected"
            Write-host "Process complete!"  -ForegroundColor Cyan
        }
        #########################################################################################
        '2' {
            ### Move one group at a time
            $Selected_Files_to_Move = $Groups_of_dupes | ForEach-Object { $_.group | Out-Gridview -Title "Select the files you want to move to the duplicate files directory" -OutputMode Multiple }
            Write-Host $Selected_Files_to_Move.Length "Duplicate files to be moved" -ForegroundColor Cyan
            Write-Host "In the folder dialog window select the directory you want to move duplicates to" -ForegroundColor Yellow
            $DuplicateDir = Folder_Dialog -Message "Select the duplicate files directory"
            
            foreach ($Object in $Selected_Files_to_Move) {
                Move-Item -Path $Object.Path -Destination $DuplicateDir -Force -Verbose ### Moving selected file to duplicate files directory
            }
            Write-host "Process complete!"  -ForegroundColor Cyan
        }
        #########################################################################################
        '3' {
            ### Move all groups at once
            $Selected_Files_to_Move = $Groups_of_dupes | Select-Object -ExpandProperty Group | Out-GridView -Title "Select the files you want to move to the duplicate files directory" -OutputMode Multiple
            Write-Host $Selected_Files_to_Move.Length "Duplicate files to be moved" -ForegroundColor Cyan
            Write-Host "In the folder dialog window select the directory you want to move duplicates to" -ForegroundColor Yellow
            $DuplicateDir = Folder_Dialog -Message "Select the duplicate files directory"
        
            foreach ($Object in $Selected_Files_to_Move) {
                Move-Item -Path $Object.Path -Destination $DuplicateDir -Force -Verbose ### Moving selected file to duplicate files directory
            }
            Write-host "Process complete!"  -ForegroundColor Cyan
        }
        #########################################################################################
        '4' {
            ### Auto move duplicate files to the duplicates directory
            $Selected_Files_to_Move = $Groups_of_dupes | ForEach-Object {
                $Files_in_group = $_.group | Measure-Object | ForEach-Object { $_.Count } ### Counting how many files are in the group
                $number_to_delete = $Files_in_group - 1 ### Subtracting one from the total in the group, by doing so will keep one file from the duplicates group
                $_.group | Select-Object -Last $number_to_delete  ### Selecting the objects to move to duplicates directory
            }
            Write-Host $Selected_Files_to_Move.Length "Duplicate files to be moved" -ForegroundColor Cyan
            Write-Host "In the folder dialog window select the directory you want to move duplicates to" -ForegroundColor Yellow
            $DuplicateDir = Folder_Dialog -Message "Select the duplicate files directory"
            
            foreach ($Object in $Selected_Files_to_Move) {
                Move-Item -Path $Object.Path -Destination $DuplicateDir -Force -Verbose ### Moving selected file to duplicate files directory
            }
            Write-host "Process complete!"  -ForegroundColor Cyan
        }
        #########################################################################################
        'Q' { Write-Host "The script has been canceled" -BackgroundColor Red -ForegroundColor White }
        Default { Write-Host "Your selection = $selection, is not valid. Please try again." -BackgroundColor Red -ForegroundColor White }
        #########################################################################################
    }
}
###############################################################################
### Function to view files in a given directory
###############################################################################
Function View_Files {
    Write-Host "Select the directory you want to view files for" -ForegroundColor Yellow
    $View_Dir = Folder_Dialog -Message "Select the directory you want to view files for"
    Get-ChildItem $View_Dir -recurse -Force | Select-Object Name, FullName, Length, LastWriteTime, Mode | Out-GridView -Title "Files in the $View_Dir directory"
    Write-host "Process complete!"  -ForegroundColor Cyan
}
###############################################################################
### Function to conduct cleanup
###############################################################################
Function Cleanup {

    Write-Host "`n ----------------- Media Sync Menu -----------------" -ForegroundColor Cyan
    Write-Host "1: Enter 1 to remove files with given extension (manual input)"
    Write-Host "2: Enter 2 to delete all empty directories"
    Write-Host "3: Enter 3 to utilize Out-GridView to highlight and delete files or folders"
    Write-Host "4: Enter 4 to find duplicate files in a given directory"
    Write-Host "5: Enter 5 to view files in a directory"
    Write-Host "Q: Enter Q to quit. `n"
    
    $selection = (Read-Host "Please make a selection").ToUpper()
    
    switch ($selection) {
        '1' { Remove_Extension }  
        '2' { Delete_Empty_Dirs }
        '3' { OutGrid_Delete }
        '4' { Duplicate_File_Finder }
        '5' { View_Files }
        'Q' { Write-Host "The script has been canceled" -BackgroundColor Red -ForegroundColor White }
        Default { Write-Host "Your selection = $selection, is not valid. Please try again." -BackgroundColor Red -ForegroundColor White }
    }
}
###############################################################################
### Running the functions / menu
###############################################################################
Write-Host "`n ----------------- Media Sync Menu -----------------" -ForegroundColor Cyan
Write-Host "1: Enter 1 to COPY files and folders"
Write-Host "2: Enter 2 to MOVE files and folders"
Write-Host "3: Enter 3 to CLEANUP files and folders"
Write-Host "Q: Enter Q to quit. `n"

$selection = (Read-Host "Please make a selection").ToUpper()

switch ($selection) {
    '1' { Photos_Videos -Mode copy }    ### Input the name of the function you want to execute when 1 is entered
    '2' { Photos_Videos -Mode move }    ### Input the name of the function you want to execute when 2 is entered
    '3' { Cleanup }  ### Input the name of the function you want to execute when 3 is entered
    'Q' { Write-Host "The script has been canceled" -BackgroundColor Red -ForegroundColor White }
    Default { Write-Host "Your selection = $selection, is not valid. Please try again." -BackgroundColor Red -ForegroundColor White }
}

Stop-Transcript
