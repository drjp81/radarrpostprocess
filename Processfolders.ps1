<#.SYNOPSIS
    This script extracts .rar files from a specified input path and moves the extracted files to a staging path.

.DESCRIPTION
    The script takes an input path as a parameter and searches for .rar files within the specified path and its subdirectories.
    It then extracts the contents of each .rar file using the unrar command and moves the extracted files to a staging path.
    Finally, it deletes the original .rar files if the extraction and move back to the original folder were successful.

.PARAMETER inputpath
    Specifies the path where the .rar files are located.

.EXAMPLE
    .\Processfolders.ps1 -inputpath "/media/plex/Movies/Completed/"
    Extracts .rar files from the "/media/plex/Movies/Completed" directory and its subdirectories, and moves the extracted files to the staging path.

.NOTES
    - This script requires the unrar command-line tool to be installed.
    - The staging path is created if it does not exist.
    - The staging path is set to a subdirectory named "tmp" within the base path of the input path.
    - The script logs the extraction process to a log file named "folderlogfile.log" located in the same directory as the script.
#>
param (
    [string]$inputpath
)

# Read and trim the log file if it exceeds the maximum line count
$logfile = Join-Path -Path $PSScriptRoot -ChildPath "folderlogfile.log"
$maxLines = 200
$logContent = Get-Content -Path $logfile
if ($logContent.Count -gt $maxLines) {
    # Keep only the last 200 lines
    $trimmedContent = $logContent | Select-Object -Last $maxLines
    Set-Content -Path $logfile -Value $trimmedContent
}

<#
.SYNOPSIS
This script defines a function to log messages to a log file and display them on the console.

.DESCRIPTION
The log function in this script allows you to log messages to a log file and also display them on the console. It appends the current date and time to the log message before writing it to the log file.

.PARAMETER message
The message parameter represents the log message that you want to write to the log file.

.EXAMPLE
log "Processing started"
This example logs the message "Processing started" to the log file and displays it on the console.

#>

function log([string]$message) {
    $logfile = Join-Path -Path $PSScriptRoot -ChildPath "folderlogfile.log"
    # Add new log message
    $NEWMSG = (Get-Date -Format "yyyy-MM-dd HH:mm:ss").ToString() + " - " + $message
    Add-Content -Path $logfile -Value $NEWMSG
    # Write log message to the console
    Write-Host $NEWMSG
}

<#
.SYNOPSIS
    Cleans up a movie folder by removing unnecessary files.

.DESCRIPTION
    The CleanMovieFolder function removes files from a specified folder that are not media files or subtitle files. 
    It also excludes files named 'folder.jpg' or 'folder.png'.

.PARAMETER folderPath
    Specifies the path of the folder to be cleaned.

.NOTES
    - This function requires PowerShell 3.0 or later.
    - This function does not delete folders or subfolders, only files within the specified folder.
    - Media file extensions include: .mkv, .mp4, .avi, .mov, .wmv, .flv, .mpeg, .mpg, .m4v, .divx, .vob, .iso, .ts, .m2ts, .mts, .webm, .ogm, .rmvb, .rm, .3gp, .asf, .wm, .wma, .ogv, .m2t, .m2v, .mpv, .mpeg1, .mpeg2, .mpeg4, .mpg2, .mpg4, .h264, .h265.
    - Subtitle file extensions include: .srt, .sub, .idx, .ssa, .ass.

.EXAMPLE
    CleanMovieFolder -folderPath "C:\Movies"

    This example cleans up the "C:\Movies" folder by removing unnecessary files.

#>
function CleanMovieFolder([string]$folderPath) {
    # Define media file extensions (including legacy types)
    $mediaExtensions = @('.mkv', '.mp4', '.avi', '.mov', '.wmv', '.flv', '.mpeg', '.mpg', '.m4v', '.divx', '.vob', '.iso', '.ts', '.m2ts', '.mts', '.webm', '.ogm', '.rmvb', '.rm', '.3gp', '.asf', '.wm', '.wma', '.ogv', '.m2t', '.m2v', '.mpv', '.mpeg1', '.mpeg2', '.mpeg4', '.mpg2', '.mpg4', '.h264', '.h265')

    # Define subtitle file extensions
    $subtitleExtensions = @('.srt', '.sub', '.idx', '.ssa', '.ass')

    # Get all files in the folder
    $files = Get-ChildItem -Path $folderPath -File

    # Loop through each file
    foreach ($file in $files) {
        # Check if the file's extension is neither a media nor a subtitle extension
        # and the file is not named 'folder.jpg' or 'folder.png'
        if ($mediaExtensions -notcontains $file.Extension -and 
            $subtitleExtensions -notcontains $file.Extension -and
            $file.Name -ne "folder.jpg" -and 
            $file.Name -ne "folder.png") {
            # Delete the file
            try {
                log("Going to delete the file: $($file.FullName)")
                Remove-Item $file.FullName -Force
                log("Deleted file: $($file.FullName)")
            }
            catch {
                log("Error in deleting old file: $($file.FullName)")
                log("The error was: $($_.Exception.Message)")
            }
            
        }
    }
}

<#
.SYNOPSIS
Deletes original archive files in a specified folder.

.DESCRIPTION
The DeeleteRarArchives function deletes original archive files in the specified folder. It searches for files with the extension ".rar" or ".rxx" and removes them from the folder.

.PARAMETER folderPath
The path of the folder where the archive files are located.

.EXAMPLE
DeeleteRarArchives -folderPath "C:\Movies"

This example deletes all the original archive files in the "C:\Movies" folder.

#>
function DeeleteRarArchives([string]$folderPath) {
    
    $oldarchives = Get-ChildItem -LiteralPath $folderPath -File -Filter "*.r??" 
    try {
        log("Deleting original archive files in: $($folderPath)")
        log("The files to be deleted are: $($oldarchives.FullName)")
        $oldarchives | Remove-Item -Force
    }
    catch {
        log("Error in deleting old archive files")
        log("The error was: $($_.Exception.Message)")
    }
}
    <#
    .SYNOPSIS
    Extracts RAR files from a specified path and moves the extracted files to a destination folder.

    .DESCRIPTION
    The Extract_RarFiles function searches for RAR files in the specified path and extracts their contents to a temporary directory. 
    It then moves the extracted files to a destination folder. If the extraction and move operations are successful, the original RAR files are deleted.

    .PARAMETER path
    The path where the RAR files are located.

    .PARAMETER destination
    The destination folder where the extracted files will be moved.

    .EXAMPLE
    Extract_RarFiles -path "C:\RARFiles" -destination "C:\ExtractedFiles"

    This example extracts RAR files from the "C:\RARFiles" directory and moves the extracted files to the "C:\ExtractedFiles" directory.

    .NOTES
    - This function requires the unrar command-line tool to be installed and accessible via the "/usr/bin/unrar" path.
    - The function uses logging functions (not shown in this code) to log the extraction process.
    - The function also calls the DeeleteRarArchives and CleanMovieFolder functions to perform additional cleanup operations.

    #>
    function Extract_RarFiles([string]$path, [string]$destination) {
        $rarFiles = Get-ChildItem -Path $path -Recurse -Filter "*.rar"
        $tempdir = $script:stagingPath
        foreach ($rarFile in $rarFiles) {
            # Construct the command to extract the archive
            $archivedir = $rarFile.DirectoryName
            Set-Location -LiteralPath $archivedir
            $fullname = $rarFile.FullName
            log("Going to unrar: $fullname to $tempdir")
            $Arguments = @('e', '-y', '-ierr', '-ai', '-o+', $fullname, $tempdir)
            $cmd = "/usr/bin/unrar"
            log("Executing: $cmd $Arguments")
            try {
                Start-Transcript -Path $logfile -Append -Force -NoClobber
                $res = Start-Process -FilePath $cmd -ArgumentList $Arguments -Wait -NoNewWindow -ErrorAction Stop -PassThru
                $status = $res.ExitCode
                if ($status -ne 0) {
                    Throw "Error processing archive, error number: $status"
                }
                log("Finished processing file $fullname") 
                $extractedFiles = Get-ChildItem -Path $destination -Recurse
                log("Moving extracted files to $archivedir")
                $extractedFiles | Move-Item -Destination $archivedir -Force
                log("Successfully extracted and moved files from $($rarFile.Name)")
                # Delete original archive files if extraction and move was successful
                DeeleteRarArchives -folderPath $archivedir
                log("Cleaning up folder: $archivedir")
                CleanMovieFolder -folderPath $archivedir
            }
            catch {
                log("Error processing archive: $fullname")
                log($_.Exception.Message)
                log($status)
                Stop-Transcript
                return 1
            }
            finally {
                Stop-Transcript -ErrorAction SilentlyContinue
            }
        }
        return 0
    }

    <#
    .SYNOPSIS
        Initializes the staging path for processing folders.

    .DESCRIPTION
        The Initialize-StagingPath function is used to set up the staging path for processing folders. 
        It checks if the original path contains the word "completed". If it does, it removes the "Completed" 
        portion from the path and appends "tmp" to create the staging path. If the original path does not 
        contain "completed", it replaces the leaf (last folder or file name) with "tmp" to create the staging path.

    .PARAMETER originalpath
        The original path for which the staging path needs to be initialized.

    .EXAMPLE
        Initialize-StagingPath -originalpath "C:\Movies\Completed\Movie1"
        Initializes the staging path for the given original path to "C:\Movies\tmp".
        
        Initialize-StagingPath -originalpath "C:\Movies\Whatever"
        Initializes the staging path for the given original path to "C:\Movies\tmp".

    #>
    function Initialize-StagingPath([string]$originalpath) {

        if ($originalpath -ilike "*completed*") {
            $basePath = $originalpath -replace 'Completed.*', ''
            $script:stagingPath = Join-Path -Path $basePath -ChildPath "tmp"
        }
        else {
            $leaf = Split-Path -Path $originalpath -Leaf
            $script:stagingPath = $originalpath -replace $leaf, "tmp"
        }
        # Check if the staging directory exists, if not, create it
        log("Checking if staging path exists: $script:stagingPath")
        if (-not (Test-Path -Path $script:stagingPath)) {
            log("Creating staging path: $script:stagingPath")
            New-Item -ItemType Directory -Path $script:stagingPath -Force
            chmod 777 -R $script:stagingPath
        }
        log("Staging path is set to: $script:stagingPath")
    }

    # Call the function with the $inputpath
    log("Initializing staging path")
    Initialize-StagingPath -originalpath $inputpath

    # Call the function with the provided search path and staging path
    log("Extracting rar files from $inputpath")
    exit (Extract_RarFiles -path $inputpath -destination $script:stagingPath)
