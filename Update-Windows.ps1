<#
.SYNOPSIS
Script for updating drivers with HP Image Assistant (HPIA).

.DESCRIPTION
This script downloads and runs the HP Image Assistant utility to update drivers and BIOS on HP computers. It requires administrative privileges to execute.

.NOTES
    Author: Jessica Rojas Mosquera
    Date: 03/06/2024
    Version: 1.0

.PARAMETER OnlyBIOS
If specified, the script will only update the BIOS.

.EXAMPLE
.\UpdateDriversWithHPIA.ps1
.\UpdateDriversWithHPIA.ps1 -OnlyBIOS
#>

param(
    [switch]$OnlyBIOS
)

function Test-Admin {
    <#
    .SYNOPSIS
    Checks if the current user has administrative privileges.

    .OUTPUTS
    [bool]
    Returns $true if the user is an administrator, otherwise $false.
    #>
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Initialize-Paths {
    <#
    .SYNOPSIS
    Initializes necessary paths and URLs.

    .OUTPUTS
    [hashtable]
    Returns a hashtable with initialized paths and URLs.
    #>
    $paths = @{
        BasePath         = "C:\ProgramData\HP\HP Image Assistant"
        FinalReportPath  = "C:\ProgramData\HP\HP Image Assistant\reports"
        DownloadPath     = "C:\ProgramData\HP\HP Image Assistant\HPIADownload"
        InstallerUrl     = "https://ftp.ext.hp.com/pub/caps-softpaq/cmit/HPIA.html"
        Installer        = "C:\ProgramData\HP\HP Image Assistant\HPIA.exe"
        LogFilePath      = "C:\ProgramData\HP\HP Image Assistant\UpdateLog.txt"
    }
    return $paths
}

function Clean-Directories {
    <#
    .SYNOPSIS
    Cleans existing files in specified directories.

    .PARAMETER BasePath
    The base directory path to clean.

    .PARAMETER FinalReportPath
    The report directory path to ensure exists.
    #>
    param (
        [string]$BasePath,
        [string]$FinalReportPath
    )

    if (Test-Path -Path $BasePath) {
        Write-Host "Cleaning up existing files in $BasePath"
        Get-ChildItem -Path $BasePath -Exclude "UpdateLog.txt" | Remove-Item -Recurse -Force
    } else {
        New-Item -Path $BasePath -ItemType Directory -Force | Out-Null
    }

    if (-Not (Test-Path -Path $FinalReportPath)) {
        New-Item -Path $FinalReportPath -ItemType Directory -Force | Out-Null
    }
}

function Download-HPIAInstaller {
    <#
    .SYNOPSIS
    Downloads the HP Image Assistant installer.

    .PARAMETER InstallerUrl
    The URL to download the installer from.

    .PARAMETER InstallerPath
    The local path to save the installer to.
    #>
    param (
        [string]$InstallerUrl,
        [string]$InstallerPath
    )

    try {
        Write-Host "Downloading HP Image Assistant (HPIA) utility"
        $HPInstallerPage = Invoke-WebRequest -Uri $InstallerUrl -UseBasicParsing
        $HPIAInstallerURL = ($HPInstallerPage.links | Where-Object { $_.href -like "*.exe" }).href
        Invoke-WebRequest -Uri $HPIAInstallerURL -OutFile $InstallerPath
    } catch {
        Write-Error "Failed to download the HPIA installer: $_"
        throw
    }
}

function Extract-HPIAInstaller {
    <#
    .SYNOPSIS
    Extracts the HP Image Assistant installer.

    .PARAMETER InstallerPath
    The local path of the downloaded installer.

    .PARAMETER BasePath
    The base directory to extract the installer to.
    #>
    param (
        [string]$InstallerPath,
        [string]$BasePath
    )

    try {
        Write-Host "Extracting HPIA installer"
        Start-Process -FilePath $InstallerPath -ArgumentList "/s /e /f "$BasePath"" -Wait
    } catch {
        Write-Error "Failed to extract the HPIA installer: $_"
        throw
    }
}

function Run-HPIAUtility {
    <#
    .SYNOPSIS
    Runs the HP Image Assistant utility.

    .PARAMETER BasePath
    The base directory where HPIA is located.

    .PARAMETER DownloadPath
    The directory to download the updates to.

    .PARAMETER OnlyBIOS
    If specified, only BIOS updates will be performed.
    #>
    param (
        [string]$BasePath,
        [string]$DownloadPath,
        [switch]$OnlyBIOS
    )

    try {
        Write-Host "Running HPIA utility"
        if ($OnlyBIOS) {
            Start-Process -FilePath "$BasePath\HPImageAssistant.exe" -ArgumentList "/Operation:Analyze /Category:BIOS /selection:All /action:Extract /Noninteractive /softpaqdownloadfolder:$DownloadPath" -Wait
        } else {
            Start-Process -FilePath "$BasePath\HPImageAssistant.exe" -ArgumentList "/Operation:Analyze /Category:All /selection:All /action:Extract /Noninteractive /softpaqdownloadfolder:$DownloadPath" -Wait
        }
    } catch {
        Write-Error "Failed to run the HPIA utility: $_"
        throw
    }
}

function Install-Updates {
    <#
    .SYNOPSIS
    Installs downloaded updates.

    .PARAMETER DownloadPath
    The path where updates are downloaded.

    .PARAMETER LogFilePath
    The path to the log file.
    #>
    param (
        [string]$DownloadPath,
        [string]$LogFilePath
    )

    try {
        if (Test-Path -Path $DownloadPath) {
            Write-Host "HPIA successfully downloaded all packages."

            Write-Host "Installing updates."
            $installScript = "$DownloadPath\InstallAll.cmd"
            if (Test-Path -Path $installScript) {
                Get-Content $installScript | ForEach-Object {
                    if ($_ -match "REM\s+Package\s+Name:\s+(.*)") {
                        Write-Host "Installing driver $($matches[1])"
                    }
                }
                Start-Process -FilePath $installScript -Wait
            } else {
                Write-Warning "InstallAll.cmd not found, skipping installation."
            }

            $logEntry = "Update completed successfully on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            Add-Content -Path $LogFilePath -Value $logEntry

            Write-Warning "HP updates installation completed. It's recommended to restart the computer before continuing."
        } else {
            Write-Host "HPIA did not find any updates."
            $logEntry = "No updates found on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            Add-Content -Path $LogFilePath -Value $logEntry
        }
    } catch {
        Write-Error "An error occurred while processing the updates: $_"
        $logEntry = "Error occurred on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $_"
        Add-Content -Path $LogFilePath -Value $logEntry
        throw
    }
}

function Clean-Up {
    <#
    .SYNOPSIS
    Cleans up temporary files and directories.

    .PARAMETER BasePath
    The base directory to clean.

    .PARAMETER FinalReportPath
    The path to the final report directory.
    #>
    param (
        [string]$BasePath,
        [string]$FinalReportPath
    )

    try {
        Write-Host "Cleaning up temporary files."
        Get-ChildItem -Path $BasePath -Exclude "UpdateLog.txt", "reports" | Remove-Item -Recurse -Force

        if (Test-Path C:\SWSetup) { Remove-Item -Path C:\SWSetup -Recurse -Force }

        Get-ChildItem c:\ | Where-Object { $_.Attributes -match "ReparsePoint" -and $_.Name -match '^sp\d+' } | Remove-Item -Force
    } catch {
        Write-Error "Failed to clean up temporary files: $_"
        throw
    }
}

function Move-Reports {
    <#
    .SYNOPSIS
    Moves report files from Downloads to the Reports folder.

    .PARAMETER FinalReportPath
    The path to the final report directory.
    #>
    param (
        [string]$FinalReportPath
    )

    try {
        $ComputerInfo = Get-WmiObject -Class Win32_ComputerSystem
        $computerModel = $ComputerInfo.Model
        Write-Host "Computer Model: $computerModel"

        $sourceFolder = "$env:USERPROFILE\Downloads"
        $fileNamePattern = "$computerModel.*"

        Get-ChildItem -Path $sourceFolder -Filter "$fileNamePattern" | ForEach-Object {
            if ($_.Name -match "\.xml$|\.html$|\.json$") {
                Write-Host "Moving file: $($_.FullName) to $FinalReportPath"
                Move-Item -Path $_.FullName -Destination $FinalReportPath -Force
            } else {
                Write-Host "Skipping file: $($_.FullName) (does not match the required extensions)"
            }
        }
    } catch {
        Write-Error "Failed to move report files: $_"
        throw
    }
}

# Main script execution
if (-not (Test-Admin)) {
    Write-Warning "This script must be run as an Administrator."
    return
}

Write-Host "Script is running with administrative privileges."

$paths = Initialize-Paths
Clean-Directories -BasePath $paths.BasePath -FinalReportPath $paths.FinalReportPath

Download-HPIAInstaller -InstallerUrl $paths.InstallerUrl -InstallerPath $paths.Installer
Extract-HPIAInstaller -InstallerPath $paths.Installer -BasePath $paths.BasePath

Run-HPIAUtility -BasePath $paths.BasePath -DownloadPath $paths.DownloadPath -OnlyBIOS:$OnlyBIOS
Install-Updates -DownloadPath $paths.DownloadPath -LogFilePath $paths.LogFilePath

Clean-Up -BasePath $paths.BasePath -FinalReportPath $paths.FinalReportPath
Move-Reports -FinalReportPath $paths.FinalReportPath
