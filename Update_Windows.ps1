<#
.SYNOPSIS
Script for checking and installing Windows updates using PSWindowsUpdate module.

.DESCRIPTION
This script checks for available Windows updates using the PSWindowsUpdate module. If updates are found,
it installs them and checks if a system restart is required.

#>

# Ensure the PSWindowsUpdate module is imported
Import-Module PSWindowsUpdate

try {
    # Check for updates
    Write-Host "Checking for updates..."
    $updates = Get-WindowsUpdate -AcceptAll -IgnoreReboot

    if ($updates) {
        Write-Host "Updates found. Installing updates..."
        Install-WindowsUpdate -AcceptAll -IgnoreReboot -Verbose

        # Check for installed updates that require a restart
        $pendingReboot = (Get-WindowsUpdate -Install -AcceptAll -IgnoreReboot | Where-Object { $_.RebootRequired })

        if ($pendingReboot) {
            Write-Host "Updates installed. A restart is required to complete the installation."
            # Prompt for restart
            $restart = Read-Host "Do you want to restart now? (y/n)"
            if ($restart -eq 'y') {
                Restart-Computer -Force
            } else {
                Write-Host "Please restart your computer manually to complete the update process."
            }
        } else {
            Write-Host "Updates installed. No restart required."
        }
    } else {
        Write-Host "No updates available."
    }

} catch {
    Write-Error "An error occurred: $_"
}

# End of script
Write-Host "Script execution completed."
