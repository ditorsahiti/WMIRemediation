<#
.SYNOPSIS
A PowerShell script designed to detect and remedy potential WMI class or namespace issues.

.DESCRIPTION
This script checks for specific errors related to the WMI class. If such errors are identified, various remediation steps are undertaken, 
including the re-registration of WBEM DLLs and EXEs, the restoration of the WMI repository, and the restarting of associated services.

.PARAMETER None required

.EXAMPLE
PS> .\WMIRemediation.ps1

.NOTES
Date Created: 2023-08-18
Author: Ditor Sahiti
#>

# Define the log file path
$logFilePath = "C:\Windows\Logs\WMIRemediation.log"

# Function to write logs
function Write-Log {
    param (
        [string]$Message
    )
    Add-Content -Path $logFilePath -Value "$(Get-Date) - $Message"
}

# Attempt to retrieve Network Adapters
$adapters = $null
$errorOccurred = $false

# Redirect error stream to variable
$adapters = Get-NetAdapter 2>&1

# Check if the error message contains the substring 'Invalid class'
if ($adapters -like "*Invalid*") {
    $errorOccurred = $true
    Write-Log "Error occurred: Invalid class in network adapters."
}

# If WMI class related error occurred
if ($errorOccurred) {
    Write-Log "WMI class appears to be missing. Attempting to fix..."

    # Attempt to stop ConfigManager service and WMI service
    Write-Log 'Stopping ConfigManager service if it exists'
    Stop-Service -Force 'ccmexec' -ErrorAction 'SilentlyContinue'

    # Temporarily disable the WMI service to perform corrective actions
    Write-Log 'Temporarily disabling Windows Management Instrumentation service'
    Set-Service -Name 'winmgmt' -StartupType 'Disabled' 

    Write-Log 'Stopping Windows Management Instrumentation service'
    Stop-Service -Force 'winmgmt' -ErrorAction 'SilentlyContinue'

    # Run the commands to re-register WMI components

    Write-Log 'Running commands to re-register WMI components'
    Set-Location "$env:windir\system32\wbem"
    cmd /c "for /f %s in ('dir /b *.dll') do regsvr32 /s %s"
    cmd /c "wmiprvse /regserver"
    cmd /c "winmgmt /regserver"
    cmd /c "for /f %s in ('dir /s /b *.mof *.mfl') do mofcomp %s"

    # Restore the WMI service to its original state and restart it
    Write-Log 'Setting Windows Management Instrumentation service back to Automatic'
    Set-Service -Name 'winmgmt' -StartupType 'Automatic'

    Write-Log 'Starting Windows Management Instrumentation service'
    Start-Service winmgmt -ErrorAction SilentlyContinue

    # Compile the provided MOF file
    Write-Log "Executed mofcomp on 'ExtendedStatus.mof'"
    mofcomp "C:\Program Files\Microsoft Policy Platform\ExtendedStatus.mof"
    
    # Run ccmeval.exe without waiting for it to finish
    Write-Log "Running ConfigMgr Evaluation"
    Start-Job -ScriptBlock {
    & "C:\Windows\ccm\ccmeval.exe"
    }
    Write-Log 'Remediation process completed'
    
    # Retry the retrieval of network adapters
    $adapters = Get-NetAdapter 2>&1
    if ($adapters -like "*Invalid*") {
        Write-Log "Error still present: Invalid class in network adapters. Running additional fix."

        Set-Location "$env:windir\system32\wbem"
        
        # Run the additional remediation steps
        cmd /c "regsvr32.exe /s /i .\NetAdapterCim.dll"
        cmd /c ".\mofcomp.exe .\NetAdapterCim.mof"
        Write-Log 'Remediation process completed'
    }
}
else {
    Write-Log "No issues detected with WMI class for network adapters."
}