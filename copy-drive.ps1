[cmdletbinding()]
param (
    [Parameter (Mandatory = $true)] [String]$ConfigFilePath
)


#if json config file does not exist, abort process
if (-not(Test-Path -Path $ConfigFilePath -PathType Leaf)) {
    throw "json config file specified at $($ConfigFilePath) does not exist, aborting process"
}
  
#if config file configured is not json format, abort process.
try {
    $PowerShellObject=Get-Content -Path $ConfigFilePath | ConvertFrom-Json
} catch {
    throw "Config file of $($ConfigFilePath) is not a valid json file, aborting process"
}

#if SourceDiskDriveLetter option does not exist in json, abort process
if ($PowerShellObject.Required.SourceDiskDriveLetter) {
    $SourceDiskDriveLetter = $PowerShellObject.Required.SourceDiskDriveLetter
} else {
    throw "SourceDiskDriveLetter does not exist in json config file, aborting process"
}

#if SourceDiskID option does not exist in json, abort process
if ($PowerShellObject.Required.SourceDiskID) {
    $SourceDiskID = $PowerShellObject.Required.SourceDiskID
} else {
    throw "SourceDiskID does not exist in json config file, aborting process"
}

#if DestinationDiskDriveLetter option does not exist in json, abort process
if ($PowerShellObject.Required.DestinationDiskDriveLetter) {
    $DestinationDiskDriveLetter = $PowerShellObject.Required.DestinationDiskDriveLetter
} else {
    throw "DestinationDiskDriveLetter does not exist in json config file, aborting process"
}

#if devicesToCheck optoin does not exist in json, abort process
if ($PowerShellObject.Required.DestinationDiskID) {
    $DestinationDiskID = $PowerShellObject.Required.DestinationDiskID
} else {
    throw "DestinationDiskID does not exist in json config file, aborting process"
}

#if errorMailSender optoin does not exist in json, abort process
if ($PowerShellObject.Required.errorMailSender) {
    $errorMailSender = $PowerShellObject.Required.errorMailSender
} else {
    throw "errorMailSender does not exist in json config file, aborting process"
}

#if errorMailRecipients option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailRecipients) {
    $errorMailRecipients = $PowerShellObject.Required.errorMailRecipients
} else {
    throw "errorMailRecipients does not exist in json config file, aborting process"
}

#if errorMailTenantID option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailTenantID) {
    $errorMailTenantID = $PowerShellObject.Required.errorMailTenantID
} else {
    throw "errorMailTenantID does not exist in json config file, aborting process"
}

#if errorMailAppID option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailAppID) {
    $errorMailAppID = $PowerShellObject.Required.errorMailAppID
} else {
    throw "errorMailAppID does not exist in json config file, aborting process"
}

#if errorMailSubjectPrefix option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailSubjectPrefix) {
    $errorMailSubjectPrefix = $PowerShellObject.Required.errorMailSubjectPrefix
} else {
    throw "errorMailSubjectPrefix does not exist in json config file, aborting process"
}

#if errorMailPasswordFile option does not exist in json, abort process
if ($PowerShellObject.Required.errorMailPasswordFile) {
    $errorMailPasswordFile = $PowerShellObject.Required.errorMailPasswordFile
} else {
    throw "errorMailPasswordFile does not exist in json config file, aborting process"
}

#set up variables
[string] $strExecDir = $PSScriptRoot
[string] $strServerName = $env:computername
[bool] $blnWriteToLog = $false
[string] $strRoboCopySourceDrive = "$($SourceDiskDriveLetter):\"
[string] $strRoboCopyDestinationDrive = "$($DestinationDiskDriveLetter):\"
[int] $intErrorCount = 0
$arrStrErrors = @()

#clear all errors before starting
$error.Clear()

[uint16] $intDaysToKeepLogFiles = 0
[string] $strServerName = $env:computername

#if path to log directory exists, set logging to true and setup log file
if (Test-Path -Path $PowerShellObject.Optional.logsDirectory -PathType Container) {
    $blnWriteToLog = $true
    [string] $strTimeStamp = $(get-date -f yyyy-MM-dd-hh_mm_ss)
    [string] $strDetailLogFilePath = $PowerShellObject.Optional.logsDirectory + "\win-drive-bckup-status-detail-" + $strTimeStamp + ".log"
    $objDetailLogFile = [System.IO.StreamWriter] $strDetailLogFilePath
}

#if days to keep log files directive exists in config file, set configured days to keep log files
if ($PowerShellObject.Optional.daysToKeepLogFiles) {
    try {
        $intDaysToKeepLogFiles = $PowerShellObject.Optional.daysToKeepLogFiles
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Using $($PowerShellObject.Optional.daysToKeepLogFiles) value specified in config file for log retention" -LogType "Info" -DisplayInConsole $false
    } catch {
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Warning: $($PowerShellObject.Optional.daysToKeepLogFiles) value specified in config file is not valid, defaulting to unlimited log retention" -LogType "Warning"
    }
}

[bool] $blnFoundSourceDisk = $false
[bool] $blnFoundDestinationDisk = $false
[bool] $blnSuccessfullyObtainedDiskInfo = $false

Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Beginning process to backup all data via robocopy from source disk $($SourceDiskDriveLetter) with disk ID of $($SourceDiskID) to destination disk $($DestinationDiskDriveLetter) with disk ID of $($DestinationDiskID) using robocopy source drive of $($strRoboCopySourceDrive) and robocopy destination drive of $($strRoboCopyDestinationDrive)" -LogType "Info"

try {
    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Getting all disks for comparision" -LogType "Info"
	$objDisks = get-volume

	foreach ($objDisk in $objDisks) {
		if ($objDisk.DriveLetter -eq $SourceDiskDriveLetter -and $objDisk.Path -like "*$SourceDiskID*") {
			$blnFoundSourceDisk = $true
            Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Found source disk" -LogType "Info"
		}
		if ($objDisk.DriveLetter -eq $DestinationDiskDriveLetter -and $objDisk.Path -like "*$DestinationDiskID*") {
			$blnFoundDestinationDisk = $true
            Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Found destination disk" -LogType "Info"
		}
	}
	
	if ($blnFoundSourceDisk -eq $true -and $blnFoundDestinationDisk -eq $true) {
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Successfully validated disk drive letters and IDs of source and destination disks, proceeding to attempt robocopy" -LogType "Info"
		try {
            Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Backing up via robocopy via the following command..." -LogType "Info"
            Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: robocopy $strRoboCopySourceDrive $strRoboCopyDestinationDrive /MIR /XD ""`$RECYCLE.BIN"" ""System Volume Information"" /Z /W:0 /R:1 /nfl /ndl /njh /njs /ns /nc /np" -LogType "Info"
			#$result = robocopy G:\ T:\ /MIR /XD "$RECYCLE.BIN" "System Volume Information" /Z /W:0 /R:1 /nfl /ndl /njh /njs /ns /nc /np
			#$result = robocopy $strRoboCopySourceDrive $strRoboCopyDestinationDrive /MIR /XD "`$RECYCLE.BIN" "System Volume Information" /Z /W:0 /R:1 /nfl /ndl /njh /njs /ns /nc /np
            Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Info: Successfully backed up data via robocopy with result $($result)" -LogType "Info"
		} catch {
			$ErrorMessage = $_.Exception.Message
			$line = $_.InvocationInfo.ScriptLineNumber
			$arrStrErrors += "Failed to backup all data via robocopy from $($SourceDrive) to $($DestinationDrive) with result of $($result) at $($line) with the following error: $ErrorMessage"
            Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Error: Failed to backup all data via robocopy from $($SourceDrive) to $($DestinationDrive) with result of $($result) at $($line) with the following error: $ErrorMessage" -LogType "Error"
		}
	} else {
		$arrStrErrors += "Error: There were no matches for source disk of $($SourceDiskDriveLetter) with disk ID of $($SourceDiskID) and destination disk of $($DestinationDiskDriveLetter) with disk ID of $($DestinationDiskID), so not proceeding.  Please verify the output of get-volume and update the drive letters/IDs if necessary"
		Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Error: There were no matches for source disk of $($SourceDiskDriveLetter) with disk ID of $($SourceDiskID) and destination disk of $($DestinationDiskDriveLetter) with disk ID of $($DestinationDiskID), so not proceeding.  Please verify the output of get-volume and update the drive letters/IDs if necessary" -LogType "Error"
	}
} catch {
	$ErrorMessage = $_.Exception.Message
	$line = $_.InvocationInfo.ScriptLineNumber
	$arrStrErrors += "Failed to get all disks for comparision at $($line) with the following error: $ErrorMessage"
    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $true -LogString "$(get-date) Error: Failed to get all disks for comparision at $($line) with the following error: $ErrorMessage" -LogType "Error"
}

#log retention
if ($intDaysToKeepLogFiles -gt 0) {
    try {
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Info: Purging log files older than $($intDaysToKeepLogFiles) days from $($PowerShellObject.Optional.logsDirectory)" -LogType "Info"
        $CurrentDate = Get-Date
        $DatetoDelete = $CurrentDate.AddDays("-$($intDaysToKeepLogFiles)")
        Get-ChildItem "$($PowerShellObject.Optional.logsDirectory)" | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
    } catch {
        $ErrorMessage = $_.Exception.Message
        $line = $_.InvocationInfo.ScriptLineNumber
        $arrStrErrors += "Failed to purge log files older than $($intDaysToKeepLogFiles) days from $($PowerShellObject.Optional.logsDirectory) with the following error: $ErrorMessage"
        Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Error: Failed to purge log files older than $($intDaysToKeepLogFiles) days from $($PowerShellObject.Optional.logsDirectory) with the following error: $ErrorMessage" -LogType "Error"
    }
}

[int] $intErrorCount = $arrStrErrors.Count

if ($intErrorCount -gt 0) {
    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Info: Encountered $intErrorCount errors, sending error report email" -LogType "Error"
    #loop through all errors and add them to email body
    foreach ($strErrorElement in $arrStrErrors) {
        $intErrorCounter = $intErrorCounter + 1
        $strEmailBody = $strEmailBody + $intErrorCounter.toString() + ") " + $strErrorElement + "<br>"
    }
    $strEmailBody = $strEmailBody + "<br>Please see $strDetailLogFilePath on $strServerName for more details"

    Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Info: Sending email error report via $($errorMailAppID) app on $($errorMailTenantID) tenant from $($errorMailSender) to $($errorMailRecipients) as specified in config file" -LogType "Info"
    $errorEmailPasswordSecure = Get-Content $errorMailPasswordFile | ConvertTo-SecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($errorEmailPasswordSecure)
    $errorEmailPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

    Send-GVMailMessage -sender $errorMailSender -TenantID $errorMailTenantID -AppID $errorMailAppID -subject "$($errorMailSubjectPrefix): Encountered $($intErrorCount) errors during process" -body $strEmailBody -ContentType "HTML" -Recipient $errorMailRecipients -ClientSecret $errorEmailPassword
}

Out-GVLogFile -LogFileObject $objDetailLogFile -WriteToLog $blnWriteToLog -LogString "$(get-date) Info: Process Complete" -LogType "Info"

$objDetailLogFile.close()