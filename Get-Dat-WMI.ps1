#Requires -RunAsAdministrator


function Log {
	param([string]$filename,[string]$text)
	Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
}


# Set New Line format to support MS Notepad
$NLCR = "`r`n"


# Set Output File
$OutFile = "C:\temp\ALL-DA-WMI-" + (get-date -Format MMddyyyy_HHmm) + ".txt"


# Set ErrLog File
$ErrFile = $OutFile.Insert(($OutFile.lastIndexOf('-')), "-Errors")


# Test For Path / File. If non-exsistant, create it
If (!(Test-Path $OutFile)) {New-Item -Path $OutFile -Force | Out-Null}
If (!(Test-Path $ErrFile)) {New-Item -Path $ErrFile -Force | Out-Null}


# Write Filenames and Paths to screen output
""
Write-Host Saving Output to -> $OutFile
Write-Host Saving Errors to -> $ErrLogPath
""

# Count WMI Objects and print to screen
Write-Host Currently there are (Get-WmiObject -List ).count WMI objects. 
""

# Start The Fun
Write-Host Let us begin...


# Get Current list of all WMI objects and itterate over name 
ForEach ($i in (Get-WmiObject -List | Select-Object Name)) {

    # Write current class and time to output file
    $txt= "Current Class-> "+$i.Name+$NLCR+"Current Time: "+(get-date -Format MMddyyyy_HHmm)+$NLCR |
        Out-File $OutFile -Append -NoClobber -encoding ASCII
    
	try {
		
		# Get WMI Object Properties and output to file set EA action to stop to force terminating error on exception
		Get-WmiObject -class $i.Name -EnableAllPrivileges -ErrorAction 'Stop'| Out-File $OutFile -Append -NoClobber -encoding ASCII
	}

	# Catch exceptions thrown by get wmi command
	catch {
		$txt = "ERROR : " + (get-date -Format MMddyyyy_HH:mm:ss) + " Exception caught for " + $i.Name + $NLCR + " Exception Detail : " + $Error[0].Exception
		write-hoste $txt
		log -filename $ErrFile -text $txt
	}  
}
