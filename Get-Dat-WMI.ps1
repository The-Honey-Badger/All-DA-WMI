# Test PowerShell Session for Elevated Privladges 
Write-Host "Testing Permissions"
& {
	$wid = [System.Security.Principal.WindowsIdentity]::GetCurrent()
	$prp = new-object System.Security.Principal.WindowsPrincipal($wid)
	$adm = [System.Security.Principal.WindowsBuiltInRole]::Administrator
	$IsAdmin = $prp.IsInRole($adm)
	if ($IsAdmin)
	{
		Write-host "Session running with Elevated Permissions"
	}
	Else
	{
		write-host "Session is not running with Elevated Permissions. please estart PowerShell with Elevated Rights."
		exit
	}
}

function Log {
	param([string]$filename,[string]$text)
	Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
}

$ErrLogPath = "C:\temp\ALL-DA-WMI-Error.log"

# Set New Line format to support MS Notepad
$NLCR = "`r`n"

# Set Output path and Filename
$OutFile = "C:\temp\ALL-DA-WMI-" + (get-date -Format MMddyyyy_HHmm) + ".txt"

# Write Filename and Path to screen output 
Write-Host Saving to -> $OutFile

# Get Current list of all WMI objects and itterate over name 
ForEach ($i in (Get-WmiObject -List | Select-Object Name)) {

    # Write current class and time to output file
    $txt= "Current Class-> "+$i.Name+$NLCR+"Current Time: "+(get-date -Format MMddyyyy_HHmm)+$NLCR 
    $txt | Out-File $OutFile -Append -NoClobber -encoding ASCII
    
	try {
		
		# Get WMI Object Properties and output to file set EA action to stop to force terminating error on exception
		Get-WmiObject -class $i.Name -EnableAllPrivileges -ErrorAction 'Stop'| Out-File $OutFile -Append -NoClobber -encoding ASCII
	}
	# Catch exceptions thrown by get wmi command
	catch {
		$txt = "ERROR : " + (get-date -Format MMddyyyy_HH:mm:ss) + " Exception caught for " + $i.Name + $NLCR + " Exception Detail : " + $Error[0].Exception
		write-hoste $txt
		log -filename $ErrLogPath -text $txt
	}  
}
