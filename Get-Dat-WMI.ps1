& {
     $wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
     $prp=new-object System.Security.Principal.WindowsPrincipal($wid)
     $adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
     $IsAdmin=$prp.IsInRole($adm)
     if ($IsAdmin)
     {
        Write-host "Script Running With elevated Privladges" 
     }
     Else 
     {
        write-host "Script not running as Admin. please restart powershell with elevated privladges."
        exit
     }
}

# Set New Line format to support MS Notepad
$NLCR = "`r`n"

# Set Output path and Filename
$OutFile = "C:\temp\ALL-DA-WMI-"+(get-date -Format MMddyyyy_HHmm)+".txt"

# Write Filename and Path to screen output 
Write-Host Saving to -> $OutFile

# Get Current list of all WMI objects and itterate over name 
ForEach ($i in (Get-WmiObject -List | Select-Object Name)) {

    # Write current class and time to output file
    $txt= "Current Class-> "+$i.Name+$NLCR+"Current Time: "+(get-date -Format MMddyyyy_HHmm)+$NLCR 
    write-host $txt #| Out-File $OutFile -Append -NoClobber
    
    # Get WMI Object Properties and output to file
    Get-WmiObject $i.Name | Out-File $OutFile -Append -NoClobber
      
}
