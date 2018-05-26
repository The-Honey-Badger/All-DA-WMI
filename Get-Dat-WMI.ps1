$OutFile = "C:\temp\ALL-DA-WMI-"+(get-date -Format MMddyyyy:HHmm)+".txt"

Write-Host Saving to -> $OutFile

$Data = ForEach ($i in (Get-WmiObject -List | Select-Object Name)) {

    echo "Current Class-> " $i.Name | Out-File $OutFile -Append -NoClobber
    Get-WmiObject $i.Name | Out-File $OutFile -Append -NoClobber
      
}

$Data