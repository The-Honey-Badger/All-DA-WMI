#Requires -RunAsAdministrator

# Add GUI Type
Add-Type -AssemblyName System.Windows.Forms

# Create GUI Object
$Form = New-Object system.Windows.Forms.Form

# Set Default Font ||| Font styles are: Regular, Bold, Italic, Underline, Strikeout
$Font = New-Object System.Drawing.Font("Comic Sans MS",18,[System.Drawing.FontStyle]::Bold)


# Set Form Properties
$Form.Text = "WMI Ass Ripper"
$Form.Font = $Font
$Form.Width = "600"
$Form.Height = "400"
$Form.BackColor="Brown"
$Form.WindowState = "Maximized"


# Create Lable/Title
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Do You Really Want To Do This?"
$Label.ForeColor  = "Red"
$Label.AutoSize = $True
$Form.Controls.Add($Label)


$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Size(35,100)
$Button.Size = New-Object System.Drawing.Size(120,23)
$Button.ForeColor="Green"
$Button.AutoSize = $True
$Button.Text = "Just Fuck My Shit Up"
$Button.Visible = $True
$Button.Enabled = $True
$Button.FlatStyle = "Flat"
$Button.AccessibilityObject
$Button.Add_Click($JFMSU_Click)
$Form.Controls.Add($Button)

$OutputBox = New-Object System.Windows.Forms.TextBox
$OutputBox.Location = New-Object System.Drawing.Size(700,50)
$OutputBox.Size = New-Object System.Drawing.Size(1000,800)
$OutputBox.Multiline = $True
$OutputBox.ReadOnly = $True
$OutputBox.BackColor="#33FF44"
$OutputBox.ForeColor="White"
$Form.Controls.Add($OutputBox)


$JFMSU_Click = {Update_textbox -String "OK Ass Hat, Here We Go....";GetDaWMI}

# Function for updating output text box. Does not work in GetDaWMI function...
function Update_textbox {
    param([string]$String)
    #$OutputBox.Clear()
    $OutputBox.AppendText("`r`n$String")
}


function Log {
	param([string]$filename,[string]$text)
	Out-File $filename -append -noclobber -inputobject $text -encoding ASCII
}


# Function to kick off the fun
function GetDaWMI {

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

    #Update-textbox -String Saving Output to -> $OutFile
    $OutputBox.AppendText("`r`nSaving Output to -> $OutFile")
    $OutputBox.AppendText("`r`nSaving Errors to -> $ErrFile")

    


    # Count WMI Objects and print to screen
    $OutputBox.AppendText("`r`nCurrently there are " +(Get-WmiObject -List ).count + " WMI objects.")

    # Start The Fun
    $OutputBox.AppendText("`r`nLet us begin...")


    # Get Current list of all WMI objects and itterate over name 
    ForEach ($i in (Get-WmiObject -List | Select-Object Name)) {

        # Write current class and time to output file
        $txt= "Current Class-> "+$i.Name+$NLCR+"Current Time: "+(get-date -Format "MM/dd/yyyy HH:mm:ss")+$NLCR |
            Out-File $OutFile -Append -NoClobber -encoding ASCII
        $OutputBox.AppendText("`r`n$txt")
    
	    try {
		
		    # Get WMI Object Properties and output to file set EA action to stop to force terminating error on exception
		    Get-WmiObject -class $i.Name -EnableAllPrivileges -ErrorAction 'Stop'| Out-File $OutFile -Append -NoClobber -encoding ASCII
        }

	    # Catch exceptions thrown by get wmi command
	    catch {
		    $txt = (get-date -Format "MM/dd/yyyy HH:mm:ss") + " -- Exception caught for " + $i.Name + $NLCR + " Exception Detail : " + $Error[0].Exception + $NLCR
		    #write-host $txt
		    log -filename $ErrFile -text $txt
	    }  
    }
#>

}

# Show The Form
$Form.ShowDialog()
