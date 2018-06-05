Add-Type –assemblyName PresentationFramework
 
# Create Initial Runspace For GUI
$Runspace = [runspacefactory]::CreateRunspace()
$Runspace.ApartmentState = "STA"
$Runspace.ThreadOptions = "ReuseThread"
$Runspace.Open()
 
# Encapsulate Entire GUI In ScriptBlock
$code = {
#Build the GUI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" x:Name="MainWindow"        
    Title="WMI Ass Ripper" Height="692.4" Width="1274.4" WindowStartupLocation="CenterScreen" Cursor="IBeam">
    <Window.Resources>
        <Style TargetType="{x:Type Button}">
            <Setter Property="Background" Value="Green"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border Background="{TemplateBinding Background}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="DarkGoldenrod"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Background="#FF565656" Margin="0,0,-6,-0.8">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="8*"/>
            <ColumnDefinition Width="375*"/>
            <ColumnDefinition Width="890*"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="Error_TxtBox" HorizontalAlignment="Left" Height="415" Margin="6.8,211,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="834" FontFamily="Comic Sans MS" FontSize="18" Grid.Column="2" Background="Silver"/>
        <Button x:Name="Run_Btn" Content="Zhu Li, Do The Thing!" Margin="34,42,85.2,368" Background="#FFFF5C00" FontWeight="ExtraBold" FontFamily="Comic Sans MS" FontSize="22" Cursor="Hand" Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center" Width="256" Height="254"/>
        <TextBlock x:Name="CurrentWMI_TxtBox" HorizontalAlignment="Left" Margin="6.8,42,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="44" Width="834" FontFamily="Comic Sans MS" Background="Silver" Grid.Column="2" FontSize="14"/>
        <TextBlock x:Name="Progress_TxtBox" HorizontalAlignment="Left" Margin="6.8,125,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="45" Width="834" FontFamily="Comic Sans MS" Background="Silver" Grid.Column="2" FontSize="14"/>
        <Label x:Name="Error_Box_Label" Content="Error Output:" Grid.Column="2" HorizontalAlignment="Left" Margin="6.8,175,0,0" VerticalAlignment="Top" FontFamily="Comic Sans MS" FontSize="16" FontWeight="Bold" Height="31" Width="119"/>
        <Button x:Name="Stop_Btn" Content="ABORT! ABORT! ABORT! &#xD;&#xA;WE'VE DUG TOO DEEP!" HorizontalAlignment="Center" Margin="34,372,86.2,38" VerticalAlignment="Center" Width="255" Height="254" FontFamily="Comic Sans MS" FontSize="20" FontWeight="Bold" Background="#FFFC3D3D" Cursor="Hand" Grid.Column="1"/>
        <Label x:Name="Current_WMI_Label" Content="Current WMI Class: " Grid.Column="2" HorizontalAlignment="Left" Margin="6.8,10,0,0" VerticalAlignment="Top" FontFamily="Comic Sans MS" FontSize="16" FontWeight="Bold" Height="31" Width="177"/>
        <Label x:Name="Progress_Box_Label" Content="Progress:" Grid.Column="2" HorizontalAlignment="Left" Margin="6.8,86,0,0" VerticalAlignment="Top" FontFamily="Comic Sans MS" FontSize="16" FontWeight="Bold" Height="31" Width="177"/>
        
    </Grid>
</Window>
"@

# Create Synchronized Hash Table
$Global:syncHash = [hashtable]::Synchronized(@{})
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
 


 # Create Runspace's on the fly
function Start-Runspace{
    param($scriptblock)
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"         
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("SyncHash",$global:syncHash)
    $psCmd = [PowerShell]::Create().AddScript($ScriptBlock)
    $psCmd.Runspace = $newRunspace
    $psCMD.BeginInvoke()
}



# Counter To Test Runspace stuff...
$WorkingCounter = {
    $Counter=0
    While ($Counter -lt 10){
        $Global:syncHash.Progress_TxtBox.Dispatcher.Invoke([action]{$Global:syncHash.Progress_TxtBox.Text = $Counter}, "Normal")
        $Counter++
        Sleep -Seconds 1
    }
}


# Set Some Global Vars
$Global:syncHash.CurrentClass = "No Class Set"
$Global:syncHash.Min = 0

# Progress Tracker To Watch WMI Collection Progress
$ProgTracker = {
    $Global:syncHash.Progress_TxtBox.Dispatcher.Invoke([action]{$Global:syncHash.Progress_TxtBox.Text = "Calculating..."}, "Normal")
    $Max = (Get-WmiObject -list | select Name)
    While ($Global:syncHash.Min -lt $Max.count){
        $GLOBAL:syncHash.Progress_TxtBox.Dispatcher.Invoke([action]{$Global:syncHash.Progress_TxtBox.Text=("" + $Global:syncHash.Min + " Out Of " + $Max.Count + " Done")}, "Normal")
        $GLOBAL:syncHash.CurrentWMI_TxtBox.Dispatcher.Invoke([action]{$Global:syncHash.CurrentWMI_TxtBox.Text = $Global:syncHash.CurrentClass}, "Normal")

        #Sleep -Seconds 2
        #$Min=$Min+300
        #$Global:syncHash.Min = $Min
    }
    $Global:syncHash.Progress_TxtBox.Dispatcher.Invoke([action]{$Global:syncHash.Progress_TxtBox.Text = "End Of The Line Mate"}, "Normal")
}


# Main WMI Script Stuff
$CollectWMI = {
    #$Global:syncHash.CurrentWMI_TxtBox.Dispatcher.Invoke([action]{$GLOBAL:syncHash.CurrentWMI_TxtBox.Text = $Global:syncHash.CurrentClass}, "Normal")



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


    # Get Current list of all WMI objects and itterate over name 
    ForEach ($i in (Get-WmiObject -List | Select-Object Name)) {

    $Global:syncHash.Min = $Global:syncHash.Min + 1
    $Global:syncHash.CurrentClass = $i.Name

    # Write Current Class Current WMI Textbox
    #$GLOBAL:syncHash.CurrentWMI_TxtBox.Dispatcher.Invoke([action]{$GLOBAL:syncHash.CurrentWMI_TxtBox.Text = $GLOBAL:syncHash.Min}, "Normal")


    # Write current class and time to output file
    $txt= "Current Class-> "+$i.Name+$NLCR+"Current Time: "+(get-date -Format "MM/dd/yyyy HH:mm:ss")+$NLCR |
        Out-File $OutFile -Append -NoClobber -encoding ASCII
    
	try {
		
		# Get WMI Object Properties and output to file set EA action to stop to force terminating error on exception
		Get-WmiObject -class $i.Name -EnableAllPrivileges -ErrorAction 'Stop'| Out-File $OutFile -Append -NoClobber -encoding ASCII
    }

	# Catch exceptions thrown by get wmi command
	catch {
		$txt = (get-date -Format "MM/dd/yyyy HH:mm:ss") + " -- Exception caught for " + $i.Name + $NLCR + " Exception Detail : " + $Error[0].Exception + $NLCR

        $Global:syncHash.Error_TxtBox.Text=$txt}
		log -filename $ErrFile -text $txt
	}  
}


# Function To Update Error Text Box.
function UpdateTextBox {$syncHash.Error_TxtBox.Text="This Is A Work In Progress"}

# Function To Close GUI When Abort Button Pressed
function StopButtonPressed {$syncHash.window.close()}
 
# XAML objects
# Run Button
$syncHash.Run_Btn = $syncHash.window.FindName("Run_Btn")

# Stop Button
$syncHash.Stop_Btn = $syncHash.window.FindName("Stop_Btn")

# Current WMI Class Textbox
$syncHash.CurrentWMI_TxtBox = $syncHash.window.FindName("CurrentWMI_TxtBox")

# Progress Tracker Textbox
$syncHash.Progress_TxtBox = $syncHash.window.FindName("Progress_TxtBox")

# Error Textbox
$syncHash.Error_TxtBox = $syncHash.window.FindName("Error_TxtBox")


# Define Button Click Actions
$syncHash.Run_Btn.Add_click({UpdateTextBox;Start-Runspace $ProgTracker;Sleep -Seconds 3;Start-Runspace $CollectWMI})
$syncHash.Stop_Btn.Add_click({StopButtonPressed})
#$syncHash.Stop_Btn.Add_click({Start-Runspace $CollectWMI})



# Show GUI
$syncHash.Window.ShowDialog()
$Runspace.Close()
$Runspace.Dispose()
 
}
 
# Put GUI ScriptBlock In Its Own RunSpace And Start Running It
$PSinstance1 = [powershell]::Create().AddScript($Code)
$PSinstance1.Runspace = $Runspace
$job = $PSinstance1.BeginInvoke()