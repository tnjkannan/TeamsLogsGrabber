###Log Grabber for Teams
### (c)2022 Kannan Dharmlinkam, GUI based Teams logs collector tool 
# including URL and other Internet Web site references, may change without notice. You bear the risk of using it. 
# This document does not provide you with any legal rights to any intellectual property in any Microsoft product.
# You may copy and use this document for your internal, reference purposes. You may modify this document for your internal purposes
# 17-Feb 2021 Kannan Dharmalinkam U2 - Updated few lines to make this script work in Nestle, ALso added step to clear cache of Edge browser. 
# 26-Feb 2021 Kannan Dharmalinkam U2 - Added recursive for Edge cache clearing and PS execution policy.
# 11-Mar 2021 Kannan Dharmalinkam U2 - Added a step to copy the Log file to Regional T server with username and time stamp
# 12-Mar 2021 Kannan Dharmalinkam U2 - Added a step to test the connectivity towards the regional T servers and copy accordingly.
# 07-Jul-2022 V 3 Kannan Dharmalinkam Unisys Added Audio logs recording
# 07-Jul-2022 V 4 Kannan Dharmalinkam Unisys Added Cache Clearing option separately
# 07-Jul-2022 V 5 Kannan Dharmalinkam Unisys UI Popup to Choose options to select
# 20-Jul-2022 V 6 Kannan Dharmalinkam Unisys Fixing C:\recording folder error bug
# 20-Jul-2022 V 6 Kannan Dharmalinkam Unisys Encoded the images
# 22-Jul-2022 V 7 Kannan Dharmalinkam Unisys Added Exit button and corrected spelling mistakes and clear VQE logs while clear cache
# 25-Jul-2022 V 8 Kannan Dharmalinkam Unisys Common Version with link in C:\temp\link.txt
Add-Type -AssemblyName System.Windows.Forms 

#Declare
$title    = 'DWS Unisys'
$question = 'Please choose the option as needed - (Microsoft Teams)'
powershell.exe Set-ExecutionPolicy remotesigned -scope currentuser -force
$TimeStamp = (get-date -Format FileDateTimeUniversal)
$userFolder = $env:USERPROFILE
 $username = gc env:username
$NameTarget = "MSTeamsLogs$UserName"
$DateTimeLocalLog = get-date
$RootLogPath = "$env:APPDATA" + "\Microsoft\Teams\"
$RecordLogPath = "c:\recordings\"
$MediaStackPath = $RootLogPath + "media-stack\"
$CallingPath = $RootLogPath + "skylib\"
$RootLogs = $RootLogPath + "\logs.txt"
$DownLoadsFolder = $userFolder + "\Downloads\"
$TeamsStartCommand = $env:LOCALAPPDATA + "\Microsoft\Teams\Update.exe"
$TargetPath = $userFolder + "\Downloads\" + $NameTarget + $TimeStamp
$TargetCompressed = $userFolder + "\Downloads\" + $NameTarget + $TimeStamp + ".zip"
$k=$NameTarget + $TimeStamp + ".zip"
$host.UI.RawUI.WindowTitle = "TeamsLogGrabber"
$NewFolder = New-Item -ItemType Directory -Path $TargetPath

 
$Form = New-Object system.Windows.Forms.Form 
$Form.Text = "TeamsLog Tool V8" 
$Form.TopMost = $true 
$Form.Width = 400 
$Form.Height = 500 
$Form.FormBorderStyle= "Fixed3D" 
$form.MaximizeBox = $false
$form.StartPosition ="centerScreen" 
$form.ShowInTaskbar = $true 
$form.BackColor = "Azure"
  $label9 = New-Object system.windows.Forms.Label 

  #encode the image
$img = "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"
  
  $Image = [Drawing.Bitmap]::FromStream([IO.MemoryStream][Convert]::FromBase64String($Img))
#$Image.Save("c:\temp\U.jpg")

#$label9.Text = "                                            
 
 
 
 #" 
$label9.AutoSize = $true 
$label9.Width = 50 
$label9.Height = 50 
$label9.location = new-object system.drawing.point(15,290) 
$label9.Font = "Microsoft Sans Serif,10,style=Bold" 
$Form.controls.Add($label9) 
#$img1 = [System.Drawing.Image]::Fromfile('c:\temp\u.jpg')
#$img1 = [System.Drawing.Image]::FromStream($Dimage)
 #$label9.Image = $img1
 #background image
 #$objImage = [system.drawing.image]::FromFile("c:\temp\BG2.jpg")
  #$objImage.Size = New-Object System.Drawing.Size(300,170)
  $form.BackgroundImage=$objImage 
  $Form.BackgroundImageLayout = "None"

  $label8 = New-Object system.windows.Forms.Label 
$label8.Text = "Teams Log Tool-Unisys" 
$label8.AutoSize = $true 
$label8.Width = 70 
$label8.Height = 10 
$label8.location = new-object system.drawing.point(55,9) 
$label8.Font = "Microsoft Sans Serif,10,style=Bold" 
$Form.controls.Add($label8) 
$label8.ForeColor = "Blue"

    
$StatusBar = New-Object System.Windows.Forms.StatusBar
$StatusBar.Text = "Ready"
$StatusBar.Height = 22
$StatusBar.Width = 200
$StatusBar.Location = New-Object System.Drawing.Point( 0, 250 )
$Form.Controls.Add($StatusBar)

####### Up Time Button #########
 
$Upbutton = New-Object system.windows.Forms.Button 
$Upbutton.BackColor = "#FFC83D" # #5bd22c
$Upbutton.Text = "Clear Cache and Enable Logs (Before Meeting)" 
$Upbutton.Width = 230 
$Upbutton.Height = 50 
$Upbutton.location = new-object system.drawing.point(70,50) 
$Upbutton.Font = "Microsoft Sans Serif,8,style=Bold" 
$Upbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$Upbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Upbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$Upbutton.Add_Click({Before}) 
$Form.controls.Add($Upbutton)
 
 
 $Afterbutton = New-Object system.windows.Forms.Button 
$Afterbutton.BackColor = "#3EFFFA" 
$Afterbutton.Text = "Collect Teams Logs (After Meeting)" 
$Afterbutton.Width = 230 
$Afterbutton.Height = 50 
$Afterbutton.location = new-object system.drawing.point(70,130) 
$Afterbutton.Font = "Microsoft Sans Serif,8,style=Bold" 
$Afterbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$Afterbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Afterbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$Afterbutton.Add_Click({After}) 
$Form.controls.Add($Afterbutton)


 $Clearbutton = New-Object system.windows.Forms.Button 
$Clearbutton.BackColor = "#b4a0ff" ##ffa500
$Clearbutton.Text = "Clear Teams Cache Only" 
$Clearbutton.Width = 230 
$Clearbutton.Height = 50 
$Clearbutton.location = new-object system.drawing.point(70,210) 
$Clearbutton.Font = "Microsoft Sans Serif,8,style=Bold" 
$Clearbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$Clearbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Clearbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$Clearbutton.Add_Click({Clearfn}) 
$Form.controls.Add($Clearbutton)


 $Exitbutton = New-Object system.windows.Forms.Button 
$Exitbutton.BackColor = "#eb347a" ##ffa500
$Exitbutton.Text = "Exit" 
$Exitbutton.Width = 80 
$Exitbutton.Height = 50 
$Exitbutton.location = new-object system.drawing.point(290,380) 
$Exitbutton.Font = "Microsoft Sans Serif,8,style=Bold" 
$Exitbutton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 255, 36)
$Exitbutton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$Exitbutton.Cursor = [System.Windows.Forms.Cursors]::Hand
$Exitbutton.Add_Click({$form.close()}) 
$Form.controls.Add($Exitbutton)
$l=get-content C:\temp\link.txt
#link Label
$LinkLabel = New-Object System.Windows.Forms.LinkLabel
$LinkLabel.Location = New-Object System.Drawing.Size(60,270)
$LinkLabel.Size = New-Object System.Drawing.Size(250,40)
$LinkLabel.LinkColor = "BLUE"
$LinkLabel.ActiveLinkColor = "RED"
$LinkLabel.Text = "Teams Meeting Best Practices"
$LinkLabel.add_Click({[system.Diagnostics.Process]::start("$l")})
$Form.Controls.Add($LinkLabel)
 
 #Download Label
$LinkLabel1 = New-Object System.Windows.Forms.LinkLabel
$LinkLabel1.Location = New-Object System.Drawing.Size(60,350)
$LinkLabel1.Size = New-Object System.Drawing.Size(250,40)
$LinkLabel1.LinkColor = "BLUE"
$LinkLabel1.ActiveLinkColor = "RED"
$LinkLabel1.Text = "Go To Logs Download Location"
$LinkLabel1.add_Click({[system.Diagnostics.Process]::start("$userFolder\Downloads")})
$Form.Controls.Add($LinkLabel1)

 ##declare

   

 #### Before Meeting functions
 function before 

 {
$statusBar.Text=("Clear cache and Enable Logs (Before Meeting)")
    #Clearing the cache
$statusBar.Text=("We clearing Teams Cache now..pls wait")
Stop-Process -Name Teams; 
Start-Sleep -Seconds 5
Get-ChildItem -Path $env:AppData\Microsoft\Teams -Directory|Where{$_ -in ('Cache','databases','blob_storage','IndexedDB','','tmp','GPUCache')}|ForEach{Remove-Item $_.FullName -Recurse -Force}
    
    write-host "Enabling Logging now..pls wait"
    $Folder = 'C:\temp\recordings'
if (Test-Path -Path $Folder) {
Get-ChildItem -Path C:\temp\recordings | ForEach{Remove-Item $_.FullName -Recurse -Force}
   
} else {
   New-Item -ItemType directory -Path "C:\temp\recordings"
    $statusBar.Text=("C:\temp\Recordings folder created..")
   Start-Sleep -Seconds 1
}

#Check Recording Folder in C Root
    $Folder1 = 'C:\recordings'
if (Test-Path -Path $Folder1) {
Get-ChildItem -Path C:\recordings | ForEach{Remove-Item $_.FullName -Recurse -Force}
   
} else {
   New-Item -ItemType directory -Path "C:\recordings"
   $statusBar.Text=("C:\Recordings folder created..")
   Start-Sleep -Seconds 1
}
    
  #powershell.exe -executionpolicy bypass $script -app_name Teams -set -domain 'MsrtcEcs' -key 'ADSP\EnableUnifiedVQEAPIRecordings' -val '1'
  #powershell.exe -executionpolicy bypass $script -app_name Teams -set -domain 'MsrtcEcs' -key 'ADSP\UnifiedVQEAPIRecordingFileName' -val 'c:\recordings\VqeApiRec*.dat'
$set = "-set"
#write-host $set
pers -app_name Teams $set -domain 'MsrtcEcs' -key 'ADSP\EnableUnifiedVQEAPIRecordings' -val '1'
pers -app_name Teams  $set -domain 'MsrtcEcs' -key 'ADSP\UnifiedVQEAPIRecordingFileName' -val 'c:\recordings\VqeApiRec*.dat'

  #Start Teams again
Start-Sleep -Seconds 5
$statusBar.Text=("Starting Teams App")
Start-Process $TeamsStartCommand -ArgumentList "--processStart Teams.exe" -PassThru
$statusBar.Text=("Teams App Started")
Start-Sleep -Seconds 2
$statusBar.Text=("Ready")
}
##### After Meeting functions
function After
{
$statusBar.Text=("Stopping Teams App..")
Stop-Process -Name Teams; 
Start-Sleep -Seconds 5
#Collect the Audio Recording Logs
    Write-Host 'Stop logging and Collect Microsoft Teams logs (After Meeting)'
    #powershell.exe -executionpolicy bypass $script -app_name Teams -del -domain 'MsrtcEcs' -key 'ADSP\EnableUnifiedVQEAPIRecordings'
    #powershell.exe -executionpolicy bypass $script -app_name Teams -del -domain 'MsrtcEcs' -key 'ADSP\UnifiedVQEAPIRecordingFileName'
    $set = "-del"
$val = "$false"
  Write-Host 'Stop logging and Collect Microsoft Teams logs (After Meeting)'
  $statusBar.Text=("Collect Microsoft Teams logs (After Meeting)..")
pers -app_name Teams $set -domain 'MsrtcEcs' -key 'ADSP\EnableUnifiedVQEAPIRecordings' $val
pers -app_name Teams $set -domain 'MsrtcEcs' -key 'ADSP\UnifiedVQEAPIRecordingFileName' $val
Start-Sleep -Seconds 2
#Collect the Teams Logs 
function writelog($File)
{​
}
function GenLogs()
{
#Start Process for Teams in case not running or closed down to Taskbar  
#$K=(get-process | where {$_.ProcessName -eq "Teams"} | select-object ProcessName​).count 
if ((get-process | where {$_.ProcessName -eq "Teams"} | select-object ProcessName​).count  -gt 0)
    {
        $p = Start-Process $TeamsStartCommand -ArgumentList "--processStart Teams.exe" -Wait -NoNewWindow -PassThru
       $statusBar.Text=("starting Teams client.")
    }
else 
    {
    $statusBar.Text=("Teams client not started, starting Teams client....")
        Write-Host "Teams client not started, starting Teams client."
        $p = Start-Process $TeamsStartCommand -ArgumentList "--processStart Teams.exe" -PassThru
        start-sleep 10
    }
    
$pids = (get-process | where {$_.ProcessName -eq "Teams"}).id
foreach ($process_id in $pids)
    {
    $ShellObj = New-Object -ComObject wscript.shell 
    $ShellObj.AppActivate($process_id) | Out-Null
    start-sleep 1
    $ShellObj.SendKeys("(+%^)1") | Out-Null
    start-sleep 3
    $shellApplication = new-object -com shell.application
$zipPackage = $shellApplication.NameSpace($zipFile)
    $FileResult =  (Get-ChildItem -path $DownLoadsFolder  | Where-Object {($_.Name -match "MSTeams Diagnostics")  -and ($_.LastWriteTime -gt (get-date).AddMinutes(-2))})
    write-host $fileResult,$DownLoadsFolder,$FileResult.count
        if ($FileResult.count -gt 0)
        {
        $ShellObj.AppActivate("TeamsLogGrabber")
              
        foreach ($File in $FileResult)
                {
                $statusBar.Text=("Copying Teams DiagLogs...")
                write-host "Copying Teams DiagLogs..."
        #$zipPackage.CopyHere($file.FullName)
                write-host $file.FullName,$TargetPath
                    Copy-Item $File.FullName -Destination $TargetPath -Recurse
                }
              #Remove-Item 
       break
     }   
    }

   
    
}
function CopyNonWeb()
{
    if ((Get-ChildItem -Path $MediaStackPath | where {$_.name -like "*msrtc*.blog"}).count -gt 0)
        {
        Copy-Item $MediaStackPath -Destination $TargetPath
        Get-ChildItem -Path $MediaStackPath | Copy-Item -Destination "$TargetPath\media-stack\" -Recurse
        }
    else 
        {
        $statusBar.Text=("Warning: Media Stack Logs not Found. Validate Diagnostic logging is enabled in the Teams Client Settings")
        write-host "Warning: Media Stack Logs not Found. Validate Diagnostic logging is enabled in the Teams Client Settings"  -ForegroundColor Yellow
        }
    if ((Get-ChildItem -Path $CallingPath | where {$_.name -like "Debug*.blog"}).count -gt 0)
        {
            Copy-Item $CallingPath -Destination $TargetPath 
            Get-ChildItem -Path $CallingPath | Copy-Item -Destination "$TargetPath\skylib\" -Recurse
        }
    else 
        {
        $statusBar.Text=("Warning: Calling Logs not Found. Validate Diagnostic logging is enabled in the Teams Client Settings")
            write-host "Warning: Calling Logs not Found. Validate Diagnostic logging is enabled in the Teams Client Settings"  -ForegroundColor Yellow
        }
if ((Get-ChildItem -Path $RecordLogPath ).count -gt 0)
        {
        Copy-Item $RecordLogPath -Destination $TargetPath
        Get-ChildItem -Path $RecordLogPath | Copy-Item -Destination "$TargetPath\Recordings\" -Recurse
        }

Copy-Item $RootLogs -Destination $TargetPath
}
function CompressArchive()
    {
        Compress-Archive -LiteralPath $TargetPath -DestinationPath $TargetCompressed
    }
CopyNonWeb
   start-sleep 5
GenLogs

if ((get-childitem -Path $TargetPath).count -gt 0) 
    {
        CompressArchive
    }
Remove-Item $TargetPath -Recurse
write-host "Complete, File written to: $DownLoadsFolder"
 $statusBar.Text=("Complete, File written to: $DownLoadsFolder")
Explorer $DownLoadsFolder


#####

#Start Teams again
Start-Sleep -Seconds 20
Write-host "Start Teams App..pls wait"
 $statusBar.Text=("Start Teams App..pls wait")
Start-Process $TeamsStartCommand -ArgumentList "--processStart Teams.exe" -PassThru
$statusBar.Text=("Started Teams App")
Start-Sleep -Seconds 2
$statusBar.Text=("Ready")
       }

function Clearfn
{
       write-host "Only Clearing Teams Cache"
       #Clearing the cache
write-host "we clearing only Teams Cache..pls wait"
$statusBar.Text=("Clearing Teams Cache files..pls wait")
Start-Sleep -Seconds 2
Stop-Process -Name Teams; 
Start-Sleep -Seconds 5
Get-ChildItem -Path $env:AppData\Microsoft\Teams -Directory|Where{$_ -in ('Cache','databases','blob_storage','IndexedDB','','tmp','GPUCache')}|ForEach{Remove-Item $_.FullName -Recurse -Force}
Get-ChildItem -Path C:\recordings | ForEach{Remove-Item $_.FullName -Recurse -Force}
#Start Teams again
Start-Sleep -Seconds 5
Write-host "Start Teams App.."
$statusBar.Text=("Starting Teams App..pls wait")
Start-Sleep -Seconds 2
Start-Process $TeamsStartCommand -ArgumentList "--processStart Teams.exe" -PassThru
$statusBar.Text=("Started Teams App")
Start-Sleep -Seconds 2
$statusBar.Text=("Ready")
       }

#Persistent Logs 
function Pers
{

Param
(
[Parameter(Mandatory=$true, Position=0)]
[String] $app_name,
[parameter(Mandatory=$true, position=1)]
[String] $set,
[parameter(Mandatory=$true, position=2)]
[String] $domain,
[parameter(Mandatory=$true, position=3)]
[String] $Key,
[parameter(Mandatory=$true, position=4)]
[String] $val
)


#powershell.exe -executionpolicy bypass $script -app_name Teams -set -domain 'MsrtcEcs' -key 'ADSP\EnableUnifiedVQEAPIRecordings' -val '1'
 # powershell.exe -executionpolicy bypass $script -app_name Teams -set -domain 'MsrtcEcs' -key 'ADSP\UnifiedVQEAPIRecordingFileName' -val 'c:\recordings\VqeApiRec*.dat'

$script_name = $MyInvocation.MyCommand.Name
function Usage()
{
    Write-Output("Usage:")
    Write-Output("-print (bool) - prints the persistent.conf file, could be combined with -set or -del options")
    Write-Output("-set (bool) - set key/value to config file, mandatory options are -domain, -key, -val. Can't be used together with -del")
    Write-Output("-del (bool) - deletes key or domain from config file, mandatory option is -domain. Can't be used together with -set")
    Write-Output("-app_name (string) - the app name, that will be searched in %appdata% folder (e.g. Teams or Skype for Desktop). If is omitted local peristent.conf is used")
    Write-Output("-domain (string) - domain, like SkypeRT, MsrtcEcs")
    Write-Output("-key (string) - key to set or delete")
    Write-Output("-val (string) - value to set")
    Write-Output("Exmaples:")
    Write-Output(".\$script_name -print -app_name Teams")
    Write-Output(".\$script_name -set -domain 'SkypeRT' -key 'testKey' -val 'testValue'")
    Write-Output(".\$script_name -del -app_name 'Skype for Desktop' -domain 'TestDomain' -key 'testKey'")
    Write-Output(".\$script_name -print -del -app_name 'TestApp' -domain 'TestDomain'")
}

function GetPersistentConfig($path)
{
    $json = New-Object -TypeName PSCustomObject
    if (Test-Path $path) {
        $conf = Get-Content -Path $path
        $json = $conf | ConvertFrom-Json
    }
    $json 
}

function WritePersistentConfig($path, $json)
{
    $res = $json | ConvertTo-Json
    Set-Content -Path $path -Value $res
}

function SetConfigKey($path, $domain, $key, $val)
{
    Write-Debug "SetConfigKey: domain=$domain key=$key val=$val"
    $json = GetPersistentConfig $path

    if ($json.$domain -eq $null) {
        $obj = New-Object -TypeName PSCustomObject
        $json | Add-Member $domain $obj
    }

    if ($json.$domain.$key -ne $null) {
        $json.$domain.$key = $val
    }
    else {
        $json.$domain | Add-Member $key $val
    }

    WritePersistentConfig $path $json
}

function DeleteConfigKey($path, $domain, $key)
{
    Write-Debug "DeleteConfigKey: domain=$domain key=$key"
    $json = GetPersistentConfig $path

    if ($json.$domain -eq $null) {
        return
    }
    $json.$domain.PSObject.Properties.Remove($key)
    WritePersistentConfig $path $json
}

function DeleteConfigDomain($path, $domain)
{
    Write-Debug "DeleteConfigDomain: $domain"
    $json = GetPersistentConfig $path
    if ($json.$domain -eq $null) {
        return
    }
    $json.PSObject.Properties.Remove($domain)
    WritePersistentConfig $path $json
}

function PrintConfig($path)
{
    # to prettify output
    $json = GetPersistentConfig $path
    $res = $json | ConvertTo-Json
    Write-Output $res
}


##################################

if ($help) {
    Usage
    return
}

if (!$print -and !$set -and !$del) {
    Write-Warning "either -print or -set or -del option should be provided"
    Usage
    return
}


if ($app_name) {
    $path = [Environment]::GetFolderPath('ApplicationData') +  "\Microsoft\" + $app_name + "\SkypeRT\persistent.conf"
    Write-Output "$path will be examined."
}
else {
    $path = ".\persistent.conf"
    Write-Output "-app_name is not set. persistent.conf in current folder will be examined."
}

if ($set -and $del) {
    Write-Warning "-set and -del should not be used together"
    Usage
    return
}

if ($set) {
    if (!$domain -or !$key -or !$val) {
        Write-Warning "-set: domain, key or val paramter is not set"
        Usage
        return
    }
    SetConfigKey $path $domain $key $val
}

if ($del) {
    if (!$domain) {
        Write-Warning "-del: domain parameter is not set"
        Usage
        return
    }
    if ($key) {
        DeleteConfigKey $path $domain $key
    }
    else {
        DeleteConfigDomain $path $domain
    }
}

if ($print) {
    Write-Output $path
    PrintConfig $path
}
}

       
[void]$Form.ShowDialog() 
$Form.Dispose()

