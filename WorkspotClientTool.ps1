# Workspot Client Uninstall / Reinstall Tool (This has made up from v5 in local repo)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

############################################################
# FORM
############################################################

$form = New-Object System.Windows.Forms.Form
$form.Text = "Workspot Client Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(520,520)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)

############################################################
# LOAD ICON
############################################################

try{
$iconUrl="https://raw.githubusercontent.com/sathishp-beep/WorkspotScripts/main/workspot.ico"
$wc=New-Object Net.WebClient
$bytes=$wc.DownloadData($iconUrl)
$ms=New-Object IO.MemoryStream(,$bytes)
$form.Icon=[Drawing.Icon]::FromHandle(([Drawing.Bitmap]::FromStream($ms)).GetHicon())
}catch{}

############################################################
# HEADER
############################################################

$header=New-Object Windows.Forms.Panel
$header.Size=New-Object Drawing.Size(520,70)
$header.BackColor=[Drawing.Color]::LightSteelBlue
$form.Controls.Add($header)

$title=New-Object Windows.Forms.Label
$title.Text="Workspot Client Uninstall / Reinstall Tool"
$title.Font=New-Object Drawing.Font("Segoe UI",14,[Drawing.FontStyle]::Bold)
$title.ForeColor=[Drawing.Color]::MidnightBlue
$title.AutoSize=$true
$title.Location=New-Object Drawing.Point(60,22)
$header.Controls.Add($title)

# Logo
try{
$logoUrl="https://raw.githubusercontent.com/sathishp-beep/WorkspotScripts/main/workspot.png"
$wc=New-Object Net.WebClient
$bytes=$wc.DownloadData($logoUrl)
$stream=New-Object IO.MemoryStream(,$bytes)

$logo=New-Object Windows.Forms.PictureBox
$logo.Size=New-Object Drawing.Size(32,32)
$logo.Location=New-Object Drawing.Point(15,18)
$logo.SizeMode="StretchImage"
$logo.Image=[Drawing.Image]::FromStream($stream)
$header.Controls.Add($logo)
}catch{}

############################################################
# STATUS BOX
############################################################

$statusBox=New-Object Windows.Forms.TextBox
$statusBox.Multiline=$true
$statusBox.ReadOnly=$true
$statusBox.ScrollBars="Vertical"
$statusBox.Size=New-Object Drawing.Size(460,220)
$statusBox.Location=New-Object Drawing.Point(20,80)
$form.Controls.Add($statusBox)

function Update-Status($msg){
$statusBox.AppendText("$msg`r`n")
$statusBox.SelectionStart=$statusBox.Text.Length
$statusBox.ScrollToCaret()
[Windows.Forms.Application]::DoEvents()
}

############################################################
# PROGRESS BAR
############################################################

$progressBar=New-Object Windows.Forms.ProgressBar
$progressBar.Size=New-Object Drawing.Size(460,20)
$progressBar.Location=New-Object Drawing.Point(20,310)
$form.Controls.Add($progressBar)

############################################################
# URL INPUT
############################################################
<#

$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Text = "*Workspot Client Download Link for Reinstall:"
$urlLabel.Location = New-Object System.Drawing.Point(20,335)
$urlLabel.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$urlLabel.AutoSize = $true
$form.Controls.Add($urlLabel)
#>

$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Text = "*Provide Workspot Client Download Link for Reinstallation:"
$urlLabel.Location = New-Object System.Drawing.Point(20,335)
$urlLabel.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$urlLabel.ForeColor = [System.Drawing.Color]::Red
$urlLabel.AutoSize = $true
$form.Controls.Add($urlLabel)

$urlTextbox=New-Object Windows.Forms.TextBox
$urlTextbox.Size=New-Object Drawing.Size(460,25)
$urlTextbox.Location=New-Object Drawing.Point(20,360)
$form.Controls.Add($urlTextbox)

############################################################
# BUTTONS
############################################################

$buttonUninstall=New-Object Windows.Forms.Button
$buttonUninstall.Text="Uninstall Workspot Client"
$buttonUninstall.Size=New-Object Drawing.Size(200,40)
$buttonUninstall.Location=New-Object Drawing.Point(60,410)
$buttonUninstall.BackColor=[Drawing.Color]::LightCoral
$buttonUninstall.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($buttonUninstall)

$buttonInstall=New-Object Windows.Forms.Button
$buttonInstall.Text="Reinstall Workspot Client"
$buttonInstall.Size=New-Object Drawing.Size(200,40)
$buttonInstall.Location=New-Object Drawing.Point(260,410)
$buttonInstall.BackColor=[Drawing.Color]::LightBlue
$buttonInstall.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($buttonInstall)

############################################################
# UNINSTALL
############################################################

$buttonUninstall.Add_Click({

Update-Status "`nSearching for Workspot installation..."

$paths = @(
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$app = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
Where-Object { $_.DisplayName -eq "Workspot Client" } |
Select-Object -First 1

if(!$app){

    Update-Status "Workspot Client not found."
    return
}

Update-Status "Workspot Client found: $($app.DisplayName)"

$uninstallString = $app.UninstallString

if([string]::IsNullOrWhiteSpace($uninstallString)){

    Update-Status "Uninstall string not found."
    return
}

Update-Status "Uninstall string detected:"
Update-Status $uninstallString


############################################################
# STOP PROCESSES
############################################################

Update-Status "Stopping Workspot processes..."

Get-Process "*workspot*" -ErrorAction SilentlyContinue |
Stop-Process -Force -ErrorAction SilentlyContinue

Update-Status "Processes stopped."


############################################################
# MSI UNINSTALL
############################################################

$guidMatch = [regex]::Match($uninstallString,"\{[A-F0-9\-]+\}","IgnoreCase")

if($guidMatch.Success){

    $productCode = $guidMatch.Value

    Update-Status "Running MSI uninstall using ProductCode:"
    Update-Status $productCode

    New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null

    Start-Process "msiexec.exe" `
    -ArgumentList "/x $productCode /qn /norestart /l*v C:\Temp\Workspot_Uninstall.log" `
    -Wait

    Update-Status "✓ MSI uninstall completed."

}

Update-Status "Uninstall command completed."

})
############################################################
# DOWNLOAD + INSTALL
############################################################

$buttonInstall.Add_Click({

$url=$urlTextbox.Text.Trim()

if(!$url){
Update-Status "Please paste Workspot MSI URL"
return
}

$buttonInstall.Enabled=$false

New-Item C:\Temp -ItemType Directory -Force | Out-Null
$dest="C:\Temp\WorkspotClient.msi"

Update-Status "Starting download..."

$request=[System.Net.HttpWebRequest]::Create($url)
$response=$request.GetResponse()
$total=$response.ContentLength
$stream=$response.GetResponseStream()

$file=[IO.File]::Create($dest)

$buffer=New-Object byte[] 8192
$read=0
$totalRead=0
$start=Get-Date

while(($read=$stream.Read($buffer,0,$buffer.Length)) -gt 0){

$file.Write($buffer,0,$read)
$totalRead+=$read

$percent=[math]::Round(($totalRead/$total)*100)

$elapsed=(Get-Date)-$start
$speed=($totalRead/1MB)/$elapsed.TotalSeconds
$remain=($total-$totalRead)/1MB
$eta=$remain/$speed

$progressBar.Value=$percent

$status="Download $percent%  {0:N1}/{1:N1} MB  Speed {2:N2} MB/s  ETA {3:N0}s" -f `
($totalRead/1MB),($total/1MB),$speed,$eta

$statusBox.Lines[-1]=$status
[Windows.Forms.Application]::DoEvents()

}

$file.Close()
$response.Close()

Update-Status "Download complete"

############################################################
# INSTALL
############################################################

Update-Status "Starting installation..."

$progressBar.Style="Marquee"

Start-Process msiexec.exe -ArgumentList "/i `"$dest`" /qn /norestart" -Wait

$progressBar.Style="Blocks"
$progressBar.Value=100

Update-Status "✓ Workspot Client installed successfully"

$buttonInstall.Enabled=$true

})

$form.ShowDialog()
