# Workspot Client Uninstall & Cleanup GUI Tool

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Workspot Client Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(520,450)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI",10)

##############################################################
# Load Workspot Icon from GitHub
##############################################################

try{
$iconUrl = "https://raw.githubusercontent.com/sathishp-beep/WorkspotScripts/main/workspot.ico"

$wc = New-Object System.Net.WebClient
$bytes = $wc.DownloadData($iconUrl)

$ms = New-Object System.IO.MemoryStream
$ms.Write($bytes,0,$bytes.Length)
$ms.Position = 0

$form.Icon = [System.Drawing.Icon]::FromHandle(
([System.Drawing.Bitmap]::FromStream($ms)).GetHicon()
)
}catch{}

##############################################################
# Header Banner
##############################################################

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Size = New-Object System.Drawing.Size(520,70)
$headerPanel.Location = New-Object System.Drawing.Point(0,0)
$headerPanel.BackColor = [System.Drawing.Color]::LightSteelBlue
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Workspot Client Uninstall / Reinstall Tool"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::MidnightBlue
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(60,22)

$headerPanel.Controls.Add($titleLabel)
# -------------------------------------------------
# Workspot Logo in Header (Load PNG from URL)
# -------------------------------------------------

try {

$logoUrl = "https://raw.githubusercontent.com/sathishp-beep/WorkspotScripts/main/workspot.png"

$logoBox = New-Object System.Windows.Forms.PictureBox
$logoBox.Size = New-Object System.Drawing.Size(32,32)
$logoBox.Location = New-Object System.Drawing.Point(15,18)
$logoBox.SizeMode = "StretchImage"

$webClient = New-Object System.Net.WebClient
$bytes = $webClient.DownloadData($logoUrl)

$stream = New-Object System.IO.MemoryStream
$stream.Write($bytes,0,$bytes.Length)
$stream.Position = 0

$logoBox.Image = [System.Drawing.Image]::FromStream($stream)

$headerPanel.Controls.Add($logoBox)

}
catch{
   # Write-Host "Logo failed to load"
}
##############################################################
# Status Box
##############################################################

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Location = New-Object System.Drawing.Point(20,80)
$statusBox.Size = New-Object System.Drawing.Size(460,210)
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$form.Controls.Add($statusBox)

function Update-Status {

param([string]$msg)

$statusBox.AppendText("$msg`r`n")
$statusBox.SelectionStart = $statusBox.Text.Length
$statusBox.ScrollToCaret()

[System.Windows.Forms.Application]::DoEvents()

}

##############################################################
# MSI URL INPUT
##############################################################

$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Text = "Workspot Client Download Link:"
$urlLabel.Location = New-Object System.Drawing.Point(20,300)
$urlLabel.Size = New-Object System.Drawing.Size(300,20)
$urlLabel.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($urlLabel)

$urlTextbox = New-Object System.Windows.Forms.TextBox
$urlTextbox.Location = New-Object System.Drawing.Point(20,325)
$urlTextbox.Size = New-Object System.Drawing.Size(460,25)
$urlTextbox.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)
$form.Controls.Add($urlTextbox)

##############################################################
# Uninstall Button
##############################################################

$buttonUninstall = New-Object System.Windows.Forms.Button
$buttonUninstall.Text = "Uninstall Workspot Client"
$buttonUninstall.Size = New-Object System.Drawing.Size(200,40)
$buttonUninstall.Location = New-Object System.Drawing.Point(60,365)
$buttonUninstall.BackColor = [System.Drawing.Color]::LightCoral
$buttonUninstall.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($buttonUninstall)

##############################################################
# Reinstall Button
##############################################################

$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Re-Install Workspot Client"
$buttonInstall.Size = New-Object System.Drawing.Size(200,40)
$buttonInstall.Location = New-Object System.Drawing.Point(260,365)
$buttonInstall.BackColor = [System.Drawing.Color]::LightBlue
$buttonInstall.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($buttonInstall)

##############################################################
# UNINSTALL LOGIC
##############################################################

$buttonUninstall.Add_Click({

Update-Status "`nStarting Workspot uninstall..."

New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null

Update-Status "[Step 0] Stopping Workspot processes..."

Get-Process "*workspot*" -ErrorAction SilentlyContinue |
Stop-Process -Force -ErrorAction SilentlyContinue

Update-Status "✓ Processes stopped"

Update-Status "`n[Step 1] Searching for Workspot installation..."

$paths = @(
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$app = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
Where-Object { $_.DisplayName -eq "Workspot Client" }

if (!$app){

Update-Status "⚠ Workspot Client not found"

}
else{

$uninstallString = $app.UninstallString

if ($uninstallString -match "MsiExec"){

$productCode = $uninstallString -replace '.*(\{.*\}).*', '$1'

$arguments = "/x $productCode /qn /norestart /l*v C:\Temp\Workspot_Uninstall.log"

Start-Process "msiexec.exe" -ArgumentList $arguments -Wait

Update-Status "✓ Workspot Client uninstalled"

}
else{

$exe = $uninstallString -replace '"',''

Start-Process "cmd.exe" -ArgumentList "/c $exe" -Wait

Update-Status "✓ EXE uninstall executed"

}

}

Update-Status "`nCleanup completed."

})

##############################################################
# REINSTALL LOGIC
##############################################################

$buttonInstall.Add_Click({

$url = $urlTextbox.Text

if([string]::IsNullOrWhiteSpace($url)){

Update-Status "⚠ Please enter MSI download URL first."
return

}

New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null

$msiPath = "C:\Temp\WorkspotClient.msi"

try{

Update-Status "`nDownloading MSI..."

Invoke-WebRequest $url -OutFile $msiPath

Update-Status "✓ Download completed"

Update-Status "Installing Workspot Client silently..."

$arguments = "/i `"$msiPath`" /qn /norestart /l*v C:\Temp\Workspot_Install.log"

Start-Process "msiexec.exe" -ArgumentList $arguments -Wait

Update-Status "✓ Workspot Client installed successfully"

}
catch{

Update-Status "✗ Installation failed: $_"

}

})

$form.ShowDialog() | Out-Null