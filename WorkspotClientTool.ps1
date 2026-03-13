Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

############################################################
# FORM
############################################################

$form = New-Object Windows.Forms.Form
$form.Text = "Workspot Client Support Tool"
$form.Size = New-Object Drawing.Size(650,640)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object Drawing.Font("Segoe UI",10)

# Optional but recommended for support tools
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

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

$header = New-Object Windows.Forms.Panel
$header.Size = New-Object Drawing.Size(650,70)
$header.BackColor = [Drawing.Color]::SteelBlue
$form.Controls.Add($header)

$title = New-Object Windows.Forms.Label
$title.Text = "Workspot Client Support Tool"
$title.Font = New-Object Drawing.Font("Segoe UI",16,[Drawing.FontStyle]::Bold)
$title.ForeColor = [Drawing.Color]::White
$title.Location = New-Object Drawing.Point(20,15)
$title.AutoSize = $true

$header.Controls.Add($title)

############################################################
# STATUS BOX
############################################################

$statusBox = New-Object Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ReadOnly = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.Size = New-Object Drawing.Size(600,230)
$statusBox.Location = New-Object Drawing.Point(20,90)

$form.Controls.Add($statusBox)

function Update-Status($msg){

$statusBox.AppendText("$msg`r`n")
$statusBox.SelectionStart = $statusBox.Text.Length
$statusBox.ScrollToCaret()

[Windows.Forms.Application]::DoEvents()

}

############################################################
# PROGRESS BAR
############################################################

$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Size = New-Object Drawing.Size(600,18)
$progressBar.Location = New-Object Drawing.Point(20,330)

$form.Controls.Add($progressBar)

############################################################
# URL FIELD
############################################################

$urlLabel = New-Object Windows.Forms.Label
$urlLabel.Text = "*Workspot Client Download Link for Install / Reinstall:"
$urlLabel.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$urlLabel.ForeColor = [Drawing.Color]::DarkRed
$urlLabel.Location = New-Object Drawing.Point(20,360)
$urlLabel.AutoSize = $true
$form.Controls.Add($urlLabel)

$urlTextbox = New-Object Windows.Forms.TextBox
$urlTextbox.Size = New-Object Drawing.Size(600,25)
$urlTextbox.Location = New-Object Drawing.Point(20,385)

$form.Controls.Add($urlTextbox)

############################################################
#CONTROL & GATEWAY URL FIELD
############################################################

$gwLabel = New-Object System.Windows.Forms.Label
$gwLabel.Text = " Connection Check to Workspot Control / Gateway:"
$gwLabel.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$gwLabel.Location = New-Object System.Drawing.Point(20,420)
$gwLabel.AutoSize = $true
$form.Controls.Add($gwLabel)

# Control URL
$portalTextbox = New-Object System.Windows.Forms.TextBox
$portalTextbox.Size = New-Object System.Drawing.Size(280,25)
$portalTextbox.Location = New-Object System.Drawing.Point(20,445)
$portalTextbox.Text = "control.workspot.com"
$form.Controls.Add($portalTextbox)

# Gateway URL
$gatewayTextbox = New-Object System.Windows.Forms.TextBox
$gatewayTextbox.Size = New-Object System.Drawing.Size(280,25)
$gatewayTextbox.Location = New-Object System.Drawing.Point(320,445)
$gatewayTextbox.Text = "<PROVIDE GATEWAY URL HERE>"
$form.Controls.Add($gatewayTextbox)

############################################################
# BUTTON CREATION FUNCTION
############################################################

function New-ToolButton($text,$x,$y){

$btn = New-Object Windows.Forms.Button
$btn.Text = $text
$btn.Size = New-Object Drawing.Size(180,40)
$btn.Location = New-Object Drawing.Point($x,$y)
$btn.Font = New-Object Drawing.Font("Segoe UI",10,[Drawing.FontStyle]::Bold)

$form.Controls.Add($btn)

return $btn
}

############################################################
# BUTTONS
############################################################

$btnDetect = New-ToolButton "Detect Client" 20 430
$btnUninstall = New-ToolButton "Uninstall" 220 430
$btnCleanup = New-ToolButton "Deep Cleanup" 420 430

$btnReinstall = New-ToolButton "Install / Reinstall" 20 480
$btnLogs = New-ToolButton "Collect Logs" 220 480
$btnTest = New-ToolButton "Test Workspot Connectivity" 420 480


$btnDetect.Location = New-Object Drawing.Point(20,490)
$btnUninstall.Location = New-Object Drawing.Point(220,490)
$btnCleanup.Location = New-Object Drawing.Point(420,490)

$btnReinstall.Location = New-Object Drawing.Point(20,540)
$btnLogs.Location = New-Object Drawing.Point(220,540)
$btnTest.Location = New-Object Drawing.Point(420,540)
############################################################
# DETECT CLIENT
############################################################

$btnDetect.Add_Click({

Update-Status "`nDetecting Workspot Client..."

$paths = @(
"HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
"HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$app = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
Where-Object { $_.DisplayName -eq "Workspot Client" } |
Select-Object -First 1

if ($app) {

    Update-Status "Client Found:"
    Update-Status "Name: $($app.DisplayName)"
    Update-Status "Version: $($app.DisplayVersion)"

}
else {

    Update-Status "Workspot Client not installed"

}

})

############################################################
# UNINSTALL
############################################################

$btnUninstall.Add_Click({

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

    #################################################################

    # STEP 2 Remove AppData
        Update-Status "`n[Step 2] Removing AppData..."

        $path = "$env:LOCALAPPDATA\Workspot"

        if (Test-Path $path) {
            Remove-Item $path -Recurse -Force
            Update-Status "✓ Removed $path"
        }
        else {
            Update-Status "⚠ AppData folder not found"
        }

        # STEP 3 Registry cleanup
        Update-Status "`n[Step 3] Removing registry..."

        $reg = "HKCU:\Software\Workspot"

        if (Test-Path $reg) {
            Remove-Item $reg -Recurse -Force
            Update-Status "✓ Registry removed"
        }
        else {
            Update-Status "⚠ Registry key not found"
        }

        Update-Status "`n================================"
        Update-Status "Cleanup completed."

        ##########################################################################


}

Update-Status "Uninstall command completed."

})

############################################################
# CLEANUP
############################################################

$btnCleanup.Add_Click({

Update-Status "`nRunning deep cleanup..."

$paths=@(
"$env:ProgramFiles\Workspot",
"$env:ProgramFiles(x86)\Workspot",
"$env:ProgramData\Workspot",
"$env:LOCALAPPDATA\Workspot",
"$env:LOCALAPPDATA\Programs\Workspot"
)

foreach($p in $paths){

if(Test-Path $p){

Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
Update-Status "Removed $p"

}
else {
    Update-Status "Check the Path not found: $p or You may not have permissions to delete."
}
}

Update-Status "Cleanup completed"

})

############################################################
# COLLECT LOGS
############################################################

$btnLogs.Add_Click({

Update-Status "`nCollecting logs..."

$dest="C:\Temp\WorkspotLogs.zip"

$logPaths=@(
"$env:LOCALAPPDATA\Workspot\Client\log"
)

Compress-Archive $logPaths $dest -Force

Update-Status "Logs exported to $dest"

})

############################################################
# TEST GATEWAY & Workspot URLs
############################################################

$btnTest.Add_Click({

Update-Status "`nTesting connectivity..."

$servers = @(
    $portalTextbox.Text.Trim(),
    $gatewayTextbox.Text.Trim()
)

foreach($s in $servers){

    # Ignore empty fields
    if([string]::IsNullOrWhiteSpace($s)){
        continue
    }

    $result = Test-NetConnection $s -Port 443 -WarningAction SilentlyContinue

    if($result.TcpTestSucceeded){

        Update-Status "$s reachable on port 443"

    }
    else{

        Update-Status "$s NOT reachable"

    }

}

})

############################################################
# DOWNLOAD + INSTALL
############################################################

$btnReinstall.Add_Click({

$url=$urlTextbox.Text.Trim()

if(!$url){
Update-Status "Please paste Workspot MSI URL in the Client Download Link field below the status box."
return
}
$btnReinstall.Enabled=$true
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

$btnReinstall.Enabled=$true

})

$form.ShowDialog()
