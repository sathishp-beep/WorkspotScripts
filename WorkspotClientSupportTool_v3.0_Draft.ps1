Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

############################################################
# FORM
############################################################

$form = New-Object Windows.Forms.Form
$form.Text = "Workspot Client Support Tool"
$form.Size = New-Object Drawing.Size(700, 720)
$form.StartPosition = "CenterScreen"
$form.AutoScroll = $true
$form.AutoScrollMargin = New-Object Drawing.Size(0, 100)
$form.Font = New-Object Drawing.Font("Segoe UI", 10)
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false

############################################################
# LOAD ICON
############################################################

try {
    $iconUrl = "https://raw.githubusercontent.com/sathishp-beep/WorkspotScripts/main/workspot.ico"
    $wc = New-Object Net.WebClient
    $bytes = $wc.DownloadData($iconUrl)
    $ms = New-Object IO.MemoryStream(, $bytes)
    $form.Icon = [Drawing.Icon]::FromHandle(([Drawing.Bitmap]::FromStream($ms)).GetHicon())
}
catch {}


############################################################
# HEADER
############################################################

$header = New-Object Windows.Forms.Panel
$header.Size = New-Object Drawing.Size(700, 60)
$header.BackColor = [Drawing.Color]::DarkBlue
$form.Controls.Add($header)

$title = New-Object Windows.Forms.Label
$title.Text = "Workspot Client Support Tool"
$title.ForeColor = "White"
$title.Font = New-Object Drawing.Font("Segoe UI", 14, [Drawing.FontStyle]::Bold)
$title.Location = New-Object Drawing.Point(20, 15)
$title.AutoSize = $true
$header.Controls.Add($title)

############################################################
# STATUS BOX
############################################################

$statusBox = New-Object Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$statusBox.Size = New-Object Drawing.Size(650, 200)
$statusBox.Location = New-Object Drawing.Point(20, 70)
$form.Controls.Add($statusBox)

function Update-Status($msg) {
    $statusBox.AppendText("$msg`r`n")
    $statusBox.SelectionStart = $statusBox.Text.Length
    $statusBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

############################################################
# PROGRESS BAR
############################################################

$progressBar = New-Object Windows.Forms.ProgressBar
$progressBar.Size = New-Object Drawing.Size(650, 18)
$progressBar.Location = New-Object Drawing.Point(20, 280)
$form.Controls.Add($progressBar)


$y = 310   # starting Y after progress bar


############################################################
# WORKSPOT SECTION
############################################################

$wsLabel = New-Object Windows.Forms.Label
$wsLabel.Text = "Workspot Client Download URL"
$wsLabel.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$wsLabel.Location = New-Object Drawing.Point(20, $y)
$wsLabel.AutoSize = $true
$form.Controls.Add($wsLabel)

$y += 25

$wsTextbox = New-Object Windows.Forms.TextBox
$wsTextbox.Size = New-Object Drawing.Size(650, 25)
$wsTextbox.Location = New-Object Drawing.Point(20, $y)
$form.Controls.Add($wsTextbox)

$y += 40


############################################################
# FABTECH SECTION
############################################################

$fLabel = New-Object Windows.Forms.Label
$fLabel.Text = "FabTech Setup (Run PS As an Admin)"
$fLabel.Font = New-Object Drawing.Font("Segoe UI", 10, [Drawing.FontStyle]::Bold)
$fLabel.Location = New-Object Drawing.Point(20, $y)
$fLabel.AutoSize = $true
$form.Controls.Add($fLabel)

$y += 30

# FabTech Client
$fClientTextbox = New-Object Windows.Forms.TextBox
$fClientTextbox.Size = New-Object Drawing.Size(650, 25)
$fClientTextbox.Location = New-Object Drawing.Point(20, $y)
$fClientTextbox.Text = "FabTech Client Download Link"
$fClientTextbox.ForeColor = [Drawing.Color]::Gray

$fClientTextbox.Add_GotFocus({
    if ($fClientTextbox.Text -eq "FabTech Client Download Link") {
        $fClientTextbox.Text = ""
        $fClientTextbox.ForeColor = [Drawing.Color]::Black
    }
})

$fClientTextbox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($fClientTextbox.Text)) {
        $fClientTextbox.Text = "FabTech Client Download Link"
        $fClientTextbox.ForeColor = [Drawing.Color]::Gray
    }
})
$form.Controls.Add($fClientTextbox)

$y += 35

# FabTech Server
$fServerTextbox = New-Object Windows.Forms.TextBox
$fServerTextbox.Size = New-Object Drawing.Size(650, 25)
$fServerTextbox.Location = New-Object Drawing.Point(20, $y)
$fServerTextbox.Text = "FabTech Server Download Link"
$fServerTextbox.ForeColor = [Drawing.Color]::Gray

$fServerTextbox.Add_GotFocus({
    if ($fServerTextbox.Text -eq "FabTech Server Download Link") {
        $fServerTextbox.Text = ""
        $fServerTextbox.ForeColor = [Drawing.Color]::Black
    }
})

$fServerTextbox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($fServerTextbox.Text)) {
        $fServerTextbox.Text = "FabTech Client Download URL"
        $fServerTextbox.ForeColor = [Drawing.Color]::Gray
    }
})
$form.Controls.Add($fServerTextbox)

$y += 40

# FabTech License Server
$fLicenseTextbox = New-Object Windows.Forms.TextBox
$fLicenseTextbox.Size = New-Object Drawing.Size(650, 25)
$fLicenseTextbox.Location = New-Object Drawing.Point(20, $y)
$fLicenseTextbox.Text = "FabTech License Server (e.g. licenseserver.company.com DNS/FQDN, Or IP)"
$fLicenseTextbox.ForeColor = [Drawing.Color]::Gray

$fLicenseTextbox.Add_GotFocus({
    if ($fLicenseTextbox.Text -eq "FabTech License Server (e.g. licenseserver.company.com DNS/FQDN, Or IP)") {
        $fLicenseTextbox.Text = ""
        $fLicenseTextbox.ForeColor = [Drawing.Color]::Black
    }
})

$fLicenseTextbox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($fLicenseTextbox.Text)) {
        $fLicenseTextbox.Text = "FabTech Client Download URL"
        $fLicenseTextbox.ForeColor = [Drawing.Color]::Gray
    }
})
$form.Controls.Add($fLicenseTextbox)

$y += 45

############################################################
# CONNECTIVITY
############################################################

$portalTextbox = New-Object Windows.Forms.TextBox
$portalTextbox.Size = New-Object Drawing.Size(300, 25)
$portalTextbox.Location = New-Object Drawing.Point(20, $y)
$portalTextbox.Text = "control.workspot.com"
$form.Controls.Add($portalTextbox)

$gatewayTextbox = New-Object Windows.Forms.TextBox
$gatewayTextbox.Size = New-Object Drawing.Size(300, 25)
$gatewayTextbox.Location = New-Object Drawing.Point(370, $y)
$form.Controls.Add($gatewayTextbox)

$y += 50

############################################################
# BUTTON CREATOR
############################################################

function New-Button($text, $x, $y) {
    $btn = New-Object Windows.Forms.Button
    $btn.Text = $text
    $btn.Size = New-Object Drawing.Size(200, 40)
    $btn.Location = New-Object Drawing.Point($x, $y)
    $btn.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
    $form.Controls.Add($btn)
    return $btn
}

############################################################
# AUTO-LAYOUT BUTTONS (NO MANUAL Y NEEDED)
############################################################

function Get-NextY($control, $padding = 20) {
    return $control.Location.Y + $control.Height + $padding
}

# Anchor from the LAST control in FabTech / Connectivity section
$baseY = Get-NextY $gatewayTextbox 30

############################################################
# ROW 1
############################################################

$btnDetect = New-Button "Detect Client"     20  $baseY
$btnUninstall = New-Button "Uninstall Workspot Client"         240 $baseY
$btnCleanup = New-Button "Deep Cleanup"      460 $baseY

############################################################
# ROW 2
############################################################

$row2Y = Get-NextY $btnDetect 10

$btnInstall = New-Button "Install Workspot Client"  20  $row2Y
$btnLogs = New-Button "Collect Logs"      240 $row2Y
$btnTest = New-Button "Test Connectivity" 460 $row2Y

############################################################
# ROW 3 (FABTECH)
############################################################

$row3Y = Get-NextY $btnInstall 10

$btnFabClient = New-Button "Install FabTech Client" 20 $row3Y
$btnFabServer = New-Button "Install FabTech Server" 240 $row3Y
$btnDebugLogs = New-Button "Enable Client Debug Logs" 460 $row3Y

############################################################
# DETECT INSTALLED APPLICATION (Workspot Client)
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
# DOWNLOAD + INSTALL (Workspot Client)
############################################################

$btnInstall.Add_Click({

        $url = $wsTextbox.Text.Trim()

        if (!$url) {
            Update-Status "Please paste Workspot MSI URL in the Client Download Link field below the status box."
            return
        }
        $btnInstall.Enabled = $true
        New-Item C:\Temp -ItemType Directory -Force | Out-Null
        $dest = "C:\Temp\WorkspotClient.msi"

        Update-Status "Starting download..."

        $request = [System.Net.HttpWebRequest]::Create($url)
        $response = $request.GetResponse()
        $total = $response.ContentLength
        $stream = $response.GetResponseStream()

        $file = [IO.File]::Create($dest)

        $buffer = New-Object byte[] 8192
        $read = 0
        $totalRead = 0
        $start = Get-Date

        while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {

            $file.Write($buffer, 0, $read)
            $totalRead += $read

            $percent = [math]::Round(($totalRead / $total) * 100)

            $elapsed = (Get-Date) - $start
            $speed = ($totalRead / 1MB) / $elapsed.TotalSeconds
            $remain = ($total - $totalRead) / 1MB
            $eta = $remain / $speed

            $progressBar.Value = $percent

            $status = "Download $percent%  {0:N1}/{1:N1} MB  Speed {2:N2} MB/s  ETA {3:N0}s" -f `
            ($totalRead / 1MB), ($total / 1MB), $speed, $eta

            $statusBox.Lines[-1] = $status
            [Windows.Forms.Application]::DoEvents()

        }

        $file.Close()
        $response.Close()

        Update-Status "Download complete"

        ############################################################
        # INSTALL
        ############################################################

        Update-Status "Starting installation..."

        $progressBar.Style = "Marquee"

        Start-Process msiexec.exe -ArgumentList "/i `"$dest`" /qn /norestart" -Wait

        $progressBar.Style = "Blocks"
        $progressBar.Value = 100

        Update-Status "✓ Workspot Client installed successfully"

        $btnInstall.Enabled = $true

    })
############################################################
# FABTECH BUTTON EVENTS
############################################################

############################################################
# GENERIC DOWNLOAD + INSTALL FUNCTION (REUSABLE)
############################################################

function Install-MSI {
    param(
        [string]$Url,
        [string]$Name,
        [string]$ExtraArgs
    )

    try {
        if ([string]::IsNullOrWhiteSpace($Url)) {
            Update-Status "❌ $Name URL is empty."
            return
        }

        New-Item C:\Temp -ItemType Directory -Force | Out-Null
        $dest = "C:\Temp\$Name.msi"
        $log = "C:\Temp\$Name-install.log"

        Update-Status "`n[$Name] Download started..."

        ############################################################
        # DOWNLOAD WITH PROGRESS
        ############################################################

        $request = [System.Net.HttpWebRequest]::Create($Url)
        $response = $request.GetResponse()
        $total = $response.ContentLength
        $stream = $response.GetResponseStream()

        $file = [IO.File]::Create($dest)

        $buffer = New-Object byte[] 8192
        $totalRead = 0
        $start = Get-Date

        while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {

            $file.Write($buffer, 0, $read)
            $totalRead += $read

            if ($total -gt 0) {
                $percent = [math]::Round(($totalRead / $total) * 100)
                $progressBar.Value = $percent
            }

            [Windows.Forms.Application]::DoEvents()
        }

        $file.Close()
        $response.Close()

        Update-Status "[$Name] Download complete"

        ############################################################
        # INSTALL
        ############################################################

        Update-Status "[$Name] Installing..."

        $progressBar.Style = "Marquee"

        $arguments = "/i `"$dest`" /qn /norestart /l*v `"$log`""

        if ($ExtraArgs) {
            $arguments += " $ExtraArgs"
        }

        $process = Start-Process msiexec.exe -ArgumentList $arguments -Wait -PassThru

        $progressBar.Style = "Blocks"

        if ($process.ExitCode -eq 0) {
            Update-Status "✅ $Name installed successfully"
        }
        else {
            Update-Status "❌ $Name installation failed. ExitCode: $($process.ExitCode)"
            Update-Status "Check log: $log"
        }

    }
    catch {
        Update-Status "❌ Error installing $Name : $_"
    }
}

############################################################
# FABTECH CLIENT INSTALL
############################################################

$btnFabClient.Add_Click({

        $btnFabClient.Enabled = $false

        $url = $fClientTextbox.Text.Trim()

        if (!$url) {
            Update-Status "❌ Please provide FabTech Client URL"
            $btnFabClient.Enabled = $true
            return
        }

        Install-MSI -Url $url -Name "FabTechClient"

        $btnFabClient.Enabled = $true
    })

############################################################
# FABTECH SERVER INSTALL (WITH LICENSE LOGIC)
############################################################

$btnFabServer.Add_Click({

        $btnFabServer.Enabled = $false

        $url = $fServerTextbox.Text.Trim()
        $license = $fLicenseTextbox.Text.Trim()

        if (!$url) {
            Update-Status "❌ Please provide FabTech Server URL"
            $btnFabServer.Enabled = $true
            return
        }

        ############################################################
        # LICENSE SERVER LOGIC
        ############################################################

        $extraArgs = ""

        if ($license) {

            Update-Status "Applying FabTech License configuration..."

            $extraArgs = @(
                "LICENSESERVER=$license"
                "PORT=33033"
            ) -join " "
        }
        else {
            Update-Status "⚠ No license server provided. Installing without license config."
        }

        Install-MSI -Url $url -Name "FabTechServer" -ExtraArgs $extraArgs

        $btnFabServer.Enabled = $true
    })

$btnTest.Add_Click({

        Update-Status "`nTesting connectivity..."

        foreach ($s in @($portalTextbox.Text, $gatewayTextbox.Text)) {

            if (!$s) { continue }

            $r = Test-NetConnection $s -Port 443 -WarningAction SilentlyContinue

            if ($r.TcpTestSucceeded) {
                Update-Status "✅ $s reachable"
            }
            else {
                Update-Status "❌ $s NOT reachable"
            }
        }
    })

$btnLogs.Add_Click({

        $zip = "C:\Temp\WorkspotLogs.zip"

        Compress-Archive "$env:LOCALAPPDATA\Workspot\Client\log" $zip -Force

        Update-Status "Logs saved: $zip"
    })

$btnCleanup.Add_Click({

        Update-Status "Running cleanup..."

        @(
            "$env:ProgramFiles\Workspot",
            "$env:LOCALAPPDATA\Workspot"
        ) | ForEach-Object {

            if (Test-Path $_) {
                Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue
                Update-Status "Removed $_"
            }
        }

        Update-Status "Cleanup complete"
    })

$btnDetect.Add_Click({

        $app = Get-ItemProperty `
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" `
            -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Workspot client*" }

        if ($app) {
            Update-Status "Installed: $($app.DisplayName) - $($app.DisplayVersion)"
        }
        else {
            Update-Status "Workspot not installed"
        }
    })

$btnUninstall.Add_Click({

        Update-Status "`nSearching for Workspot installation..."

        $paths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )

        $app = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -eq "Workspot Client" } |
        Select-Object -First 1

        if (!$app) {

            Update-Status "Workspot Client not found."
            return
        }

        Update-Status "Workspot Client found: $($app.DisplayName)"

        $uninstallString = $app.UninstallString

        if ([string]::IsNullOrWhiteSpace($uninstallString)) {

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

        $guidMatch = [regex]::Match($uninstallString, "\{[A-F0-9\-]+\}", "IgnoreCase")

        if ($guidMatch.Success) {

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

        Update-Status "Uninstall command completed"

    })

$btnDebugLogs.Add_Click({
        function Enable-DebugLogs {
            # Ensure destination folder exists
            $script:LogPath = "C:\Temp"
            # Full .reg file path under the data folder
            $script:WSDebugLogging = Join-Path $script:LogPath "WS_Advance_logging_v2.reg"
            if (-not (Test-Path $script:LogPath)) {
                New-Item -Path $script:LogPath -ItemType Directory -Force | Out-Null
            }

            $GitHubUrl = "https://raw.githubusercontent.com/sathishp-beep/WS_Advance_logging_v2/main/WS_Advance_logging_v2.reg"

            # Always (re)download to be sure we have a valid .reg file
            try {
                Write-Host "Downloading debug logging registry file..." -ForegroundColor Yellow
                Invoke-WebRequest -Uri $GitHubUrl -OutFile $script:WSDebugLogging -UseBasicParsing -ErrorAction Stop
                Write-Host "Downloaded to: $script:WSDebugLogging" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to download .reg file: $($_.Exception.Message)"
                return
            }

            if (-not (Test-Path $script:WSDebugLogging)) {
                Write-Warning "Debug logging file not found after download: $script:WSDebugLogging"
                return
            }

            # Import using reg.exe for reliability
            try {
                $quotedPath = '"' + $script:WSDebugLogging + '"'
                Write-Host "Importing registry from: $quotedPath" -ForegroundColor Yellow

                $proc = Start-Process -FilePath "$env:WINDIR\System32\reg.exe" `
                    -ArgumentList "import $quotedPath" `
                    -Wait -PassThru -NoNewWindow

                if ($proc.ExitCode -eq 0) {
                    Write-Host "Debug logging registry imported successfully." -ForegroundColor Green
                }
                else {
                    Write-Error "Registry import failed. Exit code: $($proc.ExitCode)"
                }
            }
            catch {
                Write-Error "Failed to import .reg file: $($_.Exception.Message)"
            }
        }

        Enable-DebugLogs | Out-Null
        Update-Status "✓ Client debug logging enabled. Check registry for details."
    })
$form.ShowDialog()
