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
#$fClientTextbox.PlaceholderText = "FabTech Client Download Link"

$fClientTextbox.Text = "FabTech Client Download URL"
$fClientTextbox.ForeColor = [Drawing.Color]::Gray

$fClientTextbox.Add_GotFocus({
    if ($fClientTextbox.Text -eq "FabTech Client Download URL") {
        $fClientTextbox.Text = ""
        $fClientTextbox.ForeColor = [Drawing.Color]::Black
    }
})

$fClientTextbox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($fClientTextbox.Text)) {
        $fClientTextbox.Text = "FabTech Client Download URL"
        $fClientTextbox.ForeColor = [Drawing.Color]::Gray
    }
})

$form.Controls.Add($fClientTextbox)

$y += 35

# FabTech Server
$fServerTextbox = New-Object Windows.Forms.TextBox
$fServerTextbox.Size = New-Object Drawing.Size(650, 25)
$fServerTextbox.Location = New-Object Drawing.Point(20, $y)
$fServerTextbox.PlaceholderText = "FabTech Server Download Link"
$form.Controls.Add($fServerTextbox)

$y += 40

# FabTech License Server
$fLicenseTextbox = New-Object Windows.Forms.TextBox
$fLicenseTextbox.Size = New-Object Drawing.Size(650, 25)
$fLicenseTextbox.Location = New-Object Drawing.Point(20, $y)
$fLicenseTextbox.PlaceholderText = "FabTech License Server (e.g. licenseserver.company.com DNS/FQDN, Or IP)"
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
$btnUninstall = New-Button "Uninstall"         240 $baseY
$btnCleanup = New-Button "Deep Cleanup"      460 $baseY

############################################################
# ROW 2
############################################################

$row2Y = Get-NextY $btnDetect 10

$btnInstall = New-Button "Install Workspot"  20  $row2Y
$btnLogs = New-Button "Collect Logs"      240 $row2Y
$btnTest = New-Button "Test Connectivity" 460 $row2Y

############################################################
# ROW 3 (FABTECH)
############################################################

$row3Y = Get-NextY $btnInstall 10

$btnFabClient = New-Button "Install FabTech Client" 120 $row3Y
$btnFabServer = New-Button "Install FabTech Server" 360 $row3Y

<#
############################################################
# BUTTONS (SHIFTED DOWN AUTOMATICALLY)
############################################################

$btnDetect = New-Button "Detect Client"        20  $y
$btnUninstall = New-Button "Uninstall Workspot Client"            240 $y
$btnCleanup = New-Button "Deep Cleanup"         460 $y

$y += 60

$btnInstall = New-Button "Install Workspot Client"     20  $y
$btnLogs = New-Button "Collect Logs"         240 $y
$btnTest = New-Button "Test Connectivity"    460 $y

$y += 70

$btnFabClient = New-Button "Install FabTech Client" 120 $y
$btnFabServer = New-Button "Install FabTech Server" 360 $y

$y += 80

############################################################
# DOWNLOAD + INSTALL FUNCTION (REUSABLE)
############################################################

function Install-MSI($url, $name) {

    if ([string]::IsNullOrWhiteSpace($url)) {
        Update-Status "❌ URL missing for $name"
        return
    }

    $dest = "C:\Temp\$name.msi"
    New-Item C:\Temp -ItemType Directory -Force | Out-Null

    Update-Status "`nDownloading $name..."

    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing

        Update-Status "Download complete: $dest"
        Update-Status "Installing $name..."

        $progressBar.Style = "Marquee"

        Start-Process msiexec.exe -ArgumentList "/i `"$dest`" /qn /norestart" -Wait

        $progressBar.Style = "Blocks"
        $progressBar.Value = 100

        Update-Status "✅ $name installed successfully"
    }
    catch {
        Update-Status "❌ Failed: $($_.Exception.Message)"
    }
}
#>

############################################################
# DETECT INSTALLED APPLICATION
############################################################

function Get-InstalledApp($name) {

    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    return Get-ItemProperty $paths -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "*$name*" } |
    Select-Object -First 1
}

############################################################
# DOWNLOAD + INSTALL (SMART)
############################################################

function Install-MSI($url, $name) {

    if ([string]::IsNullOrWhiteSpace($url)) {
        Update-Status "❌ URL missing for $name"
        return
    }

    ########################################################
    # PRE-CHECK (ALREADY INSTALLED)
    ########################################################

    $existing = Get-InstalledApp $name

    if ($existing) {
        Update-Status "⚠ $name already installed"
        Update-Status "Version: $($existing.DisplayVersion)"
        return
    }

    ########################################################
    # DOWNLOAD
    ########################################################

    $dest = "C:\Temp\$name.msi"
    $log = "C:\Temp\$name-Install.log"

    New-Item C:\Temp -ItemType Directory -Force | Out-Null

    Update-Status "`nDownloading $name..."

    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
        Update-Status "Download complete: $dest"
    }
    catch {
        Update-Status "❌ Download failed: $($_.Exception.Message)"
        return
    }

    ########################################################
    # BUILD INSTALL ARGUMENTS
    ########################################################

    $arguments = @(
        "/i `"$dest`""
        "/qn"
        "/norestart"
        "/l*v `"$log`""
    )

    ########################################################
    # FABTECH SPECIAL LOGIC
    ########################################################

    if ($name -like "*FabTech*") {

        Update-Status "Applying FabTech configuration..."
        # Read value from textbox
        $licenseServer = $fLicenseTextbox.Text.Trim()

        $arguments += @(
            "LICENSESERVER=$licenseServer"
            "PORT=33033"
        )

        # Optional: Server-specific install path
        if ($name -like "*Server*") {
            $arguments += 'INSTALLDIR="C:\Program Files\FabTech\Server"'
        }
    }

    ########################################################
    # INSTALL
    ########################################################

    Update-Status "Installing $name..."
    $progressBar.Style = "Marquee"

    try {
        Start-Process "msiexec.exe" `
            -ArgumentList $arguments `
            -Wait -NoNewWindow

        $progressBar.Style = "Blocks"
        $progressBar.Value = 100

        ####################################################
        # POST-INSTALL VALIDATION
        ####################################################

        Start-Sleep -Seconds 3

        $installed = Get-InstalledApp $name

        if ($installed) {
            Update-Status "✅ $name installed successfully"
            Update-Status "Version: $($installed.DisplayVersion)"
        }
        else {
            Update-Status "⚠ Installation completed but not detected"
            Update-Status "Check log: $log"
        }
    }
    catch {
        Update-Status "❌ Installation failed: $($_.Exception.Message)"
        Update-Status "Check log: $log"
    }
}
############################################################
# BUTTON EVENTS
############################################################

$btnInstall.Add_Click({
        Install-MSI $wsTextbox.Text.Trim() "WorkspotClient"
    })

$btnFabClient.Add_Click({
        Install-MSI $fClientTextbox.Text.Trim() "FabTechClient"
    })

$btnFabServer.Add_Click({
        Install-MSI $fServerTextbox.Text.Trim() "FabTechServer"
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

$form.ShowDialog()
