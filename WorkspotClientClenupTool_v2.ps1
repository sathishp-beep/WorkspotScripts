# Workspot Client Uninstall & Cleanup GUI Tool

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Workspot Client Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(520, 430)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Icon = New-Object System.Drawing.Icon("https://raw.githubusercontent.com/sathishp-beep/WorkspotScripts/refs/heads/main/workspot.ico")

# Title
$label = New-Object System.Windows.Forms.Label
$label.Text = "Workspot Client Uninstall / Reinstall Tool"
$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(460, 30)
$label.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($label)

# Status Box
$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Location = New-Object System.Drawing.Point(20, 60)
$statusBox.Size = New-Object System.Drawing.Size(460, 200)
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

# --------------------------------------------------
# MSI URL INPUT
# --------------------------------------------------

$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Text = "Workspot MSI Download URL:"
$urlLabel.Location = New-Object System.Drawing.Point(20, 270)
$urlLabel.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($urlLabel)

$urlTextbox = New-Object System.Windows.Forms.TextBox
$urlTextbox.Location = New-Object System.Drawing.Point(20, 295)
$urlTextbox.Size = New-Object System.Drawing.Size(460, 25)
$form.Controls.Add($urlTextbox)

# --------------------------------------------------
# Uninstall Button
# --------------------------------------------------

$buttonUninstall = New-Object System.Windows.Forms.Button
$buttonUninstall.Text = "Uninstall & Cleanup"
$buttonUninstall.Size = New-Object System.Drawing.Size(200, 35)
$buttonUninstall.Location = New-Object System.Drawing.Point(60, 340)
$buttonUninstall.BackColor = [System.Drawing.Color]::LightCoral
$form.Controls.Add($buttonUninstall)

# --------------------------------------------------
# Reinstall Button
# --------------------------------------------------

$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Re-Install Workspot Client"
$buttonInstall.Size = New-Object System.Drawing.Size(200, 35)
$buttonInstall.Location = New-Object System.Drawing.Point(260, 340)
$buttonInstall.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($buttonInstall)

# =====================================================
# UNINSTALL LOGIC
# =====================================================

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

        if (!$app) {

            Update-Status "⚠ Workspot Client not found"

        }
        else {

            $uninstallString = $app.UninstallString

            if ($uninstallString -match "MsiExec") {

                $productCode = $uninstallString -replace '.*(\{.*\}).*', '$1'
                $arguments = "/x $productCode /qn /norestart /l*v C:\Temp\Workspot_Uninstall.log"

                Start-Process "msiexec.exe" -ArgumentList $arguments -Wait

                Update-Status "✓ Workspot Client uninstalled"

            }
            else {

                $exe = $uninstallString -replace '"', ''
                Start-Process "cmd.exe" -ArgumentList "/c $exe" -Wait

                Update-Status "✓ EXE uninstall executed"

            }

        }

        Update-Status "`nCleanup completed."

    })

# =====================================================
# REINSTALL LOGIC
# =====================================================

$buttonInstall.Add_Click({

        $url = $urlTextbox.Text

        if ([string]::IsNullOrWhiteSpace($url)) {

            Update-Status "⚠ Please enter MSI download URL first."
            return

        }

        New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null

        $msiPath = "C:\Temp\WorkspotClient.msi"

        try {

            Update-Status "`nDownloading MSI..."

            Invoke-WebRequest $url -OutFile $msiPath

            Update-Status "✓ Download completed"

            Update-Status "Installing Workspot Client silently..."

            $arguments = "/i `"$msiPath`" /qn /norestart /l*v C:\Temp\Workspot_Install.log"

            Start-Process "msiexec.exe" -ArgumentList $arguments -Wait

            Update-Status "✓ Workspot Client installed successfully"

        }
        catch {

            Update-Status "✗ Installation failed: $_"

        }

    })


$form.ShowDialog() | Out-Null
