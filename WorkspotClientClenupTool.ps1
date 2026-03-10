# Workspot Client Uninstall & Cleanup GUI Tool

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Workspot Client Cleanup Tool"
$form.Size = New-Object System.Drawing.Size(520, 350)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Arial", 10)

# Title
$label = New-Object System.Windows.Forms.Label
$label.Text = "Workspot Client Uninstall & Cleanup"
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

# Button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Uninstall & Cleanup"
$button.Size = New-Object System.Drawing.Size(200, 35)
$button.Location = New-Object System.Drawing.Point(150, 270)
$button.BackColor = [System.Drawing.Color]::LightCoral
$form.Controls.Add($button)

$button.Add_Click({

        $button.Enabled = $false
        Update-Status "Starting Workspot uninstall & cleanup..."

        New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null

        # STEP 0 Stop processes
        Update-Status "`n[Step 0] Stopping Workspot processes..."

        Get-Process "*workspot*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

        Update-Status "✓ Processes stopped"


        # STEP 1 Detect Workspot
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

            Update-Status "✓ Found: $($app.DisplayName)"

            $uninstallString = $app.UninstallString

            if ($uninstallString -match "MsiExec") {

                Update-Status "Detected MSI uninstall"

                $productCode = $uninstallString -replace '.*(\{.*\}).*', '$1'

                $arguments = "/x $productCode /qn /norestart /l*v C:\Temp\Workspot_Uninstall.log"

                Update-Status "Running: msiexec $arguments"

                Start-Process "msiexec.exe" -ArgumentList $arguments -Wait

                Update-Status "✓ MSI uninstall completed"

            }
            else {

                Update-Status "Detected EXE uninstall"

                $exe = $uninstallString -replace '"', ''

                Start-Process "cmd.exe" -ArgumentList "/c $exe" -Wait

                Update-Status "✓ EXE uninstall executed"

            }

        }

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

        $button.Enabled = $true

    })

$form.ShowDialog() | Out-Null