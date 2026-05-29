<#
.SYNOPSIS
    GUI tool for building and running GCP instance-template gcloud commands with tags and labels.

.DESCRIPTION
    Provides a UI to fill in gcloud compute instance-templates create parameters,
    including tags and labels, and generates or executes the command directly.

.NOTES
    Run from Windows PowerShell with STA enabled if your host requires it:
    powershell.exe -NoProfile -ExecutionPolicy Bypass -STAh -File ".\Invoke-WsTemplateOperationsGui.ps1"
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latesth

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $powerShellExe = (Get-Command powershell.exe -ErrorAction SilentlyContinue).Source
    if ($powerShellExe -and $PSCommandPath) {   # ← add: -and $PSCommandPath
        Start-Process -FilePath $powerShellExe -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA',
            '-File', "`"$PSCommandPath`""
        ) -WindowStyle Normal
        return
    }
}

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function New-WsLabel {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 130
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, 20)
    $label.TextAlign = 'MiddleLeft'
    $label
}

function New-WsTextBox {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width = 260,
        [string]$Text = '',
        [switch]$Password,
        [switch]$Multiline
    )

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point($X, $Y)
    $textBox.Size = New-Object System.Drawing.Size($Width, $(if ($Multiline) { 80 } else { 22 }))
    $textBox.Text = $Text
    $textBox.Anchor = 'Top, Left, Right'
    if ($Password) {
        $textBox.UseSystemPasswordChar = $true
    }
    if ($Multiline) {
        $textBox.Multiline = $true
        $textBox.ScrollBars = 'Vertical'
        $textBox.AcceptsReturn = $true
        $textBox.AcceptsTab = $true
    }
    $textBox
}

function New-WsButton {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 140
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, 30)
    $button
}

function Write-WsConsole {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"
    $script:OutputBox.AppendText($line + [Environment]::NewLine)
}

function Format-WsJson {
    param([object]$InputObject)

    if ($null -eq $InputObject) {
        return ''
    }

    try {
        return ($InputObject | ConvertTo-Json -Depth 30)
    }
    catch {
        return [string]$InputObject
    }
}

function ConvertFrom-WsCsvList {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    $items = @()
    foreach ($part in ($Text -split ',')) {
        $value = $part.Trim()
        if ($value) {
            $items += $value
        }
    }
    @($items)
}

function ConvertFrom-WsKeyValueList {
    param([string]$Text)

    $hash = @{}
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $hash
    }

    # Accept newline-separated (multiline textbox) and/or comma-separated entries
    foreach ($part in ($Text -split '[,\r\n]+')) {
        $item = $part.Trim()
        if (-not $item) {
            continue
        }

        $pieces = @($item -split '=', 2)
        if ($pieces.Count -ne 2 -or [string]::IsNullOrWhiteSpace($pieces[0])) {
            throw "Labels must use key=value format. Invalid entry: '$item'"
        }

        $hash[$pieces[0].Trim()] = $pieces[1].Trim()
    }

    $hash
}

function ConvertTo-WsGcloudArgument {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    $Value
}

function Invoke-WsGuiAction {
    param(
        [string]$ActionName,
        [scriptblock]$Action
    )

    try {
        Write-WsConsole "Starting $ActionName..."
        & $Action | Out-Null
        Write-WsConsole "$ActionName completed."
    }
    catch {
        Write-WsConsole $_.Exception.Message 'ERROR'
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $ActionName, 'OK', 'Error') | Out-Null
    }
}

function Update-WsGcloudCommand {
    $templateName = $script:GcloudTemplateNameTextBox.Text.Trim()
    $project = $script:GcloudProjectTextBox.Text.Trim()
    $machineType = $script:GcloudMachineTypeTextBox.Text.Trim()
    $network = $script:GcloudNetworkTextBox.Text.Trim()
    $subnet = $script:GcloudSubnetTextBox.Text.Trim()
    $maintenancePolicy = $script:GcloudMaintenanceTextBox.Text.Trim()
    $regionZone = $script:GcloudRegionTextBox.Text.Trim()   # zone-level  e.g. us-east4-a  (1st --region)
    $region = $script:GcloudZoneTextBox.Text.Trim()         # region-level e.g. us-east4   (2nd --region)
    $acceleratorType = $script:GcloudAcceleratorTypeComboBox.Text.Trim()
    $acceleratorCount = if ($null -ne $script:GcloudAcceleratorCountUpDown) { $script:GcloudAcceleratorCountUpDown.Value.ToString() } else { '' }
    $deviceName = $script:GcloudDeviceNameTextBox.Text.Trim()
    $image = $script:GcloudImageTextBox.Text.Trim()
    $diskSize = $script:GcloudDiskSizeTextBox.Text.Trim()
    $diskType = $script:GcloudDiskTypeComboBox.Text.Trim()
    $metadataHash = ConvertFrom-WsKeyValueList -Text $script:GcloudMetadataTextBox.Text
    $tags = (ConvertFrom-WsCsvList -Text $script:TagsTextBox.Text) -join ','
    $labelPairs  = [System.Collections.Generic.List[string]]::new()
    $labelKeysSeen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $duplicateKeys = [System.Collections.Generic.List[string]]::new()
    foreach ($part in ($script:LabelsTextBox.Text -split '[,\r\n]+')) {
        $item = $part.Trim()
        if (-not $item) { continue }
        $pieces = @($item -split '=', 2)
        if ($pieces.Count -ne 2 -or [string]::IsNullOrWhiteSpace($pieces[0])) {
            throw "Labels must use key=value format. Invalid entry: '$item'"
        }
        $k = $pieces[0].Trim()
        $v = $pieces[1].Trim()
        if (-not $labelKeysSeen.Add($k)) {
            if (-not $duplicateKeys.Contains($k)) { $duplicateKeys.Add($k) }
        }
        $labelPairs.Add("$k=$v")
    }
    if ($duplicateKeys.Count -gt 0) {
        $dupList = $duplicateKeys -join ', '
        throw "Duplicate label key(s) detected: $dupList`n`nGCP labels require UNIQUE keys per resource — only the last value per key would be applied by GCP.`n`nPlease use distinct keys, e.g.:`n  costtag1=description`n  costtag2=description`n  costtag3=description"
    }
    $labels = $labelPairs -join ','
    $metadata = (($metadataHash.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ',')

    if ([string]::IsNullOrWhiteSpace($templateName)) {
        throw 'Gcloud template name is required.'
    }
    if ([string]::IsNullOrWhiteSpace($project)) {
        throw 'Gcloud project is required.'
    }

    $arguments = New-Object System.Collections.Generic.List[string]
    $arguments.Add('gcloud')
    $arguments.Add('compute')
    $arguments.Add('instance-templates')
    $arguments.Add('create')
    $arguments.Add((ConvertTo-WsGcloudArgument -Value $templateName))
    $arguments.Add("--project=$(ConvertTo-WsGcloudArgument -Value $project)")

    if ($machineType) {
        $arguments.Add("--machine-type=$(ConvertTo-WsGcloudArgument -Value $machineType)")
    }

    # 1st --region: zone-level value (e.g. us-east4-a), placed before --network-interface
    if ($regionZone) {
        $arguments.Add("--region=$(ConvertTo-WsGcloudArgument -Value $regionZone)")
    }

    $networkInterfaceParts = New-Object System.Collections.Generic.List[string]
    if ($network) {
        $networkInterfaceParts.Add("network=$network")
    }
    if ($subnet) {
        $networkInterfaceParts.Add("subnet=$subnet")
    }
    if ($networkInterfaceParts.Count -gt 0) {
        $arguments.Add("--network-interface=$(ConvertTo-WsGcloudArgument -Value ($networkInterfaceParts -join ','))")
    }
    if ($script:GcloudNoAddressCheckBox.Checked) {
        $arguments.Add('--no-address')
    }
    if ($script:GcloudCanIpForwardCheckBox.Checked) {
        $arguments.Add('--can-ip-forward')
    }
    if ($maintenancePolicy) {
        $arguments.Add("--maintenance-policy=$(ConvertTo-WsGcloudArgument -Value $maintenancePolicy)")
    }

    # 2nd --region: region-level value (e.g. us-east4), placed after --maintenance-policy
    if ($region) {
        $arguments.Add("--region=$(ConvertTo-WsGcloudArgument -Value $region)")
    }

    if ($acceleratorType) {
        $acceleratorValue = "type=$acceleratorType"
        if ($acceleratorCount) {
            $acceleratorValue = "$acceleratorValue,count=$acceleratorCount"
        }
        $arguments.Add("--accelerator=$(ConvertTo-WsGcloudArgument -Value $acceleratorValue)")
    }
    if ($metadata) {
        $arguments.Add("--metadata=$(ConvertTo-WsGcloudArgument -Value $metadata)")
    }

    $diskParts = New-Object System.Collections.Generic.List[string]
    $diskParts.Add('auto-delete=yes')
    $diskParts.Add('boot=yes')
    if ($deviceName) {
        $diskParts.Add("device-name=$deviceName")
    }
    if ($image) {
        $diskParts.Add("image=$image")
    }
    $diskParts.Add('mode=rw')
    if ($diskSize) {
        $diskParts.Add("size=$diskSize")
    }
    if ($diskType) {
        $diskParts.Add("type=$diskType")
    }
    $arguments.Add("--create-disk=$(ConvertTo-WsGcloudArgument -Value ($diskParts -join ','))")

    if ($tags) {
        $arguments.Add("--tags=$(ConvertTo-WsGcloudArgument -Value $tags)")
    }
    if ($labels) {
        $arguments.Add("--labels=$(ConvertTo-WsGcloudArgument -Value $labels)")
    }

    $script:GcloudCommandTextBox.Text = ($arguments -join ' ')
}

function Invoke-WsProcess {
    param(
        [string]$Command,
        [string]$Label = 'Process'
    )
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo.FileName = 'cmd.exe'
    $p.StartInfo.Arguments = "/c $Command"
    $p.StartInfo.RedirectStandardOutput = $true
    $p.StartInfo.RedirectStandardError = $true
    $p.StartInfo.UseShellExecute = $false
    $p.StartInfo.CreateNoWindow = $true
    [void]$p.Start()
    $out = $p.StandardOutput.ReadToEnd()
    $err = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    if ($out) { Write-WsConsole $out.Trim() }
    if ($err) { Write-WsConsole $err.Trim() 'ERROR' }
    Write-WsConsole "$Label exit code: $($p.ExitCode)"
    return $p.ExitCode
}

function Invoke-WsGcloudCommand {
    $commandLine = $script:GcloudCommandTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($commandLine)) {
        Update-WsGcloudCommand
        $commandLine = $script:GcloudCommandTextBox.Text.Trim()
    }
    Write-WsConsole 'Running gcloud command...'
    Invoke-WsProcess -Command $commandLine -Label 'gcloud' | Out-Null
}

function Invoke-WsImageBrowse {
    $project = $script:GcloudProjectTextBox.Text.Trim()
    $templateName = $script:GcloudTemplateNameTextBox.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($project)) {
        throw 'Project is required to browse images.'
    }

    $filterArg = if ($templateName) { "--filter=name~$templateName" } else { '' }
    $cmd = "gcloud compute images list --project=$project --no-standard-images --format=value(name) $filterArg".Trim()
    Write-WsConsole "Listing images: $cmd"

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo.FileName = 'cmd.exe'
    $p.StartInfo.Arguments = "/c $cmd"
    $p.StartInfo.RedirectStandardOutput = $true
    $p.StartInfo.RedirectStandardError = $true
    $p.StartInfo.UseShellExecute = $false
    $p.StartInfo.CreateNoWindow = $true
    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    if ($stderr) { Write-WsConsole $stderr.Trim() 'ERROR' }

    $imageNames = @($stdout -split "`r?`n" | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() })

    if ($imageNames.Count -eq 0) {
        $msg = if ($templateName) { "No images found matching '$templateName' in project '$project'." } else { "No custom images found in project '$project'." }
        [System.Windows.Forms.MessageBox]::Show($msg, 'Browse Images', 'OK', 'Information') | Out-Null
        Write-WsConsole $msg
        return
    }

    Write-WsConsole "Found $($imageNames.Count) image(s)."

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Select Image — $($imageNames.Count) found (double-click or Select)"
    $dlg.Size = New-Object System.Drawing.Size(640, 440)
    $dlg.StartPosition = 'CenterParent'
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Dock = 'Fill'
    $listBox.Font = New-Object System.Drawing.Font('Consolas', 9)
    $listBox.SelectionMode = 'One'
    foreach ($img in $imageNames) { [void]$listBox.Items.Add($img) }
    if ($listBox.Items.Count -gt 0) { $listBox.SelectedIndex = 0 }
    $dlg.Controls.Add($listBox)

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Bottom'
    $panel.Height = 44
    $dlg.Controls.Add($panel)

    $applySelection = {
        if ($null -ne $listBox.SelectedItem) {
            $imgName = [string]$listBox.SelectedItem
            $proj = $script:GcloudProjectTextBox.Text.Trim()
            $script:GcloudImageTextBox.Text = "projects/$proj/global/images/$imgName"
            Write-WsConsole "Image path set to: $($script:GcloudImageTextBox.Text)"
            $dlg.Close()
        }
    }

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = 'Select'
    $okBtn.Width = 100
    $okBtn.Height = 30
    $okBtn.Top = 7
    $okBtn.Left = 16
    $okBtn.Add_Click($applySelection)
    $panel.Controls.Add($okBtn)

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = 'Cancel'
    $cancelBtn.Width = 100
    $cancelBtn.Height = 30
    $cancelBtn.Top = 7
    $cancelBtn.Left = 124
    $cancelBtn.Add_Click({ $dlg.Close() })
    $panel.Controls.Add($cancelBtn)

    $listBox.Add_DoubleClick($applySelection)
    $dlg.ShowDialog() | Out-Null
}

function Invoke-WsApplyTagsLabels {
    $templateName = $script:GcloudTemplateNameTextBox.Text.Trim()
    $project = $script:GcloudProjectTextBox.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($templateName)) { throw 'Template Name is required.' }
    if ([string]::IsNullOrWhiteSpace($project)) { throw 'Project is required.' }

    # Always rebuild so the command reflects the current form values (labels, tags, etc.)
    Update-WsGcloudCommand
    $command = $script:GcloudCommandTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($command)) {
        throw 'No gcloud command found. Build Command failed to produce output.'
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will DELETE and RECREATE the instance template:`n`n  Template : $templateName`n  Project  : $project`n`nStep 1 — Delete the existing template`nStep 2 — Recreate it with the current tags and labels`n`nThis action cannot be undone. Continue?",
        'Apply Tags and Labels',
        'YesNo',
        'Warning'
    )
    if ($confirm -ne 'Yes') {
        Write-WsConsole 'Apply Tags and Labels cancelled by user.'
        return
    }

    # Step 1: delete
    Write-WsConsole "Step 1: Deleting instance template '$templateName' in project '$project'..."
    $deleteCmd = "gcloud compute instance-templates delete $templateName --project=$project --quiet"
    $rc1 = Invoke-WsProcess -Command $deleteCmd -Label 'Delete'
    if ($rc1 -ne 0) {
        throw "Delete failed (exit code $rc1). Template was NOT deleted. Aborting."
    }
    Write-WsConsole "Step 1 complete — '$templateName' deleted."

    # Step 2: recreate
    Write-WsConsole "Step 2: Recreating instance template '$templateName'..."
    $rc2 = Invoke-WsProcess -Command $command -Label 'Recreate'
    if ($rc2 -ne 0) {
        Write-WsConsole "WARNING: '$templateName' was deleted but recreation FAILED (exit code $rc2). Manual recovery required." 'ERROR'
        throw "Template deleted but recreation failed (exit code $rc2). Manual recovery required."
    }
    Write-WsConsole "Step 2 complete — '$templateName' recreated with updated tags and labels."
}

$form = New-Object System.Windows.Forms.Form
$modifiedTime = if ($PSCommandPath) { (Get-Item -LiteralPath $PSCommandPath).LastWriteTime } else { 'iex' }
$form.Text = "Gcloud Tags and Labels (Modified: $modifiedTime)"
#$form.Text = "Gcloud Tags and Labels (Modified: $((Get-Item -LiteralPath $PSCommandPath).LastWriteTime))"
$form.Size = New-Object System.Drawing.Size(1180, 820)
$form.MinimumSize = New-Object System.Drawing.Size(980, 720)
$form.StartPosition = 'CenterScreen'

$mainSplit = New-Object System.Windows.Forms.SplitContainer
$mainSplit.Dock = 'Fill'
$mainSplit.Orientation = 'Horizontal'
$mainSplit.Panel1MinSize = 300
$mainSplit.Panel2MinSize = 100
$form.Controls.Add($mainSplit)

$form.Add_Shown({
        $targetDistance = [Math]::Min(545, [Math]::Max($mainSplit.Panel1MinSize, $mainSplit.Height - $mainSplit.Panel2MinSize - 10))
        if ($targetDistance -ge $mainSplit.Panel1MinSize -and $targetDistance -le ($mainSplit.Height - $mainSplit.Panel2MinSize)) {
            $mainSplit.SplitterDistance = $targetDistance
        }
    })

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$mainSplit.Panel1.Controls.Add($tabs)

$tabGcloud = New-Object System.Windows.Forms.TabPage
$tabGcloud.Text = 'Gcloud Tags and Labels'
$tabs.TabPages.Add($tabGcloud)
$tabGcloud.AutoScroll = $true

$gcloudGroup = New-Object System.Windows.Forms.GroupBox
$gcloudGroup.Text = 'Gcloud Instance Template Inputs'
$gcloudGroup.Dock = 'Top'
$gcloudGroup.Height = 628
$tabGcloud.Controls.Add($gcloudGroup)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Template Name' -X 12 -Y 26))
$script:GcloudTemplateNameTextBox = New-WsTextBox -X 150 -Y 24 -Width 360 -Text 'TemplateName'
$gcloudGroup.Controls.Add($script:GcloudTemplateNameTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Project' -X 12 -Y 56))
$script:GcloudProjectTextBox = New-WsTextBox -X 150 -Y 54 -Width 360 -Text 'GCP project Name'
$gcloudGroup.Controls.Add($script:GcloudProjectTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Machine Type' -X 12 -Y 86))
$script:GcloudMachineTypeTextBox = New-WsTextBox -X 150 -Y 84 -Width 360 -Text 'Instance SKU'
$gcloudGroup.Controls.Add($script:GcloudMachineTypeTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Network' -X 12 -Y 116))
$script:GcloudNetworkTextBox = New-WsTextBox -X 150 -Y 114 -Width 360 -Text 'VPC Network Name'
$gcloudGroup.Controls.Add($script:GcloudNetworkTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Subnet' -X 12 -Y 146))
$script:GcloudSubnetTextBox = New-WsTextBox -X 150 -Y 144 -Width 360 -Text 'Subnet Name'
$gcloudGroup.Controls.Add($script:GcloudSubnetTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Maintenance' -X 12 -Y 176))
$script:GcloudMaintenanceTextBox = New-WsTextBox -X 150 -Y 174 -Width 360 -Text 'TERMINATE'
$gcloudGroup.Controls.Add($script:GcloudMaintenanceTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Region (zone)' -X 12 -Y 206))
$script:GcloudRegionTextBox = New-WsTextBox -X 150 -Y 204 -Width 360 -Text ''

$gcloudGroup.Controls.Add($script:GcloudRegionTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Region' -X 12 -Y 236))
$script:GcloudZoneTextBox = New-WsTextBox -X 150 -Y 234 -Width 360 -Text ''

$gcloudGroup.Controls.Add($script:GcloudZoneTextBox)

$script:GcloudNoAddressCheckBox = New-Object System.Windows.Forms.CheckBox
$script:GcloudNoAddressCheckBox.Text = 'No External Address'
$script:GcloudNoAddressCheckBox.Location = New-Object System.Drawing.Point(550, 206)
$script:GcloudNoAddressCheckBox.Size = New-Object System.Drawing.Size(180, 24)
$script:GcloudNoAddressCheckBox.Checked = $true
$gcloudGroup.Controls.Add($script:GcloudNoAddressCheckBox)

$script:GcloudCanIpForwardCheckBox = New-Object System.Windows.Forms.CheckBox
$script:GcloudCanIpForwardCheckBox.Text = 'Can IP Forward'
$script:GcloudCanIpForwardCheckBox.Location = New-Object System.Drawing.Point(550, 236)
$script:GcloudCanIpForwardCheckBox.Size = New-Object System.Drawing.Size(180, 24)
$gcloudGroup.Controls.Add($script:GcloudCanIpForwardCheckBox)

$innerXLabel = 12
$innerXValue = 150
$innerWidth = 360

$gcloudGroup.Controls.Add((New-WsLabel -Text 'GPU Type (optional)' -X $innerXLabel -Y 266))
$script:GcloudAcceleratorTypeComboBox = New-Object System.Windows.Forms.ComboBox
$script:GcloudAcceleratorTypeComboBox.Location = New-Object System.Drawing.Point($innerXValue, 264)
$script:GcloudAcceleratorTypeComboBox.Size = New-Object System.Drawing.Size($innerWidth, 24)
$script:GcloudAcceleratorTypeComboBox.DropDownStyle = 'DropDown'
$script:GcloudAcceleratorTypeComboBox.Items.AddRange(@('', 'nvidia-tesla-t4-vws', 'nvidia-tesla-t4', 'nvidia-tesla-l4', 'nvidia-tesla-k80', 'nvidia-tesla-p100', 'nvidia-tesla-v100', 'nvidia-tesla-a100', 'nvidia-tesla-p4', 'nvidia-tesla-p40'))
$script:GcloudAcceleratorTypeComboBox.Text = ''
$gcloudGroup.Controls.Add($script:GcloudAcceleratorTypeComboBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Accelerator Count' -X $innerXLabel -Y 296))
$script:GcloudAcceleratorCountUpDown = New-Object System.Windows.Forms.NumericUpDown
$script:GcloudAcceleratorCountUpDown.Location = New-Object System.Drawing.Point($innerXValue, 294)
$script:GcloudAcceleratorCountUpDown.Size = New-Object System.Drawing.Size(120, 24)
$script:GcloudAcceleratorCountUpDown.Minimum = 0
$script:GcloudAcceleratorCountUpDown.Maximum = 8
$script:GcloudAcceleratorCountUpDown.Value = 0
$gcloudGroup.Controls.Add($script:GcloudAcceleratorCountUpDown)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Device Name' -X $innerXLabel -Y 326))
$script:GcloudDeviceNameTextBox = New-WsTextBox -X $innerXValue -Y 324 -Width $innerWidth -Text 'TemplateName'
$gcloudGroup.Controls.Add($script:GcloudDeviceNameTextBox)
$script:GcloudDeviceNameAutoSynced = $true

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Image path' -X $innerXLabel -Y 358))
$script:GcloudImageTextBox = New-WsTextBox -X $innerXValue -Y 356 -Width 260 -Text 'Image Full Path'
$script:GcloudImageTextBox.Anchor = 'Top, Left'
$gcloudGroup.Controls.Add($script:GcloudImageTextBox)
$script:BrowseImagesButton = New-WsButton -Text 'Browse Images' -X 414 -Y 354 -Width 110
$script:BrowseImagesButton.Anchor = 'Top, Left'
$gcloudGroup.Controls.Add($script:BrowseImagesButton)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Disk Size GB' -X $innerXLabel -Y 390))
$script:GcloudDiskSizeTextBox = New-WsTextBox -X $innerXValue -Y 388 -Width 120 -Text '128'
$gcloudGroup.Controls.Add($script:GcloudDiskSizeTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Disk Type' -X $innerXLabel -Y 422))
$script:GcloudDiskTypeComboBox = New-Object System.Windows.Forms.ComboBox
$script:GcloudDiskTypeComboBox.Location = New-Object System.Drawing.Point($innerXValue, 420)
$script:GcloudDiskTypeComboBox.Size = New-Object System.Drawing.Size(180, 24)
$script:GcloudDiskTypeComboBox.DropDownStyle = 'DropDown'
$script:GcloudDiskTypeComboBox.Items.AddRange(@('pd-balanced', 'pd-standard', 'pd-ssd', 'local-ssd'))
$script:GcloudDiskTypeComboBox.Text = 'pd-balanced'
$gcloudGroup.Controls.Add($script:GcloudDiskTypeComboBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Metadata' -X ($innerXLabel + 260) -Y 390))
$script:GcloudMetadataTextBox = New-WsTextBox -X ($innerXValue + 260) -Y 388 -Width 260 -Text ''
$gcloudGroup.Controls.Add($script:GcloudMetadataTextBox)

# Tags and Labels — moved here from the former Workspot tab
$gcloudGroup.Controls.Add((New-WsLabel -Text 'Tags' -X $innerXLabel -Y 458))
$script:TagsTextBox = New-WsTextBox -X $innerXValue -Y 456 -Width $innerWidth -Text 'cloudpc,rdspool'
$gcloudGroup.Controls.Add($script:TagsTextBox)

$gcloudGroup.Controls.Add((New-WsLabel -Text 'Labels' -X $innerXLabel -Y 488))
$script:LabelsTextBox = New-WsTextBox -X $innerXValue -Y 486 -Width $innerWidth -Multiline -Text "description=rds-sgcostx"
$gcloudGroup.Controls.Add($script:LabelsTextBox)

$script:GcloudAuthButton = New-WsButton -Text 'gcloud auth login' -X $innerXLabel -Y 580 -Width 140
$gcloudGroup.Controls.Add($script:GcloudAuthButton)

$script:BuildGcloudButton = New-WsButton -Text 'Build Command' -X ($innerXValue + 120) -Y 580 -Width 140
$gcloudGroup.Controls.Add($script:BuildGcloudButton)

$script:RunGcloudButton = New-WsButton -Text 'Run gcloud' -X ($innerXValue + 270) -Y 580 -Width 140
$gcloudGroup.Controls.Add($script:RunGcloudButton)

$script:PreviewGcloudButton = New-WsButton -Text 'Preview gcloud' -X ($innerXValue + 420) -Y 580 -Width 140
$gcloudGroup.Controls.Add($script:PreviewGcloudButton)

$script:ApplyTagsLabelsButton = New-WsButton -Text 'Apply Tags && Labels' -X ($innerXValue + 570) -Y 580 -Width 160
$gcloudGroup.Controls.Add($script:ApplyTagsLabelsButton)

$script:GcloudTemplateNameTextBox.Add_TextChanged({
        if ($script:GcloudDeviceNameTextBox.Text -eq '' -or $script:GcloudDeviceNameAutoSynced) {
            $script:GcloudDeviceNameTextBox.Text = $script:GcloudTemplateNameTextBox.Text
            $script:GcloudDeviceNameAutoSynced = $true
        }
    })

$script:GcloudDeviceNameTextBox.Add_TextChanged({
        if ($script:GcloudDeviceNameTextBox.Text -ne $script:GcloudTemplateNameTextBox.Text) {
            $script:GcloudDeviceNameAutoSynced = $false
        }
    })

$commandGroup = New-Object System.Windows.Forms.GroupBox
$commandGroup.Text = 'Generated gcloud Command'
$commandGroup.Dock = 'Fill'
$tabGcloud.Controls.Add($commandGroup)

$script:GcloudCommandTextBox = New-WsTextBox -X 12 -Y 24 -Width 1100 -Multiline
$script:GcloudCommandTextBox.Height = 140
$script:GcloudCommandTextBox.ScrollBars = 'Both'
$script:GcloudCommandTextBox.Anchor = 'Top, Left, Right'
$commandGroup.Controls.Add($script:GcloudCommandTextBox)

$script:OutputBox = New-Object System.Windows.Forms.TextBox
$script:OutputBox.Multiline = $true
$script:OutputBox.ScrollBars = 'Both'
$script:OutputBox.ReadOnly = $true
$script:OutputBox.Dock = 'Fill'
$script:OutputBox.Font = New-Object System.Drawing.Font('Consolas', 9)
$mainSplit.Panel2.Controls.Add($script:OutputBox)

$script:GcloudAuthButton.Add_Click({
        try {
            Write-WsConsole 'Launching gcloud auth login in a new window...'
            Start-Process -FilePath 'cmd.exe' -ArgumentList '/k gcloud auth login' -WindowStyle Normal
        }
        catch {
            Write-WsConsole $_.Exception.Message 'ERROR'
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'gcloud auth login', 'OK', 'Error') | Out-Null
        }
    })

$script:BuildGcloudButton.Add_Click({
        Invoke-WsGuiAction -ActionName 'Build gcloud command' -Action {
            Update-WsGcloudCommand
        }
    })

$script:RunGcloudButton.Add_Click({
        Invoke-WsGuiAction -ActionName 'Run gcloud command' -Action {
            Invoke-WsGcloudCommand
        }
    })

$script:PreviewGcloudButton.Add_Click({
        try {
            Update-WsGcloudCommand
            $previewText = $script:GcloudCommandTextBox.Text

            $previewForm = New-Object System.Windows.Forms.Form
            $previewForm.Text = 'Preview gcloud command'
            $previewForm.Size = New-Object System.Drawing.Size(900, 360)
            $previewForm.StartPosition = 'CenterParent'

            $tb = New-Object System.Windows.Forms.TextBox
            $tb.Multiline = $true
            $tb.ReadOnly = $true
            $tb.ScrollBars = 'Both'
            $tb.Font = New-Object System.Drawing.Font('Consolas', 9)
            $tb.Dock = 'Top'
            $tb.Height = 260
            $tb.Text = $previewText
            $previewForm.Controls.Add($tb)

            $ok = New-Object System.Windows.Forms.Button
            $ok.Text = 'OK'
            $ok.Width = 100
            $ok.Height = 30
            $ok.Top = 270
            $ok.Left = [Math]::Max(380, ($previewForm.ClientSize.Width - $ok.Width) / 2)
            $ok.Add_Click({ $previewForm.Close() })
            $previewForm.Controls.Add($ok)

            $previewForm.Add_Shown({ $tb.Select(0, 0) })
            $previewForm.ShowDialog() | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Preview gcloud', 'OK', 'Error') | Out-Null
        }
    })

$script:BrowseImagesButton.Add_Click({
        Invoke-WsGuiAction -ActionName 'Browse images' -Action {
            Invoke-WsImageBrowse
        }
    })

$script:ApplyTagsLabelsButton.Add_Click({
        Invoke-WsGuiAction -ActionName 'Apply Tags and Labels' -Action {
            Invoke-WsApplyTagsLabels
        }
    })

Write-WsConsole 'Fill in the fields above, then click Build Command or Run gcloud.'
[void]$form.ShowDialog()
