<#
Execution command:
powershell -ExecutionPolicy Bypass -Command "iex (irm 'https://raw.githubusercontent.com/sathishp-beep/WorkspotScripts/refs/heads/main/ADTrustRelationRepairTool.ps1')"
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$Global:CurrentLog = @()
$preferredLogFolder = 'C:\Temp\DomainTrustLogs'
$fallbackLogFolder = Join-Path ([System.IO.Path]::GetTempPath()) 'DomainTrustLogs'

try {
    New-Item -Path $preferredLogFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
    $Global:LogFolder = $preferredLogFolder
}
catch {
    New-Item -Path $fallbackLogFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
    $Global:LogFolder = $fallbackLogFolder
    $Global:CurrentLog += "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') : Using fallback log folder: $fallbackLogFolder"
}

function Write-GUILog {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp : $Message"

    if ($script:textboxOutput) {
        $script:textboxOutput.AppendText($line + [Environment]::NewLine)
        $script:textboxOutput.ScrollToCaret()
    }

    $Global:CurrentLog += $line
    [System.Windows.Forms.Application]::DoEvents()
}

function Export-CurrentLog {
    try {
        $time = Get-Date -Format 'yyyyMMdd_HHmmss'
        $file = Join-Path $Global:LogFolder "TrustRepair_$time.log"

        $Global:CurrentLog | Out-File -FilePath $file -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            "Log exported to:`n$file",
            'Export Complete',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to export log:`n$($_.Exception.Message)",
            'Export Failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

function Invoke-LoggedCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$ArgumentList = @()
    )

    if (-not (Get-Command $FilePath -ErrorAction SilentlyContinue)) {
        Write-GUILog "$FilePath was not found."
        return $false
    }

    $output = & $FilePath @ArgumentList 2>&1
    foreach ($line in $output) {
        Write-GUILog ($line | Out-String).TrimEnd()
    }

    if ($LASTEXITCODE -ne 0) {
        Write-GUILog "$FilePath exited with code $LASTEXITCODE."
        return $false
    }

    return $true
}

function Get-CurrentDomainName {
    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        if ($computerSystem.PartOfDomain) {
            return $computerSystem.Domain
        }
    }
    catch {
        Write-GUILog "Unable to determine domain membership: $($_.Exception.Message)"
    }

    return $null
}

function Get-BestDomainController {
    Write-GUILog 'Selecting optimal domain controller...'

    try {
        Import-Module ActiveDirectory -ErrorAction Stop

        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop |
        Where-Object { -not $_.IsReadOnly }

        if (-not $dcs) {
            Write-GUILog 'No writable domain controllers were found.'
            return $null
        }

        $results = foreach ($dc in $dcs) {
            try {
                $reply = Test-Connection -ComputerName $dc.HostName -Count 1 -ErrorAction Stop |
                Select-Object -First 1

                $latency = if ($reply.PSObject.Properties.Name -contains 'ResponseTime') {
                    $reply.ResponseTime
                }
                else {
                    $reply.Latency
                }

                [PSCustomObject]@{
                    DC      = $dc.HostName
                    Site    = $dc.Site
                    Latency = [int]$latency
                }
            }
            catch {
                Write-GUILog "Skipping $($dc.HostName): $($_.Exception.Message)"
            }
        }

        $best = $results | Sort-Object Latency | Select-Object -First 1

        if (-not $best) {
            Write-GUILog 'No reachable writable domain controllers were found.'
            return $null
        }

        Write-GUILog "Selected DC: $($best.DC)"
        Write-GUILog "Site: $($best.Site)"
        Write-GUILog "Latency: $($best.Latency) ms"

        return $best.DC
    }
    catch {
        Write-GUILog 'Failed selecting domain controller.'
        Write-GUILog $_.Exception.Message
        return $null
    }
}

function Test-SecureChannel {
    Write-GUILog '================================================='
    Write-GUILog 'SECURE CHANNEL VALIDATION'
    Write-GUILog '================================================='

    try {
        $domainName = Get-CurrentDomainName
        if (-not $domainName) {
            Write-GUILog 'This computer is not currently joined to a domain.'
            return
        }

        $result = Test-ComputerSecureChannel -ErrorAction Stop
        Write-GUILog "Secure Channel Status: $result"

        Invoke-LoggedCommand -FilePath 'nltest.exe' -ArgumentList @("/sc_verify:$domainName") | Out-Null
    }
    catch {
        Write-GUILog 'Secure channel validation failed.'
        Write-GUILog $_.Exception.Message
    }
}

function Repair-SecureChannel {
    Write-GUILog '================================================='
    Write-GUILog 'SECURE CHANNEL REPAIR'
    Write-GUILog '================================================='

    try {
        $domainName = Get-CurrentDomainName
        if (-not $domainName) {
            Write-GUILog 'This computer is not currently joined to a domain. Secure channel repair cannot continue.'
            return
        }

        $dc = Get-BestDomainController
        if (-not $dc) {
            Write-GUILog 'Secure channel repair cannot continue without a reachable writable domain controller.'
            return
        }

        Write-GUILog 'Purging Kerberos tickets...'
        Invoke-LoggedCommand -FilePath 'klist.exe' -ArgumentList @('purge') | Out-Null

        Write-GUILog 'Prompting for credentials with permission to reset this computer account password...'
        $credential = Get-Credential -Message 'Enter domain credentials allowed to reset this computer account password.'
        if (-not $credential) {
            Write-GUILog 'Credential prompt was cancelled. Repair stopped.'
            return
        }

        Write-GUILog 'Attempting machine password reset...'
        Reset-ComputerMachinePassword -Server $dc -Credential $credential -ErrorAction Stop
        Write-GUILog 'Machine password reset successful.'

        Write-GUILog 'Attempting secure channel repair...'
        $repair = Test-ComputerSecureChannel -Repair -Server $dc -Credential $credential -ErrorAction Stop
        Write-GUILog "Repair Result: $repair"

        Invoke-LoggedCommand -FilePath 'nltest.exe' -ArgumentList @("/sc_verify:$domainName") | Out-Null

        if ($repair) {
            Write-GUILog 'Secure channel repair successful.'
        }
        else {
            Write-GUILog 'Repair completed but the secure channel test still returned False.'
        }
    }
    catch {
        Write-GUILog 'Repair failed.'
        Write-GUILog $_.Exception.Message
    }
}

function Test-ReplicationHealth {
    Write-GUILog '================================================='
    Write-GUILog 'LDAP REPLICATION VALIDATION'
    Write-GUILog '================================================='

    Invoke-LoggedCommand -FilePath 'repadmin.exe' -ArgumentList @('/replsummary') | Out-Null
    Write-GUILog ''
    Invoke-LoggedCommand -FilePath 'repadmin.exe' -ArgumentList @('/showrepl') | Out-Null
}

function Clear-KerberosTickets {
    Write-GUILog '================================================='
    Write-GUILog 'KERBEROS TICKET PURGE'
    Write-GUILog '================================================='

    if (Invoke-LoggedCommand -FilePath 'klist.exe' -ArgumentList @('purge')) {
        Write-GUILog 'Kerberos tickets purged.'
    }
}

function Build-EventTimeline {
    Write-GUILog '================================================='
    Write-GUILog 'EVENT CORRELATION TIMELINE'
    Write-GUILog '================================================='

    $eventIDs = @(3210, 5722, 5805, 5719, 4742, 4625, 4768, 4769, 4771)
    $logNames = @('System', 'Security')

    try {
        $events = foreach ($logName in $logNames) {
            Get-WinEvent -FilterHashtable @{
                LogName   = $logName
                ID        = $eventIDs
                StartTime = (Get-Date).AddDays(-7)
            } -ErrorAction SilentlyContinue
        }

        $events = $events | Sort-Object TimeCreated

        if (-not $events) {
            Write-GUILog 'No matching events found in the last 7 days.'
            return
        }

        foreach ($event in $events) {
            Write-GUILog '-------------------------------------------------'
            Write-GUILog "Time: $($event.TimeCreated)"
            Write-GUILog "EventID: $($event.Id)"
            Write-GUILog "Source: $($event.ProviderName)"

            $msg = $event.Message -replace "(`r`n|`r|`n)+", ' '
            if ($msg.Length -gt 250) {
                $msg = $msg.Substring(0, 250)
            }

            Write-GUILog "Message: $msg"
        }
    }
    catch {
        Write-GUILog 'Timeline generation failed.'
        Write-GUILog $_.Exception.Message
    }
}

function Test-ADComputerObject {
    Write-GUILog '================================================='
    Write-GUILog 'AD COMPUTER OBJECT VALIDATION'
    Write-GUILog '================================================='

    try {
        Import-Module ActiveDirectory -ErrorAction Stop

        $computer = Get-ADComputer -Identity $env:COMPUTERNAME -Properties Created, LastLogonDate, PasswordLastSet, SID -ErrorAction Stop

        Write-GUILog "Computer Name: $($computer.Name)"
        Write-GUILog "Enabled: $($computer.Enabled)"
        Write-GUILog "Created: $($computer.Created)"
        Write-GUILog "LastLogonDate: $($computer.LastLogonDate)"
        Write-GUILog "PasswordLastSet: $($computer.PasswordLastSet)"
        Write-GUILog "DistinguishedName: $($computer.DistinguishedName)"
        Write-GUILog "SID: $($computer.SID)"
    }
    catch {
        Write-GUILog 'AD validation failed.'
        Write-GUILog $_.Exception.Message
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Enterprise Domain Trust Relationship Repair Tool'
$form.Size = New-Object System.Drawing.Size(1400, 900)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor = [System.Drawing.Color]::White

$header = New-Object System.Windows.Forms.Label
$header.Text = 'Enterprise Active Directory Trust Relationship Repair Tool'
$header.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
$header.AutoSize = $true
$header.Location = New-Object System.Drawing.Point(20, 10)
$form.Controls.Add($header)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Size = New-Object System.Drawing.Size(1350, 400)
$tabs.Location = New-Object System.Drawing.Point(20, 60)
$form.Controls.Add($tabs)

$script:textboxOutput = New-Object System.Windows.Forms.RichTextBox
$script:textboxOutput.Location = New-Object System.Drawing.Point(20, 480)
$script:textboxOutput.Size = New-Object System.Drawing.Size(1350, 320)
$script:textboxOutput.BackColor = [System.Drawing.Color]::Black
$script:textboxOutput.ForeColor = [System.Drawing.Color]::Lime
$script:textboxOutput.Font = New-Object System.Drawing.Font('Consolas', 10)
$script:textboxOutput.ReadOnly = $true
$form.Controls.Add($script:textboxOutput)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Save / Export Logs'
$btnExport.Size = New-Object System.Drawing.Size(180, 35)
$btnExport.Location = New-Object System.Drawing.Point(1180, 20)
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnExport.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnExport)
$btnExport.Add_Click({ Export-CurrentLog })

$tab1 = New-Object System.Windows.Forms.TabPage
$tab1.Text = 'Secure Channel'
$tabs.TabPages.Add($tab1) | Out-Null

$btnTestSC = New-Object System.Windows.Forms.Button
$btnTestSC.Text = 'Test Secure Channel'
$btnTestSC.Size = New-Object System.Drawing.Size(220, 50)
$btnTestSC.Location = New-Object System.Drawing.Point(30, 30)
$btnTestSC.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnTestSC.ForeColor = [System.Drawing.Color]::White
$tab1.Controls.Add($btnTestSC)
$btnTestSC.Add_Click({ Test-SecureChannel })

$btnRepairSC = New-Object System.Windows.Forms.Button
$btnRepairSC.Text = 'Fix / Repair Secure Channel'
$btnRepairSC.Size = New-Object System.Drawing.Size(260, 50)
$btnRepairSC.Location = New-Object System.Drawing.Point(280, 30)
$btnRepairSC.BackColor = [System.Drawing.Color]::DarkRed
$btnRepairSC.ForeColor = [System.Drawing.Color]::White
$tab1.Controls.Add($btnRepairSC)
$btnRepairSC.Add_Click({ Repair-SecureChannel })

$tab2 = New-Object System.Windows.Forms.TabPage
$tab2.Text = 'Replication'
$tabs.TabPages.Add($tab2) | Out-Null

$btnReplication = New-Object System.Windows.Forms.Button
$btnReplication.Text = 'Validate LDAP Replication'
$btnReplication.Size = New-Object System.Drawing.Size(260, 50)
$btnReplication.Location = New-Object System.Drawing.Point(30, 30)
$btnReplication.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnReplication.ForeColor = [System.Drawing.Color]::White
$tab2.Controls.Add($btnReplication)
$btnReplication.Add_Click({ Test-ReplicationHealth })

$tab3 = New-Object System.Windows.Forms.TabPage
$tab3.Text = 'Kerberos'
$tabs.TabPages.Add($tab3) | Out-Null

$btnKerb = New-Object System.Windows.Forms.Button
$btnKerb.Text = 'Purge Kerberos Tickets'
$btnKerb.Size = New-Object System.Drawing.Size(240, 50)
$btnKerb.Location = New-Object System.Drawing.Point(30, 30)
$btnKerb.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnKerb.ForeColor = [System.Drawing.Color]::White
$tab3.Controls.Add($btnKerb)
$btnKerb.Add_Click({ Clear-KerberosTickets })

$tab4 = New-Object System.Windows.Forms.TabPage
$tab4.Text = 'Timeline'
$tabs.TabPages.Add($tab4) | Out-Null

$btnTimeline = New-Object System.Windows.Forms.Button
$btnTimeline.Text = 'Build Event Timeline'
$btnTimeline.Size = New-Object System.Drawing.Size(240, 50)
$btnTimeline.Location = New-Object System.Drawing.Point(30, 30)
$btnTimeline.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnTimeline.ForeColor = [System.Drawing.Color]::White
$tab4.Controls.Add($btnTimeline)
$btnTimeline.Add_Click({ Build-EventTimeline })

$tab5 = New-Object System.Windows.Forms.TabPage
$tab5.Text = 'AD Object'
$tabs.TabPages.Add($tab5) | Out-Null

$btnAD = New-Object System.Windows.Forms.Button
$btnAD.Text = 'Validate AD Computer Object'
$btnAD.Size = New-Object System.Drawing.Size(280, 50)
$btnAD.Location = New-Object System.Drawing.Point(30, 30)
$btnAD.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnAD.ForeColor = [System.Drawing.Color]::White
$tab5.Controls.Add($btnAD)
$btnAD.Add_Click({ Test-ADComputerObject })

$tab6 = New-Object System.Windows.Forms.TabPage
$tab6.Text = 'Domain Controller'
$tabs.TabPages.Add($tab6) | Out-Null

$btnDC = New-Object System.Windows.Forms.Button
$btnDC.Text = 'Select Best Domain Controller'
$btnDC.Size = New-Object System.Drawing.Size(300, 50)
$btnDC.Location = New-Object System.Drawing.Point(30, 30)
$btnDC.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnDC.ForeColor = [System.Drawing.Color]::White
$tab6.Controls.Add($btnDC)
$btnDC.Add_Click({ Get-BestDomainController | Out-Null })

$footer = New-Object System.Windows.Forms.Label
$footer.Text = 'Enterprise Active Directory Trust Repair Utility'
$footer.AutoSize = $true
$footer.Location = New-Object System.Drawing.Point(20, 820)
$form.Controls.Add($footer)

[void]$form.ShowDialog()
