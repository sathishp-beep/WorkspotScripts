

<#
.SYNOPSIS
    Production-grade Workspot Client Troubleshooting Menu
    Interactive script with Y/N prompts for each diagnostic task

.DESCRIPTION
    Comprehensive client data capture, Entra ID management, logging, 
    and event log upload for Workspot troubleshooting
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param()

#region Initialization
$ErrorActionPreference = 'Stop'
$script:LogPath = "C:\Temp\WSClientData"
$script:UserName = $env:USERNAME

# Full .reg file path under the data folder
$script:WSDebugLogging = Join-Path $script:LogPath "WS_Advance_logging_v2.reg"

# Prompt for target tenant information
#$script:TargetTenant = "Siemens Energy"
#$script:TargetTenantId = "254ba93e-1f6f-48f3-90f6-e2766664b477"
$script:TargetTenant = Read-Host "Enter Target Tenant Name"
$script:TargetTenantId = Read-Host "Enter Target Tenant ID"
#endregion

function Write-SectionHeader {
    param([string]$Title)
    Clear-Host
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host " $Title " -ForegroundColor White -BackgroundColor DarkCyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host ""
}

function Show-MainMenu {
    Write-SectionHeader "Workspot Client Troubleshooting Menu"
    Write-Host "1. Create Data Collection Folder" -ForegroundColor Green
    Write-Host "2. Enable Workspot Client Debug Logs" -ForegroundColor Green
    Write-Host "3. Capture dsregcmd Status" -ForegroundColor Green
    Write-Host "4. Unregister (remove Work/School) from Siemens Energy Entra ID" -ForegroundColor Green
    Write-Host "5. Re-register (add Work/School) to Siemens Energy" -ForegroundColor Green
    Write-Host "6. Collect Workspot Client Logs into Data Collection" -ForegroundColor Green
    Write-Host "7. Enable Always connect RegKey" -ForegroundColor Green
    Write-Host "8. Disable Always connect RegKey" -ForegroundColor Green
    Write-Host "9. Upload Client Event Logs to Workspot(CSR Azure)" -ForegroundColor Green
    Write-Host "10. Check Workspot Gateway Certificate Issuer" -ForegroundColor Green
    Write-Host "11. Get Installed Programs" -ForegroundColor Green
    Write-Host "12. Capture Group Policy Result (gpresult)" -ForegroundColor Green
    Write-Host "13. Get Intune Policy Results" -ForegroundColor Green
    Write-Host "0. Exit" -ForegroundColor Red
    Write-Host ""
}

function Confirm-Action {
    param([string]$Message)
    do {
        $response = Read-Host "$Message (Y/N)"
        if ($response -match '^[Yy]') { return $true }
        elseif ($response -match '^[Nn]') { return $false }
        else { Write-Host "Please enter Y or N" -ForegroundColor Yellow }
    } while ($true)
}

function Get-InstalledPrograms {
    param([string]$OutputFile)
    
    try {
        Write-Host "Collecting installed programs..." -ForegroundColor Yellow
        
        # Collect from both 32-bit and 64-bit registry
        $programs = @()
        
        # 64-bit registry
        $programs += Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
        Where-Object { $_.DisplayName }
        
        # 32-bit registry
        $programs += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
        Where-Object { $_.DisplayName }
        
        # Remove duplicates and sort
        $programs = $programs | Sort-Object DisplayName -Unique
        
        # Export to file
        $programs | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
        
        Write-Host "Installed programs saved to: $OutputFile" -ForegroundColor Green
        Write-Host "Total programs found: $($programs.Count)" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to collect installed programs: $($_.Exception.Message)"
    }
}

function Initialize-DataFolder {
    if (-not (Test-Path $script:LogPath)) {
        New-Item -Path $script:LogPath -ItemType Directory -Force | Out-Null
        Write-Host "Created folder: $script:LogPath" -ForegroundColor Green
    }
    else {
        Write-Host "Folder exists: $script:LogPath" -ForegroundColor Green
    }
}

function Enable-DebugLogs {
    # Ensure destination folder exists
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

        $proc = Start-Process -FilePath "$env:WINDIR\System32\regedit.exe" `
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

function Capture-DsregcmdStatus {
    Initialize-DataFolder
    $outputFile = "$script:LogPath\dsregcmd_status_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    dsregcmd /status | Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host "dsregcmd status saved to: $outputFile" -ForegroundColor Green
}

function Unregister-EntraID {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    $accounts = Get-WmiObject -Namespace "root\cimv2\mdm\dmmap" -Class MDM_DevDetail_Ext01 -ErrorAction SilentlyContinue
    
    $status = dsregcmd /status
    $workAccounts = ($status -join "`n") -split '\+----------------------------------------------------------------------\+'
    
    $found = $false
    foreach ($block in $workAccounts) {
        if ($block -match "WorkplaceTenantName\s+:\s+$script:TargetTenant") {
            Write-Host "Found Work Account for $script:TargetTenant" -ForegroundColor Yellow
            if ($block -match "WorkplaceDeviceId\s+:\s+([a-z0-9-]+)") {
                $deviceId = $matches[1]
                Write-Host "WorkplaceDeviceId: $deviceId" -ForegroundColor Cyan
                
                $confirm = Confirm-Action "Remove this account"
                if ($confirm) {
                    $regPath = "HKCU:\Software\Microsoft\Enrollments"
                    Get-ChildItem $regPath -ErrorAction SilentlyContinue | ForEach-Object {
                        $tenant = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).AADTenantID
                        if ($tenant -eq $script:TargetTenantId) {
                            Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Host "Removed Siemens Energy enrollment." -ForegroundColor Green
                            $found = $true
                        }
                    }
                    if ($found) {
                        Write-Host "Sign out and sign back in required." -ForegroundColor Cyan
                    }
                }
            }
        }
    }
    if (-not $found) {
        Write-Host "No Siemens Energy work account found." -ForegroundColor Yellow
    }
}

function ReRegister-EntraID {
    Write-Host "Triggering Siemens Energy Work Account Enrollment..." -ForegroundColor Green
    Start-Process "ms-settings:workplace"
    Write-Host "Settings window opened. Complete enrollment manually." -ForegroundColor Yellow
    Start-Sleep 3
}

function Capture-WorkspotLogs {
    Initialize-DataFolder
    $sourcePath = "$env:LOCALAPPDATA\workspot\client\log"
    $destFolder = "$script:LogPath\WsClientLogs_$script:UserName"
    
    if (Test-Path $sourcePath) {
        Remove-Item $destFolder -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item $sourcePath $destFolder -Recurse -Force
        Write-Host "Workspot logs copied to: $destFolder" -ForegroundColor Green
    }
    else {
        Write-Warning "Workspot log path not found: $sourcePath"
    }
}

function Add-NonNLARegistry {
    $regPath = "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services"
    $regValue = "AuthenticationLevel"
    $regData = 0  # Always connect, even if authentication fails
    
    $parentPath = Split-Path $regPath -Parent
    if (-not (Test-Path $parentPath)) {
        New-Item -Path $parentPath -Force | Out-Null
    }
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    
    Set-ItemProperty -Path $regPath -Name $regValue -Value $regData -Type DWord -Force
    Write-Host "Non-NLA registry added: $regPath\$regValue = 0" -ForegroundColor Green
}

function Remove-NonNLARegistry {
    $regPath = "HKLM:\Software\Policies\Microsoft\Windows NT\Terminal Services"
    $regValue = "AuthenticationLevel"
    
    if (Test-Path $regPath) {
        Remove-ItemProperty -Path $regPath -Name $regValue -Force -ErrorAction SilentlyContinue
        Write-Host "Non-NLA registry removed from: $regPath" -ForegroundColor Green
    }
    else {
        Write-Host "Non-NLA registry not found." -ForegroundColor Yellow
    }
}

function Upload-EventLogs {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $azCopyDir = "C:\Users\AzCopy"
        $azCopyZip = "$azCopyDir\azcopy.zip"
        $azCopyUrl = "https://aka.ms/downloadazcopy-v10-windows"
        
        # Setup AzCopy if not exists
        if (-not (Test-Path "$azCopyDir\azcopy.exe")) {
            New-Item -Path $azCopyDir -ItemType Directory -Force | Out-Null
            Invoke-WebRequest $azCopyUrl -OutFile $azCopyZip
            Expand-Archive -LiteralPath $azCopyZip -DestinationPath $azCopyDir -Force
            $azCopyFolder = Get-ChildItem -Path $azCopyDir -Name "*azcopy_W*" | Select-Object -First 1
            Rename-Item "$azCopyDir\$azCopyFolder" "azcopy" -Force
        }
        
        cd "$azCopyDir\azcopy"
        $thisMachine = [System.Net.Dns]::GetHostByName($env:COMPUTERNAME).HostName
        $source = "C:\Windows\System32\winevt\Logs"
        $sourceCopy = "${source}-$thisMachine"
        
        # Copy logs
        Copy-Item -Path $source -Recurse -Destination $sourceCopy -Force
        
        # Upload
        $DestSASKey = "https://csrsupportstorageacc.blob.core.windows.net/workspot-agent-logs?sp=racwl&st=2025-08-21T10:55:29Z&se=2050-08-21T19:10:29Z&sv=2024-11-04&sr=c&sig=3XtKGfn6%2Fo6NVgOXsk7E0S%2FD5oEE1LBEl%2FtDmkK7BeY%3D"
        .\azcopy.exe cp $sourceCopy $DestSASKey --recursive=true
        
        Write-Host "Event logs uploaded successfully for $thisMachine" -ForegroundColor Green
        
        # Cleanup
        Remove-Item $sourceCopy -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Error "Event log upload failed: $($_.Exception.Message)"
    }
}

function Capture-GroupPolicy {
    Initialize-DataFolder
    $outputFile = "$script:LogPath\gpresult_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    
    try {
        Write-Host "Capturing Group Policy Result..." -ForegroundColor Yellow
        gpresult /h $outputFile
        
        if (Test-Path $outputFile) {
            Write-Host "Group Policy result saved to: $outputFile" -ForegroundColor Green
        }
        else {
            Write-Warning "Group Policy report generation may have failed."
        }
    }
    catch {
        Write-Error "Failed to capture Group Policy result: $($_.Exception.Message)"
    }
}

function Get-WebUrlCertificate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    try {
        # Force modern TLS
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor `
            [Net.SecurityProtocolType]::Tls13

        # Create HttpWebRequest but don't care about 403, just want the cert
        $request = [System.Net.HttpWebRequest]::Create($Url)
        $request.Method = "HEAD"  # HEAD avoids fetching the body
        $request.AllowAutoRedirect = $true

        # Get response (will throw 403 if forbidden, but certificate can still be accessed)
        try {
            $response = $request.GetResponse()
            $response.Close()
        }
        catch [System.Net.WebException] {
            # Even if 403/401, the ServicePoint.Certificate is still populated
        }

        # Get certificate
        $cert = $request.ServicePoint.Certificate
        $certObj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $cert

        # Output
        $certOutput = [PSCustomObject]@{
            URL           = $Url
            Subject       = $certObj.Subject
            Issuer        = $certObj.Issuer
            NotBefore     = $certObj.NotBefore
            NotAfter      = $certObj.NotAfter
            DaysRemaining = ($certObj.NotAfter - (Get-Date)).Days
            Thumbprint    = $certObj.Thumbprint
            SerialNumber  = $certObj.SerialNumber
        }
        
        # Write output to file
        $outputFile = "$script:LogPath\WebURICertIssuer_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $certOutput | Out-File -FilePath $outputFile -Append -Force
        Write-Host "Certificate information written to $outputFile" -ForegroundColor Green
        
        $certOutput
    }
    catch {
        Write-Error "Failed to retrieve certificate from $Url. $($_.Exception.Message)"
    }
}

function Get-IntunePolicyResults {
    param([string]$OutputFile)
    
    try {
        Write-Host "Collecting Intune Policy Results from PolicyManager hive..." -ForegroundColor Yellow
        
        $policyPath = "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device"
        
        if (-not (Test-Path $policyPath)) {
            Write-Warning "PolicyManager hive not found at: $policyPath"
            return
        }
        
        $policyResults = @()
        
        # Recursively get all device policies
        $policies = Get-ChildItem -Path $policyPath -Recurse -ErrorAction SilentlyContinue
        
        foreach ($policy in $policies) {
            $properties = Get-ItemProperty -Path $policy.PSPath -ErrorAction SilentlyContinue
            
            foreach ($prop in $properties.PSObject.Properties) {
                if ($prop.Name -notmatch '^PS') {
                    # Exclude PS properties
                    $policyResults += [PSCustomObject]@{
                        Path  = $policy.PSPath
                        Name  = $prop.Name
                        Value = $prop.Value
                        Type  = $prop.MemberType
                    }
                }
            }
        }
        
        if ($policyResults.Count -eq 0) {
            Write-Host "No device policies found in PolicyManager hive." -ForegroundColor Yellow
        }
        else {
            # Export to CSV file
            $policyResults | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
            Write-Host "Intune policy results saved to: $OutputFile" -ForegroundColor Green
            Write-Host "Total policies found: $($policyResults.Count)" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to collect Intune policy results: $($_.Exception.Message)"
    }
}

# Main Loop
while ($true) {
    Show-MainMenu
    $choice = Read-Host "Select/Choose your option"
    
    switch ($choice) {
        '1' { 
            if (Confirm-Action "Would you like to create the data folder") {
                Initialize-DataFolder
                Read-Host "Press Enter to continue"
            }
        }
        '2' { 
            if (Confirm-Action "Would you like to enable debug logs") {
                Enable-DebugLogs
                Read-Host "Press Enter to continue"
            }
        }
        '3' { 
            if (Confirm-Action "Would you like to capture dsregcmd status") {
                Capture-DsregcmdStatus
                Read-Host "Press Enter to continue"
            }
        }
        '4' { 
            if (Confirm-Action "Would you like to unregister from Siemens Energy Entra ID") {
                Unregister-EntraID
                Read-Host "Press Enter to continue"
            }
        }
        '5' { 
            if (Confirm-Action "Would you like to re-register to Siemens Energy Entra ID") {
                ReRegister-EntraID
                Read-Host "Press Enter to continue"
            }
        }
        '6' { 
            if (Confirm-Action "Would you like to capture Workspot client logs") {
                Capture-WorkspotLogs
                Read-Host "Press Enter to continue"
            }
        }
        '7' { 
            if (Confirm-Action "Would you like to add Non-NLA GPO registry") {
                Add-NonNLARegistry
                Read-Host "Press Enter to continue"
            }
        }
        '8' { 
            if (Confirm-Action "Would you like to remove Non-NLA GPO registry") {
                Remove-NonNLARegistry
                Read-Host "Press Enter to continue"
            }
        }
        '9' { 
            if (Confirm-Action "Would you like to upload event logs") {
                Upload-EventLogs
                Read-Host "Press Enter to continue"
            }
        }
        '10' { 
            if (Confirm-Action "Would you like to check Workspot Gateway certificate issuer") {
                $url = Read-Host "Enter the Workspot Gateway URL to check certificate issuer"
                if ($url) {
                    Get-WebUrlCertificate -Url $url
                }
                Read-Host "Press Enter to continue"
            }
        }
        '11' { 
            if (Confirm-Action "Would you like to get installed programs") {
                $outputFile = "$script:LogPath\InstalledPrograms_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                Get-InstalledPrograms -OutputFile $outputFile
                Read-Host "Press Enter to continue"
            }
        }
        '12' { 
            if (Confirm-Action "Would you like to capture Group Policy result") {
                Capture-GroupPolicy
                Read-Host "Press Enter to continue"
            }
        }
        '13' { 
            if (Confirm-Action "Would you like to get Intune policy results") {
                $outputFile = "$script:LogPath\IntunePolicyResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                Get-IntunePolicyResults -OutputFile $outputFile
                Read-Host "Press Enter to continue"
            }
        }
        '0' { exit }
        default { Write-Host "Invalid option. Press Enter to continue"; Read-Host }
    }
}
