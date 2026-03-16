param(
    [string]$SystemDrive = "",
    [switch]$AutoReboot,
    [string]$UploadEndpoint = "https://diagnostics.microsoft.com/upload",
    [switch]$Force,
    [switch]$Silent,
    [switch]$DryRun,
    [switch]$WhatIf,
    [switch]$Rollback
)

$Global:ScriptName = "Windows Emergency Fixer"
$Global:StartTime = (Get-Date).ToString("yyyyMMdd_HHmmss")
$Global:LogRoot = "C:\RecoveryLogs"
$Global:UndoManifestPath = Join-Path $Global:LogRoot ("UndoManifest_$Global:StartTime.json")
$Global:SessionLog = Join-Path $Global:LogRoot ("Session_$Global:StartTime.log")
$Global:DesktopZip = Join-Path ([Environment]::GetFolderPath("Desktop")) ("Emergency_Windows_Fixer_Logs_$Global:StartTime.zip")
$Global:QuarantineFolder = Join-Path $Global:LogRoot ("Quarantine_$Global:StartTime")
$Global:UploadFlag = Join-Path $Global:LogRoot "UploadDone.flag"
$Global:CapabilityCache = @{}
$Global:ErrorCount = 0
$Global:RebootRequired = $false
$Global:ParallelJobs = @()
$Global:Force = $true
if ($DryRun) { $Global:DryRun = $true } else { $Global:DryRun = $false }
if ($WhatIf) { $Global:WhatIf = $true } else { $Global:WhatIf = $false }
if (-not (Test-Path $Global:LogRoot)) { New-Item -Path $Global:LogRoot -ItemType Directory -Force | Out-Null }
if (-not (Test-Path $Global:QuarantineFolder)) { New-Item -Path $Global:QuarantineFolder -ItemType Directory -Force | Out-Null }

function Ensure-Elevated {
    try {
        $isAdmin = $false
        try { $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch {}
        if (-not $isAdmin) {
            $args = $MyInvocation.UnboundArguments
            $argString = $args | ForEach-Object { if ($_ -is [switch]) { $_.ToString() } else { '"' + ($_ -replace '"','\"') + '"' } } -join " "
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = (Get-Process -Id $PID).Path
            $psi.Arguments = $argString
            $psi.Verb = "runas"
            try { [System.Diagnostics.Process]::Start($psi) | Out-Null; exit 0 } catch {}
        }
    } catch {}
}

function Start-Logging {
    try { Start-Transcript -Path $Global:SessionLog -Force -ErrorAction SilentlyContinue } catch {}
}

function Write-Log {
    param([string]$Message,[string]$Level="INFO")
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts] [$Level] $Message"
    try { Add-Content -Path $Global:SessionLog -Value $line -Force -ErrorAction SilentlyContinue } catch {}
    try { Add-Content -Path (Join-Path ([Environment]::GetFolderPath("Desktop")) "EmergencyWindowsFixer_Log.txt") -Value $line -Force -ErrorAction SilentlyContinue } catch {}
}

function Init-Capabilities {
    $list = @(
        "pnputil.exe","reg.exe","bcdedit.exe","bootrec.exe","DISM.exe","sfc.exe","chkdsk.exe","chkntfs.exe",
        "Get-MpPreference","Remove-MpPreference","Set-MpPreference","Start-MpScan","Start-MpWDOScan","Get-MpComputerStatus",
        "Get-Volume","Get-Partition","Get-Disk","Get-PnpDevice","Checkpoint-Computer","Compress-Archive","Import-Certificate",
        "Import-PfxCertificate","Get-WinEvent","Get-Process","Get-Service","Invoke-RestMethod","Start-Job","Stop-Job","Get-AuthenticodeSignature"
    )
    foreach ($c in $list) {
        $found = $false
        try { $found = (Get-Command -Name $c -ErrorAction SilentlyContinue) -ne $null } catch { $found = $false }
        $Global:CapabilityCache[$c] = $found
    }
}

function Save-UndoManifest {
    try {
        $Global:UndoManifest | ConvertTo-Json -Depth 6 | Set-Content -Path $Global:UndoManifestPath -Force
    } catch {}
}

function Start-UndoManifest {
    $Global:UndoManifest = [PSCustomObject]@{ Time = (Get-Date).ToString("o"); Script = $Global:ScriptName; Entries = @() }
    Save-UndoManifest
}

function Add-UndoEntry {
    param([string]$Type,[hashtable]$Data,[string]$RevertCommand="")
    try {
        $entry = [PSCustomObject]@{ Time = (Get-Date).ToString("o"); Type = $Type; Data = $Data; RevertCommand = $RevertCommand }
        $Global:UndoManifest.Entries += $entry
        Save-UndoManifest
    } catch {}
}

function Rollback-From-Manifest {
    param([string]$ManifestFile)
    try {
        if (-not (Test-Path $ManifestFile)) { return }
        $m = Get-Content -Path $ManifestFile -Raw | ConvertFrom-Json
        foreach ($e in ($m.Entries | Sort-Object -Property Time -Descending)) {
            try {
                switch ($e.Type) {
                    "RegistryExport" { if (Test-Path $e.Data.Path) { reg.exe import $e.Data.Path | Out-Null } }
                    "FileMove" { if (Test-Path $e.Data.Target) { Move-Item -Path $e.Data.Target -Destination $e.Data.Source -Force -ErrorAction SilentlyContinue } }
                    "DriverRestore" { if (Test-Path $e.Data.BackupFolder) { pnputil.exe /add-driver (Join-Path $e.Data.BackupFolder "*") /install | Out-Null } }
                    default {}
                }
            } catch {}
        }
    } catch {}
}

function Try-Log {
    param([scriptblock]$Action,[string]$Name)
    try { & $Action } catch {}
}

function Export-RegistryHive {
    param([string]$Hive,[string]$OutFile)
    try { reg.exe export $Hive $OutFile /y | Out-Null
        Add-UndoEntry -Type "RegistryExport" -Data @{ Hive = $Hive; Path = $OutFile } -RevertCommand "reg import `"$OutFile`"" } catch {}
}

function Create-RestorePoint {
    param([string]$Desc)
    try { if ($Global:CapabilityCache["Checkpoint-Computer"] -or (Get-Command -Name Checkpoint-Computer -ErrorAction SilentlyContinue)) {
        Checkpoint-Computer -Description $Desc -RestorePointType "MODIFY_SETTINGS" -ErrorAction SilentlyContinue
        Add-UndoEntry -Type "RestorePoint" -Data @{ Description = $Desc } -RevertCommand "" } } catch {}
}

function Backup-Drivers {
    try {
        $out = Join-Path $Global:LogRoot ("DriverBackup_$Global:StartTime")
        if (-not (Test-Path $out)) { New-Item -Path $out -ItemType Directory -Force | Out-Null }
        if ($Global:CapabilityCache["pnputil.exe"]) {
            pnputil.exe /export-driver * $out | Out-Null
            Add-UndoEntry -Type "DriverRestore" -Data @{ BackupFolder = $out } -RevertCommand "pnputil /add-driver `"$out\*`" /install"
        }
    } catch {}
}

function Detect-WinPE-Or-ToGo {
    $isWinPE = $false
    $isToGo = $false
    try { $pe = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'PEBoot' -ErrorAction SilentlyContinue).PEBoot; if ($pe -eq 1) { $isWinPE = $true } } catch {}
    try { $pt = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'PortableOperatingSystem' -ErrorAction SilentlyContinue).PortableOperatingSystem; if ($pt -eq 1) { $isToGo = $true } } catch {}
    return @{ IsWinPE = $isWinPE; IsToGo = $isToGo }
}

function Get-WindowsPartition {
    $candidates = @()
    if ($Global:CapabilityCache["Get-Volume"]) {
        try {
            $vols = Get-Volume -ErrorAction SilentlyContinue
            foreach ($v in $vols) {
                if ($v.DriveLetter) {
                    $root = "$($v.DriveLetter):\"
                    if (Test-Path (Join-Path $root "Windows\System32")) {
                        $candidates += [PSCustomObject]@{ Drive = $v.DriveLetter; Size = $v.Size; Path = $root }
                    }
                }
            }
        } catch {}
    } else {
        try {
            $drives = Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue
            foreach ($d in $drives) {
                $root = $d.Root
                if (Test-Path (Join-Path $root "Windows\System32")) {
                    $size = 0
                    try { $size = (Get-ChildItem -Path $root -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum } catch {}
                    $candidates += [PSCustomObject]@{ Drive = $d.Name; Size = $size; Path = $root }
                }
            }
        } catch {}
    }
    if ($candidates.Count -eq 0) { return $null }
    return $candidates | Sort-Object -Property Size -Descending | Select-Object -First 1
}

function Collect-Diagnostics {
    try {
        $out = Join-Path $Global:LogRoot ("Diag_$Global:StartTime")
        if (-not (Test-Path $out)) { New-Item -Path $out -ItemType Directory -Force | Out-Null }
        Try-Log -Action { Get-WinEvent -LogName System -MaxEvents 500 | Out-File (Join-Path $out "System.txt") } -Name "Collect System Log"
        Try-Log -Action { Get-WinEvent -LogName Application -MaxEvents 500 | Out-File (Join-Path $out "Application.txt") } -Name "Collect Application Log"
        Try-Log -Action { ipconfig /all | Out-File (Join-Path $out "IPConfig.txt") } -Name "Collect IPConfig"
        Try-Log -Action { Get-Process | Out-File (Join-Path $out "Processes.txt") } -Name "Collect Processes"
        Try-Log -Action { Get-Service | Out-File (Join-Path $out "Services.txt") } -Name "Collect Services"
        Add-UndoEntry -Type "Diagnostics" -Data @{ Path = $out } -RevertCommand ""
        try { Compress-Archive -Path (Join-Path $out "*") -DestinationPath $Global:DesktopZip -Force -ErrorAction SilentlyContinue } catch {}
    } catch {}
}

function Remove-Defender-Exclusions {
    try {
        if (-not $Global:CapabilityCache["Get-MpPreference"]) { return }
        $prefs = Get-MpPreference -ErrorAction SilentlyContinue
        if (-not $prefs) { return }
        $record = @{ Paths = $prefs.ExclusionPath; Extensions = $prefs.ExclusionExtension; Processes = $prefs.ExclusionProcess }
        Add-UndoEntry -Type "DefenderExclusions" -Data $record -RevertCommand ""
        foreach ($p in $prefs.ExclusionPath) { if ($p) { Remove-MpPreference -ExclusionPath $p -Force -ErrorAction SilentlyContinue } }
        foreach ($e in $prefs.ExclusionExtension) { if ($e) { Remove-MpPreference -ExclusionExtension $e -Force -ErrorAction SilentlyContinue } }
        foreach ($pr in $prefs.ExclusionProcess) { if ($pr) { Remove-MpPreference -ExclusionProcess $pr -Force -ErrorAction SilentlyContinue } }
    } catch {}
}

function Enable-Defender-PUA {
    try {
        if (-not $Global:CapabilityCache["Set-MpPreference"]) { return }
        Set-MpPreference -PUAProtection Enabled -Force -ErrorAction SilentlyContinue
    } catch {}
}

function Start-Defender-Scans {
    try {
        if ($Global:CapabilityCache["Start-MpScan"]) { Start-MpScan -ScanType FullScan -ErrorAction SilentlyContinue }
        else {
            $mp = Join-Path $env:ProgramFiles "Windows Defender\MpCmdRun.exe"
            if (Test-Path $mp) { & $mp -Scan -ScanType 2 | Out-Null }
        }
    } catch {}
}

function Schedule-DefenderOffline {
    try {
        if ($Global:CapabilityCache["Start-MpWDOScan"]) { Start-MpWDOScan -ErrorAction SilentlyContinue; $Global:RebootRequired = $true }
        $mp = Join-Path $env:ProgramFiles "Windows Defender\MpCmdRun.exe"
        if (Test-Path $mp) { & $mp -Scan -ScanType 2 -OfflineScan | Out-Null; $Global:RebootRequired = $true }
    } catch {}
}

function Upload-LogsOnce {
    param([string]$Endpoint)
    try {
        if (-not $Endpoint) { return }
        if (Test-Path $Global:UploadFlag) { return }
        if (-not (Test-Path $Global:DesktopZip)) { Collect-Diagnostics }
        if (-not (Test-Path $Global:DesktopZip)) { return }
        $bytes = [IO.File]::ReadAllBytes($Global:DesktopZip)
        $enc = [Convert]::ToBase64String($bytes)
        $payload = @{ Machine = $env:COMPUTERNAME; Time = (Get-Date).ToString("o"); Payload = $enc } | ConvertTo-Json -Depth 5
        Invoke-RestMethod -Uri $Endpoint -Method Post -Body $payload -ContentType "application/json" -TimeoutSec 120
        New-Item -Path $Global:UploadFlag -ItemType File -Force | Out-Null
        Add-UndoEntry -Type "UploadPerformed" -Data @{ Endpoint = $Endpoint } -RevertCommand ""
    } catch {}
}

function Parallel-Start {
    param([scriptblock]$Action,[string]$Name)
    try { $job = Start-Job -ScriptBlock $Action -ErrorAction SilentlyContinue; $Global:ParallelJobs += $job } catch {}
}

function Parallel-Wait {
    try {
        foreach ($j in $Global:ParallelJobs) { try { Wait-Job -Job $j -ErrorAction SilentlyContinue; Receive-Job -Job $j -ErrorAction SilentlyContinue | Out-Null; Remove-Job -Job $j -ErrorAction SilentlyContinue } catch {} }
        $Global:ParallelJobs = @()
    } catch {}
}

Ensure-Elevated
Start-Logging
Init-Capabilities
Start-UndoManifest

Try-Log -Action { Export-RegistryHive -Hive "HKLM\SOFTWARE" -OutFile (Join-Path $Global:LogRoot ("HKLM_SOFTWARE_$Global:StartTime.reg")) } -Name "Export HKLM_SOFTWARE"
Try-Log -Action { Export-RegistryHive -Hive "HKLM\SYSTEM" -OutFile (Join-Path $Global:LogRoot ("HKLM_SYSTEM_$Global:StartTime.reg")) } -Name "Export HKLM_SYSTEM"
Try-Log -Action { Export-RegistryHive -Hive "HKCU\Software" -OutFile (Join-Path $Global:LogRoot ("HKCU_Software_$Global:StartTime.reg")) } -Name "Export HKCU_Software"
Create-RestorePoint -Desc "PreChanges_$Global:StartTime"
Backup-Drivers
Remove-Defender-Exclusions
Enable-Defender-PUA
Parallel-Start -Action { Start-Defender-Scans } -Name "DefenderScan"
Parallel-Start -Action { Collect-Diagnostics } -Name "CollectDiagnostics"

# Disk/Partition/SMART and Windows protection/repair actions
Try-Log -Action { Start-Process "sfc.exe" -ArgumentList "/scannow" -Wait -ErrorAction SilentlyContinue } -Name "SFC"
Try-Log -Action { Start-Process "dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -Wait -ErrorAction SilentlyContinue } -Name "DISM"
Try-Log -Action { Start-Process "chkdsk.exe" -ArgumentList "C: /f /r" -Wait -ErrorAction SilentlyContinue } -Name "CHKDSK"
Try-Log -Action { Get-WmiObject -Query "SELECT * FROM Win32_DiskDrive" | ForEach-Object { $_.Status; $_.IsSMARTEnabled } } -Name "DiskSMART"
Try-Log -Action { Start-Process "bootrec.exe" -ArgumentList "/fixmbr" -Wait -ErrorAction SilentlyContinue } -Name "BootRec_FixMBR"
Try-Log -Action { Start-Process "bootrec.exe" -ArgumentList "/fixboot" -Wait -ErrorAction SilentlyContinue } -Name "BootRec_FixBoot"
Try-Log -Action { Start-Process "bootrec.exe" -ArgumentList "/scanos" -Wait -ErrorAction SilentlyContinue } -Name "BootRec_ScanOS"
Try-Log -Action { Start-Process "bootrec.exe" -ArgumentList "/rebuildbcd" -Wait -ErrorAction SilentlyContinue } -Name "BootRec_RebuildBCD"
Try-Log -Action { Start-Process "bcdedit.exe" -ArgumentList "/enum" -Wait -ErrorAction SilentlyContinue } -Name "BCDEdit_Enum"
Try-Log -Action { bcdedit /set "{default}" recoveryenabled Yes } -Name "BCDEdit_Recovery"
Try-Log -Action { Start-Process "mdsched.exe" -ArgumentList "/force" -ErrorAction SilentlyContinue } -Name "MemCheck"

# Enable security features
Try-Log -Action { reg add "HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard" /v EnableVirtualizationBasedSecurity /t REG_DWORD /d 1 /f /Force } -Name "MemoryIntegrity"
Try-Log -Action { reg add "HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard" /v RequirePlatformSecurityFeatures /t REG_DWORD /d 1 /f /Force } -Name "MemoryIntegrityPlat"
Try-Log -Action { reg add "HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard" /v HVCIEnabled /t REG_DWORD /d 1 /f /Force } -Name "MemoryIntegrityHVCI"
Try-Log -Action { reg add "HKLM\SYSTEM\CurrentControlSet\Control\Lsa" /v RunAsPPL /t REG_DWORD /d 1 /f /Force } -Name "LSAProtection"

# Remove temp files
Try-Log -Action { Remove-Item "$env:TEMP\*" -Force -Recurse -ErrorAction SilentlyContinue } -Name "ClearTemp"
Try-Log -Action { Remove-Item "$env:Windir\Temp\*" -Force -Recurse -ErrorAction SilentlyContinue } -Name "ClearWindirTemp"

# Restart drivers
Try-Log -Action { Get-WmiObject Win32_PnPSignedDriver | ForEach-Object { Restart-Service $_.ServiceName -Force -ErrorAction SilentlyContinue } } -Name "RestartDrivers"

# General registry/system settings tweaks
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v PublishUserActivities /t REG_DWORD /d 0 /f /Force } -Name "DisableActivityHistory"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v EnableActivityFeed /t REG_DWORD /d 0 /f /Force } -Name "DisableActivityFeed"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v UploadUserActivities /t REG_DWORD /d 0 /f /Force } -Name "DisableUploadActivities"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\CloudContent" /v DisableConsumerFeatures /t REG_DWORD /d 1 /f /Force } -Name "DisableConsumerFeatures"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" /v DisableLocation /t REG_DWORD /d 1 /f /Force } -Name "DisableLocation"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\DataCollection" /v AllowTelemetry /t REG_DWORD /d 0 /f /Force } -Name "DisableTelemetry"
Try-Log -Action { reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v TaskbarEndTask /t REG_DWORD /d 1 /f /Force } -Name "EnableTaskbarEndTask"
Try-Log -Action { reg add "HKCU\Software\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f /Force } -Name "DisableCopilot"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f /Force } -Name "DisableCopilotHKLM"

Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v HideFirstRunExperience /t REG_DWORD /d 1 /f /Force } -Name "EdgeHideFirstRun"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v MetricsReportingEnabled /t REG_DWORD /d 0 /f /Force } -Name "EdgeDisableMetrics"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v BackgroundModeEnabled /t REG_DWORD /d 0 /f /Force } -Name "EdgeDisableBackgroundMode"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v EdgeShoppingAssistantEnabled /t REG_DWORD /d 0 /f /Force } -Name "EdgeDisableShopping"
Try-Log -Action { reg add "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v EdgeCollectionsEnabled /t REG_DWORD /d 0 /f /Force } -Name "EdgeDisableCollections"

Try-Log -Action { reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v HideFileExt /t REG_DWORD /d 0 /f /Force } -Name "ShowFileExtensions"
Try-Log -Action { reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v VerboseStatus /t REG_DWORD /d 1 /f /Force } -Name "ShowVerboseStatus"
Try-Log -Action { reg add "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled /t REG_DWORD /d 0 /f /Force } -Name "DisableFastStartup"
Try-Log -Action { reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" /v ConsentPromptBehaviorAdmin /t REG_DWORD /d 2 /f /Force } -Name "UACAlwaysPrompt"

Try-Log -Action { Disable-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2 -NoRestart -Force -ErrorAction SilentlyContinue } -Name "DisablePowerShell2"
Try-Log -Action { Disable-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2Root -NoRestart -Force -ErrorAction SilentlyContinue } -Name "DisablePowerShell2Root"
Try-Log -Action { Disable-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellISE -NoRestart -Force -ErrorAction SilentlyContinue } -Name "DisablePowerShellISE"

$sysDirs = @("$env:SystemRoot\System32\WindowsPowerShell\v1.0", "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0", "$env:ProgramFiles\PowerShell\7")
$blockedPatterns = @("powershell.exe","pwsh.exe","powershell_ise.exe")
$searchPaths = @("$env:SystemDrive\","$env:LocalAppData\", "$env:ProgramFiles\","$env:ProgramFiles(x86)\")
foreach ($pattern in $blockedPatterns) {
    foreach ($path in $searchPaths) {
        try {
            Get-ChildItem -Path $path -Recurse -Force -Include $pattern -ErrorAction SilentlyContinue | ForEach-Object {
                if ($sysDirs -notcontains $_.DirectoryName) {
                    try { Rename-Item -Path $_.FullName -NewName "$($_.Name).blocked" -Force -ErrorAction SilentlyContinue } catch {}
                }
            }
        } catch {}
    }
}

Try-Log -Action { Set-Service -Name "WinDefend" -StartupType Automatic -Force -ErrorAction SilentlyContinue } -Name "EnableDefenderService"
Try-Log -Action { Start-Service -Name "WinDefend" -Force -ErrorAction SilentlyContinue } -Name "StartDefenderService"
Try-Log -Action { Set-Service -Name "wuauserv" -StartupType Automatic -Force -ErrorAction SilentlyContinue } -Name "EnableWindowsUpdate"
Try-Log -Action { Start-Service -Name "wuauserv" -Force -ErrorAction SilentlyContinue } -Name "StartWindowsUpdate"

Try-Log -Action { reg add "HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" /v FlightSettingsEnable/OtherMicrosoftProducts /t REG_DWORD /d 1 /f /Force } -Name "UpdateOtherProducts"
Try-Log -Action { reg add "HKLM\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" /v ReceiveUpdatesForOtherProducts /t REG_DWORD /d 1 /f /Force } -Name "ReceiveUpdatesOtherProducts"
Try-Log -Action { reg add "HKLM\SOFTWARE\Microsoft\Windows\Windows Troubleshooter" /v DiagnosticAutoRun /t REG_DWORD /d 1 /f /Force } -Name "TroubleshooterAutoRun"

# Winget updates and windows update scan
$internet = $false
try { $internet = Test-Connection -ComputerName "8.8.8.8" -Count 2 -Quiet -ErrorAction SilentlyContinue } catch {}
if ($internet) {
    Try-Log -Action { winget upgrade --all --force | Out-Null } -Name "WingetUpgrade"
    Try-Log -Action {
        Install-Module -Name PSWindowsUpdate -Force -ErrorAction SilentlyContinue
        Import-Module PSWindowsUpdate -ErrorAction SilentlyContinue
        Get-WindowsUpdate -Install -AcceptAll -AutoReboot -ErrorAction SilentlyContinue
    } -Name "PSWindowsUpdate"
}

Parallel-Wait
Save-UndoManifest
Upload-LogsOnce -Endpoint $UploadEndpoint
Schedule-DefenderOffline
Save-UndoManifest
try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}

if ($Global:RebootRequired -or $AutoReboot) { Restart-Computer -Force }
if ($Rollback) { Rollback-From-Manifest -ManifestFile $Global:UndoManifestPath }
exit 0