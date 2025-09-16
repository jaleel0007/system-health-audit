<# 
.SYNOPSIS
  System Health Audit (Windows)
.OUTPUTS
  - JSON and Markdown in .\reports\
#>

[CmdletBinding()]
param(
  [string]$OutDir = (Join-Path (Split-Path -Parent $PSScriptRoot) "reports")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Ensure reports dir exists
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

$ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
$hostName = $env:COMPUTERNAME
$jsonPath = Join-Path $OutDir "windows_${hostName}_$ts.json"
$mdPath   = Join-Path $OutDir "windows_${hostName}_$ts.md"

# --- Collect OS base info ---
$os = Get-CimInstance Win32_OperatingSystem

# --- UPTIME (robust; no DMTF) ---
$uptime = $null
try {
  # In PS7 this returns a [TimeSpan] directly
  $uptime = Get-Uptime
} catch { $uptime = $null }

if (-not $uptime) {
  try {
    $sec = (Get-CimInstance Win32_PerfFormattedData_PerfOS_System).SystemUpTime
    $uptime = [TimeSpan]::FromSeconds([double]$sec)
  } catch { $uptime = [TimeSpan]::Zero }
}
$uptimeDays  = [int]$uptime.TotalDays
$uptimeHours = [int][Math]::Round($uptime.TotalHours)

# --- CPU ---
$cpuInfo = Get-CimInstance Win32_Processor | Select-Object Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed
$cpuLoad = (Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 3).CounterSamples |
           Measure-Object -Property CookedValue -Average | Select-Object -ExpandProperty Average

# --- Memory ---
$memTotalMB = [math]::Round($os.TotalVisibleMemorySize/1024,0)
$memFreeMB  = [math]::Round($os.FreePhysicalMemory/1024,0)
$memUsedMB  = $memTotalMB - $memFreeMB
$memUsedPct = if ($memTotalMB -ne 0) { [math]::Round(100*$memUsedMB/$memTotalMB,2) } else { 0 }

# --- Disks ---
$disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
  $sizeGB = if ($_.Size) { [math]::Round($_.Size/1GB,2) } else { 0 }
  $freeGB = if ($_.FreeSpace) { [math]::Round($_.FreeSpace/1GB,2) } else { 0 }
  [pscustomobject]@{
    DeviceID = $_.DeviceID
    FileSystem = $_.FileSystem
    SizeGB = $sizeGB
    FreeGB = $freeGB
    UsedPct = if ($sizeGB -ne 0) { [math]::Round(100*($sizeGB-$freeGB)/$sizeGB,2) } else { 0 }
  }
}

# --- Patching (best effort) ---
$recentHotfixes = $null
try {
  $recentHotfixes = Get-HotFix | Sort-Object -Property InstalledOn -Descending | Select-Object -First 10
} catch { $recentHotfixes = @() }
$lastPatchDate  = if ($recentHotfixes) { ($recentHotfixes | Select-Object -First 1).InstalledOn } else { $null }

# --- AV (Defender if present) ---
$av = $null
try {
  $mp = Get-MpComputerStatus
  $av = [pscustomobject]@{
    Product = "Microsoft Defender"
    RealTimeProtection = $mp.RealTimeProtectionEnabled
    SignatureAgeHours = $mp.AntispywareSignatureLastUpdated.TotalHours
    EngineVersion = $mp.AMEngineVersion
  }
} catch {
  $av = [pscustomobject]@{
    Product = "Unknown/3rd-party"
    RealTimeProtection = $null
    SignatureAgeHours = $null
    EngineVersion = $null
  }
}

# --- Assemble result object ---
$result = [pscustomobject]@{
  CollectedAt = (Get-Date)
  Hostname    = $hostName
  OS          = [pscustomobject]@{
    Caption = $os.Caption
    Version = $os.Version
    Build   = $os.BuildNumber
    Uptime  = [pscustomobject]@{
      Days  = $uptimeDays
      Hours = $uptimeHours
    }
  }
  CPU = [pscustomobject]@{
    Model = $cpuInfo.Name -join ', '
    Cores = ($cpuInfo | Measure-Object NumberOfCores -Sum).Sum
    LogicalProcessors = ($cpuInfo | Measure-Object NumberOfLogicalProcessors -Sum).Sum
    AvgLoadPct = [math]::Round($cpuLoad,2)
  }
  Memory = [pscustomobject]@{
    TotalMB = $memTotalMB
    UsedMB  = $memUsedMB
    UsedPct = $memUsedPct
  }
  Disks = $disks
  Patching = [pscustomobject]@{
    LastHotfixInstalledOn = $lastPatchDate
    RecentHotfixes = $recentHotfixes | Select-Object HotFixID, Description, InstalledOn
  }
  Antivirus = $av
}

# --- Write JSON ---
$result | ConvertTo-Json -Depth 6 | Out-File -Encoding UTF8 $jsonPath

# --- Write Markdown ---
$diskTable = ($disks | ForEach-Object {
  "| $($_.DeviceID) | $($_.FileSystem) | $($_.SizeGB) | $($_.FreeGB) | $($_.UsedPct)% |"
}) -join "`n"

$md = @"
# Windows System Health Report — $hostName
**Collected:** $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))

## OS
- **Name:** $($result.OS.Caption)
- **Version/Build:** $($result.OS.Version) / $($result.OS.Build)
- **Uptime:** $($result.OS.Uptime.Days)d ($($result.OS.Uptime.Hours)h)

## CPU
- **Model:** $($result.CPU.Model)
- **Cores/Threads:** $($result.CPU.Cores) / $($result.CPU.LogicalProcessors)
- **Avg Load:** $($result.CPU.AvgLoadPct)%

## Memory
- **Total:** $($result.Memory.TotalMB) MB
- **Used:** $($result.Memory.UsedMB) MB ($($result.Memory.UsedPct)%)

## Disks
| Drive | FS | Size(GB) | Free(GB) | Used% |
|---|---|---:|---:|---:|
$diskTable

## Patching
- **Last Hotfix Installed:** $($result.Patching.LastHotfixInstalledOn)
- **Recent Hotfixes:** $(@($result.Patching.RecentHotfixes).Count)

## Antivirus
- **Product:** $($result.Antivirus.Product)
- **Real-time:** $($result.Antivirus.RealTimeProtection)
- **Signature Age (hrs):** $($result.Antivirus.SignatureAgeHours)
- **Engine:** $($result.Antivirus.EngineVersion)

> JSON: `$jsonPath`
"@

$md | Out-File -Encoding UTF8 $mdPath

Write-Host "Wrote:"
Write-Host " - $jsonPath"
Write-Host " - $mdPath"
