[CmdletBinding()]
param(
  [string]$LogPath = "$env:TEMP\Apply-Latest-Patches-script.log",
  [string]$ARLog = 'C:\Program Files (x86)\ossec-agent\active-response\active-responses.log'
)

$ErrorActionPreference = 'Stop'
$HostName = $env:COMPUTERNAME
$LogMaxKB = 100
$LogKeep = 5

function Write-Log {
  param([string]$Message,[ValidateSet('INFO','WARN','ERROR','DEBUG')]$Level='INFO')
  $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
  $line = "[$ts][$Level] $Message"
  switch ($Level) {
    'ERROR' { Write-Host $line -ForegroundColor Red }
    'WARN'  { Write-Host $line -ForegroundColor Yellow }
    'DEBUG' { if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose')) { Write-Verbose $line } }
    default { Write-Host $line }
  }
  Add-Content -Path $LogPath -Value $line
}

function Rotate-Log {
  if (Test-Path $LogPath -PathType Leaf) {
    if ((Get-Item $LogPath).Length/1KB -gt $LogMaxKB) {
      for ($i = $LogKeep - 1; $i -ge 0; $i--) {
        $old = "$LogPath.$i"; $new = "$LogPath." + ($i + 1)
        if (Test-Path $old) { Rename-Item $old $new -Force }
      }
      Rename-Item $LogPath "$LogPath.1" -Force
    }
  }
}

Rotate-Log
$runStart = Get-Date
Write-Log "=== SCRIPT START : Check & Install Critical Windows Updates ==="

try {
  if (-not (Get-Module -ListAvailable PSWindowsUpdate)) {
    Write-Log "PSWindowsUpdate module not found. Installing..." 'INFO'
    Install-PackageProvider -Name NuGet -Force -ErrorAction Stop
    Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop
  }
  Import-Module PSWindowsUpdate -Force

  Write-Log "Checking for updates on $HostName..." 'INFO'
  $updates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot | Select-Object * | ForEach-Object {
    [PSCustomObject]@{
      guid              = $_.UpdateID
      title             = $_.Title
      kb_article        = ($_.KB | Out-String).Trim()
      categories        = ($_.Categories | Out-String).Trim()
      severity          = $_.MsrcSeverity
      download_sizeMB   = [math]::Round($_.Size / 1MB, 2)
      is_downloaded     = $_.IsDownloaded
      is_installed      = $_.IsInstalled
      publication_date  = $_.LastDeploymentChangeTime
    }
  }

  $checkObj = [pscustomobject]@{
    timestamp    = (Get-Date).ToString('o')
    host         = $HostName
    action       = 'check_critical_updates'
    update_count = $updates.Count
    updates      = $updates
    status       = 'success'
  }
  $checkObj | ConvertTo-Json -Depth 4 -Compress | Out-File -FilePath $ARLog -Append -Encoding ascii -Width 2000
  Write-Log "Found $($updates.Count) updates. Logged check phase." 'INFO'

  Write-Log "Installing all available updates..." 'INFO'
  $installResults = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot -Verbose -ErrorAction Continue |
    ForEach-Object {
      [PSCustomObject]@{
        title   = $_.Title
        kb_article = ($_.KB | Out-String).Trim()
        result  = if ($_.IsInstalled) { 'Installed' } else { 'Failed' }
        reboot  = $_.RebootRequired
      }
    }

  $finalObj = [pscustomobject]@{
    timestamp = (Get-Date).ToString('o')
    host      = $HostName
    action    = 'install_critical_updates'
    installed = $installResults
    status    = 'completed'
  }
  $finalObj | ConvertTo-Json -Depth 4 -Compress | Out-File -FilePath $ARLog -Append -Encoding ascii -Width 2000
  Write-Log "Results JSON logged to $ARLog" 'INFO'

} catch {
  Write-Log $_.Exception.Message 'ERROR'
  $errorObj = [pscustomobject]@{
    timestamp = (Get-Date).ToString('o')
    host      = $HostName
    action    = 'install_critical_updates'
    status    = 'error'
    error     = $_.Exception.Message
  }
  $errorObj | ConvertTo-Json -Compress | Out-File -FilePath $ARLog -Append -Encoding ascii -Width 2000
}
finally {
  $dur = [int]((Get-Date) - $runStart).TotalSeconds
  Write-Log "=== SCRIPT END : duration ${dur}s ==="
}
