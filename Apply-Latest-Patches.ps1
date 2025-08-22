[CmdletBinding()]
param(
  [string]$LogPath = "$env:TEMP\Apply-Latest-Patches-script.log",
  [string]$ARLog  = 'C:\Program Files (x86)\ossec-agent\active-response\active-responses.log'
)

$ErrorActionPreference='Stop'
$HostName=$env:COMPUTERNAME
$LogMaxKB=100
$LogKeep=5
$runStart=Get-Date

function Write-Log {
  param([string]$Message,[ValidateSet('INFO','WARN','ERROR','DEBUG')]$Level='INFO')
  $ts=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
  $line="[$ts][$Level] $Message"
  switch($Level){
    'ERROR'{Write-Host $line -ForegroundColor Red}
    'WARN'{Write-Host $line -ForegroundColor Yellow}
    'DEBUG'{if($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose')){Write-Verbose $line}}
    default{Write-Host $line}
  }
  Add-Content -Path $LogPath -Value $line
}

function Rotate-Log {
  if(Test-Path $LogPath -PathType Leaf){
    if((Get-Item $LogPath).Length/1KB -gt $LogMaxKB){
      for($i=$LogKeep-1;$i -ge 0;$i--){
        $old="$LogPath.$i";$new="$LogPath."+($i+1)
        if(Test-Path $old){Rename-Item $old $new -Force}
      }
      Rename-Item $LogPath "$LogPath.1" -Force
    }
  }
}

Rotate-Log
Write-Log "=== SCRIPT START : Check & Install Critical Windows Updates ==="

try{
  if(-not (Get-Module -ListAvailable PSWindowsUpdate)){
    Write-Log "PSWindowsUpdate module not found. Installing..." 'INFO'
    Install-PackageProvider -Name NuGet -Force -ErrorAction Stop
    Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop
  }
  Import-Module PSWindowsUpdate -Force

  Write-Log "Checking for updates on $HostName..." 'INFO'
  $updates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot |
    ForEach-Object{
      [pscustomobject]@{
        guid=$_.UpdateID
        title=$_.Title
        kb_article= ($_.KB | Out-String).Trim()
        categories= ($_.Categories | Out-String).Trim()
        severity=$_.MsrcSeverity
        download_sizeMB=[math]::Round(($_.Size/1MB),2)
        is_downloaded=$_.IsDownloaded
        is_installed=$_.IsInstalled
        publication_date=$_.LastDeploymentChangeTime
      }
    }

  $lines=@()
  $ts=(Get-Date).ToString('o')

  $lines += ([pscustomobject]@{
    timestamp=$ts
    host=$HostName
    action='check_critical_updates_summary'
    update_count=$updates.Count
    copilot_action=$true
  } | ConvertTo-Json -Compress -Depth 3)

  foreach($u in $updates){
    $lines += ([pscustomobject]@{
      timestamp=(Get-Date).ToString('o')
      host=$HostName
      action='check_critical_updates'
      guid=$u.guid
      title=$u.title
      kb_article=$u.kb_article
      categories=$u.categories
      severity=$u.severity
      download_sizeMB=$u.download_sizeMB
      is_downloaded=$u.is_downloaded
      is_installed=$u.is_installed
      publication_date=$u.publication_date
      copilot_action=$true
    } | ConvertTo-Json -Compress -Depth 4)
  }

  Write-Log "Installing all available updates..." 'INFO'
  $installResults = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot -Verbose -ErrorAction Continue |
    ForEach-Object{
      [pscustomobject]@{
        title=$_.Title
        kb_article= ($_.KB | Out-String).Trim()
        result= (if ($_.IsInstalled) {'Installed'} else {'Failed'})
        reboot=$_.RebootRequired
      }
    }

  $lines += ([pscustomobject]@{
    timestamp=(Get-Date).ToString('o')
    host=$HostName
    action='install_critical_updates_summary'
    attempted= (if ($updates){$updates.Count}else{0})
    results_count= (if ($installResults){$installResults.Count}else{0})
    copilot_action=$true
  } | ConvertTo-Json -Compress -Depth 3)

  foreach($r in $installResults){
    $lines += ([pscustomobject]@{
      timestamp=(Get-Date).ToString('o')
      host=$HostName
      action='install_critical_updates'
      title=$r.title
      kb_article=$r.kb_article
      result=$r.result
      reboot_required=$r.reboot
      copilot_action=$true
    } | ConvertTo-Json -Compress -Depth 3)
  }

  $ndjson=[string]::Join("`n",$lines)
  $tempFile="$env:TEMP\arlog.tmp"
  Set-Content -Path $tempFile -Value $ndjson -Encoding ascii -Force
  $recordCount=$lines.Count
  try{
    Move-Item -Path $tempFile -Destination $ARLog -Force
    Write-Log "Wrote $recordCount NDJSON record(s) to $ARLog" 'INFO'
  }catch{
    Move-Item -Path $tempFile -Destination "$ARLog.new" -Force
    Write-Log "ARLog locked; wrote to $($ARLog).new" 'WARN'
  }
}
catch{
  Write-Log $_.Exception.Message 'ERROR'
  $err=[pscustomobject]@{
    timestamp=(Get-Date).ToString('o')
    host=$HostName
    action='install_critical_updates'
    status='error'
    error=$_.Exception.Message
    copilot_action=$true
  }
  $ndjson=($err | ConvertTo-Json -Compress -Depth 3)
  $tempFile="$env:TEMP\arlog.tmp"
  Set-Content -Path $tempFile -Value $ndjson -Encoding ascii -Force
  try{
    Move-Item -Path $tempFile -Destination $ARLog -Force
    Write-Log "Error JSON written to $ARLog" 'INFO'
  }catch{
    Move-Item -Path $tempFile -Destination "$ARLog.new" -Force
    Write-Log "ARLog locked; wrote error to $($ARLog).new" 'WARN'
  }
}
finally{
  $dur=[int]((Get-Date)-$runStart).TotalSeconds
  Write-Log "=== SCRIPT END : duration ${dur}s ==="
}
