[CmdletBinding()]
Param (
  $HookUrl = '',
  [switch]$Reset
)

# Maintenance bot
# Run using e.g. .\Send-DiscordStatusNotifications.ps1 -Verbose

# Stop transcript if there is any currently running.
# Typically only relevant for debug purposes when run in a console host.
Try { Stop-Transcript | Out-Null } Catch [System.InvalidOperationException] { }

# Start script

$ScriptName = ($MyInvocation.MyCommand.Name) -replace '.ps1', ''
Start-Transcript "$ScriptName.log" | Out-Null
Write-Verbose "Transcript started, output file is $ScriptName.log"

# Script variable to indicate the location of the local cache
$script:CacheFilePath = $env:LOCALAPPDATA + '\PowerShell\PCGWMaintenanceBot\discord_sn.json'

# Global configurations
$script:ProgressPreference = 'SilentlyContinue'              # Suppress progress bar (speeds up Invoke-WebRequest by a ton)

$CacheTemp  = $null
$Cache      = [ordered]@{
  Timestamp = '' # Last updated
  Output    = @()
}

# Reset the cache
If ($Reset)
{
  If ((Test-Path $script:CacheFilePath) -eq $true)
  { Remove-Item $script:CacheFilePath }
}


# Read the persistent cache
if ((Test-Path $script:CacheFilePath) -eq $true)
{
  Write-Warning "Using data from last successful run. Use -Reset to recreate or bypass the stored data."
  Try
  {
    # Try to load the cache.
    $CacheTemp = Get-Content $script:CacheFilePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    $Cache     = $CacheTemp
  }
  Catch [System.Management.Automation.ItemNotFoundException], [System.ArgumentException] {
    # Handle corrupt cache
    Write-Warning "The local cache could not be found or was corrupt.`n"
    $CacheTemp = $null
  }
  Catch
  {
    # Unknown exception
    Write-Warning "Unknown error occurred when trying to read the local cache."
    $CacheTemp = $null
  }
}

# Create the persistent cache using New-Item with -Force parameter so missing directories are also created.
else
{ New-Item -Path $script:CacheFilePath -ItemType "file" -Force | Out-Null }

# If something is wrong with the cache, reset it
if ($null -eq $Cache.Timestamp)
{
  $Cache      = [ordered]@{
    Timestamp = '' # Last updated
    Output    = @()
  }
}

# If we have no cached timestamp, default to the last 24 hours
If ([string]::IsNullOrEmpty($Cache.Timestamp))
{ $Cache.Timestamp = Get-Date ((Get-Date).AddHours(-24).ToUniversalTime()) -UFormat '+%Y-%m-%dT%H:%M:%SZ' }

# Need to declare this again in case reading the cache overwrote it...
$Cache | Add-Member -MemberType NoteProperty -Name Output -Value @()

# Status Page
$Parameters       = @{
  Uri             = 'https://status.pcgamingwiki.com/history.atom'
  UseBasicParsing = $true
  Method          = 'GET'
}

# Discord
$Body = [PSCustomObject]@{
  username   = 'PCGamingWiki'
  avatar_url = 'https://thumbnails.pcgamingwiki.com/9/93/PCGamingWiki_Favicon.svg/240px-PCGamingWiki_Favicon.svg.png'
  content    = ''
}

try
{
  $Response   = Invoke-WebRequest @Parameters

  if ($Response.StatusCode -ne 200)
  {
    throw 'HTTP status code != 200 !'
  }

  $StatusPage = [xml]($Response).Content

  # Only do something if there is actually something to do...
  if ($StatusPage.feed.updated -ne $Cache.Timestamp)
  {
    ForEach ($Entry in $StatusPage.feed.entry)
    {
      # Find the entry that corresponds to the latest updated entry
      if ($Entry.updated -ne $StatusPage.feed.updated)
      { continue }

      $ResolutionTime = ($Entry.category | Where-Object { $_.term -eq 'event:end' }).label
      $IsServiceDown  = [string]::IsNullOrWhiteSpace($ResolutionTime)

      $UnixSeconds    = ([DateTimeOffset]$Entry.updated).ToUnixTimeSeconds()

      $NewContent  = "<t:$UnixSeconds> . . **[$($Entry.title)]($($Entry.link.href))** - "

      if ($IsServiceDown)
      {
        $NewContent += "**New**: Service has become unresponsive."
      } else {
        $NewContent += "**Resolved**: Service is back up."
      }

      $Body.content += $NewContent
    }

    # Used to keep track of last processed change across runs
    $Cache.Timestamp = $StatusPage.feed.updated

    # Update the local cache after each page so we can abort at any moment without losing progress
    $Cache | Select-Object Timestamp | ConvertTo-Json | Out-File $script:CacheFilePath

    if ($HookUrl -and (-not [string]::IsNullOrWhiteSpace($Body.content)))
    {
      $Output = $null
      $Output = Invoke-RestMethod -Method POST -ContentType 'application/json; charset=utf-8' -Body ($Body | ConvertTo-Json) -Uri $HookUrl
      if ($null -ne $Output)
      { $Cache.Output += $Output }
    } else {
      Write-Warning 'No new updates were found since last check.'
    }
  }
} catch {
  throw $_
}

Stop-Transcript | Out-Null

return $Cache
