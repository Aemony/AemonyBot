[CmdletBinding()]
Param (
  $HookUrl            = '',
  $WarnHookUrl        = '',
  $Start              = $null, # Timestamp from where to start # (Get-Date).AddMinutes(-5)
  [switch]$Descending,        # defaults to using an ascending order
  [switch]$Force,             # Forces the bot to run regardless of the status of the public toggle
  [switch]$Reset
)

# Maintenance bot
# Run using e.g. .\Send-DiscordMessages.ps1 -Verbose

# Stop transcript if there is any currently running.
# Typically only relevant for debug purposes when run in a console host.
Try { Stop-Transcript | Out-Null } Catch [System.InvalidOperationException] { }

# Start script

$ScriptName = ($MyInvocation.MyCommand.Name) -replace '.ps1', ''
Start-Transcript "$ScriptName.log" | Out-Null
Write-Verbose "Transcript started, output file is $ScriptName.log"

# Script variable to indicate the location of the local cache
$script:CacheFilePath = $env:LOCALAPPDATA + '\PowerShell\PCGWMaintenanceBot\discord_nu.json'

# Global configurations
$script:EnablePage         = 'User:AemonyBot/DiscordEnabled' # Page to check between each processed page to see if the bot should continue or not.
$script:ProgressPreference = 'SilentlyContinue'              # Suppress progress bar (speeds up Invoke-WebRequest by a ton)

$CacheTemp  = $null
$Cache      = [ordered]@{
  LogID     = 0  # Used to keep track of last processed change ID
  Timestamp = '' # If used as 'rcstart' with -Ascending ('rcdir: newer'), this can be used to continue where the bot last left
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

If ([string]::IsNullOrEmpty($Start) -eq $false)
{ $Cache.Timestamp = $Start }

# If -Start is not used and we have no cached timestamp, default to the last 6 hours
If ([string]::IsNullOrEmpty($Cache.Timestamp))
{ $Cache.Timestamp = Get-Date ((Get-Date).AddHours(-6).ToUniversalTime()) -UFormat '+%Y-%m-%dT%H:%M:%SZ' }

# Need to declare this again in case reading the cache overwrote it...
$Cache | Add-Member -MemberType NoteProperty -Name Output -Value @()

$Module  = $false
$Session = $false

if (-not (Get-Module MediaWiki))
{
  Import-Module   ..\MediaWiki
  $Module = $true
}

if (-not (Get-MWSession))
{
  Connect-MWSession -Persistent -Guest
  $Session = $true
}

# A publicly accessible toggle for the bot
$Status = Get-MWPage -Name $script:EnablePage -Wikitext

if ($Force -or $Status.Wikitext -eq '1')
{
  $Parameters = @{
    Type      = 'NewUsers'
    Ascending = $true
    Start     = $Cache.Timestamp
  }

  $NewUsers = Get-MWEventLog @Parameters

  $Body = [PSCustomObject]@{
    username   = 'PCGamingWiki'
    avatar_url = 'https://thumbnails.pcgamingwiki.com/9/93/PCGamingWiki_Favicon.svg/240px-PCGamingWiki_Favicon.svg.png'
    content    = ''
  }

  ForEach ($User in $NewUsers)
  {
    if ($User.LogID -eq $Cache.LogID)
    { continue }

    $Username    = $User.User.Trim('User:')
    $UsernameURI = $Username.Replace(' ', '_')
    $UserPage    = "https://www.pcgamingwiki.com/wiki/User:$UsernameURI"
    $UserTalk    = "https://www.pcgamingwiki.com/wiki/User_talk:$UsernameURI"
    $UserContr   = "https://www.pcgamingwiki.com/wiki/Special:Contributions/$UsernameURI"
    $UserBlock   = "https://www.pcgamingwiki.com/wiki/Special:Block/$UsernameURI"
    $UnixSeconds = ([DateTimeOffset]$User.Timestamp).ToUnixTimeSeconds()

    $NewContent  = "<t:$UnixSeconds> . . **[$Username]($UserPage)** ([talk]($UserTalk) | [contribs]($UserContr) | [block]($UserBlock)) was created."

    if (($Body.content.Length + $NewContent.Length) -ge 2000)
    {
      $Output = $null
      $Output = Invoke-RestMethod -Method POST -ContentType 'application/json' -Body ($Body | ConvertTo-Json) -Uri $HookUrl
      Start-Sleep -Seconds 5
      $Body.content = ''
    }

    if (-not [string]::IsNullOrWhiteSpace($Body.content))
    {
      $Body.content += "`n"
    }

    $Body.content   += $NewContent

    # Used to keep track of last processed change across runs
    $Cache.LogID     = $User.LogID
    $Cache.Timestamp = $User.Timestamp

    # Update the local cache after each page so we can abort at any moment without losing progress
    # Only cache the LogID and Timestamp values
    $Cache | Select-Object LogID, Timestamp | ConvertTo-Json | Out-File $script:CacheFilePath
  }

  if ($HookUrl -and (-not [string]::IsNullOrWhiteSpace($Body.content)))
  {
    $Output = $null
    $Output = Invoke-RestMethod -Method POST -ContentType 'application/json' -Body ($Body | ConvertTo-Json) -Uri $HookUrl
    if ($null -ne $Output)
    { $Cache.Output += $Output }
  }

  # Risk
  $Minutes    = 60
  $Parameters = @{
    Type      = 'NewUsers'
    Ascending = $true
    Start     = (Get-Date ((Get-Date).AddMinutes(-$Minutes).ToUniversalTime()) -UFormat '+%Y-%m-%dT%H:%M:%SZ')
  }

  $RecentUsers = Get-MWEventLog @Parameters

  if ($RecentUsers.Count -gt 5)
  {
    $Body.content  = "## *$($RecentUsers.Count)* users was created on the wiki within the last $Minutes minutes!`n"
    $Body.content += "This could be an indication of an unusual amount of traffic."
    Invoke-RestMethod -Method POST -ContentType 'application/json' -Body ($Body | ConvertTo-Json) -Uri $WarnHookUrl | Out-Null
  }

} else {
  # .Timestamp is only accessible when doing an additional API request with the -Information switch, so use .Retrieved instead
  Write-Warning "Bot has been disabled as per the contents of $($Status.Name), retrieved $($Status.Retrieved)."
}

if ($Session)
{
  Disconnect-MWSession
}

if ($Module)
{
  Remove-Module MediaWiki
}

Stop-Transcript | Out-Null

return $Cache
