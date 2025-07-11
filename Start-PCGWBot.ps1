[CmdletBinding()]
Param (
  [string]$ResultSize = 100,
  [Alias("Limit")]
  $ExcludeUser        = 'AemonyBot',
  $Namespace          = @(
    '' # Main namespace
    'Company'
    'Controller'
    'Emulation'
    'Engine'
    'Glossary'
    'Guide'
    'Series'
    'Store'
  ),

  $Start              = $null, # Timestamp from where to start

  [switch]$Descending,        # defaults to using an ascending order
  [switch]$Force,             # Forces the bot to run regardless of the status of the public toggle
  [switch]$Reset,
  [switch]$Persistent
)

# Maintenance bot
# Run using e.g. .\Invoke-PCGWMaintenanceBot.ps1 -Limit 5000 -Persistent -Verbose

# Stop transcript if there is any currently running.
# Typically only relevant for debug purposes when run in a console host.
Try { Stop-Transcript | Out-Null } Catch [System.InvalidOperationException] { }

# Start script

$ScriptName = ($MyInvocation.MyCommand.Name) -replace '.ps1', ''
Start-Transcript "$ScriptName.log" | Out-Null
Write-Verbose "Transcript started, output file is $ScriptName.log"

# Script variable to indicate the location of the local cache
$script:CacheFilePath = $env:LOCALAPPDATA + '\PowerShell\PCGWMaintenanceBot\cache.json'

# Global configurations
$script:EnablePage         = 'User:AemonyBot/Enabled' # Page to check between each processed page to see if the bot should continue or not.
$script:ProgressPreference = 'SilentlyContinue'       # Suppress progress bar (speeds up Invoke-WebRequest by a ton)

$CacheTemp = $null
$Cache     = [ordered]@{
  RecentChangesID = 0  # Used to keep track of last processed change ID
  Timestamp       = '' # If used as 'rcstart' with -Ascending ('rcdir: newer'), this can be used to continue where the bot last left
  Output          = @()
}

# Reset the cache
If ($Reset)
{
  If ((Test-Path $script:CacheFilePath) -eq $true)
  { Remove-Item $script:CacheFilePath }
}


# Read the cache
If ($Persistent)
{
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

  # Create the file first using New-Item with -Force parameter so missing directories are also created.
  else
  { New-Item -Path $script:CacheFilePath -ItemType "file" -Force | Out-Null }
}

If ([string]::IsNullOrEmpty($Start) -eq $false)
{ $Cache.Timestamp = $Start }

# Need to declare this again in case reading the cache overwrote it...
$Cache.Output = @()

$Module  = $false
$Session = $false

if (-not (Get-Module MediaWiki))
{
  Import-Module   ..\MediaWiki
  $Module = $true
}

if (-not (Get-MWSession))
{
  Connect-MWSession -Persistent
  $Session = $true
}

# A publicly accessible toggle for the bot
$Status = Get-MWPage -Name $script:EnablePage -Wikitext

if ($Force -or $Status.Wikitext -eq '1')
{
  $MWRecentChangesParameters = @{
    ExcludeUser    = $ExcludeUser
    ResultSize     = $ResultSize
    Namespace      = $Namespace
    Type           = @('edit', 'new')
  }

  if (-not $Descending)
  { $MWRecentChangesParameters.Ascending = $true }
  else
  { $MWRecentChangesParameters.Descending = $true }

  if (-not $AllRevision)
  { $MWRecentChangesParameters.LatestRevision = $true }

  If ([string]::IsNullOrEmpty($Cache.Timestamp) -eq $false)
  { $MWRecentChangesParameters.Start = $Cache.Timestamp }

  $RecentChanges = Get-MWRecentChanges @MWRecentChangesParameters

  ForEach ($Change in $RecentChanges)
  {
    $Output = $null
    $Output = (.\Invoke-PCGWMaintenance.ps1 -ID $Change.ID)
    if ($null -ne $Output)
    { $Cache.Output += $Output }

    # Used to keep track of last processed change across runs
    $Cache.RecentChangesID = $Change.RecentChangesID
    $Cache.Timestamp       = $Change.Timestamp

    # Update the local cache after each page so we can abort at any moment without losing progress
    # Only cache the RecentChangesID and Timestamp values
    If ($Persistent)
    { $Cache | Select-Object RecentChangesID, Timestamp | ConvertTo-Json | Out-File $script:CacheFilePath }
    
    # Check the status after each processed page
    $Status = Get-MWPage -Name $script:EnablePage -Wikitext
    if ($Force -eq $false -and $Status.Wikitext -ne '1')
    {
      # .Timestamp is only accessible when doing an additional API request with the -Information switch, so use .Retrieved instead
      Write-Warning "Bot has aborted as per the contents of $($Status.Name), retrieved $($Status.Retrieved)."
      break
    }
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
