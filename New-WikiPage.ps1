[CmdletBinding(DefaultParameterSetName = '')]
Param (
  <#
    Generic
  #>
  [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Generic', Position=0)]
  [ValidateNotNullOrEmpty()]
  [string]$Name,

  [Parameter(Mandatory, ParameterSetName = 'Generic')]
  [ValidateSet('Singleplayer', 'Multiplayer', 'Unknown')]
  [string]$Mode,

  [Parameter(Mandatory, ParameterSetName = 'Generic')]
  [ValidateNotNullOrEmpty()]
  [string[]]$Developers,

  [Parameter(ParameterSetName = 'Generic')]
  [ValidateNotNullOrEmpty()]
  [string[]]$Publishers,

  [Parameter(ParameterSetName = 'Generic')]
  [ValidateNotNullOrEmpty()]
  [string]$ReleaseDateWindows,

  [Parameter(ParameterSetName = 'Generic')]
  [ValidateNotNullOrEmpty()]
  [string]$ReleaseDateLinux,

  [Parameter(ParameterSetName = 'Generic')]
  [ValidateNotNullOrEmpty()]
  [string]$ReleaseDateMacOS,

  <#
    Steam ID
  #>
  [Parameter(Mandatory, ParameterSetName = 'SteamId')]
  [int]$SteamAppId,

  <#
    Steam URL
  #>
  [Parameter(Mandatory, ParameterSetName = 'SteamUrl')]
  [string]$SteamUrl,

  <#
    Generic flags
  #>
  # License / commercialization
  [switch]$Commercial, # default; not used
  [switch]$Freeware,
  [Alias('F2P')]
  [switch]$FreeToPlay,
  [switch]$Shareware,

  # Misc
  [switch]$NoRetail,
  [switch]$NoDLCs,
  [switch]$NoWindows,

  <#
    Debug
  #>
  [switch]$WhatIf,
  [string]$TargetPage
)

Begin {
  # Configuration
  $ProgressPreference = 'SilentlyContinue' # Suppress progress bar (speeds up Invoke-WebRequest by a ton)

  $Templates = @{
    Singleplayer = 'PCGamingWiki:Sample article/Game (singleplayer)'
    Multiplayer  = 'PCGamingWiki:Sample article/Game (multiplayer)'
    Unknown      = 'PCGamingWiki:Sample article/Game (unknown)'
  }
}

Process
{
  $Singleplayer = $false
  $Multiplayer  = $false

  if ($NoWindows)
  {
    if ($ReleaseDateWindows)
    {
      Write-Warning '-NoWindows and -ReleaseDateWindows cannot be used at the same time!'
      return
    }

    if (-not $ReleaseDateLinux -and -not $ReleaseDateMacOS)
    {
      Write-Warning 'A Linux or macOS release date needs to be specified when using -NoWindows!'
      return
    }
  }

  $SteamData  = $null
  
  Write-Verbose $SteamAppId

  if (-not [string]::IsNullOrWhiteSpace($SteamUrl))
  { $SteamAppId = ($SteamUrl -replace '^([^\d]+\/app\/)(\d+)(\/?.*)', '$2') }

  # Extract Steam app info
  if ($SteamAppId -ne 0)
  {
    $Link = "https://store.steampowered.com/api/appdetails/?appids=$SteamAppId"
    try {
      Write-Verbose "Retrieving $Link"
      $WebPage    = Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive
      $StatusCode = $WebPage.StatusCode
    } catch {
      $StatusCode = $_.Exception.response.StatusCode.value__
    }

    if ($StatusCode -eq 200)
    {
      $Json = ConvertFrom-Json $WebPage.Content

      if ($Json.$SteamAppId.success -ne 'true')
      {
        Write-Warning 'Failed to parse Json from Steam!'
        return
      }
      else
      {
        $SteamData = $Json.$SteamAppId.data

        $Type = $SteamData.type

        if ($Type -ne 'game')
        {
          Write-Warning "$SteamAppId is not a game!"
          return
        }

        $Name = $SteamData.name

            if ($SteamData.categories | Where-Object { $_.description -eq 'Multi-player' })
        { $Mode = 'Multiplayer' }
        elseif ($SteamData.categories | Where-Object { $_.description -eq 'Single-player' })
        { $Mode = 'Singleplayer' }
        else
        { $Mode = 'Unknown' }

        $Developers = $SteamData.developers
        $Pubs = $SteamData.publishers | Where-Object { $Developers -notcontains $_ }
        if ($null -ne $Pubs)
        { $Publishers = $Pubs }
      }
    }
  }

  if ([string]::IsNullOrWhiteSpace($Developers))
  {
    Write-Warning 'A developer needs to be specified to continue!'
    return
  }

  $Developers = $Developers -replace '(?:,?\s|,)(?:Inc|Ltd|GmbH|S\.?A|LLC|V\.?O\.?F|AB)\.?$', ''
  $Publishers = $Publishers -replace '(?:,?\s|,)(?:Inc|Ltd|GmbH|S\.?A|LLC|V\.?O\.?F|AB)\.?$', ''

  $Singleplayer = ($Mode -eq 'Singleplayer')
  $Multiplayer  = ($Mode -eq 'Multiplayer')

  $Template = Get-MWPage -PageName $Templates[$Mode] -Wikitext

  # Game Name
  $Template.Wikitext = $Template.Wikitext.Replace('GAME TITLE', $Name)

  # Platforms
  $Platforms = @()

  # Steam stuff
  if ($SteamData)
  {
    $ReleaseDate = 'TBA'

    # EA
        if ($SteamData.genres.description -contains 'Early Access')
    { $ReleaseDate = 'EA' }

    # TBA
    elseif ($SteamData.release_date.date -ne '' -and
            $SteamData.release_date.date -ne 'Coming soon')
    { $ReleaseDate = $SteamData.release_date.date }

    if ($SteamData.platforms.windows -eq 'true')
    { $ReleaseDateWindows = $ReleaseDate }
    if ($SteamData.platforms.mac -eq 'true')
    { $ReleaseDateMacOS   = $ReleaseDate }
    if ($SteamData.platforms.linux -eq 'true')
    { $ReleaseDateLinux   = $ReleaseDate }
  }

  if ($ReleaseDateWindows)
  { $Platforms += 'Windows' }

  if ($ReleaseDateMacOS)
  { $Platforms += 'macOS' }

  if ($ReleaseDateLinux)
  { $Platforms += 'Linux' }

  if ($Platforms -notcontains 'Windows')
  { $NoWindows = $true }

  # Series
  $Series = ''
  $Template.Wikitext = $Template.Wikitext.Replace('PCGW Templates<!-- CHANGE TO THE ACTUAL SERIES NAME IF ONE EXISTS -->', $Series)

  # Developer
  $Template.Wikitext = $Template.Wikitext.Replace('DEVELOPER', $Developers[0])

  if ($Developers.Count -gt 1)
  {
    foreach ($Developer in $Developers)
    { $Template.Wikitext = $Template.Wikitext.Replace('|publishers   = ', "{{Infobox game/row/developer|$Developer}}`n|publishers   = ") }
  }

  # Publisher
  if (-not [string]::IsNullOrWhiteSpace($Publishers))
  {
    $Template.Wikitext = $Template.Wikitext.Replace('PUBLISHER', $Publishers[0])

    if ($Publishers.Count -gt 1)
    {
      foreach ($Publisher in $Publishers)
      { $Template.Wikitext = $Template.Wikitext.Replace('|publishers   = ', "{{Infobox game/row/developer|$Publisher}}`n|publishers   = ") }
    }
  } else {
    $Template.Wikitext = $Template.Wikitext.Replace("{{Infobox game/row/publisher|PUBLISHER}}`n", '')
  }

  # Release Date: Windows
  if (-not [string]::IsNullOrWhiteSpace($ReleaseDateWindows))
  {
    try {
      $DateTime = [datetime]::Parse($ReleaseDateWindows)
      $Template.Wikitext = $Template.Wikitext.Replace('{{Infobox game/row/date|Windows|TBA}}', "{{Infobox game/row/date|Windows|$($DateTime.ToString('MMMM d, yyyy', [CultureInfo]("en-US")))}}")
    } catch {
      $Template.Wikitext = $Template.Wikitext.Replace('{{Infobox game/row/date|Windows|TBA}}', "{{Infobox game/row/date|Windows|$ReleaseDateWindows}}")
    }
  }

  # Release Date: macOS
  if (-not [string]::IsNullOrWhiteSpace($ReleaseDateMacOS))
  {
    try {
      $DateTime = [datetime]::Parse($ReleaseDateMacOS)
      $Template.Wikitext = $Template.Wikitext.Replace('|reception    = ', "{{Infobox game/row/date|macOS|$($DateTime.ToString('MMMM d, yyyy', [CultureInfo]("en-US")))}}`n|reception    = ")
    } catch {
      $Template.Wikitext = $Template.Wikitext.Replace('|reception    = ', "{{Infobox game/row/date|macOS|$ReleaseDateMacOS}}`n|reception    = ")
    }
  }

  # Release Date: Linux
  if (-not [string]::IsNullOrWhiteSpace($ReleaseDateLinux))
  {
    try {
      $DateTime = [datetime]::Parse($ReleaseDateLinux)
      $Template.Wikitext = $Template.Wikitext.Replace('|reception    = ', "{{Infobox game/row/date|Linux|$($DateTime.ToString('MMMM d, yyyy', [CultureInfo]("en-US")))}}`n|reception    = ")
    } catch {
      $Template.Wikitext = $Template.Wikitext.Replace('|reception    = ', "{{Infobox game/row/date|Linux|$ReleaseDateLinux}}`n|reception    = ")
    }
  }

  # Website
  if ($SteamData.website)
  { $Template.Wikitext = $Template.Wikitext.Replace('|official site= ', ('|official site= ' + $SteamData.website)) }

  # No Retail
  if ($NoRetail)
  { $Template.Wikitext = $Template.Wikitext.Replace("{{Availability/row| retail | | unknown |  |  | Windows }}`n", '') }
  # Steam
  elseif ($SteamData)
  { $Template.Wikitext = $Template.Wikitext.Replace('{{Availability/row| retail | | unknown |  |  | Windows }}', "{{Availability/row| steam | $SteamAppId | steam |  |  | $($Platforms -join ', ') }}") }
  # Retail
  else
  { $Template.Wikitext = $Template.Wikitext.Replace('{{Availability/row| retail | | unknown |  |  | Windows }}', "{{Availability/row| retail | | unknown |  |  | $($Platforms -join ', ') }}") }

  # No Windows
  if ($NoWindows)
  {
    # Remove release date
    $Template.Wikitext = $Template.Wikitext.Replace("{{Infobox game/row/date|Windows|TBA}}`n", '')

    # Remove game data
    $Template.Wikitext = $Template.Wikitext.Replace("{{Game data/config|Windows|}}`n", '')
    $Template.Wikitext = $Template.Wikitext.Replace("{{Game data/saves|Windows|}}`n", '')

    # Replace Windows in the system requirements with the first supported OS
    $Template.Wikitext = $Template.Wikitext.Replace('|OSfamily = Windows', "|OSfamily = $($Platforms[0])")
  }

  # Freeware
  if ($Freeware)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext.Replace('{{Infobox game/row/taxonomy/monetization      | One-time game purchase }}', '{{Infobox game/row/taxonomy/monetization      | Freeware }}')
    $Template.Wikitext = $Template.Wikitext.Replace('|license      = ', '|license      = freeware')

    # Monetization table
    $Template.Wikitext = $Template.Wikitext.Replace('|freeware                    = ',                                                 '|freeware                    = Game is freeware.')
    $Template.Wikitext = $Template.Wikitext.Replace('|one-time game purchase      = The game requires an upfront purchase to access.', '|one-time game purchase      = ')
  }

  # Free-to-Play
  if ($FreeToPlay)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext.Replace('{{Infobox game/row/taxonomy/monetization      | One-time game purchase }}', '{{Infobox game/row/taxonomy/monetization      | Free-to-play }}')
    $Template.Wikitext = $Template.Wikitext.Replace('|license      = ', '|license      = free-to-play')

    # Monetization table
    $Template.Wikitext = $Template.Wikitext.Replace('|free-to-play                = ',                                                 '|free-to-play                = Game is free-to-play.')
    $Template.Wikitext = $Template.Wikitext.Replace('|one-time game purchase      = The game requires an upfront purchase to access.', '|one-time game purchase      = ')
  }

  # Shareware
  if ($Shareware)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext.Replace('|license      = ', '|license      = shareware')
    # Assume a shareware title is a one-time purchase as well
  }

  # No DLCs
  if ($NoDLCs)
  { $Template.Wikitext = $Template.Wikitext.Replace("`n{{DLC|`n<!-- DLC rows goes below: -->`n`n}}`n", '') }

  elseif ($SteamData.dlc)
  {
    # Retrieve info about the DLC using a separate request...

  }

  if ($SteamData)
  {
    $Template.Wikitext = $Template.Wikitext.Replace('==Availability==', @"
'''General information'''
{{mm}} [http://steamcommunity.com/app/$SteamAppId/discussions/ Steam Community Discussions]

==Availability==
"@)
  }

  # Game data
  $GameDataConfig = ''
  $GameDataSaves = ''
  foreach ($Platform in $Platforms)
  { $GameData += "{{Game data/config|$Platform|}}`n" }

  #$Template.Wikitext = $Template.Wikitext -replace ''

  # Steam Cloud
  if ($SteamData.categories.description -contains 'Steam Cloud')
  { $Template.Wikitext = $Template.Wikitext.Replace('|steam cloud               = ', '|steam cloud               = true') }

  # Controller support
  if ($SteamData.controller_support -eq 'full')
  {
    $Template.Wikitext = $Template.Wikitext.Replace('|controller support        = unknown', '|controller support        = true')
    $Template.Wikitext = $Template.Wikitext.Replace('|full controller           = unknown', '|full controller           = true')
  }


  # Create page
  if ([string]::IsNullOrWhiteSpace($TargetPage))
  { $TargetPage = $Name }

  if ($WhatIf)
  {
    Write-Host ('What if: Performing maintenance on target "' + $TargetPage + '".')
    return $Template.Wikitext
  } else {
    return Set-MWPage -Name $TargetPage -Summary 'Created page' -Major -CreateOnly -Content $Template.Wikitext
  }
}

End { }