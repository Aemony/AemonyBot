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
  function RegexEscape($UnescapedString)
  { return [regex]::Escape($UnescapedString).Replace('/', '\/') }

  function SetTemplate
  {
    param (
      [Parameter(Mandatory, Position=0)]
      [string]$Template,

      [Parameter(Mandatory, Position=1)]
      [AllowEmptyString()]
      [string]$Value,

      [Parameter(Mandatory, ValueFromPipeline)]
      [string]$String
    )
    process
    {
      if (-not [string]::IsNullOrWhiteSpace($Value))
      { $Value += ' ' }

      return ($String -replace ('(?m)^\{\{(' + (RegexEscape($Template)) + '\s*)\|.*\}\}$'), "{{`$1| $Value}}")
    }
  }

  function SetParameter
  {
    param (
      [Parameter(Mandatory, Position=0)]
      [string]$Parameter,

      [Parameter(Mandatory, Position=1)]
      [AllowEmptyString()]
      [string]$Value,

      [Parameter(Mandatory, ValueFromPipeline)]
      [string]$String
    )
    process { return $String -replace ('(?m)^(\|' + (RegexEscape($Parameter)) + '\s*=).*$'), "`$1 $Value" }
  }

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

  $SteamData      = $null
  $SteamStorePage = $null # ComObject: HTMLFile
  
  Write-Verbose $SteamAppId

  if (-not [string]::IsNullOrWhiteSpace($SteamUrl))
  { $SteamAppId = ($SteamUrl -replace '(?m)^([^\d]+\/app\/)(\d+)(\/?.*)', '$2') }

  # Extract information from Steam
  if ($SteamAppId -ne 0)
  {
    # Steam Store API
    $Link = "https://store.steampowered.com/api/appdetails/?appids=$SteamAppId&l=english"
    try {
      Write-Verbose "Retrieving $Link"
      $WebPage    = Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive
      $StatusCode = $WebPage.StatusCode
    } catch {
      $StatusCode = $_.Exception.response.StatusCode.value__
    }
    
    if ($StatusCode -ne 200)
    {
      Write-Warning 'Failed to retrieve app details from Steam!'
      return
    }

    $Json = ConvertFrom-Json $WebPage.Content

    if ($Json.$SteamAppId.success -ne 'true')
    {
      Write-Warning 'Failed to parse Json from Steam!'
      return
    }

    $SteamData = $Json.$SteamAppId.data
    $Type = $SteamData.type

    if ($Type -ne 'game')
    {
      Write-Warning "$SteamAppId is not a game!"
      return
    }

    # Steam Store
    $Link = "https://store.steampowered.com/app/$SteamAppId/&l=english"
    try {
      Write-Verbose "Retrieving $Link"

      $UAGoogleBot = 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'

      $Session      = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
      $CookiesAge   = [System.Net.Cookie]::new('birthtime', '0')
      $CookiesAdult = [System.Net.Cookie]::new('mature_content', '1')
      $Session.Cookies.Add('https://store.steampowered.com/', $CookiesAge)
      $Session.Cookies.Add('https://store.steampowered.com/', $CookiesAdult)

      $WebPage    = Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive -UserAgent $UAGoogleBot -WebSession $Session
      $StatusCode = $WebPage.StatusCode
    } catch {
      $StatusCode = $_.Exception.response.StatusCode.value__
    }
    
    if ($StatusCode -ne 200)
    {
      Write-Warning 'Failed to retrieve store page from Steam!'
      return
    }

    $SteamStorePage = New-Object -ComObject "HTMLFile"
    [string]$Body = $WebPage.Content
    $SteamStorePage.Write([ref]$Body)

    # Initial stuff

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

  if ([string]::IsNullOrWhiteSpace($Developers))
  {
    Write-Warning 'A developer needs to be specified to continue!'
    return
  }

  $Developers = $Developers -replace '(?:,?\s|,)(?:Inc|Ltd|GmbH|S\.?A|LLC|V\.?O\.?F|AB)\.?$', ''
  if ($Publishers)
  { $Publishers = $Publishers -replace '(?:,?\s|,)(?:Inc|Ltd|GmbH|S\.?A|LLC|V\.?O\.?F|AB)\.?$', '' }

  $Template = Get-MWPage -PageName $Templates[$Mode] -Wikitext

  # Game Name
  $Name = $Name.Replace('™', '')
  $Name = $Name.Replace('®', '')
  $Name = $Name.Replace('©', '')
  $Name = $Name.Replace(': ', ' - ')
  $Name = $Name.Replace(':', '')
  $Template.Wikitext = $Template.Wikitext.Replace('GAME TITLE', $Name)

  # Platforms
  $Platforms = @()

  # Steam stuff
  if ($SteamData)
  {
    $Template.Wikitext = $Template.Wikitext | SetParameter 'steam appid' -Value "$SteamAppId"

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

    # Taxonomy
    <#
      {{Infobox game/row/taxonomy/pacing            | }}
      {{Infobox game/row/taxonomy/perspectives      | }}
      {{Infobox game/row/taxonomy/controls          | }}
      {{Infobox game/row/taxonomy/genres            | }}
      {{Infobox game/row/taxonomy/sports            | }}
      {{Infobox game/row/taxonomy/vehicles          | }}
      {{Infobox game/row/taxonomy/art styles        | }}
      {{Infobox game/row/taxonomy/themes            | }}
    #>

    $Taxonomy = @{
      modes        = (Get-MWCategoryMember 'Modes'                 -Type 'subcat').Name.Replace('Category:', '')
      pacing       = (Get-MWCategoryMember 'Pacing'                -Type 'subcat').Name.Replace('Category:', '')
      perspectives = (Get-MWCategoryMember 'Perspectives'          -Type 'subcat').Name.Replace('Category:', '')
      controls     = (Get-MWCategoryMember 'Controls'              -Type 'subcat').Name.Replace('Category:', '')
      genres       = (Get-MWCategoryMember 'Genres'                -Type 'subcat').Name.Replace('Category:', '')
      sports       = (Get-MWCategoryMember 'Sports subcategories'  -Type 'subcat').Name.Replace('Category:', '')
      vehicles     = (Get-MWCategoryMember 'Vehicle subcategories' -Type 'subcat').Name.Replace('Category:', '')
     'art styles'  = (Get-MWCategoryMember 'Art styles'            -Type 'subcat').Name.Replace('Category:', '')
      themes       = (Get-MWCategoryMember 'Themes'                -Type 'subcat').Name.Replace('Category:', '')
    }

    $PopularTags = $SteamStorePage.getElementsByClassName('app_tag') | Select-Object -Expand 'innerText'
    $PopularTags = $PopularTags | Where-Object { $_ -ne '+' }

    foreach ($Key in $Taxonomy.Keys)
    {
      $Values = @()

      foreach ($Value in $Taxonomy[$Key])
      {
        $TranslatedValue = $Value

        if ($TranslatedValue -eq 'Multiplayer')
        { $TranslatedValue = 'Multi-player'}
        elseif ($TranslatedValue -eq 'Singleplayer')
        { $TranslatedValue = 'Single-player'}

        if ($SteamData.categories.description -contains $TranslatedValue -or
            $SteamData.genres.description     -contains $TranslatedValue -or
            $PopularTags                      -contains $TranslatedValue)
        { $Values += $Value }
      }

      # If no pacing, assume Real-Time
      if ($Key -eq 'pacing' -and $Values.Count -eq 0)
      { $Values += 'Real-time' }

      # Force Singleplayer to be listed first
      if ($Key -eq 'modes')
      { $Values = $Values | Sort-Object -Descending }
      
      if ($Values)
      { $Template.Wikitext = $Template.Wikitext | SetTemplate "Infobox game/row/taxonomy/$Key" -Value ($Values -join ', ') }
    }


    # In-App Purchases
    $InAppPurchases = ($SteamData.categories.description -contains 'In-App Purchases')

    if ($SteamData.is_free -eq 'true')
    {
      if ($InAppPurchases)
      { $FreeToPlay = $true }
      else
      { $Freeware   = $true }
    }
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
  $Template.Wikitext = $Template.Wikitext | SetTemplate 'Infobox game/row/taxonomy/series' -Value $Series

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
    $Template.Wikitext = $Template.Wikitext | SetParameter 'OSfamily' -Value $Platforms[0]
    #$Template.Wikitext = $Template.Wikitext.Replace('|OSfamily = Windows', "|OSfamily = $($Platforms[0])")
  }

  if ($InAppPurchases)
  { $Template.Wikitext = $Template.Wikitext | SetParameter 'none' -Value '' }

  # Free-to-Play
  if ($FreeToPlay)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext | SetTemplate 'Infobox game/row/taxonomy/monetization' -Value 'Free-to-play'
    $Template.Wikitext = $Template.Wikitext | SetParameter 'license' -Value 'free-to-play'

    # Monetization table
    $Template.Wikitext = $Template.Wikitext | SetParameter 'one-time game purchase' -Value ''
    if ($InAppPurchases)
    { $Template.Wikitext = $Template.Wikitext | SetParameter 'free-to-play' -Value 'Game is free-to-play with in-app purchases.' }
    else
    { $Template.Wikitext = $Template.Wikitext | SetParameter 'free-to-play' -Value 'Game is free-to-play.' }
  }

  # Freeware
  if ($Freeware)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext | SetTemplate 'Infobox game/row/taxonomy/monetization' -Value 'Freeware'
    $Template.Wikitext = $Template.Wikitext | SetParameter 'license' -Value 'freeware'

    # Monetization table
    $Template.Wikitext = $Template.Wikitext | SetParameter 'one-time game purchase' -Value ''
    $Template.Wikitext = $Template.Wikitext | SetParameter 'freeware' -Value 'Game is freeware.'
  }

  # Shareware
  if ($Shareware)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext | SetParameter 'license' -Value 'shareware'
    # Assume a shareware title is a one-time purchase as well
  }

  # No DLCs
  if ($NoDLCs)
  { $Template.Wikitext = $Template.Wikitext.Replace("`n{{DLC|`n<!-- DLC rows goes below: -->`n`n}}`n", '') }

  elseif ($SteamData.dlc)
  {
    # Retrieve info about the DLC using separate requests...

  }

  if ($SteamData)
  {
    $Template.Wikitext = $Template.Wikitext.Replace('==Availability==', @"
'''General information'''
{{mm}} [http://steamcommunity.com/app/$SteamAppId/discussions/ Steam Community Discussions]

==Availability==
"@)

    # Game Data
    if ($SteamData.categories.description -contains 'Steam Cloud')
    { $Template.Wikitext = $Template.Wikitext | SetParameter 'steam cloud' -Value 'true' }
    else
    { $Template.Wikitext = $Template.Wikitext | SetParameter 'steam cloud' -Value 'false' }

    # Video
    if ($SteamData.categories.description -contains 'HDR available')
    { $Template.Wikitext = $Template.Wikitext | SetParameter 'hdr' -Value 'true' }
    if ($SteamData.categories.description -contains 'Color Alternatives')
    { $Template.Wikitext = $Template.Wikitext | SetParameter 'color blind' -Value 'true' }

    # Input
    if ($SteamData.controller_support -eq 'full')
    {
      $Template.Wikitext = $Template.Wikitext | SetParameter 'controller support' -Value 'true'
      $Template.Wikitext = $Template.Wikitext | SetParameter 'full controller' -Value 'true'
    } elseif ($SteamData.categories.description -contains 'Partial Controller Support')
    {
      $Template.Wikitext = $Template.Wikitext | SetParameter 'controller support' -Value 'true'
      $Template.Wikitext = $Template.Wikitext | SetParameter 'full controller' -Value 'false'
    }
    
    # Audio
    if ($SteamData.categories.description -contains 'Custom Volume Controls')
    { $Template.Wikitext = $Template.Wikitext | SetParameter 'separate volume' -Value 'true' }

    $Sound = @()
    if ($SteamData.categories.description -contains 'Stereo Sound')
    { $Sound += 'Stereo' }
    if ($SteamData.categories.description -contains 'Surround Sound')
    { $Sound += '5.1' }

    if ($Sound)
    { $Template.Wikitext = $Template.Wikitext | SetParameter 'surround sound' -Value ($Sound -join ', ') }
  }

  # Game data
  $GameDataConfig = @()
  $GameDataSaves  = @()
  foreach ($Platform in $Platforms)
  {
    $GameDataConfig += "{{Game data/config|$Platform|}}"
    $GameDataSaves  += "{{Game data/saves|$Platform|}}"
  }

  $Template.Wikitext = $Template.Wikitext -replace '(?m)^\{\{Game data/config\s?\|.*\|?\}\}$', ($GameDataConfig -join "`n")
  $Template.Wikitext = $Template.Wikitext -replace '(?m)^\{\{Game data/saves\s?\|.*\|?\}\}$',  ($GameDataSaves  -join "`n")

  # Create page
  if ([string]::IsNullOrWhiteSpace($TargetPage))
  { $TargetPage = $Name }

  if ($WhatIf)
  {
    [Console]::BackgroundColor = 'Black'
    [Console]::ForegroundColor = 'Yellow'
    [Console]::WriteLine('What if: Performing maintenance on target "' + $TargetPage + '".')
    [Console]::ResetColor()
    return $Template.Wikitext
  } else {
    return Set-MWPage -Name $TargetPage -Summary 'Created page' -Major -CreateOnly -Content $Template.Wikitext
  }
}

End { }