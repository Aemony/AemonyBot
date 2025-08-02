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
  [string[]]$Modes,

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
  [int]$SteamAppId = 0,

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
  [string]$TargetPage,
  [switch]$WhatIf
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
  # Core object
  $Game = @{
    Name         = ''
    Developers   = @()
    Publishers   = @()
    ReleaseDates = @{
      Linux        = ''
      macOS        = ''
      Windows      = ''
    }
    Platforms    = @()
    Reception    = @{
      MetaCritic   = @{
        Rating       = ''
        URL          = ''
      }
      OpenCritic   = @{
        Rating       = ''
        URL          = ''
      }
      IGDB         = @{
        Rating       = ''
        URL          = ''
      }
    }
    Taxonomy     = @{
      modes        = @()
      pacing       = @()
      perspectives = @()
      controls     = @()
      genres       = @()
      sports       = @()
      vehicles     = @()
     'art styles'  = @()
      themes       = @()
    }
    Website      = ''
    Steam        = @{
      IDs          = @()
     'steam cloud' = 'unknown'
    }
    DLCs         = @()
    Video        = @{
     'hdr'                      = 'unknown'
     'color blind'              = 'unknown'
    }
    Input        = @{
      'controller support'      = 'unknown'
      'full controller support' = 'unknown'
    }
    Audio        = @{
      'separate volume'         = 'unknown'
      'surround sound'          = 'unknown'
    }
  }

  # Supported taxonomy tags
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

  function SetTemplateParameter
  {
    param (
      [Parameter(Mandatory, Position=0)]
      [string]$Template,

      [Parameter(Mandatory, Position=1)]
      [string]$Parameter,

      [Parameter(Mandatory, Position=2)]
      [AllowEmptyString()]
      [string]$Value,

      [Parameter(Mandatory, ValueFromPipeline)]
      [string]$String
    )
    
    process
    {
      if ($String -match "(?s){{$Template(.*?)\n\n={1,6}")
      {
        $TemplateBody = $Matches[1].Trim()
        $TemplateRepl = $TemplateBody | SetParameter $Parameter -Value $Value
        #Write-Verbose "Performing change on '$Template':`nOrg: $TemplateBody`nNew: $TemplateRepl"
        return $String.Replace($TemplateBody, $TemplateRepl)
      }

      return $String
    }
  }

  function GetSteamData($AppId)
  {
    # Extract information from Steam
    Write-Verbose "Steam App ID: $AppId"

    $Details      = $null
    $PageComObject = $null # ComObject: HTMLFile

    $UAGoogleBot  = 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'
    $Session      = New-Object Microsoft.PowerShell.Commands.WebRequestSession # [Microsoft.PowerShell.Commands.WebRequestSession]::new()
    $CookiesAge   = [System.Net.Cookie]::new('birthtime', '0')
    $CookiesAdult = [System.Net.Cookie]::new('mature_content', '1')
    $Session.Cookies.Add('https://store.steampowered.com/', $CookiesAge)
    $Session.Cookies.Add('https://store.steampowered.com/', $CookiesAdult)

    # Steam Store API
    $Link = "https://store.steampowered.com/api/appdetails/?appids=$AppId&l=english"
    try {
      Write-Verbose "Retrieving $Link"
      $WebPage    = Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive -UserAgent $UAGoogleBot -WebSession $Session
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

    if ($Json.$AppId.success -ne 'true')
    {
      Write-Warning 'Failed to parse Json from Steam!'
      return
    }

    $Details = $Json.$AppId.data
    $Type = $Details.type

    if ($Type -ne 'game')
    {
      Write-Warning "$AppId is not a game!"
      return
    }

    # Steam Store
    $Link = "https://store.steampowered.com/app/$AppId/&l=english"
    try {
      Write-Verbose "Retrieving $Link"
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

    $PageComObject = New-Object -ComObject "HTMLFile"
    [string]$PageContent = $WebPage.Content
    $PageComObject.Write([ref]$PageContent)



    # Parsing the data

    $Game.Name = $Details.name

    if ($Details.categories | Where-Object { $_.description -eq 'Multi-player' })
    { $Game.Taxonomy.modes += 'Multiplayer' }
    
    if ($Details.categories | Where-Object { $_.description -eq 'Single-player' })
    { $Game.Taxonomy.modes += 'Singleplayer' }

    $Game.Developers = $Details.developers
    $Pubs = $Details.publishers | Where-Object { $Developers -notcontains $_ }
    if ($null -ne $Pubs)
    { $Game.Publishers = $Pubs }

    $ReleaseDate   = 'TBA'

    # EA
    if ($Details.genres.description -contains 'Early Access')
    { $ReleaseDate = 'EA' }

    # TBA
    elseif ($Details.release_date.date -ne '' -and
            $Details.release_date.date -ne 'Coming soon')
    { $ReleaseDate = $Details.release_date.date }

    if ($Details.platforms.mac -eq 'true')
    { $Game.ReleaseDates.macOS   = $ReleaseDate }

    if ($Details.platforms.linux -eq 'true')
    { $Game.ReleaseDates.Linux   = $ReleaseDate }

    if ($Details.platforms.windows -eq 'true')
    { $Game.ReleaseDates.Windows = $ReleaseDate }

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

    $PopularTags = $PageComObject.getElementsByClassName('app_tag') | Select-Object -Expand 'innerText'
    $PopularTags = $PopularTags | Where-Object { $_ -ne '+' }

    foreach ($Key in $Taxoonmy.Keys)
    {
      $Values = @()

      foreach ($Value in $Taxonomy[$Key])
      {
        $TranslatedValue = $Value

        if ($TranslatedValue -eq 'Multiplayer')
        { $TranslatedValue = 'Multi-player'}
        elseif ($TranslatedValue -eq 'Singleplayer')
        { $TranslatedValue = 'Single-player'}

        if ($Details.categories.description -contains $TranslatedValue -or
            $Details.genres.description     -contains $TranslatedValue -or
            $PopularTags                      -contains $TranslatedValue)
        { $Values += $Value }
      }
      
      if ($Values)
      { $Game.Taxonomy.$Key = $Values }
    }

    # IGDB      : https://store.steampowered.com/app/1814770/Tall_Poppy/
    # OpenCritic: https://store.steampowered.com/app/1561340/Berserk_Boy/
    # MetaCritic: https://store.steampowered.com/app/1561340/Berserk_Boy/
    if ($Reviews = $PageComObject.getElementsByName('game_area_reviews') | Select-Object -Expand 'innerHtml')
    {
      $Part1  = RegexEscape('<a href="https://steamcommunity.com/linkfilter/?u=')
      $Part2 = RegexEscape('" rel=" noopener" target=_blank>')
      $Part3  = RegexEscape('</a>')
      ($Reviews -split "<br>") | Where-Object { $_ -match "^(\d+)\s.\s$Part1(.*)$Part2([\w\s]+)$Part3$" } | ForEach-Object {
        if ($Matches[3] -eq 'MetaCritic')
        {
          $Game.Reception.MetaCritic.Rating = $Matches[1]
          $Game.Reception.MetaCritic.Url    = [System.Uri]::UnescapeDataString($Matches[2]).Replace('https://www.metacritic.com/game/', '') -replace '([\w|\d|\-]+).*', '$1'
        }

        if ($Matches[3] -eq 'OpenCritic')
        {
          $Game.Reception.OpenCritic.Rating = $Matches[1]
          $Game.Reception.OpenCritic.Url    = [System.Uri]::UnescapeDataString($Matches[2]).Replace('https://opencritic.com/game/', '') -replace '(\d+\/[\w|\d|\-]+).*', '$1'
        }

        if ($Matches[3] -eq 'IGDB')
        {
          $Game.Reception.IGDB.Rating       = $Matches[1]
          $Game.Reception.IGDB.Url          = [System.Uri]::UnescapeDataString($Matches[2]).Replace('https://www.igdb.com/games/', '') -replace '([\w|\d|\-]+).*', '$1'
        }
      }
    }

    # In-App Purchases
    $InAppPurchases = ($Details.categories.description -contains 'In-App Purchases')

    if ($Details.is_free -eq 'true')
    {
      if ($InAppPurchases)
      { $FreeToPlay = $true }
      else
      { $Freeware   = $true }
    }

    # Game Data
    if ($Details.categories.description -contains 'Steam Cloud')
    { $Game.Steam.'steam cloud'        = 'true' }

    # Video
    if ($Details.categories.description -contains 'HDR available')
    { $Game.Video.'hdr'                = 'true' }
    if ($Details.categories.description -contains 'Color Alternatives')
    { $Game.Video.'color blind'        = 'true' }

    # Input
    if ($Details.controller_support -eq 'full')
    {
      $Game.Input.'controller support' = 'true'
      $Game.Input.'full controller'    = 'true'
    } elseif ($Details.categories.description -contains 'Partial Controller Support')
    {
      $Game.Input.'controller support' = 'true'
      $Game.Input.'full controller'    = 'false'
    }
    
    # Audio
    if ($Details.categories.description -contains 'Custom Volume Controls')
    { $Game.Audio.'separate volume'    = 'true' }

    $Sound = @()
    if ($Details.categories.description -contains 'Stereo Sound')
    { $Sound += 'Stereo' }
    if ($Details.categories.description -contains 'Surround Sound')
    { $Sound += '5.1' }

    if ($Sound)
    { $Game.Audio.'surround sound'     = ($Sound -join ', ') }

    return @{
      Details = $Details
      Store   = @{
        Page      = $PageContent
        ComObject = $PageComObject
      }
    }
  }


  <#
  
    START PROCESSING
  
  #>

  # Initialization: Steam Game

  if (-not [string]::IsNullOrWhiteSpace($SteamUrl))
  { $SteamAppId = ($SteamUrl -replace '(?m)^([^\d]+\/app\/)(\d+)(\/?.*)', '$2') }

  $Steam = @{
    Details = ''
    Store   = @{
      Page      = ''
      ComObject = $null # COM Object: HtmlFile
    }
  }

  if ($SteamAppId -ne 0)
  { $Steam = GetSteamData ($SteamAppId) }

  # Initialization: Generic Game

  if ($Name)
  { $Game.Name = $Name }

  if ($Modes)
  { $Game.Taxonomy.modes = $Modes }

  if ($Developers)
  { $Game.Developers = $Developers }

  if ($Publishers)
  { $Game.Publishers = $Publishers }

  if ($ReleaseDateWindows)
  { $Game.ReleaseDates.Windows = $ReleaseDateWindows }

  if ($ReleaseDateMacOS)
  { $Game.ReleaseDates.macOS = $ReleaseDateMacOS }

  if ($ReleaseDateLinux)
  { $Game.ReleaseDates.Linux = $ReleaseDateLinux }



  # Validation

  if ([string]::IsNullOrWhiteSpace($Game.Name))
  {
    Write-Warning 'A game name needs to be specified to continue!'
    return
  }

  if ([string]::IsNullOrWhiteSpace($Game.Developers))
  {
    Write-Warning 'A developer needs to be specified to continue!'
    return
  }

  if (-not [string]::IsNullOrWhiteSpace($Game.ReleaseDates.macOS))
  { $Game.Platforms += 'macOS' }

  if (-not [string]::IsNullOrWhiteSpace($Game.ReleaseDates.Linux))
  { $Game.Platforms += 'Linux' }

  if (-not [string]::IsNullOrWhiteSpace($Game.ReleaseDates.Windows))
  { $Game.Platforms += 'Windows' }

  if ($Game.Platforms -notcontains 'Windows')
  { $NoWindows = $true }

  if ($NoWindows)
  {
    if ([string]::IsNullOrWhiteSpace($Game.ReleaseDates.Linux) -and
        [string]::IsNullOrWhiteSpace($Game.ReleaseDates.macOS))
    {
      Write-Warning 'A Linux or macOS release date needs to be specified when using -NoWindows!'
      return
    }

    if (-not [string]::IsNullOrWhiteSpace($Game.ReleaseDates.Windows))
    {
      Write-Warning '-NoWindows and -ReleaseDateWindows cannot be used at the same time!'
      return
    }
  }



  # Processing
  $Game.Name         = $Game.Name.Replace('™', '')
  $Game.Name         = $Game.Name.Replace('®', '')
  $Game.Name         = $Game.Name.Replace('©', '')
  $Game.Name         = $Game.Name.Replace(': ', ' - ')
  $Game.Name         = $Game.Name.Replace(':', '')

  $Game.Developers   = $Game.Developers -replace '(?:,?\s|,)(?:Inc|Ltd|GmbH|S\.?A|LLC|V\.?O\.?F|AB)\.?$', ''
  if ($Game.Publishers)
  { $Game.Publishers = $Game.Publishers -replace '(?:,?\s|,)(?:Inc|Ltd|GmbH|S\.?A|LLC|V\.?O\.?F|AB)\.?$', '' }


  # Game Name

  $TemplateName      = if ($Modes) { $Modes[0] } else { 'Unknown' }
  $Template          = Get-MWPage -PageName $Templates[$TemplateName] -Wikitext
  $Template.Wikitext = $Template.Wikitext.Replace('GAME TITLE', $Game.Name)

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

  foreach ($Key in $Game.Taxonomy.Keys)
  {
    $Values = $Game.Taxonomy[$Key]

    # If no pacing, assume Real-Time
    if ($Key -eq 'pacing' -and $Values.Count -eq 0)
    { $Values += 'Real-time' }

    # Force Singleplayer to be listed first
    if ($Key -eq 'modes')
    { $Values = $Values | Sort-Object -Descending }
    
    if ($Values)
    { $Template.Wikitext = $Template.Wikitext | SetTemplate "Infobox game/row/taxonomy/$Key" -Value ($Values -join ', ') }
  }

  # Series
  $Series = ''
  $Template.Wikitext = $Template.Wikitext | SetTemplate 'Infobox game/row/taxonomy/series' -Value $Series

  # Developer
  $Template.Wikitext = $Template.Wikitext.Replace('DEVELOPER', $Game.Developers[0])

  if ($Game.Developers.Count -gt 1)
  {
    foreach ($Developer in $Game.Developers)
    { $Template.Wikitext = $Template.Wikitext.Replace('|publishers   = ', "{{Infobox game/row/developer|$Developer}}`n|publishers   = ") }
  }

  # Publisher
  if (-not [string]::IsNullOrWhiteSpace($Game.Publishers))
  {
    $Template.Wikitext = $Template.Wikitext.Replace('PUBLISHER', $Game.Publishers[0])

    if ($Game.Publishers.Count -gt 1)
    {
      foreach ($Publisher in $Game.Publishers)
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

  

  if ($Game.ExternalData.SteamIDs)
  {
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'steam appid' -Value "$($Game.ExternalData.SteamIDs[0])"

    if ($Game.ExternalData.SteamIDs.Count -gt 1)
    { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'steam appid side' -Value "$(($Game.ExternalData.SteamIDs[1..($Game.ExternalData.SteamIDs.Count)]) -join ', ')" }
  }
  

  # Website
  if ($Game.Website)
  { $Template.Wikitext = $Template.Wikitext.Replace('|official site= ', ('|official site= ' + $Game.Website)) }

  # No Retail
  if ($NoRetail)
  { $Template.Wikitext = $Template.Wikitext.Replace("{{Availability/row| retail | | unknown |  |  | Windows }}`n", '') }
  # Steam
  elseif ($Game.ExternalData.SteamIDs)
  { $Template.Wikitext = $Template.Wikitext.Replace('{{Availability/row| retail | | unknown |  |  | Windows }}', "{{Availability/row| steam | $($Game.ExternalData.SteamIDs[0]) | steam |  |  | $($Game.Platforms -join ', ') }}") }
  # Retail
  else
  { $Template.Wikitext = $Template.Wikitext.Replace('{{Availability/row| retail | | unknown |  |  | Windows }}', "{{Availability/row| retail | | unknown |  |  | $($Game.Platforms -join ', ') }}") }

  # No Windows
  if ($NoWindows)
  {
    # Remove release date
    $Template.Wikitext = $Template.Wikitext.Replace("{{Infobox game/row/date|Windows|TBA}}`n", '')

    # Remove game data
    $Template.Wikitext = $Template.Wikitext.Replace("{{Game data/config|Windows|}}`n", '')
    $Template.Wikitext = $Template.Wikitext.Replace("{{Game data/saves|Windows|}}`n", '')

    # Replace Windows in the system requirements with the first supported OS
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'System requirements' -Parameter 'OSfamily' -Value $Game.Platforms[0]
  }

  if ($InAppPurchases)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Microtransactions' -Parameter 'none' -Value '' }

  # Free-to-Play
  if ($FreeToPlay)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext | SetTemplate 'Infobox game/row/taxonomy/monetization' -Value 'Free-to-play'
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'license' -Value 'free-to-play'

    # Monetization table
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Monetization' -Parameter 'one-time game purchase' -Value ''
    if ($InAppPurchases)
    { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Monetization' -Parameter 'free-to-play' -Value 'Game is free-to-play with in-app purchases.' }
    else
    { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Monetization' -Parameter 'free-to-play' -Value 'Game is free-to-play.' }
  }

  # Freeware
  if ($Freeware)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext | SetTemplate 'Infobox game/row/taxonomy/monetization' -Value 'Freeware'
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'license' -Value 'freeware'

    # Monetization table
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Monetization' -Parameter 'one-time game purchase' -Value ''
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Monetization' -Parameter 'freeware' -Value 'Game is freeware.'
  }

  # Shareware
  if ($Shareware)
  {
    # Infobox game
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'license' -Value 'shareware'
    # Assume a shareware title is a one-time purchase as well
  }

  # No DLCs
  if ($NoDLCs)
  { $Template.Wikitext = $Template.Wikitext.Replace("`n{{DLC|`n<!-- DLC rows goes below: -->`n`n}}`n", '') }

  
  # Steam Community
  if ($Game.Steam.IDs[0])
  {
    $Template.Wikitext = $Template.Wikitext.Replace('==Availability==', @"
'''General information'''
{{mm}} [http://steamcommunity.com/app/$($Game.Steam.IDs[0])/discussions/ Steam Community Discussions]

==Availability==
"@)
  }

  # Game data
  $GameDataConfig = @()
  $GameDataSaves  = @()
  foreach ($Platform in $Game.Platforms)
  {
    $GameDataConfig += "{{Game data/config|$Platform|}}"
    $GameDataSaves  += "{{Game data/saves|$Platform|}}"
  }

  $Template.Wikitext = $Template.Wikitext -replace '(?m)^\{\{Game data/config\s?\|.*\|?\}\}$', ($GameDataConfig -join "`n")
  $Template.Wikitext = $Template.Wikitext -replace '(?m)^\{\{Game data/saves\s?\|.*\|?\}\}$',  ($GameDataSaves  -join "`n")



  foreach ($Key in $Game.Video.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Video' -Parameter $Key -Value $Game.Video[$Key] }

  foreach ($Key in $Game.Input.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Input' -Parameter $Key -Value $Game.Input[$Key] }

  foreach ($Key in $Game.Audio.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Audio' -Parameter $Key -Value $Game.Audio[$Key] }







  # Create page
  if ([string]::IsNullOrWhiteSpace($TargetPage))
  { $TargetPage = $Game.Name }

  if ($WhatIf)
  {
    [Console]::BackgroundColor = 'Black'
    [Console]::ForegroundColor = 'Yellow'
    [Console]::WriteLine('What if: Performing maintenance on target "' + $TargetPage + '".')
    [Console]::ResetColor()
    return @{
      Wikitext = $Template.Wikitext
      Game     = $Game
      Steam    = $Steam
    }
  } else {
    return Set-MWPage -Name $TargetPage -Summary 'Created page' -Major -CreateOnly -Content $Template.Wikitext
  }
}

End { }