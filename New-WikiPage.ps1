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
    Steam ID / URL
  #>
  [Parameter(Mandatory, ParameterSetName = 'Steam')]
  [string]$Steam,

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

  <#
    Authentication
  #>
  [switch]$Online,  # Create the page as well (requires bot password + authentication)
  [switch]$Offline, # Runs locally (default)

  <#
    Debug
  #>
  [string]$TargetPage,
  [switch]$WhatIf
)

Begin {
  # Configuration
  $ProgressPreference = 'SilentlyContinue' # Suppress progress bar (speeds up Invoke-WebRequest by a ton)
  $ScriptLoadedModule = $false
  $ScriptConnectedAPI = $false

  # Script variable to indicate the location of the saved config file
  $ConfigFileName     = $env:LOCALAPPDATA + '\PowerShell\New-WikiPage\config.json'
  $SteamDataPath      = $env:LOCALAPPDATA + '\PowerShell\New-WikiPage\steam_apps.json'

  $Templates = @{
    Singleplayer = 'PCGamingWiki:Sample article/Game (singleplayer)'
    Multiplayer  = 'PCGamingWiki:Sample article/Game (multiplayer)'
    Unknown      = 'PCGamingWiki:Sample article/Game (unknown)'
  }

  if ($null -eq (Get-Module -Name 'MediaWiki'))
  {
    if ($null -ne (Get-Module -ListAvailable -Name 'MediaWiki'))
    {
      Import-Module -Name   'MediaWiki' 6> $null # 6> $null redirects Write-Host to $null
      $ScriptLoadedModule = $true
    }
    
    elseif (Test-Path -Path '.\MediaWiki')
    {
      Import-Module -Name '.\MediaWiki' 6> $null # 6> $null redirects Write-Host to $null
      $ScriptLoadedModule = $true
    }
  }

  if ($null -ne (Get-Module -Name 'MediaWiki'))
  {
    $ApiProperties = @{
      ApiEndpoint = 'https://www.pcgamingwiki.com/w/api.php'
      Guest      = (-not $Online)
      Silent     = $true
    }

    Connect-MWSession @ApiProperties
    $ScriptConnectedAPI = $true
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
}

Process
{
  if ($null -eq (Get-Module -Name 'MediaWiki'))
  {
    Write-Warning 'MediaWiki module has not been loaded!'
    return $null
  }

  if ($null -eq (Get-MWSession))
  {
    Write-Warning 'Not connected to the PCGW API endpoint!'
    return $null
  }

  $CoverPath           = '.\cover.jpg'

  # Sanitize input
  $KnownDRMs = [PSCustomObject]@{
    'Denuvo'                  = 'Denuvo Anti-Tamper'
    'Denuvo Anti-Tamper'      = 'Denuvo Anti-Tamper'
  }

  # Sanitize input
  $KnownACs = [PSCustomObject]@{
    'EAC'                      = 'Easy Anti-Cheat'
    'Easy AntiCheat'           = 'Easy Anti-Cheat'
    'Easy Anti-Cheat'          = 'Easy Anti-Cheat'
    'BattlEye'                 = 'BattlEye'
    'NGS'                      = 'Nexon Game Security'
    'Nexon Game Security'      = 'Nexon Game Security'
    'NGS(Nexon Game Security)' = 'Nexon Game Security'
    'ACE'                      = 'Anti-Cheat Expert'
    'AntiCheat Expert'         = 'Anti-Cheat Expert'
    'Anti-Cheat Expert'        = 'Anti-Cheat Expert'
    'Anti-Cheat Expert (ACE)'  = 'Anti-Cheat Expert'
    'EA AC'                    = 'EA Javelin'
    'EA AntiCheat'             = 'EA Javelin'
    'EA Anti-Cheat'            = 'EA Javelin'
    'EA Javelin Anticheat'     = 'EA Javelin'
    'EA Javelin Anti-Cheat'    = 'EA Javelin'
    'GameGuard'                = 'nProtect GameGuard'
    'nProtect GameGuard'       = 'nProtect GameGuard'
    'Denuvo'                   = 'Denuvo Anti-Cheat'
    'Denuvo Anti-Cheat'        = 'Denuvo Anti-Cheat'
  }

  # Ignore some DLCs based on their name
  $IgnoredDLCs = @(
    'art book'
    'artbook'
    'donation'
    'sponsor'
    'wallpaper'
  )

  # Core object
  $Game = [PSCustomObject]@{
    Name         = ''
    Series       = @()
    Developers   = @()
    Publishers   = @()
    Platforms    = @() # Also holds release dates
    Reception    = @{
      Metacritic   = [PSCustomObject]@{
        Rating       = ''
        Link         = ''
      }
      OpenCritic   = [PSCustomObject]@{
        Rating       = ''
        Link         = ''
      }
      IGDB         = [PSCustomObject]@{
        Rating       = ''
        Link         = ''
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
      DRMs         = @()
    }
    Introduction = @{
      'introduction'            = '' # Allow it to auto-generate by default.
      'release history'         = ''
      'current state'           = ''
    }
    DLCs         = @()
    Cloud        = @{}
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
    Localizations= @()
    Network      = @{}
    VR           = @{}
    API          = @{
      'direct3d versions'       = @()
     #'vulkan versions'         = ''
    }
    Middleware   = @{
      'physics'                 = @()
      'multiplayer'             = @()
      'anticheat'               = @()
    }
    SystemRequirements = @(
      <#
        @{
          Platform            = 'Windows/Linux/macOS'
          Minimum/Recommended = @{
            TGT   = '' # Target/Settings (TGT)

            # Processor (CPU)
            CPU   = ''
            CPU2  = ''

            RAM   = '' # System memory (RAM)
            HD    = '' # Storage drive (HDD/SSD)

            # Video card (GPU)
            GPU   = ''
            GPU2  = ''
            GPU3  = ''
            VRAM  = ''
            OGL   = '' # OpenGL
            DX    = '' # DirectX
            SM    = '' # Shader Model

            Audio = '' # Sound (audio device)
            Cont  = '' # Controller
            Other = '' # Other
          }
        }
      #>
    )
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

  # Fallback function
  function GetSteamDataFallback
  {
    param (
      $UserAgent,
      $WebSession
    )
    # If the file exists, do not update it
    if (Test-Path $SteamDataPath)
    { return }

    $UAGoogleBot = $UserAgent
    $Session     = $WebSession

    # Fallback
    $SteamUserID = '76561198017975643' # This is just some public user with a ton of games, lol
    # Taken from https://steamladder.com/ladder/games/
    $SteamWebApiKey = $null

    if ((Test-Path $ConfigFileName) -eq $true)
    {
      Try
      {
        # Try to load the config file.
        $SteamWebApiKey = Get-Content $ConfigFileName -ErrorAction Stop | ConvertTo-SecureString -ErrorAction Stop

        # Try to convert the hashed password. This will only work on the same machine that the config file was created on.
      } Catch [System.Management.Automation.ItemNotFoundException], [System.ArgumentException] {
        # Handle corrupt config file
        Write-Warning "The stored configuration could not be found or was corrupt.`n"
        $SteamWebApiKey = $null
      } Catch [System.Security.Cryptography.CryptographicException] {
        # Handle broken key
        Write-Warning "The key in the stored configuration could not be read."
        $SteamWebApiKey = $null
      } Catch {
        # Unknown exception
        Write-Warning "Unknown error occurred when trying to read stored configuration."
        $SteamWebApiKey = $null
      }
    }

    if ($null -eq $SteamWebApiKey)
    {
      Write-Host 'Trying to retrieve game data using the Steam Web API...'
      Write-Host
      Write-Host 'A personal Steam Web API key can be created at https://steamcommunity.com/dev/apikey'
      Write-Host
      Write-Host '!!! DO NOT SHARE YOUR WEB API KEY WITH ANYONE !!!'
      Write-Host

      [SecureString]$SecurePassword = Read-Host 'Steam Web API Key' -AsSecureString
      $SteamWebApiKey = if ($SecurePassword.Length -eq 0) { $null } else { $SecurePassword | ConvertFrom-SecureString }

      if ($SteamWebApiKey)
      {
        Write-Host

        $Yes = @('y', 'yes')
        $No  = @('n', 'no')
        do { $Reply = Read-Host 'Do you want the script to remember the key for future uses? (y/n)' }
        while ($Yes -notcontains $Reply -and $No -notcontains $Reply)

        if ($Yes -contains $Reply)
        {
          # Create the file first using New-Item with -Force parameter so missing directories are also created.
          New-Item -Path $ConfigFileName -ItemType "file" -Force | Out-Null

          # Output the config to the recently created file
          $SteamWebApiKey | Out-File $ConfigFileName
        }
      }
    }

    if ($SteamWebApiKey)
    {
      $BSTR        = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SteamWebApiKey)
      $PlainApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

      $SteamWebApiLink = "https://api.steampowered.com/IPlayerService/GetOwnedGames/v1/?key=$PlainApiKey&steamid=$SteamUserID&skip_unvetted_apps=false&include_played_free_games=true&include_free_sub=true&include_appinfo=true&include_extended_appinfo=true&language=english"

      try {
        #Write-Verbose "Retrieving $SteamWebApiLink"
        Write-Verbose "https://api.steampowered.com/IPlayerService/GetOwnedGames/v1/..."
        Invoke-WebRequest -Uri $SteamWebApiLink -Method GET -UseBasicParsing -DisableKeepAlive -UserAgent $UAGoogleBot -WebSession $Session -OutFile $SteamDataPath
      } catch {
        $StatusCode = $_.Exception.response.StatusCode.value__
      }

      $PlainApiKey    = $null
      $SteamWebApiKey = $null
      $SecurePassword = $null

      if ($BSTR)
      { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR) }
    }
  }

  function GetSteamData($AppId)
  {
    # Extract information from Steam
    Write-Verbose "Steam App ID: $AppId"

    $Details         = $null
    $PageComObject   = $null # ComObject: HTMLFile
    $AssumedCoverUrl = "library_600x900_2x.jpg"

    $UAGoogleBot     = 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'
    $Session         = New-Object Microsoft.PowerShell.Commands.WebRequestSession # [Microsoft.PowerShell.Commands.WebRequestSession]::new()
    $CookiesAge      = [System.Net.Cookie]::new('birthtime', '0')
    $CookiesAdult    = [System.Net.Cookie]::new('mature_content', '1')
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
      Write-Warning 'Failed to parse game json from Steam!'
      GetSteamDataFallback -UserAgent $UAGoogleBot -WebSession $Session
      #return
    }

    $Details = $Json.$AppId.data

    if ($Details.type -and $Details.type -ne 'game')
    {
      Write-Warning "$AppId is not a game!"
      #return
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
      GetSteamDataFallback -UserAgent $UAGoogleBot -WebSession $Session
      #return
    }

    $PageComObject = New-Object -ComObject "HTMLFile"
    [string]$PageContent = $WebPage.Content
    $PageComObject.Write([ref]$PageContent)

    # If we have stored a local copy of Steam's massive app info dump, use it
    $SADBItem = $null

    if (Test-Path $SteamDataPath)
    {
      $SADB     = Get-Content $SteamDataPath -Raw | ConvertFrom-Json
      if ($SADB)
      { $SADBItem = $SADB.response.games | Where-Object { $_.appid -eq $AppId } }
    }

    if ($SADBItem)
    { $AssumedCoverUrl = $SADBItem.capsule_filename.Replace('library_600x900.jpg', 'library_600x900_2x.jpg') }

    $SteamCoverLink  = "https://steamcdn-a.akamaihd.net/steam/apps/$AppId/$AssumedCoverUrl"
    $SteamCoverLink2 = "https://shared.fastly.steamstatic.com/store_item_assets/steam/apps/$AppId/$AssumedCoverUrl"

    # Steam cover (600x900_2x.jpg)
    # Safe to use: "All capsule images (store and library) must have PG-13 appropriate artwork."
    $Link = $SteamCoverLink
    try {
      Write-Verbose "Retrieving $Link"
      Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive -UserAgent $UAGoogleBot -WebSession $Session -OutFile $CoverPath
    } catch {
      $Link = $SteamCoverLink2
      try {
        Write-Verbose "Retrieving $Link"
        Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive -UserAgent $UAGoogleBot -WebSession $Session -OutFile $CoverPath
        $StatusCode = 200
      } catch {
        $StatusCode = $_.Exception.response.StatusCode.value__
      }
    }
    
    # This can happen, but usually only for really new games
    if ($StatusCode -ne 200 -or -not (Test-Path $CoverPath))
    { Write-Warning 'Failed to retrieve game cover from Steam!' }

    # Parsing the data
    $Game.Steam.IDs = @($AppId)

    if ($Details)
    {
      $Game.Name        = $Details.name
      $Game.Developers += $Details.developers.Trim()
      $Game.Publishers += $Details.publishers.Trim() | Where-Object { $Game.Developers -notcontains $_ }
    }
    
    # Fallback
    elseif ($SADBItem)
    {
      $Game.Name        = $SADBItem.Name
      $Game.Developers += 'Unknown'
    }

    if ($null -ne $Details.drm_notice)
    {
      foreach ($DRM in @( $Details.drm_notice.Trim() ))
      {
        if ($KnownDRMs.Keys -contains $DRM)
        { $Game.Steam.DRMs             += $KnownDRMs[$DRM] }
        elseif ($KnownACs.Keys -contains $DRM)
        { $Game.Middleware.'anticheat' += $KnownACs[$DRM] }
        else
        { $Game.Steam.DRMs             += $DRM }
      }
    }

    if ($null -ne $Details.ext_user_account_notice)
    {
      $Game.Steam.DRMs += 'Account'
    }

    $ReleaseDate   = 'TBA'

    # EA
    if ($Details.genres.description -contains 'Early Access')
    { $ReleaseDate = 'EA' }

    # TBA
    elseif ($Details.release_date.date -ne '' -and
            $Details.release_date.date -ne 'Coming soon')
    { $ReleaseDate = $Details.release_date.date }

    if ($Details.platforms.windows -eq 'true')
    {
      $Game.Platforms += [PSCustomObject]@{
        Name        = 'Windows'
        ReleaseDate = $ReleaseDate
      }
    }

    if ($Details.platforms.mac -eq 'true')
    {
      $Game.Platforms += [PSCustomObject]@{
        Name        = 'OS X' # macOS
        ReleaseDate = $ReleaseDate
      }
    }

    if ($Details.platforms.linux -eq 'true')
    {
      $Game.Platforms += [PSCustomObject]@{
        Name        = 'Linux'
        ReleaseDate = $ReleaseDate
      }
    }

    # Fallback
    if ($Game.Platforms.Count -eq 0 -and $SADBItem)
    {
      $Game.Platforms += [PSCustomObject]@{
        Name        = 'Windows'
        ReleaseDate = 'Unknown'
      }
    }

    # Taxonomy
    if ($PopularTags = $PageComObject.getElementsByClassName('app_tag') | Select-Object -Expand 'innerText')
    {
      $PopularTags = $PopularTags | Where-Object { $_ -ne '+' }

      foreach ($Key in $Taxonomy.Keys)
      {
        foreach ($Value in $Taxonomy[$Key])
        {
          $TranslatedValue = $Value

          if ($TranslatedValue -eq 'Multiplayer')
          { $TranslatedValue = 'Multi-player'}
          elseif ($TranslatedValue -eq 'Singleplayer')
          { $TranslatedValue = 'Single-player'}

          if ($Details.categories.description -contains $TranslatedValue -or
              $Details.genres.description     -contains $TranslatedValue -or
              $PopularTags                    -contains $TranslatedValue)
          { $Game.Taxonomy.$Key += $Value }
        }
      }
    }

    # IGDB      : https://store.steampowered.com/app/1814770/Tall_Poppy/
    # OpenCritic: https://store.steampowered.com/app/1561340/Berserk_Boy/
    # MetaCritic: https://store.steampowered.com/app/1561340/Berserk_Boy/

    # AppDetails
    # metacritic	{ score: 78, url: "https://www.metacritic.com/game/pc/lego-star-wars-the-skywalker-saga?ftag=MCD-06-10aaa1f" }
    if ($Details.metacritic)
    {
      $Game.Reception.Metacritic.Rating = $Details.metacritic.score
      $Game.Reception.Metacritic.Link   = [System.Uri]::UnescapeDataString($Details.metacritic.url).Replace('https://www.metacritic.com/game/', '') -replace '(?:pc/)?([\w|\d|\-]+).*', '$1'
    }

    # Store Page
    if ($Reviews = $PageComObject.getElementsByName('game_area_reviews') | Select-Object -Expand 'innerHtml')
    {
      $Part1  = RegexEscape('<a href="https://steamcommunity.com/linkfilter/?u=')
      $Part2  = RegexEscape('" rel=" noopener" target=_blank>')
      $Part3  = RegexEscape('</a>')
      ($Reviews -split "<br>") | Where-Object { $_ -match "^(\d+)\s.\s$Part1(.*)$Part2([\w\s]+)$Part3$" } | ForEach-Object {
        if ($Matches[3] -eq 'MetaCritic' -and [string]::IsNullOrEmpty($Game.Reception.Metacritic.Rating))
        {
          $Game.Reception.Metacritic.Rating = $Matches[1]
          $Game.Reception.Metacritic.Link   = [System.Uri]::UnescapeDataString($Matches[2]).Replace('https://www.metacritic.com/game/', '') -replace '(?:pc/)?([\w|\d|\-]+).*', '$1'
        }

        if ($Matches[3] -eq 'OpenCritic')
        {
          $Game.Reception.OpenCritic.Rating = $Matches[1]
          $Game.Reception.OpenCritic.Link   = [System.Uri]::UnescapeDataString($Matches[2]).Replace('https://opencritic.com/game/', '') -replace '(\d+\/[\w|\d|\-]+).*', '$1'
        }

        #<# Many IGDB review listings are seemingly based on user reviews
        if ($Matches[3] -eq 'IGDB')
        {
          $Game.Reception.IGDB.Rating       = $Matches[1]
          $Game.Reception.IGDB.Link         = [System.Uri]::UnescapeDataString($Matches[2]).Replace('https://www.igdb.com/games/', '') -replace '([\w|\d|\-]+).*', '$1'
        }
        #>
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

    # Series (franchise)
    $LabelFranchise = 'Franchise: ' # Update if Steam Store layout ever changes
    $FalseFranchise = @(
      'WB Games'
    )
    foreach ($DevRow in $PageComObject.getElementsByClassName('dev_row'))
    {
      if ($DevRow.innerText -like "$LabelFranchise*")
      { $Game.Series += ($DevRow.innerText.Replace($LabelFranchise, '') -split ', ').Trim() | Where-Object { $FalseFranchise -notcontains $_ } }
    }

    # Early Access release dates
    $LabelEarlyAccess = 'Early Access Release Date: '
    if ($Block = $PageComObject.getElementsByName('genresAndManufacturer').item(0))
    {
      foreach ($Row in ($Block.innerText -split "`n"))
      {
        if ($Row -like "$LabelEarlyAccess*")
        {
          $EAReleaseDate = $Row.Replace($LabelEarlyAccess, '').Trim()
          
          try {
            $DateTime = [datetime]::Parse($EAReleaseDate)
            $Game.Introduction.'release history' += "On $($DateTime.ToString('MMMM d, yyyy', [CultureInfo]("en-US"))) the game was released to Early Access on Steam."
          } catch {
            
          }
        }
      }
    }

    # Game Data
    if ($Details.categories.description -contains 'Steam Cloud')
    { $Game.Cloud.'steam cloud'        = 'true' }
    else
    { $Game.Cloud.'steam cloud'        = 'unknown' }

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

    if ($Details.categories.description -contains 'Tracked Controller Support')
    { $Game.Input.'tracked motion controllers' = 'true' }
    
    # Audio
    if ($Details.categories.description -contains 'Custom Volume Controls')
    { $Game.Audio.'separate volume'    = 'true' }
    
    if ($Details.categories.description -contains 'Captions available')
    { $Game.Audio.'closed captions'    = 'true' }

    $Sound = @()
    if ($Details.categories.description -contains 'Stereo Sound')
    { $Sound += 'Stereo' }
    if ($Details.categories.description -contains 'Surround Sound')
    { $Sound += '5.1' }

    if ($Sound)
    { $Game.Audio.'surround sound'     = ($Sound -join ', ') }

    # Localization
    if ($L10nRows = $PageComObject.getElementsByClassName('game_language_options').item(0))
    {
      if ($FirstChild = $L10nRows.children(0))
      {
        foreach ($L10nRow in $FirstChild.children)
        {
          # Skip first row
          if ($L10nRow.innerText -like "*Full Audio*")
          { continue }

          $Game.Localizations += [PSCustomObject]@{
            Language  =            $L10nRow.children(0).innerText.Trim()
            Interface = ($null -ne $L10nRow.children(1).innerText)
            Audio     = ($null -ne $L10nRow.children(2).innerText)
            Subtitles = ($null -ne $L10nRow.children(3).innerText)
          }
        }
      }
    }

    if ($Game.Localizations | Where-Object Subtitles -eq $true)
    { $Game.Audio.'subtitles'          = 'true' }

    # VR
    if ($Details.categories.description -contains 'VR Supported')
    { $Game.VR.'vr only'               = 'false' }

    if ($Details.categories.description -contains 'VR Only')
    { $Game.VR.'vr only'               = 'true' }

    # DLCs
    foreach ($DlcId in $Details.dlc)
    {
      Start-Sleep 1

      $Link = "https://store.steampowered.com/api/appdetails/?appids=$DlcId&l=english"
      try {
        Write-Verbose "Retrieving $Link"
        $WebPage    = Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive -UserAgent $UAGoogleBot -WebSession $Session
        $StatusCode = $WebPage.StatusCode
      } catch {
        $StatusCode = $_.Exception.response.StatusCode.value__
      }
      
      if ($StatusCode -ne 200)
      {
        Write-Warning 'Failed to retrieve DLC details from Steam!'
        return
      }

      $Json = ConvertFrom-Json $WebPage.Content

      if ($Json.$DlcId.success -ne 'true')
      {
        Write-Warning 'Failed to parse DLC json from Steam!'
        return
      }

      $DlcDetails = $Json.$DlcId.data

      $Game.DLCs += [PSCustomObject]@{
        Type    =     $DlcDetails.type
        Name    =     $DlcDetails.name
        Free    =     $DlcDetails.is_free
        Notes   = if ($DlcDetails.is_free) { 'Free' } else { '' }
        Ignored = $false
      }
    }
    
    # User + Kernel Anti-Cheat
    if ($ACs = $PageComObject.getElementsByClassName('anticheat_name'))
    {
      foreach ($AC in $ACs | Select-Object -Expand 'innerHtml')
      {
        # Removes <span class="anticheat_uninstalls"> - Requires manual removal after game uninstall</span>
        $Uninstall = $PageComObject.getElementsByClassName('anticheat_uninstalls') | Select-Object -Expand 'outerHtml'
        if ($Uninstall)
        { $AC = $AC.Replace($Uninstall, '') }

        if ($KnownACs.Keys -contains $AC)
        { $Game.Middleware.'anticheat' += $KnownACs[$AC] }
        elseif ($KnownDRMs.Keys -contains $AC)
        { $Game.Steam.DRMs             += $KnownDRMs[$AC] }
        else
        { $Game.Middleware.'anticheat' += $AC }
      }
    }

    # System Requirements
    foreach ($Div in $PageComObject.getElementsByClassName('game_area_sys_req'))
    {
      foreach ($Li in $Div.getElementsByTagName('li'))
      {
        $Li.innerText
      }

      "-------"
    }

    # Return the resulting object
    return @{
      Details  = $Details
      Store    = @{
        Page      = $PageContent
        ComObject = $PageComObject
      }
      Cover    = $SteamCoverLink
      Fallback = $SADBItem
    }
  }


  <#
  
    START PROCESSING
  
  #>

  # Initialization: Steam Game
  $Steam = ($Steam -replace '(?m)^([^\d]+\/app\/)(\d+)(\/?.*)', '$2')

  $SteamApp = @{
    Details = ''
    Store   = @{
      Page      = ''
      ComObject = $null # COM Object: HtmlFile
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($Steam))
  { $SteamApp = GetSteamData ($Steam) }

  # Initialization: Generic Game

  if ($Name)
  { $Game.Name = $Name }

  if ($Modes)
  { $Game.Taxonomy.modes = $Modes }
  else # Steam games
  { $Modes = $Game.Taxonomy.modes }

  if ($Developers)
  { $Game.Developers += $Developers }

  if ($Publishers)
  { $Game.Publishers += $Publishers }

  if ($ReleaseDateWindows)
  {
    $Game.Platforms += [PSCustomObject]@{
      Name        = 'Windows'
      ReleaseDate = $ReleaseDateWindows
    }
  }

  if ($ReleaseDateMacOS)
  {
    $Game.Platforms += [PSCustomObject]@{
      Name        = 'OS X' # macOS
      ReleaseDate = $ReleaseDateMacOS
    }
  }

  if ($ReleaseDateLinux)
  {
    $Game.Platforms += [PSCustomObject]@{
      Name        = 'Linux'
      ReleaseDate = $ReleaseDateLinux
    }
  }


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

  if ($Game.Platforms.Name -notcontains 'Windows' -and
      $Game.Platforms.Name -notcontains 'OS X'   -and
      $Game.Platforms.Name -notcontains 'Linux')
  {
    Write-Warning 'A Windows, macOS, or Linux release date needs to be specified!'
    return
  }


  # Processing

  # Trim game name
  $Game.Name         = $Game.Name.Replace([string][char]0x2122, '') # ™
  $Game.Name         = $Game.Name.Replace([string][char]0x00AE, '') # ®
  $Game.Name         = $Game.Name.Replace([string][char]0x00A9, '') # ©
  $Game.Name         = $Game.Name.Replace(': ', ' - ')
  $Game.Name         = $Game.Name.Replace(':', '')
  $Game.Name         = $Game.Name.Trim(' - ')
  $Game.Name         = $Game.Name.Trim()

  foreach ($DlcObject in $Game.DLCs)
  {
    # Trim DLC names
    $DlcObject.Name = $DlcObject.Name.Replace($Game.Name, '')
    $DlcObject.Name = $DlcObject.Name.Replace([string][char]0x2122, '') # ™
    $DlcObject.Name = $DlcObject.Name.Replace([string][char]0x00AE, '') # ®
    $DlcObject.Name = $DlcObject.Name.Replace([string][char]0x00A9, '') # ©
    $DlcObject.Name = $DlcObject.Name.Replace(': ', ' - ')
    $DlcObject.Name = $DlcObject.Name.Replace(':', ' ')
    $DlcObject.Name = $DlcObject.Name.Replace('  ', ' ')
    $DlcObject.Name = $DlcObject.Name.Replace($Game.Name, '')
    $DlcObject.Name = $DlcObject.Name.Trim(' - ')
    $DlcObject.Name = $DlcObject.Name.Trim()

    # Mark some DLCs as ignored
    if ($null -ne ($IgnoredDLCs | Where-Object { $DlcObject.Name -like "*$_*" }))
    { $DlcObject.Ignored = $true }

    if ($DlcObject.Type -eq 'music')
    { $DlcObject.Ignored = $true }
  }

  $Game.Developers   = $Game.Developers -replace '(?:,?\s|,)(?:Inc|Ltd|GmbH|S\.?A|LLC|V\.?O\.?F|AB)\.?$', ''
  if ($Game.Publishers)
  { $Game.Publishers = $Game.Publishers -replace '(?:,?\s|,)(?:Inc|Ltd|GmbH|S\.?A|LLC|V\.?O\.?F|AB)\.?$', '' }


  # Game Name

  $TemplateName      = if ($Modes) { $Modes[0] } else { 'Unknown' }
  $Template          = Get-MWPage -PageName $Templates[$TemplateName] -Wikitext
  $Template.Wikitext = $Template.Wikitext.Replace('GAME TITLE', $Game.Name)

  if ($null -ne $Game.VR.'vr only')
  {
    $VRSection = (Get-MWPage 'PCGamingWiki:Sample article/Game (unknown)').Sections | Where-Object Line -eq 'VR support' | Get-MWSection -Wikitext | Select-Object -Expand Wikitext
    $Template.Wikitext = $Template.Wikitext.Replace('==Other information==', "$VRSection`n`n==Other information==")
  }

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
  $Template.Wikitext = $Template.Wikitext | SetTemplate 'Infobox game/row/taxonomy/series' -Value ($Game.Series -join ', ')

  # Developer
  $Template.Wikitext = $Template.Wikitext.Replace("{{Infobox game/row/developer|DEVELOPER}}`n", '')
  foreach ($Developer in $Game.Developers)
  { $Template.Wikitext = $Template.Wikitext.Replace('|publishers   = ', "{{Infobox game/row/developer|$Developer}}`n|publishers   = ") }

  # Publisher
  $Template.Wikitext = $Template.Wikitext.Replace("{{Infobox game/row/publisher|PUBLISHER}}`n", '')
  foreach ($Publisher in $Game.Publishers)
  { $Template.Wikitext = $Template.Wikitext.Replace('|engines      = ', "{{Infobox game/row/publisher|$Publisher}}`n|engines      = ") }

  # Remove Windows release date
  $Template.Wikitext = $Template.Wikitext.Replace("{{Infobox game/row/date|Windows|TBA}}`n", '')

  # Release Dates
  foreach ($Platform in $Game.Platforms)
  {
    try {
      $DateTime = [datetime]::Parse($Platform.ReleaseDate)
      $Template.Wikitext = $Template.Wikitext.Replace('|reception    = ', "{{Infobox game/row/date|$($Platform.Name)|$($DateTime.ToString('MMMM d, yyyy', [CultureInfo]("en-US")))}}`n|reception    = ")
    } catch {
      $Template.Wikitext = $Template.Wikitext.Replace('|reception    = ', "{{Infobox game/row/date|$($Platform.Name)|$($Platform.ReleaseDate)}}`n|reception    = ")
    }
  }

  # Reception
  foreach ($Key in $Game.Reception.Keys)
  {
    if (-not [string]::IsNullOrEmpty($Game.Reception.$Key.Rating))
    {
      $Link              = $Game.Reception.$Key.Link
      $Rating            = $Game.Reception.$Key.Rating
      $Pattern           = "$Key|link|rating"
      $Replacement       = "$Key|$Link|$Rating"
      $Template.Wikitext = $Template.Wikitext.Replace($Pattern, $Replacement)
    }
  }

  # Steam IDs
  if ($Game.Steam.IDs)
  {
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'steam appid' -Value "$($Game.Steam.IDs[0])"

    if ($Game.Steam.IDs)
    { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'steam appid side' -Value "$(($Game.Steam.IDs[1..($Game.Steam.IDs.Count)]) -join ', ')" }
  }

  # Website
  if ($Game.Website)
  { $Template.Wikitext = $Template.Wikitext.Replace('|official site= ', ('|official site= ' + $Game.Website)) }

  # Move 'Game (unknown)' template in line with the other two
  $Template.Wikitext = $Template.Wikitext.Replace('{{Availability/row| store  | id | drm | notes  | keys | Windows }}', '{{Availability/row| retail | | unknown |  |  | Windows }}')

  # No Retail
  if ($NoRetail)
  { $Template.Wikitext = $Template.Wikitext.Replace("{{Availability/row| retail | | unknown |  |  | Windows }}`n", '') }
  # Steam
  elseif ($Game.Steam.IDs)
  {
    $Notes = @()

    foreach ($DRM in $Game.Steam.DRMs)
    { $Notes += "{{DRM|$DRM}}" }
    $Notes = $Notes -join ', '

    $Template.Wikitext = $Template.Wikitext.Replace('{{Availability/row| retail | | unknown |  |  | Windows }}', "{{Availability/row| steam | $($Game.Steam.IDs[0]) | steam | $Notes | | $($Game.Platforms.Name -join ', ') }}")
  }
  # Retail
  else
  { $Template.Wikitext = $Template.Wikitext.Replace('{{Availability/row| retail | | unknown |  |  | Windows }}', "{{Availability/row| retail | | unknown |  |  | $($Game.Platforms.Name -join ', ') }}") }

  # No Windows
  if ($Game.Platforms.Name -notcontains 'Windows')
  {
    # Remove game data
    $Template.Wikitext = $Template.Wikitext.Replace("{{Game data/config|Windows|}}`n", '')
    $Template.Wikitext = $Template.Wikitext.Replace("{{Game data/saves|Windows|}}`n", '')

    # Replace Windows in the system requirements with the first supported OS
    $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'System requirements' -Parameter 'OSfamily' -Value $Game.Platforms[0].Name
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

  # Paid DLCs
  if ($Game.DLCs.Free -contains $false)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Monetization' -Parameter 'dlc' -Value 'Game offers paid DLCs.' }
  
  # Introduction
  foreach ($Key in $Game.Introduction.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Introduction' -Parameter $Key -Value $Game.Introduction[$Key] }

  # Steam Community
  if ($Game.Steam.IDs)
  {
    $Template.Wikitext = $Template.Wikitext.Replace('==Availability==', @"
'''General information'''
{{mm}} [http://steamcommunity.com/app/$($Game.Steam.IDs[0])/discussions/ Steam Community Discussions]

==Availability==
"@)
  }

  # DLCs
  $DlcEntry = @"
{{{{DLC/row| {0} | {1} | {2} }}}}`n
"@
  $DlcRows = ''

  foreach ($Dlc in ($Game.DLCs | Where-Object { $_.Ignored -eq $false }))
  { $DlcRows += $DlcEntry -f $Dlc.Name, $Dlc.Notes, ($Game.Platforms.Name -join ', ') }

  if ($DlcRows.Length -gt 0)
  { $DlcRows = "{{DLC|`n$DlcRows}}" }

  $Template.Wikitext = $Template.Wikitext -replace '\{\{DLC\|[\s\-\<\>\!\w\:]*?\}\}', $DlcRows

  # Game data
  $GameDataConfig = @()
  $GameDataSaves  = @()
  foreach ($Platform in $Game.Platforms.Name)
  {
    $GameDataConfig += "{{Game data/config|$Platform|}}"
    $GameDataSaves  += "{{Game data/saves|$Platform|}}"
  }

  $Template.Wikitext = $Template.Wikitext -replace '(?m)^\{\{Game data/config\s?\|.*\|?\}\}$', ($GameDataConfig -join "`n")
  $Template.Wikitext = $Template.Wikitext -replace '(?m)^\{\{Game data/saves\s?\|.*\|?\}\}$',  ($GameDataSaves  -join "`n")

  # Cloud
  foreach ($Key in $Game.Cloud.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Save game cloud syncing' -Parameter $Key -Value $Game.Cloud[$Key] }

  # Video
  foreach ($Key in $Game.Video.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Video'      -Parameter $Key -Value $Game.Video[$Key] }

  # Input
  foreach ($Key in $Game.Input.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Input'      -Parameter $Key -Value $Game.Input[$Key] }

  # Audio
  foreach ($Key in $Game.Audio.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Audio'      -Parameter $Key -Value $Game.Audio[$Key] }

  # Localization
  $L10nEntry = @"
{{{{L10n/switch
 |language  = {0}
 |interface = {1}
 |audio     = {2}
 |subtitles = {3}
 |notes     =
 |fan       =
 |ref       =
}}}}
"@
  $L10nRows = ''

  foreach ($L10n in $Game.Localizations)
  { $L10nRows += $L10nEntry -f $L10n.Language, $L10n.Interface, $L10n.Audio, $L10n.Subtitles }

  $Template.Wikitext = $Template.Wikitext -replace '\{\{L10n\/switch[\s\|\w=]*?\}\}', $L10nRows

  # VR
  foreach ($Key in $Game.VR.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'VR support' -Parameter $Key -Value $Game.VR[$Key] }

  # API
  foreach ($Key in $Game.API.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'API'        -Parameter $Key -Value ($Game.API[$Key] -join ', ') }

  # Middleware
  foreach ($Key in $Game.Middleware.Keys)
  { $Template.Wikitext = $Template.Wikitext | SetTemplateParameter 'Middleware' -Parameter $Key -Value ($Game.Middleware[$Key] -join ', ') }







  # Create page
  if ([string]::IsNullOrWhiteSpace($TargetPage))
  {
    $TargetPage        = $Game.Name
    $Template.Wikitext = $Template.Wikitext.Replace("$($Game.Name) cover.jpg", "$TargetPage cover.jpg")
  }

  $Result = [ordered]@{
    Game     = $Game
  }

  if ($Steam)
  { $Result.Steam = $SteamApp }

  if (Test-Path $CoverPath)
  {
    $CoverPath = Get-Item $CoverPath
    $Result.Cover = $CoverPath.FullName
  }

  $Result.Wikitext = $Template.Wikitext

  # -Online
  if ($Online)
  {
    # Upload cover first (needed for initial page rendering)
    if ($null -ne $Result.Cover)
    {
      $Result.File = Import-MWFile -Name "$TargetPage cover.jpg" -File $CoverPath.FullName
      Remove-Item -Path $CoverPath.FullName -Force
    }

    $Result.Page = Set-MWPage -Name  $TargetPage -Major -CreateOnly -Content $Template.Wikitext

    return $Result
  }

  # -WhatIf
  elseif ($WhatIf)
  {
    [Console]::BackgroundColor = 'Black'
    [Console]::ForegroundColor = 'Yellow'
    [Console]::WriteLine('What if: Performing maintenance on target "' + $TargetPage + '".')
    [Console]::ResetColor()

    return $Result
  }
  
  # -Offline (default)
  else {
    return $Result
  }
}

End {
  if ($ScriptConnectedAPI)
  { Disconnect-MWSession }

  if ($ScriptLoadedModule)
  { Remove-Module -Name 'MediaWiki' }
}