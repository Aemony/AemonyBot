[CmdletBinding(DefaultParameterSetName = 'PageID')]
Param (
  <#
    Core parameters
  #>
  [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'PageName', Position=0)]
  [ValidateNotNullOrEmpty()]
  [Alias("Title", "Identity", "PageName")]
  [string]$Name,

  [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, ParameterSetName = 'PageID', Position=0)]
  [Alias("PageID")]
  [int]$ID,

  [switch]$WhatIf
)

Begin {
  # Configuration
  $script:ProgressPreference = 'SilentlyContinue' # Suppress progress bar (speeds up Invoke-WebRequest by a ton)
}

Process {

  # Regex 101: https://regex101.com/
































# --------------------------------------------------------------------------------------- #
#                                                                                         #
#                                     HELPER CMDLETs                                      #
#                                                                                         #
# --------------------------------------------------------------------------------------- #

#region Read-Title
  function Read-Title ($Body)
  {
    $Title = '' # <title[^>]*\>((?:[^<]|\s)*)\<\/title>
    if ($Body -match '<title[^>]*\>([^-|]*\|\s*|[^<]*)\<\/title\>') # 
    {
      $Title = ($Matches[1].Replace("`n", '') -replace ('\s+', ' ')).Trim()

      if ($Title -like "* - Buy and download on GamersGate*")
      {
        $Title = "GamersGate - $Title"
        $Title = $Title.Replace(' - Buy and download on GamersGate', '')
      }
      $Matches = $null
    }
    return $Title
  }
#endregion

#region Set-Substring
  Set-Alias -Name Replace-Substring -Value Set-Substring
  function Set-Substring
  {
  <#
    .SYNOPSIS
      String helper used to replace the nth occurrence of a substring.
    .DESCRIPTION
      Function used to replace the nth occurrence of a substring within a given string,
      using an optional string comparison type.
    .PARAMETER InputObject
      String to act upon.
    .PARAMETER Substring
      Substring to search for.
    .PARAMETER NewSubstring
      The new substring to replace the found substring with.
    .PARAMETER Occurrence
      The nth occurrence to replace. Defaults to first occurrence.
    .PARAMETER Comparison
      The string comparison type to use. Defaults to InvariantCultureIgnoreCase.
    .EXAMPLE
      $ContentBlock | Set-Substring -Substring $Target -NewSubstring $NewSection -Occurrence -1
    .INPUTS
      String to act upon.
    .OUTPUTS
      Returns InputObject with the nth matching substring changed.
  #>
    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)]
      [string]$InputObject,
      
      [Parameter(Mandatory, Position=0)]
      [string]$Substring,

      [Parameter(Mandatory, Position=1)]
      [Alias('Replacement')]
      [AllowEmptyString()]
      [string]$NewSubstring,

      [Parameter()]
        [int]$Occurrence = 0, # Positive: from start; Negative: from back.

      [Parameter()]
      [StringComparison]$Comparison = [StringComparison]::InvariantCultureIgnoreCase
    )
    
    Begin { }

    Process
    {
      $Index   = -1
      $Indexes = @()

      if ($Occurrence -gt 0)
      {
        $Occurrence--
      }

      do
      {
        $Index = $InputObject.IndexOf($Substring, 1 + $Index, $Comparison)
        if ($Index -ne -1)
        {
          $Indexes += $Index
        }
      } while ($Index -ne -1)

      if ($null  -ne   $Indexes[$Occurrence]) {
        $Index       = $Indexes[$Occurrence]
        $InputObject = $InputObject.Remove($Index, $Substring.Length).Insert($Index, $NewSubstring)
      } elseif ($Indexes.Count -gt 0) {
        Write-Verbose "The specified occurrence does not exist."
      } else {
        Write-Verbose "No matching substring was found."
      }

      return $InputObject
    }

    End { }
  }
#endregion

#region Export-Metadata
  function Export-Metadata ($Link, $Title, $Date)
  {
    # Ensure TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $WebPage      =  $null
    $StatusCode   =  $null
    $LinkAnchor   = ($Link -split '#')
    $UpdatedLink  =  $Link
    $UpdatedTitle =  $Title
    $TempTitle    =  $null
    $UpdatedDate  =  $Date

    # Restore the # to the link anchor
    if ($LinkAnchor.Count -gt 1)
    {
      $LinkAnchor = "#" + $LinkAnchor[-1]
    } else {
      $LinkAnchor = '' # Clear the variable so we do not mistakenly duplicate a non-anchored link
    }

    # Try to retrieve the web page

    $ExcludeDomains = @(
      # Cloudflare protection
      'http(?:s)?:\/\/(?:www\.)?wsgf\.org'
      'http(?:s)?:\/\/(?:www\.)?superuser\.com'

      # Other
    )

    $FetchWeb = $true

    if ($ExcludeDomains | Where-Object { $Link -match $_ }) { $FetchWeb = $false }
    
    if ($FetchWeb)
    {
      try {
        Write-Verbose "Retrieving $Link"
        $WebPage    = Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive
        $StatusCode = $WebPage.StatusCode
        $TempTitle  = Read-Title $WebPage.Content
      } catch {
        $StatusCode = $_.Exception.response.StatusCode.value__
        Write-Warning "Failed (HTTP $StatusCode) trying to retrieve $Link"
      }

      # Helper array used to try to detect 404 pages through the website title
      $NotFound = @(
        "page not found",
        "not found",
        "HTTP 404"
      )

      # If the retrieval fails, try to retrieve the latest archive of it
      if ($StatusCode -eq 404 -or ($NotFound | Where-Object { $TempTitle -Like "*$_*" } ))
      {
        # TODO Use the Wayback Machine API: https://archive.org/help/wayback_api.php
        try {
          Write-Verbose "Retrieving https://web.archive.org/web/$Link"
          $WebPage    = Invoke-WebRequest -Uri "https://web.archive.org/web/$Link" -Method GET -UseBasicParsing -DisableKeepAlive
          $StatusCode = $WebPage.StatusCode
          $TempTitle  = Read-Title $WebPage.Content
        } catch {
          $StatusCode = $_.Exception.response.StatusCode.value__
          Write-Warning "Failed (HTTP $StatusCode) trying to retrieve https://web.archive.org/web/$Link"
        }
      }

      if ($WebPage -and $StatusCode -eq 200)
      {
        $UpdatedLink = "$($WebPage.baseResponse.ResponseUri)"

        if (-not ([string]::IsNullOrEmpty($LinkAnchor)) -and $UpdatedLink -NotLike "*#*")
        {
          $UpdatedLink += $LinkAnchor
        }

        if ([string]::IsNullOrEmpty($UpdatedTitle) -and ($NotFound | Where-Object { $TempTitle -NotLike "*$_*" } ))
        {
          $UpdatedTitle = $TempTitle
        }
      }
    }

    if ([string]::IsNullOrEmpty($UpdatedTitle))
    {
      # http://trolloll.ocker.derp.lasd/asokjpogaopkgopsa
      #        trolloll.ocker.derp.lasd
      $ExtratedDomain = $Link -replace '(?:https?:)(?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n]+)(?:\/.*)?', '$1'
      $UpdatedTitle = "$ExtratedDomain - Unknown page title (retrieval failure)"

      # Super User special handling
      if ($Link -match 'http(?:s)?:\/\/(?:www\.)?superuser\.com')
      {
        # https://superuser.com/questions/1334140/how-to-check-if-a-binary-is-16-bit-on-windows
       #$ThreadID     =  $Link -replace '(.*)/(\d+)/(.*)', '$2' # 1334140
        $ThreadTitle  = (($Link -replace '(.*)/(\d+)/(.*)', '$3') -replace '(.*)\?.*', '$1') -replace '-', ' ' # how-to-check-if-a-binary-is-16-bit-on-windows -> how to check if a binary is 16 bit on windows
        $ThreadTitle  = (Get-Culture).TextInfo.ToTitleCase($ThreadTitle) # Title Case
        $UpdatedTitle = "Super User - Thread #${Thread}: $ThreadTitle"
      }
    }

    if ($UpdatedLink -match 'http(?:s):\/\/web\.archive\.org' -and $UpdatedTitle -NotLike "*(archived)*")
    {
      $UpdatedTitle += " (archived)"

      # https://web.archive.org/web/20120108022353/<link>
      $WaybackTimestamp = $null
      $WaybackTimestamp = ($UpdatedLink -split ('/'))[4]
      if ($null -ne $WaybackTimestamp)
      {
        if ($WaybackTranslated = [datetime]::ParseExact($WaybackTimestamp, 'yyyyMMddHHmmss', $null))
        {
          $UpdatedDate = $WaybackTranslated.ToString('yyyy-MM-dd')
        }
      }
    }

    $Object = @{
      Link  = $UpdatedLink
      Title = $UpdatedTitle
      Date  = $UpdatedDate
    }

    return $Object
  }
#endregion

#region Add-MissingTemplateParameters
function Add-MissingTemplateParameters
{
  [CmdletBinding()]
  param(
    [Parameter(Mandatory, ValueFromPipeline)]
    [string]$InputObject,
    
    [Parameter(Mandatory)]
    [string]$DetectionString, # |peripheral devices

    [Parameter(Mandatory)]
    [string]$InsertionText,   # |peripheral devices            = `n|peripheral device types       = `n|peripheral devices notes      = `n

    [Parameter(Mandatory)]
    [string]$PrependBefore    # \|other button prompts\s*=(.+)\n
  )

  Process
  {
    if (-not ($InputObject.Contains($DetectionString)))
    {
      if ($InputObject -match $PrependBefore)
      {
        # $Matches[0] holds the full match
        # $Matches[1] holds the capture group
        $InputObject  = $InputObject.Replace($Matches[0], $InsertionText + $Matches[0])
        $Matches = $null
      }
    }

    $InputObject
  }
}

#endregion






























# --------------------------------------------------------------------------------------- #
#                                                                                         #
#                                     INITIALIZATION                                      #
#                                                                                         #
# --------------------------------------------------------------------------------------- #

#region Initialization...

  $Page   = $null
  if ($PSBoundParameters.ContainsKey('Name'))
  { $Page = $Name }
  else
  { $Page = $ID }


  if ($WhatIf)
  {
    [Console]::BackgroundColor = 'Black'
    [Console]::ForegroundColor = 'Yellow'
    [Console]::WriteLine('What if: Performing maintenance on target "' + $Page + '".')
    [Console]::ResetColor()
  }

  Write-Verbose "Working on $Page..."

  # Used to store details of the edit, if one is performed
  $Output = $null

  $Summary   = 'Maintenance:'
  $Tags       = @()
  $Today     = (Get-Date).ToString("yyyy-MM-dd")
  $ThisMonth = (Get-Date).ToString("MMMM yyyy", [CultureInfo]'en-us') # June 2025

  # Taxonomy Categories
  # $TaxonomyCategories = Get-MWCategoryMember 'Taxonomy' -Namespace Category | foreach { Get-MWCategoryMember $_.Name -Namespace Category }
  # Static for now since it would be too costly to fetch it anew for every new page we process
  $TaxonomyCategories = @(
    "Category:Abstract",
    "Category:Anime",
    "Category:Cartoon",
    "Category:Cel-shaded",
    "Category:Comic book",
    "Category:Digitized",
    "Category:FMV",
    "Category:Live action",
    "Category:Pixel art",
    "Category:Pre-rendered graphics",
    "Category:Realistic",
    "Category:Stylized",
    "Category:Vector art",
    "Category:Video backdrop",
    "Category:Voxel art",
    "Category:Direct control",
    "Category:Gestures",
    "Category:Menu-based",
    "Category:Multiple select",
    "Category:Point and select",
    "Category:Text input",
    "Category:Voice control",
    "Category:4X",
    "Category:Action",
    "Category:Adventure",
    "Category:Arcade",
    "Category:ARPG",
    "Category:Art",
    "Category:Artillery",
    "Category:Battle royale",
    "Category:Board",
    "Category:Brawler",
    "Category:Building",
    "Category:Business",
    "Category:Card/tile",
    "Category:CCG",
    "Category:Chess",
    "Category:Clicker",
    "Category:Dating",
    "Category:Driving",
    "Category:Educational",
    "Category:Endless runner",
    "Category:Exploration",
    "Category:Falling block",
    "Category:Farming",
    "Category:Fighting",
    "Category:FPS",
    "Category:Gambling/casino",
    "Category:Hack and slash",
    "Category:Hidden object",
    "Category:Hunting",
    "Category:Idle",
    "Category:Immersive sim",
    "Category:Interactive book",
    "Category:JRPG",
    "Category:Life sim",
    "Category:Mental training",
    "Category:Metroidvania",
    "Category:Mini-games",
    "Category:MMO",
    "Category:MMORPG",
    "Category:Music/rhythm",
    "Category:Open world",
    "Category:Paddle",
    "Category:Party game",
    "Category:Pinball",
    "Category:Platform",
    "Category:Puzzle",
    "Category:Quick time events",
    "Category:Racing",
    "Category:Rail shooter",
    "Category:Roguelike",
    "Category:Rolling ball",
    "Category:RPG",
    "Category:RTS",
    "Category:Sandbox",
    "Category:Shooter",
    "Category:Simulation",
    "Category:Sports",
    "Category:Stealth",
    "Category:Strategy",
    "Category:Survival",
    "Category:Survival horror",
    "Category:Tactical RPG",
    "Category:Tactical shooter",
    "Category:TBS",
    "Category:Text adventure",
    "Category:Tile matching",
    "Category:Time management",
    "Category:Tower defense",
    "Category:TPS",
    "Category:Tricks",
    "Category:Trivia/quiz",
    "Category:Vehicle combat",
    "Category:Vehicle simulator",
    "Category:Visual novel",
    "Category:Wargame",
    "Category:Word",
    "Category:Boost",
    "Category:Cosmetic",
    "Category:Currency",
    "Category:Finite spend",
    "Category:Free-to-grind",
    "Category:Infinite spend",
    "Category:Loot box",
    "Category:No microtransactions",
    "Category:Player trading",
    "Category:Time-limited",
    "Category:Unlock",
    "Category:Multiplayer",
    "Category:Singleplayer",
    "Category:Ad-supported",
    "Category:Cross-game bonus",
    "Category:DLC",
    "Category:Expansion pack",
    "Category:Free-to-play",
    "Category:Freeware",
    "Category:One-time game purchase",
    "Category:Subscription",
    "Category:Subscription gaming service",
    "Category:Continuous turn-based",
    "Category:Persistent",
    "Category:Real-time",
    "Category:Relaxed",
    "Category:Turn-based",
    "Category:Audio-based",
    "Category:Bird's-eye view",
    "Category:Cinematic camera",
    "Category:First-person",
    "Category:Flip screen",
    "Category:Free-roaming camera",
    "Category:Isometric",
    "Category:Scrolling",
    "Category:Side view",
    "Category:Text-based",
    "Category:Third-person",
    "Category:Top-down view",
    "Category:American football",
    "Category:Australian football",
    "Category:Baseball",
    "Category:Basketball",
    "Category:Bowling",
    "Category:Boxing",
    "Category:Cricket",
    "Category:Darts/target shooting",
    "Category:Dodgeball",
    "Category:Extreme sports",
    "Category:Fictional sport",
    "Category:Fishing",
    "Category:Football (Soccer)",
    "Category:Golf",
    "Category:Handball",
    "Category:Hockey",
    "Category:Horse",
    "Category:Lacrosse",
    "Category:Martial arts",
    "Category:Mixed sports",
    "Category:Paintball",
    "Category:Parachuting",
    "Category:Pool or snooker",
    "Category:Racquetball/squash",
    "Category:Rugby",
    "Category:Sailing/boating",
    "Category:Skateboarding",
    "Category:Skating",
    "Category:Snowboarding or skiing",
    "Category:Surfing",
    "Category:Table tennis",
    "Category:Tennis",
    "Category:Volleyball",
    "Category:Water sports",
    "Category:Wrestling",
    "Category:Adult",
    "Category:Africa",
    "Category:Amusement park",
    "Category:Antarctica",
    "Category:Arctic",
    "Category:Asia",
    "Category:China",
    "Category:Classical",
    "Category:Cold War",
    "Category:Comedy",
    "Category:Contemporary",
    "Category:Cyberpunk",
    "Category:Dark",
    "Category:Detective/mystery",
    "Category:Eastern Europe",
    "Category:Egypt",
    "Category:Europe",
    "Category:Fantasy",
    "Category:Healthcare",
    "Category:Historical",
    "Category:Horror",
    "Category:Industrial Age",
    "Category:Interwar",
    "Category:Japan",
    "Category:LGBTQ",
    "Category:Lovecraftian",
    "Category:Medieval",
    "Category:Middle East",
    "Category:North America",
    "Category:Oceania",
    "Category:Piracy",
    "Category:Post-apocalyptic",
    "Category:Pre-Columbian Americas",
    "Category:Prehistoric",
    "Category:Renaissance",
    "Category:Romance",
    "Category:Sci-fi",
    "Category:South America",
    "Category:Space",
    "Category:Steampunk",
    "Category:Supernatural",
    "Category:Victorian",
    "Category:Western",
    "Category:World War I",
    "Category:World War II",
    "Category:Zombies",
    "Category:Automobile",
    "Category:Bicycle",
    "Category:Bus",
    "Category:Flight",
    "Category:Helicopter",
    "Category:Hovercraft",
    "Category:Industrial",
    "Category:Motorcycle",
    "Category:Naval/watercraft",
    "Category:Off-roading",
    "Category:Robot",
    "Category:Self-propelled artillery",
    "Category:Space flight",
    "Category:Street racing",
    "Category:Tank",
    "Category:Track racing",
    "Category:Train",
    "Category:Transport",
    "Category:Truck"
  )

  # Use Get-MWPage cuz we want RevisionID and Timestamp
  $Page   = $null
  if ($PSBoundParameters.ContainsKey('Name')) {
    $Page = (Get-MWPage -Wikitext -Name $Name)
  } elseif ($PSBoundParameters.ContainsKey('ID')) {
    $Page = (Get-MWPage -Wikitext -ID   $ID)
  }

  # Extract the namespace name from the page name
  $Page | Add-Member -MemberType NoteProperty -Name 'Namespace' -Value (Get-MWNamespace -PageName $Page.Name).Name

  # If page has content AND page is not a Redirect page
  if ($null -ne $Page.Wikitext -and -not ($Page.Wikitext.Contains('#REDIRECT')))
  {
    $OriginalContent     = $null
    $OriginalContent     = $Page.Wikitext

    $IsGame = $false
    if ($Page.Wikitext -like "*{{Infobox game`n*"  -or
        $Page.Wikitext -like "*{{Infobox game|`n*")
    {
      $IsGame = $true
    }

    # For -matches:
    #   $Matches[0] holds the full match
    #   $Matches[1] holds the capture group
    # 
    # Note that PowerShell prefers using references as much as possible,
    #   so $Matches[#].Clone() must be used where needed to do a deep copy.

#endregion
































# --------------------------------------------------------------------------------------- #
#                                                                                         #
#                                       PROCESSING                                        #
#                                                                                         #
# --------------------------------------------------------------------------------------- #

#region Date References
    # Convert |March 25, 2025<ref> to |March 25, 2025|<ref>
    $Before       = $Page.Wikitext
    while ($Page.Wikitext -match '\{\{Infobox game\/row\/date\|(.+?)\|([\w-\s,]*)\<ref')
    {
      $Section = $Matches[0].Clone()
      $OS      = $Matches[1].Trim()
      $OSDate  = $Matches[2].Trim()

      $Replacement  = "{{Infobox game/row/date|$OS|$OSDate|ref=<ref"
      $Page.Wikitext = $Page.Wikitext.Replace($Section, $Replacement)
      $Matches = $null
    }
    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' ~syntax'
    }
#endregion

#region ReferenceFix #0
    # Convert <ref>link title</ref> to {{Refurl}}
    $Before       = $Page.Wikitext
    while ($Page.Wikitext -match "\<ref\>(http(?:s):\/\/[\w\/\^\~\*'@&\+$%#\?=;:._\(\)\-]+?)\s+(.*?)\<\/ref>")
    {
      $Section  = $Matches[0].Clone()
      $Link     = $Matches[1].Trim()
      $Title    = $Matches[2].Trim()

      # Process it all
      $Metadata = Export-Metadata -Link $Link -Title $Title -Date $Today
      $Link     = $Metadata.Link
      $Title    = ConvertTo-MWEscapedString $Metadata.Title
      $LinkDate = $Metadata.Date

      # Swap in the replacements
      $Replacement  = "<ref>{{Refurl|url=$Link|title=$Title|date=$LinkDate}}</ref>"
      $Page.Wikitext = $Page.Wikitext.Replace($Section, $Replacement)
      $Matches = $null
    }
    if ($Before -cne $Page.Wikitext -and $Summary -notlike "*~refurl*")
    {
      $Summary += ' ~refurl'
    }
#endregion

#region ReferenceFix #1
    # Convert <ref>[link title]</ref> to {{Refurl}}
    $Before       = $Page.Wikitext
    while ($Page.Wikitext -match "\<ref\>\[(http(?:s):\/\/[\w\/\^\~\*'@&\+$%#\?=;:._\(\)\-]+?)\s+(.*?)\]\<\/ref>")
    {
      $Section  = $Matches[0].Clone()
      $Link     = $Matches[1].Trim()
      $Title    = $Matches[2].Trim()

      # Process it all
      $Metadata = Export-Metadata -Link $Link -Title $Title -Date $Today
      $Link     = $Metadata.Link
      $Title    = ConvertTo-MWEscapedString $Metadata.Title
      $LinkDate = $Metadata.Date

      # Swap in the replacements
      $Replacement  = "<ref>{{Refurl|url=$Link|title=$Title|date=$LinkDate}}</ref>"
      $Page.Wikitext = $Page.Wikitext.Replace($Section, $Replacement)
      $Matches = $null
    }
    if ($Before -cne $Page.Wikitext -and $Summary -notlike "*~refurl*")
    {
      $Summary += ' ~refurl'
    }
#endregion

#region ReferenceFix #2
    # Convert <ref>[link]</ref> to {{Refurl}}
    $Before       = $Page.Wikitext
    while ($Page.Wikitext -match "\<ref\>\[(http(?:s):\/\/[\w\/\^\~\*'@&\+$%#\?=;:._\(\)\-]+?)\]\<\/ref>")
    {
      $Section  = $Matches[0].Clone()
      $Link     = $Matches[1].Trim()
      $Title    = ''

      # Process it all
      $Metadata = Export-Metadata -Link $Link -Title $Title -Date $Today
      $Link     = $Metadata.Link
      $Title    = ConvertTo-MWEscapedString $Metadata.Title
      $LinkDate = $Metadata.Date

      # Swap in the replacements
      $Replacement  = "<ref>{{Refurl|url=$Link|title=$Title|date=$LinkDate}}</ref>"
      $Page.Wikitext = $Page.Wikitext.Replace($Section, $Replacement)
      $Matches = $null
    }
    if ($Before -cne $Page.Wikitext -and $Summary -notlike "*~refurl*")
    {
      $Summary += ' ~refurl'
    }
#endregion

#region ReferenceFix #3
    # Convert <ref>link</ref> to {{Refurl}}
    $Before       = $Page.Wikitext
    while ($Page.Wikitext -match '\<ref\>(http(?:s):\/\/.+?)\<\/ref>')
    {
      $Section  = $Matches[0].Clone()
      $Link     = $Matches[1].Trim()
      $Title    = ''

      # Process it all
      $Metadata = Export-Metadata -Link $Link -Title $Title -Date $Today
      $Link     = $Metadata.Link
      $Title    = ConvertTo-MWEscapedString $Metadata.Title
      $LinkDate = $Metadata.Date

      # Swap in the replacements
      $Replacement  = "<ref>{{Refurl|url=$Link|title=$Title|date=$LinkDate}}</ref>"
      $Page.Wikitext = $Page.Wikitext.Replace($Section, $Replacement)
      $Matches = $null
    }
    if ($Before -cne $Page.Wikitext -and $Summary -notlike "*~refurl*")
    {
      $Summary += ' ~refurl'
    }
#endregion

#region StrategyWiki
    if ($Page.Wikitext -match '\|strategywiki\s*=(.+)\n')
    {
      # $Matches[0] holds the full match
      # $Matches[1] holds the capture group
      $Before       = $Page.Wikitext
      if ($Link = $Matches[1].Trim())
      {
        $Replacement   = $Matches[0].Replace($Link, $Link.Replace('_', ' '))
        $Page.Wikitext = $Page.Wikitext.Replace($Matches[0], $Replacement)
        if ($Before -cne $Page.Wikitext)
        {
          $Summary += ' ~strategywiki'
        }
      }
      $Matches = $null
    }
#endregion

#region DLC table
    if ($Page.Namespace -eq '')
    {
      # Only do on articles in the main namespace

      # Requires retrieving the date through the Cargo backend (admittedly easier than trying to parse it manually)
      $ReleaseDate = Get-MWCargoQuery -Table Game -Field 'Released' -Where ('_pageID = ' + $Page.ID) -Limit 1

      if (-not ([string]::IsNullOrEmpty($ReleaseDate.Released)))
      {
        # Strip out any empty values (caused by invalid input being used on the article)
        $ReleaseDate = ($ReleaseDate.Released -split ';') | Where-Object { $_ -ne '' } | Sort-Object

        # If we have multiple results, use the first one
        if ($ReleaseDate.Count -gt 1)
        {
          $ReleaseDate = $ReleaseDate[0]
        }

        # Ensure that the data we are working on is actually a string (caused by _all_ input on the article being invalid/garbage)
        if (-not ([string]::IsNullOrEmpty($ReleaseDate)))
        {
          $FirstRelease = [DateTime]::Parse($ReleaseDate)
          $LastYear     = (Get-Date).AddYears(-1)

          if ($FirstRelease -lt $LastYear)
          {
            $Before       = $Page.Wikitext
            $Page.Wikitext = $Page.Wikitext.Replace("`n{{DLC|`n<!-- DLC rows goes below: -->`n`n}}`n", '')
            if ($Before -cne $Page.Wikitext)
            {
              $Summary += ' -DLCs'
            }
          }
        }
      }
    }
#endregion

    # Removal of Key Points tend to generate additional newlines so let us first run it once with announcement,
    #  then after the key points have been handled, we run it again but silently.

#region Newlines
    $Before       = $Page.Wikitext
    # Trim multiple newlines, e.g. \n\n\n -> \n\n
    $Page.Wikitext = $Page.Wikitext -replace "(`r?`n){3,}", "`n`n" # $([Environment]::Newline)$([Environment]::Newline)
    # Trim newlines following a parameter, e.g. |current state[...]\n\n<here there be content>
    $Page.Wikitext = $Page.Wikitext -replace '(\|[\w\s]*[\s]*=[\s]\n)\n([^\||<|}])', '$1$2'

    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' -newlines'
    }
#endregion

#region Key Points
    $Before       = $Page.Wikitext
    # Restrict changes to game pages for now
    # TODO: FIX OTHER PAGE TYPES AS WELL!!!
    if ($IsGame -and $Page.Wikitext -match "\n((?:['=]){1,4}Key points(?:['=]){1,4})\n(.|\s)*?(?=(?:\n'''General information'''|\n==Availability==|\{\{introduction|\|release history|\|current state|\n\n))")
    {
      Write-Verbose 'Page has Key Points...'

      # $Matches[0] holds the full match
      # $Matches[1] holds the header
      # $Matches[2] holds the trailing newlines

      $KeyPointsBlock = $Matches[0].Clone()
      $Header         = $Matches[1].Clone()
      $Matches = $null

      # Strip bullets at the beginning of the lines
      $KeyPointsTrim = ($KeyPointsBlock.Replace($Header, '') -replace '\{\{[im+\-]{2}\}\}[\s]*', "`n").Trim()

      # Let us begin by clearing the data entirely from its current position on the page...
      $Page.Wikitext = $Page.Wikitext.Replace($KeyPointsBlock, "`n`n") # Remove the whole Key Point block with newlines. Any unnecessary newlines will be cleared up further down

      # Does the page have a 'current state' section ?
      if ($Page.Wikitext -match "\|current state(.|\s)*?(?=(\n'''General information'''|\n==Availability==)\n)")
      {
        # $Matches[0] holds the full match

        Write-Verbose 'Page has a current state section; adding key points to the bottom of it...'

        # This caused issues when the page had an empty section as the replacement ended up running on every single newline of the page...
        #   It is why we now use a full-block replacement object to do a more localized replacement before we swap that in.
        $ContentBlock = $Matches[0].Clone()
        $Target       = '}}' # We target the trailing }}
        $NewSection   = "`n$KeyPointsTrim`n}}`n`n"
        # We only replace the last occurance
        $Replacement  = $ContentBlock | Set-Substring -Substring $Target -NewSubstring $NewSection -Occurrence -1

        # Insert the key points at the bottom of the current state section
        $Page.Wikitext = $Page.Wikitext.Replace($ContentBlock, $Replacement)
        $Matches = $null
      }
      
      # No 'current state' detected, so we need to create it!
      else {
        # Does the article have an release history section at least ?
        if ($Page.Wikitext -match "\|release history(.|\s)*?(?=(\n'''General information'''|\n==Availability==)\n)")
        {
          # $Matches[0] holds the full match

          Write-Verbose 'Page has an release history section; adding current state + key points to the bottom of it...'

          $ContentBlock = $Matches[0].Clone()
          $Target       = '}}' # We target the trailing }}
          $NewSection   = "`n|current state = `n$KeyPointsTrim`n}}`n"
          # We only replace the last occurance
          $Replacement  = $ContentBlock | Set-Substring -Substring $Target -NewSubstring $NewSection -Occurrence -1

          # Insert the key points at the bottom of the current state section
          $Page.Wikitext = $Page.Wikitext.Replace($ContentBlock, $Replacement)
          $Matches = $null
        }

        # Does the article have an introduction section at least ?
        elseif ($Page.Wikitext -match "\|introduction(.|\s)*?(?=(\n'''General information'''|\n==Availability==)\n)")
        {
          # $Matches[0] holds the full match

          Write-Verbose 'Page has an introduction section; adding release history + current state + key points to the bottom of it...'

          $ContentBlock = $Matches[0].Clone()
          $Target       = '}}' # We target the trailing }}
          $NewSection   = "`n|release history   = `n`n|current state     = `n$KeyPointsTrim`n}}`n"
          # We only replace the last occurance
          $Replacement  = $ContentBlock | Set-Substring -Substring $Target -NewSubstring $NewSection -Occurrence -1

          # Insert the key points at the bottom of the current state section
          $Page.Wikitext = $Page.Wikitext.Replace($ContentBlock, $Replacement)
          $Matches = $null
        }

        else {
          # Locate the infobox game
          if ($Page.Wikitext -match "{{Infobox game(.|\s)*?(?=(?:\n'''General information'''|\n==Availability==)\n)")
          {
            # $Matches[0] holds the infobox game stuff in its entirety

            Write-Verbose 'Page has an infobox; adding introduction section + key points...'

            $ContentBlock = $Matches[0].Clone()
            $NewSection   = "`n`n{{Introduction`n|introduction      = `n{{Introduction/oneliner}}`n`n|release history   = `n`n|current state     = `n$KeyPointsTrim`n}}`n`n"
            # Append the new section to the bottom of the infobox
            $Replacement  = $ContentBlock + $NewSection

            # Insert the introduction + key points at the bottom of the infobox game section
            $Page.Wikitext = $Page.Wikitext.Replace($ContentBlock, $Replacement)
            $Matches = $null
          }
        }
      }
    }
    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' -keypoints'
      $Tags     += $('pcgw-removed-keypoints')
    }
#endregion

#region Newlines (quiet)
    $Before       = $Page.Wikitext
    # Trim multiple newlines, e.g. \n\n\n -> \n\n
    $Page.Wikitext = $Page.Wikitext -replace "(`r?`n){3,}", "`n`n" # $([Environment]::Newline)$([Environment]::Newline)
    # Trim newlines following a parameter, e.g. |current state[...]\n\n<here there be content>
    $Page.Wikitext = $Page.Wikitext -replace '(\|[\w\s]*[\s]*=[\s]\n)\n([^\||<|}])', '$1$2'
#endregion

#region Reference Spacing
    $Before       = $Page.Wikitext
    $Page.Wikitext = $Page.Wikitext -replace "}}(`r?`n)\{\{References}}", "}}`n`n{{References}}"
    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' +ref_newline'
    }
#endregion

#region Clean comments
    $Before       = $Page.Wikitext
    $Page.Wikitext = $Page.Wikitext.Replace('Comment (optional)', '')
    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' -comment'
    }
#endregion

#region Date citations
    $Before       = $Page.Wikitext
    $Page.Wikitext = $Page.Wikitext.Replace('{{cn}}', "{{cn|date=$ThisMonth}}")
    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' ~cn'
    }
#endregion

#region Middleware
    if ($Page.Wikitext -match '(?s){{Middleware(.*?)\n\n={1,6}')
    {
      $Middleware = $Matches[1]
      $Before     = $Page.Wikitext

      function CleanTemplateParameter
      {
        param (
          [Parameter(Mandatory, Position=0)]
          [string]$Template,
          [Parameter(Mandatory, Position=1)]
          [string]$Parameter,
          [Parameter(Mandatory, ValueFromPipeline)]
          [string]$String
        )

        process
        {
          if ($Template -match "\|$Parameter\s*=(.+)\n")
          {
            $Values = $Matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($Values))
            {
              # Remove any links
              $Replacement = ((((($Values -split ',') -replace '\[\[[^\|]+\|(.*)\]\]', '$1') -replace '\[\[(.*)\]\]', '$1') -replace '\[[^\s]+\s(.*)\]', '$1') -join (','))
              
              # Remove any wildcards
              $Replacement = $Replacement.Replace('*', '')

              if ($Values -ne $Replacement)
              {
                Write-Verbose "Performing change on '$Parameter':`nOrg: $Values`nNew: $Replacement"
                return $String.Replace($Values, $Replacement)
              }
            }
            $Matches = $null
          }
          return $String
        }
      }

      $Page.Wikitext = $Page.Wikitext | CleanTemplateParameter $Middleware -Parameter 'physics'
      $Page.Wikitext = $Page.Wikitext | CleanTemplateParameter $Middleware -Parameter 'audio'
      $Page.Wikitext = $Page.Wikitext | CleanTemplateParameter $Middleware -Parameter 'interface'
      $Page.Wikitext = $Page.Wikitext | CleanTemplateParameter $Middleware -Parameter 'input'
      $Page.Wikitext = $Page.Wikitext | CleanTemplateParameter $Middleware -Parameter 'cutscenes'
      $Page.Wikitext = $Page.Wikitext | CleanTemplateParameter $Middleware -Parameter 'multiplayer'
      $Page.Wikitext = $Page.Wikitext | CleanTemplateParameter $Middleware -Parameter 'anticheat'

      if ($Before -cne $Page.Wikitext)
      {
        $Summary += ' ~middleware'
      }
      $Matches = $null
    }
#endregion

#region Add Missing Template Parameters
    $Before       = $Page.Wikitext

    # Needs to run from bottom -> top, since it prepends _above_ another template as that is the easiest

    # Save game cloud syncing: iCloud
    $AddMissingTemplateParam = @{
      DetectionString = "|icloud notes"
      InsertionText   = "|icloud                    = `n|icloud notes              = `n"
      PrependBefore   = "\|steam cloud\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: directinput prompts
    $AddMissingTemplateParam = @{
      DetectionString = "|directinput prompts"
      InsertionText   = "|directinput prompts           = `n|directinput prompts notes     = `n"
      PrependBefore   = "\|playstation controllers\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: directinput controllers
    $AddMissingTemplateParam = @{
      DetectionString = "|directinput controllers"
      InsertionText   = "|directinput controllers       = `n|directinput controllers notes = `n"
      PrependBefore   = "\|directinput prompts\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: dualsense adaptive trigger support modes
    $AddMissingTemplateParam = @{
      DetectionString = "|dualsense adaptive trigger support modes"
      InsertionText   = "|dualsense adaptive trigger support modes = `n"
      PrependBefore   = "\|dualsense adaptive trigger support notes\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: nintendo controllers
    $AddMissingTemplateParam = @{
      DetectionString = "|nintendo controllers"
      InsertionText   = "|nintendo controllers          = `n|nintendo controller models    = `n|nintendo controllers notes    = `n|nintendo prompts              = `n|nintendo prompts notes        = `n|nintendo button layout        = `n|nintendo button layout notes  = `n|nintendo motion sensors       = `n|nintendo motion sensors modes = `n|nintendo motion sensors notes = `n|nintendo connection modes     = `n|nintendo connection modes notes = `n"
      PrependBefore   = "\|tracked motion controllers\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: peripheral devices
    $AddMissingTemplateParam = @{
      DetectionString = "|peripheral devices"
      InsertionText   = "|peripheral devices            = `n|peripheral device types       = `n|peripheral devices notes      = `n"
      PrependBefore   = "\|other button prompts\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: input prompt override
    $AddMissingTemplateParam = @{
      DetectionString = "|input prompt override"
      InsertionText   = "|input prompt override         = `n|input prompt override notes   = `n"
      PrependBefore   = "\|haptic feedback\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: digital movement supported
    $AddMissingTemplateParam = @{
      DetectionString = "|digital movement supported"
      InsertionText   = "|digital movement supported    = `n|digital movement supported notes = `n"
      PrependBefore   = "\|simultaneous input\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: haptic feedback hd
    $AddMissingTemplateParam = @{
      DetectionString = "|haptic feedback hd"
      InsertionText   = "|haptic feedback hd            = `n|haptic feedback hd notes      = `n|haptic feedback hd controller models = `n"
      PrependBefore   = "\|digital movement supported\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: steam input presets
    $AddMissingTemplateParam = @{
      DetectionString = "|steam input presets"
      InsertionText   = "|steam input presets           = `n|steam input presets notes     = `n"
      PrependBefore   = "\|steam cursor detection\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: steam input motion sensors
    $AddMissingTemplateParam = @{
      DetectionString = "|steam input motion sensors"
      InsertionText   = "|steam input motion sensors    = `n|steam input motion sensors modes = `n|steam input motion sensors notes = `n"
      PrependBefore   = "\|steam input presets\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: steam deck prompts
    $AddMissingTemplateParam = @{
      DetectionString = "|steam deck prompts"
      InsertionText   = "|steam deck prompts            = `n|steam deck prompts notes      = `n"
      PrependBefore   = "\|steam controller prompts\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    # Input: steam input prompts
    $AddMissingTemplateParam = @{
      DetectionString = "|steam input prompts"
      InsertionText   = "|steam input prompts           = `n|steam input prompts icons     = `n|steam input prompts styles    = `n|steam input prompts notes     = `n"
      PrependBefore   = "\|steam deck prompts\s*=(.+)\n"
    }
    $Page.Wikitext = $Page.Wikitext | Add-MissingTemplateParameters @AddMissingTemplateParam

    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' +parameters'
    }
#endregion

#region Misc
    $Before       = $Page.Wikitext

    # Change headers to lowercase
    $Page.Wikitext = $Page.Wikitext.Replace('==Issues Fixed==', '==Issues fixed==')

    # Change PAGENAME calls to the actual page name (or display title, if different)
    $ProcessedName = $Page.DisplayTitle
    if ($Page.Namespace -and $Page.DisplayTitle -ceq $Page.Name)
    { $ProcessedName = $ProcessedName.Replace(($Page.Namespace + ':'), '') }
    $Page.Wikitext = $Page.Wikitext.Replace('{{PAGENAME}}', $ProcessedName)

    # | {{DRM|SafeDisc}} disc check. | -> | {{DRM|SafeDisc}} |
    $Page.Wikitext = $Page.Wikitext -replace '\|\s*(\{\{DRM\|(?:\w)+(?:\|[^\}]*)?\}\}) disc check\.?\s*\|', '| $1 |'

    # Image caption punctuation (remove trailing periods)
    $Page.Wikitext = $Page.Wikitext -replace '\{\{Image\|\s*([\w\.]+)\s*\|\s*([^\}]*)\.\s*\}\}', '{{Image|$1|$2}}'

    # Space between bullets and words
    $Page.Wikitext = $Page.Wikitext -replace '\{\{(\+{2}|\-{2}|i{2}|m{2})\}\}(\w)', '{{$1}} $2'

    # Correct category links, but only on pages without links to lists or controller pages...
    if (-not ($Page.Wikitext.Contains('[[:Category:Lists of ')) -and -not ($Page.Wikitext.Contains('[[:Category:Controllers')))
    {
      $Page.Wikitext = $Page.Wikitext -replace "\[\[:Category:([0-9a-zA-Z\-\/ \(\)_']+)\|[0-9a-zA-Z\-\/ \(\)_']+\]\]", '{{Glossary:$1}}'
    }

    # Add references section on pages that lacks it...
    if (-not ($Page.Wikitext.Contains('{{References}}')) -and
             ($Page.Wikitext.Contains('<ref>')))
    {
        # But only if it doesn't expect any transclusion!
        if (-not ($Page.Wikitext.Contains('<noinclude>'))   -and
            -not ($Page.Wikitext.Contains('<includeonly>')) -and
            -not ($Page.Wikitext.Contains('<onlyinclude>')))
        {
          $Page.Wikitext = $Page.Wikitext + "`n`n{{References}}"
        }
    }

    # Convert some external links into external ones:
    $Page.Wikitext = $Page.Wikitext.Replace('[https://reshade.me ReShade]', '[[ReShade]]')
    $Page.Wikitext = $Page.Wikitext.Replace('[https://reshade.me/ ReShade]', '[[ReShade]]')
    $Page.Wikitext = $Page.Wikitext.Replace('[https://github.com/doitsujin/dxvk DXVK]', '[[DXVK]]')
    $Page.Wikitext = $Page.Wikitext.Replace('[https://github.com/doitsujin/dxvk/ DXVK]', '[[DXVK]]')
    $Page.Wikitext = $Page.Wikitext.Replace('[https://www.special-k.info Special K]', '[[Special K]]')
    $Page.Wikitext = $Page.Wikitext.Replace('[https://www.special-k.info/ Special K]', '[[Special K]]')

    # Move commas before any reference that may exist
    # Regex is not well suited for this -- need a regular solution using forward find for "</ref>," and then a reverse find for "<ref>"
    #$Page.Wikitext = $Page.Wikitext -replace '(<ref[^>]*\>.*?)(?=(?:<\/ref>,))', ',$1</ref>'

    # DRM-free
    # Disabled for now because it affects URLs as well................ >_<
    #$Page.Wikitext = $Page.Wikitext.Replace('drm-free', 'DRM-free')

    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' ~misc'
    }
#endregion

#region Add missing template parameters
    $Before       = $Page.Wikitext

    if ($Page.Wikitext -notmatch '\|framegen\s*=')
    {
      $Page.Wikitext = $Page.Wikitext -replace '(\|vsync\s*=)', "|framegen                   = unknown`n|framegen tech              = `n|framegen notes             = `n`$1"
    }

    if ($Before -cne $Page.Wikitext)
    {
      $Summary += ' +template usage'
    }
#endregion

































# --------------------------------------------------------------------------------------- #
#                                                                                         #
#                                      FINALIZATION                                       #
#                                                                                         #
# --------------------------------------------------------------------------------------- #
#region Applying...
    # If a change has been made, apply it
    if ($OriginalContent -cne $Page.Wikitext)
    {
      Write-Verbose $Summary
      if ($WhatIf)
      { $Output = $Page }
      else
      { $Output = $Page | Set-MWPage -Bot -NoCreate -Minor -Summary $Summary -Tags $Tags -BaseRevisionID $Page.RevisionID -StartTimestamp $Page.ServerTimestamp }
    } else {
      Write-Verbose 'No change was made to target.'
    }
#endregion

  }

  return $Output
}

End { }