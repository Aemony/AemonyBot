[CmdletBinding(DefaultParameterSetName = 'Generic')]
Param (
  <#
    Generic
  #>
  [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName = 'Generic')]
  [ValidateNotNullOrEmpty()]
  [string]$Name,

  <#
    Switches
  #>
  [switch]$Force,   # Ignore warnings (i.e. overwrite existing or previously removed files)

  <#
    Debug
  #>
  [switch]$WhatIf
)

Begin {
  # Configuration
  $ProgressPreference = 'SilentlyContinue' # Suppress progress bar (speeds up Invoke-WebRequest by a ton)
}

Process
{
  if ($null -eq (Get-Module -Name 'IGDB'))
  {
    Write-Warning 'IGDB module has not been loaded!'
    return $null
  }

  if ($null -eq (Get-Module -Name 'MediaWiki'))
  {
    Write-Warning 'MediaWiki module has not been loaded!'
    return $null
  }

  if ($null -eq (Get-IGDBSession))
  {
    Write-Warning 'Not connected to the IGDB API!'
    return $null
  }

  if ($null -eq (Get-MWSession))
  {
    Write-Warning 'Not connected to the PCGW API!'
    return $null
  }

  function RegexEscape($UnescapedString)
  { return [regex]::Escape($UnescapedString).Replace('/', '\/') }

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

  # Initialization
  
  $Result = [ordered]@{
    IGDB = $null
    PCGW = $null
  }

  $Result.PCGW = Get-MWPage $Name -Wikitext

  if (-not $Result.PCGW)
  {
    Write-Warning "Found no PCGW page for $Name."
    return
  }

  $Result.IGDB = Get-IGDBGame -Where "name = `"$Name`" & platforms = (6)" -Fields 'cover.*'

  if (-not $Result.IGDB)
  {
    Write-Warning "Found no cover on IGDB for $Name."
    return
  }

  # Process
  
  $StatusCode = 200
  $Link       = "https:" + ($Result.IGDB.cover[0].url.Replace('t_thumb', 't_original'))
  $ext        = $Link.Split('.')[-1]

  # IGDB typically uploads t_original as PNG apparently, so lets just assume most of 'em are PNGs by default
  if ($ext -eq 'jpg')
  { $ext = 'png' }

  $FilePath   = ".\cover.$ext"
  try {
    Write-Verbose "Downloading $Link..."
    Invoke-WebRequest -Uri $Link -Method GET -UseBasicParsing -DisableKeepAlive -OutFile $FilePath
  } catch {
    $StatusCode = $_.Exception.response.StatusCode.value__
  }

  if ($StatusCode -ne 200)
  {
    Write-Warning "Failed to download $Link !"
    return
  }

  # -WhatIf
  if ($WhatIf)
  {
    [Console]::BackgroundColor = 'Black'
    [Console]::ForegroundColor = 'Yellow'
    [Console]::WriteLine('What if: Performing maintenance on target "' + $Name + '".')
    [Console]::ResetColor()

    return $Result
  }

  # -Online
  else
  {
    if (-not (Test-Path $FilePath))
    {
      Write-Warning "Could not find $FilePath !"
      return
    }

    $FilePath = Get-Item $FilePath
    $PCGWFile = "$Name cover.$ext"

    $PCGWFile = $PCGWFile.Replace(': ', ' - ')
    $PCGWFile = $PCGWFile.Replace(':', ' ')
    $PCGWFile = $PCGWFile.Replace('  ', ' ')
    $PCGWFile = $PCGWFile.Trim(' - ')
    $PCGWFile = $PCGWFile.Trim()

    # Upload file to PCGW
    $UploadParams  = @{
      Name         = $PCGWFile
      File         = $FilePath.FullName
      FixExtension = $true
      Force        = $Force
      Comment      = "Cover for [[$Name]]."
    }
    $Result.Upload = Import-MWFile @UploadParams -JSON
    Remove-Item -Path $FilePath.FullName -Force

    # Update PCGW page on a successful upload
    if ($Result.Upload.upload.filename)
    {
      $FinalName = $Result.Upload.upload.filename
      $FinalName = $FinalName.Replace('_', ' ')
      $Result.PCGW.Wikitext = $Result.PCGW.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'cover' -Value $FinalName
      $Result.PCGW = Set-MWPage -Name $Name -Content $Result.PCGW.Wikitext -Bot -Minor -NoCreate -Summary 'Added game cover'
    }

    # File is a duplicate, attempt to link to that one instead
    elseif ($Result.Upload.errors.code -eq 'fileexists-no-change')
    {
      $FinalName = $Result.Upload.errors.text -replace '.*\[\[\:File\:(.*)\]\]\.', '$1'
      $Result.PCGW.Wikitext = $Result.PCGW.Wikitext | SetTemplateParameter 'Infobox game' -Parameter 'cover' -Value $FinalName
      $Result.PCGW = Set-MWPage -Name $Name -Content $Result.PCGW.Wikitext -Bot -Minor -NoCreate -Summary 'Added game cover'
    }

    return $Result
  }
}

End {
  if ($ScriptConnectedAPI)
  { Disconnect-MWSession }

  if ($ScriptLoadedModule)
  { Remove-Module -Name 'MediaWiki' }
}