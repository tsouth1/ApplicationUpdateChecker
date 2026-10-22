<#
.SYNOPSIS
    Reports which third-party applications have a newer version available from the vendor.

.DESCRIPTION
    Builds a list of applications from a CSV file or from applications deployed in
    Configuration Manager (ConfigMgr), then looks up the latest version of each one.

    Sources (value for the Source column):
      Evergreen   Evergreen PowerShell module. Matched on AppName, or SourceId if supplied.
      Nevergreen  Nevergreen PowerShell module. Matched on AppName, or SourceId if supplied.
      WinGet      Windows Package Manager default source (cdn.winget.microsoft.com).
                  SourceId = package identifier, e.g. Google.Chrome
      MSStore     Windows Package Manager msstore source. SourceId = Store product ID.
      WinGetRepo  Community repository on GitHub (microsoft/winget-pkgs), read via the GitHub API.
                  SourceId = package identifier, e.g. Google.Chrome
      Chocolatey  Chocolatey community repository. SourceId = package id, e.g. googlechrome
      GitHub      Latest non-prerelease GitHub release. SourceId = owner/repo, e.g. microsoft/terminal
      Scoop       Scoop buckets (main, extras, versions). SourceId = app or bucket/app,
                  e.g. git or extras/notepad2

    When Source is blank the script tries Evergreen and Nevergreen by name. If SourceId is
    also supplied it adds other sources based on its format: owner/repo goes to GitHub, a dotted
    identifier goes to WinGet and WinGetRepo, and a plain name goes to Chocolatey. Scoop and
    MSStore are never tried automatically; set Source explicitly for those.

    -SearchMode FirstMatch (default) returns the first source that answers.
    -SearchMode Highest queries every applicable source and reports the highest version.
    The Candidates column shows what each source returned.

    Outputs written to -OutputDirectory:
      UpdateStatus_<timestamp>.csv    every application and its status
      UpdatesRequired_<timestamp>.csv applications that need an update only
      UpdateReport_<timestamp>.html   summary with the update list and items needing review

    With -FromSccm the script only reads from ConfigMgr. It writes an import CSV you can edit
    (fill in Source and SourceId for apps that did not map) and re-run with -CsvPath.

.PARAMETER CsvPath
    CSV describing the applications. Columns:
      AppName          (required) Display name, or Evergreen/Nevergreen name when Source is one of those
      Publisher        (optional) Informational only
      InstalledVersion (optional) Version currently deployed. Blank results in status Unknown
      Source           (optional) One of the source names above. Blank = automatic
      SourceId         (optional) Identifier for the chosen source (see above)
      Architecture     (optional) x64 | x86 | arm64. Defaults to -DefaultArchitecture
      Type             (optional) Evergreen/Nevergreen type filter, e.g. msi, exe, msix
      Channel          (optional) Evergreen/Nevergreen channel filter, e.g. stable

.PARAMETER FromSccm
    Read applications targeted by deployments in ConfigMgr. Requires the console on this machine.

.PARAMETER SiteCode
    ConfigMgr site code, e.g. PS1. Required with -FromSccm.

.PARAMETER SiteServer
    ConfigMgr site server FQDN. Required with -FromSccm.

.PARAMETER OutputDirectory
    Folder for reports. Defaults to .\UpdateReports

.PARAMETER DefaultArchitecture
    Architecture used when a row does not specify one. Defaults to x64.

.PARAMETER SearchMode
    FirstMatch (default) or Highest. See DESCRIPTION.

.PARAMETER Sources
    One or more source names to try, in the order given, for rows whose Source column is blank.
    Example: -Sources GitHub,Scoop,WinGetRepo. A row's own Source value always takes precedence.
    Each source needs its identifier: SourceId for WinGet, WinGetRepo, Chocolatey, GitHub and Scoop,
    or AppName/SourceId for Evergreen and Nevergreen. Rows missing that identifier are recorded
    in the Error column. When -Sources is omitted, the automatic selection described above applies.

.EXAMPLE
    .\Get-ThirdPartyUpdateReport.ps1 -CsvPath .\apps-sample.csv

.EXAMPLE
    .\Get-ThirdPartyUpdateReport.ps1 -CsvPath .\apps-sample.csv -SearchMode Highest

.EXAMPLE
    .\Get-ThirdPartyUpdateReport.ps1 -CsvPath .\apps-sample.csv -Sources Evergreen,WinGetRepo,GitHub -SearchMode Highest

.EXAMPLE
    .\Get-ThirdPartyUpdateReport.ps1 -FromSccm -SiteCode PS1 -SiteServer cm01.contoso.local

.NOTES
    Requires PowerShell 5.1 or later.
    Optional: Evergreen and Nevergreen modules, winget (App Installer), choco.
    GitHub-based sources (GitHub, WinGetRepo, Scoop) are rate-limited to about 60 unauthenticated
    requests per hour. Set the GITHUB_TOKEN environment variable to raise the limit.
#>
[CmdletBinding()]
param(
    [string]$CsvPath,
    [switch]$FromSccm,
    [string]$SiteCode,
    [string]$SiteServer,
    [string]$OutputDirectory = (Join-Path (Get-Location) 'UpdateReports'),
    [ValidateSet('x64', 'x86', 'arm64')]
    [string]$DefaultArchitecture = 'x64',
    [ValidateSet('FirstMatch', 'Highest')]
    [string]$SearchMode = 'FirstMatch',
    [ValidateSet('Evergreen', 'Nevergreen', 'WinGet', 'MSStore', 'WinGetRepo', 'Chocolatey', 'GitHub', 'Scoop')]
    [string[]]$Sources
)

# Scoop buckets used by the Scoop source. Branch is master for all three.
$script:ScoopBuckets = @{
    'main'     = 'ScoopInstaller/Main'
    'extras'   = 'ScoopInstaller/Extras'
    'versions' = 'ScoopInstaller/Versions'
}

#region Helpers

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message)
}

function ConvertTo-ComparableVersion {
    # Pulls the first numeric version out of a string such as "24.4.0.0 (3)" or "v126.0.6478.127"
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }

    $match = [regex]::Match($Text, '\d+(\.\d+){0,3}')
    if (-not $match.Success) { return $null }

    $parts = @($match.Value -split '\.')
    while ($parts.Count -lt 4) { $parts += '0' }

    $parsed = $null
    if ([version]::TryParse(($parts -join '.'), [ref]$parsed)) { return $parsed }
    return $null
}

function Compare-AppVersion {
    param([string]$Installed, [string]$Latest)

    $installedVersion = ConvertTo-ComparableVersion $Installed
    $latestVersion    = ConvertTo-ComparableVersion $Latest

    if (-not $latestVersion)    { return 'Unknown' }
    if (-not $installedVersion) { return 'Unknown' }

    if ($latestVersion -gt $installedVersion) { return 'UpdateAvailable' }
    return 'UpToDate'
}

function Get-GitHubHeaders {
    $headers = @{
        'User-Agent' = 'Get-ThirdPartyUpdateReport'
        'Accept'     = 'application/vnd.github+json'
    }
    if ($env:GITHUB_TOKEN) { $headers['Authorization'] = "Bearer $env:GITHUB_TOKEN" }
    return $headers
}

function Get-JsonFromUrl {
    param([string]$Url)
    $response = Invoke-WebRequest -Uri $Url -UseBasicParsing -Headers (Get-GitHubHeaders)
    return ($response.Content | ConvertFrom-Json)
}

function Select-LatestCandidate {
    # Shared filter for name-based module sources (Evergreen, Nevergreen)
    param($Candidates, $Row, [string]$Arch)

    $architecture = if ($Row.Architecture) { $Row.Architecture } else { $Arch }

    $Candidates |
        Where-Object { -not $_.Architecture -or $_.Architecture -eq $architecture } |
        Where-Object { -not $Row.Channel -or -not $_.Channel -or $_.Channel -eq $Row.Channel } |
        Where-Object { -not $Row.Type -or -not $_.Type -or $_.Type -eq $Row.Type } |
        Sort-Object -Property @{ Expression = { ConvertTo-ComparableVersion $_.Version }; Descending = $true } |
        Select-Object -First 1
}

#endregion

#region Module-based sources

function Get-LatestFromEvergreen {
    param($Row, [string]$Arch)

    if (-not (Get-Command Get-EvergreenApp -ErrorAction SilentlyContinue)) {
        throw 'Evergreen module not installed. Run: Install-Module Evergreen'
    }

    $name = if ($Row.SourceId) { $Row.SourceId } else { $Row.AppName }
    $candidates = @(Get-EvergreenApp -Name $name -ErrorAction Stop)

    $best = Select-LatestCandidate -Candidates $candidates -Row $Row -Arch $Arch
    if (-not $best) { return $null }

    [pscustomobject]@{ Version = $best.Version; DownloadUrl = $best.URI }
}

function Get-LatestFromNevergreen {
    param($Row, [string]$Arch)

    if (-not (Get-Command Get-NevergreenApp -ErrorAction SilentlyContinue)) {
        throw 'Nevergreen module not installed. Run: Install-Module Nevergreen'
    }

    $name = if ($Row.SourceId) { $Row.SourceId } else { $Row.AppName }
    $candidates = @(Get-NevergreenApp -Name $name -ErrorAction Stop)

    $best = Select-LatestCandidate -Candidates $candidates -Row $Row -Arch $Arch
    if (-not $best) { return $null }

    [pscustomobject]@{ Version = $best.Version; DownloadUrl = $best.URI }
}

#endregion

#region Windows Package Manager and Chocolatey sources

function Get-LatestFromWinGet {
    # Used for the 'winget' source (cdn.winget.microsoft.com) and the 'msstore' source
    param([string]$Id, [string]$Source = 'winget')

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw 'winget not found. Install App Installer from the Microsoft Store.'
    }
    if (-not $Id) { throw "WinGet ($Source) lookup needs SourceId." }

    $output = & winget show --id $Id --exact --source $Source --accept-source-agreements --disable-interactivity 2>&1
    if ($LASTEXITCODE -ne 0) { return $null }

    # "Version:" is the English label. Other Windows display languages need a different pattern.
    $version = $null
    foreach ($line in $output) {
        if ($line -match '^\s*Version:\s*(\S.*?)\s*$') {
            $version = $Matches[1]
            break
        }
    }
    if (-not $version) { return $null }

    [pscustomobject]@{ Version = $version; DownloadUrl = "winget:$Source/$Id" }
}

function Get-LatestFromWinGetRepo {
    # Reads the community manifest folders in microsoft/winget-pkgs via the GitHub API.
    # Each version is a folder, so the highest folder name is the latest published version.
    param([string]$Id)

    if (-not $Id -or $Id -notmatch '\.') {
        throw 'WinGetRepo needs SourceId as a WinGet identifier, e.g. Google.Chrome'
    }

    $segments = @($Id -split '\.')
    $first = $segments[0].Substring(0, 1).ToLower()
    $path = "manifests/$first/" + ($segments -join '/')

    $listing = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-pkgs/contents/$path" -Headers (Get-GitHubHeaders)

    $versions = @()
    foreach ($item in $listing) {
        if ($item.type -ne 'dir') { continue }
        $parsed = ConvertTo-ComparableVersion $item.name
        if ($parsed) { $versions += [pscustomobject]@{ Name = $item.name; Parsed = $parsed } }
    }

    $latest = $versions | Sort-Object -Property Parsed -Descending | Select-Object -First 1
    if (-not $latest) { return $null }

    [pscustomobject]@{
        Version     = $latest.Name
        DownloadUrl = "https://github.com/microsoft/winget-pkgs/tree/master/$path/$($latest.Name)"
    }
}

function Get-LatestFromChocolatey {
    param([string]$Id)

    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        throw 'choco not found. Install Chocolatey from https://chocolatey.org/install'
    }
    if (-not $Id) { throw 'Chocolatey lookup needs SourceId (package id, e.g. googlechrome).' }

    # --limit-output returns lines in the form name|version
    $lines = & choco search $Id --exact --limit-output 2>$null
    foreach ($line in $lines) {
        $parts = $line -split '\|'
        if ($parts.Count -ge 2 -and $parts[0] -ieq $Id) {
            return [pscustomobject]@{ Version = $parts[1]; DownloadUrl = "choco:$Id" }
        }
    }
    return $null
}

#endregion

#region GitHub and Scoop sources

function Get-LatestFromGitHub {
    param([string]$Id)

    if ($Id -notmatch '^[^/\s]+/[^/\s]+$') {
        throw 'GitHub needs SourceId in the form owner/repo, e.g. microsoft/terminal'
    }

    # /releases/latest excludes drafts and prereleases
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$Id/releases/latest" -Headers (Get-GitHubHeaders)

    $asset = $release.assets |
        Where-Object { $_.name -match '(x64|amd64|x86_64).*\.(msi|exe|msix)$' } |
        Select-Object -First 1

    $url = if ($asset) { $asset.browser_download_url } else { $release.html_url }

    [pscustomobject]@{
        Version     = ($release.tag_name -replace '^[vV]', '')
        DownloadUrl = $url
    }
}

function Get-LatestFromScoop {
    # SourceId is "app" (main bucket) or "bucket/app".
    # main and extras hold the current manifest, so its version field is the latest.
    # versions holds pinned older releases, so the highest pinned version is returned.
    param([string]$Id)

    if (-not $Id) { throw 'Scoop lookup needs SourceId: app or bucket/app, e.g. extras/notepad2' }

    $bucket = 'main'
    $app = $Id
    if ($Id -match '^([^/]+)/(.+)$') {
        $bucket = $Matches[1].ToLower()
        $app = $Matches[2]
    }
    if (-not $script:ScoopBuckets.ContainsKey($bucket)) {
        throw "Unknown Scoop bucket '$bucket'. Use main, extras or versions."
    }
    $repo = $script:ScoopBuckets[$bucket]

    if ($bucket -eq 'versions') {
        $listing = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/contents/bucket" -Headers (Get-GitHubHeaders)
        $pattern = '^' + [regex]::Escape($app) + '[@-](.+)\.json$'

        $pinned = @()
        foreach ($item in $listing) {
            if ($item.name -match $pattern) {
                $pinned += [pscustomobject]@{ Version = $Matches[1]; Url = $item.download_url }
            }
        }
        $latest = $pinned |
            Sort-Object -Property @{ Expression = { ConvertTo-ComparableVersion $_.Version }; Descending = $true } |
            Select-Object -First 1
        if (-not $latest) { return $null }

        return [pscustomobject]@{ Version = $latest.Version; DownloadUrl = $latest.Url }
    }

    $manifestUrl = "https://raw.githubusercontent.com/$repo/master/bucket/$app.json"
    $manifest = Get-JsonFromUrl -Url $manifestUrl

    $download = $manifest.url
    if (-not $download -and $manifest.architecture) { $download = $manifest.architecture.'64bit'.url }
    if ($download -is [array]) { $download = $download[0] }
    if (-not $download) { $download = $manifestUrl }

    [pscustomobject]@{ Version = [string]$manifest.version; DownloadUrl = $download }
}

#endregion

#region Source resolution

function Get-IdBasedSources {
    # Chooses sources for a SourceId when Source is not set
    param([string]$Id)

    if ($Id -match '^[^/\s]+/[^/\s]+$') { return @('GitHub') }
    if ($Id -match '\.')                { return @('WinGet', 'WinGetRepo') }
    return @('Chocolatey')
}

function Invoke-SourceLookup {
    param([string]$Source, $Row, [string]$Arch)

    switch ($Source) {
        'Evergreen'  { Get-LatestFromEvergreen -Row $Row -Arch $Arch }
        'Nevergreen' { Get-LatestFromNevergreen -Row $Row -Arch $Arch }
        'WinGet'     { Get-LatestFromWinGet -Id $Row.SourceId -Source 'winget' }
        'MSStore'    { Get-LatestFromWinGet -Id $Row.SourceId -Source 'msstore' }
        'WinGetRepo' { Get-LatestFromWinGetRepo -Id $Row.SourceId }
        'Chocolatey' { Get-LatestFromChocolatey -Id $Row.SourceId }
        'GitHub'     { Get-LatestFromGitHub -Id $Row.SourceId }
        'Scoop'      { Get-LatestFromScoop -Id $Row.SourceId }
        default      { throw "Unknown source '$Source'. Use Evergreen, Nevergreen, WinGet, MSStore, WinGetRepo, Chocolatey, GitHub or Scoop." }
    }
}

function Resolve-LatestVersion {
    param($Row, [string]$Arch, [string]$Mode, [string[]]$CommandLineSources)

    # Precedence: the row's Source column, then -Sources on the command line, then automatic choice
    $order = if ($Row.Source) {
        @($Row.Source)
    }
    elseif ($CommandLineSources) {
        @($CommandLineSources)
    }
    elseif ($Row.SourceId) {
        @('Evergreen', 'Nevergreen') + @(Get-IdBasedSources -Id $Row.SourceId)
    }
    else {
        @('Evergreen', 'Nevergreen')
    }

    $hits = @()
    $errors = @()

    foreach ($src in $order) {
        try {
            $hit = Invoke-SourceLookup -Source $src -Row $Row -Arch $Arch
            if ($hit -and $hit.Version) {
                $hits += [pscustomobject]@{ Source = $src; Version = $hit.Version; DownloadUrl = $hit.DownloadUrl }
                if ($Mode -eq 'FirstMatch') { break }
            }
        }
        catch {
            $errors += "${src}: $($_.Exception.Message)"
        }
    }

    if ($hits.Count -eq 0) {
        return [pscustomobject]@{
            Source = $null; Version = $null; DownloadUrl = $null
            Candidates = ''; Error = ($errors -join '; ')
        }
    }

    $best = $hits |
        Sort-Object -Property @{ Expression = { ConvertTo-ComparableVersion $_.Version }; Descending = $true } |
        Select-Object -First 1

    [pscustomobject]@{
        Source      = $best.Source
        Version     = $best.Version
        DownloadUrl = $best.DownloadUrl
        Candidates  = (($hits | ForEach-Object { "$($_.Source)=$($_.Version)" }) -join '; ')
        Error       = ($errors -join '; ')
    }
}

#endregion

#region ConfigMgr

function Get-SccmDeployedApplications {
    param([string]$SiteCode, [string]$SiteServer)

    if (-not $env:SMS_ADMIN_UI_PATH) {
        throw 'ConfigMgr console not detected (SMS_ADMIN_UI_PATH is not set). Run this on a machine with the console installed.'
    }

    # SMS_ADMIN_UI_PATH points to ...\bin\i386. The module sits one level up in ...\bin
    $modulePath = Join-Path (Split-Path -Parent $env:SMS_ADMIN_UI_PATH) 'ConfigurationManager.psd1'
    if (-not (Get-Module -Name ConfigurationManager)) {
        Import-Module $modulePath -ErrorAction Stop
    }

    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
    }

    Push-Location "$($SiteCode):\"
    try {
        $names = Get-CMDeployment |
            Where-Object { $_.SoftwareName } |
            Select-Object -ExpandProperty SoftwareName -Unique

        foreach ($name in $names) {
            Write-Log "Reading ConfigMgr application '$name'"
            $app = Get-CMApplication -Name $name -Fast -ErrorAction SilentlyContinue

            [pscustomobject]@{
                AppName          = if ($app) { $app.LocalizedDisplayName } else { $name }
                Publisher        = if ($app) { $app.Manufacturer } else { '' }
                InstalledVersion = if ($app) { $app.SoftwareVersion } else { '' }
                Source           = ''
                SourceId         = ''
                Architecture     = ''
                Type             = ''
                Channel          = ''
            }
        }
    }
    finally {
        Pop-Location
    }
}

#endregion

#region Main

if ([bool]$CsvPath -eq [bool]$FromSccm) {
    throw 'Specify either -CsvPath or -FromSccm (with -SiteCode and -SiteServer), not both.'
}
if ($FromSccm -and (-not $SiteCode -or -not $SiteServer)) {
    throw '-FromSccm requires -SiteCode and -SiteServer.'
}
if ($CsvPath -and -not (Test-Path $CsvPath)) {
    throw "CSV not found: $CsvPath"
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'

if ($FromSccm) {
    Write-Log "Reading deployed applications from ConfigMgr site $SiteCode ($SiteServer)"
    $rows = @(Get-SccmDeployedApplications -SiteCode $SiteCode -SiteServer $SiteServer)

    $importPath = Join-Path $OutputDirectory "SccmImport_$stamp.csv"
    $rows | Export-Csv -Path $importPath -NoTypeInformation
    Write-Log "Import list written to $importPath. Fill in Source and SourceId for unmapped apps, then re-run with -CsvPath."
}
else {
    Write-Log "Reading application list from $CsvPath"
    $rows = @(Import-Csv -Path $CsvPath)
}

Write-Log "Checking $($rows.Count) application(s) using $SearchMode mode"

$results = foreach ($row in $rows) {
    if (-not $row.AppName) { continue }

    Write-Log "Checking $($row.AppName)"
    $latest = Resolve-LatestVersion -Row $row -Arch $DefaultArchitecture -Mode $SearchMode -CommandLineSources $Sources

    $status = if (-not $latest.Version) {
        'NoSourceFound'
    }
    else {
        Compare-AppVersion -Installed $row.InstalledVersion -Latest $latest.Version
    }

    [pscustomobject]@{
        Application      = $row.AppName
        Publisher        = $row.Publisher
        InstalledVersion = $row.InstalledVersion
        LatestVersion    = $latest.Version
        Status           = $status
        Source           = $latest.Source
        Candidates       = $latest.Candidates
        SourceId         = $row.SourceId
        DownloadUrl      = $latest.DownloadUrl
        Error            = $latest.Error
        CheckedUtc       = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    }
}

$results = @($results)
$updates = @($results | Where-Object { $_.Status -eq 'UpdateAvailable' })
$review  = @($results | Where-Object { $_.Status -in @('NoSourceFound', 'Unknown') })

$fullCsv    = Join-Path $OutputDirectory "UpdateStatus_$stamp.csv"
$updatesCsv = Join-Path $OutputDirectory "UpdatesRequired_$stamp.csv"
$htmlPath   = Join-Path $OutputDirectory "UpdateReport_$stamp.html"

$results | Export-Csv -Path $fullCsv -NoTypeInformation
$updates | Export-Csv -Path $updatesCsv -NoTypeInformation

$summary = ($results | Group-Object Status | ForEach-Object { "$($_.Name): $($_.Count)" }) -join ' | '
$preContent = @(
    "<h2>Applications requiring update: $($updates.Count) of $($results.Count)</h2>",
    "<p>Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm') using $SearchMode mode</p>",
    "<p>Status counts: $summary</p>",
    '<h3>Updates available</h3>'
)

$updateTable = $updates | Select-Object Application, Publisher, InstalledVersion, LatestVersion, Source, DownloadUrl
$reviewTable = $review  | Select-Object Application, Publisher, InstalledVersion, Status, Error

$htmlBody = @()
$htmlBody += $updateTable | ConvertTo-Html -Fragment
$htmlBody += '<h3>Needs review (no source found or installed version unknown)</h3>'
$htmlBody += $reviewTable | ConvertTo-Html -Fragment

ConvertTo-Html -Title 'Third-party application update report' -PreContent $preContent -Body $htmlBody |
    Set-Content -Path $htmlPath -Encoding UTF8

Write-Log "Done. $($updates.Count) update(s) required, $($review.Count) item(s) need review."
Write-Log "Full status:      $fullCsv"
Write-Log "Updates required: $updatesCsv"
Write-Log "HTML report:      $htmlPath"

#endregion