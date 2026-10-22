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

    Rows with a blank Publisher, or a Publisher matching *lear*, are skipped (status Skipped)
    unless the row already has Source or SourceId filled in.

    For other rows with a blank SourceId, the script searches the WinGet source by name
    (winget search --name) and stores the best match in a WinGetId column. A match is kept
    when the cleaned-up AppName equals the package name, or when the package name starts with
    it and the package Id matches the Publisher. WinGet and WinGetRepo use WinGetId when
    SourceId is blank. The mapped list is written to MappedImport_<timestamp>.csv so matches
    can be reviewed and corrected; rows that already have a WinGetId are not searched again.

    When Evergreen is in use, the same rows are also matched against the Evergreen catalog
    (Find-EvergreenApp) by Name, e.g. GoogleChrome, or Application, e.g. Google Chrome Enterprise.
    The match is stored in an EvergreenName column, which Evergreen uses in place of AppName.
    Rows that already have an EvergreenName are not matched again.
    Use -NoIdLookup to turn off both lookups.

    -SearchMode FirstMatch (default) returns the first source that answers.
    -SearchMode Highest queries every applicable source and reports the highest version.
    The Candidates column shows what each source returned.

    Outputs written to -OutputDirectory:
      UpdateStatus_<timestamp>.csv    every application and its status
      UpdatesRequired_<timestamp>.csv applications that need an update only
      UpdateReport_<timestamp>.html   summary with the update list and items needing review
      MappedImport_<timestamp>.csv    input rows with WinGetId and IdMatch (when lookup ran)

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
      WinGetId         (optional) WinGet identifier found by the automatic lookup
      EvergreenName    (optional) Evergreen app name found by the automatic lookup

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

.PARAMETER NoIdLookup
    Do not look up WinGet ids or Evergreen names for rows with a blank SourceId.

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
    Before checking, the script prepares what the sources in use need: it installs the Evergreen
    and Nevergreen modules from PSGallery when missing (CurrentUser scope, or AllUsers when
    elevated), imports them, and runs Update-Evergreen to refresh Evergreen's app definitions.
    It warns when winget (App Installer), choco or GITHUB_TOKEN are missing but needed.
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
    [string[]]$Sources,
    [switch]$NoIdLookup
)

# Scoop buckets used by the Scoop source. Branch is master for all three.
$script:ScoopBuckets = @{
    'main'     = 'ScoopInstaller/Main'
    'extras'   = 'ScoopInstaller/Extras'
    'versions' = 'ScoopInstaller/Versions'
}

# winget search results by search name, shared by rows that clean up to the same name
$script:WinGetSearchCache = @{}

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

    # Without a Channel on the row, pre-release builds are left out, e.g. Chrome Canary or VS Code Insider
    $preRelease = '^(?i)(beta|dev|canary|insiders?|preview|nightly|alpha|rc|test|unstable)$'
    $filtered = @($Candidates |
        Where-Object { $_.Version } |
        Where-Object { if ($Row.Channel) { -not $_.Channel -or $_.Channel -eq $Row.Channel } else { $_.Channel -notmatch $preRelease } } |
        Where-Object { -not $Row.Type -or -not $_.Type -or $_.Type -eq $Row.Type })

    # Prefer the requested architecture, then architecture-neutral builds, then x86 (as for
    # apps such as WinSCP that only publish x86). A row with Architecture set gets that only.
    $ranked = foreach ($candidate in $filtered) {
        $rank = if (-not $candidate.Architecture -or $candidate.Architecture -match '^(?i)(neutral|any|all)$') { 1 }
                elseif ($candidate.Architecture -eq $architecture) { 0 }
                elseif (-not $Row.Architecture -and $candidate.Architecture -eq 'x86') { 2 }
                else { continue }
        [pscustomobject]@{ Rank = $rank; Candidate = $candidate }
    }

    $best = $ranked |
        Sort-Object -Property Rank, @{ Expression = { ConvertTo-ComparableVersion $_.Candidate.Version }; Descending = $true } |
        Select-Object -First 1

    if (-not $best -and @($Candidates).Count) {
        throw "$(@($Candidates).Count) result(s), none for architecture $architecture" +
            $(if ($Row.Channel) { ", channel $($Row.Channel)" } else { ' in a stable channel' }) +
            $(if ($Row.Type) { ", type $($Row.Type)" } else { '' })
    }
    if ($best) { $best.Candidate }
}

#endregion

#region Prerequisites

function Install-RequiredModule {
    # Installs a PSGallery module when missing, then imports it. Returns $true when it is loaded.
    param([string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Log "$Name module not found, installing from PSGallery"
        try {
            # Windows PowerShell 5.1 defaults to TLS 1.0, which PSGallery rejects
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
                [Security.Principal.WindowsBuiltInRole]::Administrator)
            $scope = if ($isAdmin) { 'AllUsers' } else { 'CurrentUser' }

            # Install-Module needs the NuGet provider; installing it here avoids an interactive prompt
            $nuget = Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue |
                Where-Object { $_.Version -ge [version]'2.8.5.201' }
            if (-not $nuget) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope $scope -Force -ErrorAction Stop | Out-Null
            }

            # -Force installs from PSGallery without changing its trust setting
            Install-Module -Name $Name -Repository PSGallery -Scope $scope -Force -AllowClobber -ErrorAction Stop
            Write-Log "Installed $Name for $scope"
        }
        catch {
            Write-Log "Could not install $Name module: $($_.Exception.Message)" 'WARN'
            return $false
        }
    }

    try {
        Import-Module -Name $Name -ErrorAction Stop
        $version = (Get-Module -Name $Name).Version
        Write-Log "$Name module $version loaded"
        return $true
    }
    catch {
        Write-Log "Could not import $Name module: $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Initialize-Prerequisites {
    # Prepares the modules and tools used by the sources this run will query. A missing
    # prerequisite is reported here and again in the Error column of each affected row.
    param([string[]]$UsedSources, [bool]$IdLookup)

    if ('Evergreen' -in $UsedSources -and (Install-RequiredModule -Name 'Evergreen')) {
        # Evergreen keeps app definitions in a local cache (%LocalAppData%\Evergreen) that
        # Update-Evergreen downloads and refreshes. Without it Get-EvergreenApp knows no apps.
        try {
            Write-Log 'Updating Evergreen app definitions (Update-Evergreen)'
            Update-Evergreen -ErrorAction Stop
        }
        catch {
            $cache = Get-EvergreenAppsPath -ErrorAction SilentlyContinue
            if ($cache -and (Test-Path $cache)) {
                Write-Log "Update-Evergreen failed, using existing cache at ${cache}: $($_.Exception.Message)" 'WARN'
            }
            else {
                Write-Log "Update-Evergreen failed and no local cache exists. Evergreen lookups will fail: $($_.Exception.Message)" 'WARN'
            }
        }
    }

    # Nevergreen has its app definitions built in and needs no update step
    if ('Nevergreen' -in $UsedSources) { [void](Install-RequiredModule -Name 'Nevergreen') }

    if (($IdLookup -or @($UsedSources | Where-Object { $_ -in 'WinGet', 'MSStore' }).Count) -and
        -not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log 'winget not found. WinGet/MSStore lookups and the automatic ID lookup need App Installer from the Microsoft Store.' 'WARN'
    }

    if ('Chocolatey' -in $UsedSources -and -not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Log 'choco not found. Chocolatey lookups need Chocolatey from https://chocolatey.org/install' 'WARN'
    }

    if (@($UsedSources | Where-Object { $_ -in 'GitHub', 'WinGetRepo', 'Scoop' }).Count -and -not $env:GITHUB_TOKEN) {
        Write-Log 'GITHUB_TOKEN is not set. GitHub, WinGetRepo and Scoop share a limit of about 60 GitHub API requests per hour.' 'WARN'
    }
}

#endregion

#region Module-based sources

function Get-LatestFromEvergreen {
    param($Row, [string]$Arch)

    if (-not (Get-Command Get-EvergreenApp -ErrorAction SilentlyContinue)) {
        throw 'Evergreen module not installed. Run: Install-Module Evergreen'
    }

    # SourceId is an Evergreen name when Source is Evergreen; otherwise it may be another
    # source's id, so the name found by the automatic lookup comes first
    $name = if ($Row.Source -eq 'Evergreen' -and $Row.SourceId) { $Row.SourceId }
            elseif ($Row.EvergreenName) { $Row.EvergreenName }
            elseif ($Row.SourceId) { $Row.SourceId }
            else { $Row.AppName }
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

    # The folder can also hold related packages, e.g. DesktopRuntime/8/x86 or
    # SQLServerManagementStudio/22. Version folders start with a digit, and when any
    # dotted version exists, dotless names like 22 are treated as package folders.
    $versions = @()
    foreach ($item in $listing) {
        if ($item.type -ne 'dir' -or $item.name -notmatch '^v?\d') { continue }
        $parsed = ConvertTo-ComparableVersion $item.name
        if ($parsed) { $versions += [pscustomobject]@{ Name = $item.name; Parsed = $parsed } }
    }
    $dotted = @($versions | Where-Object { $_.Name -match '\.' })
    if ($dotted.Count) { $versions = $dotted }

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

#region Automatic WinGet ID lookup

function Get-SkipReason {
    # Rows that are mapped by hand are never skipped
    param($Row)

    if ($Row.Source -or $Row.SourceId) { return $null }
    if ([string]::IsNullOrWhiteSpace($Row.Publisher)) { return 'Skipped: no publisher' }
    if ($Row.Publisher -like '*lear*') { return "Skipped: publisher '$($Row.Publisher)'" }
    return $null
}

function ConvertTo-MatchKey {
    # Lowercase letters and digits only, so "Notepad++ (64-bit)" and "notepad 64bit" compare cleanly
    param([string]$Text)
    if (-not $Text) { return '' }
    return ($Text.ToLower() -replace '[^a-z0-9]', '')
}

function Get-SearchName {
    # Strips installer wording, versions and architecture from a ConfigMgr display name,
    # e.g. "TortoiseSVN 1.14.9.29743 (64 bit)" becomes "TortoiseSVN"
    param([string]$AppName)

    $name = $AppName -replace '(?i)^\s*(install|uninstall)\s+', ''
    $name = $name -replace '(?i)\bversion\s+\S+', ''
    $name = $name -replace '\bv?\d+(\.\d+)+\S*', ''
    $name = $name -replace '(?i)\(?\b(x64|x86|amd64|arm64|64[- ]?bit|32[- ]?bit)\b\)?', ''
    $name = $name -replace '\(\s*\)', ''
    $name = $name -replace '\s+-\s+.*$', ''
    return ($name -replace '\s+', ' ').Trim(' -_')
}

function Get-PublisherKey {
    # First significant word of the publisher, e.g. "Microsoft Corporation" becomes "microsoft"
    param([string]$Publisher)

    $words = @($Publisher -split '[\s,\.]+' | Where-Object {
        $_ -and $_ -notmatch '^(?i)(inc|corp|corporation|llc|ltd|limited|gmbh|ag|co|company|the|s\.?a)$'
    })
    if ($words.Count -eq 0) { return '' }
    return (ConvertTo-MatchKey $words[0])
}

function Search-WinGetByName {
    # Returns Name/Id pairs from "winget search --name". The output is a fixed-width table,
    # so columns are cut at the positions of the header words above the dashed line.
    param([string]$Name)

    if ($script:WinGetSearchCache.ContainsKey($Name)) { return $script:WinGetSearchCache[$Name] }
    $script:WinGetSearchCache[$Name] = @(Invoke-WinGetSearch -Name $Name)
    return $script:WinGetSearchCache[$Name]
}

function Invoke-WinGetSearch {
    param([string]$Name)

    $output = @(& winget search --name $Name --source winget --accept-source-agreements --disable-interactivity 2>&1 |
        ForEach-Object { [string]$_ })
    if ($LASTEXITCODE -ne 0) { return @() }

    $dashIndex = -1
    for ($i = 1; $i -lt $output.Count; $i++) {
        if ($output[$i] -match '^-{10,}\s*$') { $dashIndex = $i; break }
    }
    if ($dashIndex -lt 1) { return @() }

    $columns = [regex]::Matches($output[$dashIndex - 1], '\S+')
    if ($columns.Count -lt 3) { return @() }
    $idStart = $columns[1].Index
    $versionStart = $columns[2].Index

    foreach ($line in $output[($dashIndex + 1)..($output.Count - 1)]) {
        if ($line.Length -le $versionStart) { continue }
        $id = $line.Substring($idStart, $versionStart - $idStart).Trim()
        # Skip truncated ids and rows misaligned by wide characters in the name
        if ($id -notmatch '^[^\s.]+(\.[^\s.]+)+$' -or $id.Contains([char]0x2026)) { continue }
        [pscustomobject]@{ Name = $line.Substring(0, $idStart).Trim(); Id = $id }
    }
}

function Find-WinGetId {
    # Picks the best WinGet package for a row. Returns $null when nothing is confident enough.
    param($Row)

    $query = Get-SearchName $Row.AppName
    if ($query.Length -lt 3) { return $null }

    $candidates = @(Search-WinGetByName -Name $query)
    if ($candidates.Count -eq 0) {
        # Second try without a trailing number, e.g. "FormingSuite 2018" becomes "FormingSuite"
        $shorter = ($query -replace '\s+\d+$', '').Trim()
        if ($shorter -ne $query -and $shorter.Length -ge 3) {
            $query = $shorter
            $candidates = @(Search-WinGetByName -Name $query)
        }
    }
    if ($candidates.Count -eq 0) { return $null }

    $queryKey = ConvertTo-MatchKey $query
    $publisherKey = Get-PublisherKey $Row.Publisher

    # Numbers in the original name (".NET 8.0.24" gives 8, 0, 24) and its architecture, used to
    # reject ids pinned to another major version or architecture, e.g. DesktopRuntime.9 or VCRedist.2013.x86
    $nameNumbers = @([regex]::Matches($Row.AppName, '\d+') | ForEach-Object { $_.Value.TrimStart('0') })
    $nameArch = if ($Row.AppName -match '(?i)\b(x86|32[- ]?bit)\b') { 'x86' }
                elseif ($Row.AppName -match '(?i)\b(x64|amd64|64[- ]?bit)\b') { 'x64' }
                elseif ($Row.AppName -match '(?i)\barm64\b') { 'arm64' }
                else { '' }

    $scored = foreach ($candidate in $candidates) {
        $nameKey = ConvertTo-MatchKey $candidate.Name
        $idSegments = @($candidate.Id -split '\.')
        $idPublisherKey = ConvertTo-MatchKey $idSegments[0]

        $conflict = $false
        foreach ($segment in $idSegments[1..($idSegments.Count - 1)]) {
            if ($segment -match '^\d+$' -and $segment.TrimStart('0') -notin $nameNumbers) { $conflict = $true }
            # Years inside a segment, e.g. AccessDatabaseEngine2016
            foreach ($year in [regex]::Matches($segment, '(19|20)\d\d')) {
                if ($year.Value -notin $nameNumbers) { $conflict = $true }
            }
            if ($segment -match '^(?i)(x86|x64|arm64)$' -and $nameArch -and $segment -ne $nameArch) { $conflict = $true }
            if ($segment -match '^(?i)(beta|preview|nightly|weekly|dev|canary|insiders?|alpha|rc)$' -and
                $Row.AppName -notmatch "(?i)\b$segment\b") { $conflict = $true }
        }

        $nameMatch = if ($nameKey -eq $queryKey) { 'exact' } elseif ($nameKey.StartsWith($queryKey)) { 'prefix' } else { '' }
        $publisherMatch = $publisherKey.Length -ge 3 -and $idPublisherKey -and
            ($idPublisherKey.Contains($publisherKey) -or $publisherKey.Contains($idPublisherKey))

        # Exact name wins; a prefix name only counts when the publisher also matches
        $score = 0
        if ($nameMatch -eq 'exact')  { $score += 6 }
        if ($nameMatch -eq 'prefix') { $score += 2 }
        if ($publisherMatch)         { $score += 3 }
        if ($conflict)               { $score -= 6 }

        [pscustomobject]@{
            Id = $candidate.Id; Name = $candidate.Name; Score = $score
            NameMatch = $nameMatch; PublisherMatch = $publisherMatch
        }
    }

    $best = $scored |
        Where-Object { $_.Score -ge 5 } |
        Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, @{ Expression = { $_.Name.Length } } |
        Select-Object -First 1
    if (-not $best) { return $null }

    $detail = "$($best.NameMatch) name '$($best.Name)'"
    if ($best.PublisherMatch) { $detail += ', publisher match' }
    [pscustomobject]@{ Id = $best.Id; Detail = "winget search '$query': $detail" }
}

function Find-EvergreenName {
    # Matches a row against the Evergreen catalog from Find-EvergreenApp, comparing the cleaned
    # AppName with each entry's Name (e.g. GoogleChrome) and Application (e.g. Google Chrome Enterprise).
    # Returns $null when nothing is confident enough.
    param($Row, $Catalog)

    $query = Get-SearchName $Row.AppName
    $queryKey = ConvertTo-MatchKey $query
    if ($queryKey.Length -lt 3) { return $null }

    $publisherKey = Get-PublisherKey $Row.Publisher
    # "VLC Media Player" from VideoLAN also matches "VideoLAN VLC Media Player"
    $withPublisher = if ($publisherKey -and -not $queryKey.StartsWith($publisherKey)) { $publisherKey + $queryKey } else { '' }
    $nameNumbers = @([regex]::Matches($Row.AppName, '\d+') | ForEach-Object { $_.Value.TrimStart('0') })

    $scored = foreach ($entry in $Catalog) {
        $nameKey = ConvertTo-MatchKey $entry.Name
        $appKey = ConvertTo-MatchKey $entry.Application

        $match = if ($queryKey -in $nameKey, $appKey -or ($withPublisher -and $withPublisher -in $nameKey, $appKey)) { 'exact' }
                 elseif ($appKey.StartsWith($queryKey) -or $nameKey.StartsWith($queryKey) -or
                         ($withPublisher -and ($appKey.StartsWith($withPublisher) -or $nameKey.StartsWith($withPublisher)))) { 'prefix' }
                 else { '' }
        if (-not $match) { continue }

        $publisherMatch = $publisherKey.Length -ge 3 -and ($appKey.StartsWith($publisherKey) -or $nameKey.StartsWith($publisherKey))

        # Numbers in the Evergreen name that the AppName lacks point to another edition, e.g. 1Password7
        $conflict = @([regex]::Matches($entry.Name, '\d+') | Where-Object { $_.Value.TrimStart('0') -notin $nameNumbers }).Count -gt 0

        $score = 0
        if ($match -eq 'exact')  { $score += 6 }
        if ($match -eq 'prefix') { $score += 2 }
        if ($publisherMatch)     { $score += 3 }
        if ($conflict)           { $score -= 6 }

        [pscustomobject]@{ Name = $entry.Name; Application = $entry.Application; Score = $score; Match = $match; PublisherMatch = $publisherMatch }
    }

    $best = $scored |
        Where-Object { $_.Score -ge 5 } |
        Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, @{ Expression = { $_.Application.Length } } |
        Select-Object -First 1
    if (-not $best) { return $null }

    $detail = "$($best.Match) name '$($best.Application)'"
    if ($best.PublisherMatch) { $detail += ', publisher match' }
    [pscustomobject]@{ Name = $best.Name; Detail = "Find-EvergreenApp '$query': $detail" }
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

    # The automatic lookup fills WinGetId, used when SourceId is blank
    $winGetId = if ($Row.SourceId) { $Row.SourceId } else { $Row.WinGetId }

    switch ($Source) {
        'Evergreen'  { Get-LatestFromEvergreen -Row $Row -Arch $Arch }
        'Nevergreen' { Get-LatestFromNevergreen -Row $Row -Arch $Arch }
        'WinGet'     { Get-LatestFromWinGet -Id $winGetId -Source 'winget' }
        'MSStore'    { Get-LatestFromWinGet -Id $Row.SourceId -Source 'msstore' }
        'WinGetRepo' { Get-LatestFromWinGetRepo -Id $winGetId }
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
    elseif ($Row.WinGetId) {
        @('Evergreen', 'Nevergreen', 'WinGet', 'WinGetRepo')
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

# Work out which sources this run can reach, so only their prerequisites are prepared.
# Rows with a Source value use it; other rows use -Sources, or the automatic choice, which
# starts with Evergreen and Nevergreen and adds sources based on SourceId or WinGetId.
$usedSources = @()
foreach ($row in $rows) {
    if (-not $row.AppName -or (Get-SkipReason $row)) { continue }
    if ($row.Source) { $usedSources += $row.Source; continue }
    if ($Sources) { $usedSources += $Sources; continue }
    $usedSources += 'Evergreen', 'Nevergreen'
    if ($row.SourceId) { $usedSources += Get-IdBasedSources -Id $row.SourceId }
    elseif ($row.WinGetId -or -not $NoIdLookup) { $usedSources += 'WinGet', 'WinGetRepo' }
}
$usedSources = @($usedSources | Select-Object -Unique)
Write-Log "Checking prerequisites for: $($usedSources -join ', ')"
Initialize-Prerequisites -UsedSources $usedSources -IdLookup (-not $NoIdLookup)

# Every row gets the lookup columns so the mapped CSV has consistent columns
foreach ($row in $rows) {
    foreach ($column in 'WinGetId', 'IdMatch', 'EvergreenName', 'EvergreenMatch') {
        if (-not $row.PSObject.Properties[$column]) {
            $row | Add-Member -NotePropertyName $column -NotePropertyValue ''
        }
    }
}

if (-not $NoIdLookup) {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Log 'winget not found, skipping automatic ID lookup' 'WARN'
    }
    else {
        $lookupRows = @($rows | Where-Object {
            $_.AppName -and -not $_.SourceId -and -not $_.WinGetId -and
            (-not $_.Source -or $_.Source -in @('WinGet', 'WinGetRepo')) -and
            -not (Get-SkipReason $_)
        })
        Write-Log "Looking up WinGet IDs for $($lookupRows.Count) application(s)"

        $found = 0
        foreach ($row in $lookupRows) {
            $match = $null
            try {
                $match = Find-WinGetId -Row $row
            }
            catch {
                Write-Log "ID lookup failed for $($row.AppName): $($_.Exception.Message)" 'WARN'
            }

            if ($match) {
                $row.WinGetId = $match.Id
                $row.IdMatch = $match.Detail
                $found++
                Write-Log "  $($row.AppName) -> $($match.Id)"
            }
            else {
                $row.IdMatch = 'No confident WinGet match'
            }
        }

        Write-Log "Matched $found of $($lookupRows.Count) to a WinGet id"
    }

    # Evergreen names, matched against the catalog Find-EvergreenApp returns in one call
    if ('Evergreen' -in $usedSources -and (Get-Command Find-EvergreenApp -ErrorAction SilentlyContinue)) {
        $lookupRows = @($rows | Where-Object {
            $_.AppName -and -not $_.EvergreenName -and
            (-not $_.Source -or ($_.Source -eq 'Evergreen' -and -not $_.SourceId)) -and
            -not (Get-SkipReason $_)
        })
        Write-Log "Looking up Evergreen names for $($lookupRows.Count) application(s)"

        try {
            $catalog = @(Find-EvergreenApp -WarningAction SilentlyContinue -InformationAction SilentlyContinue -ErrorAction Stop)
        }
        catch {
            Write-Log "Find-EvergreenApp failed, skipping Evergreen name lookup: $($_.Exception.Message)" 'WARN'
            $catalog = @()
        }

        $found = 0
        if ($catalog.Count) {
            foreach ($row in $lookupRows) {
                $match = Find-EvergreenName -Row $row -Catalog $catalog
                if ($match) {
                    $row.EvergreenName = $match.Name
                    $row.EvergreenMatch = $match.Detail
                    $found++
                    Write-Log "  $($row.AppName) -> $($match.Name)"
                }
                else {
                    $row.EvergreenMatch = 'No confident Evergreen match'
                }
            }
            Write-Log "Matched $found of $($lookupRows.Count) to an Evergreen name"
        }
    }

    $mappedPath = Join-Path $OutputDirectory "MappedImport_$stamp.csv"
    $rows | Export-Csv -Path $mappedPath -NoTypeInformation
    Write-Log "Mapped list written to $mappedPath"
}

Write-Log "Checking $($rows.Count) application(s) using $SearchMode mode"

$results = foreach ($row in $rows) {
    if (-not $row.AppName) { continue }

    $skipReason = Get-SkipReason $row
    if ($skipReason) {
        [pscustomobject]@{
            Application      = $row.AppName
            Publisher        = $row.Publisher
            InstalledVersion = $row.InstalledVersion
            LatestVersion    = $null
            Status           = 'Skipped'
            Source           = $null
            Candidates       = ''
            SourceId         = $row.SourceId
            WinGetId         = ''
            IdMatch          = ''
            EvergreenName    = ''
            EvergreenMatch   = ''
            DownloadUrl      = $null
            Error            = $skipReason
            CheckedUtc       = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        }
        continue
    }

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
        WinGetId         = $row.WinGetId
        IdMatch          = $row.IdMatch
        EvergreenName    = $row.EvergreenName
        EvergreenMatch   = $row.EvergreenMatch
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

$updateTable = $updates | Select-Object Application, Publisher, InstalledVersion, LatestVersion, Source, WinGetId, DownloadUrl
$reviewTable = $review  | Select-Object Application, Publisher, InstalledVersion, Status, WinGetId, EvergreenName, Error

$htmlBody = @()
$htmlBody += $updateTable | ConvertTo-Html -Fragment
$htmlBody += '<h3>Needs review (no source found or installed version unknown)</h3>'
$htmlBody += $reviewTable | ConvertTo-Html -Fragment

ConvertTo-Html -Title 'Third-party application update report' -PreContent $preContent -Body $htmlBody |
    Set-Content -Path $htmlPath -Encoding UTF8

$skipped = @($results | Where-Object { $_.Status -eq 'Skipped' })
Write-Log "Done. $($updates.Count) update(s) required, $($review.Count) item(s) need review, $($skipped.Count) skipped."
Write-Log "Full status:      $fullCsv"
Write-Log "Updates required: $updatesCsv"
Write-Log "HTML report:      $htmlPath"

#endregion