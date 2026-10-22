<#
.SYNOPSIS
    Downloads the installers for applications listed in an UpdatesRequired CSV.

.DESCRIPTION
    Reads the UpdatesRequired_<timestamp>.csv written by Get-AppVersions1.ps1 and downloads the
    installer for each row into <OutputDirectory>\<AppName>_<LatestVersion>.

    How each row is downloaded, based on its DownloadUrl:
      github.com/microsoft/winget-pkgs/tree/...   Reads <Id>.installer.yaml from that manifest folder
      winget:winget/<Id>                          Reads the community manifest for <Id> at LatestVersion
      winget:msstore/<Id>                         winget download from the msstore source
      http(s)://...                               Downloads the URL directly (Evergreen, GitHub, Scoop)

    From the installer manifest the script picks one installer, preferring -Architecture, then
    neutral, then x86; -Scope machine; en-US or no locale; and msi/wix over msix over exe types.
    The file is checked against InstallerSha256 and removed if the hash does not match.

    When no installer URL can be found or the download fails, the script falls back to
    winget download from the winget source, then the msstore source.

    Manifests are read from raw.githubusercontent.com, not the GitHub API, so the API rate limit
    does not apply.

    Writes DownloadLog_<timestamp>.csv to -OutputDirectory with the result for each row.

.PARAMETER CsvPath
    UpdatesRequired CSV. Defaults to the newest UpdatesRequired_*.csv in .\UpdateReports.

.PARAMETER OutputDirectory
    Folder that receives one subfolder per application. Defaults to .\Downloads

.PARAMETER Architecture
    Preferred installer architecture. Defaults to x64.

.PARAMETER Scope
    Preferred installer scope. Defaults to machine.

.PARAMETER Force
    Download again even when the application folder already contains files.

.PARAMETER SkipHashCheck
    Keep a manifest download whose SHA256 does not match InstallerSha256 instead of deleting it.
    The mismatch is still reported in the log and the Error column, and HashChecked is False.
    Also passes --ignore-security-hash to the winget download fallback, which only works when
    an administrator has run "winget settings --enable InstallerHashOverride" and the script
    is not running elevated. Use this only for vendors that replace files at a fixed URL.

.EXAMPLE
    .\Save-AppInstallers.ps1

.EXAMPLE
    .\Save-AppInstallers.ps1 -CsvPath .\UpdateReports\UpdatesRequired_20261022_1313.csv -Architecture x86

.NOTES
    Requires PowerShell 5.1 or later. The winget fallback needs winget 1.8 or later (winget download).
    msstore downloads require signing in with a Microsoft Entra ID account and may fail unattended.
#>
[CmdletBinding()]
param(
    [string]$CsvPath,
    [string]$OutputDirectory = (Join-Path (Get-Location) 'Downloads'),
    [ValidateSet('x64', 'x86', 'arm64')]
    [string]$Architecture = 'x64',
    [ValidateSet('machine', 'user')]
    [string]$Scope = 'machine',
    [switch]$Force,
    [switch]$SkipHashCheck
)

$ProgressPreference = 'SilentlyContinue'   # The progress bar slows Invoke-WebRequest downloads considerably
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$script:RawManifestRoot = 'https://raw.githubusercontent.com/microsoft/winget-pkgs/master'

# Same user agent as winget. Some hosts (e.g. SourceForge) return an HTML page to browser user agents.
$script:UserAgent = 'winget-cli WindowsPackageManager/1.29 DesktopAppInstaller/Microsoft.DesktopAppInstaller'

# File extension for each winget InstallerType, used when the URL does not end in one
$script:TypeExtensions = @{
    'msi' = '.msi'; 'wix' = '.msi'; 'msix' = '.msix'; 'appx' = '.appx'
    'exe' = '.exe'; 'inno' = '.exe'; 'nullsoft' = '.exe'; 'burn' = '.exe'; 'portable' = '.exe'
    'zip' = '.zip'
}

#region Helpers

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message)
}

function Get-FolderName {
    # "WinSCP 6.3.4" at 6.7 becomes "WinSCP_6.7". The installed version is removed from the name
    # so the folder only carries the version being downloaded.
    param([string]$Application, [string]$Version)

    $name = $Application -replace '(?i)\bversion\s+\S+', ''
    $name = $name -replace '\bv?\d+(\.\d+)+\S*', ''
    $name = $name -replace '\(\s*\)', ''
    $name = ($name -replace '\s+', ' ').Trim(' -_')
    if (-not $name) { $name = $Application }

    $folder = "${name}_$Version"
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) { $folder = $folder.Replace([string]$char, '_') }
    return $folder.Trim()
}

function Get-ManifestFolder {
    # Manifest folder for a WinGet id and version, e.g. manifests/a/AnalogDevices/LTspice/26.1.1.1
    param([string]$Id, [string]$Version)
    $first = $Id.Substring(0, 1).ToLower()
    return "manifests/$first/" + ($Id -replace '\.', '/') + "/$Version"
}

#endregion

#region Manifest parsing

function ConvertFrom-YamlValue {
    param([string]$Value)
    $Value = $Value.Trim()
    if ($Value -match '^"(.*)"$' -or $Value -match "^'(.*)'$") { return $Matches[1] }
    return $Value
}

function Read-InstallerManifest {
    # Minimal reader for winget installer manifests. Returns the top-level values and one
    # hashtable per entry under Installers:. Only the scalar keys directly on each installer are
    # kept; nested blocks such as InstallerSwitches are not needed here and are skipped.
    param([string]$Yaml)

    $root = @{}
    $installers = New-Object System.Collections.ArrayList
    $inInstallers = $false
    $itemIndent = -1
    $current = $null

    foreach ($line in ($Yaml -split "`r?`n")) {
        if ($line -match '^\s*(#|$)') { continue }
        $indent = $line.Length - $line.TrimStart().Length

        # Installer items can start at column 0 ("- Architecture: x64"), so check them first
        if ($inInstallers -and $line -match '^(\s*)-\s+(\w+):\s*(.*)$' -and ($itemIndent -lt 0 -or $Matches[1].Length -eq $itemIndent)) {
            $itemIndent = $Matches[1].Length
            $current = @{}
            [void]$installers.Add($current)
            if ($Matches[3]) { $current[$Matches[2]] = ConvertFrom-YamlValue $Matches[3] }
            continue
        }

        if ($indent -eq 0) {
            $inInstallers = $false
            if ($line -match '^Installers:\s*$') { $inInstallers = $true; $itemIndent = -1; continue }
            if ($line -match '^(\w+):\s*(.*)$' -and $Matches[2]) { $root[$Matches[1]] = ConvertFrom-YamlValue $Matches[2] }
            continue
        }
        if (-not $inInstallers) { continue }

        if ($current -and $indent -eq ($itemIndent + 2) -and $line -match '^\s*(\w+):\s*(.*)$' -and $Matches[2]) {
            $current[$Matches[1]] = ConvertFrom-YamlValue $Matches[2]
        }
    }

    [pscustomobject]@{ Root = $root; Installers = @($installers) }
}

function Select-Installer {
    # Picks the installer that best fits -Architecture and -Scope. Values missing on an
    # installer fall back to the manifest's top-level values.
    param($Manifest, [string]$Arch, [string]$PreferredScope)

    $typeRank = @{ 'msi' = 0; 'wix' = 0; 'msix' = 1; 'appx' = 1; 'burn' = 2; 'exe' = 2; 'inno' = 2; 'nullsoft' = 2; 'zip' = 3; 'portable' = 4 }
    $archRank = @{ 'x86' = 2; 'neutral' = 1 }
    $archRank[$Arch] = 0

    $options = foreach ($installer in $Manifest.Installers) {
        $value = { param($key) if ($installer[$key]) { $installer[$key] } else { $Manifest.Root[$key] } }

        $url = & $value 'InstallerUrl'
        $installerArch = & $value 'Architecture'
        if (-not $url -or -not $archRank.ContainsKey($installerArch)) { continue }

        $type = (& $value 'InstallerType')
        $installerScope = & $value 'Scope'
        $locale = & $value 'InstallerLocale'

        [pscustomobject]@{
            Url       = $url
            Sha256    = & $value 'InstallerSha256'
            Type      = $type
            ArchRank  = $archRank[$installerArch]
            ScopeRank = if ($installerScope -eq $PreferredScope) { 0 } elseif (-not $installerScope) { 1 } else { 2 }
            LocaleRank = if (-not $locale -or $locale -eq 'en-US') { 0 } elseif ($locale -like 'en*') { 1 } else { 2 }
            TypeRank  = if ($type -and $typeRank.ContainsKey($type)) { $typeRank[$type] } else { 5 }
        }
    }

    $options |
        Sort-Object -Property ArchRank, ScopeRank, LocaleRank, TypeRank |
        Select-Object -First 1
}

function Get-InstallerFromManifest {
    # Returns the selected installer for a manifest folder, or $null when no installer fits.
    # Throws when the manifest does not exist.
    param([string]$Folder, [string]$Id)

    # Newer manifests use <Id>.installer.yaml; older single-file manifests use <Id>.yaml
    $found = $false
    foreach ($fileName in "$Id.installer.yaml", "$Id.yaml") {
        $url = "$script:RawManifestRoot/$Folder/$fileName"
        try {
            $yaml = (Invoke-WebRequest -Uri $url -UseBasicParsing -UserAgent $script:UserAgent).Content
        }
        catch {
            continue
        }
        $found = $true

        $manifest = Read-InstallerManifest -Yaml $yaml
        $installer = Select-Installer -Manifest $manifest -Arch $Architecture -PreferredScope $Scope
        if ($installer) {
            $installer | Add-Member -NotePropertyName ManifestUrl -NotePropertyValue $url
            return $installer
        }
    }
    if (-not $found) { throw "Manifest not found: $script:RawManifestRoot/$Folder" }
    return $null
}

#endregion

#region Downloads

function Get-InstallerFileName {
    # Uses the last URL segment with an installer extension, e.g. WinSCP-6.7.msi from
    # .../WinSCP-6.7.msi/download. Otherwise builds a name from the id, version and type.
    param([string]$Url, [string]$Type, [string]$FallbackName)

    $path = ([uri]$Url).AbsolutePath
    $segments = @($path -split '/' | Where-Object { $_ } | ForEach-Object { [uri]::UnescapeDataString($_) })
    [array]::Reverse($segments)
    foreach ($segment in $segments) {
        if ($segment -match '\.(msi|msix|msixbundle|appx|appxbundle|exe|zip|7z)$') { return $segment }
    }

    $extension = if ($Type -and $script:TypeExtensions.ContainsKey($Type)) { $script:TypeExtensions[$Type] } else { '.exe' }
    return "$FallbackName$extension"
}

function Save-FromUrl {
    # Downloads a file and checks it. Throws when the server returns a web page instead of a
    # file or when the SHA256 does not match the manifest. With -SkipHashCheck a mismatched file
    # is kept and the mismatch is recorded in $script:HashWarning.
    param([string]$Url, [string]$Destination, [string]$Sha256)

    $script:HashWarning = ''

    $response = Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -UserAgent $script:UserAgent -PassThru

    if ($response.Headers['Content-Type'] -match 'text/html') {
        Remove-Item -Path $Destination -Force
        throw "Server returned a web page instead of a file: $Url"
    }

    if ($Sha256) {
        $actual = (Get-FileHash -Path $Destination -Algorithm SHA256).Hash
        if ($actual -ne $Sha256) {
            $message = "SHA256 mismatch for $Url (expected $Sha256, got $actual)"
            if ($SkipHashCheck) {
                $script:HashWarning = "$message, kept because of -SkipHashCheck"
                Write-Log "  $script:HashWarning" 'WARN'
                return $false
            }
            Remove-Item -Path $Destination -Force
            throw $message
        }
        return $true
    }
    return $false
}

function Save-WithWinGet {
    # winget download writes the installer and its manifest into the folder
    param([string]$Id, [string]$Version, [string]$Source, [string]$Folder)

    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw 'winget not found. Install App Installer from the Microsoft Store.'
    }

    $arguments = @('download', '--id', $Id, '--exact', '--source', $Source, '--download-directory', $Folder,
        '--accept-source-agreements', '--accept-package-agreements', '--disable-interactivity')
    if ($Source -eq 'winget') {
        $arguments += @('--architecture', $Architecture, '--scope', $Scope)
        if ($Version) { $arguments += @('--version', $Version) }
        if ($SkipHashCheck) { $arguments += '--ignore-security-hash' }
    }
    else {
        $arguments += '--skip-license'
    }

    $output = & winget @arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = ($output | ForEach-Object { [string]$_ } | Where-Object { $_ -match '\w' } | Select-Object -Last 1)
        throw "winget download ($Source) failed with exit code $LASTEXITCODE. $message"
    }

    $file = Get-ChildItem -Path $Folder -File |
        Where-Object { $_.Extension -notin '.yaml', '.yml' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $file) { throw "winget download ($Source) finished but no installer was saved" }
    return $file.FullName
}

#endregion

#region Main

if (-not $CsvPath) {
    $latest = Get-ChildItem -Path (Join-Path (Get-Location) 'UpdateReports') -Filter 'UpdatesRequired_*.csv' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $latest) { throw 'No -CsvPath given and no UpdatesRequired_*.csv found in .\UpdateReports' }
    $CsvPath = $latest.FullName
}
if (-not (Test-Path $CsvPath)) { throw "CSV not found: $CsvPath" }

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'

$rows = @(Import-Csv -Path $CsvPath | Where-Object { $_.Application -and $_.LatestVersion })
Write-Log "Downloading installers for $($rows.Count) application(s) from $CsvPath"

$results = foreach ($row in $rows) {
    $folderName = Get-FolderName -Application $row.Application -Version $row.LatestVersion
    $folder = Join-Path $OutputDirectory $folderName
    Write-Log "$($row.Application) $($row.LatestVersion) -> $folderName"

    $result = [pscustomobject]@{
        Application = $row.Application
        Version     = $row.LatestVersion
        Folder      = $folder
        Method      = ''
        File        = ''
        HashChecked = $false
        Status      = ''
        Error       = ''
    }

    if (-not $Force -and (Test-Path $folder) -and (Get-ChildItem -Path $folder -File | Where-Object { $_.Extension -notin '.yaml', '.yml' })) {
        $result.Status = 'AlreadyDownloaded'
        Write-Log '  Folder already has an installer, skipping (use -Force to download again)'
        $result
        continue
    }
    New-Item -ItemType Directory -Force -Path $folder | Out-Null

    # Work out the WinGet id and manifest folder, if the row has one
    $url = [string]$row.DownloadUrl
    $id = if ($row.WinGetId) { $row.WinGetId } elseif ($row.SourceId -match '^[^/\s]+\.[^/\s]+$') { $row.SourceId } else { '' }
    $manifestFolder = ''
    $directUrl = ''
    $storeOnly = $false

    if ($url -match '^https://github\.com/microsoft/winget-pkgs/(tree|blob)/master/(manifests/./.+?)/?$') {
        $manifestFolder = $Matches[2] -replace '/[^/]+\.yaml$', ''
        # Id is every folder between manifests/<letter>/ and the version folder
        $parts = @($manifestFolder -split '/')
        if (-not $id) { $id = $parts[2..($parts.Count - 2)] -join '.' }
    }
    elseif ($url -match '^winget:(winget|msstore)/(.+)$') {
        $id = $Matches[2]
        if ($Matches[1] -eq 'msstore') { $storeOnly = $true }
        else { $manifestFolder = Get-ManifestFolder -Id $id -Version $row.LatestVersion }
    }
    elseif ($url -match '^https?://') {
        $directUrl = $url -replace '#.*$', ''   # Scoop urls can carry a #/rename suffix
    }
    elseif ($id) {
        $manifestFolder = Get-ManifestFolder -Id $id -Version $row.LatestVersion
    }

    $errors = @()

    # 1. Installer URL from the community manifest
    if ($manifestFolder) {
        try {
            $installer = Get-InstallerFromManifest -Folder $manifestFolder -Id $id
            if (-not $installer) { throw "No $Architecture installer found in manifest folder $manifestFolder" }

            $fileName = Get-InstallerFileName -Url $installer.Url -Type $installer.Type -FallbackName "$($id)_$($row.LatestVersion)"
            $destination = Join-Path $folder $fileName
            Write-Log "  Manifest installer: $($installer.Url)"
            $result.HashChecked = Save-FromUrl -Url $installer.Url -Destination $destination -Sha256 $installer.Sha256
            $result.Method = 'WinGetManifest'
            $result.File = $destination
            if ($script:HashWarning) { $errors += $script:HashWarning }
        }
        catch {
            $errors += "Manifest: $($_.Exception.Message)"
        }
    }

    # 2. Direct download URL from Evergreen, Nevergreen, GitHub or Scoop
    if (-not $result.File -and $directUrl) {
        try {
            $fileName = Get-InstallerFileName -Url $directUrl -Type '' -FallbackName $folderName
            $destination = Join-Path $folder $fileName
            Write-Log "  Direct download: $directUrl"
            [void](Save-FromUrl -Url $directUrl -Destination $destination -Sha256 '')
            $result.Method = 'DirectUrl'
            $result.File = $destination
        }
        catch {
            $errors += "Direct: $($_.Exception.Message)"
        }
    }

    # 3. winget download, then msstore
    if (-not $result.File -and $id) {
        $wingetSources = if ($storeOnly) { @('msstore') } else { @('winget', 'msstore') }
        foreach ($source in $wingetSources) {
            try {
                Write-Log "  Falling back to winget download from $source"
                $result.File = Save-WithWinGet -Id $id -Version $row.LatestVersion -Source $source -Folder $folder
                $result.Method = "winget:$source"
                break
            }
            catch {
                $errors += "${source}: $($_.Exception.Message)"
            }
        }
    }

    if ($result.File) {
        $result.Status = 'Downloaded'
        Write-Log "  Saved $($result.File)"
    }
    else {
        if (-not $id -and -not $directUrl) { $errors += 'No download URL or WinGet id for this row' }
        $result.Status = 'Failed'
        Write-Log "  Failed: $($errors -join '; ')" 'WARN'
        # Leave no empty folder behind
        if (-not (Get-ChildItem -Path $folder -Force)) { Remove-Item -Path $folder -Force }
    }
    $result.Error = $errors -join '; '
    $result
}

$results = @($results)
$logPath = Join-Path $OutputDirectory "DownloadLog_$stamp.csv"
$results | Export-Csv -Path $logPath -NoTypeInformation

$summary = ($results | Group-Object Status | ForEach-Object { "$($_.Name): $($_.Count)" }) -join ' | '
Write-Log "Done. $summary"
Write-Log "Download log: $logPath"

#endregion
