param(
    [string]$Repo = "OthmaneBlial/rusdox",
    [string]$Version = "latest",
    [string]$InstallDir = "$env:LOCALAPPDATA\Rusdox\bin",
    [string]$DownloadBase = ""
)

if ([string]::IsNullOrWhiteSpace($DownloadBase)) {
    $DownloadBase = "https://github.com/$Repo/releases"
}

$arch = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
switch ($arch) {
    "X64" { $target = "x86_64-pc-windows-msvc" }
    default { throw "Unsupported architecture: $arch (supported: X64)" }
}

$asset = "rusdox-$target.zip"
if ($Version -eq "latest") {
    $releaseBase = "$DownloadBase/latest/download"
} else {
    $releaseBase = "$DownloadBase/download/$Version"
}
$url = "$releaseBase/$asset"
$checksumsUrl = "$releaseBase/SHA256SUMS"

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("rusdox-install-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    $archivePath = Join-Path $tempDir $asset
    $checksumsPath = Join-Path $tempDir "SHA256SUMS"
    Write-Host "Downloading $url"
    Invoke-WebRequest -Uri $url -OutFile $archivePath
    Invoke-WebRequest -Uri $checksumsUrl -OutFile $checksumsPath

    $assetPattern = [regex]::Escape($asset)
    $checksumLine = Get-Content $checksumsPath | Where-Object {
        $_ -match "^([A-Fa-f0-9]{64})\s+\*?$assetPattern$"
    } | Select-Object -First 1
    if (-not $checksumLine) {
        throw "Checksum for $asset was not found in SHA256SUMS."
    }
    $expectedHash = ($checksumLine -split '\s+')[0].ToUpperInvariant()
    $actualHash = (Get-FileHash -Path $archivePath -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actualHash -ne $expectedHash) {
        throw "Checksum verification failed for $asset."
    }
    Write-Host "Verified SHA-256 for $asset"

    Expand-Archive -Path $archivePath -DestinationPath $tempDir -Force

    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    $rusdoxExe = Join-Path $InstallDir "rusdox.exe"
    Copy-Item (Join-Path $tempDir "rusdox.exe") $rusdoxExe -Force

    $configPath = & $rusdoxExe config path
    $configCreated = $false
    if (-not (Test-Path $configPath)) {
        & $rusdoxExe config init --template | Out-Null
        $configCreated = $true
    }

    $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
    if ([string]::IsNullOrWhiteSpace($userPath)) {
        [Environment]::SetEnvironmentVariable("Path", $InstallDir, "User")
        Write-Host "Added $InstallDir to your User PATH."
    } else {
        $parts = $userPath -split ';'
        if (-not ($parts -contains $InstallDir)) {
            [Environment]::SetEnvironmentVariable("Path", "$userPath;$InstallDir", "User")
            Write-Host "Added $InstallDir to your User PATH."
        }
    }

    Write-Host "Installed rusdox.exe to $rusdoxExe"
    Write-Host "User config: $configPath"
    if ($configCreated) {
        Write-Host "Created default config at $configPath"
    }
    Write-Host "Customize styling with:"
    Write-Host "  rusdox config wizard --level basic"
    Write-Host "  rusdox config wizard --level advanced"
    Write-Host "Create a project-local override with:"
    Write-Host "  rusdox config wizard --path .\rusdox.toml --level basic"
} finally {
    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force
    }
}
