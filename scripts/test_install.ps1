param(
    [string]$BinaryPath = "target/release/rusdox.exe",
    [int]$Port = 18732
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $BinaryPath)) {
    throw "Missing RusDox binary: $BinaryPath"
}

$version = "v0.1.0"
$target = "x86_64-pc-windows-msvc"
$asset = "rusdox-$target.zip"
$testRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("rusdox-installer-test-" + [Guid]::NewGuid().ToString("N"))
$releaseDir = Join-Path $testRoot "releases\download\$version"
$archiveDir = Join-Path $testRoot "archive"
$installDir = Join-Path $testRoot "bin"
$server = $null

try {
    New-Item -ItemType Directory -Path $releaseDir, $archiveDir, $installDir -Force | Out-Null
    Copy-Item $BinaryPath (Join-Path $archiveDir "rusdox.exe")
    Compress-Archive -Path (Join-Path $archiveDir "rusdox.exe") -DestinationPath (Join-Path $releaseDir $asset)

    $hash = (Get-FileHash -Path (Join-Path $releaseDir $asset) -Algorithm SHA256).Hash.ToLowerInvariant()
    Set-Content -Path (Join-Path $releaseDir "SHA256SUMS") -Value "$hash  $asset" -NoNewline

    $server = Start-Process -FilePath "python" -ArgumentList @(
        "-m", "http.server", "$Port", "--bind", "127.0.0.1", "--directory", (Join-Path $testRoot "releases")
    ) -PassThru -WindowStyle Hidden

    $ready = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        try {
            Invoke-WebRequest -Uri "http://127.0.0.1:$Port/" -UseBasicParsing | Out-Null
            $ready = $true
            break
        } catch {
            Start-Sleep -Milliseconds 100
        }
    }
    if (-not $ready) {
        throw "Local release fixture did not start."
    }

    $env:APPDATA = Join-Path $testRoot "appdata"
    $env:LOCALAPPDATA = Join-Path $testRoot "localappdata"
    ./scripts/install.ps1 -Version $version -InstallDir $installDir -DownloadBase "http://127.0.0.1:$Port"

    $output = & (Join-Path $installDir "rusdox.exe") --version
    if ($output -notmatch "0\.1\.0") {
        throw "Installed binary reported unexpected version: $output"
    }
    Write-Host "Windows installer fixture passed for $target."
} finally {
    if ($server -and -not $server.HasExited) {
        Stop-Process -Id $server.Id -Force
    }
    if (Test-Path $testRoot) {
        Remove-Item -Path $testRoot -Recurse -Force
    }
}
