param(
    [string]$Repo = "OthmaneBlial/rusdox",
    [string]$Version = "latest",
    [string]$InstallDir = "$env:LOCALAPPDATA\Rusdox\bin"
)

$arch = [System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
switch ($arch) {
    "X64" { $target = "x86_64-pc-windows-msvc" }
    default { throw "Unsupported architecture: $arch (supported: X64)" }
}

$asset = "rusdox-$target.zip"
if ($Version -eq "latest") {
    $url = "https://github.com/$Repo/releases/latest/download/$asset"
} else {
    $url = "https://github.com/$Repo/releases/download/$Version/$asset"
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("rusdox-install-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    $archivePath = Join-Path $tempDir $asset
    Write-Host "Downloading $url"
    Invoke-WebRequest -Uri $url -OutFile $archivePath

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
