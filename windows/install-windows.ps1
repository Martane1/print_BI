param(
  [switch]$SkipBrowserInstall
)

$ErrorActionPreference = "Stop"
try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch {
}

function Get-LocalNpmCommand {
  param([string]$ProjectDir)
  $localNpm = Join-Path $ProjectDir ".tools\node\npm.cmd"
  if (Test-Path $localNpm) {
    return $localNpm
  }
  return $null
}

function Get-GlobalNpmCommand {
  $cmd = Get-Command npm.cmd -ErrorAction SilentlyContinue
  if (-not $cmd) {
    $cmd = Get-Command npm -ErrorAction SilentlyContinue
  }
  if (-not $cmd) {
    return $null
  }
  return $cmd.Source
}

function Refresh-ProcessPath {
  $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
  $userPath = [Environment]::GetEnvironmentVariable("Path", "User")
  $merged = @()
  if ($machinePath) { $merged += $machinePath }
  if ($userPath) { $merged += $userPath }
  if ($merged.Count -gt 0) {
    $env:Path = ($merged -join ";")
  }
}

function Install-NodeViaWinget {
  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  if (-not $winget) {
    return $false
  }

  Write-Host "Tentando instalar Node.js LTS via winget..."
  & $winget.Source install --id OpenJS.NodeJS.LTS --source winget --accept-package-agreements --accept-source-agreements --silent --scope user
  if ($LASTEXITCODE -ne 0) {
    Write-Host "Falha no winget (codigo $LASTEXITCODE)."
    return $false
  }

  return $true
}

function Install-NodePortable {
  param([string]$ProjectDir)

  $toolsDir = Join-Path $ProjectDir ".tools"
  $nodeRoot = Join-Path $toolsDir "node"
  $extractDir = Join-Path $toolsDir "node-extract"

  if (-not (Test-Path $toolsDir)) {
    New-Item -ItemType Directory -Path $toolsDir | Out-Null
  }

  if (Test-Path $extractDir) {
    Remove-Item -Path $extractDir -Recurse -Force
  }
  New-Item -ItemType Directory -Path $extractDir | Out-Null

  $offlineZip = Join-Path $ProjectDir "windows\assets\node-win-x64.zip"
  $zipPath = $null

  if (Test-Path $offlineZip) {
    Write-Host "Usando pacote Node.js offline: windows/assets/node-win-x64.zip"
    $zipPath = $offlineZip
  } else {
    $baseUrl = "https://nodejs.org/dist/latest-v20.x/"
    $shaUrl = "$baseUrl" + "SHASUMS256.txt"

    Write-Host "Baixando metadados do Node.js..."
    $shaContent = (Invoke-WebRequest -Uri $shaUrl -UseBasicParsing).Content

    $zipName = $null
    foreach ($line in ($shaContent -split "`n")) {
      if ($line -match "node-v20\.[0-9]+\.[0-9]+-win-x64\.zip") {
        $parts = ($line -split "\s+") | Where-Object { $_ -and $_.Trim() -ne "" }
        $zipName = $parts[-1].Trim()
        break
      }
    }

    if (-not $zipName) {
      throw "Nao foi possivel identificar o pacote win-x64 do Node.js."
    }

    $zipUrl = "$baseUrl$zipName"
    $zipPath = Join-Path $env:TEMP $zipName

    Write-Host "Baixando Node.js portatil: $zipName"
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
  }

  if (Test-Path $nodeRoot) {
    Remove-Item -Path $nodeRoot -Recurse -Force
  }

  Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
  $extracted = Get-ChildItem -Path $extractDir -Directory | Select-Object -First 1
  if (-not $extracted) {
    throw "Falha ao extrair Node.js portatil."
  }

  Move-Item -Path $extracted.FullName -Destination $nodeRoot
  Remove-Item -Path $extractDir -Recurse -Force
}

function Get-NpmCommand {
  param([string]$ProjectDir)

  $localNpm = Get-LocalNpmCommand -ProjectDir $ProjectDir
  if ($localNpm) {
    return $localNpm
  }

  return (Get-GlobalNpmCommand)
}

function Ensure-NodeAndNpm {
  param([string]$ProjectDir)

  $npm = Get-NpmCommand -ProjectDir $ProjectDir
  if ($npm) {
    return $npm
  }

  Write-Host "Node.js/NPM nao encontrado."
  $wingetInstalled = Install-NodeViaWinget
  if ($wingetInstalled) {
    Refresh-ProcessPath
    $npm = Get-NpmCommand -ProjectDir $ProjectDir
    if ($npm) {
      return $npm
    }
  }

  Write-Host "Tentando fallback com Node.js portatil no projeto (.tools/node)..."
  Install-NodePortable -ProjectDir $ProjectDir
  $npm = Get-NpmCommand -ProjectDir $ProjectDir
  if (-not $npm) {
    throw "Node.js nao foi encontrado nem instalado (winget/portatil). Verifique acesso a https://nodejs.org."
  }

  return $npm
}

function Ensure-PlaywrightChromium {
  param([string]$ProjectDir, [string]$NpmCmd)

  $playwrightCli = Join-Path $ProjectDir "node_modules\.bin\playwright.cmd"
  if (Test-Path $playwrightCli) {
    Write-Host "Instalando Chromium do Playwright..."
    & $playwrightCli install chromium
  } else {
    Write-Host "Playwright CLI nao encontrado em node_modules/.bin. Tentando via npm exec..."
    & $NpmCmd exec -- playwright install chromium
  }

  if ($LASTEXITCODE -ne 0) {
    throw "Falha ao instalar Chromium do Playwright. Codigo: $LASTEXITCODE"
  }
}

function Create-Shortcut {
  param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$TargetPath,
    [Parameter(Mandatory = $true)][string]$Arguments,
    [Parameter(Mandatory = $true)][string]$WorkingDirectory
  )

  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut($Path)
  $shortcut.TargetPath = $TargetPath
  $shortcut.Arguments = $Arguments
  $shortcut.WorkingDirectory = $WorkingDirectory
  $shortcut.IconLocation = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe,0"
  $shortcut.Save()
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
Set-Location $projectDir

Write-Host "Projeto: $projectDir"
Write-Host "Validando Node.js/NPM..."
$npm = Ensure-NodeAndNpm -ProjectDir $projectDir

if (-not (Test-Path (Join-Path $projectDir "config.json")) -and (Test-Path (Join-Path $projectDir "config.example.json"))) {
  Copy-Item (Join-Path $projectDir "config.example.json") (Join-Path $projectDir "config.json")
  Write-Host "config.json criado a partir de config.example.json"
}

Write-Host "Instalando dependencias Node..."
& $npm install

if (-not $SkipBrowserInstall) {
  Ensure-PlaywrightChromium -ProjectDir $projectDir -NpmCmd $npm
}

$setupDir = Join-Path $projectDir ".setup"
if (-not (Test-Path $setupDir)) {
  New-Item -ItemType Directory -Path $setupDir | Out-Null
}
Set-Content -Path (Join-Path $setupDir "windows-ready.txt") -Value (Get-Date -Format o)

$desktop = [Environment]::GetFolderPath("Desktop")
$runner = Join-Path $projectDir "windows\run-print-bi.ps1"
$pwsh = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

Create-Shortcut -Path (Join-Path $desktop "PRINT BI - Launcher.lnk") `
  -TargetPath $pwsh `
  -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$runner`" -Action Menu" `
  -WorkingDirectory $projectDir

Create-Shortcut -Path (Join-Path $desktop "PRINT BI - Login.lnk") `
  -TargetPath $pwsh `
  -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$runner`" -Action Login" `
  -WorkingDirectory $projectDir

Create-Shortcut -Path (Join-Path $desktop "PRINT BI - Captura.lnk") `
  -TargetPath $pwsh `
  -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$runner`" -Action Capture" `
  -WorkingDirectory $projectDir

Create-Shortcut -Path (Join-Path $desktop "PRINT BI - Ultima Saida.lnk") `
  -TargetPath $pwsh `
  -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$runner`" -Action OpenLast" `
  -WorkingDirectory $projectDir

Write-Host ""
Write-Host "Instalacao finalizada."
Write-Host "Atalhos criados na Area de Trabalho:"
Write-Host "- PRINT BI - Launcher"
Write-Host "- PRINT BI - Login"
Write-Host "- PRINT BI - Captura"
Write-Host "- PRINT BI - Ultima Saida"
