param(
  [switch]$SkipBrowserInstall
)

$ErrorActionPreference = "Stop"

function Get-NpmCommand {
  $cmd = Get-Command npm.cmd -ErrorAction SilentlyContinue
  if (-not $cmd) {
    $cmd = Get-Command npm -ErrorAction SilentlyContinue
  }
  if (-not $cmd) { return $null }
  return $cmd.Source
}

function Ensure-NodeAndNpm {
  $npm = Get-NpmCommand
  if ($npm) { return $npm }

  Write-Host "Node.js/NPM nao encontrado."
  Write-Host "Tentando instalar automaticamente via winget..."

  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
  if (-not $winget) {
    throw "winget nao encontrado. Instale Node.js LTS manualmente em https://nodejs.org e rode novamente."
  }

  & $winget.Source install --id OpenJS.NodeJS.LTS --source winget --accept-package-agreements --accept-source-agreements --silent --scope user
  if ($LASTEXITCODE -ne 0) {
    throw "Falha ao instalar Node.js via winget. Codigo: $LASTEXITCODE"
  }

  $extraNodePaths = @(
    "$env:LocalAppData\Programs\nodejs",
    "$env:ProgramFiles\nodejs"
  )
  foreach ($p in $extraNodePaths) {
    if ((Test-Path $p) -and ($env:Path -notlike "*$p*")) {
      $env:Path = "$p;$env:Path"
    }
  }

  $npm = Get-NpmCommand
  if (-not $npm) {
    throw "Node.js foi instalado, mas npm ainda nao foi encontrado nesta sessao. Feche e abra o instalador novamente."
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
$npm = Ensure-NodeAndNpm

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
