param(
  [ValidateSet("Menu", "Login", "Capture", "OpenLast")]
  [string]$Action = "Menu"
)

$ErrorActionPreference = "Stop"

function Get-NpmCommand {
  param([string]$ProjectDir)

  $localNpm = Join-Path $ProjectDir ".tools\node\npm.cmd"
  if (Test-Path $localNpm) {
    return $localNpm
  }

  $cmd = Get-Command npm.cmd -ErrorAction SilentlyContinue
  if (-not $cmd) {
    $cmd = Get-Command npm -ErrorAction SilentlyContinue
  }
  if (-not $cmd) {
    throw "npm nao encontrado. Execute INSTALAR_WINDOWS.bat novamente."
  }
  return $cmd.Source
}

function Ensure-Dependencies {
  param([string]$ProjectDir, [string]$NpmCmd)
  $nodeModules = Join-Path $ProjectDir "node_modules"
  if (-not (Test-Path $nodeModules)) {
    Write-Host "node_modules nao encontrado. Instalando dependencias..."
    & $NpmCmd install
  }
}

function Select-ActionFromUi {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "PRINT BI"
  $form.Width = 420
  $form.Height = 250
  $form.StartPosition = "CenterScreen"
  $form.FormBorderStyle = "FixedDialog"
  $form.MaximizeBox = $false
  $form.MinimizeBox = $false
  $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.AutoSize = $true
  $label.Text = "Selecione uma acao:"
  $label.Location = New-Object System.Drawing.Point(25, 20)
  $form.Controls.Add($label)

  $selected = $null

  $btnLogin = New-Object System.Windows.Forms.Button
  $btnLogin.Text = "Login"
  $btnLogin.Width = 160
  $btnLogin.Height = 38
  $btnLogin.Location = New-Object System.Drawing.Point(25, 55)
  $btnLogin.Add_Click({ $script:selected = "Login"; $form.Close() })
  $form.Controls.Add($btnLogin)

  $btnCapture = New-Object System.Windows.Forms.Button
  $btnCapture.Text = "Captura completa"
  $btnCapture.Width = 160
  $btnCapture.Height = 38
  $btnCapture.Location = New-Object System.Drawing.Point(205, 55)
  $btnCapture.Add_Click({ $script:selected = "Capture"; $form.Close() })
  $form.Controls.Add($btnCapture)

  $btnLast = New-Object System.Windows.Forms.Button
  $btnLast.Text = "Abrir ultima saida"
  $btnLast.Width = 160
  $btnLast.Height = 38
  $btnLast.Location = New-Object System.Drawing.Point(25, 105)
  $btnLast.Add_Click({ $script:selected = "OpenLast"; $form.Close() })
  $form.Controls.Add($btnLast)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = "Cancelar"
  $btnCancel.Width = 160
  $btnCancel.Height = 38
  $btnCancel.Location = New-Object System.Drawing.Point(205, 105)
  $btnCancel.Add_Click({ $script:selected = $null; $form.Close() })
  $form.Controls.Add($btnCancel)

  [void]$form.ShowDialog()
  return $selected
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
Set-Location $projectDir

if ($Action -eq "Menu") {
  $Action = Select-ActionFromUi
  if (-not $Action) {
    exit 0
  }
}

if ($Action -eq "OpenLast") {
  $outputDir = Join-Path $projectDir "output"
  if (-not (Test-Path $outputDir)) {
    throw "Pasta output nao encontrada."
  }

  $latest = Get-ChildItem -Path $outputDir -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $latest) {
    throw "Nenhuma saida encontrada em output."
  }

  Start-Process explorer.exe $latest.FullName
  exit 0
}

$npm = Get-NpmCommand -ProjectDir $projectDir
Ensure-Dependencies -ProjectDir $projectDir -NpmCmd $npm

if ($Action -eq "Login") {
  Write-Host "Abrindo fluxo de login..."
  Write-Host "Apos autenticar no navegador, volte aqui e pressione ENTER."
  & $npm run capture -- --config config.json --login
  exit $LASTEXITCODE
}

if ($Action -eq "Capture") {
  Write-Host "Iniciando captura completa..."
  & $npm run capture -- --config config.json
  exit $LASTEXITCODE
}

throw "Acao invalida: $Action"
