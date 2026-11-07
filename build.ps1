param(
    [string]$DistPath = (Join-Path $PSScriptRoot 'dist'),
    [string]$WorkPath = (Join-Path $PSScriptRoot 'build')
)

$ErrorActionPreference = 'Stop'
$preview = Join-Path $PSScriptRoot 'docs/excel-import.png'
$entry = Join-Path $PSScriptRoot 'app.py'
New-Item -ItemType Directory -Force -Path $WorkPath | Out-Null

python -B -m PyInstaller --noconfirm --onefile --windowed --name '题库速查' `
    --distpath $DistPath --workpath $WorkPath --specpath $WorkPath `
    --exclude-module numpy --exclude-module lxml --add-data "${preview};docs" $entry
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
