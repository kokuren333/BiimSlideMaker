param([string]$TestDirectory)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $TestDirectory) { $TestDirectory = Join-Path $root ('output/frame-test-' + [guid]::NewGuid().ToString('N')) }
$TestDirectory = [IO.Path]::GetFullPath($TestDirectory)
New-Item -ItemType Directory -Path $TestDirectory -Force | Out-Null
$cli = Join-Path $root 'biim-video.ps1'
$source = Join-Path $root 'assets/frame-default.svg'
$expectedBoxes = @{
    slide = @(16,16,1440,810); notes_top = @(1498,34,388,182)
    notes_bottom = @(1498,274,388,532); subtitle = @(350,870,1528,178)
    character = @(0,740,330,332)
}
function Test-Init([string]$Name, [string]$FramePath, [string]$Extension) {
    $projectDir = Join-Path $TestDirectory $Name
    $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',$cli,'init',$projectDir)
    if ($FramePath) { $arguments += @('-FrameImage',$FramePath) }
    & powershell.exe @arguments
    if ($LASTEXITCODE -ne 0) { throw "init failed: $Name" }
    $project = Get-Content (Join-Path $projectDir 'project.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($project.assets.background -ne "assets/frame.$Extension") { throw "Incorrect background: $Name" }
    $original = if ($FramePath) { $FramePath } else { $source }
    if ((Get-FileHash $original).Hash -ne (Get-FileHash (Join-Path $projectDir $project.assets.background)).Hash) { throw "Frame copy changed: $Name" }
    foreach ($key in $expectedBoxes.Keys) {
        if (($project.layout.$key -join ',') -ne ($expectedBoxes[$key] -join ',')) { throw "Incorrect default box: $key" }
    }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $cli validate $projectDir
    if ($LASTEXITCODE -ne 0) { throw "validate failed: $Name" }
    Write-Output "PASS $Name"
}
Test-Init 'default-svg' '' 'svg'
Test-Init 'explicit-svg' $source 'svg'
Add-Type -AssemblyName System.Drawing
$pngPath = Join-Path $TestDirectory 'custom.png'
$bitmap = [Drawing.Bitmap]::new(1920,1080)
try { $bitmap.Save($pngPath, [Drawing.Imaging.ImageFormat]::Png) } finally { $bitmap.Dispose() }
Test-Init 'explicit-png' $pngPath 'png'
$defaultDir = Join-Path $TestDirectory 'default-svg'
$log = Join-Path $TestDirectory 'preview.log'
$argumentLine = (@('-NoProfile','-ExecutionPolicy','Bypass','-File',$cli,'preview',$defaultDir) | ForEach-Object { '"' + $_ + '"' }) -join ' '
$process = Start-Process powershell.exe -ArgumentList $argumentLine -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError ($log + '.stderr')
if ($process.ExitCode -ne 0) { throw "SVG preview failed: $log" }
$preview = [Drawing.Bitmap]::new((Join-Path $defaultDir 'output/preview/slide_001_001.png'))
try {
    if ($preview.Width -ne 1920 -or $preview.Height -ne 1080) { throw 'Incorrect preview dimensions' }
    foreach ($point in @(@(12,12),@(1480,12),@(1480,250),@(304,850))) {
        $pixel = $preview.GetPixel($point[0],$point[1])
        if ($pixel.R -ne 204 -or $pixel.G -ne 204 -or $pixel.B -ne 204) { throw "Missing SVG border at $point" }
    }
} finally { $preview.Dispose() }
Write-Output "PASS default SVG browser composition. Artifacts: $TestDirectory"
