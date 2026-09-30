param([string]$TestDirectory)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $TestDirectory) { $TestDirectory = Join-Path $root ('output/quality-test-' + [guid]::NewGuid().ToString('N')) }
$TestDirectory = [IO.Path]::GetFullPath($TestDirectory)
New-Item -ItemType Directory -Path $TestDirectory -Force | Out-Null
Add-Type -AssemblyName System.Drawing
# A generated fixture proves frame import without redistributing a supplied image.
$frame = Join-Path $TestDirectory 'custom-frame.png'
$bitmap = [Drawing.Bitmap]::new(1920,1080); $g = [Drawing.Graphics]::FromImage($bitmap)
$g.Clear([Drawing.Color]::FromArgb(30,30,30)); $bitmap.Save($frame,[Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $bitmap.Dispose()
$projectDir = Join-Path $TestDirectory 'project'
$cli = Join-Path $root 'biim-video.ps1'
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $cli init $projectDir -FrameImage $frame
if ($LASTEXITCODE -ne 0) { throw 'init failed' }
if ((Get-FileHash $frame).Hash -ne (Get-FileHash (Join-Path $projectDir 'assets/frame.png')).Hash) { throw 'FrameImage was not copied exactly' }
$jsonPath = Join-Path $projectDir 'project.json'
$project = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
$project.slides[0].script = 'Quality check.'; $project.slides[0].motions = @('idle')
[IO.File]::WriteAllText($jsonPath, (ConvertTo-Json $project -Depth 100), [Text.UTF8Encoding]::new($false))
$slidePath = Join-Path $projectDir 'slides/001.html'
$results = [Collections.Generic.List[object]]::new()
function Test-Render([string]$Name, [string]$Html, [string]$ExpectedCode = '') {
    [IO.File]::WriteAllText($slidePath, $Html, [Text.UTF8Encoding]::new($false))
    $logPath = Join-Path $TestDirectory ($Name + '.log')
    $reportPath = Join-Path $projectDir 'output/preview/slide_001_source.png.audit.json'
    if (Test-Path -LiteralPath $reportPath) { Remove-Item -LiteralPath $reportPath -Force }
    $argumentLine = (@('-NoProfile','-ExecutionPolicy','Bypass','-File',$cli,'preview',$projectDir) | ForEach-Object { '"' + $_ + '"' }) -join ' '
    $process = Start-Process -FilePath powershell.exe -ArgumentList $argumentLine -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput $logPath -RedirectStandardError ($logPath + '.stderr')
    $exitCode = $process.ExitCode
    if (-not (Test-Path -LiteralPath $reportPath)) { throw "No browser audit report: $Name ($logPath)" }
    $report = Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $codes = @($report.issues | ForEach-Object { $_.code })
    if ($ExpectedCode) {
        if ($exitCode -eq 0 -or $codes -notcontains $ExpectedCode) { throw "Expected $ExpectedCode in $Name; got $($codes -join ',') ($logPath)" }
    } elseif ($exitCode -ne 0 -or $report.status -ne 'passed') { throw "Valid fixture failed: $Name ($logPath)" }
    $results.Add([pscustomobject]@{ name=$Name; expected=$ExpectedCode; result='passed' })
    Write-Output "PASS $Name"
}
$templateRoot = Join-Path $root 'assets/slide-templates'
foreach ($name in @('comparison','flow','fraction','proportion')) {
    Test-Render $name ([IO.File]::ReadAllText((Join-Path $templateRoot "$name.html"), [Text.Encoding]::UTF8))
}
$comparison = [IO.File]::ReadAllText((Join-Path $templateRoot 'comparison.html'), [Text.Encoding]::UTF8)
Test-Render 'local-math' ($comparison.Replace('</h1>', '</h1><div style="position:absolute;left:48px;top:218px;font-size:32px">\(E=mc^2\)</div>'))
Test-Render 'small-font' ($comparison.Replace('</head>', '<style>.detail{font-size:12px}</style></head>')) 'small-text'
Test-Render 'css-shrink' ($comparison.Replace('</head>', '<style>.card{transform:scale(.4)}</style></head>')) 'small-text'
Test-Render 'text-overlap' ($comparison.Replace('</head>', '<style>.card{position:relative}.detail{position:absolute;top:95px;left:30px}</style></head>')) 'text-overlap'
Test-Render 'low-contrast' ($comparison.Replace('</head>', '<style>.card{background:#888}.detail{color:#999}</style></head>')) 'low-contrast'
Test-Render 'clipped-text' ($comparison.Replace('</head>', '<style>.value{width:40px;height:30px;white-space:nowrap;overflow:hidden}</style></head>')) 'clipped-text'
Test-Render 'missing-message' ([regex]::Replace([regex]::Replace($comparison, ' data-message="[^"]*"', ''), '<p class="diagram-caption">.*?</p>', '')) 'missing-diagram-message'
Test-Render 'empty-item' ($comparison.Replace('<div class="cards">', '<div class="cards"><div class="card" data-item></div>')) 'unlabeled-item'
$proportion = [IO.File]::ReadAllText((Join-Path $templateRoot 'proportion.html'), [Text.Encoding]::UTF8)
Test-Render 'wrong-bar-width' ($proportion.Replace('width:15%', 'width:50%')) 'proportion-scale'
$fraction = [IO.File]::ReadAllText((Join-Path $templateRoot 'fraction.html'), [Text.Encoding]::UTF8)
Test-Render 'missing-denominator' ($fraction.Replace('data-denominator', 'data-other')) 'fraction-labels'
Test-Render 'empty-denominator' ([regex]::Replace($fraction, '<div data-denominator>[^<]*</div>', '<div data-denominator></div>')) 'fraction-labels'
# A narrow frame must reject text that was readable before final fitting.
$project.layout.slide = @(16,16,640,360)
[IO.File]::WriteAllText($jsonPath, (ConvertTo-Json $project -Depth 100), [Text.UTF8Encoding]::new($false))
Test-Render 'small-final-frame' $comparison 'small-text'
[IO.File]::WriteAllText((Join-Path $TestDirectory 'results.json'), (ConvertTo-Json @($results) -Depth 10), [Text.UTF8Encoding]::new($false))
Write-Output "All $($results.Count) browser integration checks passed. Artifacts: $TestDirectory"
