# Browser text rendering uses local fonts, including variable weights. No system font installation.
function Get-FileUri([string]$Path) { return ([uri]::new([IO.Path]::GetFullPath($Path))).AbsoluteUri }
function Escape-Html([string]$Text) { return [System.Net.WebUtility]::HtmlEncode($Text) }

function Save-BrowserImage([string]$HtmlPath, [string]$Target, $Project, [int]$Width, [int]$Height, [bool]$Downsample = $false) {
    $browser = Get-BrowserPath $Project $Project.basePath
    $profile = Join-Path ([IO.Path]::GetTempPath()) ('biim-edge-' + [guid]::NewGuid().ToString('N'))
    $capture = if ($Downsample) { $Target + '.2x.png' } else { $Target }
    $arguments = @('--headless=new','--disable-gpu','--no-first-run','--hide-scrollbars','--allow-file-access-from-files',
        '--force-device-scale-factor=2',"--window-size=$Width,$Height",'--virtual-time-budget=5000',
        '--run-all-compositor-stages-before-draw','--dump-dom',"--user-data-dir=$profile", "--screenshot=$capture", (Get-FileUri $HtmlPath))
    try {
        $argumentLine = ($arguments | ForEach-Object { '"' + ($_ -replace '"','\"') + '"' }) -join ' '
        $domPath = $Target + '.dom.txt'
        $process = Start-Process -FilePath $browser -ArgumentList $argumentLine -Wait -PassThru -WindowStyle Hidden -RedirectStandardOutput $domPath
        if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $capture)) { throw "Browser capture failed: $HtmlPath" }
        $dom = [IO.File]::ReadAllText($domPath)
        if ($dom -match 'data-slide-audit="([^"]+)"') {
            $report = [Net.WebUtility]::HtmlDecode($Matches[1])
            [IO.File]::WriteAllText(($Target + '.audit.json'), $report, [Text.UTF8Encoding]::new($false))
        }
        if ($dom -match 'data-slide-issues="([^"]+)"') { throw "Slide readability check: $([Net.WebUtility]::HtmlDecode($Matches[1])) ($HtmlPath)" }
        if ($dom -match 'data-overflow="true"') { throw "Text exceeds its frame. Edit the sentence/note in project.json. Inspect $HtmlPath" }
        if (($Downsample -or $Project.renderer.audit_slides -ne $false) -and $dom -notmatch 'data-ready="true"') { throw "Render audit/fonts did not finish: $HtmlPath" }
        if ($Downsample) {
            $source = [Drawing.Image]::FromFile($capture)
            $bitmap = [Drawing.Bitmap]::new($Width, $Height)
            $g = [Drawing.Graphics]::FromImage($bitmap)
            try {
                $g.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $g.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $g.DrawImage($source, 0, 0, $Width, $Height)
                $bitmap.Save($Target, [Drawing.Imaging.ImageFormat]::Png)
            } finally { $g.Dispose(); $bitmap.Dispose(); $source.Dispose() }
        }
    } finally {
        if ($Downsample -and (Test-Path -LiteralPath $capture)) { Remove-Item -LiteralPath $capture -Force }
        if ($domPath -and (Test-Path -LiteralPath $domPath)) { Remove-Item -LiteralPath $domPath -Force }
        # Only remove the checked, uniquely generated browser profile inside TEMP.
        $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        if ($profile.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $profile -Leaf) -like 'biim-edge-*') {
            Remove-Item -LiteralPath $profile -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-LocalFontCss([string]$Base) {
    $noto = Resolve-Asset $Base 'assets/fonts/NotoSansJP-Variable.ttf'
    if (-not (Test-Path -LiteralPath $noto)) { $noto = Join-Path $PSScriptRoot 'assets/fonts/NotoSansJP-Variable.ttf' }
    $rounded = Resolve-Asset $Base 'assets/fonts/MPLUSRounded1c-ExtraBold.ttf'
    if (-not (Test-Path -LiteralPath $rounded)) { $rounded = Join-Path $PSScriptRoot 'assets/fonts/MPLUSRounded1c-ExtraBold.ttf' }
    return "@font-face{font-family:'Noto Sans JP';src:url('$(Get-FileUri $noto)');font-weight:100 900} @font-face{font-family:'M PLUS Rounded 1c';src:url('$(Get-FileUri $rounded)');font-weight:800}"
}

function Convert-HtmlSlide([string]$Path, [string]$Target, $Project) {
    $family = if ($Project.layout.fonts.slide) { [string]$Project.layout.fonts.slide } else { 'Noto Sans JP' }
    $family = $family.Replace("'", '').Replace('<', '')
    $style = "<style>$(Get-LocalFontCss $Project.basePath) html,body{font-family:'$family','Noto Sans JP',sans-serif !important}</style>"
    $auditEnabled = $Project.renderer.audit_slides -ne $false
    $box = Get-Box $Project.layout 'slide' $script:DefaultLayout.slide ([int]$Project.canvas.width) ([int]$Project.canvas.height)
    $config = [ordered]@{
        slideScale = [math]::Min($box[2] / 1280.0, $box[3] / 720.0)
        minTextPixels = $(if ($Project.renderer.min_text_pixels) { [double]$Project.renderer.min_text_pixels } else { 26 }) * ([int]$Project.canvas.height / 1080.0)
        minContrast = $(if ($Project.renderer.min_contrast) { [double]$Project.renderer.min_contrast } else { 3 })
        requireDiagramDescription = $Project.renderer.require_diagram_description -ne $false
    }
    $auditJs = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'slide-audit.js'), [Text.Encoding]::UTF8)
    $audit = if ($auditEnabled) { '<script>window.biimAuditConfig=' + (ConvertTo-Json $config -Compress) + ';' + $auditJs + '</script>' } else { '' }
    $html = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8).Replace('</head>', $style + $audit + '</head>')
    $renderPath = Join-Path (Split-Path $Path -Parent) ('.biim-render-' + [guid]::NewGuid().ToString('N') + '.html')
    try {
        [IO.File]::WriteAllText($renderPath, $html, [Text.UTF8Encoding]::new($false))
        Save-BrowserImage $renderPath $Target $Project 1280 720
    } finally { Remove-Item -LiteralPath $renderPath -Force -ErrorAction SilentlyContinue }
}

function Get-BoxCss([int[]]$Box) { return "left:$($Box[0])px;top:$($Box[1])px;width:$($Box[2])px;height:$($Box[3])px;" }

function New-VideoFrame($Project, [string]$Base, $Slide, [string]$Subtitle, [string]$Target,
                        [bool]$IncludeCharacter = $false, [string]$CharacterAction = 'idle') {
    $Project | Add-Member -NotePropertyName basePath -NotePropertyValue $Base -Force
    $width = [int]$Project.canvas.width; $height = [int]$Project.canvas.height; $scale = $height / 1080.0
    $layout = $Project.layout
    New-Item -ItemType Directory -Path (Split-Path $Target -Parent) -Force | Out-Null
    $slidePath = if ($Slide.renderedImage) { [string]$Slide.renderedImage } else { Resolve-Asset $Base $(if ($Slide.html) { $Slide.html } else { $Slide.image }) }
    if (-not $Slide.renderedImage -and [IO.Path]::GetExtension($slidePath) -in @('.html','.htm')) {
        $rendered = Join-Path (Split-Path $Target -Parent) ('slide_{0:D3}_source.png' -f [int]$Slide.id)
        Convert-HtmlSlide $slidePath $rendered $Project; $slidePath = $rendered
    }
    $boxes = @{}
    foreach ($name in $script:DefaultLayout.Keys) { $boxes[$name] = Get-Box $layout $name $script:DefaultLayout[$name] $width $height }
    $subtitleFamily = if ($layout.fonts.subtitle) { $layout.fonts.subtitle } else { 'M PLUS Rounded 1c' }
    $notesFamily = if ($layout.fonts.notes) { $layout.fonts.notes } else { 'Noto Sans JP' }
    $subtitleSize = if ($layout.subtitle_font_size) { $layout.subtitle_font_size * $scale } else { 64 * $scale }
    $topSize = if ($layout.note_top_font_size) { $layout.note_top_font_size * $scale } else { 44 * $scale }
    $noteSize = if ($layout.note_font_size) { $layout.note_font_size * $scale } else { 34 * $scale }
    $color = if ($layout.subtitle_color) { $layout.subtitle_color } else { '#ff3434' }
    $stroke = if ($null -ne $layout.subtitle_stroke) { $layout.subtitle_stroke * $scale } else { 4 * $scale }
    $outer = if ($null -ne $layout.subtitle_outer_stroke) { $layout.subtitle_outer_stroke * $scale } else { 9 * $scale }
    $character = ''
    if ($IncludeCharacter) {
        $gif = Get-ActionGif (Resolve-Asset $Base $Project.assets.animations) $CharacterAction
        $box = $boxes.character
        $im = [Drawing.Image]::FromFile($gif); $iw = $im.Width; $ih = $im.Height; $im.Dispose()
        $crop = if ($layout.character_crop) { @($layout.character_crop) } else { @(0, 0, $iw, $ih) }
        $fit = [math]::Min($box[2] / $crop[2], $box[3] / $crop[3])
        $cw = $crop[2] * $fit; $ch = $crop[3] * $fit
        $cx = $box[0] + ($box[2] - $cw)/2; $cy = $box[1] + $box[3] - $ch
        $character = "<div class='character' style='left:$($cx)px;top:$($cy)px;width:$($cw)px;height:$($ch)px'><img src='$(Get-FileUri $gif)' style='position:absolute;max-width:none;width:$($iw*$fit)px;height:$($ih*$fit)px;left:$(-$crop[0]*$fit)px;top:$(-$crop[1]*$fit)px'></div>"
    }
    $subtitleText = Escape-Html $Subtitle
    $html = @"
<!doctype html><html lang="ja"><head><meta charset="utf-8"><style>
$(Get-LocalFontCss $Base)
*{box-sizing:border-box}html,body{margin:0;width:$($width)px;height:$($height)px;overflow:hidden;background:#1e1e1e}
.background{position:absolute;inset:0;width:100%;height:100%}.slide{position:absolute;object-fit:contain;$(Get-BoxCss $boxes.slide)}
.text{position:absolute;color:#f3f6ff;overflow:hidden;white-space:pre-wrap;overflow-wrap:anywhere;line-break:strict;line-height:1.45;font-family:'$(Escape-Html $notesFamily)','Noto Sans JP',sans-serif}
.top{font-weight:800;font-size:$($topSize)px;$(Get-BoxCss $boxes.notes_top)}
.bottom{font-weight:600;font-size:$($noteSize)px;$(Get-BoxCss $boxes.notes_bottom)}
.caption{position:absolute;$(Get-BoxCss $boxes.subtitle);display:flex;align-items:center;justify-content:center;padding:$($outer+3)px;text-align:center}
.caption-content{width:100%;font-family:'$(Escape-Html $subtitleFamily)','Noto Sans JP',sans-serif;font-weight:800;font-size:$($subtitleSize)px;line-height:1.35;white-space:pre-wrap;overflow-wrap:anywhere;line-break:strict}
.caption-lines{display:block;position:relative;width:100%}.caption-lines span{display:block;paint-order:stroke fill}
.caption-outer{position:absolute;inset:0;color:white;-webkit-text-stroke:$($outer)px white}.caption-inner{position:relative;color:$(Escape-Html $color);-webkit-text-stroke:$($stroke)px #080808}
.character{position:absolute;overflow:hidden;z-index:10}
</style></head><body><img class="background" src="$(Get-FileUri (Resolve-Asset $Base $Project.assets.background))"><img class="slide" src="$(Get-FileUri $slidePath)">
<div class="text top">$(Escape-Html ([string]$Slide.note_top))</div><div class="text bottom">$(Escape-Html ([string]$Slide.note_bottom))</div>
<div class="caption"><div class="caption-content"><div class="caption-lines"><span class="caption-outer" aria-hidden="true">$subtitleText</span><span class="caption-inner">$subtitleText</span></div></div></div>$character
<script>
async function fitText(){await document.fonts.ready;
for(const e of document.querySelectorAll('.text,.caption-content')){
 const base=parseFloat(getComputedStyle(e).fontSize); const min=Math.max(20*$scale,base*.7);
 const limit=e.matches('.caption-content')?e.parentElement.clientHeight-2*($outer+3):e.clientHeight;
 let size=base;
 while(size>min&&(e.scrollHeight>limit+1||e.scrollWidth>e.clientWidth+1)){size=Math.max(min,size-1);e.style.fontSize=size+'px';}
 if(e.scrollHeight>limit+1||e.scrollWidth>e.clientWidth+1){document.body.dataset.overflow='true';console.error('Text does not fit. Shorten the sentence or note.');}
}document.body.dataset.ready='true';}fitText();
</script></body></html>
"@
    # Keep the editable frame HTML beside previews for visual inspection and font/overflow checks.
    $htmlPath = [IO.Path]::ChangeExtension($Target, '.html')
    [IO.File]::WriteAllText($htmlPath, $html, [Text.UTF8Encoding]::new($false))
    Save-BrowserImage $htmlPath $Target $Project $width $height $true
}
