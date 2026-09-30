param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateSet('init', 'validate', 'preview', 'build')]
    [string]$Command,

    [Parameter(Position = 1, Mandatory = $true)]
    [string]$ProjectPath,
    [string]$FrameImage,
    [ValidateRange(1, 100000)][int]$PreviewFrom = 1,
    [ValidateRange(1, 100000)][int]$PreviewTo = 100000
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$script:Actions = @('idle', 'wave', 'nod', 'think', 'point', 'cheer', 'walk', 'surprise')
$script:DefaultVoice = @{
    name = 'kokuren_3rd'
    speaker_uuid = '38d7216c-e595-4d8f-b06c-1fc376e47c0a'
    style_name = 'ノーマル'
    style_id = 1069147200
    engine_url = 'http://127.0.0.1:10101'
}
$script:DefaultLayout = @{
    slide = @(16, 16, 1440, 810)
    subtitle = @(350, 870, 1528, 178)
    notes_top = @(1498, 34, 388, 182)
    notes_bottom = @(1498, 274, 388, 532)
    character = @(0, 740, 330, 332)
}

function Resolve-ProjectFile([string]$Path) {
    $resolved = [System.IO.Path]::GetFullPath($Path)
    if (Test-Path -LiteralPath $resolved -PathType Container) {
        $resolved = Join-Path $resolved 'project.json'
    }
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "Project file not found: $resolved"
    }
    return $resolved
}

function Resolve-Asset([string]$Base, [string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    if ([System.IO.Path]::IsPathRooted($Value)) { return [System.IO.Path]::GetFullPath($Value) }
    return [System.IO.Path]::GetFullPath((Join-Path $Base $Value))
}

function Read-Project([string]$Path) {
    $file = Resolve-ProjectFile $Path
    $json = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
    return [pscustomobject]@{ Data = ($json | ConvertFrom-Json); Base = (Split-Path -Parent $file); File = $file }
}

function Split-Script([string]$Text) {
    $clean = ($Text -replace "`r", '') -replace "`n", ''
    return @([regex]::Split($clean, '(?<=[。！？!?])') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-Box($Layout, [string]$Name, [int[]]$Default, [int]$Width, [int]$Height) {
    $value = $Layout.$Name
    if ($null -eq $value) {
        $scale = $Height / 1080.0
        return @($Default | ForEach-Object { [int][math]::Round($_ * $scale) })
    }
    if (@($value).Count -ne 4) { throw "layout.$Name must be [x, y, width, height]" }
    return @($value | ForEach-Object { [int]$_ })
}

function Get-ActionGif([string]$AnimationsDir, [string]$Action) {
    if ($script:Actions -notcontains $Action) { throw "Unknown character action: $Action" }
    $manifestPath = Join-Path $AnimationsDir 'manifest.json'
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $entry = $manifest.actions.$Action
    $preview = if ($entry -and $entry.preview) { $entry.preview } else { "$Action.gif" }
    $file = Resolve-Asset $AnimationsDir $preview
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Animation GIF not found: $file" }
    return $file
}

function Test-Project($Project, [string]$Base) {
    $errors = [System.Collections.Generic.List[string]]::new()
    $width = [int]$Project.canvas.width
    $height = [int]$Project.canvas.height
    if ($width -le 0 -or $height -le 0 -or ($width * 9 -ne $height * 16)) {
        $errors.Add('canvas must be a positive 16:9 size, such as 1920x1080')
    }
    if ([int]$Project.fps -le 0 -or [int]$Project.fps -gt 60) { $errors.Add('fps must be between 1 and 60') }
    $background = Resolve-Asset $Base $Project.assets.background
    if (-not $background -or -not (Test-Path -LiteralPath $background -PathType Leaf)) {
        $errors.Add('assets.background must name an existing image')
    }
    $animations = Resolve-Asset $Base $Project.assets.animations
    if (-not $animations -or -not (Test-Path -LiteralPath (Join-Path $animations 'manifest.json') -PathType Leaf)) {
        $errors.Add('assets.animations must contain manifest.json')
    }
    if ($Project.assets.bgm) {
        $bgm = Resolve-Asset $Base $Project.assets.bgm
        if (-not (Test-Path -LiteralPath $bgm -PathType Leaf)) { $errors.Add("assets.bgm not found: $bgm") }
    }
    $slides = @($Project.slides)
    if ($slides.Count -eq 0) { $errors.Add('slides must be a non-empty array') }
    $ids = @{}
    $defaults = @{
        slide = $script:DefaultLayout.slide
        subtitle = $script:DefaultLayout.subtitle
        notes_top = $script:DefaultLayout.notes_top
        notes_bottom = $script:DefaultLayout.notes_bottom
        character = $script:DefaultLayout.character
    }
    $layout = $Project.layout
    if ($null -ne $Project.renderer.min_text_pixels -and ([double]$Project.renderer.min_text_pixels -lt 12 -or [double]$Project.renderer.min_text_pixels -gt 72)) {
        $errors.Add('renderer.min_text_pixels must be between 12 and 72 (1080p reference pixels)')
    }
    if ($null -ne $Project.renderer.min_contrast -and ([double]$Project.renderer.min_contrast -lt 1 -or [double]$Project.renderer.min_contrast -gt 21)) {
        $errors.Add('renderer.min_contrast must be between 1 and 21')
    }
    foreach ($key in @('subtitle_font_size','note_top_font_size','note_font_size')) {
        if ($null -ne $layout.$key -and ([double]$layout.$key -lt 20 -or [double]$layout.$key -gt 160)) {
            $errors.Add("layout.$key must be between 20 and 160 pixels")
        }
    }
    foreach ($key in @('subtitle_stroke','subtitle_outer_stroke')) {
        if ($null -ne $layout.$key -and ([double]$layout.$key -lt 0 -or [double]$layout.$key -gt 24)) {
            $errors.Add("layout.$key must be between 0 and 24 pixels")
        }
    }
    if ($layout.subtitle_color -and [string]$layout.subtitle_color -notmatch '^#[0-9a-fA-F]{6}$') {
        $errors.Add('layout.subtitle_color must be a #RRGGBB color')
    }
    if ($layout.character_crop) {
        $crop = @($layout.character_crop)
        if ($crop.Count -ne 4 -or $crop[0] -lt 0 -or $crop[1] -lt 0 -or $crop[2] -le 0 -or $crop[3] -le 0) {
            $errors.Add('layout.character_crop must be [x,y,width,height] in source animation pixels')
        } elseif ($animations -and (Test-Path -LiteralPath (Join-Path $animations 'manifest.json'))) {
            foreach ($action in $script:Actions) {
                try {
                    $im = [Drawing.Image]::FromFile((Get-ActionGif $animations $action))
                    try { if ($crop[0]+$crop[2] -gt $im.Width -or $crop[1]+$crop[3] -gt $im.Height) { $errors.Add("character_crop exceeds $action animation") } }
                    finally { $im.Dispose() }
                } catch { $errors.Add($_.Exception.Message) }
            }
        }
    }
    if ($width -gt 0 -and $height -gt 0) {
        foreach ($name in $defaults.Keys) {
            try {
                $box = Get-Box $layout $name $defaults[$name] $width $height
                if ($box[0] -lt 0 -or $box[1] -lt 0 -or $box[2] -le 0 -or $box[3] -le 0 -or
                    ($box[0] + $box[2]) -gt $width -or ($box[1] + $box[3]) -gt $height) {
                    $errors.Add("layout.$name must be within the canvas")
                }
            } catch { $errors.Add($_.Exception.Message) }
        }
        try {
            $sub = Get-Box $layout 'subtitle' $defaults.subtitle $width $height
            $char = Get-Box $layout 'character' $defaults.character $width $height
            if ($sub[0] -lt ($char[0] + $char[2]) -and ($sub[0] + $sub[2]) -gt $char[0] -and
                $sub[1] -lt ($char[1] + $char[3]) -and ($sub[1] + $sub[3]) -gt $char[1]) {
                $errors.Add('layout.subtitle overlaps layout.character')
            }
        } catch { }
        try {
            $top = Get-Box $layout 'notes_top' $defaults.notes_top $width $height
            $bottom = Get-Box $layout 'notes_bottom' $defaults.notes_bottom $width $height
            if ($top[0] -ne $bottom[0] -or $top[2] -ne $bottom[2] -or
                $bottom[1] -lt ($top[1] + $top[3] + [math]::Max(12, [int]($height / 90))) -or
                $top[3] -gt [int](($top[3] + $bottom[3]) * 0.4) -or $bottom[3] -lt ($top[3] * 2)) {
                $errors.Add('layout.notes_top and notes_bottom must share a left edge, leave a visible gap, and reserve roughly 1/4 and 3/4 of the notes panel')
            }
        } catch { }
    }
    # Edge also draws captions/notes for raster slides.
    try { $null = Get-BrowserPath $Project $Base }
    catch { $errors.Add($_.Exception.Message) }
    $slideIndex = 0
    foreach ($slide in $slides) {
        $slideIndex++
        $id = if ($slide.id) { [string]$slide.id } else { [string]$slideIndex }
        if ($ids.ContainsKey($id)) { $errors.Add("duplicate slide id: $id") } else { $ids[$id] = $true }
        $imageValue = if ($slide.html) { [string]$slide.html } else { [string]$slide.image }
        $image = Resolve-Asset $Base $imageValue
        if (-not $image -or -not (Test-Path -LiteralPath $image -PathType Leaf)) {
            $errors.Add("slides[$slideIndex] HTML/image file not found: $imageValue")
        } elseif ([System.IO.Path]::GetExtension($image).ToLowerInvariant() -notin @('.html', '.htm', '.png', '.jpg', '.jpeg', '.bmp', '.gif')) {
            $errors.Add("slides[$slideIndex] must use HTML, PNG, JPEG, BMP, or GIF")
        }
        $sentences = Split-Script ([string]$slide.script)
        if ($sentences.Count -eq 0) { $errors.Add("slides[$slideIndex].script is empty") }
        $motions = @($slide.motions)
        if ($motions.Count -gt $sentences.Count) { $errors.Add("slides[$slideIndex].motions has too many entries") }
        foreach ($motion in $motions) {
            if ($script:Actions -notcontains [string]$motion) { $errors.Add("slides[$slideIndex] unknown motion: $motion") }
        }
        $ttsTexts = @($slide.tts_texts)
        if ($ttsTexts.Count -gt $sentences.Count) { $errors.Add("slides[$slideIndex].tts_texts has too many entries") }
        foreach ($reading in $ttsTexts) {
            if ($null -ne $reading -and $reading -isnot [string]) { $errors.Add("slides[$slideIndex].tts_texts entries must be strings") }
        }
        if ($animations -and (Test-Path -LiteralPath (Join-Path $animations 'manifest.json') -PathType Leaf)) {
            foreach ($motion in (@($motions | ForEach-Object { [string]$_ }) + @('idle') | Select-Object -Unique)) {
                try { $null = Get-ActionGif $animations $motion }
                catch { $errors.Add("slides[$slideIndex] missing animation for '$motion'") }
            }
        }
    }
    if ([int]$Project.voice.style_id -le 0) { $errors.Add('voice.style_id must be a positive integer') }
    return $errors
}

function New-DefaultFrame([string]$Path) {
    # Original geometric fallback; does not redistribute the locally supplied frame image.
    $bitmap = [Drawing.Bitmap]::new(1920, 1080)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $pen = [Drawing.Pen]::new([Drawing.Color]::FromArgb(204,204,204), 8)
    try {
        $graphics.Clear([Drawing.Color]::FromArgb(30,30,30))
        foreach ($rect in @(@(12,12,1448,818), @(1480,12,428,218), @(1480,250,428,580), @(304,850,1604,218))) {
            $graphics.DrawRectangle($pen, [int]$rect[0], [int]$rect[1], [int]$rect[2], [int]$rect[3])
        }
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    } finally { $pen.Dispose(); $graphics.Dispose(); $bitmap.Dispose() }
}

function Initialize-Project([string]$Directory) {
    $directory = [System.IO.Path]::GetFullPath($Directory)
    $assets = Join-Path $directory 'assets'
    $slidesDir = Join-Path $directory 'slides'
    New-Item -ItemType Directory -Path $assets,$slidesDir -Force | Out-Null
    $localFrame = if ($FrameImage) { [IO.Path]::GetFullPath($FrameImage) } else { Join-Path $PSScriptRoot 'assets/frame-nc293888.png' }
    if ($FrameImage -and -not (Test-Path -LiteralPath $localFrame -PathType Leaf)) { throw "Frame image not found: $localFrame" }
    if (Test-Path -LiteralPath $localFrame -PathType Leaf) {
        Copy-Item -LiteralPath $localFrame -Destination (Join-Path $assets 'frame.png') -Force
    } else {
        New-DefaultFrame (Join-Path $assets 'frame.png')
    }
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'animations') -Destination (Join-Path $assets 'animations') -Recurse -Force
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'assets\fonts') -Destination (Join-Path $assets 'fonts') -Recurse -Force
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'assets\katex') -Destination (Join-Path $assets 'katex') -Recurse -Force
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'assets/slide-theme.css') -Destination $assets -Force
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'assets/slide-math.js') -Destination $assets -Force
    foreach ($template in Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'assets/slide-templates') -Filter '*.html') {
        Copy-Item -LiteralPath $template.FullName -Destination (Join-Path $slidesDir ($template.BaseName + '.template.html')) -Force
    }
    New-SampleHtml (Join-Path $slidesDir '001.html')
    $project = [ordered]@{
        version = 1
        title = (Split-Path $directory -Leaf)
        canvas = [ordered]@{ width = 1920; height = 1080 }
        fps = 30
        output = 'output/final.mp4'
        assets = [ordered]@{ background = 'assets/frame.png'; animations = 'assets/animations'; bgm = '' }
        voice = $script:DefaultVoice
        audio = [ordered]@{ bgm_volume = 0.2 }
        renderer = [ordered]@{ audit_slides = $true; min_text_pixels = 26; min_contrast = 3; require_diagram_description = $true }
        layout = [ordered]@{
            slide = $script:DefaultLayout.slide; subtitle = $script:DefaultLayout.subtitle
            notes_top = $script:DefaultLayout.notes_top; notes_bottom = $script:DefaultLayout.notes_bottom
            character = $script:DefaultLayout.character
            subtitle_font_size = 64; note_font_size = 34; note_top_font_size = 44
            subtitle_color = '#ff3434'; subtitle_stroke = 4; subtitle_outer_stroke = 9
            character_crop = @(36, 57, 184, 148)
            fonts = [ordered]@{ slide = 'Noto Sans JP'; subtitle = 'M PLUS Rounded 1c'; notes = 'Noto Sans JP' }
        }
        slides = @([ordered]@{
            id = 1; html = 'slides/001.html'; script = "こんにちは。`nここにナレーションを書きます。"
            motions = @('wave', 'idle'); tts_texts = @(); note_top = '今回の要点'
            note_bottom = '背景や理由、具体例などをここに補足します。スライドだけでは伝わりにくい前提や、誤解しやすい点も短く説明し、ナレーションの要約だけで終わらない内容にしてください。'
        })
    }
    $json = ConvertTo-Json -InputObject $project -Depth 100
    [System.IO.File]::WriteAllText((Join-Path $directory 'project.json'), $json, [System.Text.UTF8Encoding]::new($false))
    Write-Output "Created $directory"
}

function New-SampleHtml([string]$Path) {
    $html = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'assets/slide-templates/comparison.html'), [Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($Path, $html, [System.Text.UTF8Encoding]::new($false))
}

function Get-BrowserPath($Project, [string]$Base) {
    if ($Project.renderer.browser) {
        $configured = Resolve-Asset $Base ([string]$Project.renderer.browser)
        if (Test-Path -LiteralPath $configured -PathType Leaf) { return $configured }
        throw "renderer.browser was not found: $configured"
    }
    $appPath = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe' -ErrorAction SilentlyContinue
    if ($appPath -and (Test-Path -LiteralPath $appPath.'(default)')) { return [string]$appPath.'(default)' }
    $command = Get-Command msedge,chrome -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { return $command.Source }
    throw 'HTML slide rendering requires Microsoft Edge or Google Chrome. Install Edge or set renderer.browser in project.json.'
}

function Get-TextHash([string]$Value) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value))).Replace('-', '').Substring(0, 12)) }
    finally { $sha.Dispose() }
}

function Invoke-FFmpeg([string[]]$Arguments) {
    & $script:FFmpegPath @Arguments
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg failed with exit code $LASTEXITCODE" }
}

function Get-Narration([string]$Text, [string]$Path, $Voice, [string]$EngineUrl) {
    if (Test-Path -LiteralPath $Path -PathType Leaf) { return }
    $styleId = if ($Voice.style_id) { [int]$Voice.style_id } else { [int]$script:DefaultVoice.style_id }
    $queryUri = "$($EngineUrl.TrimEnd('/'))/audio_query?text=$([uri]::EscapeDataString($Text))&speaker=$styleId"
    $query = Invoke-RestMethod -Method Post -Uri $queryUri -TimeoutSec 60
    $body = ConvertTo-Json -InputObject $query -Depth 100
    Invoke-WebRequest -Method Post -Uri "$($EngineUrl.TrimEnd('/'))/synthesis?speaker=$styleId" `
        -ContentType 'application/json' -Body $body -OutFile $Path -TimeoutSec 240 -UseBasicParsing | Out-Null
}

function Build-Video($Project, [string]$Base) {
    $Project | Add-Member -NotePropertyName basePath -NotePropertyValue $Base -Force
    $ffmpegCommand = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if (-not $ffmpegCommand) { throw 'FFmpeg is required and must be on PATH.' }
    $script:FFmpegPath = $ffmpegCommand.Source
    $output = Resolve-Asset $Base ([string]$Project.output)
    if (-not $output) { $output = Join-Path $Base 'output/final.mp4' }
    $work = Join-Path (Split-Path -Parent $output) (([IO.Path]::GetFileNameWithoutExtension($output)) + '_work')
    $frameDir = Join-Path $work 'frames'; $audioDir = Join-Path $work 'audio'; $segmentDir = Join-Path $work 'segments'
    New-Item -ItemType Directory -Path $frameDir,$audioDir,$segmentDir -Force | Out-Null
    $engineUrl = if ($Project.voice.engine_url) { [string]$Project.voice.engine_url } else { [string]$script:DefaultVoice.engine_url }
    $fps = [int]$Project.fps
    $animationsDir = Resolve-Asset $Base $Project.assets.animations
    $segments = [System.Collections.Generic.List[string]]::new()
    $slideIndex = 0
    foreach ($slide in @($Project.slides)) {
        $slideIndex++
        $slidePath = Resolve-Asset $Base $(if ($slide.html) { [string]$slide.html } else { [string]$slide.image })
        if ([IO.Path]::GetExtension($slidePath).ToLowerInvariant() -in @('.html','.htm')) {
            $rendered = Join-Path $frameDir ('slide_{0:D3}_source.png' -f $slideIndex)
            Convert-HtmlSlide $slidePath $rendered $Project
            $slide | Add-Member -NotePropertyName renderedImage -NotePropertyValue $rendered -Force
        }
    }
    $sequence = 0
    $slideIndex = 0
    foreach ($slide in @($Project.slides)) {
        $slideIndex++
        $sentences = Split-Script ([string]$slide.script)
        $motions = @($slide.motions); $ttsTexts = @($slide.tts_texts)
        $chunkIndex = 0
        foreach ($sentence in $sentences) {
            $chunkIndex++; $sequence++
            $motion = if ($chunkIndex -le $motions.Count -and $motions[$chunkIndex-1]) { [string]$motions[$chunkIndex-1] } else { 'idle' }
            $ttsText = if ($chunkIndex -le $ttsTexts.Count -and $ttsTexts[$chunkIndex-1]) { [string]$ttsTexts[$chunkIndex-1] } else { $sentence }
            $chunkName = '{0:D3}_{1:D3}' -f $slideIndex,$chunkIndex
            $framePath = Join-Path $frameDir "$chunkName.png"
            $cache = Get-TextHash ($ttsText + "`0" + $Project.voice.speaker_uuid + "`0" + $Project.voice.style_id + "`0" + $engineUrl)
            $audioPath = Join-Path $audioDir "$($chunkName)_$cache.wav"
            $segmentPath = Join-Path $segmentDir "$chunkName.mp4"
            New-VideoFrame $Project $Base $slide $sentence $framePath
            Get-Narration $ttsText $audioPath $Project.voice $engineUrl
            $gif = Get-ActionGif $animationsDir $motion
            $char = Get-Box $Project.layout 'character' $script:DefaultLayout.character ([int]$Project.canvas.width) ([int]$Project.canvas.height)
            $cropFilter = if ($Project.layout.character_crop) { $c = @($Project.layout.character_crop); "crop=$($c[2]):$($c[3]):$($c[0]):$($c[1])," } else { '' }
            Invoke-FFmpeg @('-y','-loop','1','-framerate',"$fps",'-i',$framePath,'-stream_loop','-1','-i',$gif,'-i',$audioPath,
                '-filter_complex',"[1:v]${cropFilter}scale=$($char[2]):$($char[3]):force_original_aspect_ratio=decrease:flags=lanczos[char];[0:v][char]overlay=x='$($char[0])+($($char[2])-overlay_w)/2':y='$($char[1])+$($char[3])-overlay_h':eof_action=repeat[v]",
                '-map','[v]','-map','2:a:0','-c:v','libx264','-preset','medium','-crf','18','-r',"$fps",'-fps_mode','cfr','-pix_fmt','yuv420p',
                '-c:a','aac','-b:a','192k','-ar','48000','-shortest','-movflags','+faststart',$segmentPath)
            $segments.Add($segmentPath)
            Write-Output "[$sequence] slide=$($slide.id) motion=$motion : $sentence"
        }
    }
    $concatFile = Join-Path $work 'concat.txt'
    $concat = foreach ($segment in $segments) { "file '$($segment.Replace("'", "'\''"))'" }
    [System.IO.File]::WriteAllLines($concatFile, [string[]]$concat, [System.Text.UTF8Encoding]::new($false))
    $narrationOnly = Join-Path (Split-Path -Parent $output) (([IO.Path]::GetFileNameWithoutExtension($output)) + '_narration.mp4')
    New-Item -ItemType Directory -Path (Split-Path -Parent $output) -Force | Out-Null
    Invoke-FFmpeg @('-y','-f','concat','-safe','0','-i',$concatFile,'-c:v','libx264','-preset','medium','-crf','18','-pix_fmt','yuv420p',
        '-r',"$fps",'-c:a','aac','-b:a','192k','-ar','48000','-movflags','+faststart',$narrationOnly)
    $bgm = Resolve-Asset $Base ([string]$Project.assets.bgm)
    if ($bgm -and (Test-Path -LiteralPath $bgm -PathType Leaf)) {
        $volume = if ($Project.audio.bgm_volume -ne $null) { [double]$Project.audio.bgm_volume } else { 0.2 }
        Invoke-FFmpeg @('-y','-i',$narrationOnly,'-stream_loop','-1','-i',$bgm,'-filter_complex',
            "[1:a]volume=$volume[bgm];[0:a][bgm]amix=inputs=2:duration=first:dropout_transition=2[a]",'-map','0:v:0','-map','[a]',
            '-c:v','copy','-c:a','aac','-b:a','192k','-ar','48000','-movflags','+faststart','-shortest',$output)
    } else { Copy-Item -LiteralPath $narrationOnly -Destination $output -Force }
    Write-Output "Wrote $output"
}

function Render-Preview($Project, [string]$Base) {
    if ($PreviewFrom -gt $PreviewTo -or $PreviewFrom -gt @($Project.slides).Count) { throw 'Preview range does not include any slides' }
    $Project | Add-Member -NotePropertyName basePath -NotePropertyValue $Base -Force
    $previewDir = Resolve-Asset $Base 'output/preview'
    New-Item -ItemType Directory -Path $previewDir -Force | Out-Null
    $slideIndex = 0
    foreach ($slide in @($Project.slides)) {
        $slideIndex++
        if ($slideIndex -lt $PreviewFrom -or $slideIndex -gt $PreviewTo) { continue }
        $slidePath = Resolve-Asset $Base $(if ($slide.html) { [string]$slide.html } else { [string]$slide.image })
        if ([IO.Path]::GetExtension($slidePath).ToLowerInvariant() -in @('.html','.htm')) {
            $rendered = Join-Path $previewDir ('slide_{0:D3}_source.png' -f $slideIndex)
            Convert-HtmlSlide $slidePath $rendered $Project
            $slide | Add-Member -NotePropertyName renderedImage -NotePropertyValue $rendered -Force
        }
        $sentences = Split-Script ([string]$slide.script)
        $motions = @($slide.motions)
        for ($chunk = 0; $chunk -lt $sentences.Count; $chunk++) {
            $action = if ($chunk -lt $motions.Count -and $motions[$chunk]) { [string]$motions[$chunk] } else { 'idle' }
            $target = Join-Path $previewDir ('slide_{0:D3}_{1:D3}.png' -f $slideIndex,($chunk + 1))
            New-VideoFrame $Project $Base $slide $sentences[$chunk] $target $true $action
            Write-Output "Preview: $target"
        }
    }
}

. (Join-Path $PSScriptRoot 'render-frame.ps1')

try {
    if ($Command -eq 'init') { Initialize-Project $ProjectPath; exit 0 }
    $loaded = Read-Project $ProjectPath
    $problems = Test-Project $loaded.Data $loaded.Base
    if ($problems.Count -gt 0) { $problems | ForEach-Object { Write-Error $_ }; exit 2 }
    if ($Command -eq 'validate') { Write-Output "Valid project: $($loaded.File) ($(@($loaded.Data.slides).Count) slides)"; exit 0 }
    if ($Command -eq 'preview') { Render-Preview $loaded.Data $loaded.Base; exit 0 }
    if ($Command -eq 'build') { Build-Video $loaded.Data $loaded.Base }
} catch {
    Write-Error ("$($_.Exception.Message)`n$($_.ScriptStackTrace)")
    exit 1
}
