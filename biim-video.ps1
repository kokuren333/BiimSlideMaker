param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateSet('init', 'validate', 'build')]
    [string]$Command,

    [Parameter(Position = 1, Mandatory = $true)]
    [string]$ProjectPath
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
        return ,@($Default | ForEach-Object { [int][math]::Round($_ * $scale) })
    }
    if (@($value).Count -ne 4) { throw "layout.$Name must be [x, y, width, height]" }
    return ,@($value | ForEach-Object { [int]$_ })
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
        slide = @(40, 28, 1280, 720)
        subtitle = @(330, 847, 1520, 163)
        notes_top = @(1413, 66, 444, 324)
        notes_bottom = @(1413, 410, 444, 310)
        character = @(35, 795, 250, 250)
    }
    $layout = $Project.layout
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
    }
    $slideIndex = 0
    foreach ($slide in $slides) {
        $slideIndex++
        $id = if ($slide.id) { [string]$slide.id } else { [string]$slideIndex }
        if ($ids.ContainsKey($id)) { $errors.Add("duplicate slide id: $id") } else { $ids[$id] = $true }
        $image = Resolve-Asset $Base $slide.image
        if (-not $image -or -not (Test-Path -LiteralPath $image -PathType Leaf)) {
            $errors.Add("slides[$slideIndex].image not found: $($slide.image)")
        } elseif ([System.IO.Path]::GetExtension($image).ToLowerInvariant() -notin @('.png', '.jpg', '.jpeg', '.bmp', '.gif')) {
            $errors.Add("slides[$slideIndex].image must be PNG, JPEG, BMP, or GIF")
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

function New-SampleArtwork([string]$Path) {
    $bitmap = [System.Drawing.Bitmap]::new(1280, 720)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.Clear([System.Drawing.Color]::FromArgb(16, 24, 39))
    $titleFont = [System.Drawing.Font]::new('Meiryo', 60, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $bodyFont = [System.Drawing.Font]::new('Meiryo', 32, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $white = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
    $blue = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(185, 215, 255))
    $format = [System.Drawing.StringFormat]::new()
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $graphics.DrawString('タイトル', $titleFont, $white, [System.Drawing.RectangleF]::new(0, 230, 1280, 100), $format)
    $graphics.DrawString('内容に合わせて project.json とこの画像を編集', $bodyFont, $blue, [System.Drawing.RectangleF]::new(0, 350, 1280, 80), $format)
    $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    $format.Dispose(); $white.Dispose(); $blue.Dispose(); $titleFont.Dispose(); $bodyFont.Dispose(); $graphics.Dispose(); $bitmap.Dispose()
}

function Initialize-Project([string]$Directory) {
    $directory = [System.IO.Path]::GetFullPath($Directory)
    $assets = Join-Path $directory 'assets'
    $slidesDir = Join-Path $directory 'slides'
    New-Item -ItemType Directory -Path $assets,$slidesDir -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'biimslide_1920x1080.png') -Destination (Join-Path $assets 'frame.png') -Force
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'animations') -Destination (Join-Path $assets 'animations') -Recurse -Force
    New-SampleArtwork (Join-Path $slidesDir '001.png')
    $project = [ordered]@{
        version = 1
        title = (Split-Path $directory -Leaf)
        canvas = [ordered]@{ width = 1920; height = 1080 }
        fps = 30
        output = 'output/final.mp4'
        assets = [ordered]@{ background = 'assets/frame.png'; animations = 'assets/animations'; bgm = '' }
        voice = $script:DefaultVoice
        audio = [ordered]@{ bgm_volume = 0.2 }
        layout = [ordered]@{
            slide = @(40, 28, 1280, 720); subtitle = @(330, 847, 1520, 163)
            notes_top = @(1413, 66, 444, 324); notes_bottom = @(1413, 410, 444, 310)
            character = @(35, 795, 250, 250); subtitle_font_size = 54; note_font_size = 35
            font_family = 'Meiryo'
        }
        slides = @([ordered]@{
            id = 1; image = 'slides/001.png'; script = "こんにちは。`nここにナレーションを書きます。"
            motions = @('wave', 'idle'); tts_texts = @(); note_top = 'ポイント'; note_bottom = '補足説明'
        })
    }
    $json = ConvertTo-Json -InputObject $project -Depth 100
    [System.IO.File]::WriteAllText((Join-Path $directory 'project.json'), $json, [System.Text.UTF8Encoding]::new($false))
    Write-Output "Created $directory"
}

function Get-WrappedLines([System.Drawing.Graphics]$Graphics, [string]$Text, [System.Drawing.Font]$Font, [int]$MaxWidth) {
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($paragraph in ($Text -split "`n")) {
        $current = ''
        foreach ($character in $paragraph.ToCharArray()) {
            $candidate = $current + $character
            if ($current -and $Graphics.MeasureString($candidate, $Font).Width -gt $MaxWidth) {
                $lines.Add($current); $current = [string]$character
            } else { $current = $candidate }
        }
        $lines.Add($current)
    }
    return $lines
}

function Draw-TextBlock([System.Drawing.Graphics]$Graphics, [string]$Text, [int[]]$Box,
                        [System.Drawing.Font]$Font, [System.Drawing.Brush]$Brush, [bool]$Center) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $rect = [System.Drawing.RectangleF]::new($Box[0], $Box[1], $Box[2], $Box[3])
    $lines = Get-WrappedLines $Graphics $Text $Font $Box[2]
    $lineHeight = [math]::Max($Font.GetHeight($Graphics) + 3, 1)
    $maxLines = [math]::Max([int][math]::Floor($Box[3] / $lineHeight), 1)
    if ($lines.Count -gt $maxLines) { $lines = @($lines | Select-Object -First $maxLines) }
    $y = if ($Center) { $Box[1] + [math]::Max(0, ($Box[3] - ($lines.Count * $lineHeight)) / 2) } else { $Box[1] }
    $format = [System.Drawing.StringFormat]::new()
    $format.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    if ($Center) { $format.Alignment = [System.Drawing.StringAlignment]::Center }
    foreach ($line in $lines) {
        $lineRect = [System.Drawing.RectangleF]::new($rect.X, [single]$y, $rect.Width, [single]$lineHeight)
        if ($Center) {
            $shadow = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::Black)
            $shadowRect = [System.Drawing.RectangleF]::new($rect.X, [single]($y + 2), $rect.Width, [single]$lineHeight)
            $Graphics.DrawString($line, $Font, $shadow, $shadowRect, $format); $shadow.Dispose()
        }
        $Graphics.DrawString($line, $Font, $Brush, $lineRect, $format)
        $y += $lineHeight
    }
    $format.Dispose()
}

function New-VideoFrame($Project, [string]$Base, $Slide, [string]$Subtitle, [string]$Target) {
    $width = [int]$Project.canvas.width; $height = [int]$Project.canvas.height
    $layout = $Project.layout; $scale = $height / 1080.0
    $backgroundPath = Resolve-Asset $Base $Project.assets.background
    $bitmap = [System.Drawing.Bitmap]::new($width, $height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $background = [System.Drawing.Image]::FromFile($backgroundPath)
    $graphics.DrawImage($background, 0, 0, $width, $height); $background.Dispose()
    $slidePath = Resolve-Asset $Base $Slide.image
    $slideImage = [System.Drawing.Image]::FromFile($slidePath)
    $slideBox = Get-Box $layout 'slide' @(40,28,1280,720) $width $height
    $fit = [math]::Min($slideBox[2] / $slideImage.Width, $slideBox[3] / $slideImage.Height)
    $drawWidth = [int][math]::Round($slideImage.Width * $fit); $drawHeight = [int][math]::Round($slideImage.Height * $fit)
    $graphics.DrawImage($slideImage, [int]($slideBox[0] + ($slideBox[2] - $drawWidth)/2), [int]($slideBox[1] + ($slideBox[3] - $drawHeight)/2), $drawWidth, $drawHeight)
    $slideImage.Dispose()
    $subtitleBox = Get-Box $layout 'subtitle' @(330,847,1520,163) $width $height
    $topBox = Get-Box $layout 'notes_top' @(1413,66,444,324) $width $height
    $bottomBox = Get-Box $layout 'notes_bottom' @(1413,410,444,310) $width $height
    $subtitleFont = [System.Drawing.Font]::new([string]$layout.font_family, [single]([int]$layout.subtitle_font_size * $scale), [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $noteFont = [System.Drawing.Font]::new([string]$layout.font_family, [single]([int]$layout.note_font_size * $scale), [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $captionBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
    $noteBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(238,244,255))
    Draw-TextBlock $graphics $Subtitle $subtitleBox $subtitleFont $captionBrush $true
    Draw-TextBlock $graphics ([string]$Slide.note_top) $topBox $noteFont $noteBrush $false
    Draw-TextBlock $graphics ([string]$Slide.note_bottom) $bottomBox $noteFont $noteBrush $false
    New-Item -ItemType Directory -Path (Split-Path -Parent $Target) -Force | Out-Null
    $bitmap.Save($Target, [System.Drawing.Imaging.ImageFormat]::Png)
    $noteBrush.Dispose(); $captionBrush.Dispose(); $subtitleFont.Dispose(); $noteFont.Dispose(); $graphics.Dispose(); $bitmap.Dispose()
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
            $char = Get-Box $Project.layout 'character' @(35,795,250,250) ([int]$Project.canvas.width) ([int]$Project.canvas.height)
            Invoke-FFmpeg @('-y','-loop','1','-framerate',"$fps",'-i',$framePath,'-stream_loop','-1','-i',$gif,'-i',$audioPath,
                '-filter_complex',"[1:v]scale=$($char[2]):$($char[3]):force_original_aspect_ratio=decrease[char];[0:v][char]overlay=$($char[0]):$($char[1]):eof_action=repeat[v]",
                '-map','[v]','-map','2:a:0','-c:v','libx264','-preset','medium','-crf','18','-r',"$fps",'-vsync','cfr','-pix_fmt','yuv420p',
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

try {
    if ($Command -eq 'init') { Initialize-Project $ProjectPath; exit 0 }
    $loaded = Read-Project $ProjectPath
    $problems = Test-Project $loaded.Data $loaded.Base
    if ($problems.Count -gt 0) { $problems | ForEach-Object { Write-Error $_ }; exit 2 }
    if ($Command -eq 'validate') { Write-Output "Valid project: $($loaded.File) ($(@($loaded.Data.slides).Count) slides)"; exit 0 }
    if ($Command -eq 'build') { Build-Video $loaded.Data $loaded.Base }
} catch {
    Write-Error $_
    exit 1
}
