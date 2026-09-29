# BiimSlideMaker

Windows向けのAgent主導動画制作CLIです。PowerShell/.NETでプロジェクト、字幕、画面合成、AivisSpeech連携を行い、動画エンコードをFFmpegへ任せます。**CLI利用時にPythonやpip、Node.js、追加PowerShellモジュールは不要です。**

## 必要なもの

- Windows PowerShell 5.1以降（Windows標準）
- FFmpegを`PATH`に設定
- ナレーション生成時はAivisSpeech Engineを起動し、`http://127.0.0.1:10101`で利用可能にする

動画生成はローカルAivisSpeech Engineにも接続します。Python環境は必要ありません。

## プロジェクトを作る

```powershell
.\biim-video.ps1 init projects\my-video
.\biim-video.ps1 validate projects\my-video
.\biim-video.ps1 build projects\my-video
```

`init`は編集用の`project.json`、サンプルスライド、Biim枠、キャラクター素材を作成します。AgentはJSONとスライド画像を直接編集します。画像形式はPNG、JPEG、BMP、GIFに対応し、Marpやスライド作成アプリには固定しません。スクリプト実行がExecution Policyで止められる環境では、`powershell.exe -ExecutionPolicy Bypass -File .\biim-video.ps1 validate projects\my-video`のように起動できます。

既定の出力先は`projects\my-video\output\final.mp4`です。各発話のフレーム、合成音声、チャンク動画は`output\final_work\`に残します。読み上げ音声はテキストと話者設定のハッシュでキャッシュします。

## プロジェクト形式

`project.json`にスライド画像、ナレーション、メモ、動作、素材、画面配置を記述します。`script`は字幕として表示する正規表記です。読みづらい語だけをカタカナにする`tts_texts`を指定すると、字幕は変えずにAivisSpeechへの合成入力だけを上書きできます。詳細は[スキーマ](skills/biim-video/references/project-schema.md)を参照してください。

既定の音声は`kokuren_3rd`、speaker UUID `38d7216c-e595-4d8f-b06c-1fc376e47c0a`、スタイル「ノーマル」ID `1069147200`です。AivisSpeech APIにはスタイルIDを`speaker`として渡します。[AivisSpeech Engine公式資料](https://github.com/Aivis-Project/AivisSpeech-Engine)

既定レイアウトは16:9、1920×1080、30fpsです。キャラクターは左下（約x=35–285）、字幕はx=330から始まり、領域の重なりを検証します。出力はH.264/CRF 18、AAC 192kbps/48kHz、`yuv420p`、MP4 faststartです。

## Agentの制作ガイド

[AGENTS.md](AGENTS.md)がテーマ、長さ、希望デザイン、参考ソース、備考を確認する手順を定めています。[Biim動画Skill](skills/biim-video/SKILL.md)は枠ごとの記述方法とBiim文化を調査した設計ノートを案内し、[Aivis読み変換Skill](skills/aivis-pronunciation/SKILL.md)は誤読しやすい語だけを合成入力用に変換します。Biimシステムの由来は[作者インタビュー](https://denfaminicogamer.jp/interview/190514c)を中心に調べました。

## 旧Python GUI

`movie_maker_gui.py`は既存PDF/YAMLプロジェクト向けに残しています。旧GUIを使う場合のみPythonと`requirements.txt`のライブラリが必要です。新しいAgent主導の制作にはPowerShell CLIを使ってください。
