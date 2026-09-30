# BiimSlideMaker

Windows向けのAgent主導動画制作CLIです。PowerShell/.NETでプロジェクト、HTMLスライド、字幕、Edgeによる画面合成、AivisSpeech連携を行い、動画エンコードをFFmpegへ任せます。**Pythonやpip、Node.js、追加PowerShellモジュールは不要です。**

## 必要なもの

- Windows PowerShell 5.1以降（Windows標準）
- FFmpegを`PATH`に設定
- HTMLスライド描画用のMicrosoft Edge（既定で自動検出）またはGoogle Chrome
- ナレーション生成時はAivisSpeech Engineを起動し、`http://127.0.0.1:10101`で利用可能にする

動画生成はローカルAivisSpeech Engineにも接続します。Python環境は必要ありません。

## プロジェクトを作る

```powershell
.\biim-video.ps1 init projects\my-video
.\biim-video.ps1 validate projects\my-video
.\biim-video.ps1 preview projects\my-video
.\biim-video.ps1 build projects\my-video
```

`preview` renders slide/subtitle/character composites to `output\preview\` without starting AivisSpeech or encoding a video.

`init`は編集用の`project.json`、HTMLサンプルスライド、Biim枠、キャラクター素材、Noto Sans JP・M PLUS Rounded 1cとKaTeXを含むオフライン用アセットを作成します。AgentはJSON、HTML、画像を直接編集します。HTMLではCSSで配置と画像サイズを調整でき、KaTeXで数式を記述できます。PNG、JPEG、BMP、GIFの画像スライドも引き続き利用できます。Marpやスライド作成アプリには固定しません。スクリプト実行がExecution Policyで止められる環境では、`powershell.exe -ExecutionPolicy Bypass -File .\biim-video.ps1 validate projects\my-video`のように起動できます。

既定の出力先は`projects\my-video\output\final.mp4`です。各発話のフレーム、合成音声、チャンク動画は`output\final_work\`に残します。読み上げ音声はテキストと話者設定のハッシュでキャッシュします。

## プロジェクト形式

`project.json`にHTMLスライド、ナレーション、メモ、動作、素材、画面配置を記述します。`script`は句点などで発話単位に分かれ、字幕として表示する正規表記です。太字字幕は枠内に収まるよう縮小し、必要なら折り返します。`note_top`は右枠上部1/4の見出し、`note_bottom`は下部3/4の補足欄です。どちらも左上揃えで、上段は大きい太字、下段は読みやすい本文書体を使い、各欄からはみ出ないよう縮小・折り返しし、なお収まらない場合はエラーで編集を促します。note_bottomは背景・理由・具体例などを2〜4文、目安60〜140字で説明します。読みづらい語だけをカタカナにする`tts_texts`を指定すると、字幕は変えずにAivisSpeechへの合成入力だけを上書きできます。詳細は[スキーマ](skills/biim-video/references/project-schema.md)を参照してください。

既定の音声は`kokuren_3rd`、speaker UUID `38d7216c-e595-4d8f-b06c-1fc376e47c0a`、スタイル「ノーマル」ID `1069147200`です。AivisSpeech APIにはスタイルIDを`speaker`として渡します。[AivisSpeech Engine公式資料](https://github.com/Aivis-Project/AivisSpeech-Engine)

既定レイアウトはシンプルな矩形枠（16:9、1920×1080、30fps）です。スライドは `[16,16,1440,810]`、右上見出しは `[1498,34,388,182]`、右下補足は `[1498,274,388,532]`、字幕は `[350,870,1528,178]`、キャラは `[0,740,330,332]`。キャラは最前面に描画し、スライドへの軽い重なりを許容します。キャラと字幕の重なりは検証エラーです。

字幕は M PLUS Rounded 1c ExtraBold（800、64pxから自動調整）、赤字・黒4px/白9pxの二重フチです。見出しは Noto Sans JP 800（44px）、補足は同600（34px）。スライドも同梱のNoto Sans JPを使い、HTML/CSSでサイズ・ウェイトを編集します。フォントはすべてローカルの `@font-face` で読み込むためインストール不要です。両書体のOFLを `assets/fonts/` に保持します。KaTeXのMITライセンスも維持します。出力はH.264/CRF 18、AAC 192kbps/48kHz、`yuv420p`、MP4 faststartです。

HTMLスライドを2560×1440で撮影し、字幕・ノートもEdgeの2倍解像度で描画して、最終1920×1080へ高品質縮小します。旧版の1280×720画像の拡大とGDIの既定テキスト描画から変更しました。出力PNGとMP4のサイズは従来どおりです。プレビューはGIFの代表的な静止状態、動画はアニメーションです。

`layout.character_crop = [36,57,184,148]` は同梱GIF全フレームの可視領域を覆うクロップです。透明余白を除き、縦横比を保ってキャラボックスの下中央へ配置します。別素材に交換した場合はクロップを再設定するか削除してください。プレビューと動画は同じ配置計算を使います。

`renderer.audit_slides: true` はブラウザ描画後に24px未満の文字、枠外の文字、図の領域外の文字、見出しと図の重なりを検出して停止します。新規テンプレートに設定済みです。既存プロジェクトは明示的に有効化できます。字幕・ノートのはみ出しも描画時に検出し、末尾を省略して内容を失う代わりに編集を促すエラーにします。図の意味や数値の正しさは目視でも確認してください。

部分確認は `preview projects\my-video -PreviewFrom 14 -PreviewTo 18` のように指定できます（スライドの配列順、preview専用）。

## ローカル素材と公開範囲

`init`は外部画像を必要とせず、コードでシンプルな矩形枠を生成します。利用する権利のある枠を `assets/frame-nc293888.png` に置くと、新規プロジェクトではそのローカル画像を優先します。既存プロジェクトでは `assets.background` を変更できます。提供枠のPNG、個別の制作プロジェクト（BGM・台本・動画・プレビュー等）はGit対象外です。フォントの再配布条件と出典は `assets/fonts/README.md` および各OFLを参照してください。

## Agentの制作ガイド

[AGENTS.md](AGENTS.md)がテーマ、長さ、希望デザイン、参考ソース、備考を確認する手順を定めています。[Biim動画Skill](skills/biim-video/SKILL.md)は枠ごとの記述方法とBiim文化を調査した設計ノートを案内し、[Aivis読み変換Skill](skills/aivis-pronunciation/SKILL.md)は誤読しやすい語だけを合成入力用に変換します。Biimシステムの由来は[作者インタビュー](https://denfaminicogamer.jp/interview/190514c)を中心に調べました。

## 旧Python GUI

`movie_maker_gui.py`は既存PDF/YAMLプロジェクト向けに残しています。旧GUIを使う場合のみPythonと`requirements.txt`のライブラリが必要です。新しいAgent主導の制作にはPowerShell CLIを使ってください。
