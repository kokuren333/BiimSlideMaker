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

既定レイアウトは同梱SVGの矩形枠（16:9、1920×1080、30fps）です。スライドは `[16,16,1440,810]`、右上見出しは `[1498,34,388,182]`、右下補足は `[1498,274,388,532]`、字幕は `[350,870,1528,178]`、キャラは `[0,740,330,332]`。キャラは最前面に描画し、スライドへの軽い重なりを許容します。キャラと字幕の重なりは検証エラーです。

字幕は M PLUS Rounded 1c ExtraBold（800、64pxから自動調整）、赤字・黒4px/白9pxの二重フチです。見出しは Noto Sans JP 800（44px）、補足は同600（34px）。スライドも同梱のNoto Sans JPを使い、HTML/CSSでサイズ・ウェイトを編集します。フォントはすべてローカルの `@font-face` で読み込むためインストール不要です。両書体のOFLを `assets/fonts/` に保持します。KaTeXのMITライセンスも維持します。出力はH.264/CRF 18、AAC 192kbps/48kHz、`yuv420p`、MP4 faststartです。

HTMLスライドを2560×1440で撮影し、字幕・ノートもEdgeの2倍解像度で描画して、最終1920×1080へ高品質縮小します。旧版の1280×720画像の拡大とGDIの既定テキスト描画から変更しました。出力PNGとMP4のサイズは従来どおりです。プレビューはGIFの代表的な静止状態、動画はアニメーションです。

`layout.character_crop = [36,57,184,148]` は同梱GIF全フレームの可視領域を覆うクロップです。透明余白を除き、縦横比を保ってキャラボックスの下中央へ配置します。別素材に交換した場合はクロップを再設定するか削除してください。プレビューと動画は同じ配置計算を使います。

スライド検査は `preview` と `build` で既定有効（設定省略時も有効）です。`validate` は設定・ファイルの検証です。ブラウザ上で次を検出すると生成を停止し、PNGと同じ場所に `*.audit.json` を保存します。

- CSSのtransform/zoomと最終スライド枠への縮小を考慮した文字サイズ。既定下限は完成1080p画面で26px。
- 文字同士の重なり、枠外、図の領域外、overflowによるクリッピング。
- 計算できる単色背景に対する低コントラスト（既定比3未満）。画像・グラデーション背景は目視対象。
- 図が伝える結論の記載漏れ、比較・手順の項目不足／重なり、分数の分子・分母名の欠落。
- 部分／全体の値と一致しない割合バー、不正な数値、画像の読み込み失敗。

字幕・ノートが縮小しても収まらない場合もエラーになります。図の主張が事実として正しいか、画像内の文字、任意の図形の意味まで自動判定する機能ではありません。Agentは全スライドを完成枠内で目視し、結果と修正箇所をプロジェクトの説明に記録します。任意の図や画像については検査レポートにも目視対象を記録します。

新規プロジェクトには `assets/slide-theme.css` と `slides/comparison.template.html`、`flow.template.html`、`fraction.template.html`、`proportion.template.html` を同梱します。比較・手順・分数・割合の用途に合うものをコピーし、説明文・値を編集します。図領域を `data-diagram="comparison|flow|fraction|proportion|custom"` とし、`data-message="図から伝える結論"` または表示される `.diagram-caption` を書きます。分数の `data-numerator` / `data-denominator`、割合の `data-part` / `data-total` とバー幅は検査対象です。配色・サイズ・余白は共通CSSで管理し、内容が収まらなければ縮小せず分割してください。

`renderer.min_text_pixels`（1080p基準）、`min_contrast`、`require_diagram_description` をJSONで設定できます。既存プロジェクトにも検査が適用されます。カスタム図には説明を追加し、小さい文字は拡大してください。互換性調査のため検査を無効化する場合だけ `audit_slides: false` を明示し、その理由と目視確認結果を記録します。Agentの通常制作では無効化して完成扱いにしません。

部分確認は `preview projects\my-video -PreviewFrom 14 -PreviewTo 18` のように指定できます（スライドの配列順、preview専用）。

## ローカル素材と公開範囲

`init`は既定で同梱の [assets/frame-default.svg](assets/frame-default.svg) をプロジェクトの `assets/frame.svg` にコピーし、その枠に合うスライド・右上見出し・右下補足・下部字幕・左下キャラクターの配置を `project.json` に設定します。追加の枠指定やダウンロードは不要です。SVGは1920×1080のviewBoxと矩形・パスで構成し、スライド領域は透明です。座標・色をテキストで編集でき、拡大しても枠線の鮮明さを保ちます。

外部の枠を使う場合は `init projects/my-video -FrameImage "C:\path\to\frame.svg"` のように明示指定できます。PNG・SVG等の入力画像は、拡張子を保持してプロジェクトの `assets/frame.<拡張子>` にコピーされます。別形式の枠では `layout` の各 `[x,y,幅,高さ]` も調整してください。既存プロジェクトでは `assets.background` と `layout` を編集できます。

提供画像、個別の制作プロジェクト（BGM・台本・動画・プレビュー等）はGit対象外です。フォントの再配布条件と出典は `assets/fonts/README.md` および各OFLを参照してください。

## Agentの制作ガイド

[AGENTS.md](AGENTS.md)がテーマ、長さ、希望デザイン、参考ソース、備考を確認する手順を定めています。[Biim動画Skill](skills/biim-video/SKILL.md)は枠ごとの記述方法とBiim文化を調査した設計ノートを案内し、[Aivis読み変換Skill](skills/aivis-pronunciation/SKILL.md)は誤読しやすい語だけを合成入力用に変換します。Biimシステムの由来は[作者インタビュー](https://denfaminicogamer.jp/interview/190514c)を中心に調べました。

## 旧Python GUI

`movie_maker_gui.py`は既存PDF/YAMLプロジェクト向けに残しています。旧GUIを使う場合のみPythonと`requirements.txt`のライブラリが必要です。新しいAgent主導の制作にはPowerShell CLIを使ってください。

## 共通描画の回帰テスト

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/frame-default.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/slide-quality.ps1
```

`frame-default.ps1` は既定SVGの選択と配置、SVG・PNGの明示指定時のコピー、ブラウザ合成後の枠線と出力サイズを検証します。

特定の動画プロジェクトに依存せず、新規プロジェクト作成、枠の正確なコピー、4種の正常な図、同梱KaTeX数式、文字サイズ・CSS縮小・重なり・低コントラスト・クリッピング・図の説明欠落・割合バーの誤り・分母欠落・小さな出力枠の拒否を16ケースで検証します。結果とログはGit対象外の `output/quality-test-*/` に保存します。ブラウザとPowerShell 5.1以降が必要です。音声APIやFFmpegは使用しません。
