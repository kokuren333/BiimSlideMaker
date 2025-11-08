# BiimSlideMaker MovieMaker GUI

BiimSlideMaker MovieMaker GUI は、Markdown/Marp で作成したスライドと YAML 台本を、AivisSpeech と ffmpeg を介して 1920x1080 の解説動画へ一気通貫で変換するための Tkinter 製ツールです。PDF→PNG 変換、音声合成、BGM ミックス付きの映像書き出しまでを 3 ステップで完了できます。

---

## リポジトリ構成
| パス | 役割 |
| --- | --- |
| `movie_maker_gui.py` | GUI 本体。PDF 変換、AivisSpeech 連携、ffmpeg による結合処理を実装。 |
| `prompt.txt` | LLM（Gemini 推奨）に与えるシステムプロンプト例。テーマや目標枚数を追記して利用する。 |
| `test.md` / `test.css` / `test.pdf` | Marp スライドのサンプル。`test.md` に `test.css` を適用して PDF 化した結果が `test.pdf`。 |
| `test.yaml` | YAML 台本サンプル。`slides` 配列で `id`／`script`／`note_top`／`note_bottom` を紐づける。 |
| `biimslide_1920x1080.png` | 既定の 1920x1080 背景画像。 |
| `(Glass Weather).mp3` | 既定の BGM。GUI の「BGM」欄で参照される。 |

---

## 想定ワークフロー
1. Gemini などの LLM に `prompt.txt` をシステムプロンプトとして貼り付け、スライド化したい内容と目標枚数を指示する。
2. 返ってきた Markdown を `test.md` のように保存し、必要なら CSS（`test.css`）を調整する。
3. Marp で PDF を生成する。例：`marp test.md --theme ./test.css --pdf`
4. 同時に出力された YAML 台本（例: `test.yaml`）の `slides[].id` を PDF のページ番号と対応させる。
5. AivisSpeech Engine を起動し、ffmpeg のパスが通った状態で `python movie_maker_gui.py` を実行する。
6. GUI の「1. スライド生成」→「2. 音声合成」→「3. 動画出力」または「一括実行」で `final.mp4` を得る。

---

## 前提条件とセットアップ
- **Python 3.9+**：依存ライブラリをインストールします。
  ```powershell
  pip install -r requirements.txt
  ```
- **Marp CLI（PDF 化が前提）**：`marp test.md --theme ./test.css --pdf`
- **AivisSpeech Engine が起動していること**：既定で `http://127.0.0.1:10101` にアクセスします。`/speakers` が返る状態を確認してください。
- **ffmpeg のパスが通っていること**：`ffmpeg -version` で確認し、通っていない場合は GUI の「ffmpeg 実行ファイル」にフルパスを設定します。
- **フォント・背景・BGM**：字幕用／ノート用フォント、背景 PNG、BGM を GUI で差し替え可能です。既定値はリポジトリの同名ファイルを参照します。

---

## Gemini（長文コンテキスト）を使ったスライド原稿作成
1. Gemini Advanced などの長文コンテキスト対応モデルを開く。
2. システムプロンプト欄に `prompt.txt` の全文を貼り付ける。
3. ユーザープロンプトで「スライド化したい内容」「目標スライド枚数」「必ず伝えたいポイント」を指示する。
4. 出力された Markdown／YAML を必要に応じて修正し、Marp 用 Markdown と `slides` 配列を整える。
5. 追加修正（ノート追記、語尾調整など）が必要な場合は、同じコンテキストで Gemini に追い指示する。

---

## YAML 台本フォーマット
`test.yaml` と同様に、最上位キーは `slides` です。
| キー | 説明 |
| --- | --- |
| `id` | 1 から始まるスライド番号。PDF のページ番号と一致させる。 |
| `script` | ナレーション本文。GUI 内で句読点ごとに分割され、AivisSpeech で音声化される。 |
| `note_top` | 右上ノート枠に表示する短い要点。 |
| `note_bottom` | 右下ノート枠に表示する詳細メモ。複数行は `|`（リテラルブロック）で記述する。 |

UTF-8（BOM 無し）を推奨しますが、GUI 側で Shift_JIS などにも自動フォールバックします。

---

## MovieMaker GUI の使い方
1. **起動**：`python movie_maker_gui.py`
2. **入力**：PDF と YAML を指定し、出力ディレクトリ／マニフェスト／最終 MP4 の保存先を設定します。
3. **AivisSpeech**：Engine URL・Speaker ID・並列ワーカー数を設定し、「話者一覧」で `/speakers` の結果から ID を選択できます。
4. **合成素材**：背景 PNG、字幕フォント、ノートフォント、BGM、ffmpeg 実行ファイルを必要に応じて変更します。
5. **処理**：
   - `1. スライド生成`：PDF を 1280x720 PNG に書き出し。
   - `2. 音声合成`：`script` を分割して WAV 化し、マニフェストに記録。
   - `3. 動画出力`：テンプレ背景にテキストを描画→音声と結合→concat→BGM 追加で `final.mp4` を作成。
6. **成果物確認**：`slides`/`audio`/`frames`/`segments` などの生成物と `final.mp4` を確認します。

---

## トラブルシュート
- 文字化けする場合は Markdown/YAML を UTF-8 で保存し直す。
- ffmpeg が見つからない場合はフルパスを指定する。
- AivisSpeech が応答しない場合は Engine 側ログを確認し、必要なら `/initialize_speaker` を実行する（GUI のチェックボックスで可）。
- BGM が短い場合でも `-stream_loop -1` で自動ループしますが、音量バランスは `movie_maker_gui.py` 内の `volume=0.2` を調整してください。

---

## 次のステップ例
1. `prompt.txt` に自社用語や注意事項を追記し、チーム用テンプレートを整備する。
2. `test.css` をブランドカラーやフォントに合わせてカスタマイズする。
3. Marp→PDF→GUI の実行をスクリプト化し、定例レポート動画を自動生成するパイプラインへ発展させる。

