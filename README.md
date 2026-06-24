# Markdown → Excel 変換ガイド

AIで書いた画像付きMarkdownを、Excelのスライド風レポートに変換するプロジェクトです。

```
Markdown（AI作成） → Pythonスクリプト（AI生成） → Excel（人間が微調整）
```

---

## 必要なツール

- **Python 3.8 以上**（`python3 --version` で確認）
- **AI（Claude など）**：Markdownとスクリプトの生成に使用

Pythonライブラリは後述の手順でインストールします。

---

## ステップ1：Markdownを作る（AIに依頼）

AIに以下のような指示を出してMarkdownを書いてもらいます。

> 「〇〇についての報告資料をMarkdownで書いてください。
> 画像はネットから取得してimagesフォルダに保存してください。
> スライドの区切りは `---` を使ってください。」

### Markdownの書き方ルール

変換スクリプトは以下の記法に対応しています。

| 記法 | Excelでの表示 |
|------|--------------|
| `---` | スライドの区切り（オレンジの帯） |
| `# タイトル` | 大タイトルバー（紺背景） |
| `## 見出し` | スライドタイトルバー（紺背景） |
| `### 小見出し` | オレンジのセクション見出し |
| `- 箇条書き` / `1. 番号付き` | インデント付きリスト |
| `> 引用` | グレー背景のハイライトテキスト |
| `\| 表 \|` | ヘッダー付きの表（縞模様） |
| `![alt](images/xxx.png)` | 画像の埋め込み |

### サンプル構成

```markdown
# レポートタイトル
## サブタイトル

![ロゴ](images/logo.png)

---

## スライド1のタイトル

### セクション見出し

- 箇条書き1
- 箇条書き2

| 項目 | 内容 |
|------|------|
| A    | 説明 |

---

## スライド2のタイトル

> 重要なポイントはここに引用として書く

1. 番号付きリスト
2. 番号付きリスト
```

### 画像について

- `images/` フォルダに画像ファイルをまとめる
- Markdownからは `images/ファイル名.png` のように相対パスで参照する
- AIに「画像もネットから取得してimagesフォルダに保存してください」と依頼すると自動でやってもらえる

---

## ステップ2：Pythonスクリプトを作る（AIに依頼）

作成したMarkdownをAIに渡して、変換スクリプトを生成してもらいます。

> 「このMarkdownをExcelに変換するPythonスクリプトを作ってください。
> `---` ごとにスライドを区切り、1枚のシートに縦スクロールで並べてください。
> openpyxlとPillowを使ってください。」

このリポジトリの `md_to_excel.py` はサンプルとして参考にできます。

---

## ステップ3：実行する

### 仮想環境のセットアップ（初回のみ）

```bash
# プロジェクトフォルダに移動
cd /path/to/your/project

# 仮想環境を作成
python3 -m venv .venv

# 仮想環境を有効化
source .venv/bin/activate   # Mac / Linux
.venv\Scripts\activate      # Windows

# ライブラリをインストール
pip install -r requirements.txt
```

### 変換の実行

```bash
python3 md_to_excel.py your-report.md your-report.xlsx
```

### 仮想環境の終了

```bash
deactivate
```

---

## Excelの見た目を調整する

スクリプト冒頭の `THEME` と定数を変更するだけで見た目を変えられます。
コードの知識がなくてもAIに「この部分を変えて」と伝えれば調整してもらえます。

```python
# ---- ここを変えると見た目が変わる ----

THEME = {
    "title_bg":      "232F3E",   # タイトルバーの背景色（16進数カラーコード）
    "title_fg":      "FFFFFF",   # タイトルバーの文字色
    "accent":        "FF9900",   # アクセントの帯の色（スライド区切りにも使用）
    "h_bg":          "F2F3F3",   # 引用ブロックの背景色
    "text":          "16191F",   # 本文の文字色
    "table_head":    "232F3E",   # 表のヘッダー背景色
    "table_head_fg": "FFFFFF",   # 表のヘッダー文字色
    "table_stripe":  "F7F8F8",   # 表の縞模様の色
    "border":        "D5DBDB",   # 表の罫線の色
}

CANVAS_COLS    = 12     # スライドの横幅（列数）
COL_WIDTH      = 10.5   # 各列の幅（文字単位）
TITLE_ROW_HEIGHT = 42   # タイトルバーの行の高さ
IMG_MAX_WIDTH_PX = 760  # 画像の最大表示幅（ピクセル）
```

### 調整例

| やりたいこと | 変更箇所 |
|-------------|---------|
| タイトルを青系にしたい | `title_bg` を `"1A3A5C"` などに変更 |
| アクセントを緑にしたい | `accent` を `"2ECC71"` などに変更 |
| スライドをもっと広くしたい | `CANVAS_COLS` を `14` や `16` に増やす |
| 画像をもっと大きく表示したい | `IMG_MAX_WIDTH_PX` を `900` などに増やす |
| タイトルを太くしたい | `render_title()` 内の `Font(size=...)` を変更 |

---

## ファイル構成

```
your-project/
├── README.md               # このファイル
├── requirements.txt        # 必要なライブラリ一覧
├── md_to_excel.py          # 変換スクリプト（サンプル）
├── your-report.md          # AIが作成したMarkdown
├── your-report.xlsx        # 生成されたExcel
└── images/                 # 画像フォルダ
    ├── image1.png
    └── image2.jpg
```
