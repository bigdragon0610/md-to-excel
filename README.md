# Markdown → Excel 変換ガイド

画像付きMarkdownを、Excelのスライド風レポートに変換するプロジェクトです。

```
Markdown → Pythonスクリプト → Excel
```

---

## 必要なツール

- **Python 3.8 以上**

### 使用するPythonライブラリ

| ライブラリ | 用途 |
|-----------|------|
| [openpyxl](https://openpyxl.readthedocs.io/) | Excelファイル（.xlsx）の生成。セル結合・色・罫線・画像の埋め込みなどに対応 |
| [Pillow](https://pillow.readthedocs.io/) | 画像サイズの取得。Excelに貼り付ける際の縮小計算に使用 |

どちらも `requirements.txt` に記載済みです。

---

## ステップ1：Markdownを作る

### 書き方ルール

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

---

## ステップ2：Pythonスクリプトを作る

`md_to_excel.py` を参照してください。このスクリプトを元に改変・拡張できます。

### スクリプトの仕様

- 入力：Markdownファイル（`.md`）
- 出力：Excelファイル（`.xlsx`）
- `---` でMarkdownを分割し、1枚のシートに縦スクロールで並べる
- スライド間にはオレンジの帯を挿入して視覚的に区切る
- グリッド線は非表示にし、列幅を統一してキャンバス風のレイアウトにする

---

## ステップ3：実行する

### 仮想環境のセットアップ（初回のみ）

```bash
cd /path/to/your/project
python3 -m venv .venv
source .venv/bin/activate   # Mac / Linux
.venv\Scripts\activate      # Windows
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

`md_to_excel.py` 冒頭の `THEME` と定数を変更することで見た目を調整できます。

```python
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

CANVAS_COLS      = 12     # スライドの横幅（列数）
COL_WIDTH        = 10.5   # 各列の幅（文字単位）
TITLE_ROW_HEIGHT = 42     # タイトルバーの行の高さ
IMG_MAX_WIDTH_PX = 760    # 画像の最大表示幅（ピクセル）
```

### 調整例

| やりたいこと | 変更箇所 |
|-------------|---------|
| タイトルを青系にしたい | `title_bg` を `"1A3A5C"` などに変更 |
| アクセントを緑にしたい | `accent` を `"2ECC71"` などに変更 |
| スライドをもっと広くしたい | `CANVAS_COLS` を `14` や `16` に増やす |
| 画像をもっと大きく表示したい | `IMG_MAX_WIDTH_PX` を `900` などに増やす |
| タイトルのフォントサイズを変えたい | `render_title()` 内の `Font(size=...)` を変更 |

---

## ファイル構成

```
your-project/
├── README.md               # このファイル
├── requirements.txt        # 必要なライブラリ一覧
├── md_to_excel.py          # 変換スクリプト
├── your-report.md          # Markdownファイル
├── your-report.xlsx        # 生成されたExcel
└── images/                 # 画像フォルダ
    ├── image1.png
    └── image2.jpg
```
