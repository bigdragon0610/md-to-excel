# Markdown → Excel 変換ガイド

画像付きMarkdownを、Excelのレポートに変換するプロジェクトです。

```
Markdown → Pythonスクリプト → Excel
```

---

## 必要なツール

- **Python 3.8 以上**

### 使用するPythonライブラリ

| ライブラリ | 用途 |
|-----------|------|
| [openpyxl](https://openpyxl.readthedocs.io/) | Excelファイル（.xlsx）の生成。セル・罫線・画像の埋め込みなどに対応 |
| [Pillow](https://pillow.readthedocs.io/) | 画像サイズの取得と縮小計算、表の画像レンダリングに使用 |

どちらも `requirements.txt` に記載済みです。

---

## ステップ1：Markdownを作る

### 書き方ルール

変換スクリプトは以下の記法に対応しています。

| 記法 | Excelでの表示 |
|------|--------------|
| `---` | セクションの区切り（余白） |
| `# タイトル` | 大見出し（太字・大きめフォント） |
| `## 見出し` | 中見出し（太字） |
| `### 小見出し` | 小見出し（太字） |
| `- 箇条書き` / `1. 番号付き` | インデント付きリスト |
| `> 引用` | グレーのイタリック体 |
| `\| 表 \|` | 画像としてレンダリングされた表 |
| `![alt](images/xxx.png)` | 画像の埋め込み |

> **表について**：列幅がセルに依存しないよう、Pillowで画像として描画してから貼り付けています。列幅はテキストの長さに応じて自動調整されます。

### サンプル構成

```markdown
# レポートタイトル
## サブタイトル

![ロゴ](images/logo.png)

---

## セクション1のタイトル

### 小見出し

- 箇条書き1
- 箇条書き2

| 項目 | 内容          |
|------|---------------|
| ①   | 長い説明文... |

---

## セクション2のタイトル

> 補足情報はここに書く

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
- セクション間は余白のみで区切る
- グリッド線は非表示にし、列幅を統一してキャンバス風のレイアウトにする
- 日本語フォント：ヒラギノ（macOS）→ Arial Unicode（Windows/Linux）の順で自動検出

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
    "title":        "000000",  # タイトル文字色（16進数カラーコード）
    "text":         "000000",  # 本文文字色
    "muted":        "555555",  # 引用など補助テキストの色
    "table_stripe": "F2F2F2",  # 表のヘッダー背景色
    "border":       "999999",  # 罫線の色
}

CANVAS_COLS      = 12    # レイアウトの横幅（列数）
COL_WIDTH        = 11.0  # 各列の幅（文字単位）
TITLE_ROW_HEIGHT = 36    # 見出し行の高さ
IMG_MAX_WIDTH_PX = 700   # 画像・表の最大表示幅（ピクセル）
```

### 調整例

| やりたいこと | 変更箇所 |
|-------------|---------|
| タイトルをグレーにしたい | `title` を `"444444"` などに変更 |
| 引用をもっと薄くしたい | `muted` を `"AAAAAA"` などに変更 |
| レイアウトをもっと広くしたい | `CANVAS_COLS` を `14` や `16` に増やす |
| 画像・表をもっと大きく表示したい | `IMG_MAX_WIDTH_PX` を `900` などに増やす |
| フォントサイズを変えたい | `render_title()` 内の `Font(size=...)` を変更 |

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
