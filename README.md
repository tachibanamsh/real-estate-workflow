# 不動産販促資料生成ワークフロー

PDFの物件概要書から不動産販促スライド（PowerPoint / PDF）を自動生成するワークフロー。

## ディレクトリ構成

```
real-estate-workflow/
├── assets/
│   ├── common_input/       ← 物件概要書PDFを配置（複数可）
│   ├── page1_image/        ← 表紙用外観画像（2枚: メイン・サブ）
│   ├── page2_image/        ← 内観画像（4枚）
│   ├── page3_image/        ← コンセプト用画像（2枚）
│   ├── page4_image/        ← 設計図面（1F・2F以降）
│   ├── page5_image/        ← RENT ROLL用縦長物件画像（1枚）
│   └── page6_image/        ← 地図画像（Googleマップスクショ等）
├── outputs/
│   ├── intermediates/
│   │   ├── text/           ← Step1生成: MD形式テキスト情報
│   │   └── images/         ← Step2生成: Gemini加工済み画像
│   ├── {slug}.pptx         ← Step3生成
│   └── {slug}.pdf          ← Step4生成
├── src/
│   ├── step1_extract_text.py
│   ├── step2_process_images.py
│   ├── step3_generate_pptx.py
│   └── step4_export_pdf.py
├── config.yaml             ← 物件設定・画像加工フラグ
├── workflow.py             ← メイン実行スクリプト
└── requirements.txt
```

## スライド構成（6ページ）

| # | ページ名 | 内容 |
|---|---------|------|
| 1 | 表紙 | 外観画像2枚、物件名、キャッチコピー（Claude生成）、価格・利回り |
| 2 | 内観イメージ | 内観画像4枚グリッド |
| 3 | 設計コンセプト | コンセプト画像2枚、コンセプト文 |
| 4 | FLOOR PLAN | 設計図面（1F・2F以降）、面積情報 |
| 5 | RENT ROLL | 縦長画像、想定収入表、価格・利回り、注釈 |
| 6 | 物件立地 | 地図、物件概要、交通アクセス |

## セットアップ

```bash
pip install -r requirements.txt
cp .env.example .env
# .env に GMI_API_KEY を設定済み
```

PDF変換（Step 4）にはLibreOfficeが必要:
```bash
brew install --cask libreoffice
```

## 使い方

### 事前準備
1. `assets/common_input/` に物件概要書PDFを配置
2. `assets/pagex_image/` に各ページ用画像を配置
3. `config.yaml` の物件名・slug・画像加工フラグを設定

### 実行

```bash
# 全ステップ一括実行
python workflow.py

# ステップ個別実行
python workflow.py --step 1        # テキスト抽出のみ
python workflow.py --step 2        # 画像加工のみ
python workflow.py --step 3 4      # PPTX生成+PDF変換
python workflow.py --step 1 2 3 4  # 全ステップ（デフォルトと同じ）
```

## ワークフロー詳細

### Step 1: テキスト抽出（PDF → MD）
- モデル: `anthropic/claude-sonnet-4.6` via gmi-serving
- common_input/ のPDFをClaudeで解析し、4つのMDファイルを生成
  - `property_info.md`: 物件名・価格・利回り・所在地・面積・交通
  - `rent_roll.md`: 想定収入表
  - `concept_text.md`: コンセプト文 + **キャッチコピー（Claude生成）**
  - `location_info.md`: エリア情報・周辺施設

### Step 2: 画像加工（Gemini）
- モデル: `gemini-3.1-flash-image-preview` via gmi cloud（リクエストキュー方式）
- `config.yaml` の `image_processing` で加工するページをON/OFFで制御
- デフォルトはpage1（表紙）のみON
- 加工済み画像は `outputs/intermediates/images/` に保存

### Step 3: PowerPoint生成
- ライブラリ: python-pptx
- フォント: Noto Sans JP / Playfair Display / Helvetica Neue
- スライドサイズ: A4横（297×210mm）
- Step1/Step2の成果物を参照してスライドを生成

### Step 4: PDF変換
- LibreOffice CLI（Mac環境）
- `--headless --convert-to pdf` で変換

## TODO

- [ ] TRAIN ROUTE MAP（路線図）の自動生成（Slide 6右下）
- [ ] フォント埋め込み確認（Noto Sans JP / Playfair Display）
- [ ] Gemini APIの画像入力形式の確認（base64 data URI対応状況）
