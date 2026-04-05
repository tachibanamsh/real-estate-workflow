"""
Step 1: common_input/ のPDFからテキスト情報を抽出し、MD形式で intermediates/text/ に保存する
使用モデル: anthropic/claude-sonnet-4.6 (via gmi-serving)
"""

import os
import base64
import io
import requests
from pathlib import Path
import yaml
from dotenv import load_dotenv
import fitz  # pymupdf
from PIL import Image

load_dotenv()

SCRIPT_DIR = Path(__file__).parent
REPO_ROOT = SCRIPT_DIR.parent


def load_config() -> dict:
    with open(REPO_ROOT / "config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def pdf_to_image_blocks(pdf_paths: list[Path], max_side: int = 1024, dpi: int = 120) -> list[dict]:
    """PDFの各ページをJPEG画像に変換してimage contentブロックのリストを返す"""
    blocks = []
    for pdf_path in pdf_paths:
        print(f"  PDF変換: {pdf_path.name}")
        doc = fitz.open(str(pdf_path))
        for i, page in enumerate(doc):
            mat = fitz.Matrix(dpi / 72, dpi / 72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # 長辺 max_side にリサイズ
            w, h = img.size
            if max(w, h) > max_side:
                scale = max_side / max(w, h)
                img = img.resize((round(w * scale), round(h * scale)), Image.LANCZOS)

            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=85)
            b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")
            blocks.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/jpeg", "data": b64},
            })
            print(f"    p{i+1}/{len(doc)}: {img.size[0]}x{img.size[1]}  {len(buf.getvalue())//1024}KB")
        doc.close()
    return blocks


def call_claude(config: dict, messages: list) -> str:
    """Claude API呼び出し（gmi-serving経由）"""
    api_key = os.environ["GMI_API_KEY"]
    url = config["api"]["claude"]["base_url"]
    model = config["api"]["claude"]["model"]

    payload = {
        "model": model,
        "max_tokens": 4096,
        "messages": messages,
    }
    resp = requests.post(
        url,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        json=payload,
        timeout=120,
    )
    resp.raise_for_status()
    data = resp.json()
    return data["content"][0]["text"]



# --- 各MDファイル生成プロンプト ---

PROMPT_PROPERTY_INFO = """
添付のPDFから不動産販促資料に必要な基本情報を抽出し、以下のMarkdown形式で出力してください。
不明な項目は「未記載」と記入してください。

```markdown
# 物件基本情報

## 物件名
name_ja: （日本語物件名）
name_en: （英語物件名 または ローマ字表記）

## 価格情報
price: （販売価格 例: ¥3億8,000万円）
yield_rate: （想定表面利回り 例: 4.2%）
land_price: （土地価格 任意）
building_price: （建物参考価格 任意）

## 所在地
address: （住所）
area: （エリア名 例: 渋谷区松濤）

## 物件概要
land_area: （土地面積 例: 180.00㎡）
building_area: （建物面積 例: 320.00㎡）
structure: （構造 例: 鉄筋コンクリート造）
floors: （階数 例: 地上3階建）
completion: （竣工予定 例: 2026年9月）
legal_restrictions: （法令制限 例: 第一種低層住居専用地域）
notes: （備考 任意）

## 交通アクセス
access1: （最寄り路線・駅・徒歩 例: 東急東横線「渋谷」駅 徒歩8分）
access2: （2番目の交通手段 任意）
access3: （3番目の交通手段 任意）
```

Markdownコードブロック（```markdown〜```）の中身だけを出力してください。前後の説明は不要です。
"""

PROMPT_RENT_ROLL = """
添付のPDFから想定賃料（RENT ROLL）情報を抽出し、以下のMarkdown形式で出力してください。
表が見当たらない場合は「データなし」と記入してください。

```markdown
# RENT ROLL

## 想定収入表
| 部屋番号 | 用途 | 占有面積 | 賃料 | 共益費 | 月額合計 |
|---------|------|---------|------|--------|---------|
（各行を記入）

## 合計
monthly_total: （月額合計 例: ¥1,320,000）
annual_total: （年間合計 例: ¥15,840,000）
```

Markdownコードブロック（```markdown〜```）の中身だけを出力してください。前後の説明は不要です。
"""

PROMPT_CONCEPT_TEXT = """
添付のPDFから設計コンセプトに関する情報を抽出し、さらにキャッチコピーを生成して、以下のMarkdown形式で出力してください。

キャッチコピーの生成ルール:
- 英語で1〜2文（20〜50語程度）
- 光・風・素材感・都市との対比など建築的価値を表現
- 高級感があり詩的なトーン
- 例: "Where natural light meets refined materiality. A residence designed for those who seek tranquility within the city."

```markdown
# コンセプト情報

## キャッチコピー（英語）
catchcopy: （英語キャッチコピー1〜2文）

## 設計コンセプト（日本語）
concept_title: （コンセプトタイトル 例: 光と風を纏う邸宅）
concept_body: |
  （コンセプト本文 3〜5文 PDFから抽出またはPDFの内容を元に生成）

## デザインの特徴（箇条書き 2〜4点）
- feature1: （例: 大開口サッシによる光と緑の取り込み）
- feature2:
- feature3:
- feature4:
```

Markdownコードブロック（```markdown〜```）の中身だけを出力してください。前後の説明は不要です。
"""

PROMPT_LOCATION_INFO = """
添付のPDFから立地・エリア情報を抽出し、以下のMarkdown形式で出力してください。

```markdown
# 立地情報

## エリア特性
area_description: |
  （エリアの魅力・特徴 2〜4文）

## 周辺施設（抽出できた範囲で記入）
- 商業施設:
- 医療:
- 教育:
- 公園・緑地:

## TRAIN ROUTE MAP
# TODO: 路線図は別途手動で追加予定
route_note: ""
```

Markdownコードブロック（```markdown〜```）の中身だけを出力してください。前後の説明は不要です。
"""


def find_pdfs(common_input_dir: Path) -> list[Path]:
    pdfs = list(common_input_dir.glob("*.pdf")) + list(common_input_dir.glob("*.PDF"))
    if not pdfs:
        raise FileNotFoundError(f"PDFが見つかりません: {common_input_dir}")
    return pdfs


def get_task_pdfs(config: dict, task_name: str, all_pdfs: list[Path]) -> list[Path]:
    """config.yaml の extraction 設定に基づいてタスクで使用するPDFリストを返す。
    source_files が未設定または空の場合は全PDFを返す。"""
    source_files = config.get("extraction", {}).get(task_name, {}).get("source_files", [])
    if not source_files:
        return all_pdfs
    matched = [p for p in all_pdfs if p.name in source_files]
    if not matched:
        print(f"  警告: [{task_name}] source_files に一致するPDFが見つかりません。全PDFを使用します。")
        return all_pdfs
    return matched


def call_claude_with_images(config: dict, image_blocks: list[dict], prompt: str) -> str:
    messages = [
        {
            "role": "user",
            "content": image_blocks + [{"type": "text", "text": prompt}],
        }
    ]
    return call_claude(config, messages)


def save_md(output_dir: Path, filename: str, content: str):
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / filename
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"  保存: {path}")


def run():
    config = load_config()
    common_input_dir = REPO_ROOT / config["paths"]["common_input"]
    output_dir = REPO_ROOT / config["paths"]["intermediates_text"]

    print("=== Step 1: テキスト情報抽出 ===")
    pdfs = find_pdfs(common_input_dir)
    print(f"対象PDF ({len(pdfs)}件): {[p.name for p in pdfs]}")

    # 各PDFを個別に画像変換してキャッシュ（タスクごとの使い回しのため）
    print("PDFをページ画像に変換中...")
    pdf_blocks: dict[str, list[dict]] = {}
    for pdf in pdfs:
        blocks = pdf_to_image_blocks([pdf], max_side=1024, dpi=120)
        pdf_blocks[pdf.name] = blocks
        print(f"  {pdf.name}: {len(blocks)} ページ")
    print(f"  合計 {sum(len(b) for b in pdf_blocks.values())} ページを変換")

    tasks = [
        ("property_info.md", PROMPT_PROPERTY_INFO, "物件基本情報", "property_info"),
        ("rent_roll.md", PROMPT_RENT_ROLL, "RENT ROLL", "rent_roll"),
        ("concept_text.md", PROMPT_CONCEPT_TEXT, "コンセプト・キャッチコピー", "concept_text"),
        ("location_info.md", PROMPT_LOCATION_INFO, "立地情報", "location_info"),
    ]

    for filename, prompt, label, task_name in tasks:
        task_pdfs = get_task_pdfs(config, task_name, pdfs)
        task_file_names = [p.name for p in task_pdfs]
        print(f"\n[{label}] 抽出中... (使用ファイル: {task_file_names})")

        # 対象ファイルの画像ブロックを結合
        image_blocks = []
        for name in task_file_names:
            image_blocks.extend(pdf_blocks[name])

        result = call_claude_with_images(config, image_blocks, prompt)
        save_md(output_dir, filename, result)

    print("\n=== Step 1 完了 ===")


if __name__ == "__main__":
    run()
