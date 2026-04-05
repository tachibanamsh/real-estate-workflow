"""
Step 2: config.yaml の image_processing が true のページの画像を
Gemini (Google AI Studio API) で加工し intermediates/images/ に保存する

API構成:
  - Claude (Step1): gmi cloud  → GMI_API_KEY
  - Gemini (Step2): Google AI Studio → GOOGLE_API_KEY
    endpoint: generativelanguage.googleapis.com
    base64入力対応のため、ローカル画像をそのまま送信できる
    レスポンスに画像データ(base64)が含まれるので、ダウンロード不要
"""

import os
import base64
import shutil
import requests
from pathlib import Path
import yaml
from dotenv import load_dotenv
from image_utils import resize_for_api

load_dotenv()

SCRIPT_DIR = Path(__file__).parent
REPO_ROOT = SCRIPT_DIR.parent


def load_config() -> dict:
    with open(REPO_ROOT / "config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def call_gemini_image(config: dict, img_path: Path, prompt: str) -> bytes:
    """
    Google AI Studio Gemini API で画像加工し、結果画像のbytesを返す

    リクエスト形式:
      POST /v1beta/models/{model}:generateContent?key={GOOGLE_API_KEY}
      content: [text prompt] + [inline_data base64 jpeg]
      generationConfig.responseModalities: ["TEXT", "IMAGE"]

    レスポンス: candidates[0].content.parts の中にinline_dataで画像が返る
    """
    api_key = os.environ["GOOGLE_API_KEY"]
    model = config["api"]["gemini"]["model"]
    base_url = config["api"]["gemini"]["base_url"]
    url = f"{base_url}/{model}:generateContent?key={api_key}"

    # 画像をリサイズしてbase64取得（data URI形式から本体だけ分離）
    b64_uri, resize_info = resize_for_api(img_path, max_side=1024)
    print(f"    リサイズ: {resize_info}")
    b64_data = b64_uri.split(",")[1]  # "data:image/jpeg;base64," を除く

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": "image/jpeg",
                            "data": b64_data,
                        }
                    },
                ]
            }
        ],
        "generationConfig": {
            "responseModalities": ["TEXT", "IMAGE"],
        },
    }

    resp = requests.post(url, json=payload, timeout=120)
    if not resp.ok:
        raise RuntimeError(
            f"Gemini API エラー {resp.status_code}: {resp.text[:300]}"
        )

    data = resp.json()

    # レスポンスのpartsから画像データを取り出す
    parts = data["candidates"][0]["content"]["parts"]
    for part in parts:
        inline = part.get("inlineData") or part.get("inline_data")
        if inline:
            return base64.b64decode(inline["data"])

    # 画像なし（テキストのみ）→ Noneを返してフォールバック処理へ
    return None


def save_image_bytes(image_bytes: bytes, save_path: Path):
    save_path.parent.mkdir(parents=True, exist_ok=True)
    with open(save_path, "wb") as f:
        f.write(image_bytes)
    print(f"  保存: {save_path}  ({len(image_bytes)//1024}KB)")


# --- ページ別プロンプト ---

PAGE_PROMPTS = {
    "page1": (
        "Enhance this real estate exterior photo to luxury architectural photography standard. "
        "Apply professional color grading: slightly warmer tones, increased contrast, vivid sky. "
        "Ensure sharpness and clean shadows. Maintain photorealistic look."
    ),
    "page2": (
        "Enhance this interior real estate photo with professional staging lighting. "
        "Apply warmer color temperature, increase brightness and contrast subtly. "
        "Sharpen textures of materials. Keep photorealistic."
    ),
    "page3": (
        "Enhance this architectural detail photo to emphasize material texture and quality. "
        "Apply fine-art architectural photography style: high contrast, neutral palette, "
        "sharp focus on surface details. Keep photorealistic."
    ),
    "page5": (
        "Enhance this real estate exterior/interior photo. "
        "Apply professional color correction, clean up highlights and shadows. Photorealistic."
    ),
    "page6": (
        "Enhance the clarity and contrast of this location map image. "
        "Make text and roads sharper and more readable. Keep the map style unchanged."
    ),
}


def process_page_images(config: dict, page_key: str, image_dir: Path, output_dir: Path):
    """1ページ分の画像をGeminiで加工する"""
    prompt = PAGE_PROMPTS.get(page_key, "Enhance this real estate photo professionally.")

    images = sorted(
        [p for p in image_dir.glob("*") if p.suffix.lower() in (".jpg", ".jpeg", ".png")]
    )
    if not images:
        print(f"  [{page_key}] 画像なし → スキップ")
        return

    for i, img_path in enumerate(images):
        print(f"  [{page_key}] 加工中: {img_path.name}")
        out_filename = f"{page_key}_{i+1:02d}_{img_path.stem}_processed.jpg"
        out_path = output_dir / out_filename

        image_bytes = call_gemini_image(config, img_path, prompt)

        if image_bytes is None:
            # Geminiが画像を返さなかった場合は元画像をそのままコピー
            print(f"  [{page_key}] Geminiが画像未返却 → 元画像をコピー: {img_path.name}")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(img_path, out_path)
        else:
            save_image_bytes(image_bytes, out_path)


def run():
    config = load_config()
    image_processing = config["image_processing"]
    output_dir = REPO_ROOT / config["paths"]["intermediates_images"]

    print("=== Step 2: 画像加工（Gemini / Google AI Studio） ===")

    for page_key, enabled in image_processing.items():
        if not enabled:
            print(f"[{page_key}] スキップ（config: false）")
            continue

        image_dir = REPO_ROOT / config["paths"]["assets"] / f"{page_key}_image"
        if not image_dir.exists():
            print(f"[{page_key}] ディレクトリなし → スキップ: {image_dir}")
            continue

        print(f"\n[{page_key}] 処理開始...")
        process_page_images(config, page_key, image_dir, output_dir)

    print("\n=== Step 2 完了 ===")


if __name__ == "__main__":
    run()
