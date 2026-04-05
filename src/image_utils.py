"""
画像ユーティリティ

resize_for_api:  AIへの送信前リサイズ（長辺1024px）→ base64文字列を返す
resize_for_pptx: PPTX埋め込み前リサイズ（長辺2048px）→ 一時ファイルパスを返す
"""

import base64
import io
import tempfile
from pathlib import Path

from PIL import Image


def _resize_image(img: Image.Image, max_side: int) -> Image.Image:
    """長辺が max_side を超える場合にリサイズ（アスペクト比維持）"""
    w, h = img.size
    if max(w, h) <= max_side:
        return img
    if w >= h:
        new_w = max_side
        new_h = round(h * max_side / w)
    else:
        new_h = max_side
        new_w = round(w * max_side / h)
    return img.resize((new_w, new_h), Image.LANCZOS)


def _to_rgb(img: Image.Image) -> Image.Image:
    """RGBA / P モードを RGB に変換（JPEG保存のため）"""
    if img.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode == "RGBA":
            bg.paste(img, mask=img.split()[3])
        else:
            bg.paste(img)
        return bg
    if img.mode != "RGB":
        return img.convert("RGB")
    return img


def resize_for_api(image_path: Path, max_side: int = 1024) -> tuple[str, str]:
    """
    画像をリサイズしてbase64エンコードする（Gemini API送信用）

    Returns:
        (base64_data_uri, original_info_str)
        base64_data_uri: "data:image/jpeg;base64,..."
    """
    img = Image.open(image_path)
    orig_size = img.size

    img = _to_rgb(img)
    img = _resize_image(img, max_side)

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=90)
    b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")

    info = f"{image_path.name}: {orig_size[0]}x{orig_size[1]} → {img.size[0]}x{img.size[1]}"
    return f"data:image/jpeg;base64,{b64}", info


def resize_for_pptx(image_path: Path, max_side: int = 2048) -> Path:
    """
    画像をリサイズして一時ファイルに保存する（PPTX埋め込み用）

    リサイズ不要な場合は元のパスをそのまま返す
    リサイズが必要な場合は一時ファイルに保存してそのパスを返す
    """
    img = Image.open(image_path)
    w, h = img.size

    if max(w, h) <= max_side and img.mode not in ("RGBA", "LA", "P"):
        return image_path  # そのまま使える

    img = _to_rgb(img)
    img = _resize_image(img, max_side)

    suffix = ".jpg"
    tmp = tempfile.NamedTemporaryFile(
        suffix=suffix, delete=False, prefix=f"pptx_{image_path.stem}_"
    )
    img.save(tmp.name, format="JPEG", quality=92)
    tmp.close()
    return Path(tmp.name)
