"""
Step 3: intermediates/ のMDファイルと画像を使ってPowerPointを生成する
使用ライブラリ: python-pptx

スライド構成:
  Slide 1: 表紙
  Slide 2: 内観イメージ
  Slide 3: 設計コンセプト
  Slide 4: FLOOR PLAN
  Slide 5: RENT ROLL
  Slide 6: 物件立地

共通要素: 右上に英語物件名 + ページ番号
"""

import re
from pathlib import Path
import yaml
from pptx import Presentation
from pptx.util import Mm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from image_utils import resize_for_pptx

SCRIPT_DIR = Path(__file__).parent
REPO_ROOT = SCRIPT_DIR.parent

# --- カラーパレット ---
COLOR_BLACK = RGBColor(0x1A, 0x1A, 0x1A)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_GRAY = RGBColor(0x88, 0x88, 0x88)
COLOR_ACCENT = RGBColor(0xC8, 0xAA, 0x7A)  # ゴールド系アクセント
COLOR_BG = RGBColor(0xF8, 0xF6, 0xF3)      # オフホワイト背景


def load_config() -> dict:
    with open(REPO_ROOT / "config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def mm(value: float) -> Emu:
    return Mm(value)


def parse_md_value(md_text: str, key: str) -> str:
    """key: value 形式の値をMDから抽出する"""
    pattern = rf"^{re.escape(key)}:\s*(.+)$"
    m = re.search(pattern, md_text, re.MULTILINE)
    if m:
        return m.group(1).strip()
    return ""


def parse_md_multiline(md_text: str, key: str) -> str:
    """key: |\\n  行1\\n  行2 形式を抽出する"""
    pattern = rf"^{re.escape(key)}:\s*\|\s*\n((?:  .+\n?)*)"
    m = re.search(pattern, md_text, re.MULTILINE)
    if m:
        lines = m.group(1).split("\n")
        return "\n".join(l.strip() for l in lines if l.strip())
    # フォールバック: key: value 形式
    return parse_md_value(md_text, key)


def parse_table_rows(md_text: str) -> list[list[str]]:
    """Markdown表をリストに変換する"""
    rows = []
    in_table = False
    for line in md_text.split("\n"):
        if "|" in line:
            cells = [c.strip() for c in line.split("|") if c.strip()]
            if all(set(c) <= set("-:") for c in cells):
                continue  # セパレータ行スキップ
            rows.append(cells)
    return rows


def read_md(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8")


def find_images(directory: Path, processed_dir: Path, page_key: str) -> list[Path]:
    """加工済み画像が存在すればそちらを、なければ元画像を返す（jpg/png両対応）"""
    # 加工済みを優先（jpg・png両方検索）
    processed = sorted(
        [p for p in processed_dir.glob(f"{page_key}_*")
         if p.suffix.lower() in (".jpg", ".jpeg", ".png")]
    )
    if processed:
        return processed
    # 元画像フォルダ
    if not directory.exists():
        return []
    images = sorted(
        [p for p in directory.glob("*") if p.suffix.lower() in (".jpg", ".jpeg", ".png")]
    )
    return images


def set_slide_background(slide, color: RGBColor):
    from pptx.oxml.ns import qn
    from lxml import etree

    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text_box(
    slide,
    text: str,
    left, top, width, height,
    font_size: int = 11,
    bold: bool = False,
    color: RGBColor = COLOR_BLACK,
    align=PP_ALIGN.LEFT,
    font_name: str = "Noto Sans JP",
):
    from pptx.util import Pt

    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    return txBox


def add_image_safe(slide, image_path: Path, left, top, width, height):
    """
    画像をアスペクト比を保ってカバー配置する（object-fit: cover 相当）
    短辺に合わせてスケール → 長辺のはみ出し分を中央からトリミング
    """
    if not (image_path and image_path.exists()):
        return False

    resized = resize_for_pptx(image_path, max_side=2048)

    from PIL import Image as _Image
    img_w, img_h = _Image.open(resized).size

    target_w = int(width)
    target_h = int(height)
    img_aspect = img_w / img_h
    target_aspect = target_w / target_h

    pic = slide.shapes.add_picture(str(resized), int(left), int(top), target_w, target_h)

    if img_aspect > target_aspect:
        # 画像が横長 → 左右をトリミング
        visible_w = img_h * target_aspect
        crop_lr = (img_w - visible_w) / 2 / img_w
        pic.crop_left = crop_lr
        pic.crop_right = crop_lr
    else:
        # 画像が縦長 → 上下をトリミング
        visible_h = img_w / target_aspect
        crop_tb = (img_h - visible_h) / 2 / img_h
        pic.crop_top = crop_tb
        pic.crop_bottom = crop_tb

    return True


def add_image_fit(slide, image_path: Path, left, top, width, height):
    """
    画像をアスペクト比を保ってフィット配置する（object-fit: contain 相当）
    エリア内に収まるようスケール → 余白部分はスライド背景が透けて見える
    地図・図面など全体を見せたい画像に使用
    """
    if not (image_path and image_path.exists()):
        return False

    resized = resize_for_pptx(image_path, max_side=2048)

    from PIL import Image as _Image
    img_w, img_h = _Image.open(resized).size

    target_w = int(width)
    target_h = int(height)
    img_aspect = img_w / img_h
    target_aspect = target_w / target_h

    if img_aspect > target_aspect:
        # 画像が横長 → 幅に合わせてスケール、縦は縮む → 上下中央寄せ
        actual_w = target_w
        actual_h = int(target_w / img_aspect)
        offset_top = (target_h - actual_h) // 2
        slide.shapes.add_picture(str(resized), int(left), int(top) + offset_top, actual_w, actual_h)
    else:
        # 画像が縦長 → 高さに合わせてスケール、横は縮む → 左右中央寄せ
        actual_h = target_h
        actual_w = int(target_h * img_aspect)
        offset_left = (target_w - actual_w) // 2
        slide.shapes.add_picture(str(resized), int(left) + offset_left, int(top), actual_w, actual_h)

    return True


def add_header_common(slide, name_en: str, page_num: int, slide_width, margin_right):
    """右上に英語物件名とページ番号を追加（共通要素）"""
    header_text = f"{name_en}  |  {page_num:02d}"
    add_text_box(
        slide,
        header_text,
        left=slide_width - mm(80),
        top=mm(6),
        width=mm(75),
        height=mm(8),
        font_size=7,
        color=COLOR_GRAY,
        align=PP_ALIGN.RIGHT,
        font_name="Helvetica Neue",
    )


def add_separator_line(slide, left, top, width, color: RGBColor = COLOR_ACCENT):
    """アクセントラインを追加"""
    from pptx.util import Pt
    line = slide.shapes.add_connector(1, left, top, left + width, top)
    line.line.color.rgb = color
    line.line.width = Pt(0.5)


# =========================================================
# 各スライド生成関数
# =========================================================

def add_slide1_cover(prs, config, text_dir, img_dir, processed_dir):
    """スライド1: 表紙"""
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, COLOR_BG)

    W = prs.slide_width
    H = prs.slide_height
    ml = mm(config["slide"]["margin_left_mm"])
    mr = mm(config["slide"]["margin_right_mm"])
    mt = mm(config["slide"]["margin_top_mm"])

    property_info = read_md(text_dir / "property_info.md")
    concept_text = read_md(text_dir / "concept_text.md")

    name_ja = parse_md_value(property_info, "name_ja") or config["property"]["name_ja"]
    name_en = parse_md_value(property_info, "name_en") or config["property"]["name_en"]
    price = parse_md_value(property_info, "price")
    yield_rate = parse_md_value(property_info, "yield_rate")
    catchcopy = parse_md_value(concept_text, "catchcopy")

    images = find_images(
        REPO_ROOT / config["paths"]["assets"] / "page1_image",
        processed_dir,
        "page1",
    )

    # メイン画像（左〜中央大きめ）
    main_img_w = W * 0.62
    main_img_h = H
    if images:
        add_image_safe(slide, images[0], 0, 0, main_img_w, main_img_h)

    # サブ画像（右側縦長）
    sub_img_x = main_img_w + mm(4)
    sub_img_w = W - sub_img_x - mm(2)
    sub_img_h = H * 0.65
    if len(images) >= 2:
        add_image_safe(slide, images[1], sub_img_x, 0, sub_img_w, sub_img_h)

    # 右下エリア（物件名・価格）
    right_x = main_img_w + mm(4)
    right_y = sub_img_h + mm(6)
    right_w = W - right_x - mm(5)

    # 物件名（日本語）
    add_text_box(
        slide, name_ja,
        left=right_x, top=right_y, width=right_w, height=mm(12),
        font_size=13, bold=True, color=COLOR_BLACK,
        font_name="Noto Sans JP",
    )
    # キャッチコピー（英語）
    add_text_box(
        slide, catchcopy,
        left=right_x, top=right_y + mm(13), width=right_w, height=mm(16),
        font_size=7, color=COLOR_GRAY,
        font_name="Helvetica Neue",
    )
    # 価格・利回り
    price_text = f"{price}" if price else ""
    yield_text = f"想定表面利回り {yield_rate}" if yield_rate else ""
    add_text_box(
        slide, price_text,
        left=right_x, top=right_y + mm(30), width=right_w, height=mm(9),
        font_size=11, bold=True, color=COLOR_ACCENT,
        font_name="Helvetica Neue",
    )
    add_text_box(
        slide, yield_text,
        left=right_x, top=right_y + mm(40), width=right_w, height=mm(7),
        font_size=8, color=COLOR_GRAY,
        font_name="Noto Sans JP",
    )

    add_header_common(slide, name_en, 1, W, mr)
    return slide


def add_slide2_interior(prs, config, text_dir, img_dir, processed_dir):
    """スライド2: 内観イメージ（4枚グリッド）"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, COLOR_BG)

    W = prs.slide_width
    H = prs.slide_height
    ml = mm(config["slide"]["margin_left_mm"])
    mt = mm(config["slide"]["margin_top_mm"])
    mr = mm(config["slide"]["margin_right_mm"])
    mb = mm(config["slide"]["margin_bottom_mm"])

    name_en = config["property"]["name_en"]
    images = find_images(
        REPO_ROOT / config["paths"]["assets"] / "page2_image",
        processed_dir,
        "page2",
    )

    # タイトル
    add_text_box(
        slide, "I N T E R I O R",
        left=ml, top=mm(8), width=mm(80), height=mm(10),
        font_size=10, bold=True, color=COLOR_BLACK,
        font_name="Playfair Display",
    )
    add_separator_line(slide, ml, mm(19), mm(120))

    # 4枚グリッド配置（1枚大きめ）
    grid_top = mm(22)
    grid_h = H - grid_top - mb
    usable_w = W - ml - mr
    gap = mm(3)

    # レイアウト: 左大1枚(60%) / 右2枚縦並び(40%)
    # 下段: 左2枚
    # 実際は: 左大(上) + 右2枚縦 + 左下1枚（計4枚）
    # シンプルに: 上行(大1+小1) / 下行(小1+小1)
    cell_w_large = usable_w * 0.55 - gap / 2
    cell_w_small = usable_w * 0.45 - gap / 2
    row_h_top = grid_h * 0.6 - gap / 2
    row_h_bottom = grid_h * 0.4 - gap / 2

    positions = [
        # (left, top, width, height) - 大きめメイン
        (ml, grid_top, cell_w_large, row_h_top),
        # 右上小
        (ml + cell_w_large + gap, grid_top, cell_w_small, row_h_top),
        # 左下小
        (ml, grid_top + row_h_top + gap, cell_w_large, row_h_bottom),
        # 右下小
        (ml + cell_w_large + gap, grid_top + row_h_top + gap, cell_w_small, row_h_bottom),
    ]

    for i, (l, t, w, h) in enumerate(positions):
        if i < len(images):
            add_image_safe(slide, images[i], l, t, w, h)

    add_header_common(slide, name_en, 2, W, mr)
    return slide


def add_slide3_concept(prs, config, text_dir, img_dir, processed_dir):
    """スライド3: 設計コンセプト"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, COLOR_BG)

    W = prs.slide_width
    H = prs.slide_height
    ml = mm(config["slide"]["margin_left_mm"])
    mt = mm(config["slide"]["margin_top_mm"])
    mr = mm(config["slide"]["margin_right_mm"])
    mb = mm(config["slide"]["margin_bottom_mm"])

    name_en = config["property"]["name_en"]
    concept_text = read_md(text_dir / "concept_text.md")

    concept_title = parse_md_value(concept_text, "concept_title")
    concept_body = parse_md_multiline(concept_text, "concept_body")
    features_raw = re.findall(r"- feature\d+:\s*(.+)", concept_text)

    images = find_images(
        REPO_ROOT / config["paths"]["assets"] / "page3_image",
        processed_dir,
        "page3",
    )

    usable_w = W - ml - mr
    gap = mm(4)
    img_w = (usable_w - gap) / 2
    img_h = H * 0.58

    # 左右に画像2枚（上部）
    if images:
        add_image_safe(slide, images[0], ml, mt - mm(5), img_w, img_h)
    if len(images) >= 2:
        add_image_safe(slide, images[1], ml + img_w + gap, mt - mm(5), img_w, img_h)

    # 下部コンセプトテキスト
    text_top = mt + img_h - mm(2)
    text_h = H - text_top - mb

    # "Design" ラベル
    add_text_box(
        slide, "D E S I G N",
        left=ml, top=text_top, width=mm(60), height=mm(8),
        font_size=8, bold=True, color=COLOR_ACCENT,
        font_name="Playfair Display",
    )
    add_separator_line(slide, ml, text_top + mm(9), mm(100))

    # コンセプトタイトル
    add_text_box(
        slide, concept_title,
        left=ml, top=text_top + mm(11), width=usable_w * 0.5, height=mm(10),
        font_size=12, bold=True, color=COLOR_BLACK,
        font_name="Noto Sans JP",
    )
    # コンセプト本文
    add_text_box(
        slide, concept_body,
        left=ml, top=text_top + mm(22), width=usable_w * 0.5, height=mm(25),
        font_size=8, color=COLOR_BLACK,
        font_name="Noto Sans JP",
    )
    # 特徴リスト（右側）
    feature_text = "\n".join(f"• {f}" for f in features_raw)
    add_text_box(
        slide, feature_text,
        left=ml + usable_w * 0.52, top=text_top + mm(11), width=usable_w * 0.48, height=mm(36),
        font_size=8, color=COLOR_GRAY,
        font_name="Noto Sans JP",
    )

    add_header_common(slide, name_en, 3, W, mr)
    return slide


def add_slide4_floorplan(prs, config, text_dir, img_dir, processed_dir):
    """スライド4: FLOOR PLAN"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, COLOR_BG)

    W = prs.slide_width
    H = prs.slide_height
    ml = mm(config["slide"]["margin_left_mm"])
    mt = mm(config["slide"]["margin_top_mm"])
    mr = mm(config["slide"]["margin_right_mm"])
    mb = mm(config["slide"]["margin_bottom_mm"])

    name_en = config["property"]["name_en"]
    property_info = read_md(text_dir / "property_info.md")
    building_area = parse_md_value(property_info, "building_area")
    land_area = parse_md_value(property_info, "land_area")

    images = find_images(
        REPO_ROOT / config["paths"]["assets"] / "page4_image",
        processed_dir,
        "page4",
    )

    usable_w = W - ml - mr
    usable_h = H - mt - mb
    gap = mm(5)

    # 左下に1F図面（大きめ）
    img1_w = usable_w * 0.55
    img1_h = usable_h
    if images:
        add_image_safe(slide, images[0], ml, mt, img1_w, img1_h)

    # 右上に2F以降図面
    img2_x = ml + img1_w + gap
    img2_w = usable_w - img1_w - gap
    img2_h = usable_h * 0.62
    if len(images) >= 2:
        add_image_safe(slide, images[1], img2_x, mt, img2_w, img2_h)

    # 右下: FLOOR PLAN テキスト + 面積情報
    info_y = mt + img2_h + mm(6)
    add_text_box(
        slide, "F L O O R  P L A N",
        left=img2_x, top=info_y, width=img2_w, height=mm(10),
        font_size=11, bold=True, color=COLOR_BLACK,
        font_name="Playfair Display",
    )
    add_text_box(
        slide, config["property"]["name_en"],
        left=img2_x, top=info_y + mm(11), width=img2_w, height=mm(7),
        font_size=8, color=COLOR_GRAY,
        font_name="Helvetica Neue",
    )
    area_text = f"建物面積: {building_area}　土地面積: {land_area}"
    add_text_box(
        slide, area_text,
        left=img2_x, top=info_y + mm(19), width=img2_w, height=mm(7),
        font_size=8, color=COLOR_BLACK,
        font_name="Noto Sans JP",
    )

    add_header_common(slide, name_en, 4, W, mr)
    return slide


def add_slide5_rentroll(prs, config, text_dir, img_dir, processed_dir):
    """スライド5: RENT ROLL"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, COLOR_BG)

    W = prs.slide_width
    H = prs.slide_height
    ml = mm(config["slide"]["margin_left_mm"])
    mt = mm(config["slide"]["margin_top_mm"])
    mr = mm(config["slide"]["margin_right_mm"])
    mb = mm(config["slide"]["margin_bottom_mm"])

    name_en = config["property"]["name_en"]
    rent_roll_md = read_md(text_dir / "rent_roll.md")
    property_info = read_md(text_dir / "property_info.md")

    price = parse_md_value(property_info, "price")
    yield_rate = parse_md_value(property_info, "yield_rate")
    monthly_total = parse_md_value(rent_roll_md, "monthly_total")
    annual_total = parse_md_value(rent_roll_md, "annual_total")

    images = find_images(
        REPO_ROOT / config["paths"]["assets"] / "page5_image",
        processed_dir,
        "page5",
    )

    usable_w = W - ml - mr
    left_img_w = usable_w * 0.32

    # 左側縦長物件画像
    if images:
        add_image_safe(slide, images[0], ml, 0, left_img_w, H)

    # 右側エリア
    right_x = ml + left_img_w + mm(6)
    right_w = W - right_x - mr

    # タイトル
    add_text_box(
        slide, "R E N T  R O L L",
        left=right_x, top=mt, width=right_w, height=mm(10),
        font_size=11, bold=True, color=COLOR_BLACK,
        font_name="Playfair Display",
    )
    add_separator_line(slide, right_x, mt + mm(11), right_w)

    # 想定収入表（Markdown表をテーブルとして描画）
    table_rows = parse_table_rows(rent_roll_md)
    table_top = mt + mm(14)

    if table_rows:
        col_count = len(table_rows[0]) if table_rows else 6
        row_count = len(table_rows)
        table_h = Mm(6) * row_count + Mm(2)
        table_w = right_w

        table = slide.shapes.add_table(
            row_count, col_count,
            int(right_x), int(table_top),
            int(table_w), int(table_h),
        ).table

        col_widths = [int(table_w) // col_count] * col_count
        for ci, cw in enumerate(col_widths):
            table.columns[ci].width = cw

        for ri, row_cells in enumerate(table_rows):
            for ci, cell_text in enumerate(row_cells):
                if ci >= col_count:
                    break
                cell = table.cell(ri, ci)
                cell.text = cell_text
                para = cell.text_frame.paragraphs[0]
                para.alignment = PP_ALIGN.CENTER
                run = para.runs[0] if para.runs else para.add_run()
                run.font.size = Pt(7)
                run.font.name = "Noto Sans JP"
                run.font.bold = (ri == 0)
                run.font.color.rgb = COLOR_WHITE if ri == 0 else COLOR_BLACK
                # ヘッダ行の背景色
                if ri == 0:
                    from pptx.oxml.ns import qn
                    from lxml import etree
                    tc = cell._tc
                    tcPr = tc.get_or_add_tcPr()
                    solidFill = etree.SubElement(tcPr, qn("a:solidFill"))
                    srgbClr = etree.SubElement(solidFill, qn("a:srgbClr"))
                    srgbClr.set("val", "1A1A1A")

        # 合計行の下に価格・利回り
        price_y = table_top + table_h + mm(8)
    else:
        price_y = table_top + mm(8)

    # 価格・利回り
    add_text_box(
        slide, f"販売価格: {price}",
        left=right_x, top=price_y, width=right_w, height=mm(8),
        font_size=11, bold=True, color=COLOR_ACCENT,
        font_name="Noto Sans JP",
    )
    add_text_box(
        slide, f"想定表面利回り: {yield_rate}",
        left=right_x, top=price_y + mm(9), width=right_w, height=mm(7),
        font_size=9, color=COLOR_BLACK,
        font_name="Noto Sans JP",
    )
    if monthly_total:
        add_text_box(
            slide, f"月額合計: {monthly_total}　年間合計: {annual_total}",
            left=right_x, top=price_y + mm(17), width=right_w, height=mm(7),
            font_size=8, color=COLOR_GRAY,
            font_name="Noto Sans JP",
        )

    # 注釈（固定文）
    disclaimer = (
        "〈備考〉 ※上記想定収支表の賃貸面積等は現在計画中の建物概要であって実際の建物とは異なる場合がございます。"
        "詳細は検討中につき、随時最新情報をお問合せください。 "
        "※想定賃料は弊社のマーケティングによって算出した参考値であり建物竣工後の成約賃料を保証するものではないことをご了解ください。 "
        "※本資料への掲載内容は、設計段階の図面やコンセプトを基に作成しており、実際のものとは異なる可能性がございます。 "
        "※ＣＧの家具等はあくまで参考であり実際にはございません。"
    )
    add_text_box(
        slide, disclaimer,
        left=right_x, top=H - mb - mm(12), width=W - right_x - mr, height=mm(12),
        font_size=5.5, color=COLOR_GRAY,
        font_name="Noto Sans JP",
    )

    add_header_common(slide, name_en, 5, W, mr)
    return slide


def add_slide6_location(prs, config, text_dir, img_dir, processed_dir):
    """スライド6: 物件立地"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    set_slide_background(slide, COLOR_BG)

    W = prs.slide_width
    H = prs.slide_height
    ml = mm(config["slide"]["margin_left_mm"])
    mt = mm(config["slide"]["margin_top_mm"])
    mr = mm(config["slide"]["margin_right_mm"])
    mb = mm(config["slide"]["margin_bottom_mm"])

    name_en = config["property"]["name_en"]
    property_info = read_md(text_dir / "property_info.md")
    location_info = read_md(text_dir / "location_info.md")

    name_ja = parse_md_value(property_info, "name_ja") or config["property"]["name_ja"]
    address = parse_md_value(property_info, "address")
    land_area = parse_md_value(property_info, "land_area")
    building_area = parse_md_value(property_info, "building_area")
    structure = parse_md_value(property_info, "structure")
    floors = parse_md_value(property_info, "floors")
    legal = parse_md_value(property_info, "legal_restrictions")
    notes = parse_md_value(property_info, "notes")
    access_lines = [
        parse_md_value(property_info, "access1"),
        parse_md_value(property_info, "access2"),
        parse_md_value(property_info, "access3"),
    ]
    access_lines = [a for a in access_lines if a and a != "未記載"]

    images = find_images(
        REPO_ROOT / config["paths"]["assets"] / "page6_image",
        processed_dir,
        "page6",
    )

    # 左側: 地図（縦いっぱい・fitで全体表示）
    map_w = W * 0.48
    if images:
        add_image_fit(slide, images[0], 0, 0, map_w, H)

    # 右側エリア
    right_x = map_w + mm(6)
    right_w = W - right_x - mr

    # 物件概要タイトル
    add_text_box(
        slide, "P R O P E R T Y  O V E R V I E W",
        left=right_x, top=mt, width=right_w, height=mm(8),
        font_size=8, bold=True, color=COLOR_BLACK,
        font_name="Playfair Display",
    )
    add_separator_line(slide, right_x, mt + mm(9), right_w)

    # 物件概要テーブル
    overview_rows = [
        ("物件名", name_ja),
        ("所在地", address),
        ("土地面積", land_area),
        ("建物面積", building_area),
        ("構造", structure),
        ("階数", floors),
        ("法令制限", legal),
        ("備考", notes),
    ]
    overview_rows = [(k, v) for k, v in overview_rows if v and v != "未記載"]

    row_h = mm(8)
    for i, (label, value) in enumerate(overview_rows):
        y = mt + mm(12) + i * row_h
        add_text_box(
            slide, label,
            left=right_x, top=y, width=mm(25), height=row_h,
            font_size=7, color=COLOR_GRAY,
            font_name="Noto Sans JP",
        )
        add_text_box(
            slide, value,
            left=right_x + mm(26), top=y, width=right_w - mm(26), height=row_h,
            font_size=7, color=COLOR_BLACK,
            font_name="Noto Sans JP",
        )

    # 交通アクセス
    access_y = mt + mm(12) + len(overview_rows) * row_h + mm(8)
    add_text_box(
        slide, "T R A I N  R O U T E",
        left=right_x, top=access_y, width=right_w, height=mm(8),
        font_size=8, bold=True, color=COLOR_BLACK,
        font_name="Playfair Display",
    )
    add_separator_line(slide, right_x, access_y + mm(9), right_w)

    for i, access in enumerate(access_lines):
        add_text_box(
            slide, f"• {access}",
            left=right_x, top=access_y + mm(12) + i * mm(8),
            width=right_w, height=mm(8),
            font_size=7.5, color=COLOR_BLACK,
            font_name="Noto Sans JP",
        )

    # TODO: TRAIN ROUTE MAP（路線図）は別途追加予定
    todo_y = access_y + mm(12) + len(access_lines) * mm(8) + mm(5)
    add_text_box(
        slide, "[ TODO: TRAIN ROUTE MAP ]",
        left=right_x, top=todo_y, width=right_w, height=mm(20),
        font_size=7, color=COLOR_GRAY,
        font_name="Helvetica Neue",
    )

    add_header_common(slide, name_en, 6, W, mr)
    return slide


# =========================================================
# メイン
# =========================================================

def run():
    config = load_config()
    text_dir = REPO_ROOT / config["paths"]["intermediates_text"]
    processed_dir = REPO_ROOT / config["paths"]["intermediates_images"]
    output_dir = REPO_ROOT / config["paths"]["outputs"]
    output_dir.mkdir(parents=True, exist_ok=True)

    print("=== Step 3: PowerPoint生成 ===")

    # スライドサイズ設定（A4横）
    prs = Presentation()
    prs.slide_width = Mm(config["slide"]["width_mm"])
    prs.slide_height = Mm(config["slide"]["height_mm"])

    img_dir = REPO_ROOT / config["paths"]["assets"]

    print("スライド1: 表紙")
    add_slide1_cover(prs, config, text_dir, img_dir, processed_dir)

    print("スライド2: 内観イメージ")
    add_slide2_interior(prs, config, text_dir, img_dir, processed_dir)

    print("スライド3: 設計コンセプト")
    add_slide3_concept(prs, config, text_dir, img_dir, processed_dir)

    print("スライド4: FLOOR PLAN")
    add_slide4_floorplan(prs, config, text_dir, img_dir, processed_dir)

    print("スライド5: RENT ROLL")
    add_slide5_rentroll(prs, config, text_dir, img_dir, processed_dir)

    print("スライド6: 物件立地")
    add_slide6_location(prs, config, text_dir, img_dir, processed_dir)

    slug = config["property"]["slug"]
    out_path = output_dir / f"{slug}.pptx"
    prs.save(str(out_path))
    print(f"\nPowerPoint保存: {out_path}")
    print("=== Step 3 完了 ===")
    return out_path


if __name__ == "__main__":
    run()
