"""주소 기반 자동 pptx 리포트 생성기.

report_template/generate_report.js(pptxgenjs 목업)를 실제 라이브 데이터로 채우는
python-pptx 버전이다. 두 부분으로 나뉜다:

  1. 데이터 계층: fetch_report_data() — building_example.py의 기존 함수들
     (build_master_report, REB 공실률, 소상공인 상가업소, 서울 상권분석)을 호출해
     주소 하나에 대한 리포트 데이터를 한 딕셔너리로 모은다.
  2. 생성 계층: generate_pptx(data) — 그 딕셔너리를 python-pptx로 그려 pptx
     바이트를 반환한다.

정성적 항목(상권 성격/SNS 트렌드/개발호재/SWOT)은 어떤 공공API로도 자동 수집되지
않아 이 자동 생성 버전에서는 뺐다 — report_template/generate_report.js 목업에는
있지만, 임의 주소에 대해 지어낸 내용을 채우는 건 오히려 오해를 줄 수 있어서다.
"""

import copy
import datetime
import io
import json
import math
import os

import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.oxml.ns import qn

from building_example import (
    build_executive_summary,
    build_master_report,
    combine_zoning_sources,
    extract_year,
    get_building_ledger,
    get_nearby_stores,
    get_reb_vacancy_snapshot,
    get_reb_vacancy_trend,
    get_seoul_trade_area_locations,
    get_seoul_trade_area_quarter_dataset,
    find_nearest_seoul_trade_area,
    reb_current_quarter_id,
    REB_COMMERCIAL_VACANCY_STATBL_IDS,
    SEOUL_TRDAR_SALES_SERVICE,
    SEOUL_TRDAR_STORE_SERVICE,
    SEOUL_TRDAR_FLPOP_SERVICE,
    SEOUL_TRDAR_WRC_POPLTN_SERVICE,
)

# ------------------------------------------------------------------
# 색상 팔레트 / 글꼴 — 딥 네이비 베이스 + 비비드(인디고→핑크) 그라데이션 포인트
# ------------------------------------------------------------------
NAVY = RGBColor(0x16, 0x21, 0x3E)
NAVY_DARK = RGBColor(0x0B, 0x12, 0x20)
NAVY_LIGHT = RGBColor(0x2A, 0x39, 0x5C)
ICE = RGBColor(0xE8, 0xEC, 0xF2)
TERRACOTTA = RGBColor(0xA8, 0x55, 0xF7)  # 단색이 필요한 텍스트/얇은 선용 포인트(비비드 퍼플). 이름은 하위 호환을 위해 유지.
GRADIENT_A = RGBColor(0x63, 0x66, 0xF1)  # 그라데이션 시작 — 인디고
GRADIENT_B = RGBColor(0xEC, 0x48, 0x99)  # 그라데이션 끝 — 핑크
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT_DARK = RGBColor(0x22, 0x26, 0x2E)
MUTED = RGBColor(0x6B, 0x72, 0x80)
CARD_BG = RGBColor(0xF5, 0xF6, 0xF8)
GRID_LINE = RGBColor(0xDC, 0xE1, 0xE8)


def _gradient_fill(shape_or_chart_format, angle=45):
    """비비드 그라데이션(인디고→핑크) 채우기. 배지·포인트바·차트 시리즈 등 강조 도형에 쓴다."""
    fill = shape_or_chart_format.fill
    fill.gradient()
    stops = fill.gradient_stops
    stops[0].color.rgb = GRADIENT_A
    stops[0].position = 0.0
    stops[-1].color.rgb = GRADIENT_B
    stops[-1].position = 1.0
    fill.gradient_angle = angle

# Malgun Gothic을 bold로 쓰면 이 프로젝트가 QA에 쓰는 LibreOffice 렌더러에서 한글이
# 깨진 장식체로 나오는 버그가 있었다(report_template 작업 중 발견) — 실제 PowerPoint에서도
# 재현될 수 있어 위험하므로, 정상체/굵게 모두 깨끗하게 렌더링되는 Calibri로 통일한다.
# 한글은 OS 기본 대체 글꼴(Windows에서는 사실상 맑은 고딕)로 표시된다.
FONT_NAME = "Calibri"

SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5

# 지번이 바뀔 때마다 사용자가 PowerPoint에서 직접 고쳐 쓰는 4개 슬라이드(리포트 개요·
# 섹션 구분·개발호재·SWOT). report_template/manual_slides.pptx에서 그대로 복사해 넣는다 —
# 정성적 내용이라 공공API로 자동 채울 수 없어서다. 예시 콘텐츠는 연남동_상권분석1_보강본.pptx의
# 5/6/13/26페이지를 그대로 옮긴 것.
_MANUAL_SLIDES_PATH = os.path.join(os.path.dirname(__file__), "report_template", "manual_slides.pptx")


# ==================================================================
# 생성 계층 (python-pptx)
# ==================================================================
def _copy_manual_slide(prs, src_slide, src_slide_width):
    """다른 Presentation의 슬라이드를 도형·배경·이미지까지 그대로 복사해 새 슬라이드로 붙인다.
    원본 폭이 리포트 캔버스(13.333in)보다 좁으면(manual_slides.pptx는 10.83in) 좌우
    중앙 정렬로 배치해 비율 왜곡 없이 끼워 넣는다."""
    dest_slide = prs.slides.add_slide(prs.slide_layouts[6])
    for shp in list(dest_slide.shapes):
        shp._element.getparent().remove(shp._element)

    src_cSld = src_slide._element.find(qn("p:cSld"))
    dest_cSld = dest_slide._element.find(qn("p:cSld"))
    src_bg = src_cSld.find(qn("p:bg"))
    if src_bg is not None:
        dest_bg = dest_cSld.find(qn("p:bg"))
        if dest_bg is not None:
            dest_cSld.remove(dest_bg)
        dest_cSld.insert(0, copy.deepcopy(src_bg))

    x_offset = (prs.slide_width - src_slide_width) // 2
    dest_spTree = dest_slide.shapes._spTree
    for shape in src_slide.shapes:
        new_el = copy.deepcopy(shape._element)
        spPr = new_el.find(qn("p:spPr"))
        if spPr is not None:
            xfrm = spPr.find(qn("a:xfrm"))
            if xfrm is not None:
                off = xfrm.find(qn("a:off"))
                if off is not None and x_offset:
                    off.set("x", str(int(off.get("x")) + x_offset))
        if shape.shape_type == 13:  # PICTURE
            blip = new_el.find(qn("p:blipFill")).find(qn("a:blip"))
            old_rId = blip.get(qn("r:embed"))
            image_part = src_slide.part.related_part(old_rId)
            _, new_rId = dest_slide.part.get_or_add_image_part(io.BytesIO(image_part.blob))
            blip.set(qn("r:embed"), new_rId)
        dest_spTree.append(new_el)
    return dest_slide


def _insert_manual_slide(prs, manual_prs, index):
    """manual_slides.pptx의 index번째(0-based) 슬라이드를 그대로 삽입."""
    _copy_manual_slide(prs, manual_prs.slides[index], manual_prs.slide_width)


def _new_slide(prs, dark=False):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = NAVY if dark else WHITE
    return slide


def _textbox(slide, x, y, w, h, text, size=12, bold=False, color=TEXT_DARK,
             align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, italic=False, font=FONT_NAME,
             line_spacing=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.name = font
        run.font.color.rgb = color
    return box


def _rich_textbox(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE):
    """runs: [(text, {size,bold,color}), ...] 한 줄에 여러 스타일을 섞어 쓸 때."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    for text, opts in runs:
        run = p.add_run()
        run.text = text
        run.font.size = Pt(opts.get("size", 12))
        run.font.bold = opts.get("bold", False)
        run.font.name = FONT_NAME
        run.font.color.rgb = opts.get("color", TEXT_DARK)
    return box


def _section_title(slide, text):
    _textbox(slide, 0.6, 0.45, 11.3, 0.6, text, size=28, bold=True, color=NAVY)


def _page_footer(slide, address, label):
    _textbox(slide, 0.6, 7.12, 10, 0.3, f"{address}  ·  {label}", size=9, color=MUTED)


def _card(slide, x, y, w, h, fill=CARD_BG, line=GRID_LINE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.adjustments[0] = 0.08
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = line
    shape.line.width = Pt(0.75)
    shape.shadow.inherit = False
    return shape


def _add_picture_placeholder(slide, x, y, w, h, image_stream=None, idx=90):
    """PowerPoint에서 드래그&드롭으로 바로 교체할 수 있는 실제 사진 placeholder를 만든다.
    image_stream이 있으면(자동 생성 지도 등) 미리 채워 넣고, 없으면 빈 상자로 둔다 — 어느 쪽이든
    <p:ph type="pic">를 갖는 진짜 placeholder라서 파일탐색기에서 이미지를 끌어다 놓으면 그대로 바뀐다."""
    if image_stream is not None:
        shape = slide.shapes.add_picture(image_stream, Inches(x), Inches(y), Inches(w), Inches(h))
        nvPr = shape._element.find(qn("p:nvPicPr")).find(qn("p:nvPr"))
    else:
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        shape.fill.solid()
        shape.fill.fore_color.rgb = CARD_BG
        shape.line.color.rgb = GRID_LINE
        shape.line.width = Pt(1)
        shape.shadow.inherit = False
        nvPr = shape._element.find(qn("p:nvSpPr")).find(qn("p:nvPr"))
    nvPr.append(nvPr.makeelement(qn("p:ph"), {"type": "pic", "idx": str(idx)}))
    return shape


def _stat_card(slide, x, y, w, h, value, label, sub=None):
    _card(slide, x, y, w, h)
    _textbox(slide, x + 0.15, y + 0.12, w - 0.3, h * 0.45, value, size=22, bold=True,
             color=TERRACOTTA, valign=MSO_ANCHOR.BOTTOM)
    _textbox(slide, x + 0.15, y + h * 0.55, w - 0.3, h * 0.22, label, size=11, bold=True, color=TEXT_DARK)
    if sub:
        _textbox(slide, x + 0.15, y + h * 0.76, w - 0.3, h * 0.2, sub, size=9, color=MUTED)


def _table(slide, x, y, w, h, headers, rows, col_ratios=None, font_size=11):
    n_rows, n_cols = len(rows) + 1, len(headers)
    gframe = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y), Inches(w), Inches(h))
    table = gframe.table
    if col_ratios:
        total = sum(col_ratios)
        for i, ratio in enumerate(col_ratios):
            table.columns[i].width = Inches(w * ratio / total)

    def _fmt_cell(cell, text, bold, color, fill):
        cell.text = "" if text is None else str(text)
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
        cell.margin_top = cell.margin_bottom = Pt(4)
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(font_size)
            p.font.bold = bold
            p.font.name = FONT_NAME
            p.font.color.rgb = color

    for j, htext in enumerate(headers):
        _fmt_cell(table.cell(0, j), htext, True, WHITE, NAVY)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            _fmt_cell(table.cell(i, j), val, False, TEXT_DARK, WHITE)
    return table


def _style_chart_axes(chart):
    chart.has_legend = False
    chart.has_title = False
    try:
        chart.category_axis.tick_labels.font.size = Pt(9.5)
        chart.category_axis.tick_labels.font.color.rgb = MUTED
        chart.category_axis.format.line.color.rgb = GRID_LINE
        chart.value_axis.tick_labels.font.size = Pt(9.5)
        chart.value_axis.tick_labels.font.color.rgb = MUTED
        chart.value_axis.has_major_gridlines = True
        chart.value_axis.major_gridlines.format.line.color.rgb = GRID_LINE
        chart.category_axis.has_major_gridlines = False
    except Exception:
        pass


def _bar_chart(slide, x, y, w, h, categories, values, color, horizontal=False, num_fmt="0", label_size=9):
    chart_data = CategoryChartData()
    chart_data.categories = [str(c) for c in categories]
    chart_data.add_series("값", values)
    chart_type = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gframe = slide.shapes.add_chart(chart_type, Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
    chart = gframe.chart
    _style_chart_axes(chart)
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = num_fmt
    dl.number_format_is_linked = False
    dl.font.size = Pt(label_size)
    dl.font.color.rgb = TEXT_DARK
    series = plot.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = color
    series.format.line.fill.background()
    return chart


def _line_chart(slide, x, y, w, h, categories, values, color, num_fmt="0.0", label_size=9):
    chart_data = CategoryChartData()
    chart_data.categories = [str(c) for c in categories]
    chart_data.add_series("값", values)
    gframe = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w), Inches(h), chart_data)
    chart = gframe.chart
    _style_chart_axes(chart)
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = num_fmt
    dl.number_format_is_linked = False
    dl.font.size = Pt(label_size)
    dl.font.color.rgb = NAVY
    series = plot.series[0]
    series.format.line.color.rgb = color
    series.format.line.width = Pt(2.5)
    series.marker.style = 8  # circle
    series.marker.format.fill.solid()
    series.marker.format.fill.fore_color.rgb = color
    series.smooth = False
    return chart


def _cover_slide(prs, data):
    slide = _new_slide(prs, dark=True)
    for cx, cy, cw, ch, color, grad in [
        (8.6, 3.0, 7.5, 7.5, NAVY_LIGHT, False),
        (9.6, 4.0, 5.5, 5.5, NAVY_DARK, False),
        (10.4, 4.8, 3.9, 3.9, None, True),
    ]:
        ellipse = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx), Inches(cy), Inches(cw), Inches(ch))
        if grad:
            _gradient_fill(ellipse, angle=45)
        else:
            ellipse.fill.solid()
            ellipse.fill.fore_color.rgb = color
        ellipse.line.fill.background()
        ellipse.shadow.inherit = False

    _textbox(slide, 0.9, 2.35, 8, 0.4, "부동산 분석 리포트", size=14, bold=True, color=TERRACOTTA)
    _textbox(slide, 0.9, 2.8, 9.5, 1.6, data["address"], size=40, bold=True, color=WHITE)
    _textbox(slide, 0.9, 4.05, 8, 0.5, "건축물 · 실거래가 · 상권 통합 분석", size=16, color=ICE)

    stats = data.get("cover_stats") or []
    for i, s in enumerate(stats[:3]):
        cx = 0.9 + i * 2.9
        _textbox(slide, cx, 5.5, 2.6, 0.55, s["value"], size=24, bold=True, color=TERRACOTTA)
        _textbox(slide, cx, 6.05, 2.6, 0.5, s["label"], size=10.5, color=ICE)

    _textbox(slide, 0.9, 6.7, 6, 0.4, data["report_date"], size=11, color=MUTED)


def _toc_slide(prs, data):
    slide = _new_slide(prs)
    _section_title(slide, "목차")
    _textbox(slide, 0.6, 1.05, 11, 0.3, f"{data['address']} 분석 리포트 구성", size=12, color=MUTED)

    parts = data["toc_parts"]
    col_w, gap_x, start_x, start_y = 3.95, 0.2, 0.6, 1.55
    page_no = 1
    for pi, part in enumerate(parts):
        x = start_x + pi * (col_w + gap_x)
        _textbox(slide, x, start_y, col_w, 0.5, part["title"], size=13.5, bold=True, color=TERRACOTTA)
        iy = start_y + 0.55
        for item in part["items"]:
            badge_w = 0.4 if page_no >= 10 else 0.32
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(iy), Inches(badge_w), Inches(0.32))
            badge.adjustments[0] = 0.2
            badge.fill.solid()
            badge.fill.fore_color.rgb = NAVY
            badge.line.fill.background()
            badge.shadow.inherit = False
            tf = badge.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = str(page_no)
            run.font.size = Pt(9 if page_no >= 10 else 10)
            run.font.bold = True
            run.font.name = FONT_NAME
            run.font.color.rgb = WHITE
            _textbox(slide, x + 0.5, iy - 0.03, col_w - 0.53, 0.38, item, size=11.5, valign=MSO_ANCHOR.MIDDLE)
            iy += 0.44
            page_no += 1
    _page_footer(slide, data["address"], "목차")


def _summary_slide(prs, data):
    slide = _new_slide(prs)
    _section_title(slide, "핵심 요약")
    _card(slide, 0.6, 1.25, 12.1, 1.35)
    _textbox(slide, 0.9, 1.42, 11.5, 1.0, data["summary_text"], size=13, valign=MSO_ANCHOR.MIDDLE, line_spacing=1.25)

    stats = data["summary_stats"]
    cols, gap = 3, 0.25
    card_w = (12.1 - gap * (cols - 1)) / cols
    card_h = 1.35
    start_y = 2.85
    for i, s in enumerate(stats):
        col, row = i % cols, i // cols
        x = 0.6 + col * (card_w + gap)
        y = start_y + row * (card_h + gap)
        _stat_card(slide, x, y, card_w, card_h, s["value"], s["label"], s.get("sub"))
    _page_footer(slide, data["address"], "핵심 요약")


def _building_slide(prs, data):
    slide = _new_slide(prs)
    _section_title(slide, "건축물 개요")
    b = data["building"]

    left_x, left_y, left_w, left_h = 0.6, 1.3, 6.7, 5.6
    _card(slide, left_x, left_y, left_w, left_h)
    core = b["core"]
    row_h = (left_h - 0.4) / max(1, math.ceil(len(core) / 2))
    for i, (label, value) in enumerate(core):
        col, row = i % 2, i // 2
        x = left_x + 0.3 + col * (left_w / 2 - 0.15)
        y = left_y + 0.25 + row * row_h
        _textbox(slide, x, y, left_w / 2 - 0.5, row_h * 0.42, label, size=10.5, color=MUTED)
        _textbox(slide, x, y + row_h * 0.38, left_w / 2 - 0.5, row_h * 0.5, value, size=15, bold=True, color=NAVY)

    right_x, right_w = 7.5, 5.2
    next_y = 1.3
    if b.get("floors"):
        _textbox(slide, right_x, next_y, right_w, 0.35, "층별 면적 (㎡)", size=13, bold=True, color=NAVY)
        chart = _bar_chart(
            slide, right_x, next_y + 0.4, right_w, 2.7,
            [f["label"] for f in b["floors"]], [f["value"] for f in b["floors"]],
            TERRACOTTA, horizontal=True,
        )
        _gradient_fill(chart.plots[0].series[0].format, angle=0)
        next_y += 3.35

    if b.get("zoning"):
        _textbox(slide, right_x, next_y, right_w, 0.35, "용도지역 · 지구", size=13, bold=True, color=NAVY)
        py = next_y + 0.4
        for z in b["zoning"][:5]:
            pw = min(right_w, 0.35 + len(z) * 0.16)
            pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(right_x), Inches(py), Inches(pw), Inches(0.42))
            pill.adjustments[0] = 0.5
            pill.fill.solid()
            pill.fill.fore_color.rgb = ICE
            pill.line.fill.background()
            pill.shadow.inherit = False
            tf = pill.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            run = p.add_run()
            run.text = z
            run.font.size = Pt(10.5)
            run.font.name = FONT_NAME
            run.font.color.rgb = NAVY
            py += 0.55
    _page_footer(slide, data["address"], "건축물 개요")


def _location_slide(prs, data):
    slide = _new_slide(prs)
    _section_title(slide, "위치 및 입지")
    loc = data["location"]

    map_x, map_y, map_w, map_h = 0.6, 1.3, 7.6, 5.6
    _add_picture_placeholder(slide, map_x, map_y, map_w, map_h, image_stream=loc.get("map_image"), idx=90)
    if not loc.get("map_image"):
        _textbox(slide, map_x, map_y + map_h / 2 - 0.35, map_w, 0.7,
                 "지도 이미지를 생성하지 못했습니다.\nPowerPoint에서 이 영역에 이미지를 끌어다 놓으면 바로 채워집니다.",
                 size=12, color=MUTED, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)

    rx, rw = 8.5, 4.2
    _card(slide, rx, 1.3, rw, 1.3)
    _textbox(slide, rx + 0.25, 1.42, rw - 0.5, 0.3, "행정동", size=10.5, color=MUTED)
    _textbox(slide, rx + 0.25, 1.68, rw - 0.5, 0.4, loc.get("adong_name") or "-", size=18, bold=True, color=NAVY)
    _textbox(slide, rx + 0.25, 2.1, rw - 0.5, 0.3, "법정동", size=10.5, color=MUTED)
    _textbox(slide, rx + 0.25, 2.36, rw - 0.5, 0.2, loc.get("ldong_name") or "-", size=13)

    _textbox(slide, rx, 2.85, rw, 0.35, "인근 지하철역", size=13, bold=True, color=NAVY)
    sy = 3.25
    for st_ in loc.get("subway", [])[:3]:
        _card(slide, rx, sy, rw, 0.85, fill=CARD_BG)
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(rx + 0.18), Inches(sy + 0.2), Inches(0.44), Inches(0.44))
        badge.fill.solid()
        badge.fill.fore_color.rgb = RGBColor.from_string(st_.get("color", "1E2A44").lstrip("#"))
        badge.line.color.rgb = WHITE
        badge.line.width = Pt(1.5)
        badge.shadow.inherit = False
        tf = badge.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = st_["badge"]
        run.font.size = Pt(8.5 if len(st_["badge"]) > 1 else 12)
        run.font.bold = True
        run.font.name = FONT_NAME
        run.font.color.rgb = WHITE
        _textbox(slide, rx + 0.78, sy + 0.1, rw - 1.0, 0.35, st_["name"], size=11.5, bold=True)
        _textbox(slide, rx + 0.78, sy + 0.45, rw - 1.0, 0.3, f"{st_['line']} · 도보 약 {st_['dist']}", size=9.5, color=MUTED)
        sy += 1.0
    _page_footer(slide, data["address"], "위치 및 입지")


def _transactions_slide(prs, data):
    tx = data.get("transactions")
    if not tx or not tx.get("rows"):
        return
    slide = _new_slide(prs)
    _section_title(slide, "실거래가 동향")
    _textbox(slide, 0.6, 1.3, 5.6, 0.35, "최근 거래 내역", size=13, bold=True, color=NAVY)
    _table(slide, 0.6, 1.7, 5.6, min(2.6, 0.42 * (len(tx["rows"]) + 1)),
           ["거래일", "층", "전용면적", "거래가격"], tx["rows"], col_ratios=[1.6, 1.1, 1.4, 1.5])

    if tx.get("trend") and len(tx["trend"]) >= 2:
        _textbox(slide, 6.6, 1.3, 6.1, 0.35, "거래가 추이", size=13, bold=True, color=NAVY)
        _line_chart(slide, 6.6, 1.7, 6.1, 2.7, [t["label"] for t in tx["trend"]], [t["value"] for t in tx["trend"]], TERRACOTTA)

    if tx.get("note"):
        _card(slide, 0.6, 4.6, 12.1, 1.3)
        _textbox(slide, 0.9, 4.6, 11.5, 1.3, tx["note"], size=13, valign=MSO_ANCHOR.MIDDLE)
    _page_footer(slide, data["address"], "실거래가 동향")


def _district_slide(prs, data):
    dist = data.get("district")
    if not dist or not dist.get("age_buckets"):
        return
    slide = _new_slide(prs)
    _section_title(slide, f"동단위 시장 통계 — {data['location'].get('adong_name') or ''}")
    _textbox(slide, 0.6, 1.3, 8, 0.35, "준공연도대별 건물 수 분포 (동 전체)", size=13, bold=True, color=NAVY)
    _bar_chart(slide, 0.6, 1.7, 8.0, 4.6,
               [b["label"] for b in dist["age_buckets"]], [b["value"] for b in dist["age_buckets"]], NAVY_LIGHT)

    rx, rw = 9.0, 3.7
    if dist.get("callout"):
        _stat_card(slide, rx, 1.7, rw, 1.6, dist["callout"]["value"], dist["callout"]["label"], dist["callout"].get("sub"))
    if dist.get("note"):
        _card(slide, rx, 3.5, rw, 2.8)
        _textbox(slide, rx + 0.25, 3.7, rw - 0.5, 2.4, dist["note"], size=12, line_spacing=1.3)
    _page_footer(slide, data["address"], "동단위 시장 통계")


def _price_history_slide(prs, data):
    ph = data.get("price_history")
    if not ph or not ph.get("trend"):
        return
    slide = _new_slide(prs)
    _section_title(slide, "공시가격 시계열")
    _textbox(slide, 0.6, 1.3, 7.4, 0.35, "연도별 공시가격 추이", size=13, bold=True, color=NAVY)
    _line_chart(slide, 0.6, 1.7, 7.4, 4.6, [t["label"] for t in ph["trend"]], [t["value"] for t in ph["trend"]], TERRACOTTA)

    if ph.get("rows"):
        _textbox(slide, 8.3, 1.3, 4.4, 0.35, "연도별 변동률", size=13, bold=True, color=NAVY)
        _table(slide, 8.3, 1.7, 4.4, min(2.2, 0.42 * (len(ph["rows"]) + 1)),
               ["연도", "공시가격", "전년대비"], ph["rows"], col_ratios=[1.3, 1.6, 1.5])
    if ph.get("note"):
        _card(slide, 8.3, 4.1, 4.4, 2.2)
        _textbox(slide, 8.55, 4.1, 3.9, 2.2, ph["note"], size=12, valign=MSO_ANCHOR.MIDDLE, line_spacing=1.3)
    _page_footer(slide, data["address"], "공시가격 시계열")


def _commercial_slide(prs, data):
    com = data.get("commercial")
    if not com or (not com.get("vacancy_trend") and not com.get("top_industries")):
        return
    slide = _new_slide(prs)
    _section_title(slide, "상권 개황 — 공실률 및 주변 업종")

    if com.get("vacancy_trend"):
        _textbox(slide, 0.6, 1.3, 6.0, 0.35, f"공실률 추이 ({com.get('vacancy_label', '')})", size=13, bold=True, color=NAVY)
        _line_chart(slide, 0.6, 1.7, 6.0, 4.6,
                    [t["label"] for t in com["vacancy_trend"]], [t["value"] for t in com["vacancy_trend"]], NAVY_LIGHT)
    else:
        _card(slide, 0.6, 1.7, 6.0, 4.6)
        _textbox(slide, 0.6, 3.7, 6.0, 0.6, "해당 지역 공실률 데이터를 찾지 못했습니다.",
                 size=12, color=MUTED, align=PP_ALIGN.CENTER)

    if com.get("top_industries"):
        _textbox(slide, 6.9, 1.3, 5.8, 0.35, "반경 500m 업종 Top 5 (점포 수)", size=13, bold=True, color=NAVY)
        chart = _bar_chart(slide, 6.9, 1.7, 5.8, 4.6,
                            [t["label"] for t in com["top_industries"]], [t["value"] for t in com["top_industries"]],
                            TERRACOTTA, horizontal=True)
        _gradient_fill(chart.plots[0].series[0].format, angle=0)
    _page_footer(slide, data["address"], "상권 개황")


def _trade_area_map_slide(prs, data):
    tam = data.get("trade_area_map")
    if not tam:
        return
    slide = _new_slide(prs)
    _section_title(slide, f"상권영역 지도 — {tam.get('name', '')}")
    map_x, map_y, map_w, map_h = 0.6, 1.3, 12.1, 5.4
    _add_picture_placeholder(slide, map_x, map_y, map_w, map_h, image_stream=tam.get("map_image"), idx=91)
    _textbox(slide, map_x, map_y + map_h + 0.05, map_w, 0.3,
             "※ OpenStreetMap 기반 위치 참고 지도. 상권영역 경계 자체는 서울시 API가 중심좌표+면적만 제공해 "
             "정확한 폴리곤으로 표시할 수 없습니다. PowerPoint에서 이 영역에 이미지를 끌어다 놓으면 바로 교체됩니다.",
             size=8.5, italic=True, color=MUTED)
    _page_footer(slide, data["address"], "상권영역 지도")


def _seoul_detail_slide(prs, data):
    sd = data.get("seoul_trade_area")
    if not sd:
        return
    slide = _new_slide(prs)
    _section_title(slide, f"서울 상권 상세 — {sd.get('name', '')}")

    stats = sd.get("stats") or []
    cols, gap = 3, 0.25
    card_w = (12.1 - gap * (cols - 1)) / cols
    for i, s in enumerate(stats[:3]):
        x = 0.6 + i * (card_w + gap)
        _stat_card(slide, x, 1.3, card_w, 1.2, s["value"], s["label"])

    if sd.get("top_industries"):
        _textbox(slide, 0.6, 2.85, 5.6, 0.35, "업종별 매출 Top 5", size=13, bold=True, color=NAVY)
        _table(slide, 0.6, 3.25, 5.6, 3.0, ["업종", "당월 매출"], sd["top_industries"], col_ratios=[3.6, 2.0])

    if sd.get("weekday"):
        _textbox(slide, 6.6, 2.85, 6.1, 0.35, "요일별 매출 (억원)", size=13, bold=True, color=NAVY)
        chart = _bar_chart(slide, 6.6, 3.25, 6.1, 3.0, [d["label"] for d in sd["weekday"]], [d["value"] for d in sd["weekday"]],
                            TERRACOTTA, num_fmt="0.0")
        _gradient_fill(chart.plots[0].series[0].format, angle=0)
    _page_footer(slide, data["address"], "서울 상권 상세")


def _conclusion_slide(prs, data):
    slide = _new_slide(prs, dark=True)
    ellipse = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2.5), Inches(4.2), Inches(6.5), Inches(6.5))
    ellipse.fill.solid()
    ellipse.fill.fore_color.rgb = NAVY_LIGHT
    ellipse.line.fill.background()
    ellipse.shadow.inherit = False

    _textbox(slide, 0.9, 0.7, 8, 0.7, "종합 의견", size=32, bold=True, color=WHITE)
    cy = 1.9
    for i, point in enumerate(data.get("conclusion") or []):
        badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.9), Inches(cy + 0.05), Inches(0.36), Inches(0.36))
        _gradient_fill(badge, angle=45)
        badge.line.fill.background()
        badge.shadow.inherit = False
        tf = badge.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = str(i + 1)
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.name = FONT_NAME
        run.font.color.rgb = WHITE
        _textbox(slide, 1.5, cy, 10.6, 0.9, point, size=15, color=ICE, line_spacing=1.3)
        cy += 1.15
    _textbox(slide, 0.9, 6.9, 10, 0.4, f"{data['address']}  ·  {data['report_date']}", size=10, color=MUTED)


def generate_pptx(data: dict) -> bytes:
    """fetch_report_data()가 만든 딕셔너리로 pptx를 그려 바이트로 반환."""
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)
    manual_prs = Presentation(_MANUAL_SLIDES_PATH)

    _cover_slide(prs, data)
    _toc_slide(prs, data)
    _insert_manual_slide(prs, manual_prs, 0)   # 리포트 개요 · 핵심지표 요약 (수동 편집)
    _summary_slide(prs, data)
    _building_slide(prs, data)
    _insert_manual_slide(prs, manual_prs, 1)   # PART 1 섹션 구분: 입지 및 상권분석 (수동 편집)
    _location_slide(prs, data)
    _transactions_slide(prs, data)
    _district_slide(prs, data)
    _price_history_slide(prs, data)
    _commercial_slide(prs, data)
    _trade_area_map_slide(prs, data)
    _seoul_detail_slide(prs, data)
    _insert_manual_slide(prs, manual_prs, 2)   # 개발호재 종합 (수동 편집)
    _insert_manual_slide(prs, manual_prs, 3)   # SWOT 분석 (수동 편집)
    _conclusion_slide(prs, data)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
# ==================================================================
# 데이터 계층 — 주소 -> 실제 데이터
# ==================================================================
_SUBWAY_STATIONS_PATH = os.path.join(os.path.dirname(__file__), "address_map_component", "subway_stations.json")


def _load_subway_stations():
    with open(_SUBWAY_STATIONS_PATH, encoding="utf-8") as f:
        return json.load(f)


def _haversine_m(lat1, lon1, lat2, lon2):
    r = 6371000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return r * 2 * math.asin(math.sqrt(a))


def _nearby_subway(lat, lon, limit=3):
    stations = _load_subway_stations()
    scored = sorted(
        ((_haversine_m(lat, lon, s_lat, s_lon), name, lines) for name, s_lat, s_lon, lines in stations),
        key=lambda t: t[0],
    )
    out = []
    for d, name, lines in scored[:limit]:
        line_names = [l[0] for l in lines]
        color = lines[0][1].lstrip("#")
        badge = line_names[0][:2] if len(line_names[0]) > 2 else line_names[0]
        out.append({
            "name": f"{name}역", "line": "·".join(line_names), "color": color,
            "dist": f"{int(round(d / 10) * 10)}m", "badge": badge,
        })
    return out


def _render_location_map(lat, lon, subway=None, radius_m=None):
    """staticmap(OSM 타일, 브라우저 불필요)으로 위치 지도 PNG를 만들어 BytesIO로 반환."""
    try:
        from staticmap import StaticMap, CircleMarker
    except ImportError:
        return None
    try:
        m = StaticMap(1300, 1040, url_template="https://a.tile.openstreetmap.org/{z}/{x}/{y}.png")
        for st_ in (subway or [])[:3]:
            # 지하철 마커는 위치 참고용 점만 (역 좌표는 subway_stations.json 기준, 정밀 좌표 아님)
            pass
        m.add_marker(CircleMarker((lon, lat), "#FFFFFF", 30))
        m.add_marker(CircleMarker((lon, lat), "#C1652F", 22))
        m.add_marker(CircleMarker((lon, lat), "#FFFFFF", 8))
        image = m.render(zoom=16)
        buf = io.BytesIO()
        image.save(buf, format="PNG")
        buf.seek(0)
        return buf
    except Exception:
        return None


def _tx_field(row, candidates):
    for c in candidates:
        if c in row.index and pd.notna(row[c]) and str(row[c]).strip() not in ("", "nan"):
            return row[c]
    return None


def _format_tx_date(row):
    for c in ["계약일", "거래일"]:
        v = _tx_field(row, [c])
        if v:
            return str(v)
    y = _tx_field(row, ["년", "dealYear", "계약년도"])
    m = _tx_field(row, ["월", "dealMonth", "계약월"])
    d = _tx_field(row, ["일", "dealDay", "계약일자"])
    if y and m:
        try:
            return f"{int(y)}.{int(m):02d}" + (f".{int(d):02d}" if d else "")
        except (ValueError, TypeError):
            pass
    return "-"


def _build_transactions(master, months_lookback):
    tx = master.get("실거래가") or {}
    result = {"rows": [], "trend": [], "note": tx.get("note")}
    df = tx.get("df")
    if tx.get("status") != "matched" or df is None or df.empty:
        return result

    try:
        rows = []
        trend = []
        sort_col = None
        for c in ["계약일", "거래일", "년", "dealYear"]:
            if c in df.columns:
                sort_col = c
                break
        sorted_df = df.sort_values(sort_col, ascending=False) if sort_col else df
        for _, row in sorted_df.head(8).iterrows():
            date_label = _format_tx_date(row)
            floor = _tx_field(row, ["층"])
            area = _tx_field(row, ["전용면적", "대지면적", "연면적", "계약면적"])
            price = _tx_field(row, ["거래금액", "물건금액"])
            price_num = pd.to_numeric(str(price).replace(",", ""), errors="coerce") if price is not None else None
            price_label = f"{price_num / 1e4:.1f}억" if pd.notna(price_num) else (str(price) if price else "-")
            rows.append([
                date_label,
                f"{floor}층" if floor not in (None, "") else "-",
                f"{float(area):.1f}㎡" if area not in (None, "") and pd.notna(pd.to_numeric(area, errors="coerce")) else "-",
                price_label,
            ])
            if pd.notna(price_num):
                trend.append({"label": date_label, "value": round(float(price_num) / 1e4, 1)})
        result["rows"] = rows
        result["trend"] = list(reversed(trend))
    except Exception:
        pass
    return result


def _build_district(master):
    stats = master.get("동단위통계")
    old_df = master.get("노후도")
    subject_age = int(old_df.iloc[0]["경과연수"]) if old_df is not None and not old_df.empty else None

    result = {"age_buckets": [], "callout": None, "note": None}
    if not stats or stats.get("연대별") is None or stats["연대별"].empty:
        return result

    by_decade = stats["연대별"]
    result["age_buckets"] = [
        {"label": row["연대"], "value": int(row["건수"])} for _, row in by_decade.iterrows()
    ]
    avg_age = (stats.get("총괄") or {}).get("평균경과연수")
    if avg_age is not None:
        if subject_age is not None:
            cmp_word = "높은" if subject_age > avg_age else "낮은"
            result["note"] = (
                f"동 전체 평균 경과연수는 {avg_age}년이며, 대상 건물({subject_age}년)은 이보다 "
                f"{cmp_word} 편입니다."
            )
        else:
            result["note"] = f"동 전체 평균 경과연수는 {avg_age}년입니다."
        result["callout"] = {"value": f"{avg_age}년", "label": "동 전체 평균 경과연수"}
    return result


def _build_price_history(master):
    ph = master.get("공시가격") or {}
    units = ph.get("단위목록") or []
    if not units:
        return {"trend": [], "rows": [], "note": ph.get("경고")}
    unit = units[0]
    timeline = unit.get("추이")
    trend = []
    rows = []
    if timeline is not None and not timeline.empty:
        prev_price = None
        for _, row in timeline.sort_values("연도").iterrows():
            year, price = int(row["연도"]), float(row["주택가격"])
            trend.append({"label": str(year), "value": round(price / 1e8, 2)})
            change = f"{(price / prev_price - 1) * 100:+.1f}%" if prev_price else "-"
            rows.append([str(year), f"{price / 1e8:.1f}억", change])
            prev_price = price
        rows = list(reversed(rows))[:5]
    note = None
    cagr = unit.get("연평균상승률CAGR(%)")
    if cagr is not None:
        note = f"연평균(CAGR) 약 {cagr:.1f}% {'상승' if cagr >= 0 else '하락'}하는 흐름을 보이고 있습니다."
    return {"trend": trend, "rows": rows, "note": note}


def _build_commercial(reb_key, sangkwon_key, dong_name, lon, lat):
    """공실률/주변 상가업소 조회. 실패 사유를 삼키지 않고 result["notes"]에 남겨서
    리포트에 왜 이 섹션이 비었는지 대시보드에서 그대로 보여줄 수 있게 한다."""
    result = {"vacancy_trend": [], "vacancy_label": "", "top_industries": [], "notes": []}

    if not reb_key:
        result["notes"].append("공실률: 한국부동산원 인증키 미입력")
    else:
        try:
            statbl_id = REB_COMMERCIAL_VACANCY_STATBL_IDS["중대형 상가"]
            quarter = reb_current_quarter_id()
            snap_df, used_quarter, _ = get_reb_vacancy_snapshot(reb_key, statbl_id, quarter)
            if snap_df is None or snap_df.empty:
                result["notes"].append("공실률: 한국부동산원 응답이 비어 있음")
            else:
                keyword = dong_name[:-1] if dong_name and dong_name.endswith("동") else dong_name
                match = snap_df[snap_df["CLS_FULLNM"].astype(str).str.contains(keyword, na=False)] if keyword else pd.DataFrame()
                if match.empty:
                    result["notes"].append(f"공실률: '{keyword}'이(가) 한국부동산원 상권 분류명과 매칭되지 않음")
                else:
                    cls_id = match.iloc[0]["CLS_ID"]
                    result["vacancy_label"] = f"중대형 상가 · {match.iloc[0]['CLS_FULLNM']}"
                    trend_df, _ = get_reb_vacancy_trend(reb_key, statbl_id, cls_id, "202403", used_quarter)
                    if trend_df is None or trend_df.empty:
                        result["notes"].append("공실률: 추이 데이터 없음")
                    else:
                        trend_df = trend_df.sort_values("WRTTIME_IDTFR_ID")
                        for _, row in trend_df.iterrows():
                            wt = str(row["WRTTIME_IDTFR_ID"])
                            result["vacancy_trend"].append({
                                "label": f"{wt[2:4]}.Q{wt[4:]}",
                                "value": round(float(row["DTA_VAL"]), 1),
                            })
        except Exception as e:
            result["notes"].append(f"공실률: 조회 오류 ({e})")

    if not sangkwon_key:
        result["notes"].append("주변 상가업소: 소상공인시장진흥공단 인증키 미입력")
    elif not (lon and lat):
        result["notes"].append("주변 상가업소: 좌표(지오코딩) 실패로 조회 불가")
    else:
        try:
            stores_df = get_nearby_stores(sangkwon_key, lon, lat, radius=500)
            if stores_df is None or stores_df.empty or "indsLclsNm" not in stores_df.columns:
                result["notes"].append("주변 상가업소: 반경 500m 내 데이터 없음")
            else:
                counts = stores_df["indsSclsNm"].value_counts().head(5)
                result["top_industries"] = [{"label": k, "value": int(v)} for k, v in counts.items()]
        except Exception as e:
            result["notes"].append(f"주변 상가업소: 조회 오류 ({e})")
    return result


def _build_seoul(seoul_key, lon, lat):
    """반환값 3번째는 실패 사유(성공 시 None) — 서울 상권분석/상권영역 지도가 왜 빠졌는지
    대시보드에 그대로 보여주기 위함."""
    if not seoul_key:
        return None, None, "서울 열린데이터광장 인증키 미입력"
    if not lon or not lat:
        return None, None, "좌표(지오코딩) 실패로 조회 불가"
    try:
        locations_df = get_seoul_trade_area_locations(seoul_key)
        if locations_df is None or locations_df.empty:
            return None, None, "서울시 상권 목록 응답이 비어 있음"
        trdar_row, dist_m = find_nearest_seoul_trade_area(locations_df, lon, lat)
        if dist_m > 1500:
            return None, None, f"가장 가까운 서울시 상권이 {dist_m:.0f}m 떨어져 있어(1.5km 초과) 매칭하지 않음"
        trdar_cd = trdar_row["TRDAR_CD"]
        name = trdar_row["TRDAR_CD_NM"]

        selng_df, _ = get_seoul_trade_area_quarter_dataset(seoul_key, SEOUL_TRDAR_SALES_SERVICE)
        stor_df, _ = get_seoul_trade_area_quarter_dataset(seoul_key, SEOUL_TRDAR_STORE_SERVICE)
        flpop_df, _ = get_seoul_trade_area_quarter_dataset(seoul_key, SEOUL_TRDAR_FLPOP_SERVICE)
        wrc_df, _ = get_seoul_trade_area_quarter_dataset(seoul_key, SEOUL_TRDAR_WRC_POPLTN_SERVICE)

        sel = selng_df[selng_df["TRDAR_CD"] == trdar_cd].copy() if selng_df is not None and not selng_df.empty else pd.DataFrame()
        flp = flpop_df[flpop_df["TRDAR_CD"] == trdar_cd] if flpop_df is not None and not flpop_df.empty else pd.DataFrame()
        wrc = wrc_df[wrc_df["TRDAR_CD"] == trdar_cd] if wrc_df is not None and not wrc_df.empty else pd.DataFrame()

        stats = []
        if not sel.empty:
            total_sales = pd.to_numeric(sel["THSMON_SELNG_AMT"], errors="coerce").sum()
            stats.append({"label": "추정매출 (당월)", "value": f"{total_sales / 1e8:,.1f}억원"})
        if not flp.empty:
            total_flpop = pd.to_numeric(flp["TOT_FLPOP_CO"], errors="coerce").sum()
            stats.append({"label": "생활인구 (분기)", "value": f"{total_flpop:,.0f}명"})
        if not wrc.empty:
            total_wrc = pd.to_numeric(wrc["TOT_WRC_POPLTN_CO"], errors="coerce").sum()
            stats.append({"label": "직장인구 (분기)", "value": f"{total_wrc:,.0f}명"})

        top_industries = []
        weekday = []
        if not sel.empty:
            sel["THSMON_SELNG_AMT"] = pd.to_numeric(sel["THSMON_SELNG_AMT"], errors="coerce")
            top5 = sel.sort_values("THSMON_SELNG_AMT", ascending=False).head(5)
            top_industries = [[r["SVC_INDUTY_CD_NM"], f"{r['THSMON_SELNG_AMT'] / 1e8:.1f}억"] for _, r in top5.iterrows()]
            day_cols = [
                ("MON_SELNG_AMT", "월"), ("TUES_SELNG_AMT", "화"), ("WED_SELNG_AMT", "수"), ("THUR_SELNG_AMT", "목"),
                ("FRI_SELNG_AMT", "금"), ("SAT_SELNG_AMT", "토"), ("SUN_SELNG_AMT", "일"),
            ]
            weekday = [
                {"label": label, "value": round(pd.to_numeric(sel[col], errors="coerce").sum() / 1e8, 1)}
                for col, label in day_cols if col in sel.columns
            ]

        seoul_detail = {"name": name, "stats": stats, "top_industries": top_industries, "weekday": weekday}
        trade_area_map = {"name": name, "lat": trdar_row["lat"], "lon": trdar_row["lon"]}
        return seoul_detail, trade_area_map, None
    except Exception as e:
        return None, None, f"조회 오류 ({e})"


_SEISMIC_SHORT_LABELS = {
    "내진설계 적용(명시)": "적용",
    "내진설계 미적용(명시)": "미적용",
    "준공년도 미상": "미상",
    "1988년 이전 준공 (의무화 전, 미적용 추정)": "미적용(추정)",
    "1988~2017 준공 (기준 완화기간, 확인 필요)": "부분 적용(추정)",
    "2018년 이후 준공 (전면의무화, 적용 추정)": "적용(추정)",
}


def _short_seismic_label(classification: str) -> str:
    return _SEISMIC_SHORT_LABELS.get(classification, str(classification)[:8])


def _build_conclusion(master, data):
    points = []
    old_df = master.get("노후도")
    if old_df is not None and not old_df.empty:
        age = int(old_df.iloc[0]["경과연수"])
        if age >= 30:
            points.append(f"준공 {age}년 경과한 노후 건물로, 매입 시 리모델링/재건축 비용을 사전 검토할 필요가 있습니다.")

    seismic = (master.get("내진분석") or {}).get("취약우선목록")
    if seismic is not None and not seismic.empty and "미적용" in str(seismic.iloc[0].get("내진분류", "")):
        points.append("내진설계가 적용되지 않은 것으로 분류되어, 구조 보강 여부를 확인할 필요가 있습니다.")

    com = data.get("commercial") or {}
    if com.get("vacancy_trend") and len(com["vacancy_trend"]) >= 2:
        first, last = com["vacancy_trend"][0]["value"], com["vacancy_trend"][-1]["value"]
        trend_word = "안정적으로 유지" if abs(last - first) <= 1.5 else ("개선되는" if last < first else "악화되는")
        points.append(f"최근 공실률이 {first}%→{last}%로 {trend_word} 흐름을 보이고 있습니다.")

    ph = data.get("price_history") or {}
    # ph["note"]는 트렌드가 있을 때만 CAGR 요약 문장이고, 데이터가 없으면 일반 면책조항
    # 문구(공시가격≠시세)가 그대로 들어있어 결론 문장으로 쓰기에 부적절하다 — trend 유무로 구분.
    if ph.get("trend") and ph.get("note"):
        points.append(f"공시가격은 {ph['note']}")

    if not points:
        points.append("자동 분석 근거가 될 데이터가 충분하지 않아, 추가 리서치가 필요합니다.")
    return points


def fetch_report_data(
    *, service_key, sido, sigungu_name, dong_name, sigungu_code, bdong_code, bun, ji,
    kakao_key=None, vworld_key=None, reb_key=None, sangkwon_key=None, seoul_key=None,
    months_lookback=12, progress_callback=None,
) -> dict:
    """주소(시군구/법정동/번지) 하나에 대해 실제 데이터를 모아 generate_pptx()용 딕셔너리로 반환."""
    from PublicDataReader import BuildingLedger, TransactionPrice

    def _progress(msg):
        if progress_callback:
            progress_callback(msg)

    api = BuildingLedger(service_key)
    tp_api = TransactionPrice(service_key)
    address_label = f"{sido} {sigungu_name} {dong_name}" + (f" {bun}" if bun else "") + (f"-{ji}" if ji and str(ji) != "0" else "")

    _progress("동단위 표제부 조회 중...")
    try:
        district_title_df = get_building_ledger(
            api, ledger_type="표제부", sigungu_code=sigungu_code, bdong_code=bdong_code,
            max_rows=10000, wait_time=0.15,
        )
    except Exception:
        district_title_df = None

    _progress("건축물대장·실거래가·공시가격 조회 중...")
    master = build_master_report(
        api, tp_api, sigungu_code, bdong_code, bun, ji,
        months_lookback=months_lookback, district_title_df=district_title_df,
        sido=sido, sigungu_name=sigungu_name, vworld_key=vworld_key, kakao_key=kakao_key,
    )

    data = {
        "address": address_label,
        "report_date": datetime.date.today().strftime("%Y년 %m월 %d일 기준"),
    }

    title_df = master.get("표제부")
    core_row = title_df.iloc[0] if title_df is not None and not title_df.empty else None

    # ---- 건축물 개요 ----
    core_fields = [
        ("대지면적", "대지면적", "㎡"), ("건축면적", "건축면적", "㎡"), ("연면적", "연면적", "㎡"),
        ("건폐율", "건폐율", "%"), ("용적률", "용적률", "%"), ("구조코드명", "구조", ""),
        ("지붕코드명", "지붕", ""), ("사용승인일", "사용승인일", ""), ("허가일", "허가일", ""), ("착공일", "착공일", ""),
    ]
    core_list = []
    if core_row is not None:
        for col, label, unit in core_fields:
            val = core_row.get(col)
            if val not in (None, "", "nan") and str(val).strip():
                core_list.append((label, f"{val}{unit}"))
        floors_top = core_row.get("지상층수")
        floors_base = core_row.get("지하층수")
        if floors_top or floors_base:
            core_list.append(("지상/지하층수", f"{floors_top or 0}층 / {floors_base or 0}층"))

    zoning = combine_zoning_sources(master)
    data["building"] = {"core": core_list, "floors": [], "zoning": zoning}

    # ---- 위치 및 입지 ----
    coord = master.get("좌표")
    lon, lat = (coord if coord else (None, None))
    subway = _nearby_subway(lat, lon) if lat and lon else []
    map_buf = _render_location_map(lat, lon, subway) if lat and lon else None
    data["location"] = {
        "adong_name": dong_name, "ldong_name": dong_name,
        "subway": subway, "map_image": map_buf,
    }
    data["lon"], data["lat"] = lon, lat

    # ---- 핵심 요약 ----
    data["summary_text"] = build_executive_summary(master)
    summary_stats = []
    old_df = master.get("노후도")
    if core_row is not None:
        approval_year = extract_year(core_row.get("사용승인일"))
        age = int(old_df.iloc[0]["경과연수"]) if old_df is not None and not old_df.empty else None
        if approval_year:
            sub = f"{age}년 경과" + (" · 노후" if age and age >= 30 else "") if age is not None else None
            summary_stats.append({"label": "사용승인일 (경과연수)", "value": f"{approval_year}년", "sub": sub})
    seismic_list = (master.get("내진분석") or {}).get("취약우선목록")
    if seismic_list is not None and not seismic_list.empty:
        seismic_full = str(seismic_list.iloc[0]["내진분류"])
        summary_stats.append({"label": "내진 설계", "value": _short_seismic_label(seismic_full), "sub": seismic_full})

    # 어떤 섹션이 왜 빠졌는지(키 미입력/데이터 없음/API 오류) 대시보드에 그대로 보여주기 위한 메모.
    notes = []
    if not lon or not lat:
        notes.append("위치·좌표: 지오코딩 실패 — 카카오맵/브이월드 키를 확인하세요 (지도·주변 상가업소·서울 상권분석에 영향)")

    _progress("동단위 시장 통계 분석 중...")
    data["district"] = _build_district(master)
    if not data["district"]["age_buckets"]:
        notes.append("동단위 시장 통계: 동 표제부 데이터를 가져오지 못했거나 비어 있음")

    _progress("실거래가 정리 중...")
    data["transactions"] = _build_transactions(master, months_lookback)
    if data["transactions"]["rows"]:
        latest = data["transactions"]["rows"][0]
        summary_stats.append({"label": "최근 실거래가", "value": latest[3], "sub": f"{latest[1]} · {latest[2]} · {latest[0]}"})
    else:
        notes.append(f"실거래가 동향: 최근 {months_lookback}개월 내 거래 내역 없음")

    _progress("공시가격 시계열 분석 중...")
    data["price_history"] = _build_price_history(master)
    if not data["price_history"]["trend"]:
        notes.append("공시가격 시계열: 시계열 데이터 없음")

    _progress("상업용부동산 공실률 · 주변 상가업소 조회 중...")
    data["commercial"] = _build_commercial(reb_key, sangkwon_key, dong_name, lon, lat)
    notes.extend(data["commercial"].get("notes") or [])
    if data["commercial"]["vacancy_trend"]:
        summary_stats.append({"label": "공실률", "value": f"{data['commercial']['vacancy_trend'][-1]['value']}%", "sub": data["commercial"]["vacancy_label"]})

    is_seoul = sido.startswith("서울")
    seoul_detail, trade_area_map = None, None
    if not is_seoul:
        notes.append("서울 상권분석·상권영역 지도: 서울 열린데이터광장 API는 서울 소재 주소만 지원 (해당 없음)")
    elif not seoul_key:
        notes.append("서울 상권분석·상권영역 지도: 서울 열린데이터광장 인증키 미입력")
    else:
        _progress("서울 상권분석 데이터 조회 중... (최초 1회, 최대 1분 정도 걸릴 수 있습니다)")
        seoul_detail, trade_area_map, seoul_reason = _build_seoul(seoul_key, lon, lat)
        if seoul_reason:
            notes.append(f"서울 상권분석·상권영역 지도: {seoul_reason}")
    data["seoul_trade_area"] = seoul_detail
    if seoul_detail and seoul_detail.get("stats"):
        summary_stats.append({"label": "상권 추정매출(월)", "value": seoul_detail["stats"][0]["value"], "sub": seoul_detail["name"]})

    if trade_area_map:
        map_buf2 = _render_location_map(trade_area_map["lat"], trade_area_map["lon"])
        data["trade_area_map"] = {"name": trade_area_map["name"], "map_image": map_buf2}
    else:
        data["trade_area_map"] = None

    data["summary_stats"] = summary_stats
    data["cover_stats"] = summary_stats[:3]
    data["notes"] = notes

    toc_items_part1 = ["표지", "목차", "리포트 개요", "핵심 요약", "건축물 개요", "입지·상권 개관", "위치 및 입지"]
    toc_items_part2 = ["실거래가 동향", "동단위 시장 통계", "공시가격 시계열", "상권 개황"]
    if data["trade_area_map"]:
        toc_items_part2.append("상권영역 지도")
    if seoul_detail:
        toc_items_part2.append("서울 상권 상세")
    data["toc_parts"] = [
        {"title": "PART 01 · 입지 및 건축물 분석", "items": toc_items_part1},
        {"title": "PART 02 · 시장 데이터 분석", "items": toc_items_part2},
        {"title": "PART 03 · 종합 평가", "items": ["개발호재 종합", "SWOT 분석", "종합 의견"]},
    ]

    _progress("종합 의견 정리 중...")
    data["conclusion"] = _build_conclusion(master, data)

    return data
