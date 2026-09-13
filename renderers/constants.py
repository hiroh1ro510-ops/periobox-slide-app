"""Palette, fonts, and standard shape sizes derived from samples/_measurements.json.

All sizes are in inches (matches ``python-pptx`` ``Inches(...)`` units) unless noted.

The sample decks are hand-crafted so 1:1 reproduction is not possible; these
constants capture the *stylized* PerioBox look — pentagon heading pills, pastel
rounded-rectangle content boxes, and directional arrows — with the dominant
palette measured across all 10 samples.
"""

from pptx.dml.color import RGBColor


# --- Slide canvas ---
SLIDE_WIDTH_IN = 10.0
SLIDE_HEIGHT_IN = 7.5


# --- Fonts ---
FONT_LATIN = "Century Gothic"
FONT_JP = "MS PGothic"

FS_HEADING = 28  # e.g. "研究デザイン" pentagon pill (case02/04 top uses 24–28)
FS_SUBHEADING = 20  # subtitle pill like "多施設・後ろ向き研究"
FS_SECTION = 18  # small section labels ("評価項目")
FS_BODY = 16  # dominant body text
FS_SMALL = 14
FS_MICRO = 12  # labels inside diagrams


# --- Palette (measured hex frequencies across slides 1–3 of all 10 decks) ---
# text
TEXT_MAIN = RGBColor(0x30, 0x30, 0x30)  # 70 hits
TEXT_DARK = RGBColor(0x25, 0x25, 0x25)
TEXT_BLACK = RGBColor(0x00, 0x00, 0x00)
TEXT_RED = RGBColor(0xC0, 0x00, 0x00)  # emphasis
TEXT_GREEN = RGBColor(0x00, 0xB0, 0x50)
TEXT_YELLOW = RGBColor(0xFF, 0xFF, 0x00)
TEXT_MUTED = RGBColor(0x73, 0x73, 0x73)

# fills
FILL_PENTAGON_PINK = RGBColor(0xD5, 0xC6, 0xCC)  # 15 hits — dominant heading pill
FILL_GREEN = RGBColor(0x00, 0xB0, 0x50)
FILL_YELLOW = RGBColor(0xFF, 0xFF, 0x00)
FILL_LIGHT_GREEN = RGBColor(0xC4, 0xDE, 0xB4)
FILL_LIGHT_BLUE = RGBColor(0xDD, 0xEB, 0xF7)  # secondary content
FILL_LIGHT_PEACH = RGBColor(0xF4, 0xE0, 0xD3)
FILL_LIGHT_CREAM = RGBColor(0xFC, 0xF0, 0xC9)
FILL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# lines / outlines
LINE_RED = RGBColor(0xC0, 0x00, 0x00)
LINE_GREEN = RGBColor(0x00, 0xB0, 0x50)
LINE_YELLOW = RGBColor(0xFF, 0xFF, 0x00)
LINE_ROSE = RGBColor(0xC6, 0x34, 0x5E)
LINE_MUTED = RGBColor(0xAA, 0xAA, 0xAA)


# --- Heading pentagon pill (top of slides 2 and 3) ---
# Case02/04/09 use a small centered pentagon around x=3.9–4.5, y=0.28–0.45.
PENTAGON_TOP = 0.35
PENTAGON_WIDTH = 2.2
PENTAGON_HEIGHT = 0.65
PENTAGON_LEFT = (SLIDE_WIDTH_IN - PENTAGON_WIDTH) / 2  # centered

# Subtitle pill directly under pentagon
SUBTITLE_TOP = PENTAGON_TOP + PENTAGON_HEIGHT + 0.18
SUBTITLE_WIDTH = 6.2
SUBTITLE_HEIGHT = 0.5
SUBTITLE_LEFT = (SLIDE_WIDTH_IN - SUBTITLE_WIDTH) / 2


# --- Standard content boxes ---
CONTENT_MARGIN_X = 0.55
CONTENT_WIDTH = SLIDE_WIDTH_IN - 2 * CONTENT_MARGIN_X  # ~8.9

# Top text block (subjects / methods summary)
TOP_BOX_TOP = 2.15
TOP_BOX_HEIGHT = 1.2

# Group boxes row (comparison groups)
GROUP_ROW_TOP = 3.55
GROUP_BOX_HEIGHT = 0.9
GROUP_ROW_GAP = 0.28

# Evaluation items pill row
EVAL_LABEL_TOP = 4.85
EVAL_ROW_TOP = 5.25
EVAL_PILL_HEIGHT = 0.55
EVAL_PILL_GAP = 0.22

# "さらに詳しく" callout band (results slide)
DETAIL_BAND_TOP = 6.35
DETAIL_BAND_HEIGHT = 0.55
DETAIL_BAND_WIDTH = 1.5


# --- Arrow standard sizes ---
ARROW_WIDTH = 0.35
ARROW_HEIGHT = 0.4


# --- "さらに詳しく" pill label text ---
DETAIL_LABEL_TEXT = "さらに詳しく"
DESIGN_HEADING_TEXT = "研究デザイン"
RESULT_HEADING_TEXT = "結 果"
