"""
Dynamic text-fitting helpers.

Estimate whether a given string will fit inside a text box of (width, height)
inches at a specific font size, and pick the largest font size from a range
that keeps the text inside the box.

The character-width model is a rough average for proportional sans/serif
fonts — individual glyph widths vary, so results are approximate but good
enough for layout sizing. A safety factor is applied so typical text fits
without clipping even with wider glyphs.
"""

from __future__ import annotations

# Average advance width of one character as a fraction of the font size (pt).
# ~0.55 is a reasonable average for sans-serif body text; serif displays a
# touch wider. We use a per-font multiplier below for the few named fonts
# we care about.
_BASE_ADVANCE = 0.52

# Line-height multiplier relative to font size
_LINE_HEIGHT = 1.20

# Fonts that render noticeably wider than average sans-serif
_WIDE_FONTS = {
    "libre baskerville": 0.60,
    "oranienbaum": 0.48,
    "cambria": 0.54,
    "georgia": 0.56,
    "times new roman": 0.54,
}


def _avg_char_width_pt(font_name: str | None, font_size_pt: float) -> float:
    """Return average advance width (in points) for one character."""
    advance_ratio = _BASE_ADVANCE
    if font_name:
        advance_ratio = _WIDE_FONTS.get(font_name.strip().lower(), _BASE_ADVANCE)
    return font_size_pt * advance_ratio


def chars_per_line(width_in: float, font_size_pt: float,
                    font_name: str | None = None,
                    safety: float = 0.96) -> int:
    """Estimate how many characters fit on one line."""
    if width_in <= 0 or font_size_pt <= 0:
        return 1
    width_pt = width_in * 72.0
    advance_pt = _avg_char_width_pt(font_name, font_size_pt)
    return max(1, int(width_pt * safety / advance_pt))


def lines_needed(text: str, width_in: float, font_size_pt: float,
                  font_name: str | None = None) -> int:
    """Estimate number of wrapped lines for text at the given width/size.

    Handles explicit line breaks and word-wraps each segment, so we do not
    undercount when the source string contains newlines.
    """
    if not text:
        return 1
    cpl = chars_per_line(width_in, font_size_pt, font_name)
    total = 0
    for segment in text.split("\n"):
        if not segment:
            total += 1
            continue
        # Word-aware wrap: count words, tracking line length.
        words = segment.split()
        if not words:
            total += 1
            continue
        line_len = 0
        lines = 1
        for word in words:
            w = len(word)
            extra = w + (1 if line_len > 0 else 0)  # leading space if not start
            if line_len + extra > cpl and line_len > 0:
                lines += 1
                line_len = w
            else:
                line_len += extra
        total += lines
    return max(1, total)


def text_height_in(text: str, width_in: float, font_size_pt: float,
                    font_name: str | None = None) -> float:
    """Estimated rendered height (inches) for text at a given font size."""
    lh_pt = font_size_pt * _LINE_HEIGHT
    return lines_needed(text, width_in, font_size_pt, font_name) * lh_pt / 72.0


def fit_font_size(text: str, width_in: float, height_in: float,
                   max_pt: float = 16.0, min_pt: float = 9.0,
                   font_name: str | None = None,
                   step: float = 0.5) -> float:
    """Largest font size in [min_pt, max_pt] such that text fits in box.

    Returns min_pt when nothing fits; caller can then enable shrink-to-fit
    on the text frame to keep overflow from clipping completely.
    """
    if not text or width_in <= 0 or height_in <= 0:
        return max_pt

    size = max_pt
    while size >= min_pt:
        h = text_height_in(text, width_in, size, font_name)
        if h <= height_in * 0.98:
            return size
        size -= step
    return min_pt


def fit_multi_line_font_size(lines: list[str], width_in: float, height_in: float,
                              max_pt: float = 16.0, min_pt: float = 9.0,
                              font_name: str | None = None,
                              line_gap_pt: float = 4.0,
                              step: float = 0.5) -> float:
    """Pick font size so a list of bullet lines fit in a single box."""
    if not lines:
        return max_pt
    size = max_pt
    while size >= min_pt:
        lh_in = (size * _LINE_HEIGHT + line_gap_pt) / 72.0
        total = 0.0
        for ln in lines:
            wrapped = lines_needed(ln, width_in, size, font_name)
            total += wrapped * lh_in
        if total <= height_in * 0.98:
            return size
        size -= step
    return min_pt
