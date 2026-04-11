"""
Grid system for consistent slide layout positioning.

Based on Slide Master analysis:
  - Slide: 13.333" x 7.5" (16:9)
  - 12-column grid with 0.20" gutters
  - Margins: 0.375" left/right
  - Title area: y=0.656", h=0.58"
  - Content area starts at y=1.40"
"""

from __future__ import annotations

from dataclasses import dataclass

from pptx.util import Inches, Pt, Emu


@dataclass(frozen=True)
class GridConstants:
    """All grid measurements in inches."""
    # Slide dimensions
    SLIDE_WIDTH: float = 13.333
    SLIDE_HEIGHT: float = 7.5

    # Margins
    MARGIN_LEFT: float = 0.375
    MARGIN_RIGHT: float = 0.375
    MARGIN_TOP: float = 1.40    # Content starts below title
    MARGIN_BOTTOM: float = 0.40

    # Title area
    TITLE_LEFT: float = 0.375
    TITLE_TOP: float = 0.656
    TITLE_WIDTH: float = 11.646
    TITLE_HEIGHT: float = 0.58

    # Footer
    FOOTER_TOP: float = 7.229
    FOOTER_HEIGHT: float = 0.135
    SLIDE_NUMBER_LEFT: float = 9.938
    SLIDE_NUMBER_WIDTH: float = 3.0

    # Grid
    COLUMNS: int = 12
    GUTTER: float = 0.20

    @property
    def content_left(self) -> float:
        return self.MARGIN_LEFT

    @property
    def content_top(self) -> float:
        return self.MARGIN_TOP

    @property
    def content_width(self) -> float:
        return self.SLIDE_WIDTH - self.MARGIN_LEFT - self.MARGIN_RIGHT

    @property
    def content_height(self) -> float:
        return self.SLIDE_HEIGHT - self.MARGIN_TOP - self.MARGIN_BOTTOM

    @property
    def column_width(self) -> float:
        """Width of a single grid column (excluding gutters)."""
        total_gutters = (self.COLUMNS - 1) * self.GUTTER
        return (self.content_width - total_gutters) / self.COLUMNS


# Singleton instance
GRID = GridConstants()


class GridSystem:
    """Compute positions for elements on the slide grid."""

    def __init__(self, constants: GridConstants | None = None):
        self.g = constants or GRID

    # ── Column span helpers ──────────────────────

    def span_left(self, start_col: int) -> float:
        """X position (inches) for the left edge of a column (0-indexed)."""
        return self.g.content_left + start_col * (self.g.column_width + self.g.GUTTER)

    def span_width(self, num_cols: int) -> float:
        """Width (inches) for a span of N columns including internal gutters."""
        if num_cols <= 0:
            return 0.0
        return num_cols * self.g.column_width + (num_cols - 1) * self.g.GUTTER

    def span(self, start_col: int, num_cols: int) -> tuple[float, float]:
        """Return (left, width) in inches for a column span."""
        return self.span_left(start_col), self.span_width(num_cols)

    # ── Common layout presets ────────────────────

    def full_width(self) -> tuple[float, float]:
        """Full content width: (left, width)."""
        return self.span(0, 12)

    def half_left(self) -> tuple[float, float]:
        """Left half: (left, width)."""
        return self.span(0, 6)

    def half_right(self) -> tuple[float, float]:
        """Right half: (left, width)."""
        return self.span(6, 6)

    def third(self, index: int) -> tuple[float, float]:
        """One of three equal columns (index 0, 1, or 2)."""
        return self.span(index * 4, 4)

    def quarter(self, index: int) -> tuple[float, float]:
        """One of four equal columns (index 0, 1, 2, or 3)."""
        return self.span(index * 3, 3)

    # ── Vertical helpers ─────────────────────────

    def content_top(self, row_offset: float = 0.0) -> float:
        """Y position in content area with optional offset (inches)."""
        return self.g.content_top + row_offset

    def content_height(self) -> float:
        """Available content height (inches)."""
        return self.g.content_height

    # ── Title area ───────────────────────────────

    def title_position(self) -> tuple[float, float, float, float]:
        """Return (left, top, width, height) for title text box."""
        return (self.g.TITLE_LEFT, self.g.TITLE_TOP,
                self.g.TITLE_WIDTH, self.g.TITLE_HEIGHT)

    # ── EMU converters for python-pptx ───────────

    @staticmethod
    def inches(value: float) -> int:
        """Convert inches to EMU for python-pptx."""
        return int(Inches(value))

    @staticmethod
    def points(value: float) -> int:
        """Convert points to EMU for python-pptx."""
        return int(Pt(value))

    def span_emu(self, start_col: int, num_cols: int) -> tuple[int, int]:
        """Return (left_emu, width_emu) for a column span."""
        left, width = self.span(start_col, num_cols)
        return self.inches(left), self.inches(width)

    # ── Centering helper ─────────────────────────

    def center_horizontally(self, element_width: float) -> float:
        """X position to center an element of given width in the content area."""
        return self.g.content_left + (self.g.content_width - element_width) / 2

    def center_in_span(self, start_col: int, num_cols: int, element_width: float) -> float:
        """X position to center an element within a column span."""
        left, span_w = self.span(start_col, num_cols)
        return left + (span_w - element_width) / 2
