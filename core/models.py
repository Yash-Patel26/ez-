"""
Pydantic data models for the MD2PPTX pipeline.

These models define the contracts between agents:
  DocumentIR: Parser output → Strategist/Optimizer input
  SlidePlan: Strategist output → Optimizer/Layout input
  OptimizedSlideContent: Optimizer output → Layout/Renderer input
  SlideLayout: Layout output → Renderer input
"""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


# ──────────────────────────────────────────────
# Stage 1: Parser Output (Document IR)
# ──────────────────────────────────────────────

class ContentBlock(BaseModel):
    """A single content element within a section."""
    type: Literal[
        "text", "bullet_list", "numbered_list", "table",
        "blockquote", "code_block", "image_ref"
    ]
    content: str | None = None
    items: list[str] | None = None
    headers: list[str] | None = None
    rows: list[list[str]] | None = None
    is_visualizable: bool = False
    has_numeric_data: bool = False


class Section(BaseModel):
    """A document section with heading, content, and optional subsections."""
    heading: str
    level: int = Field(ge=1, le=6)
    content_blocks: list[ContentBlock] = []
    subsections: list[Section] = []
    word_count: int = 0
    has_numeric_data: bool = False
    has_table_data: bool = False


class DocumentIR(BaseModel):
    """Intermediate Representation of the entire parsed markdown document."""
    title: str
    subtitle: str | None = None
    sections: list[Section] = []
    total_word_count: int = 0
    total_tables: int = 0
    total_lists: int = 0
    citations: list[str] = []


# ──────────────────────────────────────────────
# Stage 2: Strategist Output (Slide Plan)
# ──────────────────────────────────────────────

SLIDE_TYPES = Literal[
    "cover", "agenda", "executive_summary", "content",
    "data_chart", "data_table", "comparison", "timeline",
    "process_flow", "kpi_callout", "divider", "conclusion", "thank_you"
]

VISUAL_TREATMENTS = Literal[
    "bullets", "two_column", "three_column",
    "chart_bar", "chart_pie", "chart_line", "chart_area", "chart_stacked_bar",
    "table", "process_flow", "timeline", "kpi_cards",
    "comparison_cards", "icon_grid", "cover_layout", "closing_layout"
]


class SlideSpec(BaseModel):
    """Specification for a single slide in the presentation plan."""
    slide_number: int
    slide_type: SLIDE_TYPES
    title: str
    key_message: str = ""
    source_sections: list[int] = Field(
        default_factory=list,
        description="Indices of IR sections that feed this slide"
    )
    visual_treatment: VISUAL_TREATMENTS = "bullets"
    content_priority: Literal["high", "medium", "low"] = "high"


class SlidePlan(BaseModel):
    """Complete slide plan for the presentation."""
    slides: list[SlideSpec]
    storyline_summary: str = ""
    sections_merged: dict[str, list[int]] = Field(default_factory=dict)


# ──────────────────────────────────────────────
# Stage 3: Content Optimizer Output
# ──────────────────────────────────────────────

class SeriesData(BaseModel):
    """A single data series for a chart."""
    name: str
    values: list[float]


class ChartData(BaseModel):
    """Data for chart generation."""
    chart_type: Literal["bar", "pie", "line", "area", "stacked_bar"]
    title: str = ""
    categories: list[str]
    series: list[SeriesData]


class TableData(BaseModel):
    """Data for table generation."""
    title: str = ""
    headers: list[str]
    rows: list[list[str]]
    highlight_row: int | None = None


class KPIItem(BaseModel):
    """A single KPI metric card."""
    value: str
    label: str
    trend: Literal["up", "down", "neutral"] | None = None
    description: str = ""


class TimelineItem(BaseModel):
    """A single item on a timeline."""
    date: str
    label: str
    description: str = ""


class ComparisonItem(BaseModel):
    """A single item in a comparison view."""
    title: str
    points: list[str] = []
    highlight: bool = False


class ProcessStep(BaseModel):
    """A single step in a process flow."""
    label: str
    description: str = ""


class OptimizedSlideContent(BaseModel):
    """Fully optimized content ready for rendering on a single slide."""
    slide_number: int
    slide_type: SLIDE_TYPES
    title: str
    subtitle: str | None = None
    key_message: str = ""
    visual_treatment: VISUAL_TREATMENTS = "bullets"

    # Content variants (only one or two populated per slide)
    bullets: list[str] | None = None
    left_column: list[str] | None = None
    right_column: list[str] | None = None
    chart_data: ChartData | None = None
    table_data: TableData | None = None
    kpi_values: list[KPIItem] | None = None
    process_steps: list[ProcessStep] | None = None
    timeline_items: list[TimelineItem] | None = None
    comparison_items: list[ComparisonItem] | None = None
    agenda_items: list[str] | None = None


# ──────────────────────────────────────────────
# Stage 4: Layout Engine Output
# ──────────────────────────────────────────────

class Position(BaseModel):
    """Position and size in inches."""
    left: float
    top: float
    width: float
    height: float


class ShapeSpec(BaseModel):
    """Specification for a shape to be rendered on a slide."""
    shape_type: Literal[
        "text_box", "rounded_rect", "rectangle", "oval",
        "line", "arrow", "triangle", "chart", "table"
    ]
    position: Position
    text: str = ""
    font_size: float | None = None
    font_bold: bool = False
    font_color: str | None = None       # hex color or theme reference
    fill_color: str | None = None       # hex color or theme reference
    border_color: str | None = None
    border_width: float | None = None
    alignment: Literal["left", "center", "right"] = "left"
    vertical_alignment: Literal["top", "middle", "bottom"] = "top"


class SlideLayout(BaseModel):
    """Complete layout specification for a single slide."""
    slide_number: int
    layout_name: str                     # Name from Slide Master
    title_text: str = ""
    shapes: list[ShapeSpec] = []
    content: OptimizedSlideContent | None = None


# ──────────────────────────────────────────────
# Theme Configuration
# ──────────────────────────────────────────────

class ThemeColors(BaseModel):
    """Color palette extracted from Slide Master theme."""
    dk1: str = "#000000"
    lt1: str = "#FFFFFF"
    dk2: str = "#2C2C2C"
    lt2: str = "#E8E8E8"
    accent1: str = "#EF4444"
    accent2: str = "#E97132"
    accent3: str = "#196B24"
    accent4: str = "#0F9ED5"
    accent5: str = "#A02B93"
    accent6: str = "#4EA72E"
    hlink: str = "#467886"
    folHlink: str = "#96607D"

    def accent_list(self) -> list[str]:
        """Return accent colors as a list for chart series coloring."""
        return [self.accent1, self.accent2, self.accent3,
                self.accent4, self.accent5, self.accent6]


class ThemeFonts(BaseModel):
    """Font families from Slide Master theme."""
    major: str = "Calibri"    # +mj-lt heading font
    minor: str = "Calibri"    # +mn-lt body font


class ThemeConfig(BaseModel):
    """Complete theme configuration extracted from a Slide Master."""
    colors: ThemeColors = Field(default_factory=ThemeColors)
    fonts: ThemeFonts = Field(default_factory=ThemeFonts)
    slide_width: float = 13.333
    slide_height: float = 7.5
