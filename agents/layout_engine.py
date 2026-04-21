"""
Agent 4: Layout Engine — Visual Richness Edition

Target: 15+ shapes per slide average (sample has 21.4)
Key additions vs previous version:
  - Decorative lines between content sections (sample uses 39 total)
  - Accent shapes as visual anchors on every slide
  - Richer KPI cards (accent bar top + large value + label + description line)
  - Bullet slides with numbered accent circles + separator lines
  - Three-column with top accent bars + accent dots + dividers
  - Footer accent line on every content slide
"""

from __future__ import annotations

from core.grid import GridSystem, GRID
from core.models import (
    OptimizedSlideContent,
    Position,
    ShapeSpec,
    SlideLayout,
    ThemeConfig,
)
from core.text_fit import fit_font_size, fit_multi_line_font_size

PAD = 0.15
GAP = 0.20

# Content area bottom y-coordinate (just above the footer line)
CONTENT_BOTTOM = GRID.SLIDE_HEIGHT - GRID.MARGIN_BOTTOM - 0.05   # ≈ 7.05"

# ═══════════════════════════════════════════════
# Domain-aware Unicode icon mapping
# Maps keywords found in slide titles/content to visual Unicode symbols.
# These render natively in PowerPoint without external image assets.
# ═══════════════════════════════════════════════
DOMAIN_ICONS: dict[str, str] = {
    # Technology & AI
    "ai": "⚡", "artificial intelligence": "⚡", "machine learning": "⚡",
    "genai": "⚡", "agentic": "⚡", "neural": "⚡",
    "technology": "⚙", "tech": "⚙", "digital": "⚙", "software": "⚙",
    "cloud": "☁", "infrastructure": "☁", "saas": "☁",
    "data": "📊", "analytics": "📊", "metrics": "📊", "dashboard": "📊",
    "automation": "⟳", "pipeline": "⟳", "workflow": "⟳",
    # Finance & Business
    "revenue": "💰", "investment": "💰", "funding": "💰", "capital": "💰",
    "market": "📈", "growth": "📈", "trend": "📈", "forecast": "📈",
    "financial": "💲", "cost": "💲", "pricing": "💲", "budget": "💲",
    "strategy": "🎯", "strategic": "🎯", "objective": "🎯", "goal": "🎯",
    "acquisition": "🤝", "merger": "🤝", "partnership": "🤝",
    # Security
    "security": "🛡", "cyber": "🛡", "threat": "🛡", "protection": "🛡",
    "identity": "🔐", "authentication": "🔐", "access": "🔐",
    # Industry & Operations
    "manufacturing": "🏭", "factory": "🏭", "production": "🏭",
    "supply chain": "🔗", "logistics": "🔗", "distribution": "🔗",
    "energy": "⚡", "solar": "☀", "renewable": "♻", "green": "♻",
    "hydrogen": "⚛", "battery": "🔋",
    "healthcare": "🏥", "medical": "🏥", "clinical": "🏥", "pharma": "💊",
    "telecom": "📡", "communication": "📡", "media": "📡",
    # People & Learning
    "talent": "👥", "workforce": "👥", "team": "👥", "employee": "👥",
    "learning": "📚", "education": "📚", "training": "📚", "reskilling": "📚",
    # Geography
    "global": "🌍", "region": "🌍", "geographic": "🌍", "international": "🌍",
    "asia": "🌏", "europe": "🌍", "africa": "🌍", "america": "🌎",
    # General
    "recommendation": "✦", "conclusion": "✦", "takeaway": "✦", "key": "✦",
    "challenge": "⚠", "risk": "⚠", "barrier": "⚠",
    "opportunity": "★", "innovation": "★", "advantage": "★",
    "integration": "⬡", "platform": "⬡", "ecosystem": "⬡",
    "agriculture": "🌾", "food": "🌾", "crop": "🌾",
    "automobile": "🚗", "vehicle": "🚗", "automotive": "🚗", "ev": "🚗",
    "pet": "🐾", "animal": "🐾",
}

# Fallback icon when no domain keyword matches
_DEFAULT_ICON = "◆"


def _pick_icon(text: str) -> str:
    """Pick the best Unicode icon for a piece of text by scanning for domain keywords."""
    if not text:
        return _DEFAULT_ICON
    text_lower = text.lower()
    # Try multi-word keys first (longer match = more specific)
    for keyword in sorted(DOMAIN_ICONS, key=len, reverse=True):
        if keyword in text_lower:
            return DOMAIN_ICONS[keyword]
    return _DEFAULT_ICON


class LayoutEngine:
    def __init__(self, theme: ThemeConfig, config: dict | None = None):
        self.theme = theme
        self.config = config or {}
        self.grid = GridSystem()

    def compute(self, slides: list[OptimizedSlideContent]) -> list[SlideLayout]:
        return [self._compute_slide(s) for s in slides]

    def _compute_slide(self, slide: OptimizedSlideContent) -> SlideLayout:
        layout_name = self._select_layout(slide)
        shapes = self._compute_shapes(slide)
        return SlideLayout(
            slide_number=slide.slide_number,
            layout_name=layout_name,
            title_text=slide.title or "",
            shapes=shapes,
            content=slide,
        )

    def _select_layout(self, slide: OptimizedSlideContent) -> str:
        if slide.slide_type == "cover":
            return "cover"
        if slide.slide_type == "thank_you":
            return "thank_you"
        if slide.slide_type == "divider":
            return "divider"
        return "title_only"

    def _compute_shapes(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        vt = slide.visual_treatment
        dispatch = {
            "cover_layout": self._shapes_cover,
            "closing_layout": lambda s: [],
            "bullets": self._shapes_bullets,
            "chart_bar": self._shapes_chart,
            "chart_pie": self._shapes_chart,
            "chart_line": self._shapes_chart,
            "chart_area": self._shapes_chart,
            "chart_stacked_bar": self._shapes_chart,
            "table": self._shapes_table,
            "kpi_cards": self._shapes_kpi,
            "process_flow": self._shapes_process,
            "timeline": self._shapes_timeline,
            "two_column": self._shapes_two_column,
            "three_column": self._shapes_three_column,
            "comparison_cards": self._shapes_comparison_cards,
            "icon_grid": self._shapes_icon_grid,
            "funnel": self._shapes_funnel,
            "divider_layout": self._shapes_divider,
        }
        return dispatch.get(vt, self._shapes_bullets)(slide)

    # ═══════════════════════════════════════════════
    # HEADER BLOCK — title + subtitle + accent line + side accent bar
    # Every content slide gets ~5 shapes just for the header
    # ═══════════════════════════════════════════════

    def _add_header(self, shapes: list[ShapeSpec], slide: OptimizedSlideContent) -> float:
        """Add slide header decorations. Title/subtitle are rendered via template
        placeholders (PH0/PH1) in the renderer, so we only add the accent bar,
        divider line, footer line, and topic icon here."""

        variant = getattr(self.theme, "style_variant", "classic")

        # Left accent bar varies by style variant:
        #   classic   → thin 0.07" vertical bar next to the title
        #   banded    → wider 0.15" block for a denser editorial feel
        #   underline → no left bar; rely on the divider underline alone
        if variant == "banded":
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=0.15, top=0.40, width=0.15, height=0.92),
                fill_color="primary",
            ))
        elif variant == "classic":
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=0.15, top=0.40, width=0.07, height=0.92),
                fill_color="primary",
            ))
        # underline variant: no left bar

        # Divider thickness also varies so the header weight differs at a glance
        divider_weight = {"classic": 1.5, "banded": 3.0, "underline": 2.25}.get(variant, 1.5)

        # Full-width accent divider line below the title/subtitle area
        divider_y = GRID.MARGIN_TOP - 0.08   # just above content start
        shapes.append(ShapeSpec(
            shape_type="line",
            position=Position(left=GRID.TITLE_LEFT, top=divider_y,
                              width=GRID.TITLE_WIDTH, height=0),
            font_color="primary", border_width=divider_weight,
        ))

        # Footer accent line at bottom
        shapes.append(ShapeSpec(
            shape_type="line",
            position=Position(left=GRID.TITLE_LEFT, top=7.05,
                              width=GRID.TITLE_WIDTH, height=0),
            font_color="lt2", border_width=0.75,
        ))

        # Footer breadcrumb: show key_message (not title) so there is no
        # duplicate of the slide title at the bottom of every slide.
        if slide.key_message:
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=GRID.TITLE_LEFT, top=7.10, width=8.0, height=0.25),
                text=slide.key_message, font_size=9, font_color="dk2", alignment="left",
            ))

        # Topic-icon disabled. Emojis render inconsistently across PowerPoint
        # installations (Segoe UI Emoji missing / tofu boxes / monochrome
        # fallback) and the marker overlaps the right edge of long titles.
        # The accent bar + divider line already provide header identity.

        return GRID.MARGIN_TOP + 0.10  # content starts at ~1.50in

    # helper shortcuts
    def _line_h(self, shapes, left, top, width, color="lt2", weight=0.75):
        shapes.append(ShapeSpec(
            shape_type="line",
            position=Position(left=left, top=top, width=width, height=0),
            font_color=color, border_width=weight,
        ))

    def _line_v(self, shapes, left, top, height, color="lt2", weight=0.75):
        shapes.append(ShapeSpec(
            shape_type="line",
            position=Position(left=left, top=top, width=0, height=height),
            font_color=color, border_width=weight,
        ))

    def _diamond(self, shapes, left, top, size, color):
        """Small diamond-shaped icon marker."""
        shapes.append(ShapeSpec(
            shape_type="rectangle",
            position=Position(left=left, top=top, width=size, height=size),
            fill_color=color,
        ))

    def _accent_bar(self, shapes, left, top, height, color="primary"):
        shapes.append(ShapeSpec(
            shape_type="rectangle",
            position=Position(left=left, top=top, width=0.06, height=height),
            fill_color=color,
        ))

    def _icon_marker(self, shapes, left, top, size, icon_text, color="primary"):
        """Add a Unicode icon inside a colored shape. Shape varies per
        template style: classic=circle, banded=rounded square, underline=square."""
        variant = getattr(self.theme, "style_variant", "classic")
        shape_type = {
            "classic": "oval",
            "banded": "rounded_rect",
            "underline": "rectangle",
        }.get(variant, "oval")
        shapes.append(ShapeSpec(
            shape_type=shape_type,
            position=Position(left=left, top=top, width=size, height=size),
            text=icon_text, font_size=int(size * 22), font_bold=False,
            font_color="lt1", fill_color=color,
            alignment="center", vertical_alignment="middle",
        ))

    def _circle(self, shapes, left, top, size, color, text="", font_size=12):
        """Numbered/step marker. Shape varies by template style variant so
        different templates produce different bullet/step geometry."""
        variant = getattr(self.theme, "style_variant", "classic")
        shape_type = {
            "classic": "oval",
            "banded": "rounded_rect",
            "underline": "rectangle",
        }.get(variant, "oval")
        shapes.append(ShapeSpec(
            shape_type=shape_type,
            position=Position(left=left, top=top, width=size, height=size),
            text=text, font_size=font_size, font_bold=True,
            font_color="lt1", fill_color=color,
            alignment="center", vertical_alignment="middle",
        ))

    # ═══════════════════════════════════════════════
    # COVER
    # ═══════════════════════════════════════════════

    def _shapes_cover(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = [
            # Decorative accent bar (left side)
            ShapeSpec(shape_type="rectangle",
                      position=Position(left=0.15, top=2.8, width=0.08, height=2.5),
                      fill_color="primary"),
            # Title
            ShapeSpec(shape_type="text_box",
                      position=Position(left=0.375, top=3.1, width=9.3, height=0.7),
                      text=slide.title, font_size=36, font_bold=True,
                      alignment="left", vertical_alignment="middle"),
            # Accent line under title
            ShapeSpec(shape_type="line",
                      position=Position(left=0.375, top=3.9, width=4.0, height=0),
                      font_color="primary", border_width=2.0),
            # Subtitle
            ShapeSpec(shape_type="text_box",
                      position=Position(left=0.375, top=4.2, width=6.6, height=0.8),
                      text=slide.subtitle or "", font_size=14, font_color="dk2",
                      alignment="left", vertical_alignment="top"),
            # Decorative small squares at bottom
            ShapeSpec(shape_type="rectangle",
                      position=Position(left=0.375, top=6.8, width=0.15, height=0.15),
                      fill_color="primary"),
            ShapeSpec(shape_type="rectangle",
                      position=Position(left=0.60, top=6.8, width=0.15, height=0.15),
                      fill_color="accent2"),
            ShapeSpec(shape_type="rectangle",
                      position=Position(left=0.83, top=6.8, width=0.15, height=0.15),
                      fill_color="accent3"),
        ]
        return shapes

    # ═══════════════════════════════════════════════
    # BULLETS — numbered circles + accent bars + divider lines
    # Target: ~16 shapes for 5 bullets (was ~3)
    # ═══════════════════════════════════════════════

    def _shapes_bullets(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        """Clean full-width numbered bullet layout.

        Each row fills the available content height evenly — no hard cap on
        row height, so 3 bullets get a taller row than 5. The text box uses
        the full row height (not just the circle height) so long bullets
        wrap across multiple lines instead of clipping. Font size is
        selected per-row to fit the actual text.
        """
        shapes = []
        y = self._add_header(shapes, slide)

        bullets = (slide.bullets or slide.agenda_items or [])[:5]
        if not bullets:
            return shapes

        bl, bw = self.grid.span(0, 12)          # full content width

        # Distribute full available height across rows — no arbitrary cap.
        available_h = CONTENT_BOTTOM - y - 0.15
        n = len(bullets)
        row_gap = 0.14                            # small gap between rows
        item_h = max(0.60, (available_h - (n - 1) * row_gap) / n)

        # Circle size scales with row height (min 0.38", max 0.60").
        circle_size = max(0.38, min(0.60, item_h * 0.55))
        text_left   = bl + circle_size + 0.20
        text_width  = bw - circle_size - 0.20
        item_top    = y + 0.10
        body_font   = self.theme.fonts.minor

        # Font consistency: compute one size that fits the LONGEST bullet so
        # every row uses the same font size. Picking sizes per-bullet makes
        # the column look mismatched when one line wraps more than the others.
        font_pt = min(
            fit_font_size(b, width_in=text_width - 0.10,
                           height_in=item_h - 0.05,
                           max_pt=16.0, min_pt=10.0,
                           font_name=body_font)
            for b in bullets
        )

        for i, bullet in enumerate(bullets):
            by     = item_top + i * (item_h + row_gap)
            accent = f"accent{(i % 4) + 1}"
            cy     = by + (item_h - circle_size) / 2   # vertically centre circle

            # Accent number circle — font sized to circle, not hardcoded.
            circle_font = max(10, int(circle_size * 28))
            self._circle(shapes, bl, cy, circle_size, accent,
                         str(i + 1), font_size=circle_font)

            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=text_left, top=by,
                                  width=text_width, height=item_h),
                text=bullet, font_size=font_pt, font_color="dk2",
                alignment="left", vertical_alignment="middle",
            ))

            # Separator line between rows — full row width
            if i < n - 1:
                sep_y = by + item_h + row_gap / 2
                self._line_h(shapes, bl, sep_y, bw, "lt2", 0.75)

        return shapes

    # ═══════════════════════════════════════════════
    # CHART — title + chart + bottom accent line
    # ═══════════════════════════════════════════════

    def _shapes_chart(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        chart_w, chart_h = 10.5, 4.3
        chart_left = self.grid.center_horizontally(chart_w)

        shapes.append(ShapeSpec(
            shape_type="chart",
            position=Position(left=chart_left, top=y + GAP + 0.1, width=chart_w, height=chart_h),
        ))

        # Decorative side accent bars
        shapes.append(ShapeSpec(
            shape_type="rectangle",
            position=Position(left=0.15, top=y + GAP, width=0.06, height=chart_h + 0.2),
            fill_color="accent2",
        ))
        shapes.append(ShapeSpec(
            shape_type="rectangle",
            position=Position(left=GRID.SLIDE_WIDTH - 0.21, top=y + GAP, width=0.06, height=chart_h + 0.2),
            fill_color="accent2",
        ))

        # Horizontal line above chart
        self._line_h(shapes, chart_left, y + GAP, chart_w, "lt2", 0.5)
        # Horizontal line below chart
        self._line_h(shapes, chart_left, y + GAP + chart_h + 0.15, chart_w, "primary", 1.0)

        # Small decorative squares at bottom-right
        corner_x = GRID.SLIDE_WIDTH - 0.83
        corner_y = GRID.SLIDE_HEIGHT - GRID.MARGIN_BOTTOM - 0.30
        self._diamond(shapes, corner_x, corner_y, 0.12, "primary")
        self._diamond(shapes, corner_x + 0.20, corner_y, 0.12, "accent3")

        # Auto-fill: key message callout below chart if space allows
        chart_bottom = y + GAP + chart_h + 0.25
        if chart_bottom < 6.2 and slide.key_message:
            kl, kw = self.grid.span(1, 10)
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=kl, top=chart_bottom + 0.10, width=kw, height=0.45),
                text=f"✦  {slide.key_message}", font_size=12, font_bold=True,
                font_color="dk1", alignment="left", vertical_alignment="middle",
            ))
            self._line_h(shapes, kl, chart_bottom + 0.05, kw, "primary", 1.0)

        return shapes

    # ═══════════════════════════════════════════════
    # TABLE
    # ═══════════════════════════════════════════════

    def _shapes_table(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        left, width = self.grid.span(1, 10)
        rows = 1 + (len(slide.table_data.rows) if slide.table_data else 0)
        table_h = min(rows * 0.50, 4.2)

        # Side accent bars
        self._accent_bar(shapes, left - 0.15, y + GAP, table_h, "primary")

        shapes.append(ShapeSpec(
            shape_type="table",
            position=Position(left=left, top=y + GAP, width=width, height=table_h),
        ))

        # Decorative lines around table
        self._line_h(shapes, left, y + GAP + table_h + 0.05, width, "primary", 1.0)

        # Auto-fill: key message below table if space allows
        table_bottom = y + GAP + table_h + 0.15
        if table_bottom < 6.2 and slide.key_message:
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=left, top=table_bottom + 0.10, width=width, height=0.45),
                text=f"✦  {slide.key_message}", font_size=12, font_bold=True,
                font_color="dk1", alignment="left", vertical_alignment="middle",
            ))

        return shapes

    # ═══════════════════════════════════════════════
    # KPI CARDS — accent top bar + large value + label + description
    # Target: ~20 shapes for 4 KPIs (was ~6)
    # ═══════════════════════════════════════════════

    def _shapes_kpi(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        kpis = slide.kpi_values or []
        n = min(len(kpis), 4)
        if n == 0:
            return self._shapes_bullets(slide)

        variant = getattr(self.theme, "style_variant", "classic")
        if variant == "banded":
            return self._shapes_kpi_banded(shapes, y, kpis, n)
        if variant == "underline":
            return self._shapes_kpi_underline(shapes, y, kpis, n)

        card_top = y + GAP + 0.05
        available_h = CONTENT_BOTTOM - card_top - 0.15

        # If we have supplementary bullets, reserve space for them below the cards.
        has_bullets = bool(slide.bullets)
        bullets_reserve = 1.30 if has_bullets else 0.0
        card_h = max(2.20, min(3.00, available_h - bullets_reserve))

        accents = ["accent1", "accent2", "accent3", "accent4"]
        body_font = self.theme.fonts.minor

        # Sub-regions inside a card, as fractions of card_h.
        icon_h   = 0.24 * card_h
        value_h  = 0.30 * card_h
        label_h  = card_h - icon_h - value_h - 0.20

        # Cards share equal width; compute it once for the shared-font pass.
        if n <= 3:
            _, sample_w = self.grid.span(0, 4)
        else:
            _, sample_w = self.grid.quarter(0)
        sample_cw = sample_w - PAD * 2

        # Unified fonts for value + label across ALL KPI cards so the row
        # reads as a consistent set rather than four cards each at a
        # different size.
        value_font_pt = min(
            fit_font_size(k.value, width_in=sample_cw - 0.20,
                           height_in=value_h,
                           max_pt=32.0, min_pt=16.0,
                           font_name=self.theme.fonts.major)
            for k in kpis[:n]
        )
        label_font_pt = min(
            fit_font_size(k.label, width_in=sample_cw - 0.20,
                           height_in=max(0.20, label_h),
                           max_pt=11.0, min_pt=8.0,
                           font_name=body_font)
            for k in kpis[:n]
        )

        variant = getattr(self.theme, "style_variant", "classic")

        for i in range(n):
            if n <= 3:
                left, width = self.grid.span(i * 4, 4)
            else:
                left, width = self.grid.quarter(i)

            accent = accents[i % len(accents)]
            cl = left + PAD
            cw = width - PAD * 2

            # Card shape varies by style:
            #   classic   → rounded_rect card with horizontal accent at TOP
            #   banded    → sharp rectangle card with vertical accent at LEFT
            #   underline → rounded_rect card with accent at BOTTOM only
            card_shape = "rectangle" if variant == "banded" else "rounded_rect"
            shapes.append(ShapeSpec(
                shape_type=card_shape,
                position=Position(left=cl, top=card_top, width=cw, height=card_h),
                fill_color="lt2",
            ))

            if variant == "banded":
                # Left vertical accent bar, full card height
                shapes.append(ShapeSpec(
                    shape_type="rectangle",
                    position=Position(left=cl, top=card_top,
                                      width=0.10, height=card_h),
                    fill_color=accent,
                ))
            elif variant == "underline":
                # Thin bottom underline
                shapes.append(ShapeSpec(
                    shape_type="rectangle",
                    position=Position(left=cl + 0.10, top=card_top + card_h - 0.05,
                                      width=cw - 0.20, height=0.05),
                    fill_color=accent,
                ))
            else:  # classic
                shapes.append(ShapeSpec(
                    shape_type="rectangle",
                    position=Position(left=cl, top=card_top, width=cw, height=0.07),
                    fill_color=accent,
                ))

            # Domain icon — size scales with card height
            icon_size = max(0.40, min(0.60, icon_h * 0.70))
            self._icon_marker(shapes, cl + cw / 2 - icon_size / 2,
                              card_top + 0.14,
                              icon_size, _pick_icon(kpis[i].label), accent)

            # KPI value — uses unified font size
            val_top = card_top + icon_h + 0.10
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl + 0.10, top=val_top,
                                  width=cw - 0.20, height=value_h),
                text=kpis[i].value, font_size=value_font_pt, font_bold=True,
                font_color="dk1", alignment="center", vertical_alignment="middle",
            ))

            # Separator line between value and label
            sep_y = val_top + value_h + 0.04
            self._line_h(shapes, cl + 0.30, sep_y, cw - 0.60, accent, 1.0)

            # Label — uses unified font size
            label_top = sep_y + 0.05
            remaining = card_top + card_h - label_top - 0.10
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl + 0.10, top=label_top,
                                  width=cw - 0.20, height=remaining),
                text=kpis[i].label, font_size=label_font_pt, font_color="dk2",
                alignment="center", vertical_alignment="top",
            ))

            # Vertical divider between cards
            if i < n - 1:
                self._line_v(shapes, left + width, card_top + 0.2, card_h - 0.4)

        # Supplementary bullets fill the remaining space below the cards
        if has_bullets:
            bleft, bwidth = self.grid.span(1, 10)
            bullets = slide.bullets[:3]
            bullet_text = '\n'.join(f'• {b}' for b in bullets)
            bullets_top = card_top + card_h + 0.20
            bullets_h   = CONTENT_BOTTOM - bullets_top - 0.10
            if bullets_h > 0.40:
                bullet_font = fit_multi_line_font_size(
                    bullets, width_in=bwidth - 0.10, height_in=bullets_h,
                    max_pt=13.0, min_pt=9.0, font_name=body_font,
                )
                shapes.append(ShapeSpec(
                    shape_type="text_box",
                    position=Position(left=bleft, top=bullets_top,
                                      width=bwidth, height=bullets_h),
                    text=bullet_text, font_size=bullet_font, font_color="dk2",
                    alignment="left", vertical_alignment="top",
                ))

        return shapes

    # ═══════════════════════════════════════════════
    # KPI — banded variant (full-width horizontal strips)
    # Row-per-KPI dashboard list: accent rail | icon | value | label
    # ═══════════════════════════════════════════════
    def _shapes_kpi_banded(self, shapes, y, kpis, n):
        row_top = y + GAP + 0.10
        available_h = CONTENT_BOTTOM - row_top - 0.15
        row_gap = 0.10
        row_h = max(0.85, (available_h - (n - 1) * row_gap) / n)

        accents = ["accent1", "accent2", "accent3", "accent4"]
        body_font = self.theme.fonts.minor
        bl, bw = self.grid.span(0, 12)

        value_col_w = 2.80
        value_font_pt = min(
            fit_font_size(k.value, width_in=value_col_w - 0.20,
                           height_in=row_h * 0.80,
                           max_pt=36.0, min_pt=18.0,
                           font_name=self.theme.fonts.major)
            for k in kpis[:n]
        )
        label_col_w = bw - (0.12 + 0.30 + 0.60 + 0.30 + value_col_w + 0.30) - 0.20
        label_font_pt = min(
            fit_font_size(k.label, width_in=label_col_w,
                           height_in=row_h * 0.70,
                           max_pt=15.0, min_pt=10.0,
                           font_name=body_font)
            for k in kpis[:n]
        )

        for i in range(n):
            accent = accents[i % len(accents)]
            ry = row_top + i * (row_h + row_gap)

            # Background strip
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=bl, top=ry, width=bw, height=row_h),
                fill_color="lt2",
            ))
            # Left accent rail, full height
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=bl, top=ry, width=0.12, height=row_h),
                fill_color=accent,
            ))

            # Icon
            icon_size = min(row_h - 0.25, 0.60)
            icon_left = bl + 0.30
            icon_top = ry + (row_h - icon_size) / 2
            self._icon_marker(shapes, icon_left, icon_top, icon_size,
                              _pick_icon(kpis[i].label), accent)

            # Value (left block)
            value_left = icon_left + icon_size + 0.30
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=value_left, top=ry,
                                  width=value_col_w, height=row_h),
                text=kpis[i].value, font_size=value_font_pt,
                font_bold=True, font_color="dk1",
                alignment="left", vertical_alignment="middle",
            ))

            # Vertical separator
            sep_x = value_left + value_col_w + 0.15
            self._line_v(shapes, sep_x, ry + 0.12, row_h - 0.24, "lt1", 1.5)

            # Label (right block, takes remaining width)
            label_left = sep_x + 0.25
            label_width = bl + bw - label_left - 0.20
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=label_left, top=ry,
                                  width=label_width, height=row_h),
                text=kpis[i].label, font_size=label_font_pt,
                font_color="dk2", alignment="left", vertical_alignment="middle",
            ))

        return shapes

    # ═══════════════════════════════════════════════
    # KPI — underline variant (editorial typography, no cards)
    # Huge numbers in accent color + thin underline + label beneath
    # ═══════════════════════════════════════════════
    def _shapes_kpi_underline(self, shapes, y, kpis, n):
        top = y + GAP + 0.25
        available_h = CONTENT_BOTTOM - top - 0.30

        accents = ["accent1", "accent2", "accent3", "accent4"]
        body_font = self.theme.fonts.minor

        cell_w = GRID.content_width / n
        value_h = available_h * 0.55
        label_h = available_h * 0.30

        value_font_pt = min(
            fit_font_size(k.value, width_in=cell_w - 0.40, height_in=value_h,
                           max_pt=64.0, min_pt=26.0,
                           font_name=self.theme.fonts.major)
            for k in kpis[:n]
        )
        label_font_pt = min(
            fit_font_size(k.label, width_in=cell_w - 0.40, height_in=label_h,
                           max_pt=13.0, min_pt=9.0,
                           font_name=body_font)
            for k in kpis[:n]
        )

        for i in range(n):
            accent = accents[i % len(accents)]
            cl = GRID.content_left + i * cell_w

            # Big value (accent color, left aligned in cell)
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl + 0.15, top=top,
                                  width=cell_w - 0.30, height=value_h),
                text=kpis[i].value, font_size=value_font_pt, font_bold=True,
                font_color=accent, alignment="left", vertical_alignment="bottom",
            ))

            # Underline below value
            ul_y = top + value_h + 0.08
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=cl + 0.15, top=ul_y,
                                  width=cell_w * 0.45, height=0.05),
                fill_color=accent,
            ))

            # Label below
            lbl_top = ul_y + 0.18
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl + 0.15, top=lbl_top,
                                  width=cell_w - 0.30, height=label_h),
                text=kpis[i].label, font_size=label_font_pt, font_bold=False,
                font_color="dk2", alignment="left", vertical_alignment="top",
            ))

            # Thin vertical divider between cells (skip after last)
            if i < n - 1:
                self._line_v(shapes, GRID.content_left + (i + 1) * cell_w,
                             top + 0.10, available_h - 0.20, "lt2", 0.75)

        return shapes

    # ═══════════════════════════════════════════════
    # PROCESS FLOW — step circles + boxes + arrows
    # ═══════════════════════════════════════════════

    def _shapes_process(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        steps = slide.process_steps or []
        n = min(len(steps), 5)
        if n == 0:
            return self._shapes_bullets(slide)

        # Centre the step row vertically in the remaining content area.
        avail_h = CONTENT_BOTTOM - y - 0.30
        step_h  = max(1.40, min(2.40, avail_h * 0.55))
        step_top = y + (avail_h - step_h) / 2 + 0.20
        arrow_w = 0.5
        total_w = GRID.content_width
        step_w = (total_w - (n - 1) * arrow_w) / n
        body_font = self.theme.fonts.minor
        box_w = step_w - 0.10
        box_h = step_h - 0.50

        # Unified label font across all process steps
        label_font_pt = min(
            fit_font_size(steps[i].label, width_in=box_w - 0.20,
                           height_in=box_h - 0.10,
                           max_pt=14.0, min_pt=9.0,
                           font_name=body_font)
            for i in range(n)
        )

        for i in range(n):
            x = GRID.content_left + i * (step_w + arrow_w)
            accent = f"accent{(i % 6) + 1}"

            # Step icon circle on top (domain-aware, fallback to number)
            step_icon = _pick_icon(steps[i].label)
            if step_icon == _DEFAULT_ICON:
                step_icon = str(i + 1)
            self._icon_marker(shapes, x + step_w / 2 - 0.22, step_top - 0.15,
                              0.44, step_icon, accent)

            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=x, top=step_top + 0.40,
                                  width=box_w, height=box_h),
                text=steps[i].label, font_size=label_font_pt, font_bold=True,
                font_color="dk1", fill_color="lt2",
                alignment="center", vertical_alignment="middle",
            ))

            # Bottom accent line under each box
            self._line_h(shapes, x, step_top + step_h - 0.05, box_w, accent, 2.0)

            # Arrow connector
            if i < n - 1:
                shapes.append(ShapeSpec(
                    shape_type="arrow",
                    position=Position(left=x + step_w - 0.05,
                                      top=step_top + step_h / 2,
                                      width=arrow_w, height=0.22),
                    font_color="primary",
                ))

        return shapes

    # ═══════════════════════════════════════════════
    # TIMELINE — markers + lines + labels
    # ═══════════════════════════════════════════════

    def _shapes_timeline(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        items = slide.timeline_items or []
        n = min(len(items), 6)
        if n == 0:
            return self._shapes_bullets(slide)

        left, width = self.grid.span(1, 10)

        # Centre the timeline vertically in the remaining content area
        avail_h = CONTENT_BOTTOM - y - 0.30
        line_y  = y + 0.30 + avail_h * 0.40
        date_box_h = 0.40
        label_top  = line_y + 0.45
        label_h    = CONTENT_BOTTOM - label_top - 0.10

        # Main horizontal line (thick accent)
        self._line_h(shapes, left, line_y, width, "primary", 3.0)

        spacing = width / (n + 1)
        body_font = self.theme.fonts.minor

        # Unified label font across all timeline items
        label_font_pt = min(
            fit_font_size(items[i].label, width_in=1.6,
                           height_in=label_h,
                           max_pt=11.0, min_pt=8.0,
                           font_name=body_font)
            for i in range(n)
        )

        for i in range(n):
            x = left + spacing * (i + 1)
            accent = f"accent{(i % 6) + 1}"

            # Milestone marker circle
            self._circle(shapes, x - 0.22, line_y - 0.22, 0.44, accent)

            # Connecting vertical line from circle to date
            self._line_v(shapes, x, line_y - 0.85, 0.60, accent, 1.5)

            # Date above
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=x - 0.65, top=line_y - 1.3,
                                  width=1.3, height=date_box_h),
                text=items[i].date, font_size=12, font_bold=True,
                font_color="lt1", fill_color=accent,
                alignment="center", vertical_alignment="middle",
            ))

            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=x - 0.8, top=label_top,
                                  width=1.6, height=label_h),
                text=items[i].label, font_size=label_font_pt,
                font_color="dk2", alignment="center", vertical_alignment="top",
            ))

        return shapes

    # ═══════════════════════════════════════════════
    # TWO COLUMN — column headers + divider + accent bars
    # Target: ~14 shapes (was ~4)
    # ═══════════════════════════════════════════════

    def _shapes_two_column(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        ll, lw = self.grid.half_left()
        rl, rw = self.grid.half_right()
        col_top = y + GAP
        col_h   = CONTENT_BOTTOM - col_top - 0.15   # fill to footer

        left_title = "Key Points"
        right_title = "Details"
        if slide.comparison_items and len(slide.comparison_items) >= 2:
            left_title = slide.comparison_items[0].title
            right_title = slide.comparison_items[1].title

        left_bullets = (slide.left_column or [])[:4]
        right_bullets = (slide.right_column or [])[:4]

        header_h = 0.50
        body_top = col_top + header_h + 0.08
        body_h   = col_h - header_h - 0.12
        body_font = self.theme.fonts.minor

        # Per-column geometry (identical for both columns since widths match)
        inner_width = lw - PAD * 2
        n_left = len(left_bullets) or 1
        n_right = len(right_bullets) or 1
        row_gap = 0.10
        # Row height uses the column with the most bullets so both columns
        # share the same row geometry → same dot / text box sizing.
        n_for_rows = max(n_left, n_right)
        row_h = max(0.50, (body_h - 0.20 - (n_for_rows - 1) * row_gap) / n_for_rows)
        dot_size = max(0.14, min(0.22, row_h * 0.28))
        text_width = inner_width - 0.30 - dot_size

        # Unified body font across BOTH columns — pick the size that fits the
        # longest bullet anywhere in the slide so left/right rows match.
        all_bullets = [b for b in (left_bullets + right_bullets) if b]
        if all_bullets:
            body_font_pt = min(
                fit_font_size(b, width_in=text_width - 0.10,
                               height_in=row_h - 0.06,
                               max_pt=14.0, min_pt=10.0,
                               font_name=body_font)
                for b in all_bullets
            )
        else:
            body_font_pt = 14.0

        # Unified header font across both columns
        header_font_pt = min(
            fit_font_size(t, width_in=inner_width - 0.60,
                           height_in=header_h - 0.10,
                           max_pt=15.0, min_pt=11.0,
                           font_name=self.theme.fonts.major)
            for t in (left_title, right_title)
        )

        for col_idx, (cl, cw, title, bullets) in enumerate([
            (ll, lw, left_title, left_bullets),
            (rl, rw, right_title, right_bullets),
        ]):
            accent = f"accent{col_idx + 1}"
            inner_left  = cl + PAD

            # Domain icon in column header
            col_icon = _pick_icon(title)
            self._icon_marker(shapes, inner_left + 0.08, col_top + 0.05, 0.40, col_icon, accent)

            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=inner_left, top=col_top,
                                  width=inner_width, height=header_h),
                text=f"  {title}", font_size=header_font_pt, font_bold=True,
                font_color="lt1", fill_color=accent,
                alignment="center", vertical_alignment="middle",
            ))

            # Column body background fills remaining vertical space
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=inner_left, top=body_top,
                                  width=inner_width, height=body_h),
                fill_color="lt2",
            ))

            if not bullets:
                continue

            text_left = inner_left + 0.16 + dot_size + 0.14

            for j, b in enumerate(bullets):
                by = body_top + 0.10 + j * (row_h + row_gap)

                # Accent dot (vertically centred in the row)
                self._circle(shapes, inner_left + 0.16,
                              by + (row_h - dot_size) / 2,
                              dot_size, accent)

                shapes.append(ShapeSpec(
                    shape_type="text_box",
                    position=Position(left=text_left, top=by,
                                      width=text_width, height=row_h),
                    text=b, font_size=body_font_pt, font_color="dk2",
                    alignment="left", vertical_alignment="middle",
                ))

                # Separator line between rows
                if j < len(bullets) - 1:
                    self._line_h(shapes, text_left, by + row_h + row_gap / 2,
                                  text_width, "lt2", 0.5)

        # Vertical divider between columns
        mid_x = (ll + lw + rl) / 2
        self._line_v(shapes, mid_x, col_top, col_h)

        return shapes

    # ═══════════════════════════════════════════════
    # THREE COLUMN — accent strip + header + accent dots + dividers
    # Target: ~30 shapes for 3 columns (was ~7)
    # ═══════════════════════════════════════════════

    def _shapes_three_column(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        items = slide.comparison_items or []
        n = min(len(items), 3)
        col_top = y + GAP
        body_h  = CONTENT_BOTTOM - col_top - 0.15   # fill remaining space

        header_block_h = 1.00   # icon + header + underline
        body_font = self.theme.fonts.minor

        # Columns share equal width/row geometry; compute once.
        sample_left, sample_width = self.grid.third(0)
        cw = sample_width - PAD * 2
        max_points = max((len(item.points[:3]) for item in items[:n]), default=1) or 1
        row_gap = 0.12
        avail_body_h = body_h - header_block_h - 0.20
        row_h = max(0.55, (avail_body_h - (max_points - 1) * row_gap) / max_points)
        dot_size = max(0.12, min(0.20, row_h * 0.22))
        text_width = cw - 0.20 - dot_size

        # Unified font sizes across all three columns
        all_points = [p for item in items[:n] for p in item.points[:3]]
        if all_points:
            body_font_pt = min(
                fit_font_size(p, width_in=text_width - 0.05,
                               height_in=row_h - 0.06,
                               max_pt=14.0, min_pt=10.0,
                               font_name=body_font)
                for p in all_points
            )
        else:
            body_font_pt = 14.0

        header_font_pt = min(
            fit_font_size(items[i].title, width_in=cw - 0.10,
                           height_in=0.35,
                           max_pt=15.0, min_pt=11.0,
                           font_name=self.theme.fonts.major)
            for i in range(n)
        )

        for i in range(n):
            left, width = self.grid.third(i)
            item = items[i]
            accent = f"accent{(i % 6) + 1}"
            cl = left + PAD

            # Accent top strip
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=cl, top=col_top, width=cw, height=0.07),
                fill_color=accent,
            ))

            # Domain icon above column header
            col_icon = _pick_icon(item.title)
            self._icon_marker(shapes, cl + cw / 2 - 0.20, col_top + 0.14, 0.40, col_icon, accent)

            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl, top=col_top + 0.58, width=cw, height=0.35),
                text=item.title, font_size=header_font_pt, font_bold=True,
                font_color="dk1", alignment="center", vertical_alignment="middle",
            ))

            # Header underline
            self._line_h(shapes, cl + 0.15, col_top + 0.96, cw - 0.30, accent, 1.0)

            # Body points
            points = item.points[:3]
            if not points:
                continue
            points_top = col_top + header_block_h + 0.10
            text_left  = cl + 0.10 + dot_size + 0.10
            pn = len(points)

            for j, point in enumerate(points):
                py = points_top + j * (row_h + row_gap)

                self._circle(shapes, cl + 0.10, py + (row_h - dot_size) / 2,
                              dot_size, accent)

                shapes.append(ShapeSpec(
                    shape_type="text_box",
                    position=Position(left=text_left, top=py,
                                      width=text_width, height=row_h),
                    text=point, font_size=body_font_pt, font_color="dk2",
                    alignment="left", vertical_alignment="middle",
                ))

                # Separator between body rows
                if j < pn - 1:
                    self._line_h(shapes, cl + 0.10, py + row_h + row_gap / 2,
                                  cw - 0.20, "lt2", 0.5)

            # Vertical divider between columns
            if i < n - 1:
                self._line_v(shapes, left + width, col_top + 0.1, body_h)

        return shapes

    # ═══════════════════════════════════════════════
    # ICON GRID — 2×2 or 2×3 cards each with large icon + title + description
    # Best for: 4–6 distinct categorised items (features, benefits, principles)
    # ═══════════════════════════════════════════════

    def _shapes_icon_grid(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        items = (slide.comparison_items or [])[:6]
        n = len(items)
        if n == 0:
            return self._shapes_bullets(slide)

        cols = 3 if n >= 5 else 2
        rows = (n + cols - 1) // cols  # ceiling division

        available_h = CONTENT_BOTTOM - y - 0.10
        card_w = (GRID.content_width - (cols - 1) * GAP) / cols
        card_h = (available_h - (rows - 1) * GAP) / rows

        body_font = self.theme.fonts.minor
        title_font_name = self.theme.fonts.major

        # Unified title/description font sizes across all cards
        title_h = 0.40
        desc_top_offset = 0.14 + min(card_h * 0.30, 0.58) + 0.08 + title_h + 0.10
        desc_h = max(0.20, card_h - desc_top_offset - 0.10)
        title_font_pt = min(
            fit_font_size(item.title, width_in=card_w - 0.24,
                           height_in=title_h,
                           max_pt=14.0, min_pt=10.0,
                           font_name=title_font_name)
            for item in items
        )
        desc_texts = [item.points[0] for item in items if item.points]
        if desc_texts:
            desc_font_pt = min(
                fit_font_size(t, width_in=card_w - 0.24,
                               height_in=desc_h,
                               max_pt=12.0, min_pt=9.0,
                               font_name=body_font)
                for t in desc_texts
            )
        else:
            desc_font_pt = 12.0

        for i, item in enumerate(items):
            col_idx = i % cols
            row_idx = i // cols

            cl = GRID.content_left + col_idx * (card_w + GAP)
            ct = y + 0.05 + row_idx * (card_h + GAP)
            accent = f"accent{(i % 6) + 1}"

            # Card background
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=cl, top=ct, width=card_w, height=card_h),
                fill_color="lt2",
            ))
            # Top accent strip
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=cl, top=ct, width=card_w, height=0.06),
                fill_color=accent,
            ))

            # Centered icon — size scales with card height
            icon_size = min(card_h * 0.30, 0.58)
            icon_left = cl + card_w / 2 - icon_size / 2
            icon_top = ct + 0.14
            self._icon_marker(shapes, icon_left, icon_top,
                              icon_size, _pick_icon(item.title), accent)

            # Title — uses the unified font size computed above
            title_top = icon_top + icon_size + 0.08
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl + 0.12, top=title_top,
                                  width=card_w - 0.24, height=title_h),
                text=item.title, font_size=title_font_pt, font_bold=True,
                font_color="dk1", alignment="center", vertical_alignment="middle",
            ))

            # Thin accent separator under title
            sep_y = title_top + title_h + 0.04
            self._line_h(shapes, cl + 0.25, sep_y, card_w - 0.50, accent, 0.75)

            # Description — fills the card's remaining vertical space
            if item.points:
                desc_top = sep_y + 0.06
                remaining = card_h - (desc_top - ct) - 0.10
                if remaining > 0.18:
                    shapes.append(ShapeSpec(
                        shape_type="text_box",
                        position=Position(left=cl + 0.12, top=desc_top,
                                          width=card_w - 0.24, height=remaining),
                        text=item.points[0], font_size=desc_font_pt, font_color="dk2",
                        alignment="center", vertical_alignment="top",
                    ))

            # Vertical separator between cards in the same row (not after last col)
            if col_idx < cols - 1 and i < n - 1:
                self._line_v(shapes, cl + card_w + GAP / 2, ct + 0.15, card_h - 0.30)

        return shapes

    # ═══════════════════════════════════════════════
    # FUNNEL — stacked decreasing-width bars for pipeline / stage data
    # Best for: sales funnels, hiring pipelines, adoption stages
    # ═══════════════════════════════════════════════

    def _shapes_funnel(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        items = (slide.funnel_items or [])[:5]
        n = len(items)
        if n == 0:
            return self._shapes_bullets(slide)

        cx = GRID.SLIDE_WIDTH / 2           # horizontal center of slide
        max_w = GRID.content_width * 0.86   # widest bar
        min_w = max_w * 0.28                # narrowest bar

        # Fill the available vertical space — bar height scales with count.
        avail_h = CONTENT_BOTTOM - y - 0.20
        gap = max(0.10, avail_h * 0.04)
        bar_h = (avail_h - (n + 1) * gap) / n
        bar_h = max(0.55, min(1.10, bar_h))
        body_font = self.theme.fonts.major

        # Unified label font across all bars — narrowest bar is the limiting
        # factor, so it sets the size the wider bars will also use.
        label_texts = [
            (f"{it.label}  ·  {it.value}" if it.value else it.label)
            for it in items
        ]
        # Width of the narrowest bar (bottom)
        narrowest_w = min_w if n > 1 else max_w
        label_font_pt = min(
            fit_font_size(t, width_in=narrowest_w - 0.24,
                           height_in=bar_h - 0.08,
                           max_pt=16.0, min_pt=10.0,
                           font_name=body_font)
            for t in label_texts
        )

        for i, item in enumerate(items):
            # Width linearly interpolated from max → min (top to bottom)
            t = i / (n - 1) if n > 1 else 0.0
            w = max_w - (max_w - min_w) * t
            left = cx - w / 2
            top = y + gap + i * (bar_h + gap)
            accent = f"accent{(i % 6) + 1}"

            # Funnel bar
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=left, top=top, width=w, height=bar_h),
                fill_color=accent,
            ))

            label_text = label_texts[i]
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=left + 0.12, top=top,
                                  width=w - 0.24, height=bar_h),
                text=label_text, font_size=label_font_pt, font_bold=True,
                font_color="lt1", alignment="center", vertical_alignment="middle",
            ))

            # Step number badge on left edge
            self._circle(shapes, left - 0.32, top + bar_h / 2 - 0.18,
                          0.36, accent, str(i + 1), 12)

            # Small connector diamond between bars
            if i < n - 1:
                next_t = (i + 1) / (n - 1) if n > 1 else 0.0
                next_w = max_w - (max_w - min_w) * next_t
                diamond_w = max(next_w * 0.06, 0.14)
                self._diamond(shapes, cx - diamond_w / 2,
                               top + bar_h + gap / 2 - diamond_w / 2,
                               diamond_w, accent)

        # Right-side annotation: value labels if present and we haven't already inline'd them
        if any(item.value for item in items):
            for i, item in enumerate(items):
                if not item.value:
                    continue
                t = i / (n - 1) if n > 1 else 0.0
                w = max_w - (max_w - min_w) * t
                bar_right = cx + w / 2
                top = y + gap + i * (bar_h + gap)
                if bar_right + 0.20 < GRID.SLIDE_WIDTH - 0.20:
                    shapes.append(ShapeSpec(
                        shape_type="text_box",
                        position=Position(left=bar_right + 0.10, top=top,
                                          width=1.50, height=bar_h),
                        text=item.value, font_size=12, font_bold=True,
                        font_color=f"accent{(i % 6) + 1}",
                        alignment="left", vertical_alignment="middle",
                    ))

        return shapes

    # ═══════════════════════════════════════════════
    # COMPARISON CARDS — VS badge + two feature cards with check marks
    # Best for: option A vs option B, before vs after, method comparisons
    # ═══════════════════════════════════════════════

    def _shapes_comparison_cards(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes = []
        y = self._add_header(shapes, slide)

        items = (slide.comparison_items or [])[:2]
        if len(items) < 2:
            return self._shapes_two_column(slide)

        card_top = y + 0.15
        card_h   = CONTENT_BOTTOM - card_top - 0.10   # fill vertical space
        # Split content width: two cards with a VS-badge gap between them
        badge_gap = 1.40
        card_w = (GRID.content_width - badge_gap) / 2

        left_x = GRID.content_left
        right_x = GRID.content_left + card_w + badge_gap
        badge_cx = GRID.content_left + card_w + badge_gap / 2   # center of gap
        accents = ["accent1", "accent2"]
        header_h = 0.68
        body_font = self.theme.fonts.minor

        # Unified row geometry based on the card with the most points
        max_points = max((len(item.points[:4]) for item in items), default=1) or 1
        row_gap = 0.12
        avail_body_h = card_h - header_h - 0.30
        row_h = max(0.55, (avail_body_h - (max_points - 1) * row_gap) / max_points)
        check_size = max(0.24, min(0.34, row_h * 0.35))
        text_width = card_w - 0.28 - check_size

        # Unified body + header font across both cards
        all_points = [p for item in items for p in item.points[:4]]
        if all_points:
            body_font_pt = min(
                fit_font_size(p, width_in=text_width - 0.05,
                               height_in=row_h - 0.06,
                               max_pt=14.0, min_pt=10.0,
                               font_name=body_font)
                for p in all_points
            )
        else:
            body_font_pt = 14.0

        header_font_pt = min(
            fit_font_size(item.title, width_in=card_w - 0.86,
                           height_in=0.44,
                           max_pt=15.0, min_pt=11.0,
                           font_name=self.theme.fonts.major)
            for item in items
        )

        for ci, (cx, item) in enumerate(zip([left_x, right_x], items)):
            accent = accents[ci]

            # Card shadow/background
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=cx, top=card_top, width=card_w, height=card_h),
                fill_color="lt2",
            ))
            # Coloured header banner
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=cx, top=card_top, width=card_w, height=header_h),
                fill_color=accent,
            ))
            # Domain icon in header (left-aligned)
            col_icon = _pick_icon(item.title)
            self._icon_marker(shapes, cx + 0.10, card_top + 0.10, 0.46, col_icon, accent)
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cx + 0.66, top=card_top + 0.12,
                                  width=card_w - 0.76, height=0.44),
                text=item.title, font_size=header_font_pt, font_bold=True,
                font_color="lt1", alignment="left", vertical_alignment="middle",
            ))

            # Feature points — divide remaining card height evenly
            points = item.points[:4]
            pn = len(points)
            if pn == 0:
                continue
            points_top = card_top + header_h + 0.15
            text_left  = cx + 0.14 + check_size + 0.14

            for j, point in enumerate(points):
                py = points_top + j * (row_h + row_gap)

                # Check mark — vertically centered in the row
                self._circle(shapes, cx + 0.14,
                              py + (row_h - check_size) / 2,
                              check_size, accent, "✓", max(9, int(check_size * 26)))

                shapes.append(ShapeSpec(
                    shape_type="text_box",
                    position=Position(left=text_left, top=py,
                                      width=text_width, height=row_h),
                    text=point, font_size=body_font_pt, font_color="dk2",
                    alignment="left", vertical_alignment="middle",
                ))

                # Separator line (not after last point)
                if j < pn - 1:
                    self._line_h(shapes, cx + 0.14, py + row_h + row_gap / 2,
                                  card_w - 0.28, "lt2", 0.5)

        # VS badge centered in the gap
        vs_size = 0.78
        vs_x = badge_cx - vs_size / 2
        vs_y = card_top + card_h / 2 - vs_size / 2
        shapes.append(ShapeSpec(
            shape_type="oval",
            position=Position(left=vs_x, top=vs_y, width=vs_size, height=vs_size),
            text="VS", font_size=14, font_bold=True,
            font_color="lt1", fill_color="dk2",
            alignment="center", vertical_alignment="middle",
        ))

        return shapes

    # ═══════════════════════════════════════════════
    # SECTION DIVIDER — bold full-slide break between major sections
    # Uses the "C_Section blue" template layout which provides a dark background.
    # All text here must be white/light to read against the dark template.
    # ═══════════════════════════════════════════════

    def _shapes_divider(self, slide: OptimizedSlideContent) -> list[ShapeSpec]:
        shapes: list[ShapeSpec] = []

        cx = GRID.SLIDE_WIDTH / 2
        cy = GRID.SLIDE_HEIGHT / 2

        # No full-slide background rectangle here — the Divider layout in every
        # supported template already provides an appropriate dark background.
        # Adding one here would override it with dk1 which is black in some themes.

        # Full-width accent bar just above center
        self._line_h(shapes, 0.40, cy - 0.85, GRID.SLIDE_WIDTH - 0.80, "primary", 2.5)

        # Large section title — taller box so long titles never overflow into the line below
        title_box_h = 1.40
        shapes.append(ShapeSpec(
            shape_type="text_box",
            position=Position(left=0.43, top=cy - 0.72, width=9.50, height=title_box_h),
            text=slide.title, font_size=36, font_bold=True,
            font_color="lt1", alignment="left",
        ))

        # Thin line below title — placed safely below title box bottom
        line_below_y = cy - 0.72 + title_box_h + 0.12   # 0.12" clear gap after box
        self._line_h(shapes, 0.40, line_below_y, 8.0, "lt2", 0.75)

        # Subtitle / key message — starts below the line
        if slide.key_message:
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=0.43, top=line_below_y + 0.10, width=9.50, height=0.55),
                text=slide.key_message, font_size=18, font_bold=False,
                font_color="lt2", alignment="left",
            ))

        # Left accent bar — slim (0.10" wide) so it doesn't crowd content
        shapes.append(ShapeSpec(
            shape_type="rectangle",
            position=Position(left=0.00, top=0.25, width=0.10, height=5.58),
            fill_color="primary",
        ))

        # Section-number badge (accent-colored circle, top-right quadrant)
        topic_icon = _pick_icon(slide.title)
        self._icon_marker(shapes, GRID.SLIDE_WIDTH - 2.20, cy - 1.10,
                          1.10, topic_icon, "accent2")

        # Three decorative dots at bottom-right (matches sample's corner squares)
        corner_x = GRID.SLIDE_WIDTH - 0.80
        corner_y = GRID.SLIDE_HEIGHT - 0.35
        for i, accent in enumerate(["accent1", "accent2", "accent3"]):
            self._diamond(shapes, corner_x + i * 0.22, corner_y, 0.14, accent)

        return shapes
