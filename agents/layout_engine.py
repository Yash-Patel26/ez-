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

PAD = 0.15
GAP = 0.20

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
        return "blank"

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
            "comparison_cards": self._shapes_two_column,
        }
        return dispatch.get(vt, self._shapes_bullets)(slide)

    # ═══════════════════════════════════════════════
    # HEADER BLOCK — title + subtitle + accent line + side accent bar
    # Every content slide gets ~5 shapes just for the header
    # ═══════════════════════════════════════════════

    def _add_header(self, shapes: list[ShapeSpec], slide: OptimizedSlideContent) -> float:
        # Left accent bar (vertical colored strip along title)
        shapes.append(ShapeSpec(
            shape_type="rectangle",
            position=Position(left=0.15, top=GRID.TITLE_TOP, width=0.07, height=0.58),
            fill_color="accent1",
        ))

        # Title
        shapes.append(ShapeSpec(
            shape_type="text_box",
            position=Position(left=GRID.TITLE_LEFT, top=GRID.TITLE_TOP,
                              width=GRID.TITLE_WIDTH, height=GRID.TITLE_HEIGHT),
            text=slide.title, font_size=28, font_bold=True,
            font_color="dk1", alignment="left",
        ))

        y = GRID.TITLE_TOP + GRID.TITLE_HEIGHT + 0.05

        # Key message subtitle
        if slide.key_message:
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=GRID.TITLE_LEFT, top=y,
                                  width=GRID.TITLE_WIDTH, height=0.30),
                text=slide.key_message, font_size=13, font_color="dk2", alignment="left",
            ))
            y += 0.32

        # Full-width accent divider line
        shapes.append(ShapeSpec(
            shape_type="line",
            position=Position(left=GRID.TITLE_LEFT, top=y, width=GRID.TITLE_WIDTH, height=0),
            font_color="accent1", border_width=1.5,
        ))

        # Footer accent line at bottom
        shapes.append(ShapeSpec(
            shape_type="line",
            position=Position(left=GRID.TITLE_LEFT, top=7.05, width=GRID.TITLE_WIDTH, height=0),
            font_color="lt2", border_width=0.75,
        ))

        # Footer context text (like sample: "Title | Category")
        shapes.append(ShapeSpec(
            shape_type="text_box",
            position=Position(left=GRID.TITLE_LEFT, top=7.10, width=8.0, height=0.25),
            text=slide.title, font_size=8, font_color="dk2", alignment="left",
        ))

        # Topic icon in top-right corner for visual identity
        topic_icon = _pick_icon(slide.title)
        if topic_icon != _DEFAULT_ICON:
            self._icon_marker(shapes, GRID.SLIDE_WIDTH - 1.0, GRID.TITLE_TOP + 0.05,
                              0.48, topic_icon, "accent2")

        return y + 0.15

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

    def _accent_bar(self, shapes, left, top, height, color="accent1"):
        shapes.append(ShapeSpec(
            shape_type="rectangle",
            position=Position(left=left, top=top, width=0.06, height=height),
            fill_color=color,
        ))

    def _icon_marker(self, shapes, left, top, size, icon_text, color="accent1"):
        """Add a Unicode icon inside a colored circle as a visual marker."""
        shapes.append(ShapeSpec(
            shape_type="oval",
            position=Position(left=left, top=top, width=size, height=size),
            text=icon_text, font_size=int(size * 22), font_bold=False,
            font_color="lt1", fill_color=color,
            alignment="center", vertical_alignment="middle",
        ))

    def _circle(self, shapes, left, top, size, color, text="", font_size=12):
        shapes.append(ShapeSpec(
            shape_type="oval",
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
                      fill_color="accent1"),
            # Title
            ShapeSpec(shape_type="text_box",
                      position=Position(left=0.375, top=3.1, width=9.3, height=0.7),
                      text=slide.title, font_size=36, font_bold=True,
                      alignment="left", vertical_alignment="middle"),
            # Accent line under title
            ShapeSpec(shape_type="line",
                      position=Position(left=0.375, top=3.9, width=4.0, height=0),
                      font_color="accent1", border_width=2.0),
            # Subtitle
            ShapeSpec(shape_type="text_box",
                      position=Position(left=0.375, top=4.2, width=6.6, height=0.8),
                      text=slide.subtitle or "", font_size=16, font_color="dk2",
                      alignment="left", vertical_alignment="top"),
            # Decorative small squares at bottom
            ShapeSpec(shape_type="rectangle",
                      position=Position(left=0.375, top=6.8, width=0.15, height=0.15),
                      fill_color="accent1"),
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
        shapes = []
        y = self._add_header(shapes, slide)  # ~5 shapes

        bullets = (slide.bullets or slide.agenda_items or [])[:5]
        if not bullets:
            return shapes

        left, width = self.grid.span(0, 11)
        item_h = 0.60
        item_top = y + GAP

        # Pick a slide-level icon from the title
        slide_icon = _pick_icon(slide.title)

        for i, bullet in enumerate(bullets):
            by = item_top + i * (item_h + 0.18)
            accent = f"accent{(i % 6) + 1}"

            # Domain-aware icon marker (falls back to numbered if no match)
            bullet_icon = _pick_icon(bullet)
            if bullet_icon == _DEFAULT_ICON:
                bullet_icon = slide_icon if slide_icon != _DEFAULT_ICON else str(i + 1)
            self._icon_marker(shapes, left, by + 0.05, 0.42, bullet_icon, accent)

            # Bullet text
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=left + 0.55, top=by,
                                  width=width - 0.65, height=item_h),
                text=bullet, font_size=13, font_color="dk2",
                alignment="left", vertical_alignment="middle",
            ))

            # Separator line between bullets
            if i < len(bullets) - 1:
                self._line_h(shapes, left + 0.55, by + item_h + 0.07, width - 0.65)

        # Right-side decorative accent block
        rl, rw = self.grid.span(11, 1)
        bullets_height = len(bullets) * (item_h + 0.18) - 0.1
        shapes.append(ShapeSpec(
            shape_type="rounded_rect",
            position=Position(left=rl + 0.1, top=item_top, width=rw - 0.1, height=bullets_height),
            fill_color="accent1",
        ))

        # ── Auto-fill: if bullets use less than ~60% of vertical space, add
        #    a summary/highlight box below to fill the gap ──
        used_bottom = item_top + bullets_height + 0.3
        available_bottom = 6.8  # above footer
        remaining = available_bottom - used_bottom

        if remaining > 0.9 and slide.key_message:
            fill_left, fill_w = self.grid.span(1, 10)
            # Key takeaway highlight box
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=fill_left, top=used_bottom,
                                  width=fill_w, height=min(remaining - 0.2, 0.80)),
                fill_color="lt2",
            ))
            self._accent_bar(shapes, fill_left, used_bottom, min(remaining - 0.2, 0.80), "accent1")
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=fill_left + 0.20, top=used_bottom + 0.08,
                                  width=fill_w - 0.30, height=min(remaining - 0.35, 0.65)),
                text=f"✦  {slide.key_message}", font_size=12, font_bold=True,
                font_color="dk1", alignment="left", vertical_alignment="middle",
            ))

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
            position=Position(left=13.10, top=y + GAP, width=0.06, height=chart_h + 0.2),
            fill_color="accent2",
        ))

        # Horizontal line above chart
        self._line_h(shapes, chart_left, y + GAP, chart_w, "lt2", 0.5)
        # Horizontal line below chart
        self._line_h(shapes, chart_left, y + GAP + chart_h + 0.15, chart_w, "accent1", 1.0)

        # Small decorative squares at bottom-right
        self._diamond(shapes, 12.5, 6.6, 0.12, "accent1")
        self._diamond(shapes, 12.7, 6.6, 0.12, "accent3")

        # Auto-fill: key message callout below chart if space allows
        chart_bottom = y + GAP + chart_h + 0.25
        if chart_bottom < 6.2 and slide.key_message:
            kl, kw = self.grid.span(1, 10)
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=kl, top=chart_bottom + 0.10, width=kw, height=0.45),
                text=f"✦  {slide.key_message}", font_size=11, font_bold=True,
                font_color="dk1", alignment="left", vertical_alignment="middle",
            ))
            self._line_h(shapes, kl, chart_bottom + 0.05, kw, "accent1", 1.0)

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
        self._accent_bar(shapes, left - 0.15, y + GAP, table_h, "accent1")

        shapes.append(ShapeSpec(
            shape_type="table",
            position=Position(left=left, top=y + GAP, width=width, height=table_h),
        ))

        # Decorative lines around table
        self._line_h(shapes, left, y + GAP + table_h + 0.05, width, "accent1", 1.0)

        # Auto-fill: key message below table if space allows
        table_bottom = y + GAP + table_h + 0.15
        if table_bottom < 6.2 and slide.key_message:
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=left, top=table_bottom + 0.10, width=width, height=0.45),
                text=f"✦  {slide.key_message}", font_size=11, font_bold=True,
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

        card_h = 2.5
        card_top = y + GAP + 0.1
        accents = ["accent1", "accent2", "accent3", "accent4"]

        for i in range(n):
            if n <= 3:
                left, width = self.grid.span(i * 4, 4)
            else:
                left, width = self.grid.quarter(i)

            accent = accents[i % len(accents)]
            cl = left + PAD
            cw = width - PAD * 2

            # Card background
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=cl, top=card_top, width=cw, height=card_h),
                fill_color="lt2",
            ))

            # Accent top bar (thin colored strip)
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=cl, top=card_top, width=cw, height=0.07),
                fill_color=accent,
            ))

            # Domain icon or accent circle with KPI context
            kpi_icon = _pick_icon(kpis[i].label)
            self._icon_marker(shapes, cl + cw / 2 - 0.25, card_top + 0.20, 0.50, kpi_icon, accent)

            # Large KPI value
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl + 0.1, top=card_top + 0.80, width=cw - 0.2, height=0.70),
                text=kpis[i].value, font_size=28, font_bold=True,
                font_color="dk1", alignment="center", vertical_alignment="middle",
            ))

            # Separator line
            self._line_h(shapes, cl + 0.3, card_top + 1.55, cw - 0.6, accent, 1.0)

            # Label
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl + 0.1, top=card_top + 1.65, width=cw - 0.2, height=0.70),
                text=kpis[i].label, font_size=10, font_color="dk2",
                alignment="center", vertical_alignment="top",
            ))

            # Vertical divider between cards
            if i < n - 1:
                self._line_v(shapes, left + width, card_top + 0.2, card_h - 0.4)

        # Supplementary bullets below
        if slide.bullets:
            bleft, bwidth = self.grid.span(1, 10)
            bullet_text = '\n'.join(f'• {b}' for b in slide.bullets[:3])
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=bleft, top=card_top + card_h + GAP * 2,
                                  width=bwidth, height=1.0),
                text=bullet_text, font_size=11, font_color="dk2",
            ))

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

        step_h = 1.6
        step_top = y + 1.0
        arrow_w = 0.5
        total_w = GRID.content_width
        step_w = (total_w - (n - 1) * arrow_w) / n

        for i in range(n):
            x = GRID.content_left + i * (step_w + arrow_w)
            accent = f"accent{(i % 6) + 1}"

            # Step icon circle on top (domain-aware, fallback to number)
            step_icon = _pick_icon(steps[i].label)
            if step_icon == _DEFAULT_ICON:
                step_icon = str(i + 1)
            self._icon_marker(shapes, x + step_w / 2 - 0.22, step_top - 0.15, 0.44, step_icon, accent)

            # Step box
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=x, top=step_top + 0.40, width=step_w - 0.1, height=step_h - 0.5),
                text=steps[i].label, font_size=11, font_bold=True,
                font_color="dk1", fill_color="lt2",
                alignment="center", vertical_alignment="middle",
            ))

            # Bottom accent line under each box
            self._line_h(shapes, x, step_top + step_h - 0.05, step_w - 0.1, accent, 2.0)

            # Arrow connector
            if i < n - 1:
                shapes.append(ShapeSpec(
                    shape_type="arrow",
                    position=Position(left=x + step_w - 0.05, top=step_top + step_h / 2,
                                      width=arrow_w, height=0.22),
                    font_color="accent1",
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
        line_y = y + 2.0

        # Main horizontal line (thick accent)
        self._line_h(shapes, left, line_y, width, "accent1", 3.0)

        spacing = width / (n + 1)
        for i in range(n):
            x = left + spacing * (i + 1)
            accent = f"accent{(i % 6) + 1}"

            # Milestone marker circle
            self._circle(shapes, x - 0.22, line_y - 0.22, 0.44, accent)

            # Connecting vertical line from circle to label
            self._line_v(shapes, x, line_y - 0.85, 0.60, accent, 1.5)

            # Date above
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=x - 0.65, top=line_y - 1.3, width=1.3, height=0.40),
                text=items[i].date, font_size=12, font_bold=True,
                font_color="lt1", fill_color=accent,
                alignment="center", vertical_alignment="middle",
            ))

            # Description below
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=x - 0.8, top=line_y + 0.45, width=1.6, height=0.9),
                text=items[i].label, font_size=10, font_color="dk2", alignment="center",
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
        col_h = 3.3
        col_top = y + GAP

        left_title = "Key Points"
        right_title = "Details"
        if slide.comparison_items and len(slide.comparison_items) >= 2:
            left_title = slide.comparison_items[0].title
            right_title = slide.comparison_items[1].title

        left_bullets = (slide.left_column or [])[:4]
        right_bullets = (slide.right_column or [])[:4]

        for col_idx, (cl, cw, title, bullets) in enumerate([
            (ll, lw, left_title, left_bullets),
            (rl, rw, right_title, right_bullets),
        ]):
            accent = f"accent{col_idx + 1}"

            # Domain icon in column header
            col_icon = _pick_icon(title)
            self._icon_marker(shapes, cl + PAD + 0.08, col_top + 0.05, 0.40, col_icon, accent)

            # Column header bar with icon offset
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=cl + PAD, top=col_top, width=cw - PAD * 2, height=0.50),
                text=f"  {title}", font_size=14, font_bold=True,
                font_color="lt1", fill_color=accent,
                alignment="center", vertical_alignment="middle",
            ))

            # Column body background
            shapes.append(ShapeSpec(
                shape_type="rounded_rect",
                position=Position(left=cl + PAD, top=col_top + 0.58, width=cw - PAD * 2, height=col_h - 0.65),
                fill_color="lt2",
            ))

            # Individual bullet items with accent dots
            for j, b in enumerate(bullets):
                by = col_top + 0.70 + j * 0.65

                # Accent dot
                self._circle(shapes, cl + PAD + 0.12, by + 0.12, 0.16, accent)

                # Bullet text
                shapes.append(ShapeSpec(
                    shape_type="text_box",
                    position=Position(left=cl + PAD + 0.38, top=by,
                                      width=cw - PAD * 2 - 0.45, height=0.55),
                    text=b, font_size=11, font_color="dk2",
                    alignment="left", vertical_alignment="middle",
                ))

                # Separator line
                if j < len(bullets) - 1:
                    self._line_h(shapes, cl + PAD + 0.38, by + 0.58, cw - PAD * 2 - 0.45)

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
        body_h = 3.2
        col_top = y + GAP

        for i in range(n):
            left, width = self.grid.third(i)
            item = items[i]
            accent = f"accent{(i % 6) + 1}"
            cl = left + PAD
            cw = width - PAD * 2

            # Accent top strip
            shapes.append(ShapeSpec(
                shape_type="rectangle",
                position=Position(left=cl, top=col_top, width=cw, height=0.07),
                fill_color=accent,
            ))

            # Domain icon above column header
            col_icon = _pick_icon(item.title)
            self._icon_marker(shapes, cl + cw / 2 - 0.20, col_top + 0.14, 0.40, col_icon, accent)

            # Column header (shifted down to accommodate icon)
            shapes.append(ShapeSpec(
                shape_type="text_box",
                position=Position(left=cl, top=col_top + 0.58, width=cw, height=0.35),
                text=item.title, font_size=13, font_bold=True,
                font_color="dk1", alignment="center", vertical_alignment="middle",
            ))

            # Header underline
            self._line_h(shapes, cl + 0.15, col_top + 0.96, cw - 0.30, accent, 1.0)

            # Body points with accent dots and separators
            points = item.points[:3]
            for j, point in enumerate(points):
                py = col_top + 1.10 + j * 0.70

                # Accent dot
                self._circle(shapes, cl + 0.08, py + 0.10, 0.14, accent)

                # Point text
                shapes.append(ShapeSpec(
                    shape_type="text_box",
                    position=Position(left=cl + 0.30, top=py, width=cw - 0.35, height=0.60),
                    text=point, font_size=11, font_color="dk2",
                    alignment="left", vertical_alignment="top",
                ))

                # Separator between points
                if j < len(points) - 1:
                    self._line_h(shapes, cl + 0.10, py + 0.68, cw - 0.20)

            # Vertical divider between columns
            if i < n - 1:
                self._line_v(shapes, left + width, col_top + 0.1, body_h)

        return shapes
