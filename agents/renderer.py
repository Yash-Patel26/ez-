"""
Agent 6: PPTX Renderer

Assembles the final .pptx file by combining layouts, optimized content,
and visual elements with the Slide Master template.

Responsibilities:
  - Load Slide Master template as base presentation
  - Add slides using correct layouts
  - Place text, shapes, charts, and tables via VisualGenerator
  - Apply theme fonts and colors consistently
  - Handle cover and thank-you slides via template placeholders

Input:  List[SlideLayout] + List[OptimizedSlideContent] + template path
Output: Saved .pptx file
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

from agents.visual_generator import VisualGenerator, _resolve_color
from core.models import (
    OptimizedSlideContent,
    ShapeSpec,
    SlideLayout,
    ThemeConfig,
)


class PPTXRenderer:
    """Render the final PPTX presentation."""

    def __init__(
        self,
        template_path: str,
        theme: ThemeConfig,
        config: dict | None = None,
    ):
        self.template_path = template_path
        self.theme = theme
        self.config = config or {}
        self.visual_gen = VisualGenerator(theme, config)

    def render(
        self,
        layouts: list[SlideLayout],
        slides_content: list[OptimizedSlideContent],
        output_path: str,
    ) -> None:
        """
        Render all slides and save the presentation.

        Args:
            layouts: Computed layout specifications.
            slides_content: Optimized content for each slide.
            output_path: Where to save the .pptx file.
        """
        prs = Presentation(self.template_path)

        # Remove existing slides from template (keep only layouts/master)
        self._remove_existing_slides(prs)

        # Build a layout name → layout object mapping
        layout_map = self._build_layout_map(prs)

        # Add each slide
        for layout_spec, content in zip(layouts, slides_content):
            pptx_layout = self._find_layout(layout_map, layout_spec.layout_name)
            slide = prs.slides.add_slide(pptx_layout)

            # Render based on slide type
            if content.slide_type == "cover":
                self._render_cover(slide, content, pptx_layout)
            elif content.slide_type == "thank_you":
                self._render_thank_you(slide, content)
            else:
                self._render_content_slide(slide, layout_spec, content)

        # Ensure output directory exists
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        prs.save(output_path)

    def _remove_existing_slides(self, prs: Presentation) -> None:
        """Remove all existing slides from the presentation."""
        from lxml import etree
        ns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

        while len(prs.slides) > 0:
            sldId = prs.slides._sldIdLst[0]
            rId = sldId.get(f'{ns}id') or sldId.get('r:id')
            if rId:
                try:
                    prs.part.drop_rel(rId)
                except KeyError:
                    pass
            prs.slides._sldIdLst.remove(sldId)

    def _build_layout_map(self, prs: Presentation) -> dict[str, object]:
        """Build a mapping from layout name to layout object."""
        layout_map = {}
        for layout in prs.slide_layouts:
            layout_map[layout.name.lower()] = layout
            # Also map by index
            layout_map[f"idx_{prs.slide_layouts.index(layout)}"] = layout
        return layout_map

    def _find_layout(self, layout_map: dict, layout_name: str) -> object:
        """Find the best matching layout by name."""
        name_lower = layout_name.lower()

        # Direct match
        if name_lower in layout_map:
            return layout_map[name_lower]

        # Try common names based on type
        type_names = {
            "cover": ["1_cover", "2_cover", "cover", "0_title company",
                       "title company", "title slide"],
            "divider": ["divider", "c_section blue"],
            "blank": ["blank", "1_e_title, subtitle and body"],
            "title_only": ["title only", "1_e_title, subtitle and body"],
            "thank_you": ["thank you", "1_thank you", "0_title company",
                          "title company"],
        }

        if name_lower in type_names:
            for candidate in type_names[name_lower]:
                if candidate in layout_map:
                    return layout_map[candidate]

        # Fuzzy match
        for key in layout_map:
            if name_lower in key or key in name_lower:
                return layout_map[key]

        # Fallback: try title-related, then first layout with placeholders, then first
        for fallback in ["blank", "title only", "1_e_title, subtitle and body"]:
            if fallback in layout_map:
                return layout_map[fallback]

        # Last resort: first layout
        return list(layout_map.values())[0]

    def _render_cover(
        self, slide, content: OptimizedSlideContent, layout
    ) -> None:
        """Render the cover slide using template placeholders."""
        # Try to use template placeholders first
        placeholders = list(slide.placeholders)

        if len(placeholders) >= 2:
            placeholders[0].text = content.title or ""
            if content.subtitle and len(placeholders) > 1:
                placeholders[1].text = content.subtitle or ""

            # Style + zero margins on unfilled placeholders (Common Mistake #14)
            for ph in placeholders:
                ph.text_frame.margin_left = 0
                ph.text_frame.margin_right = 0
                ph.text_frame.margin_top = 0
                ph.text_frame.margin_bottom = 0
                for para in ph.text_frame.paragraphs:
                    para.font.name = self.theme.fonts.major
        elif len(placeholders) == 1:
            placeholders[0].text = content.title or ""
        else:
            # No placeholders — add text boxes
            for spec in content.bullets or []:
                pass  # Handled by shapes below

            # Add title as text box
            txBox = slide.shapes.add_textbox(
                Inches(0.375), Inches(3.1), Inches(9.3), Inches(0.7)
            )
            tf = txBox.text_frame
            tf.paragraphs[0].text = content.title or ""
            tf.paragraphs[0].font.size = Pt(36)
            tf.paragraphs[0].font.bold = True
            tf.paragraphs[0].font.name = self.theme.fonts.major
            tf.paragraphs[0].font.color.rgb = _resolve_color("dk1", self.theme)

            if content.subtitle:
                txBox2 = slide.shapes.add_textbox(
                    Inches(0.375), Inches(4.3), Inches(6.6), Inches(0.8)
                )
                tf2 = txBox2.text_frame
                tf2.paragraphs[0].text = content.subtitle
                tf2.paragraphs[0].font.size = Pt(16)
                tf2.paragraphs[0].font.name = self.theme.fonts.minor
                tf2.paragraphs[0].font.color.rgb = _resolve_color("dk2", self.theme)

    def _render_thank_you(self, slide, content: OptimizedSlideContent) -> None:
        """Render the thank you/closing slide.

        Uses template placeholders if available; adds text boxes as fallback
        so the slide is never completely empty.
        """
        placeholders = list(slide.placeholders)

        if placeholders:
            # Use template placeholders
            placeholders[0].text = "Thank You"
            for ph in placeholders:
                ph.text_frame.margin_left = 0
                ph.text_frame.margin_right = 0
                for para in ph.text_frame.paragraphs:
                    para.font.name = self.theme.fonts.major
        else:
            # Fallback: add "Thank You" text box so slide is never blank
            txBox = slide.shapes.add_textbox(
                Inches(2.0), Inches(2.5), Inches(9.3), Inches(1.0)
            )
            tf = txBox.text_frame
            tf.paragraphs[0].text = "Thank You"
            tf.paragraphs[0].font.size = Pt(44)
            tf.paragraphs[0].font.bold = True
            tf.paragraphs[0].font.name = self.theme.fonts.major
            tf.paragraphs[0].font.color.rgb = _resolve_color("dk1", self.theme)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    def _render_content_slide(
        self, slide, layout: SlideLayout, content: OptimizedSlideContent
    ) -> None:
        """Render a content slide with shapes, charts, tables.

        Guideline 2B: Title uses MAJOR font, body uses MINOR font.
        Guideline 4A: Max 2 fonts, consistent sizing.
        """
        # Render each shape from the layout
        for spec in layout.shapes:
            # Determine if this is a title shape (Guideline 2B: use major font)
            is_title = (spec.font_size and spec.font_size >= 28
                        and spec.position.top < 1.0)

            if spec.shape_type == "chart":
                if content.chart_data:
                    self.visual_gen.add_chart(slide, content.chart_data, spec)
            elif spec.shape_type == "table":
                if content.table_data:
                    self.visual_gen.add_table(slide, content.table_data, spec)
            else:
                # Render generic shape
                self.visual_gen.add_shape(slide, spec, self.theme)
