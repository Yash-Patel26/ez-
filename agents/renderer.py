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
                self._render_cover(slide, content, layout_spec)
            elif content.slide_type == "thank_you":
                self._render_thank_you(slide, content)
            elif content.slide_type == "divider":
                self._render_divider(slide, layout_spec, content)
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
        """Build a mapping from layout name to layout object.

        Templates often have duplicate layout names (e.g. three copies of
        '1_E_Title, Subtitle and Body', where the last one is a 'Thank you!'
        variant with that text baked into its shapes).  Overwriting the map
        with the last duplicate would inject 'Thank you!' onto every content
        slide.  Fix: store only the FIRST occurrence under the plain name, and
        separately capture any layout whose shapes contain 'thank you' text so
        the closing slide can find it explicitly.
        """
        layout_map = {}
        for layout in prs.slide_layouts:
            idx_key = f"idx_{prs.slide_layouts.index(layout)}"
            layout_map[idx_key] = layout

            name_lower = layout.name.lower()

            # Detect layouts that have 'thank you' text baked into their shapes
            has_thankyou_text = any(
                "thank you" in sh.text_frame.text.lower()
                for sh in layout.shapes
                if sh.has_text_frame
            )
            if has_thankyou_text:
                # Store under a special key for the closing slide, but do NOT
                # overwrite the clean content-layout entry under the plain name
                layout_map.setdefault("_thankyou_layout", layout)
                continue

            # Only store the first occurrence under the plain name
            layout_map.setdefault(name_lower, layout)

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
            # Prefer the layout that already has 'thank you' text baked in
            "thank_you": ["_thankyou_layout", "thank you", "1_thank you",
                          "0_title company", "title company"],
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
        """Render the cover slide using template placeholders + decorative shapes."""
        placeholders = list(slide.placeholders)

        if len(placeholders) >= 2:
            placeholders[0].text = content.title or ""
            if content.subtitle:
                placeholders[1].text = content.subtitle or ""
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

        # Add the layout engine's decorative shapes on top of placeholders.
        # Skip pure text-box shapes (title/subtitle already handled above)
        # to avoid duplicates; only add non-text decorative elements.
        for spec in getattr(layout, 'shapes', []):
            if spec.shape_type == "text_box":
                continue  # title/subtitle from placeholders already
            self.visual_gen.add_shape(slide, spec, self.theme)

    def _render_thank_you(self, slide, content: OptimizedSlideContent) -> None:
        """Render the thank you/closing slide.

        If the template layout already has 'Thank you' text baked into its
        shapes, we leave it alone — adding text would create a duplicate.
        Only populate placeholders or add a fallback text box when the layout
        has no such pre-existing text.
        """
        # Check if the layout already has "thank you" text baked in
        layout_has_thankyou = any(
            "thank you" in sh.text_frame.text.lower()
            for sh in slide.slide_layout.shapes
            if sh.has_text_frame
        )
        if layout_has_thankyou:
            # Template provides the text — nothing to add
            return

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

    def _render_divider(
        self, slide, layout: SlideLayout, content: OptimizedSlideContent
    ) -> None:
        """Render a section-divider slide.

        The C_Section blue template layout provides the dark background.
        We only add our decorative shapes on top — no placeholder text is set
        so the template's own look is preserved cleanly.
        """
        for spec in layout.shapes:
            self.visual_gen.add_shape(slide, spec, self.theme)

    def _populate_placeholders(self, slide, title: str, subtitle: str = "") -> None:
        """Set title and subtitle template placeholders using the template's own fonts.

        Using placeholders ensures the slide master's font, size, and positioning
        are respected — giving consistent rendering across templates.
        """
        for ph in slide.placeholders:
            try:
                idx = ph.placeholder_format.idx
            except Exception:
                continue

            if idx == 0 and title:
                ph.text = title
                tf = ph.text_frame
                tf.margin_left = 0
                tf.margin_top = 0
                for para in tf.paragraphs:
                    para.font.name = self.theme.fonts.major
                    para.font.bold = True

            elif idx == 1 and subtitle:
                ph.text = subtitle
                tf = ph.text_frame
                tf.margin_left = 0
                tf.margin_top = 0
                for para in tf.paragraphs:
                    para.font.name = self.theme.fonts.minor
                    para.font.size = Pt(13)

    def _render_content_slide(
        self, slide, layout: SlideLayout, content: OptimizedSlideContent
    ) -> None:
        """Render a content slide with shapes, charts, tables.

        Title goes into the template's PH0 (preserves template font/positioning).
        Key message goes into PH1 as subtitle.
        All decorative and content shapes are then added on top.
        """
        # Populate template placeholders first so the slide title uses
        # the template's own font size, color, and position.
        self._populate_placeholders(slide, content.title, content.key_message)

        # Render each shape from the layout
        for spec in layout.shapes:
            if spec.shape_type == "chart":
                if content.chart_data:
                    self.visual_gen.add_chart(slide, content.chart_data, spec)
            elif spec.shape_type == "table":
                if content.table_data:
                    self.visual_gen.add_table(slide, content.table_data, spec)
            else:
                self.visual_gen.add_shape(slide, spec, self.theme)
