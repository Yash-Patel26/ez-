"""
Agent 3: Content Optimizer

Transforms verbose markdown content into concise, slide-ready content.
Uses Claude API for intelligent condensation, with rule-based fallback.

Responsibilities:
  - Condense paragraphs into 4-6 bullet points (max 15 words each)
  - Extract chart data from tables and numeric content
  - Identify KPI values, process steps, timeline items, comparisons
  - Apply "infographic-first" principle: visualize before text
  - Strip citations and clean up content for presentation

Input:  SlidePlan + DocumentIR
Output: List[OptimizedSlideContent]
"""

from __future__ import annotations

import json
import os
import re
from typing import Any

from core.models import (
    ChartData,
    ComparisonItem,
    ContentBlock,
    DocumentIR,
    FunnelItem,
    KPIItem,
    OptimizedSlideContent,
    ProcessStep,
    Section,
    SeriesData,
    SlideSpec,
    SlidePlan,
    TableData,
    TimelineItem,
)


# Regex for extracting KPI-like values from text
_KPI_PATTERNS = [
    re.compile(r'(\$[\d,.]+\s*[BMKTbmkt](?:illion|rillion)?)', re.IGNORECASE),
    re.compile(r'([\d,.]+\s*%)', re.IGNORECASE),
    re.compile(r'((?:USD|EUR|INR|₹|€|£)\s*[\d,.]+\s*[BMKTbmkt]?)', re.IGNORECASE),
    re.compile(r'([\d,.]+\s*(?:MW|GW|TWh|GWh))', re.IGNORECASE),
    re.compile(r'([\d,.]+\s*(?:billion|million|trillion|crore|lakh))', re.IGNORECASE),
]

_CITATION_PATTERN = re.compile(r'\s*\[\d+\]\s*')
# Markdown formatting patterns: **bold**, *italic*, __bold__, _italic_, ##heading
_MARKDOWN_PATTERN = re.compile(
    r'\*\*(.+?)\*\*'      # **bold**
    r'|\*(.+?)\*'          # *italic*
    r'|__(.+?)__'          # __bold__
    r'|_(.+?)_'            # _italic_
    r'|`(.+?)`'            # `code`
    r'|#{1,6}\s+'          # ## headings
    r'|\*\*\*(.+?)\*\*\*'  # ***bold-italic***
)


def _strip_citations(text: str) -> str:
    return _CITATION_PATTERN.sub(' ', text).strip()


def _strip_markdown(text: str) -> str:
    """Strip markdown formatting from text, keeping the plain content."""
    if not text:
        return text
    # Replace **bold**, *italic*, etc. with their plain content
    def _replace(m: re.Match) -> str:
        # Return whichever capture group matched
        return next((g for g in m.groups() if g is not None), '')
    text = re.sub(
        r'\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|__(.+?)__|_(.+?)_|`(.+?)`',
        _replace, text
    )
    # Strip heading markers (## Title → Title)
    text = re.sub(r'^#{1,6}\s+', '', text, flags=re.MULTILINE)
    # Strip horizontal rules
    text = re.sub(r'^[-*_]{3,}\s*$', '', text, flags=re.MULTILINE)
    return text.strip()


def _truncate_bullet(text: str, max_words: int = 15) -> str:
    """Truncate a bullet to max_words, keeping meaningful content."""
    text = _strip_markdown(_strip_citations(text))
    words = text.split()
    if len(words) <= max_words:
        return text
    return ' '.join(words[:max_words]) + '...'


class ContentOptimizer:
    """Optimize document content for slide presentation."""

    def __init__(self, config: dict | None = None):
        self.config = config or {}
        self.content_config = self.config.get("content", {})
        self.max_bullets = self.content_config.get("max_bullets_per_slide", 5)
        self.max_words = self.content_config.get("max_words_per_bullet", 12)
        self.max_table_rows = self.content_config.get("max_table_rows", 6)
        self.api_key = os.environ.get("GEMINI_API_KEY", "")
        self.llm_config = self.config.get("llm", {})

    def optimize(
        self, plan: SlidePlan, doc: DocumentIR
    ) -> list[OptimizedSlideContent]:
        """
        Optimize content for each slide in the plan.

        Tries LLM for intelligent condensation, falls back to rules.
        """
        results = []

        for slide_spec in plan.slides:
            # Gather source content
            source_sections = self._gather_sections(slide_spec, doc)
            optimized = self._optimize_slide(slide_spec, source_sections, doc)
            results.append(optimized)

        return results

    def _gather_sections(
        self, spec: SlideSpec, doc: DocumentIR
    ) -> list[Section]:
        """Collect sections referenced by a slide spec.

        If the slide title matches a subsection heading (expanded slide),
        return only that subsection for more targeted content.
        """
        sections = []
        for idx in spec.source_sections:
            if 0 <= idx < len(doc.sections):
                parent = doc.sections[idx]
                # Check if slide title matches a subsection
                matched_sub = None
                for sub in parent.subsections:
                    clean_sub = re.sub(r'^\d+(\.\d+)*\.?\s*', '', sub.heading).strip()
                    if (spec.title and
                        (clean_sub[:30].lower() in spec.title.lower()
                         or spec.title.lower() in clean_sub.lower())):
                        matched_sub = sub
                        break
                if matched_sub:
                    # Only use the subsection if it has actual content;
                    # otherwise fall back to the parent so bullets aren't empty.
                    if matched_sub.content_blocks or matched_sub.subsections:
                        sections.append(matched_sub)
                    else:
                        sections.append(parent)
                else:
                    sections.append(parent)
        return sections

    def _optimize_slide(
        self,
        spec: SlideSpec,
        sections: list[Section],
        doc: DocumentIR,
    ) -> OptimizedSlideContent:
        """Optimize content for a single slide based on its type and treatment."""

        # Fixed slide types that don't need content optimization
        if spec.slide_type == "cover":
            return OptimizedSlideContent(
                slide_number=spec.slide_number,
                slide_type=spec.slide_type,
                title=doc.title,
                subtitle=doc.subtitle,
                visual_treatment=spec.visual_treatment,
                key_message=spec.key_message,
            )

        if spec.slide_type == "thank_you":
            return OptimizedSlideContent(
                slide_number=spec.slide_number,
                slide_type=spec.slide_type,
                title="Thank You",
                visual_treatment=spec.visual_treatment,
            )

        if spec.slide_type == "agenda":
            agenda_items = self._build_agenda(doc)
            return OptimizedSlideContent(
                slide_number=spec.slide_number,
                slide_type=spec.slide_type,
                title="Agenda",
                visual_treatment="bullets",
                agenda_items=agenda_items,
                bullets=agenda_items,
            )

        # Content-bearing slides
        all_blocks = self._collect_blocks(sections)
        all_text = self._collect_text(sections)

        # Dispatch based on visual treatment
        treatment = spec.visual_treatment

        if treatment in ("chart_bar", "chart_pie", "chart_line", "chart_area"):
            return self._optimize_chart(spec, sections, all_blocks, all_text)

        if treatment == "table":
            return self._optimize_table(spec, sections, all_blocks, all_text)

        if treatment == "kpi_cards":
            return self._optimize_kpi(spec, sections, all_text)

        if treatment == "process_flow":
            return self._optimize_process(spec, sections, all_text)

        if treatment == "timeline":
            return self._optimize_timeline(spec, sections, all_text)

        if treatment == "comparison_cards":
            return self._optimize_comparison(spec, sections, all_text, treatment="comparison_cards")

        if treatment == "two_column":
            return self._optimize_comparison(spec, sections, all_text)

        if treatment == "three_column":
            return self._optimize_three_column(spec, sections, all_text)

        if treatment == "icon_grid":
            return self._optimize_icon_grid(spec, sections, all_text)

        if treatment == "funnel":
            return self._optimize_funnel(spec, sections, all_text)

        if treatment == "divider_layout":
            return OptimizedSlideContent(
                slide_number=spec.slide_number,
                slide_type=spec.slide_type,
                title=spec.title,
                visual_treatment="divider_layout",
                key_message=spec.key_message,
            )

        # Default: bullets
        return self._optimize_bullets(spec, sections, all_text)

    # ──────────────────────────────────────────────
    # Optimization strategies by visual treatment
    # ──────────────────────────────────────────────

    def _llm_condense(self, text: str, max_bullets: int, title: str) -> list[str] | None:
        """Use Gemini to intelligently condense text into concise bullets."""
        if not self.api_key or len(text) < 50:
            return None
        try:
            import google.generativeai as genai
            genai.configure(api_key=self.api_key)
            model = genai.GenerativeModel(
                self.llm_config.get("model", "gemini-2.0-flash")
            )

            prompt = f"""Condense this content into exactly {max_bullets} concise bullet points for a presentation slide titled "{title}".

RULES:
- Each bullet must be max 12 words
- Focus on key insights, numbers, and actionable points
- No filler words, no citations, no references
- Return ONLY the bullets, one per line, no numbering, no dashes

CONTENT:
{text[:2000]}"""

            response = model.generate_content(
                prompt,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.2,
                    max_output_tokens=500,
                ),
            )

            lines = [l.strip().lstrip('•-*0123456789. ')
                     for l in response.text.strip().split('\n')
                     if l.strip() and len(l.strip()) > 5]
            return lines[:max_bullets] if lines else None
        except Exception:
            return None

    def _optimize_bullets(
        self, spec: SlideSpec, sections: list[Section], all_text: str
    ) -> OptimizedSlideContent:
        """Extract the best bullet points from sections."""

        # Try LLM-powered condensation first
        llm_bullets = self._llm_condense(all_text, self.max_bullets, spec.title)
        if llm_bullets:
            return OptimizedSlideContent(
                slide_number=spec.slide_number,
                slide_type=spec.slide_type,
                title=spec.title,
                visual_treatment="bullets",
                key_message=spec.key_message,
                bullets=llm_bullets,
            )

        # Fallback: rule-based extraction
        bullets = []

        for section in sections:
            for block in self._iter_blocks(section):
                if block.type in ("bullet_list", "numbered_list") and block.items:
                    for item in block.items:
                        bullets.append(_truncate_bullet(item, self.max_words))
                elif block.type == "text" and block.content:
                    sentences = re.split(r'[.!?]\s+', block.content)
                    for sent in sentences[:2]:
                        sent = sent.strip()
                        if len(sent.split()) >= 4:
                            bullets.append(_truncate_bullet(sent, self.max_words))

        # Deduplicate and limit
        seen = set()
        unique = []
        for b in bullets:
            key = b.lower().strip()[:50]
            if key not in seen:
                seen.add(key)
                unique.append(b)

        # Last resort: if still empty, use section/subsection headings as bullets
        # so the slide always shows something rather than a blank body.
        if not unique:
            for section in sections:
                h = re.sub(r'^\d+(\.\d+)*\.?\s*', '', section.heading).strip()
                if h and h.lower() != (spec.title or "").lower():
                    unique.append(h)
                for sub in section.subsections[:4]:
                    sh = re.sub(r'^\d+(\.\d+)*\.?\s*', '', sub.heading).strip()
                    if sh and sh not in unique:
                        unique.append(sh)

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment="bullets",
            key_message=spec.key_message,
            bullets=unique[:self.max_bullets],
        )

    def _optimize_chart(
        self,
        spec: SlideSpec,
        sections: list[Section],
        all_blocks: list[ContentBlock],
        all_text: str,
    ) -> OptimizedSlideContent:
        """Extract chart data from tables or numeric content."""
        chart_data = None

        # Try to find a table with numeric data
        for block in all_blocks:
            if block.type == "table" and block.headers and block.rows:
                chart_data = self._table_to_chart(
                    block.headers, block.rows, spec.visual_treatment
                )
                if chart_data:
                    break

        # Fallback: create from bullet/text data if no table found
        if not chart_data:
            chart_data = self._extract_chart_from_text(all_text, spec.visual_treatment)

        # If still no chart data, fall back to bullets
        if not chart_data:
            return self._optimize_bullets(spec, sections, all_text)

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment=spec.visual_treatment,
            key_message=spec.key_message,
            chart_data=chart_data,
        )

    def _optimize_table(
        self,
        spec: SlideSpec,
        sections: list[Section],
        all_blocks: list[ContentBlock],
        all_text: str,
    ) -> OptimizedSlideContent:
        """Extract and clean table data."""
        for block in all_blocks:
            if block.type == "table" and block.headers and block.rows:
                # Limit rows and clean content
                rows = block.rows[:self.max_table_rows]
                headers = block.headers[:6]  # max 6 columns
                rows = [[_strip_markdown(_strip_citations(cell))[:50] for cell in row[:6]] for row in rows]

                # Pad rows to match header count
                for row in rows:
                    while len(row) < len(headers):
                        row.append("")

                table_data = TableData(
                    title=spec.title,
                    headers=headers,
                    rows=rows,
                )

                return OptimizedSlideContent(
                    slide_number=spec.slide_number,
                    slide_type=spec.slide_type,
                    title=spec.title,
                    visual_treatment="table",
                    key_message=spec.key_message,
                    table_data=table_data,
                )

        # Fallback to bullets
        return self._optimize_bullets(spec, sections, all_text)

    def _optimize_kpi(
        self, spec: SlideSpec, sections: list[Section], all_text: str
    ) -> OptimizedSlideContent:
        """Extract KPI values from text content."""
        kpis = []

        for section in sections:
            for block in self._iter_blocks(section):
                text = block.content or ' '.join(block.items or [])
                if not text:
                    continue

                for pattern in _KPI_PATTERNS:
                    for match in pattern.finditer(text):
                        value = match.group(1).strip()
                        # Snap context boundaries to whole words so we never
                        # return mid-word fragments like 'pacity contribute...'
                        raw_start = max(0, match.start() - 70)
                        # Move forward to the next space (word boundary)
                        if raw_start > 0:
                            sp = text.find(' ', raw_start)
                            raw_start = sp + 1 if sp != -1 else raw_start
                        end = min(len(text), match.end() + 70)
                        context = text[raw_start:end]

                        # Extract a label from surrounding words
                        label = self._extract_kpi_label(context, value)
                        if label and len(kpis) < 4:
                            # Determine trend
                            trend = None
                            context_lower = context.lower()
                            if any(w in context_lower for w in ('growth', 'increase', 'rise', 'up')):
                                trend = "up"
                            elif any(w in context_lower for w in ('decline', 'decrease', 'drop', 'down')):
                                trend = "down"

                            kpis.append(KPIItem(
                                value=value,
                                label=label,
                                trend=trend,
                            ))

        # Deduplicate by value
        seen_values = set()
        unique_kpis = []
        for kpi in kpis:
            if kpi.value not in seen_values:
                seen_values.add(kpi.value)
                unique_kpis.append(kpi)

        kpis = unique_kpis[:4]

        # If fewer than 2 KPIs, add bullets as supplement
        bullets = None
        if len(kpis) < 2:
            fallback = self._optimize_bullets(spec, sections, all_text)
            bullets = fallback.bullets

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment="kpi_cards" if len(kpis) >= 2 else "bullets",
            key_message=spec.key_message,
            kpi_values=kpis if kpis else None,
            bullets=bullets,
        )

    def _optimize_process(
        self, spec: SlideSpec, sections: list[Section], all_text: str
    ) -> OptimizedSlideContent:
        """Extract process steps from content."""
        steps = []
        for section in sections:
            for block in self._iter_blocks(section):
                if block.type in ("bullet_list", "numbered_list") and block.items:
                    for item in block.items[:5]:
                        clean = _truncate_bullet(item, 10)
                        steps.append(ProcessStep(label=clean))

        if len(steps) < 3:
            return self._optimize_bullets(spec, sections, all_text)

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment="process_flow",
            key_message=spec.key_message,
            process_steps=steps[:5],
        )

    def _optimize_timeline(
        self, spec: SlideSpec, sections: list[Section], all_text: str
    ) -> OptimizedSlideContent:
        """Extract timeline items from content."""
        items = []
        year_pattern = re.compile(r'\b((?:19|20)\d{2})\b')

        for section in sections:
            for block in self._iter_blocks(section):
                text = block.content or ' '.join(block.items or [])
                if not text:
                    continue
                matches = year_pattern.findall(text)
                for year in matches:
                    # Get surrounding context for label
                    idx = text.find(year)
                    context = text[max(0, idx-30):idx+len(year)+50].strip()
                    label = _truncate_bullet(context, 8)
                    items.append(TimelineItem(date=year, label=label))

        # Deduplicate by date
        seen = set()
        unique = []
        for item in items:
            if item.date not in seen:
                seen.add(item.date)
                unique.append(item)

        if len(unique) < 3:
            return self._optimize_bullets(spec, sections, all_text)

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment="timeline",
            key_message=spec.key_message,
            timeline_items=sorted(unique, key=lambda x: x.date)[:6],
        )

    def _optimize_comparison(
        self, spec: SlideSpec, sections: list[Section], all_text: str,
        treatment: str = "two_column",
    ) -> OptimizedSlideContent:
        """Create two-column comparison content."""
        left = []
        right = []

        if len(sections) >= 2:
            # Two sections → one per column
            for block in self._iter_blocks(sections[0]):
                if block.type in ("bullet_list", "numbered_list") and block.items:
                    left.extend(_truncate_bullet(i, self.max_words) for i in block.items)
                elif block.type == "text" and block.content:
                    left.append(_truncate_bullet(block.content, self.max_words))

            for block in self._iter_blocks(sections[1]):
                if block.type in ("bullet_list", "numbered_list") and block.items:
                    right.extend(_truncate_bullet(i, self.max_words) for i in block.items)
                elif block.type == "text" and block.content:
                    right.append(_truncate_bullet(block.content, self.max_words))
        elif sections:
            # One section with subsections
            subs = sections[0].subsections
            if len(subs) >= 2:
                for block in self._iter_blocks(subs[0]):
                    if block.items:
                        left.extend(_truncate_bullet(i, self.max_words) for i in block.items)
                    elif block.content:
                        left.append(_truncate_bullet(block.content, self.max_words))
                for block in self._iter_blocks(subs[1]):
                    if block.items:
                        right.extend(_truncate_bullet(i, self.max_words) for i in block.items)
                    elif block.content:
                        right.append(_truncate_bullet(block.content, self.max_words))

        left = [_truncate_bullet(b, 10) for b in left[:4]]
        right = [_truncate_bullet(b, 10) for b in right[:4]]

        if not left and not right:
            return self._optimize_bullets(spec, sections, all_text)

        # For comparison_cards, also populate comparison_items for VS-badge layout
        comp_items = None
        if treatment == "comparison_cards":
            comp_titles = ["Option A", "Option B"]
            subs = sections[0].subsections if sections and sections[0].subsections else []
            if len(subs) >= 2:
                comp_titles = [
                    re.sub(r'^\d+(\.\d+)*\.?\s*', '', subs[0].heading)[:30],
                    re.sub(r'^\d+(\.\d+)*\.?\s*', '', subs[1].heading)[:30],
                ]
            elif len(sections) >= 2:
                comp_titles = [sections[0].heading[:30], sections[1].heading[:30]]
            comp_items = [
                ComparisonItem(title=comp_titles[0], points=left or ["Key point"]),
                ComparisonItem(title=comp_titles[1], points=right or ["Key point"]),
            ]

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment=treatment,
            key_message=spec.key_message,
            left_column=left or ["No data available"],
            right_column=right or ["No data available"],
            comparison_items=comp_items,
        )

    def _optimize_three_column(
        self, spec: SlideSpec, sections: list[Section], all_text: str
    ) -> OptimizedSlideContent:
        """Create three-column content from subsections."""
        columns = [[], [], []]

        if sections:
            subs = sections[0].subsections if sections[0].subsections else sections[:3]
            for col_idx, sub in enumerate(subs[:3]):
                target = sub if isinstance(sub, Section) else sections[min(col_idx, len(sections)-1)]
                for block in self._iter_blocks(target):
                    if block.type in ("bullet_list", "numbered_list") and block.items:
                        columns[col_idx].extend(
                            _truncate_bullet(i, 8) for i in block.items[:2]
                        )
                    elif block.type == "text" and block.content:
                        columns[col_idx].append(_truncate_bullet(block.content, 8))

        # Build as comparison items
        comp_items = []
        subsection_titles = []
        if sections and sections[0].subsections:
            subsection_titles = [s.heading for s in sections[0].subsections[:3]]

        for idx, col in enumerate(columns[:3]):
            title = subsection_titles[idx] if idx < len(subsection_titles) else f"Point {idx+1}"
            title = re.sub(r'^\d+(\.\d+)*\.?\s*', '', title)  # Strip numbering
            words = title.split()
            title = ' '.join(words[:6]) if len(words) > 6 else title
            comp_items.append(ComparisonItem(
                title=title,
                points=col[:3] if col else ["Key insight"],
            ))

        if not comp_items:
            return self._optimize_bullets(spec, sections, all_text)

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment="three_column",
            key_message=spec.key_message,
            comparison_items=comp_items,
        )

    def _optimize_icon_grid(
        self, spec: SlideSpec, sections: list[Section], all_text: str
    ) -> OptimizedSlideContent:
        """Extract 4–6 categorised items for an icon-grid infographic card layout.

        Each card gets a title (first 3–5 words) and a description (the rest).
        """
        items: list[ComparisonItem] = []

        for block in self._collect_blocks(sections):
            if block.type not in ("bullet_list", "numbered_list") or not block.items:
                continue
            for raw in block.items[:6]:
                raw = _strip_citations(raw).strip()
                # Split on first ':', '—', ' – ', or ' - '
                for sep in (":", " — ", " – ", " - "):
                    if sep in raw:
                        title, desc = raw.split(sep, 1)
                        break
                else:
                    # No separator — use first 4 words as title
                    words = raw.split()
                    title = " ".join(words[:4])
                    desc = " ".join(words[4:])

                title = title.strip()[:40]
                desc = desc.strip()[:100]
                if title:
                    items.append(ComparisonItem(
                        title=title,
                        points=[desc] if desc else [],
                    ))
            if items:
                break  # Use the first list block only

        if len(items) < 4:
            # Fallback: build items from subsection headings + first bullet each
            items = []
            for sub in (sections[0].subsections if sections else [])[:6]:
                heading = re.sub(r'^\d+(\.\d+)*\.?\s*', '', sub.heading).strip()[:40]
                first_point = ""
                for blk in sub.content_blocks:
                    if blk.items:
                        first_point = _truncate_bullet(blk.items[0], 12)
                        break
                    if blk.content:
                        first_point = _truncate_bullet(blk.content, 12)
                        break
                if heading:
                    items.append(ComparisonItem(title=heading, points=[first_point] if first_point else []))

        if len(items) < 3:
            return self._optimize_bullets(spec, sections, all_text)

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment="icon_grid",
            key_message=spec.key_message,
            comparison_items=items[:6],
        )

    def _optimize_funnel(
        self, spec: SlideSpec, sections: list[Section], all_text: str
    ) -> OptimizedSlideContent:
        """Extract 3–5 funnel stages from numbered/bullet lists.

        Values (numbers, percentages, currencies) are separated from labels.
        """
        _value_re = re.compile(
            r'([\$€£₹]?\s*[\d,]+\.?\d*\s*(?:%|[KMBkm](?:illion|rillion)?|'
            r'leads?|applicants?|candidates?|users?|customers?))',
            re.IGNORECASE,
        )

        funnel_items: list[FunnelItem] = []

        for block in self._collect_blocks(sections):
            if block.type not in ("numbered_list", "bullet_list") or not block.items:
                continue
            for raw in block.items[:5]:
                raw = _strip_citations(raw).strip()
                # Extract embedded numeric value
                m = _value_re.search(raw)
                value = m.group(1).strip() if m else ""
                label = _value_re.sub("", raw).strip(" :–—-")
                label = _truncate_bullet(label, 7)
                if label:
                    funnel_items.append(FunnelItem(label=label, value=value))
            if funnel_items:
                break

        if len(funnel_items) < 3:
            return self._optimize_bullets(spec, sections, all_text)

        return OptimizedSlideContent(
            slide_number=spec.slide_number,
            slide_type=spec.slide_type,
            title=spec.title,
            visual_treatment="funnel",
            key_message=spec.key_message,
            funnel_items=funnel_items[:5],
        )

    # ──────────────────────────────────────────────
    # Helper methods
    # ──────────────────────────────────────────────

    def _build_agenda(self, doc: DocumentIR) -> list[str]:
        """Build agenda items from document sections."""
        skip = {"table of contents", "references", "source documentation",
                "bibliography", "appendix", "executive summary"}
        items = []
        for s in doc.sections:
            heading_lower = s.heading.lower().strip()
            if any(kw in heading_lower for kw in skip):
                continue
            clean = re.sub(r'^\d+(\.\d+)*\.?\s*', '', s.heading).strip()
            if clean and len(items) < 6:
                items.append(clean)

        # Fallback: all sections were filtered — use subsection headings instead
        if not items:
            for s in doc.sections:
                for sub in s.subsections:
                    clean = re.sub(r'^\d+(\.\d+)*\.?\s*', '', sub.heading).strip()
                    if clean and len(items) < 6:
                        items.append(clean)
                if len(items) >= 3:
                    break

        return items

    def _collect_blocks(self, sections: list[Section]) -> list[ContentBlock]:
        """Collect all content blocks from sections and subsections."""
        blocks = []
        for section in sections:
            blocks.extend(section.content_blocks)
            for sub in section.subsections:
                blocks.extend(sub.content_blocks)
        return blocks

    def _iter_blocks(self, section: Section):
        """Iterate over all blocks in a section and its subsections."""
        yield from section.content_blocks
        for sub in section.subsections:
            yield from sub.content_blocks

    def _collect_text(self, sections: list[Section]) -> str:
        """Collect all text from sections into a single string."""
        parts = []
        for section in sections:
            for block in self._iter_blocks(section):
                if block.content:
                    parts.append(block.content)
                if block.items:
                    parts.extend(block.items)
        return ' '.join(parts)

    def _table_to_chart(
        self, headers: list[str], rows: list[list[str]], treatment: str
    ) -> ChartData | None:
        """Convert a table to chart data if possible."""
        if not headers or not rows:
            return None

        # Try to find numeric columns
        numeric_cols = []
        for col_idx in range(len(headers)):
            numeric_count = 0
            for row in rows:
                if col_idx < len(row):
                    val = re.sub(r'[,$%€£₹]', '', row[col_idx].strip())
                    try:
                        float(val.replace(',', ''))
                        numeric_count += 1
                    except (ValueError, AttributeError):
                        pass
            if numeric_count >= len(rows) * 0.5:
                numeric_cols.append(col_idx)

        if not numeric_cols:
            return None

        # First non-numeric column = categories
        cat_col = 0
        for i in range(len(headers)):
            if i not in numeric_cols:
                cat_col = i
                break

        categories = []
        for row in rows[:self.max_table_rows]:
            if cat_col < len(row):
                categories.append(row[cat_col][:20])
            else:
                categories.append(f"Item {len(categories)+1}")

        # Build series from numeric columns (max 3)
        series = []
        for nc in numeric_cols[:3]:
            values = []
            for row in rows[:self.max_table_rows]:
                if nc < len(row):
                    val = re.sub(r'[,$%€£₹\s]', '', row[nc])
                    try:
                        values.append(float(val.replace(',', '')))
                    except (ValueError, AttributeError):
                        values.append(0.0)
                else:
                    values.append(0.0)
            series.append(SeriesData(name=headers[nc], values=values))

        chart_type_map = {
            "chart_bar": "bar",
            "chart_pie": "pie",
            "chart_line": "line",
            "chart_area": "area",
        }
        ct = chart_type_map.get(treatment, "bar")

        return ChartData(
            chart_type=ct,
            title="",
            categories=categories,
            series=series,
        )

    def _extract_chart_from_text(
        self, text: str, treatment: str
    ) -> ChartData | None:
        """Attempt to extract chart data from text (simple pattern matching)."""
        # Look for patterns like "X: Y%" or "X - Y%"
        pattern = re.compile(r'([A-Za-z\s]{3,30})[\s:–-]+(\d+(?:\.\d+)?)\s*%')
        matches = pattern.findall(text)

        if len(matches) >= 3:
            categories = [m[0].strip()[:20] for m in matches[:6]]
            values = [float(m[1]) for m in matches[:6]]
            chart_type_map = {
                "chart_bar": "bar",
                "chart_pie": "pie",
                "chart_line": "line",
                "chart_area": "area",
            }
            return ChartData(
                chart_type=chart_type_map.get(treatment, "bar"),
                title="",
                categories=categories,
                series=[SeriesData(name="Value", values=values)],
            )

        return None

    def _extract_kpi_label(self, context: str, value: str) -> str | None:
        """Extract a short label for a KPI value from surrounding text.

        Strategy: prefer words that come BEFORE the value in the sentence
        (they are almost always the metric name), fall back to words after.
        """
        context = _strip_citations(context)
        val_idx = context.find(value)
        if val_idx == -1:
            val_idx = len(context)

        before = context[:val_idx].strip()
        after  = context[val_idx + len(value):].strip()

        _stop = {'the', 'a', 'an', 'of', 'in', 'at', 'to', 'for', 'and',
                 'or', 'is', 'are', 'was', 'has', 'have', 'its', 'with',
                 'by', 'that', 'this', 'from', 'as', 'on', 'be', 'it',
                 # Connector/modifier words that appear right before values
                 'approximately', 'around', 'about', 'nearly', 'roughly',
                 'over', 'above', 'below', 'under', 'stood', 'reached',
                 'hit', 'exceeding', 'exceeded', 'increasing', 'growing',
                 'declining', 'fallen', 'rising', 'totaling', 'totalling',
                 'representing', 'accounting', 'comprising', 'reaching',
                 'surpassing', 'surpassed', 'estimated', 'projected',
                 'recorded', 'reported', 'achieved', 'generated',
                 'currently', 'approximately'}

        def _clean_words(text: str) -> list[str]:
            words = text.split()
            return [w.strip('.,;:()[]"\'') for w in words
                    if len(w) > 2
                    and w.lower() not in _stop
                    and not w.startswith('http')]

        # Try the last 4 meaningful words before the value first
        before_words = _clean_words(before)
        if before_words:
            label = ' '.join(before_words[-4:])
            if len(label) >= 4:
                return label[:40]

        # Fall back to first meaningful words after the value
        after_words = _clean_words(after)
        if after_words:
            return ' '.join(after_words[:4])[:40]

        return None
