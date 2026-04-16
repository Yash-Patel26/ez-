"""
Agent 2: Storyline Strategist

Creates an intelligent slide plan from the parsed DocumentIR.
Uses Google Gemini API for semantic section merging and storyline creation,
with a rule-based fallback when no API key is available.

Responsibilities:
  - Decide how many slides to create (10-15)
  - Map document sections to slides (merge related sections)
  - Assign visual treatment per slide (chart, table, bullets, etc.)
  - Enforce structured flow: Cover → Agenda → Exec Summary → Content → Data → Conclusion

Input:  DocumentIR + target slide count
Output: SlidePlan
"""

from __future__ import annotations

import json
import os
import re
from typing import Any

from core.models import DocumentIR, Section, SlideSpec, SlidePlan


class Strategist:
    """Create a slide storyline plan from document structure."""

    def __init__(self, config: dict | None = None):
        self.config = config or {}
        self.llm_config = self.config.get("llm", {})
        self.api_key = os.environ.get("GEMINI_API_KEY", "")

    def create_plan(self, doc: DocumentIR, target_slides: int = 12) -> SlidePlan:
        """
        Create a slide plan from the document IR.

        Tries LLM-based planning first, falls back to rule-based.
        """
        target_slides = max(10, min(15, target_slides))

        if self.api_key:
            try:
                plan = self._llm_plan(doc, target_slides)
                # Validate: LLM must return within ±2 of target
                if abs(len(plan.slides) - target_slides) <= 2:
                    return plan
                else:
                    print(f"       LLM returned only {len(plan.slides)} slides, using rule-based fallback")
            except Exception as e:
                print(f"       LLM strategist failed ({e}), using rule-based fallback")

        return self._rule_based_plan(doc, target_slides)

    # ──────────────────────────────────────────────
    # LLM-based planning
    # ──────────────────────────────────────────────

    def _llm_plan(self, doc: DocumentIR, target_slides: int) -> SlidePlan:
        """Use Google Gemini API to create an intelligent slide storyline."""
        import google.generativeai as genai

        genai.configure(api_key=self.api_key)
        model = genai.GenerativeModel(
            self.llm_config.get("model", "gemini-2.0-flash")
        )

        doc_summary = self._build_doc_summary(doc)

        prompt = f"""You are a presentation strategist. Create a slide plan for a PowerPoint presentation.

DOCUMENT SUMMARY:
Title: {doc.title}
Subtitle: {doc.subtitle or 'N/A'}
Total sections: {len(doc.sections)}
Total words: {doc.total_word_count}
Total tables: {doc.total_tables}

SECTIONS:
{doc_summary}

REQUIREMENTS:
- Create exactly {target_slides} slides
- Follow this flow: Cover → Agenda (optional) → Executive Summary → Section Content → Data/Charts → Conclusion → Thank You
- Merge related sections when needed to fit the target slide count
- Sections with numeric data or tables should use chart/table visual treatments
- Each slide must have exactly 1 key message
- Prioritize visual treatments (charts, KPI cards, process flows) over plain bullets
- Skip "Table of Contents" and "References" sections — don't give them dedicated slides

VISUAL TREATMENTS (pick one per slide):
- "cover_layout": Title slide
- "bullets": Standard bullet points
- "two_column": Two-column comparison or split content
- "three_column": Three items side by side
- "chart_bar": Bar/column chart for categorical comparison
- "chart_pie": Pie chart for distribution/share data
- "chart_line": Line chart for time-series/trends
- "chart_area": Area chart for cumulative trends
- "table": Data table
- "kpi_cards": 3-4 large metric callout cards
- "process_flow": Step-by-step process visualization
- "timeline": Chronological timeline
- "comparison_cards": Side-by-side comparison cards
- "closing_layout": Thank you / closing slide

Return ONLY a valid JSON array of slide objects. Each object must have:
- "slide_number": integer (1-indexed)
- "slide_type": one of "cover", "agenda", "executive_summary", "content", "data_chart", "data_table", "comparison", "timeline", "process_flow", "kpi_callout", "conclusion", "thank_you"
- "title": string (concise slide title, max 8 words)
- "key_message": string (1 sentence key takeaway)
- "source_sections": array of section indices (0-indexed) from the SECTIONS list above
- "visual_treatment": one of the treatments listed above
- "content_priority": "high", "medium", or "low"

Return ONLY the JSON array, no other text."""

        response = model.generate_content(
            prompt,
            generation_config=genai.types.GenerationConfig(
                temperature=self.llm_config.get("temperature", 0.3),
                max_output_tokens=self.llm_config.get("max_tokens", 4096),
                response_mime_type="application/json",
            ),
        )

        response_text = response.text.strip()
        return self._parse_llm_response(response_text, doc, target_slides)

    def _build_doc_summary(self, doc: DocumentIR) -> str:
        """Build compact section summary for the LLM prompt."""
        lines = []
        for i, section in enumerate(doc.sections):
            flags = []
            if section.has_numeric_data:
                flags.append("HAS_NUMERIC_DATA")
            if section.has_table_data:
                flags.append("HAS_TABLES")

            sub_count = len(section.subsections)
            flags_str = f" [{', '.join(flags)}]" if flags else ""
            lines.append(
                f"[{i}] {section.heading} "
                f"({section.word_count}w, {sub_count} subsections){flags_str}"
            )

            # Include subsection headings for context
            for j, sub in enumerate(section.subsections[:5]):  # Limit to 5
                sub_flags = []
                if sub.has_numeric_data:
                    sub_flags.append("NUM")
                if sub.has_table_data:
                    sub_flags.append("TBL")
                sf = f" [{','.join(sub_flags)}]" if sub_flags else ""
                lines.append(f"    {j}. {sub.heading[:80]}{sf}")

            if sub_count > 5:
                lines.append(f"    ... and {sub_count - 5} more subsections")

        return "\n".join(lines)

    def _parse_llm_response(self, text: str, doc: DocumentIR, target: int) -> SlidePlan:
        """Parse JSON response from LLM into SlidePlan. Handles messy Gemini output."""
        text = text.strip()

        # Strip markdown code fences
        text = re.sub(r'^```\w*\n?', '', text)
        text = re.sub(r'\n?```\s*$', '', text)
        text = text.strip()

        # Extract JSON array — find the [ ... ] portion
        start = text.find('[')
        end = text.rfind(']')
        if start != -1 and end != -1 and end > start:
            text = text[start:end + 1]

        # Try parsing directly first
        try:
            slides_data = json.loads(text)
        except json.JSONDecodeError:
            # Fix truncated JSON: remove trailing incomplete object and close
            text = re.sub(r',\s*\{[^}]*$', '', text)  # Remove last incomplete object
            text = text.rstrip(',\n ')
            open_braces = text.count('{') - text.count('}')
            open_brackets = text.count('[') - text.count(']')
            text += '}' * max(0, open_braces)
            text += ']' * max(0, open_brackets)

            try:
                slides_data = json.loads(text)
            except json.JSONDecodeError:
                # Last resort: extract individual objects with regex
                objs = re.findall(
                    r'\{[^{}]*"slide_number"\s*:\s*\d+[^{}]*\}', text
                )
                if objs:
                    slides_data = [json.loads(o) for o in objs]
                else:
                    raise

        slides = []
        for sd in slides_data:
            slides.append(SlideSpec(
                slide_number=sd.get("slide_number", len(slides) + 1),
                slide_type=sd.get("slide_type", "content"),
                title=sd.get("title", ""),
                key_message=sd.get("key_message", ""),
                source_sections=sd.get("source_sections", []),
                visual_treatment=sd.get("visual_treatment", "bullets"),
                content_priority=sd.get("content_priority", "high"),
            ))

        return SlidePlan(
            slides=slides,
            storyline_summary=f"LLM-generated {len(slides)}-slide plan for '{doc.title}'",
        )

    # ──────────────────────────────────────────────
    # Rule-based fallback
    # ──────────────────────────────────────────────

    def _rule_based_plan(self, doc: DocumentIR, target_slides: int) -> SlidePlan:
        """Create a slide plan using deterministic rules (no LLM needed)."""
        slides: list[SlideSpec] = []
        slide_num = 1

        # Filter out non-content sections
        skip_keywords = {"table of contents", "references", "source documentation",
                         "bibliography", "appendix", "disclaimer"}
        content_sections = []
        for i, s in enumerate(doc.sections):
            heading_lower = s.heading.lower().strip()
            if any(kw in heading_lower for kw in skip_keywords):
                continue
            content_sections.append((i, s))

        # ── Slide 1: Cover ───────────────────────
        slides.append(SlideSpec(
            slide_number=slide_num,
            slide_type="cover",
            title=doc.title,
            key_message=doc.subtitle or doc.title,
            source_sections=[],
            visual_treatment="cover_layout",
            content_priority="high",
        ))
        slide_num += 1

        # ── Slide 2: Agenda ──────────────────────
        slides.append(SlideSpec(
            slide_number=slide_num,
            slide_type="agenda",
            title="Agenda",
            key_message="Overview of topics covered",
            source_sections=[],
            visual_treatment="bullets",
            content_priority="medium",
        ))
        slide_num += 1

        # ── Slide 3: Executive Summary ───────────
        exec_sections = [i for i, s in content_sections
                         if "executive summary" in s.heading.lower()
                         or "overview" in s.heading.lower()]
        # Fallback: use first section only (avoid consuming too many)
        if not exec_sections and content_sections:
            exec_sections = [content_sections[0][0]]
        slides.append(SlideSpec(
            slide_number=slide_num,
            slide_type="executive_summary",
            title="Executive Summary",
            key_message="Key findings and highlights",
            source_sections=exec_sections[:1],
            visual_treatment="kpi_cards",
            content_priority="high",
        ))
        slide_num += 1

        # ── Content slides: distribute remaining sections ──
        # Remove only the specific sections used for exec summary
        exec_set = set(exec_sections[:1])  # Only remove 1 section for exec
        remaining = [(i, s) for i, s in content_sections if i not in exec_set]

        # How many content group slides do we need?
        # slide_num is the NEXT slide number (1-indexed), so slides added so far = slide_num - 1
        # Remaining = target - (slide_num - 1) = conclusion(1) + thank_you(1) + content_groups
        # content_groups = target - (slide_num - 1) - 2
        content_group_count = target_slides - (slide_num - 1) - 2
        content_group_count = max(content_group_count, 1)

        if len(remaining) <= content_group_count:
            groups = [([i], s) for i, s in remaining]
            extra_slots = content_group_count - len(groups)
            if extra_slots > 0:
                groups = self._expand_sections(groups, remaining, extra_slots)
        else:
            groups = self._merge_sections(remaining, content_group_count)

        # Ensure we have exactly content_group_count groups
        while len(groups) < content_group_count:
            expanded = False
            for g_idx, (indices, sec) in enumerate(list(groups)):
                for sub in sec.subsections:
                    sub_title = self._clean_title(sub.heading)
                    if sub.word_count >= 30 and not any(
                        g[1].heading == sub.heading for g in groups
                    ):
                        fake = Section(
                            heading=sub.heading, level=sub.level,
                            content_blocks=sub.content_blocks,
                            subsections=sub.subsections,
                            word_count=sub.word_count,
                            has_numeric_data=sub.has_numeric_data,
                            has_table_data=sub.has_table_data,
                        )
                        groups.insert(g_idx + 1, (indices, fake))
                        expanded = True
                        break
                if expanded:
                    break
            if not expanded:
                break  # Truly nothing left to expand

        # ── Insert section divider markers ──────────────────────
        # For documents with 3+ section groups, inject a divider slide at the
        # natural content midpoint.  A sentinel value (None, section) signals
        # "emit a divider here" during the loop below.
        if len(groups) >= 3:
            mid = len(groups) // 2
            divider_sec = groups[mid][1]
            groups = list(groups[:mid]) + [(None, divider_sec)] + list(groups[mid:])

        # Track recent treatments to enforce visual variety
        recent_treatments: list[str] = []

        for group_indices, primary_section in groups:
            if slide_num >= target_slides - 1:
                break

            # ── Section divider sentinel ──────────────────────────
            if group_indices is None:
                slides.append(SlideSpec(
                    slide_number=slide_num,
                    slide_type="divider",
                    title=self._clean_title(primary_section.heading),
                    key_message="",
                    source_sections=[],
                    visual_treatment="divider_layout",
                    content_priority="low",
                ))
                slide_num += 1
                continue

            treatment = self._pick_visual_treatment(primary_section)

            # ── Visual variety enforcement ──
            from collections import Counter
            t_counts = Counter(recent_treatments)

            # Rule 1: No same treatment 2+ times in a row
            if (len(recent_treatments) >= 2
                    and recent_treatments[-1] == treatment
                    and recent_treatments[-2] == treatment):
                treatment = self._rotate_treatment(treatment, primary_section)

            # Rule 2: No treatment more than 3 times total — force hard rotation
            if t_counts.get(treatment, 0) >= 3:
                # Cycle through all options until we find one under 3
                cycle = ["two_column", "process_flow", "three_column",
                         "chart_bar", "kpi_cards", "bullets", "table"]
                for alt in cycle:
                    if t_counts.get(alt, 0) < 2:
                        treatment = alt
                        break

            slide_type = self._pick_slide_type(primary_section, treatment)
            recent_treatments.append(treatment)

            slides.append(SlideSpec(
                slide_number=slide_num,
                slide_type=slide_type,
                title=self._clean_title(primary_section.heading),
                key_message="",
                source_sections=group_indices,
                visual_treatment=treatment,
                content_priority="high" if primary_section.has_numeric_data else "medium",
            ))
            slide_num += 1

        # ── Conclusion slide ─────────────────────
        conclusion_sections = [i for i, s in content_sections
                               if any(kw in s.heading.lower()
                                      for kw in ("conclusion", "recommendation", "takeaway",
                                                  "key findings", "summary", "way forward"))]
        # Fallback: use last content section if no explicit conclusion found
        if not conclusion_sections and content_sections:
            conclusion_sections = [content_sections[-1][0]]
        slides.append(SlideSpec(
            slide_number=slide_num,
            slide_type="conclusion",
            title="Key Takeaways & Recommendations",
            key_message="Strategic recommendations and next steps",
            source_sections=conclusion_sections[:2],
            visual_treatment="bullets",
            content_priority="high",
        ))
        slide_num += 1

        # ── Pad to exact target slide count ──
        # target_slides includes the Thank You slide, so pad until
        # slide_num == target_slides - 1 (the Thank You slot)
        target_before_thankyou = target_slides - 1
        used_subsection_headings = set()

        # Pass 1: expand subsections (word_count >= 30)
        while slide_num < target_before_thankyou:
            expanded = False
            for idx_s, sec in content_sections:
                for sub in sec.subsections:
                    sub_title = self._clean_title(sub.heading)
                    if (sub.word_count >= 30
                            and sub_title not in used_subsection_headings
                            and not any(sl.title == sub_title for sl in slides)):
                        treatment = self._pick_visual_treatment(sub)
                        stype = self._pick_slide_type(sub, treatment)
                        slides.append(SlideSpec(
                            slide_number=slide_num,
                            slide_type=stype,
                            title=sub_title,
                            key_message="",
                            source_sections=[idx_s],
                            visual_treatment=treatment,
                            content_priority="medium",
                        ))
                        used_subsection_headings.add(sub_title)
                        slide_num += 1
                        expanded = True
                        break
                if expanded:
                    break
            if not expanded:
                break

        # Pass 2: if still short, re-present content sections with
        #         varied visual treatments to fill the target exactly
        alt_treatments = ["two_column", "process_flow", "kpi_cards",
                          "three_column", "bullets", "table",
                          "chart_bar", "timeline"]
        alt_idx = 0
        # Use ALL content sections (not just remaining) for maximum coverage
        all_sections_for_padding = content_sections if content_sections else (
            [(0, doc.sections[0])] if doc.sections else []
        )
        existing_titles = {sl.title for sl in slides}
        while slide_num < target_before_thankyou:
            # Generate a unique title by incrementing a counter
            placed = False
            for idx_s, sec in all_sections_for_padding:
                if slide_num >= target_before_thankyou:
                    break
                base_title = self._clean_title(sec.heading)
                # Find a unique title variant
                title = base_title
                counter = 2
                while title in existing_titles:
                    title = f"{base_title} — Part {counter}"
                    counter += 1
                treatment = alt_treatments[alt_idx % len(alt_treatments)]
                alt_idx += 1
                slides.append(SlideSpec(
                    slide_number=slide_num,
                    slide_type="content",
                    title=title,
                    key_message="",
                    source_sections=[idx_s],
                    visual_treatment=treatment,
                    content_priority="medium",
                ))
                existing_titles.add(title)
                slide_num += 1
                placed = True
            if not placed:
                break  # Safety: no sections at all

        # ── Final slide count enforcement ────────
        # Current slides + 1 (thank_you) must equal target_slides
        # Insert BEFORE conclusion to maintain structured flow
        while len(slides) + 1 < target_slides:
            src_idx = (len(slides) - 3) % max(len(content_sections), 1)
            src_sec = content_sections[src_idx] if content_sections else (
                (0, doc.sections[0]) if doc.sections else None
            )
            if src_sec is None:
                break
            idx_s, sec = src_sec
            base = self._clean_title(sec.heading)
            title = base
            part = 2
            existing = {sl.title for sl in slides}
            while title in existing:
                title = f"{base} — Part {part}"
                part += 1
            t = alt_treatments[len(slides) % len(alt_treatments)]
            # Insert before conclusion (second-to-last slide)
            insert_pos = len(slides) - 1  # Before conclusion
            slides.insert(insert_pos, SlideSpec(
                slide_number=insert_pos + 1,
                slide_type="content",
                title=title,
                key_message="",
                source_sections=[idx_s],
                visual_treatment=t,
                content_priority="medium",
            ))

        # ── Thank You slide ──────────────────────
        slides.append(SlideSpec(
            slide_number=len(slides) + 1,
            slide_type="thank_you",
            title="Thank You",
            key_message="",
            source_sections=[],
            visual_treatment="closing_layout",
            content_priority="low",
        ))

        # Renumber slides sequentially
        for idx, sl in enumerate(slides):
            sl.slide_number = idx + 1

        return SlidePlan(
            slides=slides,
            storyline_summary=f"Rule-based {len(slides)}-slide plan for '{doc.title}'",
        )

    def _expand_sections(
        self,
        groups: list[tuple[list[int], Section]],
        remaining: list[tuple[int, Section]],
        extra_slots: int,
    ) -> list[tuple[list[int], Section]]:
        """Expand data-rich sections into multiple slides to fill target count."""
        if extra_slots <= 0:
            return groups

        # Find sections with subsections that have table/numeric data — best for expansion
        expandable = []
        for idx, (indices, section) in enumerate(groups):
            data_subs = [s for s in section.subsections
                         if s.has_table_data or s.has_numeric_data]
            if len(data_subs) >= 2:
                expandable.append((idx, section, data_subs))

        # Sort by most data-rich subsections first
        expandable.sort(key=lambda x: len(x[2]), reverse=True)

        new_groups = list(groups)
        added = 0

        # First pass: expand data-rich sections
        for orig_idx, section, data_subs in expandable:
            if added >= extra_slots:
                break
            best_sub = data_subs[0]
            fake_section = Section(
                heading=best_sub.heading,
                level=best_sub.level,
                content_blocks=best_sub.content_blocks,
                word_count=best_sub.word_count,
                has_numeric_data=best_sub.has_numeric_data,
                has_table_data=best_sub.has_table_data,
            )
            insert_pos = min(orig_idx + 1 + added, len(new_groups))
            parent_indices = new_groups[orig_idx][0] if orig_idx < len(new_groups) else []
            new_groups.insert(insert_pos, (parent_indices, fake_section))
            added += 1

        # Second pass: if still short, expand any section with 2+ subsections
        if added < extra_slots:
            all_expandable = [
                (idx, section)
                for idx, (indices, section) in enumerate(groups)
                if len(section.subsections) >= 2
            ]
            all_expandable.sort(key=lambda x: x[1].word_count, reverse=True)

            for orig_idx, section in all_expandable:
                if added >= extra_slots:
                    break
                # Pick a subsection not already expanded
                for sub in section.subsections:
                    if added >= extra_slots:
                        break
                    if sub.word_count >= 50:
                        fake = Section(
                            heading=sub.heading,
                            level=sub.level,
                            content_blocks=sub.content_blocks,
                            subsections=sub.subsections,
                            word_count=sub.word_count,
                            has_numeric_data=sub.has_numeric_data,
                            has_table_data=sub.has_table_data,
                        )
                        parent_indices = new_groups[orig_idx][0] if orig_idx < len(new_groups) else []
                        insert_pos = min(orig_idx + 1 + added, len(new_groups))
                        new_groups.insert(insert_pos, (parent_indices, fake))
                        added += 1

        return new_groups

    def _merge_sections(
        self,
        sections: list[tuple[int, Section]],
        max_groups: int,
    ) -> list[tuple[list[int], Section]]:
        """Merge sections into exactly max_groups groups."""
        if not sections:
            return []

        n = len(sections)
        if n <= max_groups:
            return [([i], s) for i, s in sections]

        # Distribute n sections into max_groups groups as evenly as possible
        # base = sections per group, remainder groups get one extra
        base = n // max_groups
        remainder = n % max_groups

        groups: list[tuple[list[int], Section]] = []
        pos = 0
        for g in range(max_groups):
            size = base + (1 if g < remainder else 0)
            chunk = sections[pos:pos + size]
            pos += size
            indices = [i for i, _ in chunk]
            primary = max(chunk, key=lambda x: (
                x[1].has_table_data,
                x[1].has_numeric_data,
                x[1].word_count,
            ))[1]
            groups.append((indices, primary))

        return groups

    # Keywords that suggest a funnel / pipeline slide
    _FUNNEL_KW = {
        "funnel", "pipeline", "conversion", "hiring process", "recruitment process",
        "sales process", "lead", "adoption", "onboarding", "stages", "journey",
        "attrition", "dropout", "retention funnel",
    }

    def _pick_visual_treatment(self, section: Section) -> str:
        """Choose the best visual treatment based on section content.

        Infographic-first: prioritize charts > KPIs > tables > columns > bullets.
        """
        heading_lower = section.heading.lower()

        # Priority 1: Tables with numeric data → chart
        if section.has_table_data:
            for block in self._all_blocks(section):
                if block.type == "table" and block.rows and block.has_numeric_data:
                    if len(block.rows) <= 8:
                        return "chart_bar"
                    return "table"
            # Table without numeric data → still show as table
            for block in self._all_blocks(section):
                if block.type == "table" and block.rows:
                    return "table"

        # Priority 2: Numeric data → KPI cards (lowered threshold from 3 to 2)
        if section.has_numeric_data:
            numeric_blocks = sum(
                1 for b in self._all_blocks(section) if b.has_numeric_data
            )
            if numeric_blocks >= 2:
                return "kpi_cards"
            # Even 1 numeric block → try chart from text
            return "chart_bar"

        # Priority 2.5 (infographic): Funnel diagram for pipeline / stage content
        if any(kw in heading_lower for kw in self._FUNNEL_KW):
            has_list = any(
                b.type in ("numbered_list", "bullet_list") and b.items
                for b in self._all_blocks(section)
            )
            if has_list:
                return "funnel"

        # Priority 3: Multi-column for subsection structure
        if len(section.subsections) == 2:
            # Use comparison_cards when subsection headings signal a VS / options framing
            sub_text = " ".join(s.heading.lower() for s in section.subsections)
            _cmp_kw = {
                "vs", "versus", "vs.", "option", "approach", "alternative",
                "traditional", "modern", "before", "after", "pros", "cons",
                "advantage", "disadvantage", "method", "strategy a", "strategy b",
            }
            if any(kw in sub_text for kw in _cmp_kw):
                return "comparison_cards"
            return "two_column"
        if len(section.subsections) >= 3:
            return "three_column"

        # Priority 3.5 (infographic): Icon grid for 4–6 categorised bullet items
        for block in self._all_blocks(section):
            if block.type in ("bullet_list", "numbered_list") and block.items:
                if 4 <= len(block.items) <= 6:
                    return "icon_grid"
                break  # Only inspect the first list block

        # Priority 4: Fallback to bullets (last resort)
        return "bullets"

    def _rotate_treatment(self, current: str, section: Section) -> str:
        """Force a different visual treatment to avoid repetition.

        Rotation priority: ensures visual variety across slides.
        """
        # Define rotation options for each treatment type
        rotations = {
            "three_column":    ["two_column", "kpi_cards", "icon_grid", "bullets", "process_flow"],
            "bullets":         ["icon_grid", "two_column", "kpi_cards", "three_column", "process_flow"],
            "kpi_cards":       ["chart_bar", "two_column", "three_column", "icon_grid", "bullets"],
            "chart_bar":       ["table", "kpi_cards", "two_column", "bullets"],
            "two_column":      ["comparison_cards", "three_column", "bullets", "kpi_cards", "process_flow"],
            "table":           ["chart_bar", "kpi_cards", "bullets", "two_column"],
            "icon_grid":       ["two_column", "bullets", "three_column", "kpi_cards"],
            "funnel":          ["process_flow", "bullets", "kpi_cards"],
            "comparison_cards":["two_column", "three_column", "bullets"],
        }
        alternatives = rotations.get(current, ["bullets", "two_column", "kpi_cards"])

        # Pick first alternative that makes sense for the content
        for alt in alternatives:
            if alt == "chart_bar" and section.has_numeric_data:
                return alt
            if alt == "kpi_cards" and section.has_numeric_data:
                return alt
            if alt == "table" and section.has_table_data:
                return alt
            if alt == "two_column" and len(section.subsections) >= 2:
                return alt
            if alt == "three_column" and len(section.subsections) >= 3:
                return alt
            if alt == "process_flow" and any(
                b.type in ("bullet_list", "numbered_list") for b in self._all_blocks(section)
            ):
                return alt
            if alt == "icon_grid" and any(
                b.type in ("bullet_list", "numbered_list") and b.items and 4 <= len(b.items) <= 6
                for b in self._all_blocks(section)
            ):
                return alt
            if alt in ("comparison_cards", "funnel") and any(
                b.type in ("bullet_list", "numbered_list") for b in self._all_blocks(section)
            ):
                return alt
            if alt == "bullets":
                return alt

        return "bullets"

    def _pick_slide_type(self, section: Section, treatment: str) -> str:
        """Map visual treatment to slide type."""
        chart_treatments = {"chart_bar", "chart_pie", "chart_line", "chart_area"}
        if treatment in chart_treatments:
            return "data_chart"
        if treatment == "table":
            return "data_table"
        if treatment == "kpi_cards":
            return "kpi_callout"
        if treatment in ("two_column", "comparison_cards"):
            return "comparison"
        if treatment in ("process_flow", "funnel"):
            return "process_flow"
        if treatment == "timeline":
            return "timeline"
        if treatment == "icon_grid":
            return "content"
        return "content"

    def _clean_title(self, heading: str) -> str:
        """Clean up section heading for use as slide title."""
        # Remove numbering prefix like "1." or "1.2."
        title = re.sub(r'^\d+(\.\d+)*\.?\s*', '', heading)
        # Truncate to ~8 words
        words = title.split()
        if len(words) > 10:
            title = ' '.join(words[:10]) + '...'
        return title.strip()

    def _all_blocks(self, section: Section):
        """Yield all content blocks from section and its subsections."""
        yield from section.content_blocks
        for sub in section.subsections:
            yield from sub.content_blocks
