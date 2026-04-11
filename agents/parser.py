"""
Agent 1: Markdown Parser

Converts raw Markdown text into a structured DocumentIR (Intermediate Representation)
that all downstream agents consume.

Responsibilities:
  - Parse headings, paragraphs, lists, tables using markdown-it-py tokens
  - Detect numeric data (currencies, percentages, years, growth rates)
  - Strip citations ([N]) from display text
  - Extract bold/emphasis markers
  - Handle edge cases: missing H1, irregular nesting, empty sections

Input:  Raw markdown string
Output: DocumentIR (Pydantic model)
"""

from __future__ import annotations

import os
import re
from typing import Literal

from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode

from core.models import ContentBlock, DocumentIR, Section


# ── Numeric data detection patterns ──────────────

_NUMERIC_PATTERNS = [
    re.compile(r'\$[\d,.]+\s*[BMKTbmkt](?:illion|rillion)?', re.IGNORECASE),  # $5.9B, $2.7 billion
    re.compile(r'[\d,.]+\s*%'),                                                # 78%, 24.4%
    re.compile(r'(?:USD|EUR|INR|GBP|₹|€|£)\s*[\d,.]+'),                       # USD 252B, ₹1.2L
    re.compile(r'[\d,.]+\s*(?:MW|GW|TWh|GWh|kWh)'),                           # 3,860 MW
    re.compile(r'\b\d{4}\s*[-–]\s*\d{4}\b'),                                  # 2020-2025
    re.compile(r'(?:CAGR|YoY|MoM)\s*(?:of\s*)?\s*[\d,.]+\s*%', re.IGNORECASE),  # CAGR of 24.4%
    re.compile(r'[\d,.]+\s*(?:billion|million|trillion|crore|lakh)', re.IGNORECASE),
]

_CITATION_PATTERN = re.compile(r'\s*\[\d+\]\s*')
_CITATION_URL_PATTERN = re.compile(r'\(https?://[^\)]+\)')
_BOLD_PATTERN = re.compile(r'\*\*(.*?)\*\*')


def _has_numeric_data(text: str) -> bool:
    """Check if text contains numeric/quantitative data."""
    return any(p.search(text) for p in _NUMERIC_PATTERNS)


def _strip_citations(text: str) -> str:
    """Remove inline citation references like [1], [22] and citation URLs."""
    text = _CITATION_PATTERN.sub(' ', text)
    text = _CITATION_URL_PATTERN.sub('', text)
    return text.strip()


def _extract_inline_text(token) -> str:
    """Extract plain text from an inline token, handling children."""
    if token.children:
        parts = []
        for child in token.children:
            if child.type == 'text':
                parts.append(child.content)
            elif child.type == 'softbreak':
                parts.append(' ')
            elif child.type == 'code_inline':
                parts.append(child.content)
            elif child.type in ('strong_open', 'strong_close',
                                'em_open', 'em_close',
                                'link_open', 'link_close',
                                's_open', 's_close'):
                pass  # Skip formatting markers
            else:
                if child.content:
                    parts.append(child.content)
        return ''.join(parts)
    return token.content or ''


def _count_words(text: str) -> int:
    """Count words in a text string."""
    return len(text.split())


class MarkdownParser:
    """Parse markdown into a structured DocumentIR."""

    def __init__(self):
        self.md = MarkdownIt().enable('table')

    def parse(self, md_text: str, source_filename: str = "") -> DocumentIR:
        """
        Parse markdown text into DocumentIR.

        Args:
            md_text: Raw markdown content.
            source_filename: Original filename (used as title fallback).

        Returns:
            DocumentIR with all sections, content blocks, and metadata.
        """
        tokens = self.md.parse(md_text)
        return self._build_ir(tokens, source_filename)

    def _build_ir(self, tokens: list, source_filename: str) -> DocumentIR:
        """Walk token list and build the document IR."""
        title = ""
        subtitle = None
        sections: list[Section] = []
        citations: list[str] = []
        total_tables = 0
        total_lists = 0
        total_words = 0

        i = 0
        current_h2: Section | None = None
        current_h3: Section | None = None
        first_heading_found = False

        while i < len(tokens):
            token = tokens[i]

            # ── Headings ─────────────────────────
            if token.type == 'heading_open':
                level = int(token.tag[1])  # h1→1, h2→2, etc.
                # Next token is the inline content
                inline_token = tokens[i + 1] if i + 1 < len(tokens) else None
                heading_text = _extract_inline_text(inline_token) if inline_token else ""
                heading_text = _strip_citations(heading_text)

                if level == 1 and not title:
                    title = heading_text
                    first_heading_found = True
                elif level == 1:
                    # Additional H1s treated as H2 sections
                    current_h3 = None
                    current_h2 = Section(heading=heading_text, level=2)
                    sections.append(current_h2)
                elif level == 2:
                    current_h3 = None
                    current_h2 = Section(heading=heading_text, level=2)
                    sections.append(current_h2)
                elif level == 3:
                    if not first_heading_found:
                        # H3 before any H2 — could be subtitle
                        pass
                    if not title and not first_heading_found:
                        title = heading_text
                        first_heading_found = True
                    elif title and subtitle is None and current_h2 is None and len(sections) == 0:
                        # H3 right after H1 = subtitle
                        subtitle = heading_text
                    elif current_h2 is not None:
                        current_h3 = Section(heading=heading_text, level=3)
                        current_h2.subsections.append(current_h3)
                    else:
                        # H3 without parent H2, create standalone section
                        current_h3 = None
                        current_h2 = Section(heading=heading_text, level=2)
                        sections.append(current_h2)
                elif level >= 4:
                    # H4+ treated as content within current section
                    target = current_h3 or current_h2
                    if target:
                        target.content_blocks.append(ContentBlock(
                            type="text",
                            content=f"**{heading_text}**",
                            has_numeric_data=_has_numeric_data(heading_text),
                        ))

                i += 2  # Skip heading_open + inline
                if i < len(tokens) and tokens[i].type == 'heading_close':
                    i += 1
                continue

            # ── Paragraphs ───────────────────────
            if token.type == 'paragraph_open':
                inline_token = tokens[i + 1] if i + 1 < len(tokens) else None
                text = _extract_inline_text(inline_token) if inline_token else ""
                text = _strip_citations(text)

                if text.strip():
                    target = current_h3 or current_h2
                    if target:
                        has_numeric = _has_numeric_data(text)
                        target.content_blocks.append(ContentBlock(
                            type="text",
                            content=text,
                            has_numeric_data=has_numeric,
                            is_visualizable=has_numeric,
                        ))
                        target.word_count += _count_words(text)
                        total_words += _count_words(text)
                        if has_numeric:
                            target.has_numeric_data = True
                            if current_h3 and current_h2:
                                current_h2.has_numeric_data = True

                i += 2  # Skip paragraph_open + inline
                if i < len(tokens) and tokens[i].type == 'paragraph_close':
                    i += 1
                continue

            # ── Bullet / Ordered Lists ───────────
            if token.type in ('bullet_list_open', 'ordered_list_open'):
                list_type: Literal["bullet_list", "numbered_list"] = (
                    "bullet_list" if token.type == 'bullet_list_open' else "numbered_list"
                )
                items, end_idx = self._parse_list(tokens, i)
                items = [_strip_citations(item) for item in items]

                total_lists += 1
                combined = ' '.join(items)
                has_numeric = _has_numeric_data(combined)

                target = current_h3 or current_h2
                if target and items:
                    target.content_blocks.append(ContentBlock(
                        type=list_type,
                        items=items,
                        has_numeric_data=has_numeric,
                        is_visualizable=has_numeric,
                    ))
                    wc = _count_words(combined)
                    target.word_count += wc
                    total_words += wc
                    if has_numeric:
                        target.has_numeric_data = True
                        if current_h3 and current_h2:
                            current_h2.has_numeric_data = True

                i = end_idx + 1
                continue

            # ── Tables ───────────────────────────
            if token.type == 'table_open':
                headers, rows, end_idx = self._parse_table(tokens, i)
                total_tables += 1

                combined = ' '.join(headers) + ' ' + ' '.join(
                    cell for row in rows for cell in row
                )
                has_numeric = _has_numeric_data(combined)

                target = current_h3 or current_h2
                if target:
                    target.content_blocks.append(ContentBlock(
                        type="table",
                        headers=headers,
                        rows=rows,
                        has_numeric_data=has_numeric,
                        is_visualizable=has_numeric,
                    ))
                    target.has_table_data = True
                    if has_numeric:
                        target.has_numeric_data = True
                        if current_h3 and current_h2:
                            current_h2.has_numeric_data = True
                            current_h2.has_table_data = True

                i = end_idx + 1
                continue

            # ── Blockquotes ──────────────────────
            if token.type == 'blockquote_open':
                text_parts = []
                j = i + 1
                while j < len(tokens) and tokens[j].type != 'blockquote_close':
                    if tokens[j].type == 'inline':
                        text_parts.append(_extract_inline_text(tokens[j]))
                    j += 1
                text = ' '.join(text_parts)
                text = _strip_citations(text)

                target = current_h3 or current_h2
                if target and text.strip():
                    target.content_blocks.append(ContentBlock(
                        type="blockquote",
                        content=text,
                    ))

                i = j + 1
                continue

            # ── Code blocks ──────────────────────
            if token.type == 'fence' or token.type == 'code_block':
                target = current_h3 or current_h2
                if target and token.content and token.content.strip():
                    target.content_blocks.append(ContentBlock(
                        type="code_block",
                        content=token.content,
                    ))
                i += 1
                continue

            i += 1

        # ── Fallback title from filename ─────
        if not title and source_filename:
            title = os.path.splitext(source_filename)[0].replace('_', ' ').replace('-', ' ')
        if not title:
            title = "Untitled Presentation"

        # ── Compute section word counts for H2 (sum of subsections) ──
        for section in sections:
            if section.subsections:
                section.word_count += sum(s.word_count for s in section.subsections)

        # ── Extract citations from raw text ──
        citation_matches = re.findall(r'\[(\d+)\]', ' '.join(
            b.content or '' for s in sections
            for b in s.content_blocks
        ))
        citations = sorted(set(citation_matches), key=lambda x: int(x))

        return DocumentIR(
            title=title,
            subtitle=subtitle,
            sections=sections,
            total_word_count=total_words,
            total_tables=total_tables,
            total_lists=total_lists,
            citations=citations,
        )

    def _parse_list(self, tokens: list, start: int) -> tuple[list[str], int]:
        """Parse a bullet/ordered list and return (items, end_index)."""
        items = []
        close_type = 'bullet_list_close' if tokens[start].type == 'bullet_list_open' else 'ordered_list_close'
        i = start + 1
        depth = 1

        current_item_parts = []

        while i < len(tokens):
            t = tokens[i]

            if t.type == close_type:
                depth -= 1
                if depth == 0:
                    if current_item_parts:
                        items.append(' '.join(current_item_parts))
                    return items, i

            if t.type in ('bullet_list_open', 'ordered_list_open'):
                depth += 1

            if t.type == 'list_item_open' and depth == 1:
                if current_item_parts:
                    items.append(' '.join(current_item_parts))
                    current_item_parts = []

            if t.type == 'inline':
                text = _extract_inline_text(t)
                if text.strip():
                    current_item_parts.append(text.strip())

            i += 1

        if current_item_parts:
            items.append(' '.join(current_item_parts))

        return items, i - 1

    def _parse_table(self, tokens: list, start: int) -> tuple[list[str], list[list[str]], int]:
        """Parse a table and return (headers, rows, end_index)."""
        headers = []
        rows = []
        current_row = []
        in_header = False
        i = start + 1

        while i < len(tokens):
            t = tokens[i]

            if t.type == 'table_close':
                if current_row:
                    rows.append(current_row)
                return headers, rows, i

            if t.type == 'thead_open':
                in_header = True
            elif t.type == 'thead_close':
                in_header = False
            elif t.type == 'tbody_open':
                pass
            elif t.type == 'tr_open':
                current_row = []
            elif t.type == 'tr_close':
                if in_header:
                    headers = current_row
                else:
                    if current_row:
                        rows.append(current_row)
                current_row = []
            elif t.type == 'inline':
                text = _extract_inline_text(t)
                current_row.append(text.strip())

            i += 1

        return headers, rows, i - 1
