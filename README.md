# MD2PPTX

**Intelligent Markdown → PowerPoint converter built on a 6-agent pipeline.**

Feed it a `.md` research report and a Slide Master `.pptx`. Get back a 10–15 slide deck that's branded, structured, and visually consistent — with native charts, tables, and infographics generated programmatically from the content.

Built for the **Code EZ: Master of Agents** hackathon (April 2026).

---

## Why this exists

Analysts spend hours turning research markdown into client-ready decks: condensing prose into bullets, building charts from tables, fighting with the slide master to keep branding consistent. MD2PPTX does that in under a minute, with a system designed around three principles:

1. **Storyline, not transformation.** The pipeline doesn't map sections to slides 1:1. An LLM-powered Strategist reasons about which sections deserve focus, generates an executive summary, and enforces the exact target slide count.
2. **Bullets are the last resort.** A visual treatment cascade prefers charts, KPI cards, tables, process flows, and multi-column layouts over walls of text.
3. **The Slide Master is the source of truth.** Colors, fonts, dimensions, and grid all come from the provided template — never hardcoded. Dynamically generated layouts inherit the same design language.

---

## Quick start

### Prerequisites

- Python 3.11+
- *(Optional)* Google Gemini API key for LLM-powered planning. Without it, the system falls back to deterministic rule-based logic and still produces valid decks.

### Install

```bash
cd md2pptx
pip install -r requirements.txt
```

### Run a single file

```bash
python main.py \
  --input "test_cases/Banking ROE Competitive Benchmarking Analysis.md" \
  --template "templates/Template_Accenture Tech Acquisition Analysis.pptx" \
  --output output/banking_roe.pptx \
  --slides 15
```

### Run all 24 test cases (batch mode)

```bash
python main.py --batch \
  --input-dir test_cases/ \
  --templates-dir templates/ \
  --output-dir generated_outputs/
```

Passing `--templates-dir` lets the pipeline **auto-pick the best-matching slide master per input** (by filename similarity) and rotate through all available templates instead of reusing one. Omitting `--slides` enables **adaptive slide sizing**: the pipeline picks 10, 12, 14, or 15 slides based on the document's word count (<4k / <7k / <12k / ≥12k).

### Inspect a Slide Master

```bash
python main.py --list-layouts \
  --template "templates/Template_Accenture Tech Acquisition Analysis.pptx"
```

### Environment

| Variable | Required | Description |
|---|---|---|
| `GEMINI_API_KEY` | No | Google Gemini API key. Enables LLM-powered slide planning and content condensation. Falls back to rule-based logic if unset — the system runs fully offline without it. |

---

## Architecture — the 6-agent pipeline

```
       .md file
          │
          ▼
┌─────────────────────┐
│ Agent 1: Parser     │  markdown-it-py tokenizer
│ parser.py           │  → DocumentIR  (sections, tables, numerics, metadata)
└─────────────────────┘
          │
          ▼
┌─────────────────────┐
│ Agent 2: Strategist │  Storyline planning (Gemini + rule-based fallback)
│ strategist.py       │  → SlidePlan   (slide types, treatments, section mapping)
└─────────────────────┘                 Guarantees exact target slide count
          │
          ▼
┌─────────────────────┐
│ Agent 3: Optimizer  │  Bullet condensation, KPI extraction, chart-data prep
│ content_optimizer.py│  → OptimizedSlideContent[]
└─────────────────────┘
          │
          ▼
┌─────────────────────┐
│ Agent 4: Layout     │  12-column grid, shape positioning, dynamic layouts
│ layout_engine.py    │  → SlideLayout[]
└─────────────────────┘
          │
          ▼
┌─────────────────────┐
│ Agent 5: Visual Gen │  Native PPTX charts, tables, infographics
│ visual_generator.py │  Trend arrows, color vibrancy, shape formatting
└─────────────────────┘
          │
          ▼
┌─────────────────────┐
│ Agent 6: Renderer   │  Assembles final .pptx against the Slide Master
│ renderer.py         │  Layout matching + placeholder fallback
└─────────────────────┘
          │
          ▼
   Quality Checker → output.pptx
```

### Pipeline contracts

Every agent boundary is a strictly typed Pydantic model in `core/models.py`. No agent is allowed to read another agent's internal state — only the IR it receives.

| Model | Producer | Consumers | Purpose |
|---|---|---|---|
| `DocumentIR` | Parser | Strategist, Optimizer | Structured representation of the markdown |
| `SlidePlan` | Strategist | Optimizer | Slide count, types, visual treatments, section mapping |
| `OptimizedSlideContent` | Optimizer | Layout, Renderer | Slide-ready bullets, KPIs, chart data |
| `SlideLayout` | Layout Engine | Renderer | Exact shape positions, sizes, colors per slide |
| `ThemeConfig` | Theme extractor | Layout, Visual, Renderer | Colors, fonts, dimensions extracted from the Slide Master XML |

This separation is what makes the system scale to new templates and unseen markdown shapes — you can swap any single agent without touching the others.

---

## Key design decisions

### 1. Storyline over transformation

The Strategist treats the markdown as raw material, not a script. Section weighting is driven by content density: a banking report with a heavy competitive-analysis section gets three slides on competitors and one on methodology, not vice-versa. A final multi-pass enforcement loop guarantees the deck lands at exactly the target slide count — not within ±1.

### 2. Visual treatment cascade

Slide treatments are assigned in priority order:

```
Charts > KPI Cards > Tables > Process Flows > Multi-Column > Bullets
```

Bullets are the last resort. Sections with numeric data automatically get charts or KPI cards; sections with parallel subsections get two- or three-column layouts. This is the reason the data-heavy demos barely contain any bulleted text.

### 3. 12-column grid system

All element positioning uses a 12-column grid (`core/grid.py`) derived from analyzing the provided Slide Masters:

- Slide: 13.333" × 7.5" (16:9)
- Margins: 0.375" left/right, 1.40" top (below title), 0.40" bottom
- Gutter: 0.20"

This is what gives every slide consistent alignment regardless of which treatment was chosen.

### 4. Theme fidelity — the Slide Master is the source of truth

Colors, fonts, and dimensions are extracted from the Slide Master XML at runtime (`core/theme.py`), never hardcoded. A saturation boost prevents washed-out accent colors from rendering poorly against white backgrounds. **Dynamically generated layouts** (for content shapes the master doesn't cover) inherit the same theme, so the deck stays visually cohesive.

### 5. LLM with deterministic fallback

Both the Strategist and the Content Optimizer try Google Gemini first (if `GEMINI_API_KEY` is set), then fall back to deterministic rule-based logic. **The system works fully without any API key** — you trade some narrative polish for full offline operation.

### 6. Domain-aware Unicode iconography

A keyword-to-icon mapping (`layout_engine.py:DOMAIN_ICONS`) gives every section a relevant glyph without external image assets. Over 60 keywords across 12 domains (AI, Finance, Security, Energy, Healthcare, etc.) map to Unicode symbols rendered inside themed circles. Zero copyrighted graphics, zero asset downloads.

### 7. Post-render quality checker

After the Renderer writes the file, the Quality Checker re-opens it and validates: slide count in range, no empty placeholders, no overflowing text frames, font consistency. Failures surface in the terminal, not in the delivered deck.

---

## Slide treatments

| Treatment | When used | What it produces |
|---|---|---|
| `cover_layout` | Slide 1 | Title + subtitle + accent bars from master |
| `kpi_cards` | 2–4 numeric values | Big values, accent bars, domain icons |
| `chart_bar / pie / line / area` | Tables with numeric columns | Native python-pptx charts (fully editable) |
| `table` | Tabular data | Themed headers, capped at 6 rows, trend arrows (▲/▼) |
| `process_flow` | Sequential steps | Step circles + boxes + arrows |
| `timeline` | Date-based events | Horizontal line + milestone markers |
| `two_column` / `three_column` | Parallel subsections | Header bars, accent strips, body points |
| `bullets` | Last resort, text-only sections | Numbered circles + separator lines |
| `closing_layout` | Final slide | Template-driven with text fallback |

---

## Robustness & validation

Tested against all 24 provided test cases:

| Metric | Result |
|---|---|
| Decks generated | 24 / 24 |
| Quality checks passed | 24 / 24 |
| Crashes | 0 |
| Input range | 2.8K – 29K words, 0 – 35 tables |
| Slide count distribution (adaptive) | 10 → 2 decks · 12 → 4 · 14 → 13 · 15 → 5 |
| Slide masters used across batch | 3 / 3 templates rotated by filename match |
| Decks containing native charts | 22 / 24 |
| Decks containing tables | 13 / 24 |

**Edge cases handled gracefully:**
- Inputs with zero numeric data → chart treatments skipped, system falls back to text/structure layouts
- Inputs with malformed or missing headings → parser falls back to flat sections
- Inputs exceeding 5MB → caught at the CLI boundary
- Missing or unrecognized layouts in the master → dynamic layout generator inherits theme and builds one
- Three different Slide Masters tested — each produces a visually distinct deck with **zero code changes**

**Compatibility:** output `.pptx` opens correctly in Microsoft 365 PowerPoint, Google Slides, and LibreOffice Impress.

---

## Project structure

```
md2pptx/
├── main.py                  CLI entry point (Click)
├── config.yaml              All tunable parameters
├── requirements.txt         Python dependencies
│
├── agents/
│   ├── parser.py            Agent 1 — Markdown → DocumentIR
│   ├── strategist.py        Agent 2 — Slide planning + storyline
│   ├── content_optimizer.py Agent 3 — Bullet condensation, data extraction
│   ├── layout_engine.py     Agent 4 — Grid-based shape positioning
│   ├── visual_generator.py  Agent 5 — Charts, tables, infographics
│   └── renderer.py          Agent 6 — Final .pptx assembly
│
├── core/
│   ├── models.py            Pydantic models — pipeline contracts
│   ├── grid.py              12-column grid system
│   ├── theme.py             Slide Master theme extraction
│   └── quality_checker.py   Post-render validation
│
├── templates/               3 Slide Master .pptx files
├── test_cases/              24 markdown test inputs
└── output/                  Generated .pptx outputs
```

---

## CLI reference

```
Usage: python main.py [OPTIONS]

Options:
  --input PATH          Input .md file
  --template PATH       Slide Master .pptx (optional if --templates-dir set)
  --templates-dir PATH  Directory of Template_*.pptx; auto-picks best match per input
  --output PATH         Output .pptx file path
  --slides INTEGER      Target slide count, 10–15. Omit for adaptive sizing.
  --config PATH         Custom config.yaml path
  --batch               Process all .md files in --input-dir
  --input-dir PATH      Directory with .md files (batch mode)
  --output-dir PATH     Output directory (batch mode)
  --list-layouts        List available layouts in the template
```

---

## Configuration

All tunables live in `config.yaml`:

| Key | Default | Purpose |
|---|---|---|
| `slide_count.default` | 12 | Target slide count |
| `content.max_bullets_per_slide` | 5 | Hard cap on bullets |
| `content.max_words_per_bullet` | 12 | Hard cap on words per bullet |
| `content.max_table_rows` | 6 | Table truncation cap |
| `content.max_kpi_cards` | 4 | KPI card cap |
| `llm.model` | gemini-2.5-flash | LLM for Strategist + Optimizer |
| `llm.temperature` | 0.3 | Generation temperature |

---

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| python-pptx | ≥ 1.0.2 | PPTX generation, native charts, tables |
| markdown-it-py | ≥ 4.0.0 | Markdown parsing |
| mdit-py-plugins | ≥ 0.4.0 | Extended markdown support |
| click | ≥ 8.3.0 | CLI |
| PyYAML | ≥ 6.0.3 | Configuration loading |
| pydantic | ≥ 2.0.0 | Pipeline data contracts |
| google-generativeai | ≥ 0.8.0 | LLM-powered planning *(optional)* |

---

## License & IP

Built for the Code EZ hackathon. Per the brief, IP rights for winning submissions transfer to EZ.
