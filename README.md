# MD2PPTX — Intelligent Markdown to PowerPoint Converter

A multi-agent pipeline that transforms `.md` research reports into visually polished, 10-15 slide `.pptx` presentations using a provided Slide Master template.

Built for the **Code EZ: Master of Agents** hackathon.

---

## Quick Start

### Prerequisites

- Python 3.11+
- (Optional) Google Gemini API key for LLM-powered content optimization

### Installation

```bash
cd md2pptx
pip install -r requirements.txt
```

### Run — Single File

```bash
python main.py \
  --input "test_cases/Accenture Tech Acquisition Analysis.md" \
  --template "templates/Template_Accenture Tech Acquisition Analysis.pptx" \
  --output output/result.pptx \
  --slides 12
```

### Run — Batch Mode (All Test Cases)

```bash
python main.py --batch \
  --input-dir test_cases/ \
  --template "templates/Template_Accenture Tech Acquisition Analysis.pptx" \
  --output-dir output/ \
  --slides 12
```

### Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `GEMINI_API_KEY` | No | Google Gemini API key. Enables LLM-powered slide planning and content condensation. Falls back to rule-based logic if unset. |

---

## System Architecture

```
                    .md file
                       |
                       v
             +--------------------+
             |  Agent 1: Parser   |    markdown-it-py tokenizer
             |  (parser.py)       |    Sections, tables, numeric detection
             +--------------------+
                       |
                   DocumentIR
                       |
                       v
             +--------------------+
             | Agent 2: Strategist|    LLM or rule-based slide planning
             | (strategist.py)    |    Section merging, visual treatment assignment
             +--------------------+
                       |
                   SlidePlan
                       |
                       v
             +--------------------+
             | Agent 3: Optimizer |    Content condensation (LLM/rules)
             | (content_optimizer)|    KPI extraction, chart data, bullet truncation
             +--------------------+
                       |
              OptimizedSlideContent[]
                       |
                       v
             +--------------------+
             | Agent 4: Layout    |    12-column grid system
             | (layout_engine.py) |    Shape positioning, icons, decorative elements
             +--------------------+
                       |
                  SlideLayout[]
                       |
                       v
             +--------------------+
             | Agent 5: Visual    |    Charts, tables, trend indicators
             | (visual_generator) |    Color vibrancy, shape formatting
             +--------------------+
                       |
                       v
             +--------------------+
             | Agent 6: Renderer  |    Assembles final .pptx from template
             | (renderer.py)      |    Layout matching, placeholder fallback
             +--------------------+
                       |
                       v
                  output.pptx
```

### Data Flow Contracts

Each agent communicates through strictly typed Pydantic models defined in `core/models.py`:

| Model | Producer | Consumer | Purpose |
|-------|----------|----------|---------|
| `DocumentIR` | Parser | Strategist, Optimizer | Structured representation of the markdown |
| `SlidePlan` | Strategist | Optimizer | Slide count, types, visual treatments, section mapping |
| `OptimizedSlideContent` | Optimizer | Layout, Renderer | Slide-ready content: bullets, KPIs, chart data, etc. |
| `SlideLayout` | Layout Engine | Renderer | Exact shape positions, sizes, colors for each slide |
| `ThemeConfig` | Theme extractor | Layout, Visual, Renderer | Colors, fonts, dimensions from Slide Master |

---

## Key Design Decisions

### 1. Infographic-First Visual Treatment

The Strategist assigns visual treatments using a priority cascade:

```
Charts > KPI Cards > Tables > Process Flows > Columns > Bullets
```

Bullet points are the last resort. Sections with numeric data automatically get charts or KPI cards. Sections with subsections get multi-column layouts.

### 2. 12-Column Grid System

All element positioning uses a 12-column grid (`core/grid.py`) derived from analyzing the provided Slide Master templates:

- Slide: 13.333" x 7.5" (16:9)
- Margins: 0.375" left/right
- Column gutter: 0.20"
- Content area starts below title at y=1.40"

This ensures consistent alignment across all slide types.

### 3. Domain-Aware Unicode Icons

A keyword-to-icon mapping system (`layout_engine.py:DOMAIN_ICONS`) provides visual identity without external image assets. Over 60 keywords across 12 domains (AI, Finance, Security, Energy, Healthcare, etc.) map to Unicode symbols rendered inside colored circles.

### 4. Exact Slide Count Enforcement

The Strategist guarantees the exact target slide count through a multi-pass approach:

1. Section merging/expansion to fill content slots
2. Subsection expansion for additional slides
3. Final enforcement loop before Thank You slide

Validated on all 24 test cases at target=12.

### 5. Theme Fidelity

Colors, fonts, and slide dimensions are extracted from the Slide Master XML (`core/theme.py`), not hardcoded. A saturation boost prevents washed-out accent colors from rendering poorly against white backgrounds.

### 6. LLM with Rule-Based Fallback

Both the Strategist and Content Optimizer try Google Gemini first (if `GEMINI_API_KEY` is set), then fall back to deterministic rule-based logic. The system works fully without any API key.

---

## Project Structure

```
md2pptx/
  main.py                  CLI entry point (Click)
  config.yaml              All configurable parameters
  requirements.txt         Python dependencies

  agents/
    parser.py              Agent 1 - Markdown to DocumentIR
    strategist.py          Agent 2 - Slide planning & section mapping
    content_optimizer.py   Agent 3 - Content condensation & data extraction
    layout_engine.py       Agent 4 - Grid-based shape positioning
    visual_generator.py    Agent 5 - Charts, tables, shape rendering
    renderer.py            Agent 6 - Final .pptx assembly

  core/
    models.py              Pydantic data models (pipeline contracts)
    grid.py                12-column grid system
    theme.py               Slide Master theme extraction
    quality_checker.py     Post-render validation

  templates/               Slide Master .pptx files
  test_cases/              24 markdown test inputs
  output/                  Generated .pptx outputs
```

---

## Slide Types & Visual Treatments

| Treatment | When Used | Shapes/Slide |
|-----------|-----------|-------------|
| `cover_layout` | Slide 1 | Title + subtitle + accent bars |
| `bullets` | Text-heavy sections | Numbered circles + separator lines (~19) |
| `kpi_cards` | Numeric data (2-4 values) | Accent bars + large values + icons (~27) |
| `chart_bar/pie/line/area` | Tables with numbers | Native python-pptx charts + framing |
| `table` | Tabular data | Themed headers + trend arrows (▲/▼) |
| `process_flow` | Sequential steps | Step circles + boxes + arrows |
| `timeline` | Date-based data | Horizontal line + milestone markers |
| `two_column` | 2 subsections | Header bars + accent dots + dividers |
| `three_column` | 3+ subsections | Icons + accent strips + body points (~38) |
| `closing_layout` | Last slide | Template-driven + text fallback |

---

## Configuration

All parameters are in `config.yaml`:

- `slide_count.default`: Target slide count (default: 12)
- `content.max_bullets_per_slide`: Max bullet points (default: 5)
- `content.max_words_per_bullet`: Max words per bullet (default: 12)
- `content.max_table_rows`: Max table rows (default: 6)
- `llm.model`: Gemini model name (default: gemini-2.5-flash)
- `llm.temperature`: Generation temperature (default: 0.3)

---

## CLI Reference

```
Usage: python main.py [OPTIONS]

Options:
  --input PATH          Input .md file
  --template PATH       Slide Master .pptx template (required)
  --output PATH         Output .pptx file path
  --slides INTEGER      Target slide count, 10-15 (default: 12)
  --config PATH         Custom config.yaml path
  --batch               Process all .md files in --input-dir
  --input-dir PATH      Directory with .md files (batch mode)
  --output-dir PATH     Output directory (batch mode)
  --list-layouts        List available layouts in template
```

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| python-pptx | >= 1.0.2 | PPTX generation, charts, tables |
| markdown-it-py | >= 4.0.0 | Markdown parsing |
| mdit-py-plugins | >= 0.4.0 | Extended markdown support |
| click | >= 8.3.0 | CLI interface |
| PyYAML | >= 6.0.3 | Configuration loading |
| google-generativeai | >= 0.8.0 | LLM-powered optimization (optional) |
| pydantic | >= 2.0.0 | Data validation & contracts |

---

## Validation Results

All 24 provided test cases produce valid output:

- Exact slide count: 24/24 at target=12
- Quality checks: 24/24 passed
- No crashes on any input (2.8K - 29K words, 0-35 tables)
- Average ~19 shapes per slide (sample target: 21.4)
- Structured flow verified: Cover -> Agenda -> Exec Summary -> Content -> Conclusion -> Thank You
