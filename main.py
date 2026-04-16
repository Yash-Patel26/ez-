"""
MD2PPTX - Intelligent Markdown to PowerPoint Converter

A multi-agent pipeline that transforms .md research reports into
visually polished 10-15 slide .pptx presentations.

Usage:
    python main.py --input report.md --template template.pptx --output output.pptx
    python main.py --input report.md --template template.pptx --slides 12
    python main.py --batch --input-dir test_cases/ --template template.pptx --output-dir output/
"""

import os
import re
import sys
import time
from pathlib import Path

import click
import yaml

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from core.theme import extract_theme, get_layout_names


def adaptive_slide_count(word_count: int) -> int:
    """Pick a target slide count based on document length.

    Brief requires 10-15 slides "based on the generated storyline and
    content distribution" — uniform 12 for every input ignores that.
    """
    if word_count < 4000:
        return 10
    if word_count < 7000:
        return 12
    if word_count < 12000:
        return 14
    return 15


def pick_template_for_input(md_path: Path, templates_dir: Path, fallback: str) -> str:
    """Find the best matching template for a markdown input.

    Strategy: exact stem match (Template_<stem>.pptx), then fuzzy substring
    match on the input stem against available template names, then fallback.
    """
    if not templates_dir or not templates_dir.exists():
        return fallback

    stem = md_path.stem
    exact = templates_dir / f"Template_{stem}.pptx"
    if exact.exists():
        return str(exact)

    candidates = sorted(templates_dir.glob("Template_*.pptx"))
    if not candidates:
        return fallback

    stem_lc = stem.lower()
    stem_words = {w for w in re.split(r"[^a-z0-9]+", stem_lc) if len(w) > 3}
    best, best_score = None, 0
    for cand in candidates:
        name_lc = cand.stem.lower()
        name_words = {w for w in re.split(r"[^a-z0-9]+", name_lc) if len(w) > 3}
        score = len(stem_words & name_words)
        if score > best_score:
            best, best_score = cand, score

    if best is not None and best_score >= 1:
        return str(best)

    # Rotate templates by hash so the batch doesn't reuse one master
    idx = hash(stem) % len(candidates)
    return str(candidates[idx])


def load_config(config_path: str | None = None) -> dict:
    """Load configuration from YAML file."""
    if config_path is None:
        config_path = str(Path(__file__).parent / "config.yaml")
    with open(config_path, "r") as f:
        return yaml.safe_load(f)


def run_pipeline(
    md_path: str,
    template_path: str,
    output_path: str,
    target_slides: int | None = 12,
    config: dict | None = None,
) -> str:
    """
    Execute the full MD → PPTX conversion pipeline.

    Args:
        md_path: Path to input markdown file.
        template_path: Path to Slide Master .pptx template.
        output_path: Path for output .pptx file.
        target_slides: Target number of slides (10-15).
        config: Optional config dict override.

    Returns:
        Path to the generated .pptx file.
    """
    if config is None:
        config = load_config()

    click.echo(f"\n{'='*60}")
    click.echo(f"  MD2PPTX Pipeline")
    click.echo(f"{'='*60}")
    click.echo(f"  Input:    {md_path}")
    click.echo(f"  Template: {template_path}")
    click.echo(f"  Output:   {output_path}")
    click.echo(f"  Slides:   {target_slides if target_slides else 'auto'}")
    click.echo(f"{'='*60}\n")

    start_time = time.time()

    # ── Agent 1: Parse Markdown ──────────────────
    click.echo("[1/6] Parsing markdown...")
    from agents.parser import MarkdownParser

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    parser = MarkdownParser()
    document_ir = parser.parse(md_text, source_filename=os.path.basename(md_path))
    click.echo(f"       Parsed: {document_ir.total_word_count} words, "
               f"{len(document_ir.sections)} sections, "
               f"{document_ir.total_tables} tables")

    # Adaptive slide count if caller didn't specify
    if target_slides is None:
        target_slides = adaptive_slide_count(document_ir.total_word_count)
        click.echo(f"       Adaptive slide count: {target_slides} "
                   f"(from {document_ir.total_word_count} words)")

    # ── Agent 2: Create Storyline ────────────────
    click.echo("[2/6] Creating slide storyline...")
    from agents.strategist import Strategist

    strategist = Strategist(config=config)
    slide_plan = strategist.create_plan(document_ir, target_slides)
    click.echo(f"       Planned: {len(slide_plan.slides)} slides")

    # ── Agent 3: Optimize Content ────────────────
    click.echo("[3/6] Optimizing content for slides...")
    from agents.content_optimizer import ContentOptimizer

    optimizer = ContentOptimizer(config=config)
    optimized_slides = optimizer.optimize(slide_plan, document_ir)
    click.echo(f"       Optimized: {len(optimized_slides)} slides ready")

    # ── Agent 4: Compute Layouts ─────────────────
    click.echo("[4/6] Computing slide layouts...")
    from agents.layout_engine import LayoutEngine

    theme = extract_theme(template_path)
    layout_engine = LayoutEngine(theme=theme, config=config)
    slide_layouts = layout_engine.compute(optimized_slides)
    click.echo(f"       Layouts: {len(slide_layouts)} computed")

    # ── Agent 5 & 6: Generate Visuals & Render ───
    click.echo("[5/6] Generating visuals (charts, tables, infographics)...")
    click.echo("[6/6] Rendering final PPTX...")
    from agents.renderer import PPTXRenderer

    renderer = PPTXRenderer(template_path=template_path, theme=theme, config=config)
    renderer.render(slide_layouts, optimized_slides, output_path)

    # ── Quality Check ────────────────────────────
    from core.quality_checker import QualityChecker

    checker = QualityChecker(config=config)
    issues = checker.validate(output_path)
    if issues:
        click.echo(f"\n  Quality warnings ({len(issues)}):")
        for issue in issues[:5]:
            click.echo(f"    - {issue}")
    else:
        click.echo(f"\n  Quality check: All passed!")

    elapsed = time.time() - start_time
    click.echo(f"\n  Done in {elapsed:.1f}s -> {output_path}")
    click.echo(f"{'='*60}\n")

    return output_path


@click.command()
@click.option("--input", "input_path", required=False, type=click.Path(exists=True),
              help="Path to input .md file")
@click.option("--template", "template_path", required=False, type=click.Path(exists=True),
              help="Path to Slide Master .pptx template (optional if --templates-dir set)")
@click.option("--templates-dir", "templates_dir", required=False, type=click.Path(exists=True),
              help="Directory of Template_*.pptx files; auto-picks best match per input")
@click.option("--output", "output_path", required=False, type=click.Path(),
              help="Path for output .pptx file")
@click.option("--slides", "target_slides", default=None, type=click.IntRange(10, 15),
              help="Target number of slides (10-15). Omit for adaptive sizing by word count.")
@click.option("--config", "config_path", default=None, type=click.Path(exists=True),
              help="Path to config.yaml (default: ./config.yaml)")
@click.option("--batch", is_flag=True, help="Batch mode: process all .md files in --input-dir")
@click.option("--input-dir", "input_dir", type=click.Path(exists=True),
              help="Directory with .md files (batch mode)")
@click.option("--output-dir", "output_dir", type=click.Path(),
              help="Output directory (batch mode)")
@click.option("--list-layouts", is_flag=True, help="List available layouts in template and exit")
@click.option("--clean", is_flag=True, help="Remove existing .pptx files from output dir before generating")
def main(input_path, template_path, templates_dir, output_path, target_slides, config_path,
         batch, input_dir, output_dir, list_layouts, clean):
    """MD2PPTX: Convert Markdown research reports to polished PowerPoint presentations."""

    config = load_config(config_path)

    if not template_path and not templates_dir:
        click.echo("Error: provide --template or --templates-dir", err=True)
        sys.exit(1)

    templates_dir_path = Path(templates_dir) if templates_dir else None

    # List layouts mode
    if list_layouts:
        if not template_path:
            click.echo("Error: --template required for --list-layouts", err=True)
            sys.exit(1)
        layouts = get_layout_names(template_path)
        click.echo(f"\nLayouts in {template_path}:")
        for layout in layouts:
            click.echo(f"  [{layout['index']}] {layout['name']}")
            for ph in layout["placeholders"]:
                click.echo(f"       Placeholder {ph['idx']}: {ph['type']} "
                          f"({ph['name']}) @ ({ph['left']}\", {ph['top']}\")")
        return

    # Batch mode
    if batch:
        if not input_dir:
            click.echo("Error: --input-dir required in batch mode", err=True)
            sys.exit(1)

        input_dir = Path(input_dir)
        out_dir = Path(output_dir) if output_dir else input_dir.parent / "output"
        out_dir.mkdir(parents=True, exist_ok=True)

        # Remove existing .pptx files before generating
        if clean:
            existing = list(out_dir.glob("*.pptx"))
            removed, skipped = 0, 0
            for f in existing:
                try:
                    f.unlink()
                    removed += 1
                except PermissionError:
                    skipped += 1
            if removed or skipped:
                msg = f"  Cleared {removed} file(s) from {out_dir}"
                if skipped:
                    msg += f" ({skipped} skipped — close them in PowerPoint first)"
                click.echo(msg + "\n")

        md_files = sorted(input_dir.glob("*.md"))
        if not md_files:
            click.echo(f"No .md files found in {input_dir}", err=True)
            sys.exit(1)

        click.echo(f"\nBatch mode: {len(md_files)} files to process\n")
        results = []

        for i, md_file in enumerate(md_files, 1):
            click.echo(f"\n[{i}/{len(md_files)}] Processing: {md_file.name}")
            out_file = out_dir / md_file.with_suffix(".pptx").name
            # Per-input template selection: match by filename, fall back to --template
            if templates_dir_path:
                chosen_template = pick_template_for_input(
                    md_file, templates_dir_path, template_path or ""
                )
                if not chosen_template:
                    click.echo(f"  ERROR: no template available for {md_file.name}", err=True)
                    results.append((md_file.name, "FAIL: no template"))
                    continue
            else:
                chosen_template = template_path
            try:
                run_pipeline(
                    md_path=str(md_file),
                    template_path=chosen_template,
                    output_path=str(out_file),
                    target_slides=target_slides,
                    config=config,
                )
                results.append((md_file.name, "OK"))
            except Exception as e:
                click.echo(f"  ERROR: {e}", err=True)
                results.append((md_file.name, f"FAIL: {e}"))

        click.echo(f"\n{'='*60}")
        click.echo(f"  Batch Results: {sum(1 for _, s in results if s == 'OK')}/{len(results)} succeeded")
        for name, status in results:
            icon = "+" if status == "OK" else "x"
            click.echo(f"  [{icon}] {name}: {status}")
        click.echo(f"{'='*60}\n")
        return

    # Single file mode
    if not input_path:
        click.echo("Error: --input required (or use --batch)", err=True)
        sys.exit(1)

    if not output_path:
        output_path = str(Path(input_path).with_suffix(".pptx"))

    # Ensure output directory exists
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # Resolve template for single-file mode (supports --templates-dir auto-pick)
    if templates_dir_path and not template_path:
        chosen_template = pick_template_for_input(
            Path(input_path), templates_dir_path, ""
        )
        if not chosen_template:
            click.echo("Error: no template found in --templates-dir", err=True)
            sys.exit(1)
    elif templates_dir_path:
        chosen_template = pick_template_for_input(
            Path(input_path), templates_dir_path, template_path
        )
    else:
        chosen_template = template_path

    run_pipeline(
        md_path=input_path,
        template_path=chosen_template,
        output_path=output_path,
        target_slides=target_slides,
        config=config,
    )


if __name__ == "__main__":
    main()
