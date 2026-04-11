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
import sys
import time
from pathlib import Path

import click
import yaml

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from core.theme import extract_theme, get_layout_names


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
    target_slides: int = 12,
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
    click.echo(f"  Slides:   {target_slides}")
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
@click.option("--template", "template_path", required=True, type=click.Path(exists=True),
              help="Path to Slide Master .pptx template")
@click.option("--output", "output_path", required=False, type=click.Path(),
              help="Path for output .pptx file")
@click.option("--slides", "target_slides", default=12, type=click.IntRange(10, 15),
              help="Target number of slides (10-15)")
@click.option("--config", "config_path", default=None, type=click.Path(exists=True),
              help="Path to config.yaml (default: ./config.yaml)")
@click.option("--batch", is_flag=True, help="Batch mode: process all .md files in --input-dir")
@click.option("--input-dir", "input_dir", type=click.Path(exists=True),
              help="Directory with .md files (batch mode)")
@click.option("--output-dir", "output_dir", type=click.Path(),
              help="Output directory (batch mode)")
@click.option("--list-layouts", is_flag=True, help="List available layouts in template and exit")
def main(input_path, template_path, output_path, target_slides, config_path,
         batch, input_dir, output_dir, list_layouts):
    """MD2PPTX: Convert Markdown research reports to polished PowerPoint presentations."""

    config = load_config(config_path)

    # List layouts mode
    if list_layouts:
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

        md_files = sorted(input_dir.glob("*.md"))
        if not md_files:
            click.echo(f"No .md files found in {input_dir}", err=True)
            sys.exit(1)

        click.echo(f"\nBatch mode: {len(md_files)} files to process\n")
        results = []

        for i, md_file in enumerate(md_files, 1):
            click.echo(f"\n[{i}/{len(md_files)}] Processing: {md_file.name}")
            out_file = out_dir / md_file.with_suffix(".pptx").name
            try:
                run_pipeline(
                    md_path=str(md_file),
                    template_path=template_path,
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

    run_pipeline(
        md_path=input_path,
        template_path=template_path,
        output_path=output_path,
        target_slides=target_slides,
        config=config,
    )


if __name__ == "__main__":
    main()
