"""
MD2PPTX Agents - Multi-agent pipeline for Markdown to PowerPoint conversion.

Each agent handles a single responsibility in the pipeline:
  1. Parser: Markdown → Structured IR
  2. Strategist: IR → Slide Plan (LLM-powered)
  3. ContentOptimizer: Verbose content → Slide-ready content (LLM-powered)
  4. LayoutEngine: Content → Positioned layout with grid system
  5. VisualGenerator: Data → Charts, tables, infographics
  6. Renderer: All inputs → Final .pptx file
"""
