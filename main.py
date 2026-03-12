"""
main.py
───────
Content Migration Tool – Entry Point

Reads a Word document (.docx), converts its content to HTML, and uploads
the resulting article to Document360 via their REST API.

Usage:
    python main.py                          # uses sample_document.docx
    python main.py path/to/document.docx    # uses a custom file
    python main.py --preview                # parse + convert only (no upload)
    python main.py --setup                  # fetch project versions & categories
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

from docx_parser import DocxParser
from html_converter import HtmlConverter
from document360_client import Document360Client, Document360Config

# ─── Logging setup ───────────────────────────────────────────────────────────

LOG_FORMAT = "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s"


def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format=LOG_FORMAT)
    # Quiet noisy libraries
    logging.getLogger("urllib3").setLevel(logging.WARNING)


# ─── Configuration ───────────────────────────────────────────────────────────

def load_config() -> Document360Config:
    """Load Document360 configuration from environment variables."""
    load_dotenv()

    api_token = os.getenv("DOCUMENT360_API_TOKEN", "")
    if not api_token or api_token == "your_api_token_here":
        logging.warning(
            "⚠️  DOCUMENT360_API_TOKEN is not set. "
            "API upload will fail. Use --preview to test without uploading."
        )

    return Document360Config(
        api_token=api_token,
        base_url=os.getenv("DOCUMENT360_BASE_URL", "https://apihub.document360.io"),
        project_version_id=os.getenv("DOCUMENT360_PROJECT_VERSION_ID", ""),
        category_id=os.getenv("DOCUMENT360_CATEGORY_ID", ""),
        user_id=os.getenv("DOCUMENT360_USER_ID", ""),
    )


# ─── CLI ─────────────────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Migrate Word documents (.docx) to Document360",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python main.py                        Parse & upload sample_document.docx\n"
            "  python main.py my_doc.docx             Parse & upload a specific file\n"
            "  python main.py --preview               Parse & preview HTML (no upload)\n"
            "  python main.py --setup                 Fetch project/category IDs from API\n"
        ),
    )
    parser.add_argument(
        "docx_file",
        nargs="?",
        default="sample_document.docx",
        help="Path to the .docx file to process (default: sample_document.docx)",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Parse and show HTML output without uploading to Document360",
    )
    parser.add_argument(
        "--save-html",
        metavar="FILE",
        help="Save the generated HTML to a file",
    )
    parser.add_argument(
        "--title",
        help="Article title (default: derived from the document filename)",
    )
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Fetch and display project versions and categories from Document360",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose/debug logging",
    )
    return parser


# ─── Workflow steps ──────────────────────────────────────────────────────────

def step_parse_document(filepath: str) -> list:
    """Step 1 & 2: Load and parse the .docx file."""
    logger = logging.getLogger("main")
    logger.info("=" * 60)
    logger.info("STEP 1 & 2: Loading and parsing document")
    logger.info("=" * 60)

    parser = DocxParser(filepath)
    elements = parser.parse()

    logger.info("  Parsed %d structural elements:", len(elements))
    from collections import Counter
    from docx_parser import ElementType
    counts = Counter(e.element_type.name for e in elements)
    for etype, count in sorted(counts.items()):
        logger.info("    %-20s %d", etype, count)

    return elements


def step_convert_to_html(elements: list, parser: DocxParser = None) -> str:
    """Step 3: Convert extracted elements to HTML."""
    logger = logging.getLogger("main")
    logger.info("")
    logger.info("=" * 60)
    logger.info("STEP 3: Converting to HTML")
    logger.info("=" * 60)

    converter = HtmlConverter()
    html_content = converter.convert(elements, parser)

    logger.info("  Generated %d characters of HTML", len(html_content))
    return html_content


def step_upload_to_document360(
    config: Document360Config, title: str, html_content: str
) -> dict:
    """Step 4, 5 & 6: Prepare payload, send POST, log response."""
    logger = logging.getLogger("main")
    logger.info("")
    logger.info("=" * 60)
    logger.info("STEP 4-6: Uploading to Document360")
    logger.info("=" * 60)

    client = Document360Client(config)

    logger.info("  Title     : %s", title)
    logger.info("  Category  : %s", config.category_id)
    logger.info("  Version   : %s", config.project_version_id)

    result = client.create_article(title=title, html_content=html_content)

    logger.info("")
    logger.info("API Response:")
    import json
    logger.info(json.dumps(result, indent=2))

    return result


def run_setup(config: Document360Config) -> None:
    """Fetch project versions and categories from Document360 API."""
    logger = logging.getLogger("main")
    client = Document360Client(config)

    logger.info("Fetching project versions...")
    versions = client.get_project_versions()
    import json
    logger.info("Project Versions:\n%s", json.dumps(versions, indent=2))

    # If we have a project version ID, also fetch categories
    if config.project_version_id:
        logger.info("\nFetching categories for version %s...", config.project_version_id)
        categories = client.get_categories(config.project_version_id)
        logger.info("Categories:\n%s", json.dumps(categories, indent=2))


# ─── Main entrypoint ────────────────────────────────────────────────────────

def main() -> None:
    cli = build_parser()
    args = cli.parse_args()
    setup_logging(args.verbose)
    logger = logging.getLogger("main")

    config = load_config()

    # ── Setup mode ───────────────────────────────────────────────────────
    if args.setup:
        run_setup(config)
        return

    # ── Normal migration flow ────────────────────────────────────────────
    filepath = args.docx_file
    if not os.path.isfile(filepath):
        logger.error("❌ File not found: %s", filepath)
        sys.exit(1)

    logger.info("🚀 Starting content migration")
    logger.info("   Source file: %s", os.path.abspath(filepath))
    logger.info("")

    # Step 1 & 2: Parse
    parser = DocxParser(filepath)
    elements = parser.parse()

    # Step 3: Convert to HTML
    html_content = step_convert_to_html(elements, parser)

    # Save HTML to file if requested
    if args.save_html:
        output_path = args.save_html
        Path(output_path).write_text(html_content, encoding="utf-8")
        logger.info("  Saved HTML to: %s", os.path.abspath(output_path))

    # Preview mode: print HTML and exit
    if args.preview:
        logger.info("")
        logger.info("=" * 60)
        logger.info("PREVIEW MODE – Generated HTML:")
        logger.info("=" * 60)
        print("\n" + html_content + "\n")
        logger.info("✅ Preview complete (no upload performed)")
        return

    # Step 4-6: Upload
    title = args.title or Path(filepath).stem.replace("_", " ").title()
    step_upload_to_document360(config, title, html_content)

    logger.info("")
    logger.info("✅ Content migration complete!")


if __name__ == "__main__":
    main()
