"""
html_converter.py
─────────────────
Converts a list of `ContentElement` objects (produced by `DocxParser`) into
clean, well-structured HTML.
"""

from __future__ import annotations

import html
import logging
from typing import List, Optional

from docx_parser import ContentElement, ElementType, InlineContent, DocxParser

logger = logging.getLogger(__name__)


class HtmlConverter:
    """Converts parsed document elements into an HTML string."""

    def __init__(self):
        self.parser = None

    def convert(self, elements: List[ContentElement], parser: Optional[DocxParser] = None) -> str:
        """Convert a list of ContentElement objects to an HTML string."""
        self.parser = parser
        html_parts: List[str] = []

        for element in elements:
            if element.element_type == ElementType.HEADING:
                html_parts.append(self._render_heading(element))

            elif element.element_type == ElementType.PARAGRAPH:
                html_parts.append(self._render_paragraph(element))

            elif element.element_type == ElementType.BULLET_LIST:
                html_parts.append(self._render_bullet_list(element))

            elif element.element_type == ElementType.NUMBERED_LIST:
                html_parts.append(self._render_numbered_list(element))

            elif element.element_type == ElementType.TABLE:
                html_parts.append(self._render_table(element))

        full_html = "\n\n".join(html_parts)
        logger.info("Generated HTML content (%d characters)", len(full_html))
        return full_html

    # ── Rendering helpers ────────────────────────────────────────────────────

    def _render_heading(self, element: ContentElement) -> str:
        level = max(1, min(element.level, 6))
        inner = self._render_inline(element.inline_contents)
        return f"<h{level}>{inner}</h{level}>"

    def _render_paragraph(self, element: ContentElement) -> str:
        inner = self._render_inline(element.inline_contents)
        return f"<p>{inner}</p>"

    def _render_bullet_list(self, element: ContentElement) -> str:
        items = "\n".join(
            f"  <li>{self._render_inline(item)}</li>" for item in element.items
        )
        return f"<ul>\n{items}\n</ul>"

    def _render_numbered_list(self, element: ContentElement) -> str:
        items = "\n".join(
            f"  <li>{self._render_inline(item)}</li>" for item in element.items
        )
        return f"<ol>\n{items}\n</ol>"

    def _render_table(self, element: ContentElement) -> str:
        if not element.table_data:
            return ""

        parts: List[str] = ["<table>"]

        for row_idx, row in enumerate(element.table_data):
            if row_idx == 0 and element.has_header_row:
                parts.append("  <thead>")
                parts.append("    <tr>")
                for cell in row:
                    parts.append(f"      <th>{html.escape(cell)}</th>")
                parts.append("    </tr>")
                parts.append("  </thead>")
                parts.append("  <tbody>")
            else:
                parts.append("    <tr>")
                for cell in row:
                    parts.append(f"      <td>{html.escape(cell)}</td>")
                parts.append("    </tr>")

        if element.has_header_row:
            parts.append("  </tbody>")

        parts.append("</table>")
        return "\n".join(parts)

    # ── Inline rendering ─────────────────────────────────────────────────────

    def _render_inline(self, contents: List[InlineContent]) -> str:
        """Render a list of InlineContent segments to an HTML string."""
        parts: List[str] = []

        for content in contents:
            if content.is_image and content.image_id and self.parser:
                # Handle image
                image_tag = self._render_image_tag(content.image_id)
                if image_tag:
                    parts.append(image_tag)
                continue

            text = html.escape(content.text)

            if content.hyperlink_url:
                url = html.escape(content.hyperlink_url, quote=True)
                text = f'<a href="{url}">{text}</a>'
            else:
                if content.bold:
                    text = f"<strong>{text}</strong>"
                if content.italic:
                    text = f"<em>{text}</em>"

            parts.append(text)

        return "".join(parts)

    def _render_image_tag(self, r_id: str) -> Optional[str]:
        """Convert image context to a base64 <img> tag."""
        if not self.parser:
            return None

        import base64
        data, ext = self.parser.get_image_payload(r_id)
        if data:
            b64 = base64.b64encode(data).decode('utf-8')
            mime = f"image/{ext.lstrip('.') if ext else 'png'}"
            return f'<img src="data:{mime};base64,{b64}" style="max-width:100%; height:auto; margin: 1rem 0; display:block; border-radius:8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">'
        return None
