"""
document360_client.py
─────────────────────
Handles uploading HTML content to Document360 using their v2 API.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from typing import Any, Dict, Optional

import requests

logger = logging.getLogger(__name__)


@dataclass
class Document360Config:
    """Configuration for the Document360 API."""
    api_token: str
    base_url: str = "https://apihub.document360.io"
    project_version_id: str = ""
    category_id: str = ""
    user_id: str = ""


class Document360Client:
    """Client for interacting with the Document360 API."""

    def __init__(self, config: Document360Config) -> None:
        self.config = config
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "application/json",
            "api_token": config.api_token,
        })

    def create_article(
        self,
        title: str,
        html_content: str,
        category_id: Optional[str] = None,
        project_version_id: Optional[str] = None,
        user_id: Optional[str] = None,
        order: int = 0,
    ) -> Dict[str, Any]:
        """
        Create a new article in Document360.

        Args:
            title: Article title.
            html_content: The HTML body of the article.
            category_id: Override the configured category ID.
            project_version_id: Override the configured project version ID.
            user_id: Override the configured user ID.
            order: Article position within the category (default 0).

        Returns:
            The JSON response from the API.

        Raises:
            requests.HTTPError: If the API returns an error status code.
        """
        url = f"{self.config.base_url.rstrip('/')}/v2/Articles"

        payload = {
            "title": title,
            "content": html_content,
            "category_id": category_id or self.config.category_id,
            "project_version_id": project_version_id or self.config.project_version_id,
            "user_id": user_id or self.config.user_id,
            "order": order,
            "content_type": 1,  # 1 = WYSIWYG editor (HTML)
        }

        logger.info("Sending POST request to %s", url)
        logger.debug("Payload: %s", json.dumps(payload, indent=2)[:500])

        try:
            response = self.session.post(url, json=payload, timeout=30)
            response.raise_for_status()

            result = response.json()
            logger.info(
                "✅ Article created successfully! Status: %d", response.status_code
            )
            logger.info(
                "   Article ID: %s",
                result.get("data", {}).get("id", "N/A"),
            )
            return result

        except requests.exceptions.HTTPError as e:
            logger.error(
                "❌ API Error: %s – %s",
                response.status_code,
                response.text[:300],
            )
            raise
        except requests.exceptions.ConnectionError as e:
            logger.error("❌ Connection error: %s", e)
            raise
        except requests.exceptions.Timeout:
            logger.error("❌ Request timed out")
            raise
        except requests.exceptions.RequestException as e:
            logger.error("❌ Unexpected request error: %s", e)
            raise

    def get_project_versions(self) -> Dict[str, Any]:
        """Retrieve available project versions (useful for initial setup)."""
        url = f"{self.config.base_url.rstrip('/')}/v2/ProjectVersions"
        response = self.session.get(url, timeout=30)
        response.raise_for_status()
        return response.json()

    def get_categories(self, project_version_id: str) -> Dict[str, Any]:
        """Retrieve categories for a project version (useful for initial setup)."""
        url = (
            f"{self.config.base_url.rstrip('/')}/v2/ProjectVersions/"
            f"{project_version_id}/categories"
        )
        params = {
            "excludeArticles": True,
            "langCode": "en",
            "includeCategoryDescription": False,
        }
        response = self.session.get(url, params=params, timeout=30)
        response.raise_for_status()
        return response.json()
