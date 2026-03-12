"""
app.py
──────
Web interface for the Content Migration Tool.
Upload a .docx file, preview the converted HTML, and optionally push to Document360.
"""

import os
import json
import logging
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

from docx_parser import DocxParser
from html_converter import HtmlConverter
from document360_client import Document360Client, Document360Config

load_dotenv()

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["UPLOAD_FOLDER"] = os.path.join(os.path.dirname(__file__), "uploads")
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500 MB max

os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def get_d360_config():
    return Document360Config(
        api_token=os.getenv("DOCUMENT360_API_TOKEN", ""),
        base_url=os.getenv("DOCUMENT360_BASE_URL", "https://apihub.document360.io"),
        project_version_id=os.getenv("DOCUMENT360_PROJECT_VERSION_ID", ""),
        category_id=os.getenv("DOCUMENT360_CATEGORY_ID", ""),
        user_id=os.getenv("DOCUMENT360_USER_ID", ""),
    )


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    """Upload a .docx file, parse it, convert to HTML, return the result."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not file.filename.lower().endswith(".docx"):
        return jsonify({"error": "Only .docx files are supported"}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(filepath)

    try:
        parser = DocxParser(filepath)
        elements = parser.parse()

        converter = HtmlConverter()
        html_content = converter.convert(elements, parser)

        # Save the HTML output
        html_path = os.path.join(app.config["UPLOAD_FOLDER"], Path(filename).stem + ".html")
        Path(html_path).write_text(html_content, encoding="utf-8")

        # Build summary stats
        from collections import Counter
        from docx_parser import ElementType
        counts = Counter(e.element_type.name for e in elements)

        return jsonify({
            "success": True,
            "html": html_content,
            "filename": filename,
            "stats": {
                "total_elements": len(elements),
                "headings": counts.get("HEADING", 0),
                "paragraphs": counts.get("PARAGRAPH", 0),
                "bullet_lists": counts.get("BULLET_LIST", 0),
                "numbered_lists": counts.get("NUMBERED_LIST", 0),
                "tables": counts.get("TABLE", 0),
            },
            "html_size": len(html_content),
        })
    except Exception as e:
        logger.exception("Conversion failed")
        return jsonify({"error": str(e)}), 500


@app.route("/discover-versions", methods=["POST"])
def discover_versions():
    """Fetch project versions using the provided API token."""
    data = request.get_json()
    api_token = data.get("api_token")
    
    if not api_token:
        return jsonify({"error": "API token is required"}), 400
        
    config = get_d360_config()
    config.api_token = api_token
    
    try:
        client = Document360Client(config)
        result = client.get_project_versions()
        return jsonify({"success": True, "data": result})
    except Exception as e:
        logger.exception("Discovery failed")
        return jsonify({"error": str(e)}), 500


@app.route("/discover-categories", methods=["POST"])
def discover_categories():
    """Fetch categories for a specific project version."""
    data = request.get_json()
    api_token = data.get("api_token")
    version_id = data.get("version_id")
    
    if not api_token or not version_id:
        return jsonify({"error": "API token and Version ID are required"}), 400
        
    config = get_d360_config()
    config.api_token = api_token
    
    try:
        client = Document360Client(config)
        result = client.get_categories(version_id)
        return jsonify({"success": True, "data": result})
    except Exception as e:
        logger.exception("Category discovery failed")
        return jsonify({"error": str(e)}), 500


@app.route("/upload-to-d360", methods=["POST"])
def upload_to_d360():
    """Upload the converted HTML to Document360 with Smart Discovery."""
    data = request.get_json()
    title = data.get("title", "Untitled Article")
    html_content = data.get("html", "")
    
    # Credentials from UI or Environment
    api_token = data.get("api_token") or os.getenv("DOCUMENT360_API_TOKEN")
    user_id = data.get("user_id") or os.getenv("DOCUMENT360_USER_ID")
    category_id = data.get("category_id") or os.getenv("DOCUMENT360_CATEGORY_ID")
    project_version_id = data.get("project_version_id") or os.getenv("DOCUMENT360_PROJECT_VERSION_ID")

    if not html_content:
        return jsonify({"error": "No HTML content to upload"}), 400

    if not api_token or api_token == "your_api_token_here":
        return jsonify({"error": "Document360 API token is not configured. Please enter it in the Auth settings."}), 400

    config = Document360Config(
        api_token=api_token,
        user_id=user_id,
        category_id=category_id,
        project_version_id=project_version_id
    )

    try:
        client = Document360Client(config)
        
        # --- Smart Discovery ---
        # 1. If Version ID is missing, find the default version
        if not config.project_version_id:
            logger.info("Project Version ID missing. Discovering default version...")
            versions_res = client.get_project_versions()
            versions = versions_res.get("data", [])
            # Try to find default, otherwise take first
            target_v = next((v for v in versions if v.get("is_default")), None)
            if not target_v and versions:
                target_v = versions[0]
            
            if target_v:
                config.project_version_id = target_v["id"]
                logger.info("Found Version: %s", target_v.get("version_number"))
            else:
                return jsonify({"error": "Could not find any Project Versions. Please check your API Key."}), 400

        # 2. If Category ID is missing, find the first category
        if not config.category_id:
            logger.info("Category ID missing. Discovering first category for version %s...", config.project_version_id)
            categories_res = client.get_categories(config.project_version_id)
            categories = categories_res.get("data", [])
            if categories:
                config.category_id = categories[0]["id"]
                logger.info("Found Category: %s", categories[0].get("name"))
            else:
                return jsonify({"error": f"No categories found for version {config.project_version_id}. Please create a category in Document360 first."}), 400

        # --- Final Upload ---
        result = client.create_article(title=title, html_content=html_content)
        return jsonify({"success": True, "response": result})
        
    except Exception as e:
        logger.exception("Upload to Document360 failed")
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
