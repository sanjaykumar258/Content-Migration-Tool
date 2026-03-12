# ⚡ DOCX to Document360 Content Migration Tool

A powerful, high-performance web application designed to convert complex Microsoft Word (.docx) documents into clean, semantic HTML and publish them directly to your **Document360** knowledge base using the v2 API.

Author : Sanjaykumar M

## 🚀 Approach & Architecture

The tool follows a modular architecture to ensure reliability and scalability:

1.  **Robust DOCX Parsing**: Uses a custom `DocxParser` engine built on `python-docx` that recursively iterates through document runs to detect:
    *   **Images**: Supports modern `w:drawing` and legacy `w:pict` (VML) formats.
    *   **Structures**: Handles nested bullet/numbered lists and multi-row tables accurately.
2.  **Semantic HTML Conversion**: Converts parsed elements into article-ready HTML with embedded base64 images, ensuring the output is self-contained and perfectly formatted for Knowledge Bases.
3.  **Swagger-Powered Smart Discovery**: 
    *   **Problem**: Users often only have their API Key and User ID but don't know their internal `ProjectVersionID` or `CategoryID`.
    *   **Solution**: The app automatically queries the Document360 Swagger endpoints (`/v2/ProjectVersions` and `/v2/categories`) to "discover" the default project and the first available category, making the integration plug-and-play.
4.  **Premium User Experience**: A unified Flask-based dashboard featuring a glassmorphic dark theme, real-time conversion statistics, and a dual-view (Rendered vs. Source Code) preview panel.

---

## 🛠️ Technology Stack

*   **Backend**: Python 3.10+
*   **Web Framework**: Flask
*   **Parsing Tools**: `python-docx`, `lxml`, `Pillow`
*   **API Integration**: `requests` (Document360 v2 API)
*   **Frontend**: Semantic HTML5, Vanilla CSS3 (Modern Glassmorphism), JavaScript (ES6+)
*   **Configuration**: `python-dotenv`

---

## 🏎️ Steps to Run the Application

### 1. Prerquisites
Ensure you have Python 3 installed on your system.

### 2. Install Dependencies
Clone the repository and install the required libraries:
```bash
pip install -r requirements.txt
```

### 3. Configuration
Rename `.env.example` to `.env` or create a new `.env` file in the root directory:
```env
DOCUMENT360_API_TOKEN=your_token_here
DOCUMENT360_USER_ID=your_user_id_here
DOCUMENT360_BASE_URL=https://apihub.document360.io
```
*Note: The app will automatically discover your Version and Category IDs using Swagger endpoints.*

### 4. Launch the App
Run the Flask server:
```bash
python app.py
```
Open your browser and navigate to: **[http://localhost:5000](http://localhost:5000)**

---

## 📈 Key Features
*   ✅ **500MB Upload Limit**: Optimized for large technical manuals.
*   ✅ **Image Auto-Embedding**: Base64 encoding for seamless cross-platform display.
*   ✅ **Tabbed Preview**: Switch between "Rendered View" and "Clean HTML Source".
*   ✅ **Swagger Console**: Real-time logging of API responses (Articles, Categories, etc.).
