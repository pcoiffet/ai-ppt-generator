# AI PowerPoint Generator

An AI-powered web application that generates professional PowerPoint presentations from natural language descriptions.

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Flask](https://img.shields.io/badge/Flask-3.0+-green.svg)
![OpenAI](https://img.shields.io/badge/OpenAI-GPT--4-orange.svg)

## Features

- **Natural Language Input**: Describe your presentation topic and let AI structure it
- **Multiple Content Types**: Automatically generates slides with text, bullet points, tables, charts, and images
- **Customizable Output**: Choose slide count (5-15) and language (English/French)
- **Live Preview**: Preview slides before downloading
- **Professional Templates**: Uses PowerPoint templates for consistent styling
- **Image Integration**: Fetches relevant images from Unsplash (optional)

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         Frontend (HTML/JS)                       │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                       Flask API Server                           │
└─────────────────────────────────────────────────────────────────┘
                              │
            ┌─────────────────┴─────────────────┐
            ▼                                   ▼
┌───────────────────────┐           ┌───────────────────────┐
│   LangChain + GPT-4   │           │   python-pptx         │
│   (Structure Gen)     │           │   (PPTX Generation)   │
└───────────────────────┘           └───────────────────────┘
```

## Quick Start

### Prerequisites

- Python 3.10+
- OpenAI API key

### Installation

1. Clone the repository:

```bash
git clone https://github.com/yourusername/ai-ppt-generator.git
cd ai-ppt-generator
```

2. Create a virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# or
venv\Scripts\activate     # Windows
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Set up environment variables:

```bash
cp .env.example .env
# Edit .env and add your OpenAI API key
```

5. Run the server:

```bash
python server.py
```

6. Open http://localhost:5000 in your browser

## Configuration

### Environment Variables

| Variable              | Required | Description                  |
| --------------------- | -------- | ---------------------------- |
| `OPENAI_API_KEY`      | Yes      | Your OpenAI API key          |
| `UNSPLASH_ACCESS_KEY` | No       | Unsplash API key for images  |
| `FLASK_DEBUG`         | No       | Set to `true` for debug mode |

### PowerPoint Template

The application uses a PowerPoint template located at `templates/template.pptx`. The template must include these slide layouts:

- Title Slide
- Content Only
- Image Right
- Image Left
- Image Full
- Table
- Chart
- Two Columns

## API Endpoints

### POST /generate-ppt

Generate a presentation from a topic.

**Request:**

```json
{
  "topic": "The future of renewable energy",
  "slide_count": 8,
  "language": "en"
}
```

**Response:**

```json
{
  "filename": "The future of renewable energy.pptx",
  "file_base64": "UEsDBBQ...",
  "structure": { ... }
}
```

### GET /health

Health check endpoint.

## Project Structure

```
├── server.py                 # Flask application
├── schemas.py                # Pydantic models
├── generators/
│   └── llm_generator.py      # LangChain/OpenAI integration
├── converters/
│   └── json_to_ppt.py        # JSON to PPTX conversion
├── static/
│   ├── index.html            # Frontend
│   └── styles.css            # Styles
├── templates/
│   └── template.pptx         # PowerPoint template
├── images/
│   └── placeholder.jpg       # Fallback image
├── requirements.txt
└── README.md
```

## Technologies

- **Backend**: Flask, Flask-CORS
- **AI**: LangChain, OpenAI GPT-4
- **PPT Generation**: python-pptx
- **Validation**: Pydantic
- **Images**: Unsplash API (optional)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
