"""
Flask server for AI-powered PowerPoint generation.
"""
import logging
import base64
import os
import io

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

from converters.json_to_ppt import generate_presentation_stream, PPTGenerationError
from generators.llm_generator import generate_presentation_structure

app = Flask(__name__, static_folder='static')
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'templates', 'template.pptx')

# API Key from environment variable (optional at startup)
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
if not OPENAI_API_KEY:
    logger.warning("OPENAI_API_KEY not set. Generation will fail until configured.")

IMAGES_PATH = os.path.join(os.path.dirname(__file__), 'images')

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')


@app.route('/images/<path:filename>')
def serve_image(filename):
    return send_from_directory(IMAGES_PATH, filename)


@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        req_data = request.json
        if not req_data:
            return jsonify({"error": "No data provided"}), 400

        # Case 1: Generate from topic
        if 'topic' in req_data:
            if not OPENAI_API_KEY:
                return jsonify({"error": "OPENAI_API_KEY not configured. Set it with: export OPENAI_API_KEY='your-key'"}), 500
            topic = req_data['topic']
            slide_count = req_data.get('slide_count', 8)
            language = req_data.get('language', 'en')
            
            # Validate slide_count (5-15)
            if not isinstance(slide_count, int) or slide_count < 5 or slide_count > 15:
                slide_count = max(5, min(15, int(slide_count) if slide_count else 8))
            
            logger.info(f"Generating presentation for: {topic} (slides: {slide_count}, lang: {language})")
            
            # Call LLM generator
            ppt_structure = generate_presentation_structure(topic, OPENAI_API_KEY, slide_count, language)
            
            # Pydantic -> Dict
            json_data = ppt_structure.model_dump()
            
        # Case 2: Generate from JSON directly
        elif 'slides' in req_data:
            json_data = req_data
        else:
            return jsonify({"error": "Invalid request. Provide 'topic' or 'slides'."}), 400

        # Check template exists
        if not os.path.exists(TEMPLATE_PATH):
            logger.error(f"Template not found: {TEMPLATE_PATH}")
            return jsonify({"error": f"Template missing: {TEMPLATE_PATH}"}), 500

        # Generate PPTX
        with open(TEMPLATE_PATH, 'rb') as f:
            template_stream = io.BytesIO(f.read())

        logger.info("Generating PPTX...")
        ppt_stream = generate_presentation_stream(json_data, template_stream)
        logger.info("Generation complete")
        ppt_base64 = base64.b64encode(ppt_stream.read()).decode('utf-8')

        filename = f"{json_data.get('title', 'presentation')}.pptx"
        filename = "".join(c for c in filename if c.isalnum() or c in ' -_.').strip()

        return jsonify({
            "filename": filename, 
            "file_base64": ppt_base64,
            "structure": json_data
        })

    except PPTGenerationError as e:
        logger.warning(f"Generation error: {e.message}")
        return jsonify({"error": e.message, "details": e.details}), 400
    except FileNotFoundError as e:
        logger.exception("File not found during generation")
        return jsonify({"error": f"File not found: {str(e)}"}), 500
    except Exception as e:
        logger.exception("Unexpected server error")
        return jsonify({"error": str(e)}), 500


@app.route('/health')
def health():
    return jsonify({
        "status": "ok", 
        "template": os.path.exists(TEMPLATE_PATH),
        "api_key_configured": bool(OPENAI_API_KEY)
    })


if __name__ == '__main__':
    debug = os.getenv('FLASK_DEBUG', 'false').lower() == 'true'
    app.run(debug=debug, host='0.0.0.0', port=5000)
