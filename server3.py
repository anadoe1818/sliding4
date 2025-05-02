from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import os
from dotenv import load_dotenv
import openai
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import tempfile
import shutil
import json

# Load environment variables
load_dotenv()

app = Flask(__name__, static_folder='sliding')
CORS(app)

# Configure OpenAI
openai.api_key = os.getenv('OPENAI_API_KEY')
if not openai.api_key:
    print("Warning: OPENAI_API_KEY not found in environment variables")

UPLOAD_FOLDER = 'uploads'

@app.route('/')
def serve_index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/<path:path>')
def serve_static(path):
    return send_from_directory(app.static_folder, path)

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and file.filename.endswith(('.ppt', '.pptx')):
        filename = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filename)
        return jsonify({'message': 'File uploaded successfully', 'filename': file.filename})
    
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/api/generate-content', methods=['POST'])
def generate_content():
    if not openai.api_key:
        return jsonify({'error': 'OpenAI API key not configured'}), 500

    data = request.json
    title = data.get('title', '')
    
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant that creates concise bullet points for presentations. Each bullet point should have a short title and brief content (maximum 2 lines) separated by a colon. Never generate more than 5 bullet points."},
                {"role": "user", "content": f"Create 3-5 brief bullet points for a presentation about: {title}. Format each point as 'Title: Content'. Keep content concise and to the point."}
            ],
            max_tokens=150  # Limit response length
        )
        content = response.choices[0].message.content
        return jsonify({'content': content})
    except Exception as e:
        print(f"Error generating content: {str(e)}")
        return jsonify({'error': 'Failed to generate content'}), 500

@app.route('/api/save-presentation', methods=['POST'])
def save_presentation():
    try:
        data = request.json
        slides = data.get('slides', [])
        filename = data.get('filename', 'presentation.pptx')
        
        prs = Presentation()
        # Set default slide dimensions (16:9 aspect ratio)
        prs.slide_width = Inches(16)
        prs.slide_height = Inches(9)
        
        # Create the first slide for the presentation title
        title_slide_layout = prs.slide_layouts[6]  # Blank layout
        title_slide = prs.slides.add_slide(title_slide_layout)
        
        # Create title box
        title_box = title_slide.shapes.add_shape(
            1,  # Rectangle shape
            Inches(0.61),  # Left position
            Inches(2.16),  # Top position
            Inches(8.7),   # Width
            Inches(1.46)   # Height
        )
        title_box.fill.background()  # No fill
        title_box.line.fill.background()  # No border
        
        # Add title text
        title_frame = title_box.text_frame
        title_frame.word_wrap = True
        title_paragraph = title_frame.paragraphs[0]
        title_paragraph.alignment = PP_ALIGN.LEFT
        title_run = title_paragraph.add_run()
        title_run.text = filename.replace('.pptx', '').title()  # Capitalize every word
        title_run.font.name = 'Arial'
        title_run.font.size = Pt(40)
        title_run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Create presenter box
        presenter_box = title_slide.shapes.add_shape(
            1,  # Rectangle shape
            Inches(0.61),  # Left position
            Inches(4.72),  # Top position
            Inches(7.48),  # Width
            Inches(0.3)    # Height
        )
        presenter_box.fill.background()  # No fill
        presenter_box.line.fill.background()  # No border
        
        # Add presenter text
        presenter_frame = presenter_box.text_frame
        presenter_frame.word_wrap = True
        presenter_paragraph = presenter_frame.paragraphs[0]
        presenter_paragraph.alignment = PP_ALIGN.LEFT
        presenter_run = presenter_paragraph.add_run()
        presenter_run.text = "Presenter"
        presenter_run.font.name = 'Arial'
        presenter_run.font.size = Pt(16)
        presenter_run.font.color.rgb = RGBColor(0, 0, 0)
        
        for slide_data in slides:
            # Always use blank layout (layout 6)
            slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(slide_layout)
            
            # Create title shape at the top
            title_shape = slide.shapes.add_shape(
                1,  # Rectangle shape
                Inches(1),  # Left position
                Inches(0.5),  # Top position
                Inches(14),  # Width
                Inches(1)  # Height
            )
            title_shape.fill.background()  # No fill
            title_shape.line.fill.background()  # No border
            
            # Add title text
            title_frame = title_shape.text_frame
            title_frame.word_wrap = True
            title_paragraph = title_frame.paragraphs[0]
            title_paragraph.alignment = PP_ALIGN.LEFT
            title_run = title_paragraph.add_run()
            title_run.text = slide_data.get('title', '')
            title_run.font.size = Pt(44)
            title_run.font.color.rgb = RGBColor(0, 0, 0)
            
            # Process content
            content = slide_data.get('content', [])
            if isinstance(content, str):
                # Split content into lines and filter out empty lines
                content = [line.strip() for line in content.split('\n') if line.strip()]
                # Convert each line into a title:content format if it's not already
                content = [{'title': line.split(':')[0].strip(), 'content': line.split(':')[1].strip() if ':' in line else ''} for line in content]
                # Limit to 5 items
                content = content[:5]
            create_list_boxes(slide, content)
        
        # Create uploads directory if it doesn't exist
        if not os.path.exists(UPLOAD_FOLDER):
            os.makedirs(UPLOAD_FOLDER)
        
        # Save the presentation
        output_path = os.path.join(UPLOAD_FOLDER, 'presentation_edited.pptx')
        prs.save(output_path)
        
        if not os.path.exists(output_path):
            raise Exception("Failed to create presentation file")
        
        return send_from_directory(UPLOAD_FOLDER, 'presentation_edited.pptx', as_attachment=True)
    
    except Exception as e:
        print(f"Error saving presentation: {str(e)}")
        return jsonify({'error': f'Failed to save presentation: {str(e)}'}), 500

def create_list_boxes(slide, content):
    """Create boxes for list items."""
    # Filter out any empty content items
    content = [item for item in content if item and (isinstance(item, dict) and (item.get('title') or item.get('content')) or isinstance(item, str) and item.strip())]
    
    num_points = len(content)
    if num_points > 0:
        # Get the actual slide width from the presentation
        slide_width = Inches(16)  # Standard 16:9 slide width
        slide_margin = Inches(1)  # Margin from edges
        
        # Calculate available width for boxes
        available_width = slide_width - (2 * slide_margin)
        
        # Calculate box dimensions
        box_margin = Inches(0.3)  # Margin between boxes
        total_margins = box_margin * (num_points - 1)  # Total space needed for margins
        box_width = (available_width - total_margins) / num_points
        
        # Calculate starting position to center the boxes
        start_x = slide_margin
        start_y = Inches(2.5)  # Position below title
        
        for i, point in enumerate(content):
            # Calculate box position
            left = start_x + (i * (box_width + box_margin))
            top = start_y
            
            # Create box shape
            box = slide.shapes.add_shape(1, left, top, box_width, Inches(2.5))
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(240, 240, 240)
            box.line.fill.background()
            
            # Configure text frame
            text_frame = box.text_frame
            text_frame.word_wrap = True
            text_frame.margin_left = Inches(0.2)
            text_frame.margin_right = Inches(0.2)
            text_frame.margin_top = Inches(0.2)
            text_frame.margin_bottom = Inches(0.2)
            
            # Clear any default paragraphs
            for paragraph in text_frame.paragraphs:
                p = paragraph._element
                p.getparent().remove(p)
            
            if isinstance(point, dict):
                title = point.get('title', '')
                content = point.get('content', '')
                
                # Add title
                p = text_frame.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                run = p.add_run()
                run.text = title
                run.font.size = Pt(16)
                run.font.color.rgb = RGBColor(220, 53, 69)
                run.font.bold = True
                
                # Add content
                if content:
                    p = text_frame.add_paragraph()
                    p.alignment = PP_ALIGN.LEFT
                    run = p.add_run()
                    run.text = content
                    run.font.size = Pt(14)
                    run.font.color.rgb = RGBColor(0, 0, 0)
            else:
                # Add simple text
                p = text_frame.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                run = p.add_run()
                run.text = point
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(0, 0, 0)

def create_flow_chart(slide, content):
    pass

def create_data_visualization(slide, content):
    pass

if __name__ == '__main__':
    if not openai.api_key:
        print("\nWarning: OpenAI API key not found!")
        print("Please add your API key to the .env file:")
        print("OPENAI_API_KEY=your_api_key_here\n")
    app.run(debug=True, port=5001) 
