# Presentation Editor

A web application for creating and editing PowerPoint presentations with AI assistance.

## Features

- Upload existing PowerPoint presentations
- Create new presentations
- Add slides with AI-generated or manual content
- Save presentations with all changes
- Real-time preview of slides
- Interactive chat interface

## Setup

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Set up environment variables:
Create a `.env` file in the root directory with:
```
OPENAI_API_KEY=your_openai_api_key_here
```

3. Start the backend server:
```bash
python server.py
```

4. Open `index.html` in your web browser or serve it using a local web server.

## Usage

1. Upload a PowerPoint file or create a new presentation
2. Use the "Add New Slide" button to add slides
3. Choose between AI-generated or manual content for each slide
4. Save your presentation when done

## File Structure

- `index.html` - Main HTML file
- `styles.css` - CSS styles
- `script.js` - Frontend JavaScript
- `server.py` - Backend Flask server
- `requirements.txt` - Python dependencies
- `uploads/` - Directory for uploaded and saved presentations

## Requirements

- Python 3.7+
- Modern web browser
- OpenAI API key 