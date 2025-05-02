// DOM Elements
const chatMessages = document.getElementById('chatMessages');
const messageInput = document.getElementById('messageInput');
const sendButton = document.getElementById('sendMessage');
const loadPresentationBtn = document.getElementById('loadPresentation');
const newPresentationBtn = document.getElementById('newPresentation');
const addSlideBtn = document.getElementById('addSlide');
const savePresentationBtn = document.getElementById('savePresentation');
const fileInput = document.getElementById('fileInput');
const presentationPreview = document.getElementById('presentation-preview');

// State
let currentPresentation = null;
let presentationState = {
    slides: [],
    currentSlideIndex: -1,
    awaitingInput: {
        type: null, // 'title', 'content', 'content_choice'
        slideIndex: null
    }
};

// Event Listeners
sendButton.addEventListener('click', handleSendMessage);
messageInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSendMessage();
    }
});

loadPresentationBtn.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', handleFileUpload);
newPresentationBtn.addEventListener('click', handleNewPresentation);
addSlideBtn.addEventListener('click', handleAddSlide);
savePresentationBtn.addEventListener('click', handleSavePresentation);

// Auto-resize textarea
messageInput.addEventListener('input', () => {
    messageInput.style.height = 'auto';
    messageInput.style.height = messageInput.scrollHeight + 'px';
});

// Handle file upload
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (!file.name.match(/\.(ppt|pptx)$/i)) {
        addMessage('Please select a valid PowerPoint file (.ppt or .pptx)', 'system');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('http://localhost:5001/api/upload', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error('Upload failed');
        }

        const data = await response.json();
        currentPresentation = {
            name: file.name,
            size: formatFileSize(file.size),
            lastModified: new Date(file.lastModified).toLocaleString()
        };

        presentationState = {
            slides: [],
            currentSlideIndex: -1,
            awaitingInput: {
                type: null,
                slideIndex: null
            }
        };

        updatePresentationPreview();
        addMessage(`Successfully loaded presentation: ${file.name}`, 'system');
    } catch (error) {
        console.error('Error uploading file:', error);
        addMessage('Error uploading presentation. Please try again.', 'system');
    }
}

// Handle new presentation
function handleNewPresentation() {
    currentPresentation = {
        name: 'New Presentation',
        size: '0 KB',
        lastModified: new Date().toLocaleString()
    };

    presentationState = {
        slides: [],
        currentSlideIndex: -1,
        awaitingInput: {
            type: null,
            slideIndex: null
        }
    };

    updatePresentationPreview();
    addMessage('Created new presentation', 'system');
}

// Handle add slide
function handleAddSlide() {
    if (!currentPresentation) {
        addMessage('Please create or load a presentation first', 'system');
        return;
    }

    const newSlide = {
        title: '',
        content: '',
        index: presentationState.slides.length
    };

    presentationState.slides.push(newSlide);
    presentationState.currentSlideIndex = newSlide.index;
    presentationState.awaitingInput = {
        type: 'content_choice',
        slideIndex: newSlide.index
    };

    updatePresentationPreview();
    addMessage("How would you like to create the slide content?", 'system');
    addMessage("1. Type 'AI' to generate content using AI\n2. Type 'MANUAL' to write your own content", 'system');
}

// Handle save presentation
async function handleSavePresentation() {
    if (!currentPresentation || presentationState.slides.length === 0) {
        addMessage('No presentation to save', 'system');
        return;
    }

    try {
        const response = await fetch('http://localhost:5001/api/save-presentation', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                slides: presentationState.slides,
                filename: currentPresentation.name
            })
        });

        if (!response.ok) {
            throw new Error('Save failed');
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'presentation_edited.pptx';
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        addMessage('Presentation saved successfully!', 'system');
    } catch (error) {
        console.error('Error saving presentation:', error);
        addMessage('Error saving presentation. Please try again.', 'system');
    }
}

// Handle sending messages
function handleSendMessage() {
    const message = messageInput.value.trim();
    if (!message) return;

    addMessage(message, 'user');
    messageInput.value = '';
    messageInput.style.height = 'auto';

    processMessage(message);
}

// Process the user's message
async function processMessage(message) {
    if (presentationState.awaitingInput.type) {
        await handleSlideInput(message);
        return;
    }
}

// Handle slide input
async function handleSlideInput(message) {
    const { type, slideIndex } = presentationState.awaitingInput;
    const slide = presentationState.slides[slideIndex];

    if (type === 'content_choice') {
        if (message.toUpperCase() === 'AI') {
            addMessage("Please enter a topic or title for the AI to generate content:", 'system');
            presentationState.awaitingInput.type = 'ai_title';
        } else if (message.toUpperCase() === 'MANUAL') {
            addMessage("Please enter the title for the slide:", 'system');
            presentationState.awaitingInput.type = 'title';
        } else {
            addMessage("Please type either 'AI' or 'MANUAL' to proceed.", 'system');
        }
    } else if (type === 'ai_title') {
        try {
            const response = await fetch('http://localhost:5001/api/generate-content', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ title: message })
            });

            if (!response.ok) {
                throw new Error('Content generation failed');
            }

            const data = await response.json();
            slide.title = message;
            slide.content = formatContentToBulletPoints(data.content);
            
            presentationState.awaitingInput = {
                type: null,
                slideIndex: null
            };
            
            updatePresentationPreview();
            addMessage("Slide has been created with AI-generated content!", 'system');
        } catch (error) {
            console.error('Error generating content:', error);
            addMessage("Error generating content. Please try manual input by typing 'MANUAL'.", 'system');
            presentationState.awaitingInput.type = 'content_choice';
        }
    } else if (type === 'title') {
        slide.title = message;
        addMessage("Please enter the content for the slide (use * for bullet points):", 'system');
        addMessage("Example:\n* Point 1: Description for point 1\n* Point 2: Description for point 2\n* Point 3: Description for point 3", 'system');
        presentationState.awaitingInput.type = 'content';
    } else if (type === 'content') {
        slide.content = formatContentToBulletPoints(message);
        presentationState.awaitingInput = {
            type: null,
            slideIndex: null
        };
        updatePresentationPreview();
        addMessage("Slide has been created successfully!", 'system');
    }
}

// Helper function to format content into bullet points
function formatContentToBulletPoints(content) {
    // Split content into lines
    const lines = content.split('\n');
    
    // Process each line
    const bulletPoints = lines.map(line => {
        // Remove any existing bullet points or numbers
        line = line.replace(/^[\d\.\-\*]+/, '').trim();
        
        // Skip empty lines
        if (!line) return '';
        
        // Split the line into title and content if it contains a colon
        const [title, ...contentParts] = line.split(':');
        const content = contentParts.join(':').trim();
        
        return {
            title: title.trim(),
            content: content || title.trim() // If no content after colon, use the whole line as content
        };
    }).filter(item => item.title !== ''); // Remove empty lines
    
    return bulletPoints;
}

// Update presentation preview
function updatePresentationPreview() {
    if (!currentPresentation) return;

    let previewHTML = `
        <div class="presentation-info">
            <i class="fas fa-file-powerpoint"></i>
            <h4>${currentPresentation.name}</h4>
            <p>Size: ${currentPresentation.size}</p>
            <p>Last Modified: ${currentPresentation.lastModified}</p>
            <div class="slides-preview">
                <h5>Slides (${presentationState.slides.length})</h5>
                <div class="slides-list">
    `;

    presentationState.slides.forEach((slide, index) => {
        previewHTML += `
            <div class="slide-item ${index === presentationState.currentSlideIndex ? 'active' : ''}">
                <span class="slide-number">${index + 1}</span>
                <span class="slide-title">${slide.title || 'Untitled'}</span>
            </div>
        `;
    });

    previewHTML += `
                </div>
            </div>
        </div>
    `;

    presentationPreview.innerHTML = previewHTML;
}

// Add a message to the chat
function addMessage(content, type) {
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${type}`;
    
    const messageContent = document.createElement('div');
    messageContent.className = 'message-content';
    messageContent.innerHTML = content;
    
    messageDiv.appendChild(messageContent);
    chatMessages.appendChild(messageDiv);
    
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

// Format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
} 
