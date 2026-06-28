# Use the official lightweight Python image.
FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PORT=10000

# Install system dependencies for OCR and PDF handling
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    poppler-utils \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set work directory
WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Download the spaCy language model
RUN python -m spacy download en_core_web_sm

# Copy the application code
COPY . .

# Expose the default port
EXPOSE 10000

# Run gunicorn to serve the Flask app, binding to the port set by the PORT env variable
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:$PORT app:app"]
