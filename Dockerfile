FROM python:3.11-slim

# Install poppler for PDF processing
RUN apt-get update && apt-get install -y \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Create directories
RUN mkdir -p uploads outputs

# Expose port
EXPOSE 5001

# Run
CMD ["python", "app.py"]
