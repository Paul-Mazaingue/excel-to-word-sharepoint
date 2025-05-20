FROM python:3.10-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create temp directory for file storage
RUN mkdir -p temp

# Install rclone & pandoc & wkhtmltopdf
RUN apt-get update && \
    apt-get install -y curl unzip pandoc wkhtmltopdf && \
    curl https://rclone.org/install.sh | bash && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PYTHONUNBUFFERED=1

# Run the application
CMD ["python", "main.py"]
