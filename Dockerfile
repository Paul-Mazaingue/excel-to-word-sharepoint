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

# Create directory for rclone config
RUN mkdir -p /root/.config/rclone

# Set environment variables
ENV PYTHONUNBUFFERED=1

# Create entrypoint script
RUN echo '#!/bin/bash\n\
# Check if rclone.conf exists in the host and copy it\n\
if [ -f /tmp/rclone.conf ]; then\n\
  echo "Found rclone.conf, copying to container"\n\
  cp /tmp/rclone.conf /root/.config/rclone/rclone.conf\n\
fi\n\
\n\
# Execute the main command\n\
exec "$@"' > /app/entrypoint.sh && chmod +x /app/entrypoint.sh

# Set the entrypoint
ENTRYPOINT ["/app/entrypoint.sh"]

# Run the application
CMD ["python", "main.py"]
