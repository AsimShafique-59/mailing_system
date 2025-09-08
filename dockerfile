# Use lightweight Python
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies required for reportlab/pdfrw
RUN apt-get update && apt-get install -y \
    build-essential \
    libxml2-dev \
    libxslt1-dev \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all source code into container
COPY . .

# Default command to run your agent
CMD ["python", "sign.py"]
