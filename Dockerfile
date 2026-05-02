FROM python:3.11-slim

# Prevents Python from writing .pyc files and enables stdout/stderr logging
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Install system dependencies
# - build-essential: needed for some Python packages that compile C extensions
# - curl: used by the ECS health check
# - libreoffice: needed by Docx2txtLoader and UnstructuredPowerPointLoader
# - libmagic1: needed by python-magic for file type detection
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    libreoffice \
    libmagic1 \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app source
COPY app/ .

# Expose Streamlit port
EXPOSE 8501

# Health check for ECS
HEALTHCHECK --interval=30s --timeout=10s --start-period=60s --retries=3 \
    CMD curl -f http://localhost:8501/_stcore/health || exit 1

# Run Streamlit
CMD ["streamlit", "run", "main.py", \
     "--server.port=8501", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--browser.gatherUsageStats=false"]
