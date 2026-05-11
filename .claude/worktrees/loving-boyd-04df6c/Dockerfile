FROM python:3.10-slim

# Install system dependencies
# libxml2-dev and libxslt-dev might be needed for some lxml/bs4 operations if not pre-built
# netcat is useful for healthchecks (waiting for db)
RUN apt-get update && apt-get install -y \
    gcc \
    python3-dev \
    netcat-openbsd \
    tini \
    perl \
    default-jre \
    cpanminus \
    make \
    libxml-libxml-perl \
    libarchive-zip-perl \
    libfile-copy-recursive-perl \
    libtry-tiny-perl \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Perl dependencies
RUN cpanm --notest \
    File::HomeDir \
    String::Substitution

# Copy requirements first for cache efficiency
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install gunicorn

# Copy the rest of the application
COPY . .

# Create necessary directories
RUN mkdir -p logs reports S4C-Processed-Documents S4c-Macros

# Expose the port Gunicorn will run on
EXPOSE 8000

# Run with Gunicorn using Tini as entrypoint
ENTRYPOINT ["/usr/bin/tini", "--"]
CMD ["gunicorn", "-c", "gunicorn_config.py", "app_server:app"]
