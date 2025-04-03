# Generated by https://smithery.ai. See: https://smithery.ai/docs/config#dockerfile
FROM python:3.10-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy all files into the container
COPY . /app

# Create a virtual environment and install Python dependencies
# Using pip install instead of editing setup items
RUN python -m venv /opt/venv \
    && . /opt/venv/bin/activate \
    && pip install --upgrade pip \
    && pip install .

# Expose necessary ports if any (MCP over stdio, so likely not needed)

# Set environment PATH to include virtual environment binaries
ENV PATH="/opt/venv/bin:$PATH"

# Set the entry point to run the MCP server
CMD ["python", "run.py"]
