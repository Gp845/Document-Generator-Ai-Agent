# ── Stage: runtime ────────────────────────────────────────────────────────────
FROM python:3.11-slim

# Keeps Python from buffering stdout/stderr (important for interactive CLI)
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

WORKDIR /app

# Install OS-level dependencies needed by ReportLab (font rendering)
RUN apt-get update && apt-get install -y --no-install-recommends \
        libfreetype6 \
        libfontconfig1 \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies first (layer cache-friendly)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source
COPY main.py toolset.py ./

# Output files land in /app/outputs — mount a host directory here
# so generated .pptx / .xlsx / .pdf files are accessible on the host.
RUN mkdir -p /app/outputs

# Default command — interactive chat loop
CMD ["python", "main.py"]
