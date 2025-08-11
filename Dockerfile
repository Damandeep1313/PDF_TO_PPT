# Use official Python image as base
FROM python:3.12-slim

# Set working directory inside container
WORKDIR /app

# Copy requirements if you have one, else install directly
COPY requirements.txt .

# Install dependencies (including Flask, python-pptx, openai, cloudinary, requests, dotenv, tqdm)
RUN pip install --no-cache-dir -r requirements.txt

# Copy all app files into container
COPY . .

# Expose port 5000 (Flask default)
EXPOSE 5000

# Set environment variables (optional: better to pass at runtime)
ENV FLASK_APP=server.py
ENV FLASK_RUN_HOST=0.0.0.0

# Run Flask server
CMD ["flask", "run", "--host=0.0.0.0", "--port=5000"]
