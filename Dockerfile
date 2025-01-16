# Use the official Python image from the Docker Hub
FROM python:3.12

# Set the working directory in the container
WORKDIR /app

# Install LibreOffice
RUN apt-get update && apt-get install -y libreoffice && apt-get clean
# Copy the requirements and install dependencies
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code
COPY . .

# Expose the port that the app runs on
EXPOSE 8501

# Run the application with gunicorn
CMD ["gunicorn", "-b", "0.0.0.0:8080", "api:app"]
