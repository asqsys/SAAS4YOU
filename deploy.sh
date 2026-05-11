#!/bin/bash

# Configuration
APP_NAME="facturegen-app"
PORT=8080

echo "🚀 Starting deployment for $APP_NAME..."

# 1. Stop and remove existing container if it exists
if podman ps -aq -f name=$APP_NAME | grep -q .; then
    echo "Stopping and removing existing container..."
    podman stop $APP_NAME
    podman rm $APP_NAME
fi

# 2. Build the new image
echo "Building new image..."
podman build -t $APP_NAME:latest .

# 3. Run the new container
echo "Starting new container on port $PORT..."
podman run -d --name $APP_NAME -p $PORT:80 $APP_NAME:latest

echo "✅ Deployment complete! Access your app at http://localhost:$PORT"
