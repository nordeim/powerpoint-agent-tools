#!/bin/bash
set -e

# Get the root directory of the project
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )/.." && pwd )"

echo "🐳 Building PowerPoint Agent Tools Docker Image..."
#docker build -t ppt-agent-tools:latest -f "$DIR/docker/Dockerfile" "$DIR"
DIR=`pwd` &&  docker build -t ppt-agent-tools:latest -f "$DIR/docker/Dockerfile" "$DIR"

echo "✅ Build Complete. Image: ppt-agent-tools:latest"

