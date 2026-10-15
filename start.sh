#!/bin/bash
echo ""
echo "========================================="
echo "  Test Plan Generator - Starting..."
echo "========================================="
echo ""

# Check Node.js
if ! command -v node &> /dev/null; then
    echo "ERROR: Node.js not installed."
    echo "Install from: https://nodejs.org"
    exit 1
fi

echo "Node.js: $(node --version)"
echo ""

# Install dependencies if node_modules missing
if [ ! -d "node_modules" ]; then
    echo "Installing dependencies..."
    npm install
    echo ""
fi

echo "Starting on http://localhost:3377"
echo "Press Ctrl+C to stop"
echo ""

node server.js
