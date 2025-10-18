#!/bin/bash

# Create Streamlit directory
mkdir -p ~/.streamlit/

# Create config.toml for Heroku
cat > ~/.streamlit/config.toml <<EOF
[theme]
primaryColor = "#667eea"
backgroundColor = "#ffffff"
secondaryBackgroundColor = "#f0f2f6"
textColor = "#31333F"
font = "sans serif"

[client]
showErrorDetails = true

[logger]
level = "info"

[server]
headless = true
port = \$PORT
enableCORS = false
maxUploadSize = 200
EOF

echo "Streamlit configured for Heroku"
