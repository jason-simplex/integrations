#!/usr/bin/env bash

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"

# Define the plugin path (pointing to the directory CONTAINING the plugin folder)
# In this case, the plugin folder is 'connect-plugin', so we point to the project root.
PLUGIN_PATH="${SCRIPT_DIR}"

# Export the environment variable
export FTRACK_CONNECT_PLUGIN_PATH="$PLUGIN_PATH"

echo "================================================================"
echo "Environment Configured:"
echo "FTRACK_CONNECT_PLUGIN_PATH = $FTRACK_CONNECT_PLUGIN_PATH"
echo "================================================================"

# Path to the ftrack-connect repository
CONNECT_APP_DIR="/Users/jason/fwork/repos/integrations/apps/connect"

if [ ! -d "$CONNECT_APP_DIR" ]; then
    echo "Error: Directory not found: $CONNECT_APP_DIR"
    exit 1
fi

# Change to the connect app directory
cd "$CONNECT_APP_DIR" || exit

echo "Launching ftrack-connect..."
# Run the application using poetry
poetry run python -m ftrack_connect
