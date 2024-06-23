#!/bin/bash

# Define the parent directory containing the virtual environment
PARENT_DIR="/workspaces/codespaces-blank"
VENV_NAME="env" # The name of the virtual environment directory

# Define the Python script to run
SCRIPT="car_wash_updater/carwash/weekly_sender/weekly_sender.py"

# Define the log file with date and timestamp
LOG_DIR="$PARENT_DIR/logs"
TIMESTAMP=$(date +"%Y-%m-%d_%H-%M-%S")
LOG_FILE="$LOG_DIR/log_$TIMESTAMP.txt"

# Create the log directory if it doesn't exist
mkdir -p "$LOG_DIR"

# Check if the virtual environment is already activated
if [[ -z "$VIRTUAL_ENV" ]]; then
    echo "Virtual environment is not active. Activating..." | tee -a "$LOG_FILE"
    source "$PARENT_DIR/$VENV_NAME/bin/activate"
else
    echo "Virtual environment is already active." | tee -a "$LOG_FILE"
fi

# Set the PYTHONPATH to include the parent directory
export PYTHONPATH="$PARENT_DIR/carwash"

# Ensure required modules are installed
pip install -r "$PARENT_DIR/requirements.txt" >> "$LOG_FILE" 2>&1

# Run the Python script and log the output
echo "Running $SCRIPT..." | tee -a "$LOG_FILE"
python "$PARENT_DIR/$SCRIPT" >> "$LOG_FILE" 2>&1
echo "Finished running $SCRIPT" | tee -a "$LOG_FILE"



echo "Script has been executed. Logs can be found in $LOG_FILE" | tee -a "$LOG_FILE"
