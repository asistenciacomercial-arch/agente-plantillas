#!/bin/bash

echo "=== DEBUG START ==="
pwd
ls -la

echo "=== STARTING UVICORN ==="
uvicorn app:app --host 0.0.0.0 --port $PORT
