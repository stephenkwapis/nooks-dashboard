#!/bin/bash
# Local development script - watches Excel file for changes and auto-updates

set -e

EXCEL_FILE="${1:-data/*.xlsx}"
OUTPUT_JSON="${2:-public/data/financial_data.json}"

echo "🔍 Watching for changes to Excel files..."
echo "   Excel: $EXCEL_FILE"
echo "   Output: $OUTPUT_JSON"
echo ""
echo "Press Ctrl+C to stop"
echo ""

# Initial extraction
python3 scripts/extract_data.py "$EXCEL_FILE" "$OUTPUT_JSON"

# Watch for changes (requires fswatch or inotifywait)
if command -v fswatch &> /dev/null; then
    # macOS
    fswatch -o "$EXCEL_FILE" | while read; do
        echo "📊 Excel file changed, extracting data..."
        python3 scripts/extract_data.py "$EXCEL_FILE" "$OUTPUT_JSON"
        echo "✓ Data updated at $(date)"
        
        # Auto-commit if in git repo
        if [ -d .git ]; then
            echo "📝 Committing changes..."
            git add "$OUTPUT_JSON"
            git commit -m "Update financial data - $(date '+%Y-%m-%d %H:%M')" || true
            echo "✓ Ready to push with: git push"
        fi
    done
elif command -v inotifywait &> /dev/null; then
    # Linux
    while inotifywait -e close_write "$EXCEL_FILE"; do
        echo "📊 Excel file changed, extracting data..."
        python3 scripts/extract_data.py "$EXCEL_FILE" "$OUTPUT_JSON"
        echo "✓ Data updated at $(date)"
        
        # Auto-commit if in git repo
        if [ -d .git ]; then
            echo "📝 Committing changes..."
            git add "$OUTPUT_JSON"
            git commit -m "Update financial data - $(date '+%Y-%m-%d %H:%M')" || true
            echo "✓ Ready to push with: git push"
        fi
    done
else
    echo "❌ Please install fswatch (brew install fswatch) or inotifywait"
    exit 1
fi
