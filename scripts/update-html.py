#!/usr/bin/env python3
"""
Update HTML file with new scenario data from JSON
"""

import sys
import json
import re
from pathlib import Path

# File paths
JSON_FILE = Path('data/scenarios.json')
HTML_FILE = Path('nooks_updated__7_.html')
OUTPUT_HTML = Path('index.html')  # Final output for deployment


def load_json_data():
    """Load extracted scenario data"""
    
    if not JSON_FILE.exists():
        print(f"❌ JSON file not found: {JSON_FILE}")
        sys.exit(1)
    
    with open(JSON_FILE, 'r') as f:
        return json.load(f)


def generate_js_scenarios(data):
    """Generate JavaScript SCENARIOS object from JSON data"""
    
    # This will replace the hardcoded SCENARIOS object in the HTML
    js_code = "const SCENARIOS = {\n"
    
    for key, scenario in data.items():
        js_code += f"  {key}: {{\n"
        js_code += f"    name: '{scenario['name']}',\n"
        
        # Add description and other metadata (keeping original structure)
        if key == 'ops':
            js_code += f"    desc: 'No Series B raise · Current buildout plan · Conservative organic growth · FY29E target {scenario['rev29']} revenue',\n"
            js_code += f"    raise: 'No Raise',\n"
            js_code += f"    raiseColor: 'var(--muted2)',\n"
        elif key == 'bridge':
            js_code += f"    desc: '$20M equity + $15M add-on ($35M total) raised Jan 2027 · Moderate expansion · Huntsville + 2 new sites · FY29E target {scenario['rev29']} revenue',\n"
            js_code += f"    raise: '$35M Raise',\n"
            js_code += f"    raiseColor: 'var(--green)',\n"
        elif key == 'full':
            js_code += f"    desc: '$80M equity + $30M debt raised Jan 2027 · Aggressive expansion · 6 new sites · FY29E target {scenario['rev29']} revenue',\n"
            js_code += f"    raise: '$95M Raise',\n"
            js_code += f"    raiseColor: 'var(--gold2)',\n"
        
        # Add KPI values
        js_code += f"    rev26: '{scenario['rev26']}', rev29: '{scenario['rev29']}', margin: '{scenario['margin']}',\n"
        
        # Add metadata
        js_code += f"    capital: '—', capitalSub: 'No Series B',\n"
        js_code += f"    rev26Sub: 'Current ops estimate',\n"
        js_code += f"    rev29Sub: '<span class=\"kpi-delta pos\">{scenario['cagr']} CAGR</span>',\n"
        js_code += f"    marginSub: 'FY29E target',\n"
        js_code += f"    chartSub: 'FY23A — FY29E · {scenario['name']} · No Series B raise',\n"
        js_code += f"    cashSub: 'Monthly FY26 · No capital raise · Organic cash trajectory',\n"
        js_code += f"    tableTitle: '{scenario['name']} — 5-Year P&L Summary',\n"
        
        # Add data arrays
        js_code += f"    revenue:  {json.dumps(scenario['revenue'])},\n"
        js_code += f"    ebitda:   {json.dumps(scenario['ebitda'])},\n"
        js_code += f"    payroll:  {json.dumps(scenario['payroll'])},\n"
        js_code += f"    margin:   {json.dumps(scenario['margin_array'])},\n"
        
        # Cash balance (keeping placeholder for now - can be extracted from Excel later)
        js_code += f"    cashBal:  [10265, 6563, 3403, 4940, 4679, 4786, 4789, 6882, 2426, 3156, 2797, 1554],\n"
        
        # Colors
        if key == 'ops':
            js_code += f"    color: 'rgba(10,22,40,0.7)',\n"
            js_code += f"    borderColor: '#0a1628',\n"
        elif key == 'bridge':
            js_code += f"    color: 'rgba(30,107,58,0.65)',\n"
            js_code += f"    borderColor: '#1e6b3a',\n"
        elif key == 'full':
            js_code += f"    color: 'rgba(201,168,76,0.75)',\n"
            js_code += f"    borderColor: '#c9a84c',\n"
        
        js_code += "  },\n"
    
    js_code += "};"
    
    return js_code


def update_html(html_content, new_js_scenarios):
    """Replace SCENARIOS object in HTML"""
    
    # Pattern to match the SCENARIOS object
    # Matches from "const SCENARIOS = {" to the closing "};" with everything in between
    pattern = r'const SCENARIOS = \{[^}]*(?:\{[^}]*\}[^}]*)*\};'
    
    # Replace with new data
    updated_html = re.sub(
        pattern,
        new_js_scenarios,
        html_content,
        flags=re.DOTALL
    )
    
    # Check if replacement happened
    if updated_html == html_content:
        print("⚠️  WARNING: SCENARIOS object not found or not replaced")
        return None
    
    return updated_html


def main():
    """Main execution"""
    
    print("=" * 60)
    print("HTML Update")
    print("=" * 60)
    
    # Load JSON data
    print(f"📖 Loading data from: {JSON_FILE}")
    data = load_json_data()
    
    # Generate new JavaScript
    print("🔧 Generating JavaScript...")
    new_js = generate_js_scenarios(data)
    
    # Read HTML
    if not HTML_FILE.exists():
        print(f"❌ HTML file not found: {HTML_FILE}")
        sys.exit(1)
    
    print(f"📖 Reading HTML: {HTML_FILE}")
    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # Update HTML
    print("🔄 Updating SCENARIOS object...")
    updated_html = update_html(html_content, new_js)
    
    if updated_html is None:
        print("❌ Failed to update HTML")
        sys.exit(1)
    
    # Write output
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(updated_html)
    
    print(f"✅ Updated HTML saved to: {OUTPUT_HTML}")
    print("\n📊 Summary:")
    for key, scenario in data.items():
        print(f"  ✓ {scenario['name']} updated")
    
    return 0


if __name__ == '__main__':
    sys.exit(main())
