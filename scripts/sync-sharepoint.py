#!/usr/bin/env python3
"""
SharePoint Excel Sync Script
Downloads the latest Excel file from SharePoint folder
"""

import os
import sys
import hashlib
from pathlib import Path
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# Configuration - Update these values
SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')  # e.g., https://yourcompany.sharepoint.com/sites/Finance
SHAREPOINT_FOLDER = os.getenv('SHAREPOINT_FOLDER')      # e.g., Shared Documents/Financial Models
EXCEL_FILENAME = os.getenv('EXCEL_FILENAME', 'nooks_q2_model.xlsx')  # Exact filename or pattern
USERNAME = os.getenv('SHAREPOINT_USERNAME')
PASSWORD = os.getenv('SHAREPOINT_PASSWORD')

# Local paths
OUTPUT_DIR = Path('data')
OUTPUT_FILE = OUTPUT_DIR / EXCEL_FILENAME
HASH_FILE = OUTPUT_DIR / '.last_hash'


def get_file_hash(filepath):
    """Calculate MD5 hash of file to detect changes"""
    if not filepath.exists():
        return None
    
    hash_md5 = hashlib.md5()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


def download_from_sharepoint():
    """Download Excel file from SharePoint"""
    
    # Validate environment variables
    if not all([SHAREPOINT_SITE_URL, SHAREPOINT_FOLDER, USERNAME, PASSWORD]):
        print("❌ Missing required environment variables:")
        print("   SHAREPOINT_SITE_URL, SHAREPOINT_FOLDER, SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD")
        sys.exit(1)
    
    print(f"📡 Connecting to SharePoint: {SHAREPOINT_SITE_URL}")
    
    try:
        # Authenticate
        credentials = UserCredential(USERNAME, PASSWORD)
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credentials)
        
        # Build file path
        file_relative_url = f"{SHAREPOINT_FOLDER}/{EXCEL_FILENAME}"
        print(f"📥 Downloading: {file_relative_url}")
        
        # Download file
        response = File.open_binary(ctx, file_relative_url)
        
        # Ensure output directory exists
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        
        # Write to local file
        with open(OUTPUT_FILE, 'wb') as local_file:
            local_file.write(response.content)
        
        print(f"✅ Downloaded to: {OUTPUT_FILE}")
        return True
        
    except Exception as e:
        print(f"❌ Error downloading from SharePoint: {e}")
        print(f"\nTroubleshooting:")
        print(f"  1. Check SHAREPOINT_SITE_URL: {SHAREPOINT_SITE_URL}")
        print(f"  2. Check SHAREPOINT_FOLDER: {SHAREPOINT_FOLDER}")
        print(f"  3. Check file exists: {file_relative_url}")
        print(f"  4. Verify credentials have read access")
        return False


def check_if_changed():
    """Check if downloaded file is different from last sync"""
    
    current_hash = get_file_hash(OUTPUT_FILE)
    
    if not HASH_FILE.exists():
        # First run - save hash
        print("🆕 First sync - no previous version to compare")
        HASH_FILE.write_text(current_hash)
        return True
    
    last_hash = HASH_FILE.read_text().strip()
    
    if current_hash == last_hash:
        print("⏭️  File unchanged since last sync")
        return False
    else:
        print("📝 File changed - updating hash")
        HASH_FILE.write_text(current_hash)
        return True


def main():
    """Main sync workflow"""
    
    print("=" * 60)
    print("SharePoint → GitHub Sync")
    print("=" * 60)
    
    # Download from SharePoint
    success = download_from_sharepoint()
    if not success:
        sys.exit(1)
    
    # Check if file changed
    changed = check_if_changed()
    
    if changed:
        print("\n✅ File updated - proceed with data extraction")
        sys.exit(0)  # Signal to GitHub Actions to continue
    else:
        print("\n⏭️  No changes detected - skipping deployment")
        sys.exit(1)  # Signal to GitHub Actions to skip remaining steps


if __name__ == '__main__':
    main()
