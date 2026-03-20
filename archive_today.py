"""
Run this at the start of each session to snapshot today's version into archive/.
Usage: python archive_today.py
"""
import shutil
import os
from datetime import datetime

today = datetime.now().strftime("%Y-%m-%d")
dest = os.path.join("archive", today)

if os.path.exists(dest):
    print(f"[OK] Archive for {today} already exists at {dest}/")
else:
    os.makedirs(dest)
    shutil.copy2("wcl_craft_audit.py", os.path.join(dest, "wcl_craft_audit.py"))
    print(f"[OK] Archived today's version to {dest}/")

print("You're good to start making changes.")
