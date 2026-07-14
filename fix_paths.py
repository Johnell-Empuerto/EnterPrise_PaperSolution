import os, re

cwd = os.getcwd()
print("CWD:", cwd)

fp = os.path.join(cwd, "phase11j18_investigation.py")
print("Reading:", fp)

with open(fp, "r", encoding="utf-8") as f:
    content = f.read()

# Replace the hardcoded OUT path with dynamic one
old_out = r'C:\Users\MCF-JOHNELLEEMUERTO\Documents\Johnell\PaperLess Enterprise\Investigation_546'
content = content.replace(old_out, 'os.path.join(BASE_DIR, "Investigation_546")')

old_pdf = r'C:\Users\MCF-JOHNELLEEMUERTO\Documents\Johnell\PaperLess Enterprise\pdf_extracts'
content = content.replace(old_pdf, 'os.path.join(BASE_DIR, "pdf_extracts")')

with open(fp, "w", encoding="utf-8") as f:
    f.write(content)

print("Fixed paths in:", fp)
