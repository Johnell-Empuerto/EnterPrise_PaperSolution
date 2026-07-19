"""Find our generated output workbook matching FormTest"""
import os, json, glob

forms_dir = 'ExcelAPI/ExcelAPI/Forms'

# Check runtime JSON files for matching form name
for jf in sorted(glob.glob(os.path.join(forms_dir, '*.json'))):
    try:
        data = json.load(open(jf, encoding='utf-8'))
        title = ''
        wb = data.get('workbook') or data.get('Workbook') or {}
        if isinstance(wb, dict):
            title = wb.get('title') or wb.get('Title') or ''
        if isinstance(title, str) and ('FormTest' in title or 'formtest' in title.lower()):
            xlsx = jf.replace('.runtime.json', '.xlsx')
            if os.path.exists(xlsx):
                print(f"MATCH: {jf}")
                print(f"  Title: {title}")
                print(f"  XLSX: {xlsx}")
                print(f"  Size: {os.path.getsize(xlsx)} bytes")
    except:
        pass

# Also check all xlsx files by modification time
print("\nMost recent 10 xlsx files in Forms/:")
for fx in sorted(glob.glob(os.path.join(forms_dir, '*.xlsx')), key=os.path.getmtime, reverse=True)[:10]:
    import hashlib
    with open(fx, 'rb') as f:
        md5 = hashlib.md5(f.read()).hexdigest()[:12]
    print(f"  {os.path.basename(fx)}: {os.path.getsize(fx):>8} bytes, MD5={md5}")
