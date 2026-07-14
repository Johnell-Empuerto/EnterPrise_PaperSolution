"""
Test the fixed reader on all 4 workbooks.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from render_service.excel_cluster_reader import read_fields

DOCS = os.path.expanduser("~\\Documents")

workbooks = [
    ("Text by HandWriting",  os.path.join(DOCS, "Text by HandWriting.xlsx")),
    ("[V3.1_Sample]",         os.path.join(DOCS, "[V3.1_Sample]アンケート用紙.xlsx")),
    ("FormTest - Copy",       os.path.join(DOCS, "FormTest - Copy.xlsx")),
    ("Sample A",              os.path.join(DOCS, "Sample A.xlsx")),
]

passed = 0
failed = 0

for label, path in workbooks:
    print(f"\n{'='*60}")
    print(f"Testing: {label}")
    print(f"  Path: {path}")
    
    if not os.path.isfile(path):
        print(f"  ⚠ FILE NOT FOUND")
        failed += 1
        continue
    
    try:
        result = read_fields(path)
        print(f"  Fields found: {len(result)}")
        for fd in result:
            print(f"    {fd.addr}: type={fd.data_type} "
                  f"L={fd.ratio_left:.4f} T={fd.ratio_top:.4f} "
                  f"R={fd.ratio_right:.4f} B={fd.ratio_bottom:.4f}")
        
        if len(result) > 0:
            print(f"  ✅ PASSED")
            passed += 1
        else:
            print(f"  ❌ FAILED — no fields found")
            failed += 1
    except Exception as e:
        print(f"  ❌ ERROR: {e}")
        failed += 1

print(f"\n{'='*60}")
print(f"RESULTS: {passed} passed, {failed} failed out of {len(workbooks)}")
