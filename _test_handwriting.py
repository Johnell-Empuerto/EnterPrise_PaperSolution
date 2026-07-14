"""
Test Text by HandWriting specifically with the updated reader.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from render_service.excel_cluster_reader import read_fields

path = os.path.expanduser("~\\Documents\\Text by HandWriting.xlsx")
print(f"Path: {path}")
print(f"Exists: {os.path.isfile(path)}")

result = read_fields(path)
print(f"Fields found: {len(result)}")
for fd in result:
    print(f"  {fd.addr}: type={fd.data_type} "
          f"L={fd.ratio_left:.4f} T={fd.ratio_top:.4f} "
          f"R={fd.ratio_right:.4f} B={fd.ratio_bottom:.4f}")
