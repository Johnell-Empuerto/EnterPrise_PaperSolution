"""
Test the fixed reader on FormTest - Copy.xlsx.
Calls read_fields() directly and prints results.
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from render_service.excel_cluster_reader import read_fields

result = read_fields("FormTest - Copy.xlsx")
print(f"\n=== RESULTS ===")
print(f"Fields found: {len(result)}")
for fd in result:
    print(f"  {fd.addr}: type={fd.data_type} "
          f"L={fd.ratio_left:.4f} T={fd.ratio_top:.4f} "
          f"R={fd.ratio_right:.4f} B={fd.ratio_bottom:.4f}")
print(f"\nTest {'PASSED' if len(result) > 0 else 'FAILED'}")
