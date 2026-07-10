"""Analyze exported PDF: extract content bounding box and grid positions"""
import zlib, re, sys, os

cwd = os.getcwd()
pdf_path = os.path.join(cwd, "ExcelAPI", "ExcelAPI", "Preview", "instrument_sheet1.pdf")

with open(pdf_path, "rb") as f:
    raw = f.read()

text = raw.decode("latin-1")

streams = re.findall(r"stream\n(.+?)\nendstream", text, re.DOTALL)
all_drawing = []
page_h = 792.0

for s in streams:
    raw_bytes = s.encode("latin-1")
    try:
        dec = zlib.decompress(raw_bytes)
    except:
        continue
    dec_text = dec.decode("latin-1")
    for line in dec_text.split("\n"):
        line = line.strip()
        if not line: continue
        parts = line.split()
        if len(parts) >= 2:
            last = parts[-1]
            rest = parts[:-1]
            if last in ("m", "l", "re"):
                try:
                    coords = [float(p) for p in rest]
                    all_drawing.append((coords, last))
                except:
                    pass

print("Total drawing commands: %d" % len(all_drawing))
if not all_drawing:
    print("NO DRAWING COMMANDS FOUND")
    sys.exit(0)

xs, ys = [], []
for coords, cmd in all_drawing:
    if cmd == "re":
        xs.extend([coords[0], coords[0] + coords[2]])
        ys.extend([coords[1], coords[1] + coords[3]])
    else:
        xs.append(coords[0]); ys.append(coords[1])

min_x, max_x = min(xs), max(xs)
min_y, max_y = min(ys), max(ys)
w = max_x - min_x; h = max_y - min_y

print("BBox PDF: L=%.4f B=%.4f R=%.4f T=%.4f" % (min_x, min_y, max_x, max_y))
print("Size: %.4f x %.4f pt" % (w, h))
print("Top-left: L=%.4f T=%.4f" % (min_x, page_h - max_y))
print()

moves = [(c[0], c[1]) for c, cmd in all_drawing if cmd == "m" and len(c) >= 2]
lines_l = [(c[0], c[1]) for c, cmd in all_drawing if cmd == "l" and len(c) >= 2]

v_lines = set()
for mx, my in moves:
    for lx, ly in lines_l:
        if abs(lx - mx) < 0.1 and abs(ly - my) > 5:
            v_lines.add(round(mx, 2))

h_lines = set()
for mx, my in moves:
    for lx, ly in lines_l:
        if abs(ly - my) < 0.1 and abs(lx - mx) > 5:
            h_lines.add(round(my, 2))

sorted_v = sorted(v_lines)
sorted_h = sorted(h_lines)
non_zero_v = [v for v in sorted_v if v > 0.5]

print("Vertical lines: " + str(sorted_v))
print("Horizontal lines: " + str(sorted_h))
print()

if non_zero_v and len(sorted_v) >= 2 and len(h_lines) >= 2:
    left = non_zero_v[0]
    right = sorted_v[-1]
    top_pdf = max(h_lines)
    bottom_pdf = min(h_lines)
    top_page = page_h - top_pdf
    
    pdf_w = right - left
    pdf_h = top_pdf - bottom_pdf
    pdf_x = left
    pdf_y = top_page
    
    lm = 51.0236; rm = 51.0236; tm = 53.8583; bm = 53.8583
    pw = 612 - lm - rm; ph = 792 - tm - bm
    pr_w = 192.0; pr_h = 172.8
    com_x = lm + (pw - pr_w) / 2
    com_y = tm + (ph - pr_h) / 2
    
    print("=== COMPARISON TABLE ===")
    print("%-30s %-15s %-15s %-12s" % ("Measurement", "COM (pt)", "PDF (pt)", "Diff (pt)"))
    print("-" * 72)
    print("%-30s %-15.4f %-15.4f %-12.4f" % ("print_area_width", pr_w, pdf_w, pdf_w - pr_w))
    print("%-30s %-15.4f %-15.4f %-12.4f" % ("print_area_height", pr_h, pdf_h, pdf_h - pr_h))
    print("%-30s %-15.4f %-15.4f %-12.4f" % ("origin_left", com_x, pdf_x, pdf_x - com_x))
    print("%-30s %-15.4f %-15.4f %-12.4f" % ("origin_top", com_y, pdf_y, pdf_y - com_y))
    print()
    
    x_off = (pdf_w - pr_w) / 2
    y_off = (pdf_h - pr_h) / 2
    
    print("=== OVERLAY OFFSET PREDICTION ===")
    print("Predicted X: (%.4f - %.4f) / 2 = %.4f pt" % (pdf_w, pr_w, x_off))
    print("Predicted Y: (%.4f - %.4f) / 2 = %.4f pt" % (pdf_h, pr_h, y_off))
    print()
    print("Previously measured fields overlay offset: X=5.28pt, Y=6.00pt")
    off_x = abs(x_off - 5.28)
    off_y = abs(y_off - 6.00)
    print("Match X: %.2fpt off | Match Y: %.2fpt off" % (off_x, off_y))
    if off_x < 1.0 and off_y < 1.0:
        print("=> VERDICT: HYPOTHESIS CONFIRMED")
    else:
        print("=> HYPOTHESIS PARTIALLY CONFIRMED, additional factors may be present")
else:
    print("Could not determine grid lines from drawing commands")
    print("non_zero_v=%s, sorted_v=%s, h_lines=%s" % (str(non_zero_v), str(sorted_v), str(h_lines)))
