import sys, re

with open('LibExcelController.il', 'r', encoding='utf-8') as f:
    content = f.read()

# Find class bodies
class_data = []
for m in re.finditer(r'\.class\s', content):
    start = m.start()
    body_start = content.find('{', start)
    if body_start == -1:
        continue
    depth = 0
    i = body_start
    while i < len(content):
        if content[i] == '{':
            depth += 1
        elif content[i] == '}':
            depth -= 1
            if depth == 0:
                break
        i += 1
    class_body = content[start:i+1]
    header = content[start:body_start][:300]
    name_match = re.search(r'class\s+(?:\S+\s+)*(\S+)', header)
    cls_name = name_match.group(1).strip('"') if name_match else 'unknown'
    if '/' in cls_name:
        cls_name = cls_name.split('/')[-1]
    class_data.append((cls_name, class_body))

targets = ['ExcelControllerInterop', 'CalculateCell', 'ExcelRangeBase', 'ExcelRangeCom', 'ExcelRangeInfra']

for cls_name, body in class_data:
    if not any(t in cls_name for t in targets):
        continue
    
    print(f'\n{"="*60}')
    print(f'CLASS: {cls_name} ({len(body)} bytes)')
    print(f'{"="*60}')
    
    # All method declarations
    methods = re.findall(r'\.method\s+(?:\S+\s+)*(\S+)\s*\(', body)
    all_interesting = [m for m in methods if not m.startswith('get_') and not m.startswith('set_') and not m.startswith('<')]
    all_props = [m for m in methods if m.startswith('get_') or m.startswith('set_')]
    print(f'Methods ({len(all_interesting)}):')
    for m in all_interesting:
        print(f'  {m}')
    print(f'Properties ({len(all_props)}):')
    for m in all_props:
        print(f'  {m}')
    
    # Extract ALL call instructions for comment/field related methods
    all_calls = re.findall(r'call(?:virt)?\s+(?:instance|class|void|string|object|bool|int32|float64)[^;\n]*', body)
    comment_calls = [c for c in all_calls if any(kw.lower() in c.lower() for kw in 
        ['comment', 'field', 'cluster', 'note', 'shape', 'text', 'value', 'print',
         'merge', 'address', 'cell', 'sheet', 'workbook', 'excel', 'interior', 'border',
         'font', 'align', 'column', 'row', 'range', 'visible'])]
    
    if comment_calls:
        print(f'\nNotable calls ({len(comment_calls)}):')
        for c in comment_calls[:20]:
            print(f'  {c.strip()[:180]}')
