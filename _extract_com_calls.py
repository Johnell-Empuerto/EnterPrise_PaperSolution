import sys, re

with open('LibExcelController.il', 'r', encoding='utf-8') as f:
    content = f.read()

# Find class declarations with their full bodies
# Collect all class bodies and names
class_data = []
for m in re.finditer(r'\.class\s', content):
    start = m.start()
    # Find opening brace
    body_start = content.find('{', start)
    if body_start == -1:
        continue
    
    # Track braces to find end
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
    
    # Extract class name
    name_match = re.search(r'class\s+(?:\S+\s+)*(\S+)', header)
    cls_name = name_match.group(1).strip('"') if name_match else 'unknown'
    if '/' in cls_name:
        cls_name = cls_name.split('/')[-1]
    if '`' in cls_name:
        cls_name = cls_name.split('`')[0] + '(generic)'
    
    class_data.append((cls_name, class_body))

# Focus on key interop ranges
targets = ['ExcelRangeCom', 'ExcelRangeBase', 'ExcelRangeInfra',
           'ExcelWorksheetCom', 'ExcelWorksheetBase',
           'ExcelWorkbookCom', 'ExcelWorkbookBase',
           'CalculateCell', 'ExcelControllerBase', 'ExcelControllerInterop']

for cls_name, body in class_data:
    if not any(t in cls_name for t in targets):
        continue
    
    print(f'\n{"="*80}')
    print(f'CLASS: {cls_name} ({len(body)} bytes)')
    print(f'{"="*80}')
    
    # Extract method declarations with their bodies
    # Find .method declarations and their bodies
    for method_match in re.finditer(r'(\.method\s+(?:[^{}]+?\{)(?:[^{}]*(?:\{[^{}]*\}[^{}]*)*)\})', body, re.DOTALL):
        method_body = method_match.group(0)
        
        # Get method name
        name_match2 = re.search(r'\.method\s+(?:\S+\s+)*(\S+)\s*\(', method_body)
        if not name_match2:
            continue
        method_name = name_match2.group(1)
        
        if method_name.startswith('get_') or method_name.startswith('set_'):
            continue
        
        # Check if this method calls interesting COM APIs
        interesting_keywords = [
            'Comment', 'Note', 'Threaded', 'Shape', 'Text',
            'Left', 'Top', 'Width', 'Height', 'Visible',
            'MergeCell', 'MergeArea', 'Address', 'UsedRange',
            'PrintArea', 'PageSetup', 'Cells', 'Range',
            'Col', 'Row', 'Column', 'Value', 'Clear',
            'Borders', 'Interior', 'Font', 'Format',
            'CommentText', 'get_Comment', 'CommentThreaded'
        ]
        
        calls_found = []
        for kw in interesting_keywords:
            for call_match in re.finditer(r'call(?:virt)?\s+.*?' + re.escape(kw) + r'[^;]*', method_body, re.IGNORECASE):
                calls_found.append(call_match.group(0).strip()[:150])
        
        if calls_found:
            print(f'\n  METHOD: {method_name}')
            for c in calls_found[:8]:
                print(f'    COM CALL: {c}')
