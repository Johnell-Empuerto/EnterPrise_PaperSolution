import sys, re

with open('LibExcelController.il', 'r', encoding='utf-8') as f:
    content = f.read()

# Find all class declarations with their full bodies
# Approach: find .class declaration, track braces to get full body
class_start_pattern = re.compile(r'\.class\s+(?:.+?)')

# Find all positions of class declarations
positions = []
for m in re.finditer(r'\.class\s', content):
    start = m.start()
    # Find class name
    line_end = content.find('\n', start)
    header = content[start:line_end if line_end > 0 else start+200]
    
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
    
    # Extract class name from header
    name_match = re.search(r'class\s+(?:\S+\s+)*(\S+)', header)
    cls_name = name_match.group(1).strip('"') if name_match else 'unknown'
    if '/' in cls_name:
        cls_name = cls_name.split('/')[-1]
    if '`' in cls_name:
        cls_name = cls_name.split('`')[0] + '(generic)'
    
    positions.append((start, cls_name, class_body))

# Filter to key classes
key_classes = ['ExcelControllerInterop', 'ExcelControllerBase', 
               'ExcelWorkbookCom', 'ExcelWorkbookBase',
               'ExcelWorksheetCom', 'ExcelWorksheetBase',
               'ExcelRangeCom', 'ExcelRangeBase',
               'CalculateCell', 'ImageUtility',
               'ClusterRect', 'ClusterType',
               'ExcelWorkbookInfra', 'ExcelWorksheetInfra', 'ExcelRangeInfra']

for start_pos, cls_name, body in positions:
    if not any(k in cls_name for k in key_classes):
        continue
    
    print(f'\n{"="*60}')
    print(f'CLASS: {cls_name} ({len(body)} bytes)')
    print(f'{"="*60}')
    
    # Extract .method declarations
    methods = re.findall(r'\.method\s+(.+?)(?=\.method|\n\s*\nclass|\Z)', body, re.DOTALL)
    
    for m in methods:
        m = m.strip()
        if not m:
            continue
        first_line = m.split('\n')[0].strip()[:200]
        
        # Check for special methods
        if 'get_' in first_line or 'set_' in first_line:
            continue  # skip property accessors for brevity
        
        print(f'\n  METHOD: {first_line}')
        
        # Extract all string literals in this method
        strings = re.findall(r'"([^"]*)"', m)
        if strings:
            for s in strings:
                if len(s) > 1 and not s.startswith('<'):
                    print(f'    string: "{s}"')
        
        # Extract all call instructions
        calls = re.findall(r'call(?:virt)?\s+(?:instance|class|void|string|object|bool|int32|float64)[^;]*', m)
        # Extract call to specific interesting types
        interesting_calls = [c for c in calls if any(t in c for t in 
            ['Comment', 'Notes', 'Threaded', 'Shape', 'Text', 
             'Range', 'Left', 'Top', 'Width', 'Height',
             'Merge', 'Visible', 'Value', 'Address',
             'PrintArea', 'PageSetup', 'UsedRange',
             'Cells', 'Count', 'Item', 'Field',
             'CommentText'])]
        if interesting_calls:
            print(f'    calls:')
            for c in interesting_calls[:10]:
                print(f'      {c.strip()[:150]}')
