import sys, re

with open('LibExcelController.il', 'r', encoding='utf-8') as f:
    content = f.read()

# Find the class bodies for key types
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
    if '`' in cls_name:
        cls_name = cls_name.split('`')[0] + '(generic)'
    class_data.append((cls_name, class_body))

# Extract ALL strings from key classes
for cls_name, body in class_data:
    if not any(t in cls_name for t in ['ExcelControllerBase', 'ExcelControllerInterop', 'CalculateCell', 'ExcelRangeBase']):
        continue
    
    print(f'\n{"="*80}')
    print(f'CLASS: {cls_name} ({len(body)} bytes)')
    print(f'{"="*80}')
    
    # All method names in this class
    method_names = re.findall(r'\.method\s+(?:\S+\s+)*(\S+)\s*\(', body)
    interesting = [m for m in method_names if not m.startswith('get_') and not m.startswith('set_') and not m.startswith('<')]
    print(f'Methods ({len(interesting)}):')
    for m in interesting:
        print(f'  {m}')
    
    # All string literals
    strings = re.findall(r'"([^"]*)"', body)
    relevant_strings = [s for s in strings if len(s) > 2 and not s.startswith('<') and not s.startswith('_')]
    print(f'\nStrings:')
    for s in sorted(set(relevant_strings)):
        print(f'  "{s}"')
    
    # All property/field accesses on COM/Infragistics objects  
    # Extract get_ calls that indicate which properties are read
    get_calls = re.findall(r'callvirt\s+instance\s+\S+\s+(\S+)::get_(\w+)', body)
    if get_calls:
        print(f'\nProperty reads:')
        for obj, prop in sorted(set(get_calls)):
            print(f'  {obj}::{prop}')
    
    # Call instructions - look for specific method calls  
    method_calls = re.findall(r'callvirt\s+instance\s+\S+\s+(\S+)::(\w+)', body)
    notable_calls = [(obj, m) for obj, m in method_calls 
                     if any(kw in m.lower() for kw in ['comment', 'note', 'field', 'shape', 'cluster', 'print', 'page', 'cell', 'range', 'value', 'visible'])]
    if notable_calls:
        print(f'\nNotable method calls:')
        for obj, m in sorted(set(notable_calls))[:20]:
            print(f'  {obj}::{m}')
