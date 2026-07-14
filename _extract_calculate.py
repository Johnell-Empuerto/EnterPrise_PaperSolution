import sys, re

with open('LibExcelController.il', 'r', encoding='utf-8') as f:
    content = f.read()

# Find CalculateCell class
idx = content.find('CalculateCell')
if idx < 0:
    print('CalculateCell not found!')
    sys.exit(1)

# Go back to the .class declaration
class_start = content.rfind('.class', 0, idx)
body_start = content.find('{', class_start)
depth = 0
i = body_start
while i < len(content):
    if content[i] == '{': depth += 1
    elif content[i] == '}':
        depth -= 1
        if depth == 0: break
    i += 1

body = content[class_start:i+1]
print(f'CalculateCell class ({len(body)} bytes)')

# ALL field declarations (instance variables)
fields = re.findall(r'\.field\s+(?:public|private|family|assembly)\s+(?:\S+\s+)*(\S+)', body)
print(f'\nFields:')
for f in fields:
    print(f'  {f}')

# ALL method declarations
methods = re.findall(r'\.method\s+(?:.+?)\b(\w+)\s*\(', body)
print(f'\nMethods ({len(methods)}):')
for m in sorted(set(methods)):
    print(f'  {m}')

# ALL string literals
strings = re.findall(r'"([^"]*)"', body)
print(f'\nStrings:')
for s in sorted(set(strings)):
    if len(s) > 1 and not s.startswith('<'):
        print(f'  "{s}"')

# ALL call and callvirt instructions
calls = re.findall(r'call(?:virt)?\s+[^;\n]*', body)
print(f'\nCalls ({len(calls)}):')
for c in calls[:30]:
    print(f'  {c.strip()[:150]}')
