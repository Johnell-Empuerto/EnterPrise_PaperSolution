import sys, re

with open('LibExcelController.il', 'r', encoding='utf-8') as f:
    content = f.read()

# Find ALL class declarations
class_pattern = re.compile(
    r'\.class\s+(?:.+?)\b(class|interface|value)\s+(\S+)', 
    re.DOTALL
)
matches = list(class_pattern.finditer(content))
print(f'Total classes found: {len(matches)}')

for m in matches:
    cls_name = m.group(2).strip('"')
    
    # Clean up the name
    if '/' in cls_name:
        cls_name = cls_name.split('/')[-1]
    if '`' in cls_name:
        cls_name = cls_name.split('`')[0] + '(generic)'
    
    print(f'\nClass: {cls_name}')
    print(f'  Raw: {m.group(0)[:150].strip()}')
