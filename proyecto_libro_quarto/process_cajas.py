import os
import re

def process_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.split('\n')
    new_lines = []
    i = 0
    modified = False

    while i < len(lines):
        line = lines[i]
        stripped_line = line.strip()
        
        # 1. Single cell table: | Caja X. Texto... |
        m1 = re.match(r'^\|\s*(Caja\s+\d+.*?)\s*\|$', stripped_line, re.IGNORECASE)
        if m1:
            text = m1.group(1).strip()
            text = text.replace('**', '')
            text = re.sub(r'^(Caja\s+\d+\.?)', r'**\1**', text, flags=re.IGNORECASE)
            
            new_lines.append("::: {.caja-box}")
            new_lines.append(text)
            
            if i + 1 < len(lines) and re.match(r'^\|\s*-+\s*\|$', lines[i+1].strip()):
                new_lines.append("")
                i += 2
                while i < len(lines) and lines[i].strip().startswith('|'):
                    row_m = re.match(r'^\|\s*(.*?)\s*\|$', lines[i].strip())
                    if row_m:
                        row_text = row_m.group(1).replace('**', '')
                        new_lines.append(row_text + "\n")
                    else:
                        new_lines.append(lines[i].replace('**', '') + "\n")
                    i += 1
                new_lines.append(":::")
                modified = True
                continue
            else:
                new_lines.append(":::")
                i += 1
                modified = True
                continue
                
        # 2. Double column table header: | Caja X... | Caja X... |
        m2 = re.match(r'^\|\s*(Caja\s+\d+.*?)\s*\|\s*Caja\s+\d+.*?\s*\|$', stripped_line, re.IGNORECASE)
        if m2:
            text = m2.group(1).strip()
            text = text.replace('**', '')
            text = re.sub(r'^(Caja\s+\d+\.?)', r'**\1**', text, flags=re.IGNORECASE)
            
            new_lines.append("::: {.caja-box}")
            new_lines.append(text)
            new_lines.append("")
            
            if i + 1 < len(lines) and re.match(r'^\|\s*-+\s*\|\s*-+\s*\|$', lines[i+1]):
                new_lines.append("| | |")
                new_lines.append(lines[i+1])
                i += 2
                while i < len(lines) and lines[i].strip().startswith('|'):
                    # Remove bolding from table content as well
                    table_line = lines[i].replace('**', '')
                    new_lines.append(table_line)
                    i += 1
                new_lines.append(":::")
                modified = True
                continue
            else:
                new_lines.append(":::")
                i += 1
                modified = True
                continue

        # 3. Regular paragraph starting with Caja X
        m3 = re.match(r'^(\*\*)?(Caja\s+\d+\.?)(\*\*)?(.*)$', stripped_line, re.IGNORECASE)
        # Avoid matching lines that are just regular text without Caja at start
        if m3 and not stripped_line.startswith('|') and not stripped_line.startswith(':::') and not stripped_line.startswith('#'):
            caja_part = m3.group(2)
            rest_part = m3.group(4)
            
            rest_part = rest_part.replace('**', '')
            
            new_lines.append("::: {.caja-box}")
            new_lines.append(f"**{caja_part}**{rest_part}")
            
            i += 1
            while i < len(lines) and lines[i].strip() != '' and not lines[i].startswith('|') and not lines[i].startswith(':::') and not lines[i].startswith('#'):
                new_lines.append(lines[i].replace('**', ''))
                i += 1
            new_lines.append(":::")
            modified = True
            continue

        new_lines.append(line)
        i += 1

    if modified:
        new_content = '\n'.join(new_lines)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Modificado {filepath}")

if __name__ == "__main__":
    for file in os.listdir('.'):
        if file.endswith('.qmd'):
            process_file(file)
