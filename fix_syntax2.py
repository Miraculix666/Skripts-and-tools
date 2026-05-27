import sys

with open("UserCopyCreate.ps1", "r") as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if line.strip() == 'throw "Das ActiveDirectory-Modul wird benötigt."':
        new_lines.append('        # throw "Das ActiveDirectory-Modul wird benötigt."\n')
    else:
        new_lines.append(line)

with open("UserCopyCreate.ps1", "w") as f:
    f.writelines(new_lines)
