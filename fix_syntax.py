import sys

with open("UserCopyCreate.ps1", "r") as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if line.strip() == '# Try to load ActiveDirectory module':
        new_lines.append("if ($MyInvocation.InvocationName -ne '.') {\n")
        new_lines.append(line)
    else:
        new_lines.append(line)

new_lines.append("            }\n")
new_lines.append("        }\n")
new_lines.append("    } catch {\n")
new_lines.append("        Write-CustomLog \"Fehler beim Erstellen des Benutzers `$Name: `$_\" -Level \"FEHLER\"\n")
new_lines.append("    }\n")
new_lines.append("}\n")

with open("UserCopyCreate.ps1", "w") as f:
    f.writelines(new_lines)
