import sys

with open('app.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3570, len(lines)):
    line = lines[i]
    if 'Detalle ARBA' in line:
        line = line.replace("'Detalle ARBA'", "f'Detalle {organismo}'")
        line = line.replace('"Detalle ARBA"', 'f"Detalle {organismo}"')
        
    if 'Total ARBA' in line:
        if "f'Total ARBA" in line or "f\"Total ARBA" in line:
            line = line.replace("Total ARBA", "Total {organismo}")
        else:
            line = line.replace("'Total ARBA", "f'Total {organismo}").replace('"Total ARBA', 'f"Total {organismo}')
            
    if 'CRUCE ARBA' in line:
        line = line.replace('CRUCE ARBA', 'CRUCE {organismo}')
        
    if 'Diferencia (ARBA - Mendez)' in line:
        line = line.replace("'Diferencia (ARBA - Mendez)'", "f'Diferencia ({organismo} - Mendez)'")
        line = line.replace('"Diferencia (ARBA - Mendez)"', 'f"Diferencia ({organismo} - Mendez)"')
            
    if 'Falta en ARBA' in line:
        line = line.replace('Falta en ARBA', 'Falta en {organismo}')
        
    if 'según ARBA' in line:
        line = line.replace('según ARBA', 'según {organismo}')
        if '= "' in line and not line.strip().startswith('f'):
             line = line.replace('= "', '= f"')
        elif "= '" in line and not line.strip().startswith("f"):
             line = line.replace("= '", "= f'")
             
    lines[i] = line

with open('app.py', 'w', encoding='utf-8') as f:
    f.writelines(lines)
print('Done!')
