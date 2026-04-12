import sys

with open('app.py', 'r', encoding='utf-8') as f:
    text = f.read()

# Fix static fills in Cruce x CUIT
text = text.replace("fill_row = FILL_OK if round(m_a - m_m, 2) == 0.0 else FILL_DIFF", """# Ya no usamos FILL_DIFF estatico, será condicional.""")
text = text.replace("cell.fill = fill_row", "if round(m_a - m_m, 2) == 0.0: cell.fill = FILL_OK")

# Clear the old rule rule_diff_cruce
text = text.replace("""                                FILL_YELLOW = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                                rule_diff_cruce = FormulaRule(formula=['$F6<>"✓ OK"'], stopIfTrue=False, fill=FILL_YELLOW)
                                ws1.conditional_formatting.add(f'A6:F{len(cuits_sorted)+5}', rule_diff_cruce)""", "")

cf_cruce = """                                FILL_YELLOW = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
                                FILL_RED    = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
                                FILL_ORANGE = PatternFill(start_color='F8CBAD', end_color='F8CBAD', fill_type='solid')

                                f_dif = FormulaRule(formula=['$F6="⚠ Diferencia"'], stopIfTrue=False, fill=FILL_YELLOW)
                                f_fal_org = FormulaRule(formula=[f'$F6="⚠ Falta en {organismo}"'], stopIfTrue=False, fill=FILL_RED)
                                f_fal_men = FormulaRule(formula=['$F6="⚠ Falta en Mendez"'], stopIfTrue=False, fill=FILL_ORANGE)

                                ws1.conditional_formatting.add(f'A6:F{len(cuits_sorted)+5}', f_dif)
                                ws1.conditional_formatting.add(f'A6:F{len(cuits_sorted)+5}', f_fal_org)
                                ws1.conditional_formatting.add(f'A6:F{len(cuits_sorted)+5}', f_fal_men)"""
text = text.replace("                                _autofit_ws(ws1, n1)", "                                _autofit_ws(ws1, n1)\n\n" + cf_cruce)


cf_arba = """                                
                                FILL_YELLOW = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
                                FILL_ORANGE = PatternFill(start_color='F8CBAD', end_color='F8CBAD', fill_type='solid')
                                str_col = openpyxl.utils.get_column_letter(idx_match_arba)
                                r_dif = FormulaRule(formula=[f'${str_col}6="⚠ Diferencia"'], stopIfTrue=False, fill=FILL_YELLOW)
                                r_men = FormulaRule(formula=[f'${str_col}6="⚠ Falta en Mendez"'], stopIfTrue=False, fill=FILL_ORANGE)
                                ws2.conditional_formatting.add(f'A6:{openpyxl.utils.get_column_letter(n2)}{len(df_arba_det)+5}', r_dif)
                                ws2.conditional_formatting.add(f'A6:{openpyxl.utils.get_column_letter(n2)}{len(df_arba_det)+5}', r_men)
"""
text = text.replace("""                                rule_diff_arr = FormulaRule(formula=[f'LEFT(${openpyxl.utils.get_column_letter(idx_match_arba)}6, 1)="⚠"'], stopIfTrue=False, fill=PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid'))
                                ws2.conditional_formatting.add(f'A6:{openpyxl.utils.get_column_letter(n2)}{len(df_arba_det)+5}', rule_diff_arr)""", cf_arba)

cf_men = """
                                FILL_YELLOW = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
                                FILL_RED    = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
                                str_col_m = openpyxl.utils.get_column_letter(idx_match_men)
                                r_dif_m = FormulaRule(formula=[f'${str_col_m}6="⚠ Diferencia"'], stopIfTrue=False, fill=FILL_YELLOW)
                                r_org_m = FormulaRule(formula=[f'${str_col_m}6="⚠ Falta en {organismo}"'], stopIfTrue=False, fill=FILL_RED)
                                ws3.conditional_formatting.add(f'A6:{openpyxl.utils.get_column_letter(n3)}{len(df_mendez_det)+5}', r_dif_m)
                                ws3.conditional_formatting.add(f'A6:{openpyxl.utils.get_column_letter(n3)}{len(df_mendez_det)+5}', r_org_m)
"""
text = text.replace("""                                rule_diff_men = FormulaRule(formula=[f'LEFT(${openpyxl.utils.get_column_letter(idx_match_men)}6, 1)="⚠"'], stopIfTrue=False, fill=PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid'))
                                ws3.conditional_formatting.add(f'A6:{openpyxl.utils.get_column_letter(n3)}{len(df_mendez_det)+5}', rule_diff_men)""", cf_men)


an_dif = """                                    for row_i in range(6, len(df_matriz_dif) + 6):
                                        origen = ws_dif.cell(row=row_i, column=idx_orig).value
                                        cuit_val = str(ws_dif.cell(row=row_i, column=1).value).replace("-", "")
                                        
                                        m_a = arba_por_cuit.get(cuit_val, 0)
                                        m_m = mendez_por_cuit.get(cuit_val, 0)
                                        
                                        is_sub = (origen == 'DIFERENCIA CUIT:')
                                        
                                        if m_m == 0: fill = PatternFill('solid', fgColor='F8CBAD')
                                        elif m_a == 0: fill = PatternFill('solid', fgColor='FCE4D6')
                                        else: fill = PatternFill('solid', fgColor='FFF2CC')
                                        
                                        for c in range(1, n_dif + 1):
                                            cell = ws_dif.cell(row=row_i, column=c)
                                            cell.alignment = CTR
                                            cell.border = THIN
                                            if fill: cell.fill = fill
                                            if is_sub: cell.font = Font(bold=True)
                                            if c == idx_m_dif: cell.number_format = FMT_MONEY"""

text = text.replace("""                                    for row_i in range(6, len(df_matriz_dif) + 6):
                                        origen = ws_dif.cell(row=row_i, column=idx_orig).value
                                        is_sub = (origen == 'DIFERENCIA CUIT:')
                                        fill = FILL_SUB_DIF if is_sub else (FILL_ARBA if origen == 'ARBA' else FILL_MEN)
                                        
                                        for c in range(1, n_dif + 1):
                                            cell = ws_dif.cell(row=row_i, column=c)
                                            cell.alignment = CTR
                                            cell.border = THIN
                                            if fill: cell.fill = fill
                                            if is_sub: cell.font = Font(bold=True)
                                            if c == idx_m_dif: cell.number_format = FMT_MONEY""", an_dif)

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(text)
print("Done!")
