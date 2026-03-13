import re
import sys
import io
import pandas as pd

# Eliminar el wrapping global de sys.stdout que causa error en Streamlit
# (Se movió al bloque main())
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CONCEPTOS_MAP = {
    "1": "Mercaderia c/iva", "2": "mercaderia s/iva", "3": "perecederos", "4": "carnes",
    "5": "verduras", "6": "huevos", "7": "pollos", "8": "no perecederos",
    "9": "materia prima c/iva", "10": "materia prima s/iva", "11": "materiales c/iva",
    "12": "materiales s/iva", "13": "productos varios", "14": "alimentos balanceados",
    "15": "bienes de cambio", "16": "Combustible para la venta", "18": "gs de impo/expo",
    "19": "gastos de prestamo", "20": "gastos generales c/iva", "21": "gastos generales s/iva",
    "22": "gastos bancarios c/iva", "23": "gastos bancarios s/iva", "24": "gastos adm. C/iva",
    "25": "gastos adm. S/iva", "26": "gs.comercializacion c/iva", "27": "gs.comercializacion S/iva",
    "28": "servicios varios", "29": "imp. Tasas y contribuciones", "30": "serv x cta de 3° c/iva",
    "31": "serv x cta de 3° s/iva", "32": "gastos despachantes", "33": "honorarios c/ivA",
    "34": "honorarios s/iva", "35": "derechos de importacion", "36": "prestamos",
    "37": "leasing", "38": "intereses", "39": "gastos de tarjeta", "40": "insumos",
    "41": "material de embalaje", "42": "seguros comerciales", "43": "seguro de vida",
    "44": "seguro de vehiculo", "45": "Gastos de vehiculo c/iva", "46": "Gastos de vehiculo S/iva",
    "47": "combustible", "48": "fletes c/iva", "49": "fletes s/iva", "50": "alquiler con iva",
    "51": "alquiler sin iva", "52": "gsts ch/rechazados", "53": "comisiones pagadas",
    "54": "mant y rep bs. De uso", "55": "mant y rep edificio", "56": "alquiler maquinarias",
    "57": "descuentos otorgados", "58": "indumentaria", "59": "anticipo de materiales",
    "60": "hipoteca", "61": "comitentes", "62": "inmobiliario", "63": "descuentos obtenidos",
    "64": "rodados", "65": "instalaciones", "66": "maquinarias", "67": "sistemas informaticos",
    "68": "compra de mue y utiles c/iva", "69": "compra de mue y utiles s/iva",
    "70": "compra de bs uso c/iva", "71": "compra de bs de uso s/iva", "72": "mejoras",
    "73": "vacunos", "74": "equinos", "75": "conejos", "76": "gallinas ponedoras",
    "77": "mejoras inmuebles ajenos", "78": "moldes y matrices", "79": "restitucion de gastos",
    "80": "venta de mercaderia c/iva", "81": "venta de mercaderia s/iva", "82": "venta cons final",
    "83": "venta resumen del dia", "84": "venta bs de uso c/iva", "85": "venta bs de uso s/iva",
    "86": "venta de combustible", "87": "prestacion de servicios", "88": "honorarios c/iva",
    "89": "honorarios s/iva", "90": "alquiler inmuebles", "91": "alquiler de vehiculos",
    "92": "comisiones cobradas", "93": "liquidacion verduleria", "94": "liquidacion carniceria",
    "95": "liquidacion panaderia", "96": "montajes", "97": "venta mayorista", "98": "fabricacion",
    "99": "ch/rechazados", "100": "intereses por prestamos", "101": "recibo anulado",
    "102": "venta de maquinarias", "103": "liquidacion perfumeria", "104": "venta ganado x cta de 3°",
    "105": "pastaje", "106": "venta de exportacion", "107": "liquidacion agropecuaria",
    "109": "toros", "110": "licencia", "111": "donaciones", "112": "diferencia de cambio",
    "113": "gastos financieros", "114": "arrendamientos", "115": "dto. De valores",
    "116": "mano de obra de 3°", "117": "negociacion de valores", "118": "liquidacion dto de cheques",
    "119": "patentes de vehiculos", "120": "transporte", "121": "alquiler particular c/iva",
    "122": "alquiler particular s/iva", "123": "alquiler comercial c/iva",
    "124": "alquiler comercial s/iva", "125": "utiles y herramientas", "126": "premios",
    "127": "reconocimientos", "128": "publicidad y propaganda", "129": "gastos de seguridad",
    "130": "servivio de transporte", "131": "comprobante anulado", "132": "anticipos",
    "133": "gastos carrera", "134": "Mejora inmueble Propio", "135": "Gastos de medicina",
    "136": "Tasa de Fondeadero", "137": "Seguros Leasing", "138": "Repuestos e Insumos",
    "139": "Ofrendas y Limosnas", "140": "Gastos de Comedor", "141": "Ganado propio c/iva",
    "142": "Compra de Ganado", "143": "Vta. Carne Vacuna", "144": "Envases y Accesorios",
    "145": "Venta de Ganado", "146": "Ajuste Contable", "147": "Fondo de Comercio",
    "148": "Servicios Personales", "149": "COMPRA DE CARNE", "150": "Insumos Papas",
    "151": "Gastos de Arrendamiento", "152": "Venta de Vehiculo", "153": "Compra de Vehiculo",
    "154": "Obras en Curso", "155": "Boletos y Pasajes", "156": "Alquiler Barco",
    "157": "Materiales de Decoracion", "158": "Alquiler y Expensas", "159": "Alquiler de Herramientas",
    "160": "Viandas", "161": "Intereses", "162": "Seguros de Caucion", "163": "Impresiones",
    "164": "Gastos de Producción", "165": "Prestadores", "166": "Devolucion de Mercaderias",
    "167": "Alquiler de Maquinarias", "168": "Certificados Revisión Técnica",
    "169": "Registro Control Modelo", "170": "Camara Arg. De Talleres", "171": "Honorarios Directores",
    "172": "Fondo de Reparo", "173": "Gastos de Sanidad", "174": "Plan de ahorro",
    "175": "Alquiler Temporario", "176": "Alquiler y Logistica", "177": "Alquiler Bs. Muebles",
    "178": "Gastos de Capacitación", "179": "Maq. Y equipos medicos", "180": "gastos de organización",
    "181": "equipos de comunicación", "182": "Gas para la venta", "183": "Venta Flete Internacional",
    "184": "Flete Internacional", "185": "Gastos de Obra", "186": "Gastos de Desarrollo",
    "187": "Embarcaciones", "188": "Gastos de embarcacion", "189": "Venta de Papa",
    "190": "Utiles y elementos de cocina", "191": "cubiertos y vajillas", "192": "elementos ortopedicos",
    "193": "pines", "194": "golosinas", "195": "rotary internacional", "196": "distrito rotario 4825",
    "197": "ret- seguridad e higiene a", "198": "gastos de representacion",
    "199": "ativo de caja  (compra + v", "201": "C.M.", "202": "ativo de caja  (compra + v",
    "203": "Alimentos", "204": "Enfriado",
}


# ──────────────────────────────────────────────────────────────
# Regex
# ──────────────────────────────────────────────────────────────

# Regex para la línea principal de una transacción
# Ejemplo: " 1 FC 05009-07466844A AUTOPISTAS URBANAS S A Ins. 30-57487647-4  45 B Exento           743,65          0,00          0,00        743,65"
# El día puede ser 1 o 2 dígitos, el tipo 2-3 chars, el número de comprobante variable
RE_MAIN = re.compile(
    r'^\s*(\d{1,2})\s+'                            # Dia
    r'(FC|NC|ND|TF|TK)\s+'                          # Tipo comprobante
    r'(\d{5}-\d{1,12}[A-Z ]?)\s*'                   # Numero (más flexible para exportación)
    r'(.+?)\s+'                                     # Proveedor (Flexible hasta Cond IVA)
    r'(Ins\.|Mono|Monot|Exe |Exe\.|C\.F\.|Exp\.|Resp\.)\s+' # Cond IVA
    r'([\d-]{7,13})?\s+'                            # CUIT/DNI (Opcional)
    r'(\d{1,3})\s+'                                 # Concepto
    r'([A-Z])\s+'                                   # Jurisdicción (Letra A-Z)
    r'(.+)$'                                        # Resto (tasa + montos)
)

# Regex para líneas de continuación (sub-conceptos)
# Ejemplo: "                                                                       Imp.Inter        385,94          0,00          0,00       5802,89"
RE_CONT = re.compile(
    r'^\s{50,}'                                # Gran cantidad de espacios
    r'(\S.+)$'                                 # contenido
)

# Regex para extraer montos (formato argentino: 1.234,56 o -1.234,56)
RE_MONTO = re.compile(r'-?[\d]+(?:\.[\d]{3})*,\d{2}')

# Líneas a ignorar
RE_IGNORE = re.compile(
    r'^\s*$|'
    r'^\s*Pag\.:|'
    r'^\s*CLASIFICADORURAL|'
    r'^\s*ESTADOS UNIDOS|'
    r'^\s+Numero de CUIT|'
    r'^\s*[A-Z ]?\s*IVA COMPRAS|'
    r'^\s*[A-Z ]?\s*Desde el|'
    r'^\s*Dia\s+Numero|'
    r'^\s*TC\s+|'
    r'^-- --|'
    r'^-{10,}|'
    r'==>|'                                    # Cualquier línea con flecha (subtotales)
    r'TOTALES\s+POR|'                          # Encabezados de tablas de resumen
    r'^\s*TOTAL\s+GENERAL|'
    r'^Cod\s+Concepto|'
    r'^Cod\s+Detalle|'
    r'^\s*\d+\s+Factura|'
    r'^\s*\d+\s+Nota de|'
    r'^\s*\d+\s+Tiquet|'
    r'^\s*[A-Z]\s+(Exento|Resp\.|Resp\.)|'
    r'^I: Valor neto|'
    r'^\x0c|'          # Form feed
    r'^\x0f|'          # Control chars
    r'^\x1b',          # ESC sequences
    re.IGNORECASE
)


def limpiar_control(texto: str) -> str:
    """Elimina caracteres de control y escape del texto."""
    texto = re.sub(r'\x1b[A-Za-z@]', '', texto)   # ESC + letra
    texto = re.sub(r'\x1b[A-Z]', '', texto)         
    texto = re.sub(r'[\x00-\x09\x0b-\x0c\x0e-\x1f]', '', texto)
    return texto.rstrip('\r\n')


def parse_monto(s: str) -> float:
    """Convierte un string '1.234,56' o '-1.234,56' al float correspondiente."""
    s = s.strip()
    s = s.replace('.', '').replace(',', '.')
    return float(s)


def extraer_montos_resto(resto: str):
    """
    Del 'resto' de la línea principal, extrae la Tasa y los 4 montos.
    Retorna: (tasa_str, neto, iva, percepcion, total)
    """
    montos = RE_MONTO.findall(resto)
    
    # Determinar la tasa
    tasa_str = resto.split(montos[0])[0].strip() if montos else resto.strip()
    
    neto = parse_monto(montos[0]) if len(montos) >= 1 else 0.0
    iva = parse_monto(montos[1]) if len(montos) >= 2 else 0.0
    percepcion = parse_monto(montos[2]) if len(montos) >= 3 else 0.0
    total = parse_monto(montos[3]) if len(montos) >= 4 else 0.0
    
    return tasa_str, neto, iva, percepcion, total


def extraer_montos_continuacion(contenido: str):
    """
    De una línea de continuación extrae concepto y montos.
    Retorna: (concepto, neto, iva, percepcion, total_parcial)
    """
    montos = RE_MONTO.findall(contenido)
    concepto = contenido.split(montos[0])[0].strip() if montos else contenido.strip()
    
    neto = parse_monto(montos[0]) if len(montos) >= 1 else 0.0
    iva = parse_monto(montos[1]) if len(montos) >= 2 else 0.0
    percepcion = parse_monto(montos[2]) if len(montos) >= 3 else 0.0
    total = parse_monto(montos[3]) if len(montos) >= 4 else 0.0
    
    return concepto, neto, iva, percepcion, total


def limpiar_para_excel(texto: str) -> str:
    """Elimina caracteres de control no permitidos en Excel/XML."""
    if not texto: return ""
    # Quitar caracteres de control ASCII 0-31 (excepto newline si fuera necesario, pero aqui no)
    # y otros caracteres no imprimibles detectados
    return re.sub(r'[\x00-\x1f\x7f-\x9f]', '', texto).strip()


def parsear_archivo(path: Path = None, content: str = None) -> tuple[list[dict], dict]:
    """Lee el archivo .txt (desde path o contenido directo) y extrae transacciones y metadata."""
    if content is None and path is not None:
        with open(path, 'r', encoding='ansi') as f:
            content = f.read()
    
    if not content:
        return [], {}

    lines = content.splitlines()
    transacciones = []
    current = None
    
    # Metadata del contribuyente
    meta = {
        'razon_social': '',
        'cuit_empresa': '',
        'periodo': ''
    }
    
    # Extraer metadata de las primeras líneas con limpieza
    if len(lines) > 5:
        meta['razon_social'] = limpiar_para_excel(lines[1])
        cuit_match = re.search(r'CUIT:([\d-]+)', lines[3])
        if cuit_match:
            meta['cuit_empresa'] = cuit_match.group(1)
        
        # El tipo de reporte (IVA COMPRAS / IVA VENTAS) suele estar en la linea 5
        reporte_raw = limpiar_para_excel(lines[4])
        # Limpiar prefijos como "E " o "F " que a veces aparecen en el TXT
        meta['tipo_reporte'] = re.sub(r'^[A-Z]\s+', '', reporte_raw).strip()

        # El periodo suele estar en la linea 6. Intentamos captar solo el texto "Desde... hasta..."
        periodo_raw = lines[5]
        p_match = re.search(r'(Desde .* hasta .*)', periodo_raw)
        if p_match:
            meta['periodo'] = limpiar_para_excel(p_match.group(1))
        else:
            meta['periodo'] = limpiar_para_excel(periodo_raw)

    for line in lines:
        linea = limpiar_control(line)
        
        # Ignorar líneas de encabezado, separadores, subtotales, etc.
        if RE_IGNORE.search(linea):
            # NO cerramos la transacción en separadores/subtotales
            # porque puede continuar en la página siguiente
            continue
        
        # Intentar match de línea principal
        m = RE_MAIN.match(linea)
        if m:
            dia = int(m.group(1))
            tipo = m.group(2).strip()
            numero = m.group(3).strip()
            proveedor = m.group(4).strip()
            cond_iva = m.group(5).strip()
            cuit = m.group(6).strip() if m.group(6) else ""
            concepto = int(m.group(7))
            letra = m.group(8).strip()
            resto = m.group(9)
            
            tasa_str, neto, iva, percepcion, total = extraer_montos_resto(resto)
            
            # Si es el MISMO comprobante (continuación tras salto de página),
            # tratar como sub-concepto en vez de nueva transacción.
            # Agregamos CUIT y Proveedor para evitar agrupar movimientos distintos con mismo número (ej: SIRCREB)
            if (current and
                current['Fecha'] == dia and
                current['Tipo'] == tipo and
                current['Numero'] == numero and
                current['CUIT'] == cuit and
                current['Proveedor'] == proveedor):
                # Es el mismo comprobante (salto de página):
                # Agregamos los montos como sub-conceptos para que se distribuyan 
                # correctamente según su propia 'tasa_str'.
                # NO sumamos a current['Neto']/IVA/Percepcion directamente 
                # porque eso forzaría los valores al bucket de la primera página.
                if total != 0.0:
                    current['Total'] = total
                current['SubConceptos'].append({
                    'Concepto': tasa_str,
                    'Neto': neto,
                    'IVA': iva,
                    'Percepcion': percepcion,
                    'Total': total
                })
                continue
            
            # Es una transacción nueva → guardar la previa
            if current:
                transacciones.append(current)
            
            current = {
                'Fecha': dia,
                'Tipo': tipo,
                'Numero': numero,
                'Proveedor': proveedor,
                'Cond_IVA': cond_iva,
                'CUIT': cuit,
                'Concepto': concepto,
                'Letra': letra,
                'Tasa': tasa_str,
                'Neto': neto,
                'IVA': iva,
                'Percepcion': percepcion,
                'Total': total,
                'SubConceptos': []
            }
            continue
        
        # Intentar match de línea de continuación
        mc = RE_CONT.match(linea)
        if mc and current:
            contenido = mc.group(1)
            concepto_sub, neto_s, iva_s, perc_s, total_s = extraer_montos_continuacion(contenido)
            
            # El total de la última línea de continuación tiene el total correcto
            if total_s != 0.0:
                current['Total'] = total_s
            
            current['SubConceptos'].append({
                'Concepto': concepto_sub,
                'Neto': neto_s,
                'IVA': iva_s,
                'Percepcion': perc_s,
                'Total': total_s
            })
            continue
    
    # Guardar última transacción si existe
    if current:
        transacciones.append(current)
    
    return transacciones, meta


def _autofit(ws, n_cols, start_row=6):
    """Ajusta el ancho de todas las columnas de una hoja al contenido.
    Empieza desde start_row para ignorar filas de titulo mergeadas."""
    for col_idx in range(1, n_cols + 1):
        letter = get_column_letter(col_idx)
        max_len = 0
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx,
                                min_row=start_row, max_row=ws.max_row):
            for cell in row:
                val = cell.value
                if val is not None:
                    if isinstance(val, str) and val.startswith('='):
                        text = '($999,999.99)'  # ancho estimado para formulas
                    elif isinstance(val, (int, float)):
                        text = f'${val:,.2f}'
                    else:
                        text = str(val)
                    if len(text) > max_len:
                        max_len = len(text)
        ws.column_dimensions[letter].width = max(max_len + 3, 8)


def crear_excel(transacciones: list[dict], meta: dict, output_path, con_resumenes=True, con_auxiliar=False, cruce_arca=False, df_arca=None):
    """Crea un Excel formateado. Cada tasa de IVA tiene sus propias columnas
    Neto/IVA y cada percepcion/retencion tiene su propia columna.
    output_path puede ser una ruta o un BytesIO buffer."""

    # ── Mapeo de tasas IVA a columnas ─────────────────────────
    IVA_RATES = {
        'Tasa 21%':  ('Neto IVA 21',   'IVA 21'),
        'T.21%':     ('Neto IVA 21',   'IVA 21'),
        'C.F.21%':   ('Neto C.F. 21',  'IVA C.F. 21'),
        'Tasa 27%':  ('Neto IVA 27',   'IVA 27'),
        'T.27%':     ('Neto IVA 27',   'IVA 27'),
        'Tasa 10.5%': ('Neto IVA 10.5', 'IVA 10.5'),
        'Tasa 10,5%': ('Neto IVA 10.5', 'IVA 10.5'),
        'T.10.5%':   ('Neto IVA 10.5', 'IVA 10.5'),
        'T.10,5%':   ('Neto IVA 10.5', 'IVA 10.5'),
        'C.F.10.5%': ('Neto C.F. 10.5', 'IVA C.F. 10.5'),
        'C.F.10,5%': ('Neto C.F. 10.5', 'IVA C.F. 10.5'),
        'Tasa 5%':   ('Neto IVA 5',    'IVA 5'),
        'T.5%':      ('Neto IVA 5',    'IVA 5'),
        'Tasa 2.5%': ('Neto IVA 2.5',  'IVA 2.5'),
        'Tasa 2,5%': ('Neto IVA 2.5',  'IVA 2.5'),
        'T.2.5%':    ('Neto IVA 2.5',  'IVA 2.5'),
        'T.2,5%':    ('Neto IVA 2.5',  'IVA 2.5'),
        'T.IMP 21%': ('Neto Imp. 21',  'IVA Imp. 21'),
        'T.IMP 10%': ('Neto Imp. 10.5','IVA Imp. 10.5'),
        'Exento':    ('Exento',    None),
    }

    # En ventas, expandir monotributo a columnas Neto/IVA; en compras, dejar tasa cruda
    es_ventas = 'VENTA' in meta.get('tipo_reporte', '').upper()
    if es_ventas:
        IVA_RATES['R.Monot21'] = ('Neto Monot. 21', 'IVA Monot. 21')
        IVA_RATES['R.Mont.10'] = ('Neto Monot. 10.5', 'IVA Monot. 10.5')
    
    DESIRED_IVA_ORDER = [
        'Neto IVA 21', 'IVA 21',
        'Neto C.F. 21', 'IVA C.F. 21',
        'Neto IVA 27', 'IVA 27',
        'Neto IVA 10.5', 'IVA 10.5',
        'Neto C.F. 10.5', 'IVA C.F. 10.5',
        'Neto IVA 5', 'IVA 5',
        'Neto IVA 2.5', 'IVA 2.5',
        'Neto Monot. 21', 'IVA Monot. 21',
        'Neto Monot. 10.5', 'IVA Monot. 10.5',
        'Neto Imp. 21', 'IVA Imp. 21',
        'Neto Imp. 10.5', 'IVA Imp. 10.5',
        'Exento', 'Monotributo',
    ]

    # ── 1. Recopilar sub-conceptos y tasas presentes ────────
    if not transacciones:
        return

    present_iva_cols = set()
    found_others = []  # Lista ordenada (preserva orden de aparición en TXT)
    
    for t in transacciones:
        # Tasa principal
        tasa = t['Tasa']
        if tasa in IVA_RATES:
            neto_col, iva_col = IVA_RATES[tasa]
            present_iva_cols.add(neto_col)
            if iva_col: present_iva_cols.add(iva_col)
        elif tasa and tasa.strip():
            t_clean = tasa.strip()
            if t_clean not in found_others:
                found_others.append(t_clean)
            
        # Sub-conceptos
        for s in t['SubConceptos']:
            conc = s['Concepto']
            if conc in IVA_RATES:
                neto_col, iva_col = IVA_RATES[conc]
                present_iva_cols.add(neto_col)
                if iva_col: present_iva_cols.add(iva_col)
            elif conc and conc.strip():
                c_clean = conc.strip()
                if c_clean not in found_others:
                    found_others.append(c_clean)

    IVA_COL_ORDER = [c for c in DESIRED_IVA_ORDER if c in present_iva_cols]
    if not IVA_COL_ORDER:
        # Si no hay IVA (ej. todo exento o solo percepciones), usar una columna genérica o solo others
        if not found_others and not present_iva_cols:
             return # No hay nada que escribir
        IVA_COL_ORDER = sorted(list(present_iva_cols)) # fallback
    
    other_cols = list(found_others)  # Preservar orden de aparición del TXT

    # Helper: detectar si un nombre de columna es una deducción (PERC/RET/SIRCREB)
    _DEDUCCION_KW = ("PERC", "PER.", "PER ", "RET", "SIRCREB", "SIRTAC")
    def _es_deduccion(nombre: str) -> bool:
        nu = nombre.upper()
        return any(kw in nu for kw in _DEDUCCION_KW)

    # Ordenar: primero no-deducciones (amarillo), luego deducciones (verde)
    other_cols = [c for c in other_cols if not _es_deduccion(c)] + \
                 [c for c in other_cols if _es_deduccion(c)]

    rows = []
    for t in transacciones:
        # Separar Numero (ej: 00002-00000018A) en PV, Nro., Letra
        numero_raw = t['Numero']
        partes = numero_raw.replace('-', '')
        # Punto de venta = primeros 5 digitos, Nro = digitos restantes, Letra = ultimo char
        pv = numero_raw.split('-')[0] if '-' in numero_raw else numero_raw[:5]
        resto_num = numero_raw.split('-')[1] if '-' in numero_raw else numero_raw[5:]
        letra = resto_num[-1] if resto_num and resto_num[-1].isalpha() else ''
        nro = resto_num[:-1] if letra else resto_num

        # Limpiar CUIT: quitar guiones y convertir a numerico
        cuit_raw = t['CUIT'].replace('-', '') if t['CUIT'] else ''
        cuit_val = int(cuit_raw) if cuit_raw and cuit_raw.isdigit() else cuit_raw

        row = {
            'Fecha': t['Fecha'],
            'Tipo': t['Tipo'],
            'PV': int(pv),
            'Nro.': int(nro) if nro.isdigit() else nro,
            'Letra': letra,
            'Proveedor': t['Proveedor'],
            'Cond. IVA': t['Cond_IVA'],
            'CUIT': cuit_val,
            'Concepto': t['Concepto'],
            'Jur.': t['Letra'],
        }
        # Inicializar columnas IVA
        for col in IVA_COL_ORDER:
            row[col] = 0.0
        # Inicializar columnas de otros conceptos
        for col in other_cols:
            row[col] = 0.0

        # Colocar montos de la linea principal
        tasa = t['Tasa']
        if tasa in IVA_RATES:
            neto_col, iva_col = IVA_RATES[tasa]
            row[neto_col] += t['Neto']
            if iva_col:
                row[iva_col] += t['IVA']
        elif tasa:
            # La tasa es una percepcion/retencion
            row[tasa] += t['Neto']

        # Colocar montos de sub-conceptos
        for s in t['SubConceptos']:
            nombre = s['Concepto']
            if not nombre:
                continue
            if nombre in IVA_RATES:
                neto_col, iva_col = IVA_RATES[nombre]
                row[neto_col] += s['Neto']
                if iva_col:
                    row[iva_col] += s['IVA']
            else:
                monto = s['Neto'] if s['Neto'] != 0.0 else s['Percepcion']
                row[nombre] += monto

        # Columna Auxiliar: placeholder (formula se agrega después de escribir el Excel)
        if con_auxiliar or cruce_arca:
            row['Auxiliar'] = ''

        row['Total'] = t['Total']

        # Notas de credito (NC): invertimos el signo de todos los montos
        if t['Tipo'] == 'NC':
            for col in IVA_COL_ORDER + other_cols:
                row[col] = -row[col]
            row['Total'] = -row['Total']

        rows.append(row)

    df = pd.DataFrame(rows)

    all_dynamic = IVA_COL_ORDER + other_cols

    print(f"  Total de transacciones parseadas: {len(df)}")
    print(f"   - FC (Factura): {len(df[df['Tipo'] == 'FC'])}")
    print(f"   - NC (Nota Credito): {len(df[df['Tipo'] == 'NC'])}")
    print(f"   - ND (Nota Debito): {len(df[df['Tipo'] == 'ND'])}")
    print(f"   - TF (Ticket Factura): {len(df[df['Tipo'] == 'TF'])}")

    # ── 3. Escribir Excel ─────────────────────────────────────
    total_cols = len(df.columns)
    last_col_letter = get_column_letter(total_cols)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # ── Estilos Reutilizables ─────────────────────────────
        title_font = Font(bold=True, size=14, color='FFFFFF')
        title_fill = PatternFill('solid', fgColor='2F5496')
        
        header_font = Font(bold=True, size=10, color='FFFFFF')
        header_fill = PatternFill('solid', fgColor='4472C4')
        iva_header_fill = PatternFill('solid', fgColor='BF8F00')
        perc_header_fill = PatternFill('solid', fgColor='70AD47')
        
        header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        center_align = Alignment(horizontal='center', vertical='center')
        
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        zebra_fill = PatternFill('solid', fgColor='D6E4F0')
        money_fmt = '$#,##0.00'
        # Formato contabilidad Peso para cruce ARCA
        accounting_fmt = '_-"$"* #,##0.00_-;-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'

        # ── Hoja Movimientos / SISTEMA ──────────────────────────────────
        # startrow=5 significa que los encabezados del DataFrame van en la fila 6
        mov_sheet_name = 'SISTEMA' if cruce_arca else 'Movimientos'
        df.to_excel(writer, sheet_name=mov_sheet_name, index=False, startrow=5)
        ws = writer.sheets[mov_sheet_name]

        # Estilo de Reporte (Rojo/Diferente para resaltar)
        report_type_font = Font(bold=True, size=12, color='C00000') # Rojo oscuro

        ws.merge_cells(f'A1:{last_col_letter}1')
        ws['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
        ws['A1'].font = title_font
        ws['A1'].fill = title_fill
        ws['A1'].alignment = center_align

        ws.merge_cells(f'A2:{last_col_letter}2')
        ws['A2'] = meta['tipo_reporte'].upper() if meta['tipo_reporte'] else 'REPORTE DE MOVIMIENTOS'
        ws['A2'].font = report_type_font
        ws['A2'].alignment = center_align

        ws.merge_cells(f'A3:{last_col_letter}3')
        ws['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
        ws['A3'].font = Font(bold=True, size=11, color='2F5496')
        ws['A3'].alignment = center_align

        ws.merge_cells(f'A4:{last_col_letter}4')
        ws['A4'] = f'Total: {len(df)} transacciones'
        ws['A4'].font = Font(italic=True, size=10, color='4472C4')
        ws['A4'].alignment = center_align

        # Fila 5 queda vacía como separador

        col_list = list(df.columns)
        iva_set = set(IVA_COL_ORDER)
        other_set = set(other_cols)
        # Deducciones (PERC/PER./RET/SIRCREB) → verde; otros impuestos → amarillo
        deduccion_set = {c for c in other_cols if _es_deduccion(c)}

        # Los encabezados ahora van en la fila 6
        header_row = 6
        data_start_row = 7
        for col_idx in range(1, total_cols + 1):
            cell = ws.cell(row=header_row, column=col_idx)
            cell.font = header_font
            cell.alignment = header_align
            cell.border = thin_border
            col_name = col_list[col_idx - 1]
            if col_name in iva_set or (col_name in other_set and col_name not in deduccion_set):
                cell.fill = iva_header_fill
            elif col_name in deduccion_set:
                cell.fill = perc_header_fill
            else:
                cell.fill = header_fill

        money_col_names = IVA_COL_ORDER + other_cols + ['Total']
        money_col_indices = [col_list.index(c) + 1 for c in money_col_names if c in col_list]
        active_money_fmt = accounting_fmt if cruce_arca else money_fmt
        cuit_col_idx = (col_list.index('CUIT') + 1) if 'CUIT' in col_list else None

        first_sum_col = get_column_letter(col_list.index(IVA_COL_ORDER[0]) + 1)
        last_sum_col = get_column_letter(col_list.index(other_cols[-1]) + 1) if other_cols else get_column_letter(col_list.index(IVA_COL_ORDER[-1]) + 1)
        total_col_idx = col_list.index('Total') + 1

        for row in range(data_start_row, len(df) + data_start_row):
            # Alinear todas las celdas al centro
            for col_idx in range(1, total_cols + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.alignment = center_align
                if col_idx in money_col_indices:
                    cell.number_format = active_money_fmt

            # Formula SUM
            ws.cell(row=row, column=total_col_idx).value = f'=SUM({first_sum_col}{row}:{last_sum_col}{row})'
            ws.cell(row=row, column=total_col_idx).number_format = active_money_fmt

            # Estilo Zebra (reusando el objeto fill)
            if (row - data_start_row) % 2 == 0:
                for col_idx in range(1, total_cols + 1):
                    ws.cell(row=row, column=col_idx).fill = zebra_fill

        # Fila TOTAL GENERAL en Movimientos
        total_row_mov = len(df) + data_start_row
        first_money_idx = col_list.index(IVA_COL_ORDER[0]) + 1
        
        ws.merge_cells(f'A{total_row_mov}:{get_column_letter(first_money_idx-1)}{total_row_mov}')
        ws[f'A{total_row_mov}'] = "TOTAL GENERAL"
        ws[f'A{total_row_mov}'].font = Font(bold=True)
        ws[f'A{total_row_mov}'].alignment = Alignment(horizontal='right')
        
        for col_idx in range(first_money_idx, total_cols + 1):
            col_l = get_column_letter(col_idx)
            cell = ws.cell(row=total_row_mov, column=col_idx)
            cell.value = f'=SUM({col_l}{data_start_row}:{col_l}{total_row_mov-1})'
            cell.font = Font(bold=True)
            cell.border = Border(top=Side(style='double'))
            cell.number_format = active_money_fmt
            cell.alignment = center_align

        # ── Formulas Auxiliar (interactivas) ───────────────────
        if (con_auxiliar or cruce_arca) and 'Auxiliar' in col_list:
            aux_col_idx = col_list.index('Auxiliar') + 1
            aux_col_letter = get_column_letter(aux_col_idx)
            tipo_letter = get_column_letter(col_list.index('Tipo') + 1)
            letra_letter = get_column_letter(col_list.index('Letra') + 1)
            pv_letter = get_column_letter(col_list.index('PV') + 1)
            nro_letter = get_column_letter(col_list.index('Nro.') + 1)
            cuit_letter = get_column_letter(col_list.index('CUIT') + 1)
            total_col_letter = get_column_letter(col_list.index('Total') + 1)
            for row in range(data_start_row, len(df) + data_start_row):
                ws.cell(row=row, column=aux_col_idx).value = (
                    f'={tipo_letter}{row}&" "&{letra_letter}{row}&{pv_letter}{row}&{nro_letter}{row}&{cuit_letter}{row}'
                )

        # ── CRUCE + DIFF formulas en SISTEMA ──────────────────────
        if cruce_arca and 'Auxiliar' in col_list and df_arca is not None and not df_arca.empty:
            # Agregar columnas CRUCE y DIFF al final
            cruce_col_idx = total_cols + 1
            diff_col_idx = total_cols + 2
            cruce_col_letter = get_column_letter(cruce_col_idx)
            diff_col_letter = get_column_letter(diff_col_idx)

            # Headers
            ws.cell(row=6, column=cruce_col_idx).value = 'CRUCE'
            ws.cell(row=6, column=cruce_col_idx).font = header_font
            ws.cell(row=6, column=cruce_col_idx).fill = PatternFill('solid', fgColor='7030A0')
            ws.cell(row=6, column=cruce_col_idx).alignment = header_align
            ws.cell(row=6, column=cruce_col_idx).border = thin_border

            ws.cell(row=6, column=diff_col_idx).value = 'DIFF'
            ws.cell(row=6, column=diff_col_idx).font = header_font
            ws.cell(row=6, column=diff_col_idx).fill = PatternFill('solid', fgColor='7030A0')
            ws.cell(row=6, column=diff_col_idx).alignment = header_align
            ws.cell(row=6, column=diff_col_idx).border = thin_border

            # Calcular rango de lookup en ARCA (Auxiliar:Total son las 2 ultimas cols)
            arca_col_list = list(df_arca.columns)
            arca_aux_col_letter = get_column_letter(arca_col_list.index('Auxiliar') + 1) if 'Auxiliar' in arca_col_list else 'A'
            arca_total_col_letter = get_column_letter(arca_col_list.index('Total') + 1) if 'Total' in arca_col_list else 'B'
            arca_total_col_offset = arca_col_list.index('Total') - arca_col_list.index('Auxiliar') + 1 if 'Auxiliar' in arca_col_list and 'Total' in arca_col_list else 2
            arca_last_data_row = len(df_arca) + 6  # data starts at row 7

            for row in range(data_start_row, len(df) + data_start_row):
                # CRUCE: VLOOKUP en ARCA buscando Auxiliar, trayendo Total
                ws.cell(row=row, column=cruce_col_idx).value = (
                    f'=IFERROR(VLOOKUP({aux_col_letter}{row},'
                    f"ARCA!${arca_aux_col_letter}$7:${arca_total_col_letter}${arca_last_data_row},"
                    f'{arca_total_col_offset},FALSE),"NO ENCONTRADO")'
                )
                ws.cell(row=row, column=cruce_col_idx).number_format = money_fmt
                ws.cell(row=row, column=cruce_col_idx).alignment = center_align

                # DIFF: Total - CRUCE (solo si CRUCE es numerico)
                ws.cell(row=row, column=diff_col_idx).value = (
                    f'=IF({cruce_col_letter}{row}="NO ENCONTRADO","",'
                    f'{total_col_letter}{row}-{cruce_col_letter}{row})'
                )
                ws.cell(row=row, column=diff_col_idx).number_format = money_fmt
                ws.cell(row=row, column=diff_col_idx).alignment = center_align

        _autofit(ws, total_cols)
        ws.column_dimensions['A'].width = 8 # Ancho fijo para columna Fecha

        # ── Hojas de Resumen (Solo si se solicita) ────────────
        if con_resumenes:

            # Resto del código de resúmenes...
            resumen = df.copy()
        
            # Separar conceptos en Deducciones (PERC/RET) y Otros (IMP.CIG, etc.)
            deduccion_cols = [c for c in other_cols if _es_deduccion(c)]
            individual_other_cols = [c for c in other_cols if c not in deduccion_cols]
        
            res_header_row = 6
            res_data_start = 7
            n_mov = len(df)
            mov_cuit_col = get_column_letter(col_list.index('CUIT') + 1)
            mov_tipo_col = get_column_letter(col_list.index('Tipo') + 1)
            mov_conc_col = get_column_letter(col_list.index('Concepto') + 1)
            mov_jur_col = get_column_letter(col_list.index('Jur.') + 1)

            # ── Hoja Resumen por Impuesto (INTERACTIVA) ──────────
            res_imp_data = []
            seen_cols = set()
        
            # Tasas estándar de IVA
            for tasa_label, (neto_col, iva_col) in IVA_RATES.items():
                if neto_col in df.columns and (neto_col, iva_col) not in seen_cols:
                    n_idx = col_list.index(neto_col) + 1
                    i_idx = (col_list.index(iva_col) + 1) if (iva_col and iva_col in df.columns) else None
                    res_imp_data.append({
                        'Tasa': tasa_label,
                        'Neto_Col_M': get_column_letter(n_idx),
                        'IVA_Col_M': get_column_letter(i_idx) if i_idx else None,
                        'Ded_Col_M': None
                    })
                    seen_cols.add((neto_col, iva_col))
        
            # Deducciones y otros
            for col in other_cols:
                c_idx = col_list.index(col) + 1
                col_upper = col.upper()
                if "PERC" in col_upper or "RET" in col_upper or "SIRCREB" in col_upper:
                    res_imp_data.append({
                        'Tasa': col, 'Neto_Col_M': None, 'IVA_Col_M': None, 'Ded_Col_M': get_column_letter(c_idx)
                    })
                else:
                    # Todo lo demás que no es IVA (como IMP.CIG, impuestos internos, ajustes, etc.) va a Neto
                    res_imp_data.append({
                        'Tasa': col, 'Neto_Col_M': get_column_letter(c_idx), 'IVA_Col_M': None, 'Ded_Col_M': None
                    })

            # ── Orden de conceptos: impuestos primero (por código), luego deducciones ──
            TASA_ORDER_MAP = {
                'Exento': 1,
                'Tasa 21%': 2, 'T.21%': 2,
                'Tasa 27%': 3, 'T.27%': 3,
                'T.10.5%': 4, 'T.10,5%': 4, 'Tasa 10.5%': 4, 'Tasa 10,5%': 4,
                'Tasa 21+5': 5,
                'Tasa 27+5': 6,
                'Imp.Inter': 7, 'Imp.Inter.': 7,
                'Cons.Fin.': 8,
                'R.Monot21': 9,
                'R.Mont.10': 10,
                'C.F.21%': 11,
                'C.F.10.5%': 12, 'C.F.10,5%': 12,
                'CPTE.ANUL': 13,
                'IMP.COMB.': 14,
                'IMP.CIG.': 15,
                'ABASTO.': 16,
                'L.25413': 17,
                'IMP.SELLO': 18,
                'D.976/01': 19,
                'BONIFIC.': 20,
                'itida en': 21,
                'TJT Prep.': 22,
                'L25413(2)': 23,
                'AJUST RED': 24,
                'IVA PEAJE': 25,
                'TRANPORT': 26,
                'AJUST IVA': 27,
                'DESC OTOR': 28,
                'DESC.S/IV': 29,
                'DESC.10.5': 30,
                'Valor Cri': 31,
                'Incre Iva': 32,
                'S41': 33,
                'DESC CF': 34,
                'DESC.MONO': 35,
                'Tasa 16,5': 36,
                'Tasa 22%': 37,
                'T.IMP 21%': 38,
                'T.IMP 10%': 39,
                'TASA 16,6': 40,
                'Tasa 2.5%': 41, 'Tasa 2,5%': 41, 'T.2.5%': 41, 'T.2,5%': 41,
                'Tasa 0%': 42,
                'TASA 9%': 43,
                'TASA 5%': 44, 'Tasa 5%': 44, 'T.5%': 44,
                'L.27264': 45,
                'IMP.FONDE': 46,
                'R.Mont.27': 47,
                'TurIVA': 48,
                'REC. GAS.': 49,
                'CRI 10,5': 50,
                'CCF 10,5': 51,
                'CRM 10,50': 52,
                'AJ.IMPORT': 53,
                'CM 21%': 54,
                'CM CF 21%': 55,
                'IMP. PAIS': 56,
                'GST IIBB': 57,
                'T21% 4240': 58,
            }

            # ── Orden de deducciones (percepciones/retenciones) ──
            DEDUCCION_ORDER_MAP = {
                'PERC.I.V.A.': 1, 'PERC.IVA': 1, 'PERC IVA': 1,
                'PERC.GCIAS.': 2, 'PERC.GCIAS': 2, 'PERC GCIAS': 2,
                'PERC.IB.CAP.FED.': 3, 'PERC.IB.CAP.FED': 3,
                'PERC.IB.BS.AS.': 4, 'PERC.IB.BS.AS': 4,
                'PERC.IB.CORDOBA': 5, 'PERC.IB.CÓRDOBA': 5,
                'PERC.IB.MENDOZA': 6,
                'PERC.IB.MISIONES': 7,
                'RET.GCIAS': 8, 'RET GCIAS': 8,
                'RET.IB.BS.AS.': 9, 'RET.IB.BS.AS': 9,
                'RET.IB. CAP.FED': 10, 'RET.IB.CAP.FED': 10, 'RET.IB.CAP.FED.': 10,
                'RET.IB.CORDOBA': 11, 'RET.IB.CÓRDOBA': 11,
                'RET.IB.MENDOZA': 12,
                'RET.IB.MISIONES': 13,
                'RET.SIRCREB CORDOBA': 14, 'RET.SIRCREB CÓRDOBA': 14,
                'RET.SIRCREB MENDOZA': 15,
                'RET.SIRCREB JUJUY': 16,
                'RET. SIRCREB C.A.B.A': 17, 'RET.SIRCREB C.A.B.A': 17, 'RET.SIRCREB CABA': 17,
                'RET.SIRCREB R.NEGRO': 18, 'RET.SIRCREB RIO NEGRO': 18,
                'PERC.ADUANERA C.FED.': 19, 'PERC.ADUANERA C.FED': 19,
                'PERC.ADUANERA BSAS': 20, 'PERC.ADUANERA BS.AS.': 20,
                'PERCEP.ADUAN.CORDOBA': 21, 'PERCEP.ADUAN.CÓRDOBA': 21,
                'PERCEP.ADUAN.MENDOZA': 22,
                'SIRCREB CORRIENTES': 23,
                'PERC.ADUANA CORRIENT': 24, 'PERC.ADUANA CORRIENTES': 24,
                'PERC.ADUAN. RIO NEG.': 25, 'PERC.ADUAN. RIO NEG': 25, 'PERC.ADUAN.RIO NEG': 25,
                'PERC ADUANA JUJUY': 26, 'PERC.ADUANA JUJUY': 26,
            }

            def _get_deduccion_code(nombre):
                """Busca el código de deducción, primero exacto, luego por prefijo."""
                if nombre in DEDUCCION_ORDER_MAP:
                    return DEDUCCION_ORDER_MAP[nombre]
                # Buscar por prefijo (para variantes con números entre paréntesis, etc.)
                nombre_limpio = nombre.split('(')[0].strip()
                if nombre_limpio in DEDUCCION_ORDER_MAP:
                    return DEDUCCION_ORDER_MAP[nombre_limpio]
                return 999

            def _tasa_sort_key(item):
                es_deduccion = 1 if item.get('Ded_Col_M') else 0
                if es_deduccion:
                    codigo = _get_deduccion_code(item['Tasa'])
                else:
                    codigo = TASA_ORDER_MAP.get(item['Tasa'], 999)
                return (es_deduccion, codigo, item['Tasa'])

            res_imp_data.sort(key=_tasa_sort_key)

            n_ri_cols = 5
            ws_ri_name = 'Resumen x Impuesto'
            pd.DataFrame([{'Tasa': r['Tasa']} for r in res_imp_data]).to_excel(writer, sheet_name=ws_ri_name, index=False, startrow=5)
            ws_ri = writer.sheets[ws_ri_name]
        
            ws_ri.merge_cells(f'A1:{get_column_letter(n_ri_cols)}1')
            ws_ri['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
            ws_ri['A1'].font = title_font; ws_ri['A1'].fill = title_fill; ws_ri['A1'].alignment = center_align

            ws_ri.merge_cells(f'A2:{get_column_letter(n_ri_cols)}2')
            ws_ri['A2'] = f"{meta['tipo_reporte'].upper()} - RESUMEN POR IMPUESTO"
            ws_ri['A2'].font = report_type_font; ws_ri['A2'].alignment = center_align

            ws_ri.merge_cells(f'A3:{get_column_letter(n_ri_cols)}3')
            ws_ri['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
            ws_ri['A3'].font = Font(bold=True, size=11, color='2F5496'); ws_ri['A3'].alignment = center_align
        
            ri_headers = ['Tasa', 'Neto', 'IVA', 'Deducciones', 'Total']
            for col_idx, h in enumerate(ri_headers):
                cell = ws_ri.cell(row=res_header_row, column=col_idx+1) # Reusing res_header_row
                cell.value = h
                cell.font = header_font; cell.fill = header_fill; cell.alignment = header_align; cell.border = thin_border

            for idx, r_data in enumerate(res_imp_data):
                curr_row = res_data_start + idx # Reusing res_data_start
                ws_ri.cell(row=curr_row, column=1).value = r_data['Tasa']
                ws_ri.cell(row=curr_row, column=1).alignment = center_align
            
                # Neto
                if r_data['Neto_Col_M']: ws_ri.cell(row=curr_row, column=2).value = f"={mov_sheet_name}!{r_data['Neto_Col_M']}{total_row_mov}"
                else: ws_ri.cell(row=curr_row, column=2).value = 0.0
            
                # IVA
                if r_data['IVA_Col_M']: ws_ri.cell(row=curr_row, column=3).value = f"={mov_sheet_name}!{r_data['IVA_Col_M']}{total_row_mov}"
                else: ws_ri.cell(row=curr_row, column=3).value = 0.0

                # Deducciones
                if r_data['Ded_Col_M']: ws_ri.cell(row=curr_row, column=4).value = f"={mov_sheet_name}!{r_data['Ded_Col_M']}{total_row_mov}"
                else: ws_ri.cell(row=curr_row, column=4).value = 0.0
            
                # Total
                ws_ri.cell(row=curr_row, column=5).value = f"=B{curr_row}+C{curr_row}+D{curr_row}"
            
                for c in range(2, 6):
                    ws_ri.cell(row=curr_row, column=c).number_format = money_fmt
                    ws_ri.cell(row=curr_row, column=c).alignment = center_align

            total_row_ri = res_data_start + len(res_imp_data) # Reusing res_data_start
            ws_ri[f'A{total_row_ri}'] = "TOTAL GENERAL"
            ws_ri[f'A{total_row_ri}'].font = Font(bold=True); ws_ri[f'A{total_row_ri}'].alignment = Alignment(horizontal='right')
            for col_idx in range(2, 6):
                 col_l = get_column_letter(col_idx)
                 cell = ws_ri.cell(row=total_row_ri, column=col_idx)
                 cell.value = f'=SUM({col_l}{res_data_start}:{col_l}{total_row_ri-1})' # Reusing res_data_start
                 cell.font = Font(bold=True); cell.border = Border(top=Side(style='double'))
                 cell.number_format = money_fmt; cell.alignment = center_align

            _autofit(ws_ri, n_ri_cols)


            # ── Hoja Resumen por Tipo ─────────────────────────────
            res_tipo = resumen.groupby('Tipo').agg(
                **{c: (c, 'sum') for c in IVA_COL_ORDER},
                **{c: (c, 'sum') for c in individual_other_cols},
                Deducciones=('Total', 'count'), # placeholder
                Cantidad=('Total', 'count'),
            ).reset_index()
            res_tipo['Total'] = 0.0
            cols_order_rt = ['Tipo'] + IVA_COL_ORDER + individual_other_cols + ['Deducciones', 'Total', 'Cantidad']
            cols_order_rt = [c for c in cols_order_rt if c in res_tipo.columns]
            res_tipo = res_tipo[cols_order_rt]
            # Sort logic
            sum_cols = IVA_COL_ORDER + individual_other_cols + deduccion_cols
            res_tipo['_sort'] = resumen.groupby('Tipo')[sum_cols].sum().sum(axis=1).values
            res_tipo = res_tipo.sort_values('_sort', ascending=False).drop(columns='_sort')

            n_rt_cols = len(res_tipo.columns)
            # startrow=5 -> fila 6
            res_tipo.to_excel(writer, sheet_name='Resumen x Comprobante', index=False, startrow=5)
            ws3 = writer.sheets['Resumen x Comprobante']
        
            ws3.merge_cells(f'A1:{get_column_letter(n_rt_cols)}1')
            ws3['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
            ws3['A1'].font = title_font
            ws3['A1'].fill = title_fill
            ws3['A1'].alignment = center_align

            ws3.merge_cells(f'A2:{get_column_letter(n_rt_cols)}2')
            ws3['A2'] = meta['tipo_reporte'].upper() if meta['tipo_reporte'] else 'RESUMEN POR TIPO'
            ws3['A2'].font = report_type_font
            ws3['A2'].alignment = center_align

            ws3.merge_cells(f'A3:{get_column_letter(n_rt_cols)}3')
            ws3['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
            ws3['A3'].font = Font(bold=True, size=11, color='2F5496')
            ws3['A3'].alignment = center_align

            ws3.merge_cells(f'A4:{get_column_letter(n_rt_cols)}4')
            ws3['A4'] = f'Total: {len(res_tipo)} tipos'
            ws3['A4'].font = Font(italic=True, size=10, color='4472C4')
            ws3['A4'].alignment = center_align

            for col_idx in range(1, n_rt_cols + 1):
                cell = ws3.cell(row=res_header_row, column=col_idx) # reusando res_header_row=6
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border

            rt_col_list = list(res_tipo.columns)
            first_iva_idx_rt = rt_col_list.index(IVA_COL_ORDER[0]) + 1
            ded_idx_rt = rt_col_list.index('Deducciones') + 1
            total_idx_rt = rt_col_list.index('Total') + 1
            first_iva_letter_rt = get_column_letter(first_iva_idx_rt)
            ded_letter_rt = get_column_letter(ded_idx_rt)

            for row in range(res_data_start, len(res_tipo) + res_data_start):
                for col_idx in range(1, n_rt_cols + 1):
                    cell = ws3.cell(row=row, column=col_idx)
                    cell.alignment = center_align
                    col_name = rt_col_list[col_idx - 1]
                
                    if first_iva_idx_rt <= col_idx <= total_idx_rt:
                        if col_name in col_list:
                            v_col = get_column_letter(col_list.index(col_name) + 1)
                            cell.value = f'=SUMIFS(Movimientos!${v_col}${7}:${v_col}${n_mov+7-1}, Movimientos!${mov_tipo_col}${7}:${mov_tipo_col}${n_mov+7-1}, $A{row})'
                            cell.number_format = money_fmt
                        elif col_name == 'Deducciones':
                            formula_parts = []
                            for dc in deduccion_cols:
                                v_col = get_column_letter(col_list.index(dc) + 1)
                                formula_parts.append(f'SUMIFS(Movimientos!${v_col}${7}:${v_col}${n_mov+7-1}, Movimientos!${mov_tipo_col}${7}:${mov_tipo_col}${n_mov+7-1}, $A{row})')
                            cell.value = '=' + '+'.join(formula_parts) if formula_parts else 0
                            cell.number_format = money_fmt
                        elif col_name == 'Total':
                            # Sumar desde el primer IVA hasta Deducciones
                            cell.value = f'=SUM({first_iva_letter_rt}{row}:{ded_letter_rt}{row})'
                            cell.number_format = money_fmt
                    elif col_name == 'Cantidad':
                        cell.value = f'=COUNTIFS(Movimientos!${mov_tipo_col}${7}:${mov_tipo_col}${n_mov+7-1}, $A{row})'

            # Fila TOTAL GENERAL
            total_row_rt = len(res_tipo) + res_data_start
            ws3.merge_cells(f'A{total_row_rt}:A{total_row_rt}')
            ws3[f'A{total_row_rt}'] = "TOTAL GENERAL"
            ws3[f'A{total_row_rt}'].font = Font(bold=True)
            ws3[f'A{total_row_rt}'].alignment = Alignment(horizontal='right')
        
            for col_idx in range(first_iva_idx_rt, n_rt_cols + 1):
                col_letter = get_column_letter(col_idx)
                cell = ws3.cell(row=total_row_rt, column=col_idx)
                cell.value = f'=SUM({col_letter}{res_data_start}:{col_letter}{total_row_rt-1})'
                cell.font = Font(bold=True)
                cell.border = Border(top=Side(style='double'))
                if col_idx < n_rt_cols:
                    cell.number_format = money_fmt

            _autofit(ws3, n_rt_cols)

            # ── Hoja Resumen por Concepto ─────────────────────────
            res_conc = resumen.groupby('Concepto').agg(
                **{c: (c, 'sum') for c in IVA_COL_ORDER},
                **{c: (c, 'sum') for c in individual_other_cols},
                Deducciones=('Total', 'count'), # placeholder
                Cantidad=('Total', 'count'),
            ).reset_index()
        
            # Ordenar por Concepto numérico
            res_conc['Concepto_Num'] = pd.to_numeric(res_conc['Concepto'], errors='coerce')
            res_conc = res_conc.sort_values('Concepto_Num').drop(columns='Concepto_Num')
        
            res_conc['Descripcion'] = res_conc['Concepto'].apply(
                lambda x: CONCEPTOS_MAP.get(str(x), "").replace("°", "o.").upper()
            )
        
            res_conc['Total'] = 0.0
            cols_order_rc = ['Concepto', 'Descripcion'] + IVA_COL_ORDER + individual_other_cols + ['Deducciones', 'Total', 'Cantidad']
            cols_order_rc = [c for c in cols_order_rc if c in res_conc.columns]
            res_conc = res_conc[cols_order_rc]

            n_rc_cols = len(res_conc.columns)
            # startrow=5 -> fila 6
            res_conc.to_excel(writer, sheet_name='Resumen x Concepto', index=False, startrow=5)
            ws4 = writer.sheets['Resumen x Concepto']
        
            ws4.merge_cells(f'A1:{get_column_letter(n_rc_cols)}1')
            ws4['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
            ws4['A1'].font = title_font
            ws4['A1'].fill = title_fill
            ws4['A1'].alignment = center_align

            ws4.merge_cells(f'A2:{get_column_letter(n_rc_cols)}2')
            ws4['A2'] = meta['tipo_reporte'].upper() if meta['tipo_reporte'] else 'RESUMEN POR CONCEPTO'
            ws4['A2'].font = report_type_font
            ws4['A2'].alignment = center_align

            ws4.merge_cells(f'A3:{get_column_letter(n_rc_cols)}3')
            ws4['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
            ws4['A3'].font = Font(bold=True, size=11, color='2F5496')
            ws4['A3'].alignment = center_align

            ws4.merge_cells(f'A4:{get_column_letter(n_rc_cols)}4')
            ws4['A4'] = f'Total: {len(res_conc)} conceptos'
            ws4['A4'].font = Font(italic=True, size=10, color='4472C4')
            ws4['A4'].alignment = center_align

            for col_idx in range(1, n_rc_cols + 1):
                cell = ws4.cell(row=res_header_row, column=col_idx) # reusando res_header_row=6
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border

            rc_col_list = list(res_conc.columns)
            first_iva_idx_rc = rc_col_list.index(IVA_COL_ORDER[0]) + 1
            ded_idx_rc = rc_col_list.index('Deducciones') + 1
            total_idx_rc = rc_col_list.index('Total') + 1
            first_iva_letter_rc = get_column_letter(first_iva_idx_rc)
            ded_letter_rc = get_column_letter(ded_idx_rc)

            for row in range(res_data_start, len(res_conc) + res_data_start):
                for col_idx in range(1, n_rc_cols + 1):
                    cell = ws4.cell(row=row, column=col_idx)
                    cell.alignment = center_align
                    col_name = rc_col_list[col_idx - 1]

                    if first_iva_idx_rc <= col_idx <= total_idx_rc:
                        if col_name in col_list:
                            v_col = get_column_letter(col_list.index(col_name) + 1)
                            cell.value = f'=SUMIFS(Movimientos!${v_col}${7}:${v_col}${n_mov+7-1}, Movimientos!${mov_conc_col}${7}:${mov_conc_col}${n_mov+7-1}, $A{row})'
                            cell.number_format = money_fmt
                        elif col_name == 'Deducciones':
                            formula_parts = []
                            for dc in deduccion_cols:
                                v_col = get_column_letter(col_list.index(dc) + 1)
                                formula_parts.append(f'SUMIFS(Movimientos!${v_col}${7}:${v_col}${n_mov+7-1}, Movimientos!${mov_conc_col}${7}:${mov_conc_col}${n_mov+7-1}, $A{row})')
                            cell.value = '=' + '+'.join(formula_parts) if formula_parts else 0
                            cell.number_format = money_fmt
                        elif col_name == 'Total':
                            cell.value = f'=SUM({first_iva_letter_rc}{row}:{ded_letter_rc}{row})'
                            cell.number_format = money_fmt
                    elif col_name == 'Cantidad':
                        cell.value = f'=COUNTIFS(Movimientos!${mov_conc_col}${7}:${mov_conc_col}${n_mov+7-1}, $A{row})'
        
            # Fila TOTAL GENERAL
            total_row_rc = len(res_conc) + res_data_start
            ws4.merge_cells(f'A{total_row_rc}:B{total_row_rc}')
            ws4[f'A{total_row_rc}'] = "TOTAL GENERAL"
            ws4[f'A{total_row_rc}'].font = Font(bold=True)
            ws4[f'A{total_row_rc}'].alignment = Alignment(horizontal='right')
        
            for col_idx in range(first_iva_idx_rc, n_rc_cols + 1):
                col_letter = get_column_letter(col_idx)
                cell = ws4.cell(row=total_row_rc, column=col_idx)
                cell.value = f'=SUM({col_letter}{res_data_start}:{col_letter}{total_row_rc-1})'
                cell.font = Font(bold=True)
                cell.border = Border(top=Side(style='double'))
                if col_idx < n_rc_cols:
                    cell.number_format = money_fmt

            _autofit(ws4, n_rc_cols)

            # ── Hoja Resumen por Concepto y Jur. (Pivot para CM05) ──
            # 1. Identificar columnas que forman parte del "Neto" (Base Imponible)
            # Incluimos Netos, Exento, Monotributo y cualquier otro que no sea IVA/PERC/RET/DEDUCC
            cm05_neto_cols = [
                c for c in IVA_COL_ORDER 
                if any(x in c for x in ['Neto', 'Exento', 'Monotributo'])
            ]
            otros_adicionales = [
                c for c in other_cols 
                if not any(x in c.upper() for x in ['PERC', 'RET', 'DEDUCC', 'IVA'])
            ]
            cm05_neto_cols += otros_adicionales
        
            # 2. Obtener Jurisdicciones y Conceptos únicos
            unique_jurs = sorted([str(j) for j in df['Jur.'].unique() if pd.notna(j) and str(j).strip()])
            if not unique_jurs: unique_jurs = ["S/D"]
        
            conceptos_unicos_df = res_conc[['Concepto', 'Descripcion']].copy()
        
            n_rj_cols = 3 + len(unique_jurs) # Concepto, Desc, Jurs..., Total
            res_jur_sheet_name = 'Resumen x Concepto y Jur.'
            if res_jur_sheet_name in writer.book.sheetnames:
                del writer.book[res_jur_sheet_name]
            
            # startrow=5 -> fila 6
            conceptos_unicos_df.to_excel(writer, sheet_name='Resumen x Concepto y Jur.', index=False, startrow=5)
            ws_rj = writer.sheets['Resumen x Concepto y Jur.']
        
            # Titulos y Estilos
            ws_rj.merge_cells(f'A1:{get_column_letter(n_rj_cols)}1')
            ws_rj['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
            ws_rj['A1'].font = title_font
            ws_rj['A1'].fill = title_fill
            ws_rj['A1'].alignment = center_align

            ws_rj.merge_cells(f'A2:{get_column_letter(n_rj_cols)}2')
            ws_rj['A2'] = meta['tipo_reporte'].upper() if meta['tipo_reporte'] else 'RESUMEN POR CONCEPTO Y JUR.'
            ws_rj['A2'].font = report_type_font
            ws_rj['A2'].alignment = center_align

            ws_rj.merge_cells(f'A3:{get_column_letter(n_rj_cols)}3')
            ws_rj['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
            ws_rj['A3'].font = Font(bold=True, size=11, color='2F5496')
            ws_rj['A3'].alignment = center_align

            ws_rj.merge_cells(f'A4:{get_column_letter(n_rj_cols)}4')
            ws_rj['A4'] = f'Total: {len(conceptos_unicos_df)} conceptos x {len(unique_jurs)} jur.'
            ws_rj['A4'].font = Font(italic=True, size=10, color='4472C4')
            ws_rj['A4'].alignment = center_align

            ws_rj.cell(row=res_header_row, column=1).value = 'Concepto'
            ws_rj.cell(row=res_header_row, column=2).value = 'Descripcion'
            for i, jur in enumerate(unique_jurs):
                cell = ws_rj.cell(row=res_header_row, column=3+i)
                cell.value = f"Jur {jur}"
            ws_rj.cell(row=res_header_row, column=3+len(unique_jurs)).value = 'TOTAL'
        
            for col_idx in range(1, n_rj_cols + 1):
                cell = ws_rj.cell(row=res_header_row, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border

            mov_jur_col = get_column_letter(col_list.index('Jur.') + 1)
            for idx_c, (idx_df, row_data) in enumerate(conceptos_unicos_df.iterrows()):
                curr_row = res_data_start + idx_c
                ws_rj.cell(row=curr_row, column=1).value = row_data['Concepto']
                ws_rj.cell(row=curr_row, column=2).value = row_data['Descripcion']
            
                for j_idx, jur in enumerate(unique_jurs):
                    col_target = 3 + j_idx
                    formula_parts = []
                    for n_col in cm05_neto_cols:
                        v_col_l = get_column_letter(col_list.index(n_col) + 1)
                        formula_parts.append(
                            f'SUMIFS(Movimientos!${v_col_l}${7}:${v_col_l}${n_mov+7-1}, '
                            f'Movimientos!${mov_conc_col}${7}:${mov_conc_col}${n_mov+7-1}, $A{curr_row}, '
                            f'Movimientos!${mov_jur_col}${7}:${mov_jur_col}${n_mov+7-1}, "{jur}")'
                        )
                    ws_rj.cell(row=curr_row, column=col_target).value = "=" + "+".join(formula_parts) if formula_parts else 0
                    ws_rj.cell(row=curr_row, column=col_target).number_format = money_fmt
                    ws_rj.cell(row=curr_row, column=col_target).alignment = center_align
            
                first_jur_letter = get_column_letter(3)
                last_jur_letter = get_column_letter(3 + len(unique_jurs) - 1)
                total_cell = ws_rj.cell(row=curr_row, column=3 + len(unique_jurs))
                total_cell.value = f'=SUM({first_jur_letter}{curr_row}:{last_jur_letter}{curr_row})'
                total_cell.number_format = money_fmt
                total_cell.alignment = center_align
                total_cell.font = Font(bold=True)

            total_row_rj = res_data_start + len(conceptos_unicos_df)
            ws_rj.merge_cells(f'A{total_row_rj}:B{total_row_rj}')
            ws_rj[f'A{total_row_rj}'] = "TOTAL GENERAL"
            ws_rj[f'A{total_row_rj}'].font = Font(bold=True)
            ws_rj[f'A{total_row_rj}'].alignment = Alignment(horizontal='right')
        
            for col_idx in range(3, n_rj_cols + 1):
                col_l = get_column_letter(col_idx)
                cell = ws_rj.cell(row=total_row_rj, column=col_idx)
                cell.value = f'=SUM({col_l}{res_data_start}:{col_l}{total_row_rj-1})'
                cell.font = Font(bold=True)
                cell.border = Border(top=Side(style='double'))
                cell.number_format = money_fmt
                cell.alignment = center_align

            _autofit(ws_rj, n_rj_cols)

            # ── Hoja Resumen por Proveedor (agrupado por CUIT) ────
            res = resumen.groupby('CUIT').agg(
                Proveedor=('Proveedor', 'first'),
                **{c: (c, 'sum') for c in IVA_COL_ORDER},
                **{c: (c, 'sum') for c in individual_other_cols},
                Deducciones=('Total', 'count'), # placeholder
                Cantidad=('Total', 'count'),
            ).reset_index()

            res['Total'] = 0.0
            cols_order = ['CUIT', 'Proveedor'] + IVA_COL_ORDER + individual_other_cols + ['Deducciones', 'Total', 'Cantidad']
            cols_order = [c for c in cols_order if c in res.columns]
            res = res[cols_order]
            # Sort logic
            res['_sort'] = resumen.groupby('CUIT')[sum_cols].sum().sum(axis=1).values
            res = res.sort_values('_sort', ascending=False).drop(columns='_sort')

            res.to_excel(writer, sheet_name='Resumen x Proveedor', index=False, startrow=5)
            ws2 = writer.sheets['Resumen x Proveedor']
            n_res_cols = len(res.columns)

            ws2.merge_cells(f'A1:{get_column_letter(n_res_cols)}1')
            ws2['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
            ws2['A1'].font = title_font; ws2['A1'].fill = title_fill; ws2['A1'].alignment = center_align

            ws2.merge_cells(f'A2:{get_column_letter(n_res_cols)}2')
            ws2['A2'] = meta['tipo_reporte'].upper() if meta['tipo_reporte'] else 'RESUMEN POR PROVEEDOR'
            ws2['A2'].font = report_type_font; ws2['A2'].alignment = center_align
        
            ws2.merge_cells(f'A3:{get_column_letter(n_res_cols)}3')
            ws2['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
            ws2['A3'].font = Font(bold=True, size=11, color='2F5496'); ws2['A3'].alignment = center_align

            ws2.merge_cells(f'A4:{get_column_letter(n_res_cols)}4')
            ws2['A4'] = f'Total: {len(res)} proveedores'
            ws2['A4'].font = Font(italic=True, size=10, color='4472C4'); ws2['A4'].alignment = center_align
        
            for col_idx in range(1, n_res_cols + 1):
                cell = ws2.cell(row=res_header_row, column=col_idx)
                cell.font = header_font; cell.fill = header_fill; cell.alignment = header_align; cell.border = thin_border

            res_col_list = list(res.columns)
            first_iva_idx_res = res_col_list.index(IVA_COL_ORDER[0]) + 1
            ded_idx_res = res_col_list.index('Deducciones') + 1
            total_idx_res = res_col_list.index('Total') + 1
            first_iva_letter_res = get_column_letter(first_iva_idx_res)
            ded_letter_res = get_column_letter(ded_idx_res)
        
            for row in range(res_data_start, len(res) + res_data_start):
                for col_idx in range(1, n_res_cols + 1):
                    cell = ws2.cell(row=row, column=col_idx)
                    cell.alignment = center_align
                    col_name = res_col_list[col_idx - 1]
                
                    if first_iva_idx_res <= col_idx <= total_idx_res:
                        if col_name in col_list:
                            v_col = get_column_letter(col_list.index(col_name) + 1)
                            cell.value = f'=SUMIFS(Movimientos!${v_col}${7}:${v_col}${n_mov+7-1}, Movimientos!${mov_cuit_col}${7}:${mov_cuit_col}${n_mov+7-1}, $A{row})'
                            cell.number_format = money_fmt
                        elif col_name == 'Deducciones':
                            formula_parts = []
                            for dc in deduccion_cols:
                                v_col = get_column_letter(col_list.index(dc) + 1)
                                formula_parts.append(f'SUMIFS(Movimientos!${v_col}${7}:${v_col}${n_mov+7-1}, Movimientos!${mov_cuit_col}${7}:${mov_cuit_col}${n_mov+7-1}, $A{row})')
                            cell.value = '=' + '+'.join(formula_parts) if formula_parts else 0
                            cell.number_format = money_fmt
                        elif col_name == 'Total':
                            cell.value = f'=SUM({first_iva_letter_res}{row}:{ded_letter_res}{row})'
                            cell.number_format = money_fmt
                    elif col_name == 'Cantidad':
                        cell.value = f'=COUNTIFS(Movimientos!${mov_cuit_col}${7}:${mov_cuit_col}${n_mov+7-1}, $A{row})'
                
                    # Formato CUIT como texto

            total_row = len(res) + res_data_start
            ws2.merge_cells(f'A{total_row}:B{total_row}')
            ws2[f'A{total_row}'] = "TOTAL GENERAL"
            ws2[f'A{total_row}'].font = Font(bold=True); ws2[f'A{total_row}'].alignment = Alignment(horizontal='right')
        
            for col_idx in range(first_iva_idx_res, n_res_cols + 1):
                col_letter = get_column_letter(col_idx)
                cell = ws2.cell(row=total_row, column=col_idx)
                cell.value = f'=SUM({col_letter}{res_data_start}:{col_letter}{total_row-1})'
                cell.font = Font(bold=True); cell.border = Border(top=Side(style='double'))
                if col_idx < n_res_cols: cell.number_format = money_fmt
            
            _autofit(ws2, n_res_cols)

            # ── Hoja Mayor x Proveedor ────────────────────────────
            df_with_idx = df.copy()
            # Original data rows in Movimientos start at Row 7
            df_with_idx['_orig_row'] = range(7, len(df) + 7)
        
            mayor = df_with_idx.sort_values(['CUIT', 'Fecha'])
        
            def format_comp(r):
                pv_s = f"{r['PV']:05d}"
                nro_s = f"{r['Nro.']:08d}" if isinstance(r['Nro.'], int) else str(r['Nro.'])
                return f"{pv_s}-{nro_s}{r['Letra']}"
            
            mayor['Comp.'] = mayor.apply(format_comp, axis=1)
            mayor['Saldo Acumulado'] = mayor.groupby('CUIT')['Total'].cumsum()
        
            cols_mayor = ['CUIT', 'Proveedor', 'Fecha', 'Tipo', 'Comp.', 'Concepto', 'Total', 'Saldo Acumulado', '_orig_row']
            mayor = mayor[cols_mayor]
        
            n_mayor_cols = len(mayor.columns) - 1
            # startrow=5 -> fila 6
            mayor.to_excel(writer, sheet_name='Mayor x Proveedor', index=False, startrow=5)
            ws5 = writer.sheets['Mayor x Proveedor']

            ws5.merge_cells(f'A1:{get_column_letter(n_mayor_cols)}1')
            ws5['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
            ws5['A1'].font = title_font
            ws5['A1'].fill = title_fill
            ws5['A1'].alignment = center_align

            ws5.merge_cells(f'A2:{get_column_letter(n_mayor_cols)}2')
            ws5['A2'] = meta['tipo_reporte'].upper() if meta['tipo_reporte'] else 'MAYOR AUXILIAR'
            ws5['A2'].font = report_type_font
            ws5['A2'].alignment = center_align

            ws5.merge_cells(f'A3:{get_column_letter(n_mayor_cols)}3')
            ws5['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
            ws5['A3'].font = Font(bold=True, size=11, color='2F5496')
            ws5['A3'].alignment = center_align

            ws5.merge_cells(f'A4:{get_column_letter(n_mayor_cols)}4')
            ws5['A4'] = f'Total: {len(mayor)} movimientos'
            ws5['A4'].font = Font(italic=True, size=10, color='4472C4')
            ws5['A4'].alignment = center_align
        
            total_mov_col = get_column_letter(col_list.index('Total') + 1)

            for col_idx in range(1, n_mayor_cols + 1):
                cell = ws5.cell(row=res_header_row, column=col_idx) # reusando res_header_row=6
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border
            
            for row_idx in range(res_data_start, len(mayor) + res_data_start): # res_data_start=7
                orig_row_idx = row_idx - res_data_start
                orig_row = mayor.iloc[orig_row_idx]['_orig_row']
                for col_idx in range(1, n_mayor_cols + 1):
                    cell = ws5.cell(row=row_idx, column=col_idx)
                    cell.alignment = center_align
                    if col_idx == 7: # Total
                        cell.value = f'={mov_sheet_name}!{total_mov_col}{orig_row}'
                        cell.number_format = money_fmt
                    elif col_idx == 8: # Saldo Acumulado
                        # Formula unificada: IF(mismo CUIT que arriba, saldo_ant + total_actual, total_actual)
                        if row_idx == res_data_start:
                            cell.value = f'=G{row_idx}'
                        else:
                            cell.value = f'=IF(A{row_idx}=A{row_idx-1}, H{row_idx-1}+G{row_idx}, G{row_idx})'
                        cell.number_format = money_fmt
        
            ws5.delete_cols(n_mayor_cols + 1) # Borrar columna auxiliar
            _autofit(ws5, n_mayor_cols)

        # ── Hoja ARCA (datos del CSV de ARCA) ──────────────────
        if cruce_arca and df_arca is not None and not df_arca.empty:
            df_arca.to_excel(writer, sheet_name='ARCA', index=False, startrow=5)
            ws_arca = writer.sheets['ARCA']
            n_arca_cols = len(df_arca.columns)

            ws_arca.merge_cells(f'A1:{get_column_letter(n_arca_cols)}1')
            ws_arca['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
            ws_arca['A1'].font = title_font; ws_arca['A1'].fill = title_fill; ws_arca['A1'].alignment = center_align

            ws_arca.merge_cells(f'A2:{get_column_letter(n_arca_cols)}2')
            ws_arca['A2'] = f"{meta['tipo_reporte'].upper()} - COMPROBANTES ARCA"
            ws_arca['A2'].font = report_type_font; ws_arca['A2'].alignment = center_align

            ws_arca.merge_cells(f'A3:{get_column_letter(n_arca_cols)}3')
            ws_arca['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
            ws_arca['A3'].font = Font(bold=True, size=11, color='2F5496'); ws_arca['A3'].alignment = center_align

            ws_arca.merge_cells(f'A4:{get_column_letter(n_arca_cols)}4')
            ws_arca['A4'] = f'Total: {len(df_arca)} comprobantes'
            ws_arca['A4'].font = Font(italic=True, size=10, color='4472C4'); ws_arca['A4'].alignment = center_align

            for col_idx in range(1, n_arca_cols + 1):
                cell = ws_arca.cell(row=6, column=col_idx)
                cell.font = header_font; cell.fill = header_fill
                cell.alignment = header_align; cell.border = thin_border

            # Identificar columnas monetarias
            arca_col_list_final = list(df_arca.columns)
            non_money = {'Fecha', 'Comprobante', 'PV', 'Nro.', 'Tipo Doc.', 'CUIT', 'Razon Social', 'Auxiliar'}
            arca_money_indices = []
            for ci, cn in enumerate(arca_col_list_final):
                if cn not in non_money and df_arca[cn].dtype in ('float64', 'int64', 'float32', 'int32'):
                    arca_money_indices.append(ci + 1)

            for row_idx in range(7, len(df_arca) + 7):
                for col_idx in range(1, n_arca_cols + 1):
                    cell = ws_arca.cell(row=row_idx, column=col_idx)
                    cell.alignment = center_align
                    if col_idx in arca_money_indices:
                        cell.number_format = accounting_fmt
                if (row_idx - 7) % 2 == 0:
                    for col_idx in range(1, n_arca_cols + 1):
                        ws_arca.cell(row=row_idx, column=col_idx).fill = zebra_fill

            # ── CRUCE + DIFF en ARCA (busca en SISTEMA) ──────────────
            if 'Auxiliar' in arca_col_list_final and 'Total' in arca_col_list_final and 'Auxiliar' in col_list:
                arca_cruce_col_idx = n_arca_cols + 1
                arca_diff_col_idx = n_arca_cols + 2
                arca_cruce_letter = get_column_letter(arca_cruce_col_idx)
                arca_diff_letter = get_column_letter(arca_diff_col_idx)
                arca_aux_col_idx = arca_col_list_final.index('Auxiliar') + 1
                arca_aux_letter = get_column_letter(arca_aux_col_idx)
                arca_total_col_idx = arca_col_list_final.index('Total') + 1
                arca_total_letter = get_column_letter(arca_total_col_idx)

                # Rango de lookup en SISTEMA
                sys_aux_letter = get_column_letter(col_list.index('Auxiliar') + 1)
                sys_total_letter = get_column_letter(col_list.index('Total') + 1)
                sys_total_offset = col_list.index('Total') - col_list.index('Auxiliar') + 1
                sys_last_data_row = len(df) + 6

                # Headers CRUCE
                ws_arca.cell(row=6, column=arca_cruce_col_idx).value = 'CRUCE'
                ws_arca.cell(row=6, column=arca_cruce_col_idx).font = header_font
                ws_arca.cell(row=6, column=arca_cruce_col_idx).fill = PatternFill('solid', fgColor='7030A0')
                ws_arca.cell(row=6, column=arca_cruce_col_idx).alignment = header_align
                ws_arca.cell(row=6, column=arca_cruce_col_idx).border = thin_border

                ws_arca.cell(row=6, column=arca_diff_col_idx).value = 'DIFF'
                ws_arca.cell(row=6, column=arca_diff_col_idx).font = header_font
                ws_arca.cell(row=6, column=arca_diff_col_idx).fill = PatternFill('solid', fgColor='7030A0')
                ws_arca.cell(row=6, column=arca_diff_col_idx).alignment = header_align
                ws_arca.cell(row=6, column=arca_diff_col_idx).border = thin_border

                for row_idx in range(7, len(df_arca) + 7):
                    ws_arca.cell(row=row_idx, column=arca_cruce_col_idx).value = (
                        f'=IFERROR(VLOOKUP({arca_aux_letter}{row_idx},'
                        f"SISTEMA!${sys_aux_letter}$7:${sys_total_letter}${sys_last_data_row},"
                        f'{sys_total_offset},FALSE),"NO ENCONTRADO")'
                    )
                    ws_arca.cell(row=row_idx, column=arca_cruce_col_idx).number_format = accounting_fmt
                    ws_arca.cell(row=row_idx, column=arca_cruce_col_idx).alignment = center_align

                    ws_arca.cell(row=row_idx, column=arca_diff_col_idx).value = (
                        f'=IF({arca_cruce_letter}{row_idx}="NO ENCONTRADO","",'
                        f'{arca_total_letter}{row_idx}-{arca_cruce_letter}{row_idx})'
                    )
                    ws_arca.cell(row=row_idx, column=arca_diff_col_idx).number_format = accounting_fmt
                    ws_arca.cell(row=row_idx, column=arca_diff_col_idx).alignment = center_align

            _autofit(ws_arca, n_arca_cols + 2)

            # ── Hojas de overflow: DE MAS EN SISTEMA / FALTANTES ARCA ─────
            # Construir sets de auxiliares para comparar
            if 'Auxiliar' in arca_col_list_final and 'Auxiliar' in col_list:
                # Auxiliar de ARCA: valores del df
                arca_aux_set = set(df_arca['Auxiliar'].dropna().astype(str).values)
                # Auxiliar de SISTEMA: construir igual que la formula
                sistema_aux_values = (
                    df['Tipo'].astype(str) + ' ' + df['Letra'].astype(str) +
                    df['PV'].astype(str) + df['Nro.'].astype(str) + df['CUIT'].astype(str)
                )
                sistema_aux_set = set(sistema_aux_values.values)

                # DE MAS EN SISTEMA: filas del SISTEMA no encontradas en ARCA
                mask_extra_sistema = ~sistema_aux_values.isin(arca_aux_set)
                df_extra_sistema = df[mask_extra_sistema].copy()
                if 'Auxiliar' in df_extra_sistema.columns:
                    df_extra_sistema = df_extra_sistema.drop(columns=['Auxiliar'])
                if not df_extra_sistema.empty:
                    df_extra_sistema.to_excel(writer, sheet_name='DE MAS EN SISTEMA', index=False, startrow=5)
                    ws_extra = writer.sheets['DE MAS EN SISTEMA']
                    n_extra_cols = len(df_extra_sistema.columns)
                    ws_extra.merge_cells(f'A1:{get_column_letter(n_extra_cols)}1')
                    ws_extra['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
                    ws_extra['A1'].font = title_font; ws_extra['A1'].fill = title_fill; ws_extra['A1'].alignment = center_align
                    ws_extra.merge_cells(f'A2:{get_column_letter(n_extra_cols)}2')
                    ws_extra['A2'] = 'DE MAS EN SISTEMA'
                    ws_extra['A2'].font = Font(bold=True, size=14, color='FFFFFF')
                    ws_extra['A2'].fill = PatternFill('solid', fgColor='C00000')
                    ws_extra['A2'].alignment = center_align
                    ws_extra.merge_cells(f'A3:{get_column_letter(n_extra_cols)}3')
                    ws_extra['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
                    ws_extra['A3'].font = Font(bold=True, size=11, color='2F5496'); ws_extra['A3'].alignment = center_align
                    ws_extra.merge_cells(f'A4:{get_column_letter(n_extra_cols)}4')
                    ws_extra['A4'] = f'{len(df_extra_sistema)} comprobantes en SISTEMA no encontrados en ARCA'
                    ws_extra['A4'].font = Font(italic=True, size=10, color='C00000')
                    ws_extra['A4'].alignment = center_align
                    for ci in range(1, n_extra_cols + 1):
                        c = ws_extra.cell(row=6, column=ci)
                        c.font = header_font; c.fill = header_fill
                        c.alignment = header_align; c.border = thin_border
                    # Aplicar formato contabilidad Peso a columnas numéricas
                    extra_col_list = list(df_extra_sistema.columns)
                    extra_non_money = {'Fecha', 'Tipo', 'PV', 'Nro.', 'Letra', 'Proveedor', 'Cond. IVA', 'CUIT', 'Concepto', 'Jur.'}
                    # Calcular rango de SUM para la columna Total (misma lógica que Movimientos)
                    extra_total_col_idx = extra_col_list.index('Total') + 1 if 'Total' in extra_col_list else None
                    extra_first_sum = None
                    extra_last_sum = None
                    if extra_total_col_idx:
                        # Buscar primera columna IVA presente
                        for iva_c in IVA_COL_ORDER:
                            if iva_c in extra_col_list:
                                extra_first_sum = get_column_letter(extra_col_list.index(iva_c) + 1)
                                break
                        # Buscar última columna antes de Total (other_cols o última IVA)
                        if other_cols:
                            for oc in reversed(other_cols):
                                if oc in extra_col_list:
                                    extra_last_sum = get_column_letter(extra_col_list.index(oc) + 1)
                                    break
                        if not extra_last_sum:
                            for iva_c in reversed(IVA_COL_ORDER):
                                if iva_c in extra_col_list:
                                    extra_last_sum = get_column_letter(extra_col_list.index(iva_c) + 1)
                                    break
                    for row_idx in range(7, len(df_extra_sistema) + 7):
                        for ci, cn in enumerate(extra_col_list):
                            cell = ws_extra.cell(row=row_idx, column=ci + 1)
                            cell.alignment = center_align
                            if cn not in extra_non_money:
                                cell.number_format = accounting_fmt
                        # Formula SUM en columna Total
                        if extra_total_col_idx and extra_first_sum and extra_last_sum:
                            ws_extra.cell(row=row_idx, column=extra_total_col_idx).value = f'=SUM({extra_first_sum}{row_idx}:{extra_last_sum}{row_idx})'
                            ws_extra.cell(row=row_idx, column=extra_total_col_idx).number_format = accounting_fmt

                    _autofit(ws_extra, n_extra_cols)

                # FALTANTES ARCA: filas de ARCA no encontradas en SISTEMA
                mask_falt_arca = ~df_arca['Auxiliar'].astype(str).isin(sistema_aux_set)
                df_falt_arca = df_arca[mask_falt_arca].copy()
                if 'Auxiliar' in df_falt_arca.columns:
                    df_falt_arca = df_falt_arca.drop(columns=['Auxiliar'])
                if not df_falt_arca.empty:
                    df_falt_arca.to_excel(writer, sheet_name='FALTANTES ARCA', index=False, startrow=5)
                    ws_falt = writer.sheets['FALTANTES ARCA']
                    n_falt_cols = len(df_falt_arca.columns)
                    ws_falt.merge_cells(f'A1:{get_column_letter(n_falt_cols)}1')
                    ws_falt['A1'] = meta['razon_social'].upper() if meta['razon_social'] else 'CONTRIBUYENTE'
                    ws_falt['A1'].font = title_font; ws_falt['A1'].fill = title_fill; ws_falt['A1'].alignment = center_align
                    ws_falt.merge_cells(f'A2:{get_column_letter(n_falt_cols)}2')
                    ws_falt['A2'] = f"Compras Faltantes ({meta['periodo']})"
                    ws_falt['A2'].font = Font(bold=True, size=14, color='FFFFFF')
                    ws_falt['A2'].fill = PatternFill('solid', fgColor='C00000')
                    ws_falt['A2'].alignment = center_align
                    ws_falt.merge_cells(f'A3:{get_column_letter(n_falt_cols)}3')
                    ws_falt['A3'] = f"CUIT: {meta['cuit_empresa']} | Periodo: {meta['periodo']}"
                    ws_falt['A3'].font = Font(bold=True, size=11, color='2F5496'); ws_falt['A3'].alignment = center_align
                    ws_falt.merge_cells(f'A4:{get_column_letter(n_falt_cols)}4')
                    ws_falt['A4'] = f'{len(df_falt_arca)} comprobantes en ARCA no encontrados en SISTEMA'
                    ws_falt['A4'].font = Font(italic=True, size=10, color='C00000')
                    ws_falt['A4'].alignment = center_align
                    for ci in range(1, n_falt_cols + 1):
                        c = ws_falt.cell(row=6, column=ci)
                        c.font = header_font; c.fill = header_fill
                        c.alignment = header_align; c.border = thin_border
                    # Aplicar formato contabilidad Peso a columnas numéricas
                    falt_col_list = list(df_falt_arca.columns)
                    falt_non_money = {'Fecha', 'Comprobante', 'PV', 'Nro.', 'Tipo Doc.', 'CUIT', 'Razon Social'}
                    for row_idx in range(7, len(df_falt_arca) + 7):
                        for ci, cn in enumerate(falt_col_list):
                            cell = ws_falt.cell(row=row_idx, column=ci + 1)
                            cell.alignment = center_align
                            if cn not in falt_non_money:
                                cell.number_format = accounting_fmt
                    _autofit(ws_falt, n_falt_cols)

    print(f"\n  Excel guardado en: {output_path}")


def generar_sifere_txt(transacciones: list[dict], meta: dict) -> str:
    """Genera un archivo TXT con formato SIFERE para percepciones de IIBB.
    Cada línea: CodJurisdiccion(3) + CUIT(11) + Fecha(DD/MM/YYYY) + PV(4) + Nro(8) + TipoComp(2) + Monto(11)
    """
    # ── Mapeo de nombre de percepción → código de jurisdicción SIFERE ──
    CODIGOS_JURISDICCION = {
        "PERC.IB.CAP.FED.": "901",
        "PERC.IB.CABA C.ELECT": "901",
        "PERC.IB.BS.AS.": "902",
        "PER. IIBB CATAMARCA": "903",
        "PERC.IB.CORDOBA": "904",
        "PERC. CORRIENTES": "905",
        "PERC. IIBB CHACO": "906",
        "PERC IIBB CHUBUT": "907",
        "PERCEP IB ENTRE RIOS": "908",
        "PERC. IIBB FORMOSA": "909",
        "PERC.IIBB JUJUY": "910",
        "PERC.LA PAMPA": "911",
        "PERC.IB.LA RIOJA": "912",
        "PERC.IB.MENDOZA": "913",
        "PERC.IB MISIONES": "914",
        "Perc.IIBB Neuquen": "915",
        "PERC. IB RIO NEGRO": "916",
        "PERC.IB.SALTA": "917",
        "PERC.IB SAN JUAN": "918",
        "PERC. SAN LUIS": "919",
        "PERCEP IIBB STA CRUZ": "920",
        "PERC IIBB SANTA FE": "921",
        "PERC IIBB SGO ESTERO": "922",
        "PERC. TIERRA D.FUEGO": "923",
        "PERCEP IIBB TUCUMAN": "924",
    }

    # ── Mapeo de tipo de comprobante para SIFERE ──
    TIPO_COMP_SIFERE = {
        "FC": "FA",
        "ND": "DA",
        "NC": "CA",
        "TF": "FA",
        "TK": "FA",
    }

    # ── Tasas IVA (para excluirlas de percepciones) ──
    IVA_RATES = {
        'Tasa 21%', 'T.21%', 'C.F.21%', 'Tasa 27%', 'T.27%',
        'Tasa 10.5%', 'Tasa 10,5%', 'T.10.5%', 'T.10,5%',
        'C.F.10.5%', 'C.F.10,5%', 'Tasa 5%', 'T.5%',
        'Tasa 2.5%', 'Tasa 2,5%', 'T.2.5%', 'T.2,5%',
        'T.IMP 21%', 'T.IMP 10%', 'Exento',
        'R.Monot21', 'R.Mont.10',
    }

    # ── Extraer periodo (mes/año) del meta ──
    periodo_str = meta.get('periodo', '')
    # El periodo viene como "Desde el 01/MM/YYYY hasta el DD/MM/YYYY"
    p_match = re.search(r'(\d{2})/(\d{4})', periodo_str)
    if p_match:
        mes_periodo = p_match.group(1)
        anio_periodo = p_match.group(2)
    else:
        # Fallback: intentar extraer de otra forma
        nums = re.findall(r'\d+', periodo_str)
        if len(nums) >= 5:
            # Formato DD/MM/YYYY → posiciones 1=mes, 2=año
            mes_periodo = nums[1]
            anio_periodo = nums[2]
        else:
            mes_periodo = "01"
            anio_periodo = "2025"

    # ── Recopilar percepciones IIBB de cada transacción ──
    lineas_txt = []

    for t in transacciones:
        # Datos base de la transacción
        dia = t['Fecha']
        tipo = t['Tipo']
        numero_raw = t['Numero']
        cuit_raw = t['CUIT'] if t['CUIT'] else ''
        # Formatear CUIT con guiones: XX-XXXXXXXX-X (13 chars)
        if '-' in cuit_raw:
            cuit_formateado = cuit_raw
        else:
            cuit_limpio = cuit_raw.replace('-', '')
            if len(cuit_limpio) == 11:
                cuit_formateado = f"{cuit_limpio[:2]}-{cuit_limpio[2:10]}-{cuit_limpio[10]}"
            else:
                cuit_formateado = cuit_limpio

        # Separar PV y Nro del número de comprobante
        if '-' in numero_raw:
            pv_str = numero_raw.split('-')[0]
            resto_num = numero_raw.split('-')[1]
        else:
            pv_str = numero_raw[:5]
            resto_num = numero_raw[5:]

        # Quitar letra del final si existe
        letra = resto_num[-1] if resto_num and resto_num[-1].isalpha() else ''
        nro_str = resto_num[:-1] if letra else resto_num

        # Formatear fecha completa
        fecha_completa = f"{int(dia):02d}/{mes_periodo}/{anio_periodo}"

        # Formatear PV y Nro
        pv_formateado = pv_str[-4:].zfill(4)
        nro_formateado = nro_str.zfill(8)

        # Tipo de comprobante SIFERE
        tipo_sifere = TIPO_COMP_SIFERE.get(tipo, tipo)

        # Comprobante = PV + Nro + Tipo
        comprobante_sifere = f"{pv_formateado}{nro_formateado}{tipo_sifere}"

        # ── Recopilar percepciones de esta transacción ──
        percepciones = {}  # nombre_percepcion -> monto

        # Desde la tasa principal (si no es IVA)
        tasa = t['Tasa']
        if tasa and tasa not in IVA_RATES:
            nombre_upper = tasa.upper()
            if "PERC" in nombre_upper and "ADUA" not in nombre_upper and \
               "I.V.A" not in nombre_upper and "GCIAS" not in nombre_upper and \
               "IVA" not in nombre_upper:
                percepciones[tasa] = percepciones.get(tasa, 0.0) + t['Neto']

        # Desde sub-conceptos
        for s in t['SubConceptos']:
            nombre = s['Concepto']
            if not nombre or nombre in IVA_RATES:
                continue
            nombre_upper = nombre.upper()
            if "PERC" in nombre_upper and "ADUA" not in nombre_upper and \
               "I.V.A" not in nombre_upper and "GCIAS" not in nombre_upper and \
               "IVA" not in nombre_upper:
                monto = s['Neto'] if s['Neto'] != 0.0 else s['Percepcion']
                percepciones[nombre] = percepciones.get(nombre, 0.0) + monto

        # ── Generar líneas TXT para cada percepción ──
        for nombre_perc, monto in percepciones.items():
            if monto == 0.0:
                continue

            # Buscar código de jurisdicción
            codigo = CODIGOS_JURISDICCION.get(nombre_perc, None)
            if codigo is None:
                # Intento fuzzy: buscar por contenido parcial
                for key, val in CODIGOS_JURISDICCION.items():
                    if key.upper() in nombre_perc.upper() or nombre_perc.upper() in key.upper():
                        codigo = val
                        break
                if codigo is None:
                    continue  # No se encontró jurisdicción, saltar

            # Invertir signo para NC (el extractor ya invierte, pero el formato
            # SIFERE espera el monto con signo negativo explícito para CA)
            # En nuestro extractor, NC ya tienen montos negativos en los SubConceptos NO,
            # la inversión se hace en crear_excel. Aquí trabajamos con datos crudos.
            monto_final = monto
            es_nc = (tipo == 'NC')

            # Formatear monto
            valor_abs = abs(monto_final)
            parte_entera = int(valor_abs)
            parte_decimal = f"{valor_abs:.2f}".split('.')[1]

            if es_nc:
                monto_formateado = f"-{parte_entera:07d},{parte_decimal}"
            else:
                monto_formateado = f"{parte_entera:08d},{parte_decimal}"

            # Construir línea
            linea = (
                f"{codigo}"
                f"{cuit_formateado}"
                f"{fecha_completa}"
                f"{comprobante_sifere}"
                f"{monto_formateado}"
            )
            lineas_txt.append(linea)

    return "\n".join(lineas_txt)


def generar_sifere_retenciones_txt(transacciones: list[dict], meta: dict) -> str:
    """Genera un archivo TXT con formato SIFERE Formato Nº 1 para retenciones de IIBB.
    Cada línea (79 chars): CodJurisdiccion(3) + CUIT(13) + Fecha(10) + Sucursal(4)
    + NroConstancia(16) + TipoComp(1) + LetraComp(1) + NroCompOriginal(20) + Importe(11)
    """
    # ── Mapeo provincia → código de jurisdicción (reutiliza el de percepciones) ──
    # Palabras clave de provincia extraídas de los nombres de retención
    PROVINCIA_A_JURISDICCION = {
        "CAP.FED": "901", "CABA": "901", "C.A.B.A": "901",
        "BS.AS": "902", "BSAS": "902", "BS AS": "902", "BUENOS AIRES": "902",
        "CATAMARCA": "903",
        "CORDOBA": "904", "CÓRDOBA": "904",
        "CORRIENTES": "905",
        "CHACO": "906",
        "CHUBUT": "907",
        "ENTRE RIOS": "908", "ENTRE RÍOS": "908",
        "FORMOSA": "909",
        "JUJUY": "910",
        "LA PAMPA": "911", "PAMPA": "911",
        "LA RIOJA": "912", "RIOJA": "912",
        "MENDOZA": "913",
        "MISIONES": "914",
        "NEUQUEN": "915", "NEUQUÉN": "915",
        "RIO NEGRO": "916", "RÍO NEGRO": "916", "R.NEGRO": "916",
        "SALTA": "917",
        "SAN JUAN": "918",
        "SAN LUIS": "919",
        "STA CRUZ": "920", "SANTA CRUZ": "920",
        "SANTA FE": "921",
        "SGO ESTERO": "922", "SGO.ESTERO": "922", "SANTIAGO": "922",
        "TIERRA D.FUEGO": "923", "TIERRA DEL FUEGO": "923",
        "TUCUMAN": "924", "TUCUMÁN": "924",
    }

    # ── Mapeo tipo comprobante del sistema → tipo SIFERE retenciones (1 char) ──
    TIPO_COMP_RET = {
        "FC": "F", "TF": "F", "TK": "F",
        "NC": "C",
        "ND": "D",
    }

    # ── Palabras clave a EXCLUIR de retenciones ──
    EXCLUIR = {"SIRCREB", "SIRTAC", "BCO", "GCIAS", "IVA", "I.V.A", "BANCO", "BANCAR"}

    # ── Tasas IVA (para excluirlas) ──
    IVA_RATES = {
        'Tasa 21%', 'T.21%', 'C.F.21%', 'Tasa 27%', 'T.27%',
        'Tasa 10.5%', 'Tasa 10,5%', 'T.10.5%', 'T.10,5%',
        'C.F.10.5%', 'C.F.10,5%', 'Tasa 5%', 'T.5%',
        'Tasa 2.5%', 'Tasa 2,5%', 'T.2.5%', 'T.2,5%',
        'T.IMP 21%', 'T.IMP 10%', 'Exento',
        'R.Monot21', 'R.Mont.10',
    }

    def _buscar_jurisdiccion(nombre_ret: str) -> str | None:
        """Busca el código de jurisdicción extrayendo la provincia del nombre."""
        nombre_upper = nombre_ret.upper()
        for provincia, codigo in PROVINCIA_A_JURISDICCION.items():
            if provincia in nombre_upper:
                return codigo
        return None

    def _es_retencion_iibb(nombre: str) -> bool:
        """Retorna True si el concepto es una retención IIBB (no bancaria/gcias/iva)."""
        nombre_upper = nombre.upper()
        if "RET" not in nombre_upper:
            return False
        for excl in EXCLUIR:
            if excl in nombre_upper:
                return False
        return True

    # ── Extraer periodo (mes/año) del meta ──
    periodo_str = meta.get('periodo', '')
    p_match = re.search(r'(\d{2})/(\d{4})', periodo_str)
    if p_match:
        mes_periodo = p_match.group(1)
        anio_periodo = p_match.group(2)
    else:
        nums = re.findall(r'\d+', periodo_str)
        if len(nums) >= 5:
            mes_periodo = nums[1]
            anio_periodo = nums[2]
        else:
            mes_periodo = "01"
            anio_periodo = "2025"

    # ── Procesar transacciones ──
    lineas_txt = []

    for t in transacciones:
        dia = t['Fecha']
        tipo = t['Tipo']
        numero_raw = t['Numero']
        cuit_raw = t['CUIT'] if t['CUIT'] else ''
        letra = t.get('Letra', '')

        # CUIT del agente (proveedor) con guiones, 13 chars
        # Si ya tiene guiones, usar directo; si no, formatear XX-XXXXXXXX-X
        if '-' in cuit_raw:
            cuit_formateado = cuit_raw
        else:
            cuit_limpio = cuit_raw.replace('-', '')
            if len(cuit_limpio) == 11:
                cuit_formateado = f"{cuit_limpio[:2]}-{cuit_limpio[2:10]}-{cuit_limpio[10]}"
            else:
                cuit_formateado = cuit_limpio
        # Asegurar 13 chars
        cuit_formateado = cuit_formateado[:13].ljust(13)

        # Separar PV y Nro del número de comprobante
        if '-' in numero_raw:
            pv_str = numero_raw.split('-')[0]
            resto_num = numero_raw.split('-')[1]
        else:
            pv_str = numero_raw[:5]
            resto_num = numero_raw[5:]

        # Quitar letra del final si existe en el número
        if resto_num and resto_num[-1].isalpha():
            letra_comp = resto_num[-1]
            nro_str = resto_num[:-1]
        else:
            letra_comp = letra if letra else 'A'
            nro_str = resto_num

        # Fecha dd/mm/yyyy
        fecha_completa = f"{int(dia):02d}/{mes_periodo}/{anio_periodo}"

        # Sucursal (PV, 4 dígitos, ceros a izquierda) — default 1 si no tiene
        sucursal = pv_str.strip().lstrip('0') or "1"
        sucursal = sucursal[-4:].zfill(4)

        # Nro. Constancia (16 dígitos, ceros a izquierda) = Nro comprobante
        nro_constancia = nro_str.zfill(16)

        # Tipo de comprobante SIFERE retención (1 char) — siempre "O" (Otros)
        tipo_sifere = "O"

        # Letra del comprobante (1 char) — espacio en blanco para retenciones
        letra_sifere = " "

        # Nro. Comprobante Original (20 chars, ceros a izquierda) = mismo nro repetido
        nro_comp_original = nro_str.zfill(20)

        # ── Recopilar retenciones IIBB de esta transacción ──
        retenciones = {}  # nombre_retencion -> monto

        # Desde la tasa principal
        tasa = t['Tasa']
        if tasa and tasa not in IVA_RATES and _es_retencion_iibb(tasa):
            retenciones[tasa] = retenciones.get(tasa, 0.0) + t['Neto']

        # Desde sub-conceptos
        for s in t['SubConceptos']:
            nombre = s['Concepto']
            if not nombre or nombre in IVA_RATES:
                continue
            if _es_retencion_iibb(nombre):
                monto = s['Neto'] if s['Neto'] != 0.0 else s['Percepcion']
                retenciones[nombre] = retenciones.get(nombre, 0.0) + monto

        # ── Generar líneas TXT para cada retención ──
        for nombre_ret, monto in retenciones.items():
            if monto == 0.0:
                continue

            codigo = _buscar_jurisdiccion(nombre_ret)
            if codigo is None:
                continue  # No se encontró jurisdicción

            # Montos negativos solo para NC (tipo C o H)
            monto_final = monto
            es_nc = (tipo in ('NC',))

            valor_abs = abs(monto_final)
            parte_entera = int(valor_abs)
            parte_decimal = f"{valor_abs:.2f}".split('.')[1]

            if es_nc:
                monto_formateado = f"-{parte_entera:07d},{parte_decimal}"
            else:
                monto_formateado = f"{parte_entera:08d},{parte_decimal}"

            # Construir línea Formato 1 (79 chars)
            linea = (
                f"{codigo}"                # pos 1-3:   Jurisdicción (3)
                f"{cuit_formateado}"        # pos 4-16:  CUIT agente (13)
                f"{fecha_completa}"         # pos 17-26: Fecha (10)
                f"{sucursal}"              # pos 27-30: Sucursal (4)
                f"{nro_constancia}"        # pos 31-46: Nro Constancia (16)
                f"{tipo_sifere}"           # pos 47:    Tipo Comprobante (1)
                f"{letra_sifere}"          # pos 48:    Letra Comprobante (1)
                f"{nro_comp_original}"     # pos 49-68: Nro Comp Original (20)
                f"{monto_formateado}"      # pos 69-79: Importe (11)
            )
            lineas_txt.append(linea)

    return "\n".join(lineas_txt)


def generar_percepciones_arba_txt(transacciones: list[dict], meta: dict) -> str:
    """Genera un archivo TXT con formato ARBA para percepciones IIBB de ventas.
    Cada línea (81 chars): CUIT(13) + Fecha(10) + TipoComp(1) + Letra(1) + PV(5)
    + NroComp(8) + BaseImponible(14) + Alicuota(5) + ImportePerc(13) + Fecha(10) + LetraFija(1)
    """
    # ── Mapeo tipo comprobante del sistema → código ARBA (1 char) ──
    TIPO_COMP_ARBA = {
        "FC": "F", "TF": "F", "TK": "F",
        "NC": "C",
        "ND": "D",
        "RC": "R",
    }

    # ── Tasas IVA (para excluirlas al buscar percepciones) ──
    IVA_RATES = {
        'Tasa 21%', 'T.21%', 'C.F.21%', 'Tasa 27%', 'T.27%',
        'Tasa 10.5%', 'Tasa 10,5%', 'T.10.5%', 'T.10,5%',
        'C.F.10.5%', 'C.F.10,5%', 'Tasa 5%', 'T.5%',
        'Tasa 2.5%', 'Tasa 2,5%', 'T.2.5%', 'T.2,5%',
        'T.IMP 21%', 'T.IMP 10%', 'Exento',
        'R.Monot21', 'R.Mont.10',
    }

    # ── Palabras clave para identificar percepción IIBB Buenos Aires ──
    KEYWORDS_BS_AS = ["BS.AS", "BSAS", "BS AS", "BUENOS AIRES"]

    def _es_percepcion_bs_as(nombre: str) -> bool:
        """Retorna True si el concepto es una percepción IIBB Buenos Aires."""
        nombre_upper = nombre.upper()
        if "PERC" not in nombre_upper:
            return False
        # Excluir aduanera, IVA, ganancias
        if any(x in nombre_upper for x in ("ADUA", "I.V.A", "GCIAS", "IVA")):
            return False
        return any(kw in nombre_upper for kw in KEYWORDS_BS_AS)

    # ── Extraer periodo (mes/año) del meta ──
    periodo_str = meta.get('periodo', '')
    p_match = re.search(r'(\d{2})/(\d{4})', periodo_str)
    if p_match:
        mes_periodo = p_match.group(1)
        anio_periodo = p_match.group(2)
    else:
        nums = re.findall(r'\d+', periodo_str)
        if len(nums) >= 5:
            mes_periodo = nums[1]
            anio_periodo = nums[2]
        else:
            mes_periodo = "01"
            anio_periodo = "2025"

    # ── Procesar transacciones ──
    lineas_txt = []

    for t in transacciones:
        dia = t['Fecha']
        tipo = t['Tipo']
        numero_raw = t['Numero']
        cuit_raw = t['CUIT'] if t['CUIT'] else ''

        # ── CUIT con guiones (13 chars: XX-XXXXXXXX-X) ──
        if '-' in cuit_raw:
            cuit_formateado = cuit_raw
        else:
            cuit_limpio = cuit_raw.replace('-', '')
            if len(cuit_limpio) == 11:
                cuit_formateado = f"{cuit_limpio[:2]}-{cuit_limpio[2:10]}-{cuit_limpio[10]}"
            else:
                cuit_formateado = cuit_limpio
        cuit_formateado = cuit_formateado[:13].ljust(13)

        # ── Fecha completa DD/MM/YYYY ──
        fecha_completa = f"{int(dia):02d}/{mes_periodo}/{anio_periodo}"

        # ── Tipo comprobante ARBA (1 char) ──
        tipo_arba = TIPO_COMP_ARBA.get(tipo, tipo[0] if tipo else "F")

        # ── Separar PV y Nro del número de comprobante ──
        if '-' in numero_raw:
            pv_str = numero_raw.split('-')[0]
            resto_num = numero_raw.split('-')[1]
        else:
            pv_str = numero_raw[:5]
            resto_num = numero_raw[5:]

        # Quitar letra del final si existe → esa es la letra del comprobante
        if resto_num and resto_num[-1].isalpha():
            letra_comp = resto_num[-1]
            nro_str = resto_num[:-1]
        else:
            letra_comp = 'A'
            nro_str = resto_num

        pv_formateado = pv_str[-5:].zfill(5)
        nro_formateado = nro_str.zfill(8)

        # ── Buscar percepción IIBB BS.AS. en esta transacción ──
        monto_percepcion = 0.0

        # Desde la tasa principal
        tasa = t['Tasa']
        if tasa and tasa not in IVA_RATES and _es_percepcion_bs_as(tasa):
            monto_percepcion += t['Neto']

        # Desde sub-conceptos
        for s in t['SubConceptos']:
            nombre = s['Concepto']
            if not nombre or nombre in IVA_RATES:
                continue
            if _es_percepcion_bs_as(nombre):
                monto = s['Neto'] if s['Neto'] != 0.0 else s['Percepcion']
                monto_percepcion += monto

        # Si no hay percepción BS.AS., saltar esta transacción
        if monto_percepcion == 0.0:
            continue

        # ── Base imponible = Neto gravado del movimiento ──
        # Recopilar neto de todas las tasas IVA (excluyendo percepciones/retenciones)
        base_imponible = 0.0

        # Neto de la tasa principal (si es IVA)
        if tasa and tasa in IVA_RATES:
            base_imponible += t['Neto']

        # Neto de sub-conceptos que son tasas IVA
        for s in t['SubConceptos']:
            nombre = s['Concepto']
            if nombre and nombre in IVA_RATES:
                base_imponible += s['Neto']

        # Si la base es 0, intentar usar el Neto principal
        if base_imponible == 0.0:
            base_imponible = t['Neto']

        # ── Calcular alícuota = Percepción / Base * 100 ──
        if base_imponible != 0.0:
            alicuota = abs(monto_percepcion) / abs(base_imponible) * 100
        else:
            alicuota = 0.0

        # ── Determinar si es NC (montos negativos) ──
        es_nc = (tipo == 'NC')

        # ── Formatear Base Imponible (14 chars: 11 enteros + . + 2 decimales) ──
        base_abs = abs(base_imponible)
        if es_nc:
            # Signo negativo reemplaza un cero de relleno
            base_str = f"-{int(base_abs):010d}.{base_abs:.2f}".split('.')
            base_formateada = f"-{int(base_abs):010d}.{f'{base_abs:.2f}'.split('.')[1]}"
        else:
            base_formateada = f"{int(base_abs):011d}.{f'{base_abs:.2f}'.split('.')[1]}"

        # ── Formatear Alícuota (5 chars: 2 enteros + . + 2 decimales) ──
        alic_formateada = f"{int(alicuota):02d}.{f'{alicuota:.2f}'.split('.')[1]}"

        # ── Formatear Importe Percepción (13 chars: 10 enteros + . + 2 decimales) ──
        perc_abs = abs(monto_percepcion)
        if es_nc:
            perc_formateada = f"-{int(perc_abs):09d}.{f'{perc_abs:.2f}'.split('.')[1]}"
        else:
            perc_formateada = f"{int(perc_abs):010d}.{f'{perc_abs:.2f}'.split('.')[1]}"

        # ── Construir línea (81 chars) ──
        linea = (
            f"{cuit_formateado}"          # pos 1-13:  CUIT (13)
            f"{fecha_completa}"           # pos 14-23: Fecha (10)
            f"{tipo_arba}"                # pos 24:    Tipo comprobante (1)
            f"{letra_comp}"               # pos 25:    Letra comprobante (1)
            f"{pv_formateado}"            # pos 26-30: Punto de venta (5)
            f"{nro_formateado}"           # pos 31-38: Nro comprobante (8)
            f"{base_formateada}"          # pos 39-52: Base imponible (14)
            f"{alic_formateada}"          # pos 53-57: Alícuota (5)
            f"{perc_formateada}"          # pos 58-70: Importe percepción (13)
            f"{fecha_completa}"           # pos 71-80: Fecha repetida (10)
            f"A"                          # pos 81:    Letra fija (1)
        )
        lineas_txt.append(linea)

    return "\n".join(lineas_txt)


def seleccionar_archivo() -> Path:
    """Abre un diálogo para que el usuario seleccione un archivo .txt."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()  # Ocultar ventana principal
    root.attributes('-topmost', True)  # Traer diálogo al frente

    archivo = filedialog.askopenfilename(
        title='Seleccionar archivo de movimientos',
        filetypes=[('Archivos de texto', '*.txt'), ('Todos los archivos', '*.*')],
        initialdir=Path(__file__).parent
    )

    root.destroy()

    if not archivo:
        print("❌ No se seleccionó ningún archivo. Saliendo...")
        sys.exit(0)

    return Path(archivo)


def main():
    # Forzar UTF-8 en la consola de Windows (solo cuando se corre como script)
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

    if len(sys.argv) < 2:
        input_file = seleccionar_archivo()
    else:
        input_file = Path(sys.argv[1])

    if not input_file.exists():
        print(f"❌ No se encontró el archivo: {input_file}")
        sys.exit(1)

    output_file = input_file.with_suffix('.xlsx')

    print(f"📖 Leyendo: {input_file}")
    transacciones, meta = parsear_archivo(path=input_file)

    if not transacciones:
        print("❌ No se encontraron transacciones en el archivo.")
        sys.exit(1)

    crear_excel(transacciones, meta, output_file)
    print("✅ Proceso completado exitosamente.")


if __name__ == '__main__':
    main()
