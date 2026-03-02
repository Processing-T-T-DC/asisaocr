import csv
import re
import os
import glob
import sys
import pandas as pd
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8")# type: ignore

# --- CONSTANTES DE COLUMNAS (Globales para usarlas en ambas funciones) ---
COLUMNAS_ORDENADAS_PIA = [
    "Nombre del Archivo", # Agregamos esta para saber de qué HTML viene la fila
    "Introducción", "Título - Nombre y tratamiento", "Identificación y descripción de la finalidad",
    "1. Contexto - descripción del proyecto y análisis de la necesidad",
    "1.1.1.Descripción del proyecto, sistema o producto sobre el que realizar una evaluación de impacto",
    "Descripción de la finalidad y objetivos a alcanzar al realizar un PIA",
    "1.1.2. Análisis de la necesidad del PIA - Situaciones que aconsejan realizar una mnueva evaluación de impacto o actualizar la ya realizada para este tratamiento de datos",
    "¿Se encuentra la presente actividad de tratamiento en los supuestos indicados como de \"Riesgo alto\" para los derechos y libertades de las personas o existen otros motivos que justifiquen elaborar un PIA",
    "Detalle situación que aconseja realizar una nueva evaluación de impacto o actualizar la ya existente",
    "Términos y alcance de la Evaluación de Impacto para la Privacidad",
    "2.1. Preparar y realizar el PIA",
    "2.1.1.Flujos de información e implicaciones de uso. ¿Cómo se recopilan los datos?",
    "Orígenes de datos", "¿Cómo se tratan los datos? ", "Sistemas en los que tratarán los datos",
    "¿Existe una política de conservación y eliminación de datos?",
    "Procedimiento de notificación a las autoridades pertinentes",
    "Identificación de las implicaciones del uso de información por parte de los usuarios",
    "2.1.2. Idoneidad, necesidad y proporcionalidad de la finalidad - ¿El tratamiento permite y logra alcanzar la finalidad perseguida?",
    "Justificación_1", "Los datos recogidos se van a usar exclusivamente para la finalidad declarada y no para ninguna otra no informada ni incompatible con la legitimidad de su uso (principio de limitación de la finalidad)",
    "Justificación_2", "La finalidad que se pretende cubrir requiere de todos los datos a recabar y para todas las personas/interesados afectados (principio de minimización de datos)",
    "Justificación_3", "Las tecnologías empleadas son adecuadas para la finalidad establecida desde el punto de vista del cumplimiento de los principios fundamentales de la privacidad",
    "Justificación_4", "Los datos no se mantienen más tiempo del necesario para las finalidades del tratamiento (principio de limitación del plazo de conservación)",
    "Justificación_5", "Conclusión",
    "2.1.3. Requerimientos de Privacidad - Indique las legislaciones y regulaciones que afectan a la evaluación",
    "Indicar conjuntos de control de seguridad de la información",
    "Indicar requisitos de privacidad asociados a la evaluación son los siguientes",
    "Indique controles ya planeados o existentes que se espera cumplan con los requisitos de privacidad identificados en la evaluación, son los siguientes",
    "2.1.4. Equipo y criterios - Responsable EIPD", "Responsable de firma del informe PIA",
    "Otros intervinientes en el proceso de revisión del PIA", "Criterios usados para estimar la valoración de la probabilidad",
    "Criterios usados para estimar la valoración del impacto",
    "2.1.5 Plan y Recursos - Describa el plan de proyecto, o el procedimiento que realiza la organización para realizar un PIA",
    "Los recursos asignados al PIA",
    "2.1.6 Consulta a los interesados - Partes interesadas identificadas",
    "El plan de consulta establecido para la comunicación con las partes interesadas",
    "2.1.7 Comprensión de la evaluación - Descripción detallada del programa, proceso o sistema a evaluar",
    "3. Identificación de riesgos", "3.1. Operaciones relacionadas con los fines de tratamiento",
    "3.2. Tipos de datos utilizados", "3.3. Extensión y alcance del tratamiento",
    "3.4. Categorías de interesados", "3.5. Factores técnicos del tratamiento",
    "3.6. Recogida y generación de datos", "3.7. Efectos colaterales del tratamiento",
    "3.8. Categoría del responsable", "3.9. Comunicaciones de datos", "3.10. Seguridad de los tratamientos"
]

COLUMNAS_ORDENADAS_AARR = [
    "Nombre del Archivo", # Identificador de origen
    "Tratamiento / Finalidad",
    "Evaluación objetiva",
    "Descripción",
    "¿Figuras implicadas en la EIPD?",
    "Notificar a",
    "Nº sujetos afectados",
    "Categoría de datos tratados",
    "Orígenes de datos",
    "Duración del tratamiento",
    "Extensión geográfica",
    "Legitimación del tratamiento",
    "¿Se van a tratar datos personales para hacer ratings/scoring o toma de decisiones?",
    "Detalle/Justificación (Ratings/Scoring)",
    "¿El tratamiento de los datos implica una toma de decisiones automatizada sin que haya ninguna persona que intervenga en la decisión o valore los resultados?",
    "Detalle/Justificación (Decisiones automatizadas)",
    "¿La recogida de los datos tiene como finalidad la monitorización o evaluación sistemática y exhaustiva de aspectos personales?",
    "Detalle/Justificación (Monitorización)",
    "¿La recogida de los datos tiene como finalidad el tratamiento de datos especialmente protegidos?",
    "Detalle/Justificación (Datos protegidos)",
    "¿La recogida de los datos tiene como finalidad el tratamiento a gran escala?",
    "Detalle/Justificación (Gran escala)",
    "Para llevar a cabo este tratamiento, ¿se combinan conjuntos de datos utilizados por otros responsables de tratamiento cuya finalidad diste en exceso de las expectativas del interesado?",
    "Detalle/Justificación (Combinación datos)",
    "¿La finalidad del tratamiento implica el uso específico de datos de personas con discapacidad o cualquier otro colectivo en situación de especial vulnerabilidad?",
    "Detalle/Justificación (Vulnerabilidad)",
    "¿Se prevé el uso de tecnologías que se pueden percibir como inmaduras, de reciente creación o salida al mercado, cuyo alcance no puede ser previsto por el interesado de forma clara o razonable e implique elevado riesgo para el acceso no autorizado?",
    "Detalle/Justificación (Tecnologías inmaduras)",
    "¿El tratamiento involucra contacto con los interesados de manera que, dicho contacto, pueda resultar intrusivo o se prevé el uso de tecnologías que se pueden percibir como especialmente intrusivas en la privacidad?",
    "Detalle/Justificación (Contacto intrusivo)",
    "¿Se enriquece la información de los interesados mediante la recogida de nuevas categorías de datos o se usen las existentes con nuevas finalidades que antes no se contemplaban, en particular, si estas finalidades son más intrusivas o inesperadas para los afectados, o incluso pueda llegar a bloquear el disfrute de algún servicio?",
    "Detalle/justificación (Enriquecimiento info)",
    "¿El tratamiento implica que un elevado número de personas (más allá de las necesarias para llevar a cabo el mismo) tenga acceso a los datos personales tratados?",
    "Detalle/Justificación (Acceso elevado)",
    "¿Se van a tratar datos relativos a la observación de zonas de acceso público?",
    "Detalle/Justificación (Zonas públicas)",
    "¿Se utilizan datos de carácter personal no disociados o no anonimizados de forma irreversible con fines estadísticos, históricos o de investigación científica?",
    "Detalle/Justificación (Fines estadísticos)",
    "¿Puede el tratamiento impedir ejercer un derecho, utilizar un servicio o ejecutar un contrato?",
    "Detalle/Justificación (Impedir derechos)",
    "¿Se realizan cesiones de datos a otras entidades privadas u otras organizaciones, ya sean del mismo grupo o proveedores externos al mismo?",
    "Detalle de las cesiones realizadas",
    "¿Se realizan transferencias internacionales de datos a países fuera de la Unión Europea y que no cuenten con medidas de protección de datos de carácter personal similares a las establecidas por la Autoridad de Control?",
    "Detalle/Justificación (Transferencias internacionales)",
    "¿Es este tratamiento similar a otro para el que haya sido necesario realizar un EIPD?",
    "Detalle/Justificación (Similitud EIPD)",
    "¿Este tratamiento puede conllevar una pérdida o alteración de la información?",
    "Justificación (Pérdida/Alteración)",
    "¿Se utilizada documentación en papel para tratar datos personales?",
    "Detalle/Justificación (Papel)",
    "¿Interviene algún proveedor en el proceso?",
    "Justificación (Proveedores)",
    "Indique los sistemas en los que se tratarán los datos (Medios electrónicos (físicos o en la nube)/Papel)",
    "¿Existen fines secundarios/intermedios con el tratamiento de los datos?",
    "Especificar cuáles son los fines secundarios/intermedios",
    "¿Con qué frecuencia se recaban los datos?",
    "¿Se recaban todos los datos afectados de forma conjunta o en distintos momentos?",
    "Descripción del ciclo de vida",
    "Operaciones de tratamiento",
    "¿Se infieren u obtienen datos adicionales a partir del tratamiento de datos original?",
    "Detalle/justificación (Datos inferidos)",
    "¿Se prevén efectos colaterales o adversos para los interesados?",
    "Detalle/justificación (Efectos adversos)",
    "Resultado del análisis"
]

def procesar_un_html_pia(html_content, nombre_archivo):
    """
    Recibe el contenido HTML y devuelve un DICCIONARIO con los datos limpios.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Mapeo de variables (Indices basados en la lista global)
    # Nota: Usamos la lista global para mantener coherencia
    vars_cols = {i: nombre for i, nombre in enumerate(COLUMNAS_ORDENADAS_PIA)}
    
    # LISTA DE JUSTIFICACIONES (Indices en la lista global)
    # Si estoy en la columna 19 y veo "Justificación", salto a la 20.
    relacion_pregunta_justificacion = {
        19: 20, # Idoneidad... -> Justificación_1
        21: 22, # Limitación... -> Justificación_2
        23: 24, # Minimización... -> Justificación_3
        25: 26, # Tecnologías... -> Justificación_4
        27: 28  # Conservación... -> Justificación_5
    }

    # MAPEO DE PALABRAS CLAVE (HTML -> Indice en la lista global)
    mapeo_indices = {
        "introducción": 1,
        "identificación y descripción de la finalidad": 3,
        "1.1 contexto": 4,
        "producto sobre el que realizar": 5,
        "objetivos a alcanzar al realizar": 6,
        "aconsejan realizar una nueva evaluación": 7,
        # "riesgo alto para los derechos": 8,
        "se encuentra la presente actividad de tratamiento en los supuestos indicados": 8,
        "la situación que aconseja realizar una nueva evaluación": 9,
        "los términos y alcance aprobados para la evaluación de impacto ": 10,
        "2.1 preparar y realizar": 11,
        "datos son recopilados": 12,
        "orígenes de datos": 13,
        "¿cómo se tratarán los datos?": 14,
        "sistemas en los que se tratarán": 15,
        "existe una política de conservación y eliminación de datos": 16,
        "implicaciones del uso de información": 18,
        # PREGUNTAS QUE TIENEN JUSTIFICACIÓN ASOCIADA
        "alcanzar la finalidad perseguida": 19,
        "exclusivamente para la finalidad declarada": 21,
        "la finalidad que se pretende cubrir requiere de todos los datos a recabar": 23,
        "las tecnologías empleadas son adecuadas para la finalidad establecida": 25,
        "los datos no se mantienen más tiempo del necesario para las finalidades del tratamiento": 27,

        "conclusión": 29,
        "los requisitos de protección de privacidad legislativos y regulatorios": 30,
        "los conjuntos de control de seguridad de la información": 31,
        "asociados a la evaluación son los siguientes": 32,
        "controles ya planeados": 33,
        "intervinientes en el proceso de revisión del pia": 36,
        "valoración de la probabilidad": 37,
        "valoración del impacto": 38,
        # "2.1.5 plan y recursos": 39,
        "procedimiento que realiza la organización": 39,
        "recursos asignados al proyecto": 40,
        # "2.1.6 consulta a los interesados": 42,
        "partes interesadas identificadas": 41,
        "plan de consulta establecido": 42,
        # "2.1.7 comprensión de la evaluación": 43,
        "descripción detallada del programa, proceso o sistema a evaluar": 43,
        "3 identificación de riesgos": 44,
        "1. operaciones relacionadas": 45,
        "2. tipos de datos utilizados": 46,
        "3. extensión y alcance": 47,
        "4. categorías de interesados": 48,
        "5. factores técnicos": 49,
        "6. recogida y generación": 50,
        "7. efectos colaterales": 51,
        "8. categoría del responsable": 52,
        "9. comunicaciones de datos": 53,
        "10. seguridad de los tratamientos": 54
    }

    # Diccionario resultado inicializado con vacíos
    resultado = {col: [] for col in COLUMNAS_ORDENADAS_PIA}
    resultado["Nombre del Archivo"] = [nombre_archivo] # Guardamos el nombre
    
    # 1. EXTRACCIÓN PROYECTO - TRATAMIENTO
    proy_tag = soup.find(class_="report_proyecto_name")
    trat_tag = soup.find('span', string=re.compile(r'Tratamiento:', re.I))# type: ignore
    if proy_tag or trat_tag:
        p_txt = proy_tag.get_text(strip=True) if proy_tag else ""
        t_txt = trat_tag.next_sibling.strip() if (trat_tag and trat_tag.next_sibling) else ""
        resultado[vars_cols[2]].append(f"{p_txt} - {t_txt}")

    # 2. EXTRACCIÓN DE CONTENIDO
    # IMPORTANTE: Ahora trackeamos el INDICE (int), no solo el nombre
    curr_col_idx = 1

    print(f"\n{'='*80}")
    print(f"Procesando: {nombre_archivo}")
    print(f"{'='*80}")
    
    for el in soup.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'table', 'li', 'div']):
        if el.name == 'div' and el.find(['p', 'table']): continue 
        if el.find_parent('table') and el.name != 'table': continue
        
        texto_raw = el.get_text(" ", strip=True)
        # Limpieza ISO
        texto_limpio = re.sub(r'\[.*?\]', '', texto_raw).strip()
        
        if not texto_limpio: continue
        t_low = texto_limpio.lower()

        # --- LÓGICA DE DETECCIÓN --- #
        es_seccion = False
        nuevo_idx_detectado = None
        
        # A) ¿Es una Justificación?
        # Miramos dónde estamos (curr_col_idx) para saber a dónde ir
        if t_low.startswith("justificación"):
            if curr_col_idx in relacion_pregunta_justificacion:
                nuevo_idx_detectado = relacion_pregunta_justificacion[curr_col_idx]
                es_seccion = True
                # Cambiamos de columna y saltamos el título "Justificación"
                print(f"  🔄 Justificación: [{curr_col_idx}] → [{nuevo_idx_detectado}]")
                curr_col_idx = nuevo_idx_detectado
                continue
            # Si ya estamos en una columna de justificación (20, 22, 24, 26, 28),
            # NO cambiamos de columna, solo hacemos continue para saltar el título
            elif curr_col_idx in [20, 22, 24, 26, 28]:
                print(f"  🔄 Justificación adicional en [{curr_col_idx}] (se concatena)")
                continue
        
        # B) ¿Es otra sección del mapeo?
        if not es_seccion:
            for kw, idx in mapeo_indices.items():
                if kw in t_low:
                    nuevo_idx_detectado = idx
                    es_seccion = True
                    # print(f"  🎯 Keyword '{kw[:40]}...' → columna [{idx}]")
                    break

        # 3. ¿Es una sección del HTML que NO queremos en el Excel? (EL FRENO)
        # Si el elemento es una cabecera (H1-H4) o tiene clase 'pregunta', pero no entró en el mapeo anterior
        clases = el.get('class', [])# type: ignore
        es_titulo_visual = (el.name in ['h1', 'h2', 'h3', 'h4'] or 'pregunta' in clases)# type: ignore

        if not es_seccion and es_titulo_visual:
            # print(f"  ⛔ DESCARTE (no mapeado): '{texto_limpio[:60]}...'")
            curr_col_idx = -1
            continue

        if es_seccion and nuevo_idx_detectado is not None:
            curr_col_idx = nuevo_idx_detectado
            continue

        # --- GUARDAR CONTENIDO ---
        if curr_col_idx != -1: # Si no estamos en descarte
            curr_col_name = vars_cols[curr_col_idx] # Obtenemos nombre actual por índice

            val = ""
            if el.name == 'table':
                filas = []
                for r in el.find_all('tr'):
                    celdas = [c.get_text(strip=True) for c in r.find_all(['td', 'th'])]
                    if any(celdas): filas.append(f"[{' | '.join(celdas)}]")
                val = "\n".join(filas)
            else:
                prefix = "• " if el.name == 'li' else ""
                val = prefix + texto_raw

            # DEBUG: Mostrar cuando capturamos SI o NO
            if texto_limpio.upper() in ['SI', 'NO', 'SÍ']:
                print(f"  ✅ Capturando '{texto_limpio}' en [{curr_col_idx}] '{curr_col_name[:50]}...'")
            
            # Evitar duplicados y títulos
            if val.strip().lower() != curr_col_name.lower():
                if not resultado[curr_col_name] or resultado[curr_col_name][-1] != val:
                    resultado[curr_col_name].append(val)

    # 3. APLANAR RESULTADO (Lista -> String)
    fila_final = {}
    for col in COLUMNAS_ORDENADAS_PIA:
        # Renombrar Justificaciones para el CSV final
        fila_final[col] = "\n".join(resultado[col]).strip()
        
    return fila_final

def procesar_un_html_aarr(html_content, nombre_archivo):
    """
    Recibe el contenido HTML y devuelve un DICCIONARIO con los datos limpios.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Mapeo de variables (Indices basados en la lista global)
    # Nota: Usamos la lista global para mantener coherencia
    vars_cols = {i: nombre for i, nombre in enumerate(COLUMNAS_ORDENADAS_AARR)}
    
    # LISTA DE JUSTIFICACIONES (Indices en la lista global)
    # Si estoy en la columna 19 y veo "Justificación", salto a la 20.
    relacion_pregunta_justificacion = {
        19: 20, # Idoneidad... -> Justificación_1
        21: 22, # Limitación... -> Justificación_2
        23: 24, # Minimización... -> Justificación_3
        25: 26, # Tecnologías... -> Justificación_4
        27: 28  # Conservación... -> Justificación_5
    }

    # MAPEO DE PALABRAS CLAVE (HTML -> Indice en la lista global)
    mapeo_indices = {
        # Metadatos
        "nombre del archivo": 0,
        "inicio_documento": 1,
        
        # Identificación básica
        "evaluación objetiva": 2,
        "descripción": 3,
        "figuras implicadas": 4,
        "notificar a": 5,
        
        # Alcance y Datos
        "número de sujetos afectados": 6,
        "categoría de datos tratados": 7,
        "orígenes de datos": 8,
        "duración del tratamiento": 9,
        "extensión geográfica": 10,
        "legitimación del tratamiento": 11,

        # Análisis de Riesgos Específicos (Pregunta + Justificación)
        "ratings/scoring o para la toma de decisiones": 12,
        "justificación ratings/scoring": 13,
        
        "toma de decisiones automatizada": 14,
        "justificación decisiones automatizadas": 15,
        
        "finalidad la monitorización o evaluación sistemática y exhaustiva": 16,
        "justificación monitorización": 17,
        
        "tratamiento de datos especialmente protegidos": 18,
        "justificación datos protegidos": 19,
        
        "finalidad el tratamiento a gran escala": 20,
        "justificación gran escala": 21,
        
        "se combinan conjuntos de datos utilizados": 22,
        "justificación combinación datos": 23,
        
        "colectivo en situación de especial vulnerabilidad": 24,
        "justificación vulnerabilidad": 25,
        
        "tecnologías que se pueden percibir como inmaduras, de reciente creación o salida al mercado": 26,
        "detalle de las tecnologías que puedan ser percibidas como inmaduras utilizadas": 27,
        
        "tecnologías que se pueden percibir como especialmente intrusivas": 28,
        "justificación contacto intrusivo": 29,
        
        "enriquece la información de los interesados mediante la recogida": 30,
        "justificación enriquecimiento": 31,
        
        "implica que un elevado número de personas": 32,
        "justificación acceso elevado": 33,
        
        "observación de zonas de acceso público": 34,
        "justificación zonas públicas": 35,
        
        "personal no disociados o no anonimizados de forma irreversible con fines estadísticos": 36,
        "justificación fines estadísticos": 37,
        
        "ejercer un derecho, utilizar un servicio o ejecutar un contrato": 38,
        "justificación impedir derechos": 39,

        # Relaciones Externas y Otros
        "realizan cesiones de datos a otras entidades ": 40,
        "detalle de las cesiones realizadas": 41,
        
        "transferencias internacionales de datos a países fuera": 42,
        "detalle de las transferencias realizadas": 43,
        
        "similar a otro para el que haya sido necesario realizar un eipd": 44,
        "justificación de la percepción por parte del responsable de la actividad de tratamiento respecto de la similitud a otro tratamiento para el que haya sido necesario realizar un eipd": 45,
        
        "conllevar una pérdida o alteración de la información": 46,
        "justificación de la percepción por parte del responsable de la actividad de tratamiento respecto de la posibilidad de pérdida o alteración de la información": 47,
        
        "utilizada documentación en papel para tratar datos personales": 48,
        "justificación por parte del responsable de la actividad de tratamiento de las medidas aplicadas a la documentación en papel": 49,
        
        "interviene algún proveedor en el proceso": 50,
        "justificación de los proveedores que intervienen en el proceso": 51,

        # Tratamiento de los datos
        "indique los sistemas en los que se tratarán los datos": 52,
        "existen fines secundarios/intermedios con el tratamiento de los datos": 53,
        "especificar cuáles son los fines secundarios/intermedios": 54,
        "con qué frecuencia se recaban los datos": 55,
        "se recaban todos los datos afectados de forma conjunta o en distintos momentos": 56,
        "descripción del ciclo de vida": 57,

        # Operaciones y Conclusión
        "operaciones de tratamiento": 58,

        "infieren u obtienen datos adicionales a partir del tratamiento": 59,
        "justificación datos inferidos": 60,
        
        "prevén efectos colaterales o adversos para los interesados": 61,
        "justificación efectos adversos": 62,
        
        # Resultado del análisis
        # "resultado del análisis": 56
        "en base a las respuestas realizadas en el cuestionario de evaluación objetiva, debe calificarse la presente actividad de tratamiento como": 63,
        "ha indicado que la presente actividad presenta riesgo escaso o nulo": 64,
        "evidencias relativas al resultado del análisis: con riesgo escaso o nulo": 65,
        "ha indicado que la presente actividad presenta riesgo alto": 66,
        "evidencias relativas al resultado del análisis: riesgo alto": 67,
        "otra información relevante de procedimientos anteriores": 68,
        "evidencias de otro tipo de información relevante adicional": 69
    }

    # Diccionario resultado inicializado con vacíos
    resultado = {col: [] for col in COLUMNAS_ORDENADAS_AARR}
    resultado["Nombre del Archivo"] = [nombre_archivo] # Guardamos el nombre
    
    # 1. EXTRACCIÓN PROYECTO - TRATAMIENTO
    proy_tag = soup.find(class_="report_proyecto_name")
    trat_tag = soup.find('span', string=re.compile(r'Tratamiento:', re.I))# type: ignore
    if proy_tag or trat_tag:
        p_txt = proy_tag.get_text(strip=True) if proy_tag else ""
        t_txt = trat_tag.next_sibling.strip() if (trat_tag and trat_tag.next_sibling) else ""
        resultado[vars_cols[2]].append(f"{p_txt} - {t_txt}")

    # 2. EXTRACCIÓN DE CONTENIDO
    # IMPORTANTE: Ahora trackeamos el INDICE (int), no solo el nombre
    curr_col_idx = 1

    print(f"\n{'='*80}")
    print(f"Procesando: {nombre_archivo}")
    print(f"{'='*80}")
    
    for el in soup.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'table', 'li', 'div']):
        if el.name == 'div' and el.find(['p', 'table']): continue 
        if el.find_parent('table') and el.name != 'table': continue
        
        texto_raw = el.get_text(" ", strip=True)
        # Limpieza ISO
        texto_limpio = re.sub(r'\[.*?\]', '', texto_raw).strip()
        
        if not texto_limpio: continue
        t_low = texto_limpio.lower()

        # --- LÓGICA DE DETECCIÓN --- #
        es_seccion = False
        nuevo_idx_detectado = None
        
        # A) ¿Es una Justificación?
        # Miramos dónde estamos (curr_col_idx) para saber a dónde ir
        if t_low.startswith("justificación"):
            if curr_col_idx in relacion_pregunta_justificacion:
                nuevo_idx_detectado = relacion_pregunta_justificacion[curr_col_idx]
                es_seccion = True
                # Cambiamos de columna y saltamos el título "Justificación"
                print(f"  🔄 Justificación: [{curr_col_idx}] → [{nuevo_idx_detectado}]")
                curr_col_idx = nuevo_idx_detectado
                continue
            # Si ya estamos en una columna de justificación (20, 22, 24, 26, 28),
            # NO cambiamos de columna, solo hacemos continue para saltar el título
            elif curr_col_idx in [20, 22, 24, 26, 28]:
                print(f"  🔄 Justificación adicional en [{curr_col_idx}] (se concatena)")
                continue
        
        # B) ¿Es otra sección del mapeo?
        if not es_seccion:
            for kw, idx in mapeo_indices.items():
                if kw in t_low:
                    nuevo_idx_detectado = idx
                    es_seccion = True
                    # print(f"  🎯 Keyword '{kw[:40]}...' → columna [{idx}]")
                    break

        # 3. ¿Es una sección del HTML que NO queremos en el Excel? (EL FRENO)
        # Si el elemento es una cabecera (H1-H4) o tiene clase 'pregunta', pero no entró en el mapeo anterior
        clases = el.get('class', [])# type: ignore
        es_titulo_visual = (el.name in ['h1', 'h2', 'h3', 'h4'] or 'pregunta' in clases)# type: ignore

        if not es_seccion and es_titulo_visual:
            # print(f"  ⛔ DESCARTE (no mapeado): '{texto_limpio[:60]}...'")
            curr_col_idx = -1
            continue

        if es_seccion and nuevo_idx_detectado is not None:
            curr_col_idx = nuevo_idx_detectado
            continue

        # --- GUARDAR CONTENIDO ---
        if curr_col_idx != -1: # Si no estamos en descarte
            curr_col_name = vars_cols[curr_col_idx] # Obtenemos nombre actual por índice

            val = ""
            if el.name == 'table':
                filas = []
                for r in el.find_all('tr'):
                    celdas = [c.get_text(strip=True) for c in r.find_all(['td', 'th'])]
                    if any(celdas): filas.append(f"[{' | '.join(celdas)}]")
                val = "\n".join(filas)
            else:
                prefix = "• " if el.name == 'li' else ""
                val = prefix + texto_raw

            # DEBUG: Mostrar cuando capturamos SI o NO
            if texto_limpio.upper() in ['SI', 'NO', 'SÍ']:
                print(f"  ✅ Capturando '{texto_limpio}' en [{curr_col_idx}] '{curr_col_name[:50]}...'")
            
            # Evitar duplicados y títulos
            if val.strip().lower() != curr_col_name.lower():
                if not resultado[curr_col_name] or resultado[curr_col_name][-1] != val:
                    resultado[curr_col_name].append(val)

    # 3. APLANAR RESULTADO (Lista -> String)
    fila_final = {}
    for col in COLUMNAS_ORDENADAS_AARR:
        # Renombrar Justificaciones para el CSV final
        fila_final[col] = "\n".join(resultado[col]).strip()
        
    return fila_final

def procesar_un_html_estructurado(html_content, nombre_archivo):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Mapeo de variables (Indices basados en la lista global)
    vars_cols = {i: nombre for i, nombre in enumerate(COLUMNAS_ORDENADAS_AARR)}
    
    # Tu lógica de relación (Pregunta -> Su Justificación)
    relacion_pregunta_justificacion = {
        19: 20, 21: 22, 23: 24, 25: 26, 27: 28 
    }

    # El mapeo_indices que ya tienes definido
    mapeo_indices = {
        # Metadatos
        "nombre del archivo": 0,
        "inicio_documento": 1,
        
        # Identificación básica
        "evaluación objetiva": 2,
        "descripción": 3,
        "figuras implicadas": 4,
        "notificar a": 5,
        
        # Alcance y Datos
        "número de sujetos afectados": 6,
        "categoría de datos tratados": 7,
        "orígenes de datos": 8,
        "duración del tratamiento": 9,
        "extensión geográfica": 10,
        "legitimación del tratamiento": 11,

        # Análisis de Riesgos Específicos (Pregunta + Justificación)
        "ratings/scoring o para la toma de decisiones": 12,
        "justificación ratings/scoring": 13,
        
        "toma de decisiones automatizada": 14,
        "justificación decisiones automatizadas": 15,
        
        "finalidad la monitorización o evaluación sistemática y exhaustiva": 16,
        "justificación monitorización": 17,
        
        "tratamiento de datos especialmente protegidos": 18,
        "justificación datos protegidos": 19,
        
        "finalidad el tratamiento a gran escala": 20,
        "justificación gran escala": 21,
        
        "se combinan conjuntos de datos utilizados": 22,
        "justificación combinación datos": 23,
        
        "colectivo en situación de especial vulnerabilidad": 24,
        "justificación vulnerabilidad": 25,
        
        "tecnologías que se pueden percibir como inmaduras, de reciente creación o salida al mercado": 26,
        "detalle de las tecnologías que puedan ser percibidas como inmaduras utilizadas": 27,
        
        "tecnologías que se pueden percibir como especialmente intrusivas": 28,
        "justificación contacto intrusivo": 29,
        
        "enriquece la información de los interesados mediante la recogida": 30,
        "justificación enriquecimiento": 31,
        
        "implica que un elevado número de personas": 32,
        "justificación acceso elevado": 33,
        
        "observación de zonas de acceso público": 34,
        "justificación zonas públicas": 35,
        
        "personal no disociados o no anonimizados de forma irreversible con fines estadísticos": 36,
        "justificación fines estadísticos": 37,
        
        "ejercer un derecho, utilizar un servicio o ejecutar un contrato": 38,
        "justificación impedir derechos": 39,

        # Relaciones Externas y Otros
        "realizan cesiones de datos a otras entidades ": 40,
        "detalle de las cesiones realizadas": 41,
        
        "transferencias internacionales de datos a países fuera": 42,
        "detalle de las transferencias realizadas": 43,
        
        "similar a otro para el que haya sido necesario realizar un eipd": 44,
        "justificación de la percepción por parte del responsable de la actividad de tratamiento respecto de la similitud a otro tratamiento para el que haya sido necesario realizar un eipd": 45,
        
        "conllevar una pérdida o alteración de la información": 46,
        "justificación de la percepción por parte del responsable de la actividad de tratamiento respecto de la posibilidad de pérdida o alteración de la información": 47,
        
        "utilizada documentación en papel para tratar datos personales": 48,
        "justificación por parte del responsable de la actividad de tratamiento de las medidas aplicadas a la documentación en papel": 49,
        
        "interviene algún proveedor en el proceso": 50,
        "justificación de los proveedores que intervienen en el proceso": 51,

        # Tratamiento de los datos
        "indique los sistemas en los que se tratarán los datos": 52,
        "existen fines secundarios/intermedios con el tratamiento de los datos": 53,
        "especificar cuáles son los fines secundarios/intermedios": 54,
        "con qué frecuencia se recaban los datos": 55,
        "se recaban todos los datos afectados de forma conjunta o en distintos momentos": 56,
        "descripción del ciclo de vida": 57,

        # Operaciones y Conclusión
        "operaciones de tratamiento": 58,

        "infieren u obtienen datos adicionales a partir del tratamiento": 59,
        "justificación datos inferidos": 60,
        
        "prevén efectos colaterales o adversos para los interesados": 61,
        "justificación efectos adversos": 62,
        
        # Resultado del análisis
        # "resultado del análisis": 56
        "en base a las respuestas realizadas en el cuestionario de evaluación objetiva, debe calificarse la presente actividad de tratamiento como": 63,
        "ha indicado que la presente actividad presenta riesgo escaso o nulo": 64,
        "evidencias relativas al resultado del análisis: con riesgo escaso o nulo": 65,
        "ha indicado que la presente actividad presenta riesgo alto": 66,
        "evidencias relativas al resultado del análisis: riesgo alto": 67,
        "otra información relevante de procedimientos anteriores": 68,
        "evidencias de otro tipo de información relevante adicional": 69
    }

    resultado = {col: [] for col in COLUMNAS_ORDENADAS_AARR}
    resultado["Nombre del Archivo"] = [nombre_archivo]
    
    # 1. EXTRACCIÓN METADATOS (Proyecto / Tratamiento) - Se mantiene igual
    proy_tag = soup.find(class_="report_proyecto_name")
    trat_tag = soup.find('span', string=re.compile(r'Tratamiento:', re.I))# type: ignore
    if proy_tag or trat_tag:
        p_txt = proy_tag.get_text(strip=True) if proy_tag else ""
        t_txt = trat_tag.next_sibling.strip() if (trat_tag and trat_tag.next_sibling) else ""
        if 2 in vars_cols:
            resultado[vars_cols[2]].append(f"{p_txt} - {t_txt}")

    # 2. PROCESAMIENTO POR BLOQUES
    # Buscamos todos los divs que empiecen por levelQuestionClass_
    bloques = soup.find_all('div', class_=re.compile(r'levelQuestionClass_[01]'))# type: ignore
    
    # Variable para recordar en qué pregunta estamos (el "ancla" para la justificación)
    ultimo_idx_pregunta = None

    for bloque in bloques:
        # Extraemos la pregunta y la respuesta dentro del bloque
        preg_el = bloque.find(class_="pregunta")
        resp_el = bloque.find(class_="respuesta")
        
        if not preg_el or not resp_el:
            continue
            
        texto_pregunta = preg_el.get_text(" ", strip=True).lower()
        # La respuesta puede ser un P o una tabla, usamos get_text con separador
        texto_respuesta = resp_el.get_text("\n", strip=True)
        
        # --- CASO A: Es una pregunta principal (_0) ---
        if "levelQuestionClass_0" in bloque.get('class', []):# type: ignore
            encontrado = False
            for kw, idx in mapeo_indices.items():
                if kw.lower() in texto_pregunta:
                    col_name = vars_cols[idx]
                    resultado[col_name].append(texto_respuesta)
                    ultimo_idx_pregunta = idx # Guardamos el índice para la posible justificación
                    encontrado = True
                    break
            
            if not encontrado:
                # Opcional: print(f"Pregunta no mapeada: {texto_pregunta[:50]}")
                ultimo_idx_pregunta = None # Resetear si la pregunta no nos interesa

        # --- CASO B: Es una justificación (_1) ---
        elif "levelQuestionClass_1" in bloque.get('class', []):# type: ignore
            # Si sabemos a qué pregunta pertenece esta justificación
            if ultimo_idx_pregunta in relacion_pregunta_justificacion:
                idx_just = relacion_pregunta_justificacion[ultimo_idx_pregunta]
                col_name = vars_cols[idx_just]
                resultado[col_name].append(texto_respuesta)
                # print(f"Capturada justificación para la pregunta {ultimo_idx_pregunta}")

    # 3. APLANAR RESULTADO
    fila_final = {col: "\n".join(resultado[col]).strip() for col in COLUMNAS_ORDENADAS_AARR}
    return fila_final

def procesar_un_html_bloques(html_content, nombre_archivo):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 1. Preparar contenedores
    vars_cols = {i: nombre for i, nombre in enumerate(COLUMNAS_ORDENADAS_AARR)}
    resultado = {col: [] for col in COLUMNAS_ORDENADAS_AARR}
    resultado["Nombre del Archivo"] = [nombre_archivo]

    mapeo_indices = {
        # Metadatos
        "nombre del archivo": 0,
        "inicio_documento": 1,
        
        # Identificación básica
        "evaluación objetiva": 2,
        "descripción": 3,
        "figuras implicadas": 4,
        "notificar a": 5,
        
        # Alcance y Datos
        "número de sujetos afectados": 6,
        "categoría de datos tratados": 7,
        "orígenes de datos": 8,
        "duración del tratamiento": 9,
        "extensión geográfica": 10,
        "legitimación del tratamiento": 11,

        # Análisis de Riesgos Específicos (Pregunta + Justificación)
        "ratings/scoring o para la toma de decisiones": 12,
        "justificación ratings/scoring": 13,
        
        "toma de decisiones automatizada": 14,
        "justificación decisiones automatizadas": 15,
        
        "finalidad la monitorización o evaluación sistemática y exhaustiva": 16,
        "justificación monitorización": 17,
        
        "tratamiento de datos especialmente protegidos": 18,
        "justificación datos protegidos": 19,
        
        "finalidad el tratamiento a gran escala": 20,
        "justificación gran escala": 21,
        
        "se combinan conjuntos de datos utilizados": 22,
        "justificación combinación datos": 23,
        
        "colectivo en situación de especial vulnerabilidad": 24,
        "justificación vulnerabilidad": 25,
        
        "tecnologías que se pueden percibir como inmaduras, de reciente creación o salida al mercado": 26,
        "detalle de las tecnologías que puedan ser percibidas como inmaduras utilizadas": 27,
        
        "tecnologías que se pueden percibir como especialmente intrusivas": 28,
        "justificación contacto intrusivo": 29,
        
        "enriquece la información de los interesados mediante la recogida": 30,
        "justificación enriquecimiento": 31,
        
        "implica que un elevado número de personas": 32,
        "justificación acceso elevado": 33,
        
        "observación de zonas de acceso público": 34,
        "justificación zonas públicas": 35,
        
        "personal no disociados o no anonimizados de forma irreversible con fines estadísticos": 36,
        "justificación fines estadísticos": 37,
        
        "ejercer un derecho, utilizar un servicio o ejecutar un contrato": 38,
        "justificación impedir derechos": 39,

        # Relaciones Externas y Otros
        "realizan cesiones de datos a otras entidades ": 40,
        "detalle de las cesiones realizadas": 41,
        
        "transferencias internacionales de datos a países fuera": 42,
        "detalle de las transferencias realizadas": 43,
        
        "similar a otro para el que haya sido necesario realizar un eipd": 44,
        "justificación de la percepción por parte del responsable de la actividad de tratamiento respecto de la similitud a otro tratamiento para el que haya sido necesario realizar un eipd": 45,
        
        "conllevar una pérdida o alteración de la información": 46,
        "justificación de la percepción por parte del responsable de la actividad de tratamiento respecto de la posibilidad de pérdida o alteración de la información": 47,
        
        "utilizada documentación en papel para tratar datos personales": 48,
        "justificación por parte del responsable de la actividad de tratamiento de las medidas aplicadas a la documentación en papel": 49,
        
        "interviene algún proveedor en el proceso": 50,
        "justificación de los proveedores que intervienen en el proceso": 51,

        # Tratamiento de los datos
        "indique los sistemas en los que se tratarán los datos": 52,
        "existen fines secundarios/intermedios con el tratamiento de los datos": 53,
        "especificar cuáles son los fines secundarios/intermedios": 54,
        "con qué frecuencia se recaban los datos": 55,
        "se recaban todos los datos afectados de forma conjunta o en distintos momentos": 56,
        "descripción del ciclo de vida": 57,

        # Operaciones y Conclusión
        "operaciones de tratamiento": 58,

        "infieren u obtienen datos adicionales a partir del tratamiento": 59,
        "justificación datos inferidos": 60,
        
        "prevén efectos colaterales o adversos para los interesados": 61,
        "justificación efectos adversos": 62,
        
        # Resultado del análisis
        # "resultado del análisis": 56
        "en base a las respuestas realizadas en el cuestionario de evaluación objetiva, debe calificarse la presente actividad de tratamiento como": 63,
        "ha indicado que la presente actividad presenta riesgo escaso o nulo": 64,
        "evidencias relativas al resultado del análisis: con riesgo escaso o nulo": 65,
        "ha indicado que la presente actividad presenta riesgo alto": 66,
        "evidencias relativas al resultado del análisis: riesgo alto": 67,
        "otra información relevante de procedimientos anteriores": 68,
        "evidencias de otro tipo de información relevante adicional": 69
    }

    # --- 1. EXTRACCIÓN PROYECTO / TRATAMIENTO (Cabecera) ---
    col_titulo = "Tratamiento / Finalidad" 
    
    # Prioridad: Buscamos el H1 con la clase específica que indicas
    h1_titulo = soup.find('h1', class_='type_form_title')
    
    if h1_titulo:
        texto_raw = h1_titulo.get_text(strip=True)
        # re.sub(r'^\d+\s*', '', ...) quita los dígitos (^\d+) y el espacio (\s*) al inicio
        titulo_final = re.sub(r'^\d+\s*', '', texto_raw)
        resultado[col_titulo].append(titulo_final)
    else:
        # PLAN B: Si no existe el H1, intentamos el método anterior con limpieza similar
        proy_tag = soup.find(class_="report_proyecto_name")
        trat_tag = soup.find('span', string=re.compile(r'Tratamiento:', re.I))
        if (proy_tag or trat_tag) and col_titulo in resultado:
            p_txt = proy_tag.get_text(strip=True) if proy_tag else ""
            t_txt = trat_tag.next_sibling.strip() if (trat_tag and trat_tag.next_sibling) else ""
            cabecera_sucia = f"{p_txt} - {t_txt}"
            titulo_final = re.sub(r'^\d+\s*', '', cabecera_sucia)
            resultado[col_titulo].append(titulo_final)

    # --- PROCESAMIENTO DE PREGUNTAS Y JUSTIFICACIONES ---
    # Buscamos todos los bloques _0 y _1
    bloques = soup.find_all('div', class_=re.compile(r'levelQuestionClass_[01]'))

    for bloque in bloques:
        preg_el = bloque.find(class_="pregunta")
        resp_el = bloque.find(class_="respuesta")
        
        if not preg_el or not resp_el:
            continue
            
        # Limpiamos el texto de la pregunta para buscar en el mapeo
        t_preg_raw = preg_el.get_text(" ", strip=True)
        t_preg_low = t_preg_raw.lower()
        
        # Extraemos la respuesta (puede ser texto simple o tabla)
        # Usamos "\n" para preservar saltos de línea en justificaciones largas
        texto_respuesta = resp_el.get_text("\n", strip=True)

        # Buscamos en el mapeo_indices
        for kw, idx in mapeo_indices.items():
            if kw.lower() in t_preg_low:
                col_nombre = vars_cols[idx]
                
                # Guardamos la respuesta tal cual, sin el ruido de la cabecera
                if texto_respuesta and (not resultado[col_nombre] or resultado[col_nombre][-1] != texto_respuesta):
                    resultado[col_nombre].append(texto_respuesta)
                
                break

    # 3. APLANAR RESULTADO
    fila_final = {col: "\n".join(resultado[col]).strip() for col in COLUMNAS_ORDENADAS_AARR}
    return fila_final

def procesar_carpeta_htmls_pia(ruta_carpeta, nombre_csv_salida="consolidado_pias.csv"):
    # 1. Buscar todos los .html
    # patron = os.path.join(ruta_carpeta, "*.html")
    # archivos = glob.glob(patron)
    patron = os.path.join(ruta_carpeta, "**", "*.html")
    archivos = glob.glob(patron, recursive=True)
    
    if not archivos:
        print(f"No se encontraron archivos HTML en: {ruta_carpeta}")
        return

    print(f"Encontrados {len(archivos)} archivos. Procesando...")
    
    todas_las_filas = []

    # 2. Iterar sobre cada archivo
    for archivo in archivos:
        nombre_fichero = os.path.basename(archivo)
        try:
            # Importante: encoding utf-8 para tildes
            with open(archivo, 'r', encoding='utf-8') as f:
                contenido = f.read()
            
            # LLAMADA A LA FUNCIÓN DE EXTRACCIÓN (definida arriba)
            datos_fila = procesar_un_html_pia(contenido, nombre_fichero)
            todas_las_filas.append(datos_fila)
            print(f"OK: {nombre_fichero}")
            
        except Exception as e:
            print(f"ERROR en {nombre_fichero}: {e}")

    # 3. Crear DataFrame y Exportar
    if todas_las_filas:
        df_final = pd.DataFrame(todas_las_filas)
        
        # Reordenar columnas para asegurar que sigan el orden estricto de Excel
        # Nota: Ajustamos los nombres de Justificación en la lista de orden
        cols_finales = []
        for c in COLUMNAS_ORDENADAS_PIA:
            if "Justificación_" in c:
                cols_finales.append("Justificación")
            else:
                cols_finales.append(c)
        
        df_final.to_csv(nombre_csv_salida, index=False, encoding='utf-8-sig', sep=';', quoting=csv.QUOTE_ALL)
        print(f"\n¡ÉXITO! Se ha generado '{nombre_csv_salida}' con {len(df_final)} filas.")
    else:
        print("No se generaron datos.")

def procesar_carpeta_htmls_aarr(ruta_carpeta, nombre_csv_salida="consolidado_aarr.csv"):
    # 1. Buscar todos los .html
    # patron = os.path.join(ruta_carpeta, "*.html")
    # archivos = glob.glob(patron)
    patron = os.path.join(ruta_carpeta, "**", "*.html")
    archivos = glob.glob(patron, recursive=True)
    
    if not archivos:
        print(f"No se encontraron archivos HTML en: {ruta_carpeta}")
        return

    print(f"Encontrados {len(archivos)} archivos. Procesando...")
    
    todas_las_filas = []

    # 2. Iterar sobre cada archivo
    for archivo in archivos:
        nombre_fichero = os.path.basename(archivo)
        try:
            # Importante: encoding utf-8 para tildes
            with open(archivo, 'r', encoding='utf-8') as f:
                contenido = f.read()
            
            # LLAMADA A LA FUNCIÓN DE EXTRACCIÓN (definida arriba)
            # datos_fila = procesar_un_html_aarr(contenido, nombre_fichero)
            # datos_fila = procesar_un_html_estructurado(contenido, nombre_fichero)
            datos_fila = procesar_un_html_bloques(contenido, nombre_fichero)
            todas_las_filas.append(datos_fila)
            print(f"OK: {nombre_fichero}")
            
        except Exception as e:
            print(f"ERROR en {nombre_fichero}: {e}")

    # 3. Crear DataFrame y Exportar
    if todas_las_filas:
        df_final = pd.DataFrame(todas_las_filas)
        
        # Reordenar columnas para asegurar que sigan el orden estricto de Excel
        # Nota: Ajustamos los nombres de Justificación en la lista de orden
        cols_finales = []
        for c in COLUMNAS_ORDENADAS_AARR:
            if "Justificación_" in c:
                cols_finales.append("Justificación")
            else:
                cols_finales.append(c)
        
        df_final.to_csv(nombre_csv_salida, index=False, encoding='utf-8-sig', sep=';', quoting=csv.QUOTE_ALL)
        print(f"\n¡ÉXITO! Se ha generado '{nombre_csv_salida}' con {len(df_final)} filas.")
    else:
        print("No se generaron datos.")

def inspeccion_cruda_html(html_content, nombre_archivo):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Lista para guardar lo que encontremos tal cual aparece
    datos_detectados = []
    
    # Buscamos elementos que suelen ser títulos o preguntas
    elementos = soup.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'table', 'li'])
    
    current_titulo = "INICIO_DOCUMENTO"
    current_texto = []

    def guardar_bloque():
        if current_texto:
            # Unimos el contenido acumulado
            contenido = "\n".join(current_texto).strip()
            if contenido:
                datos_detectados.append({
                    "Archivo": nombre_archivo,
                    "Seccion_Detectada": current_titulo,
                    "Contenido_Extraido": contenido
                })
        current_texto.clear()

    for el in elementos:
        # Ignorar lo que está dentro de una tabla (lo procesamos cuando llegue la 'table')
        if el.find_parent('table') and el.name != 'table': continue
        
        texto_raw = el.get_text(" ", strip=True)
        if not texto_raw: continue
        
        clases = el.get('class', [])# type: ignore
        
        # CRITERIO DE TÍTULO: Es una cabecera H o tiene la clase 'pregunta'
        if el.name in ['h1', 'h2', 'h3', 'h4'] or 'pregunta' in clases:# type: ignore
            guardar_bloque()
            current_titulo = texto_raw
            continue

        # CRITERIO DE CONTENIDO
        if el.name == 'table':
            filas = [" | ".join(c.get_text(strip=True) for c in r.find_all(['td', 'th'])) for r in el.find_all('tr')]
            current_texto.append("\n".join(f"[{f}]" for f in filas if f.strip()))
        else:
            prefix = "• " if el.name == 'li' else ""
            current_texto.append(prefix + texto_raw)

    guardar_bloque() # Guardar el último
    return datos_detectados

def ejecutar_inspeccion(ruta_carpeta):
    archivos = glob.glob(os.path.join(ruta_carpeta, "**", "*.html"), recursive=True)
    todos_los_hallazgos = []

    for archivo in archivos:
        # Extraemos las partes de la ruta
        nombre_archivo = os.path.basename(archivo)        # Ejemplo: "reporte.html"
        ruta_directorio = os.path.dirname(archivo)        # Ejemplo: "/home/user/data/carpeta_A"
        nombre_carpeta = os.path.basename(ruta_directorio) # Ejemplo: "carpeta_A" (la carpeta inmediata)

        # Imprimimos para que lo veas en el output
        print(f"🔍 Procesando: [{nombre_carpeta}] -> {nombre_archivo}")

        with open(archivo, 'r', encoding='utf-8') as f:
            resultado = inspeccion_cruda_html(f.read(), archivo)
            todos_los_hallazgos.extend(resultado)

    if todos_los_hallazgos:
        # Generamos un CSV vertical (mucho más fácil de leer para depurar)
        df_debug = pd.DataFrame(todos_los_hallazgos)
        df_debug.to_csv('LECTURA_CRUDA_HTML.csv', index=False, encoding='utf-8-sig', sep=';')
        
        print("--- INSPECCIÓN FINALIZADA ---")
        print(f"Se han detectado {len(df_debug)} bloques de información.")
        print("Revisa el archivo 'LECTURA_CRUDA_HTML.csv' para ver los nombres reales de las columnas.")
    else:
        print("No se encontró nada.")

# --- EJECUCIÓN ---
# Cambia '.' por la ruta de tu carpeta si los HTML están en otro sitio
# ruta_mis_htmls = r'C:\.git\asisa_ocr\asisaocr\src' 
# ruta_mis_htmls = r'C:\Users\jbertrandelis\OneDrive - Deloitte (O365D)\Desktop\Proyectos\18.Asisa - OCR\ASISA'
ruta_mis_htmls = r'C:\Users\jbertrandelis\OneDrive - Deloitte (O365D)\Desktop\Proyectos\18.Asisa - OCR\html evaluaciones objetivas'

ejecutar_inspeccion(ruta_mis_htmls)
# file_path = r"C:\.git\asisa_ocr\asisaocr\src\eprivacy.ecix.tech.html"
# procesar_carpeta_htmls_pia(ruta_mis_htmls)
procesar_carpeta_htmls_aarr(ruta_mis_htmls)
