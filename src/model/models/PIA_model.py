import os
import re
from typing import cast

from bs4 import BeautifulSoup
from src.model.model import Model, ParsedFile, Section
from src.model.writers.excel_writer import ExcelWriter


class PIA_Model(Model):

    FONT_SIZES_MAPPING = {
        18: "title",
        14: "H1",
        12: "H2",
        11: "H3",
        9: "H4",
    }

    COLUMNAS_ORDENADAS = [
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
        "3. Identificación de riesgos", "3.1. Operaciones relacionadas con los fines de tratamiento",
        "3.2. Tipos de datos utilizados", "3.3. Extensión y alcance del tratamiento",
        "3.4. Categorías de interesados", "3.5. Factores técnicos del tratamiento",
        "3.6. Recogida y generación de datos", "3.7. Efectos colaterales del tratamiento",
        "3.8. Categoría del responsable", "3.9. Comunicaciones de datos", "3.10. Seguridad de los tratamientos"
    ]

    target: str | None
    parsed_file: ParsedFile | None

    def process(self, html_content: bytes, nombre_archivo: str, target: str) -> ParsedFile:
        """
        Recibe el contenido HTML y devuelve un DICCIONARIO con los datos limpios.
        """
        soup = BeautifulSoup(html_content, 'html.parser')
        self.target = target
        
        # Mapeo de variables (Indices basados en la lista global)
        # Nota: Usamos la lista global para mantener coherencia
        vars_cols = {i: nombre for i, nombre in enumerate(self.COLUMNAS_ORDENADAS)}
        
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
            "alcance de la evaluación de impacto para la privacidad": 10,
            "2.1 preparar y realizar": 11,
            "datos son recopilados": 12,
            "orígenes de datos": 13,
            "cómo se tratan los datos?": 14,
            "sistemas en los que se tratarán": 15,
            "política de conservación": 16,
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
            "3 identificación de riesgos": 43,
            "1. operaciones relacionadas": 44,
            "2. tipos de datos utilizados": 45,
            "3. extensión y alcance": 46,
            "4. categorías de interesados": 47,
            "5. factores técnicos": 48,
            "6. recogida y generación": 49,
            "7. efectos colaterales": 50,
            "8. categoría del responsable": 51,
            "9. comunicaciones de datos": 52,
            "10. seguridad de los tratamientos": 53
        }

        # Diccionario resultado inicializado con vacíos
        resultado = {col: [] for col in self.COLUMNAS_ORDENADAS}
        resultado["Nombre del Archivo"] = [nombre_archivo] # Guardamos el nombre
        
        # 1. EXTRACCIÓN PROYECTO - TRATAMIENTO
        proy_tag = soup.find(class_="report_proyecto_name")
        trat_tag = soup.find('span', string=re.compile(r'Tratamiento:', re.I)) # type: ignore
        if proy_tag or trat_tag:
            p_txt = proy_tag.get_text(strip=True) if proy_tag else ""
            t_txt = trat_tag.next_sibling.strip() if (trat_tag and trat_tag.next_sibling) else ""
            resultado[vars_cols[2]].append(f"{p_txt} - {t_txt}")

        # 2. EXTRACCIÓN DE CONTENIDO
        # IMPORTANTE: Ahora trackeamos el INDICE (int), no solo el nombre
        curr_col_idx = 1
        
        for el in soup.find_all(['h1', 'h2', 'h3', 'h4', 'p', 'table', 'li', 'div']):
            if el.name == 'div' and el.find(['p', 'table']): 
                continue 
            if el.find_parent('table') and el.name != 'table': 
                continue
            
            texto_raw = el.get_text(" ", strip=True)
            # Limpieza ISO
            texto_limpio = re.sub(r'\[.*?\]', '', texto_raw).strip()
            
            if not texto_limpio: 
                continue
            
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
                    # print(f"  🔄 Justificación: [{curr_col_idx}] → [{nuevo_idx_detectado}]")
                    curr_col_idx = nuevo_idx_detectado
                    continue
                # Si ya estamos en una columna de justificación (20, 22, 24, 26, 28),
                # NO cambiamos de columna, solo hacemos continue para saltar el título
                elif curr_col_idx in [20, 22, 24, 26, 28]:
                    # # print(f"  🔄 Justificación adicional en [{curr_col_idx}] (se concatena)")
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
            clases = el.get('class', [])  # type: ignore
            es_titulo_visual = (el.name in ['h1', 'h2', 'h3', 'h4'] or 'pregunta' in clases) # type: ignore

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
                        if any(celdas): 
                            filas.append(f"[{' | '.join(celdas)}]")
                    val = "\n".join(filas)
                else:
                    prefix = "• " if el.name == 'li' else ""
                    val = prefix + texto_raw

                # DEBUG: Mostrar cuando capturamos SI o NO
                # if texto_limpio.upper() in ['SI', 'NO', 'SÍ']:
                    # print(f"  ✅ Capturando '{texto_limpio}' en [{curr_col_idx}] '{curr_col_name[:50]}...'")
                
                # Evitar duplicados y títulos
                if val.strip().lower() != curr_col_name.lower() and \
                    (not resultado[curr_col_name] or resultado[curr_col_name][-1] != val):
                        resultado[curr_col_name].append(val)

        # 3. APLANAR RESULTADO (Lista -> String)
        fila_final = {}
        for col in self.COLUMNAS_ORDENADAS:
            # Renombrar Justificaciones para el CSV final
            fila_final[col] = "\n".join(resultado[col]).strip()
        
        parsed_file = ParsedFile()
        parsed_file.raw = str(soup)
        parsed_file.title = nombre_archivo
        parsed_file.sections = []

        for col in self.COLUMNAS_ORDENADAS:
            sec = Section()
            sec.title = col
            sec.content = cast(str, fila_final[col])
            parsed_file.sections.append(sec)

        return parsed_file

    def create_writable_file(self, parsed_file: ParsedFile) -> ExcelWriter.WritableExcelFile:
        assert self.target is not None

        entries: list[ExcelWriter.WritableExcelFile.WritableExcelEntry] = []

        for section in parsed_file.sections:
            entry = ExcelWriter.WritableExcelFile.WritableExcelEntry(
                content = cast(str, section.content),
                parameters = ExcelWriter.WritableExcelFile.WritableExcelParameters(
                    row=1,
                    column=self.COLUMNAS_ORDENADAS.index(section.title) + 1
                )
            )
            entries.append(entry)

        

        file = ExcelWriter.WritableExcelFile(self.target, entries)

        return file


    # def procesar_carpeta_htmls(self, nombre_fichero: str, nombre_csv_salida: str = "consolidado_pias.csv"):
    #     # 1. Buscar todos los .html
        
    #     todas_las_filas = []

    #     # 2. Iterar sobre cada archivo
    #     try:
            
    #         # LLAMADA A LA FUNCIÓN DE EXTRACCIÓN (definida arriba)
    #         datos_fila = procesar_un_html(contenido, nombre_fichero)
    #         todas_las_filas.append(datos_fila)
    #         print(f"OK: {nombre_fichero}")
            
    #     except Exception as e:
    #         print(f"ERROR en {nombre_fichero}: {e}")

    #     # 3. Crear DataFrame y Exportar
    #     if todas_las_filas:
    #         df_final = pd.DataFrame(todas_las_filas)
            
    #         # Reordenar columnas para asegurar que sigan el orden estricto de Excel
    #         # Nota: Ajustamos los nombres de Justificación en la lista de orden
    #         cols_finales = []
    #         for c in COLUMNAS_ORDENADAS:
    #             if "Justificación_" in c:
    #                 cols_finales.append("Justificación")
    #             else:
    #                 cols_finales.append(c)
            
    #         df_final.to_csv(nombre_csv_salida, index=False, encoding='utf-8-sig', sep=';', quoting=csv.QUOTE_ALL)
    #         print(f"\n¡ÉXITO! Se ha generado '{nombre_csv_salida}' con {len(df_final)} filas.")
    #     else:
    #         print("No se generaron datos.")

