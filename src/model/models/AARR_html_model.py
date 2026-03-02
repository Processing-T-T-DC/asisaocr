import re
from functools import cached_property
from typing import cast

from bs4 import BeautifulSoup

from src.model.model import Field, Model, ParsedFile, Section
from src.model.writers.excel_writer import ExcelWriter


class AARR_HTML_Model(Model):

    FONT_SIZES_MAPPING = {
        18: "title",
        14: "H1",
        12: "H2",
        11: "H3",
        9: "H4",
    }

    @cached_property
    def COLUMNAS_ORDENADAS_AARR(self) -> list[str]:
        # return [field.name for field in self.fields]
        return ["Nombre del Archivo"] + [field.name for field in self.fields]
    
    target: str | None
    parsed_file: ParsedFile | None

    @cached_property
    def fields(self) -> list[Field]:
        
        return [
            Field("Tratamiento / Finalidad", "B2", "text", True),
            Field("Evaluación objetiva", "C2", "text", True),
            Field("Descripción", "D2", "text", False),
            Field("Figuras implicadas en la EIPD?", "E2", "text", False),
            Field("Notificar a ", "F2", "text", False),
            Field("Nº sujetos afectados", "G2", "fixed_number_range", True),
            Field("Categoría de datos tratados", "H2", "text", True),
            Field("Orígenes de datos", "I2", "text", True),
            Field("Duración del tratamiento", "J2", "treatment_duration", True),
            Field("Extensión geográfica", "K2", "geographic_extension", True),
            Field("Legitimación del tratamiento", "L2", "text", True),
            Field("¿Se van a tratar datos personales para hacer ratings/scoring o toma de decisiones?", "M2", "yes_no", True),
            Field("Detalle/Justificación ratings/scoring", "N2", "text", False),
            Field("¿El tratamiento de los datos implica una toma de decisiones automatizada sin que haya ninguna persona que intervenga en la decisión o valore los resultados?", "O2", "yes_no", True),
            Field("Detalle/Justificación tratamiento de los datos", "P2", "text", False),
            Field("La recogida de los datos tiene como finalidad la monitorización o evaluación sistemática y exhaustiva de aspectos personales?", "Q2", "yes_no", False),
            Field("Detalle/Justificación monitorización", "R2", "text", False),
            Field("¿La recogida de los datos tiene como finalidad el tratamiento de datos especialmente protegidos?", "S2", "yes_no", True),
            Field("Detalle/Justificación tratamiento de datos", "T2", "text", False),
            Field("¿La recogida de los datos tiene como finalidad el tratamiento a gran escala?", "U2", "yes_no", True),
            Field("Detalle/Justificación tratamiento a gran escala", "V2", "text", False),
            Field("Para llevar a cabo este tratamiento, ¿se combinan conjuntos de datos utilizados por otros responsables de tratamiento cuya finalidad diste en exceso de las expectativas del interesado?", "W2", "yes_no", True),
            Field("Detalle/Justificación combinan conjuntos de datos", "X2", "text", False),
            Field("¿La finalidad del tratamiento implica el uso específico de datos de personas con discapacidad o cualquier otro colectivo en situación de especial vulnerabilidad?", "Y2", "yes_no", True),
            Field("Detalle/Justificación datos de personas con discapacidad", "Z2", "text", False),
            Field("¿Se prevé el uso de tecnologías que se pueden percibir como inmaduras, de reciente creación o salida al mercado, cuyo alcance no puede ser previsto por el interesado de forma clara o razonable e implique elevado riesgo para el acceso no autorizado?", "AA2", "yes_no", True),
            Field("Detalle/Justificación inmaduras", "AB2", "text", False),
            Field("¿El tratamiento involucra contacto con los interesados de manera que, dicho contacto, pueda resultar intrusivo o se prevé el uso de tecnologías que se pueden percibir como especialmente intrusivas en la privacidad?", "AC2", "yes_no", True),
            Field("Detalle/Justificación intrusivo", "AD2", "text", False),
            Field("¿Se enriquece la información de los interesados mediante la recogida de nuevas categorías de datos o se usen las existentes con nuevas finalidades que antes no se contemplaban, en particular, si estas finalidades son más intrusivas o inesperadas para los afectados, o incluso pueda llegar a bloquear el disfrute de algún servicio?", "AE2", "yes_no", True),
            Field("Detalle/Justificación enriquece", "AF2", "text", False),
            Field("¿El tratamiento implica que un elevado número de personas (más allá de las necesarias para llevar a cabo el mismo) tenga acceso a los datos personales tratados?", "AG2", "yes_no", True),
            Field("Detalle/Justificación elevado número", "AH2", "text", False),
            Field("¿Se van a tratar datos relativos a la observación de zonas de acceso público?", "AI2", "yes_no", True),
            Field("Detalle/Justificación acceso público", "AJ2", "text", False),
            Field("¿Se utilizan datos de carácter personal no disociados o no anonimizados de forma irreversible con fines estadísticos, históricos o de investigación científica?", "AK2", "yes_no", True),
            Field("Detalle/Justificación no anonimizados", "AL2", "text", False),
            Field("¿Puede el tratamiento impedir ejercer un derecho, utilizar un servicio o ejecutar un contrato?", "AM2", "yes_no", True),
            Field("Detalle/Justificación derecho", "AN2", "text", False),
            Field("¿Se realizan cesiones de datos a otras entidades privadas u otras organizaciones, ya sean del mismo grupo o proveedores externos al mismo?", "AO2", "yes_no", True),
            Field("Detalle de las cesiones realizadas", "AP2", "text", False),
            Field("¿Se realizan transferencias internacionales de datos a países fuera de la Unión Europea y que no cuenten con medidas de protección de datos de carácter personal similares a las establecidas por la Autoridad de Control ?", "AQ2", "yes_no", True),
            Field("Detalle/Justificación transferencias internacionales", "AR2", "text", False),
            Field("¿Es este tratamiento similar a otro para el que haya sido necesario realizar un EIPD?", "AS2", "yes_no", True),
            Field("Detalle/Justificación EIPD", "AT2", "text", False),
            Field("¿Este tratamiento puede conllevar una pérdida o alteración de la información?", "AU2", "yes_no", True),
            Field("Detalle/Justificación pérdida o alteración", "AV2", "text", False),
            Field("¿Es utilizada documentación en papel para tratar datos personales?", "AW2", "yes_no", True),
            Field("Detalle/Justificación en papel", "AX2", "text", False),
            Field("¿Interviene algún proveedor en el proceso?", "AY2", "yes_no", True),
            Field("Detalle/Justificación proveedor", "AZ2", "text", False),
            Field("Indique los sistemas en los que se tratarán los datos (Medios electrónicos (físicos o en la nube)/Papel)", "BA2", "text", False),
            Field("¿Existen fines secundarios/intermedios con el tratamiento de los datos?", "BB2", "yes_no", False),
            Field("Especificar cuáles son los fines secundarios/intermedios", "BC2", "text", False),
            Field("¿Con qué frecuencia se recaban los datos?", "BD2", "text", False),
            Field("¿Se recaban todos los datos afectados de forma conjunta o en distintos momentos?", "BE2", "text", False),
            Field("Descripción del ciclo de vida", "BF2", "text", False),
            Field("Operaciones de tratamiento", "BG2", "text", False),
            Field("¿Se infieren u obtienen datos adicionales a partir del tratamiento de datos original?", "BH2", "yes_no", True),
            Field("Detalle/Justificación datos adicionales", "BI2", "text", False),
            Field("¿Se prevén efectos colaterales o adversos para los interesados?", "BJ2", "yes_no", False),
            Field("Detalle/Justificación efectos colaterales o adversos", "BK2", "text", False),
            Field("¿En base a las respuestas realizadas en el cuestionario de evaluación objetiva, debe calificarse la presente actividad de tratamiento como de \"RIESGO ALTO\" para los derechos y libertades de las personas?", "BL2", "yes_no", False),
            Field("Ha indicado que la presente actividad presenta riesgo ESCASO o NULO, por lo tanto no existe la obligación de realizar un PIA. Por favor, indique los motivos que justifican esta decisión como resultado del análisis ", "BM2", "text", False),
            Field("Evidencias relativas al resultado del análisis: con riesgo ESCASO o NULO", "BN2", "text", False),
            Field("Ha indicado que la presente actividad presenta RIESGO ALTO, por lo tanto existe la obligación de realizar un PIA. Por favor, indique los motivos que justifican esta decisión como resultado del análisis ", "BO2", "text", False),
            Field("Evidencias relativas al resultado del análisis: RIESGO ALTO", "BP2", "text", False),
            Field("Otra información relevante de procedimientos anteriores", "BQ2", "text", False),
            Field("Evidencias de otro tipo de información relevante adicional", "BR2", "text", False)
        ]

    def process(self, html_content: bytes, nombre_archivo: str, target: str) -> ParsedFile:
        """
        Recibe el contenido HTML y devuelve un DICCIONARIO con los datos limpios.
        """
        soup = BeautifulSoup(html_content, 'html.parser')
        self.target = target
        
        # MAPEO DE PALABRAS CLAVE (HTML -> Indice en la lista global)
        mapeo_indices = {
            # Metadatos
            "nombre del archivo": 0,
            "tratamiento / finalidad": 1,

            # Adelantamos estos valores por coincidencias con otras preguntas, para evitar falsos positivos en el mapeo
            "descripción del ciclo de vida": 57,
            "en base a las respuestas realizadas en el cuestionario de evaluación objetiva, debe calificarse la presente actividad de tratamiento como": 63,

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

            # Operaciones y Conclusión
            "operaciones de tratamiento": 58,

            "infieren u obtienen datos adicionales a partir del tratamiento": 59,
            "justificación datos inferidos": 60,
            
            "prevén efectos colaterales o adversos para los interesados": 61,
            "justificación efectos adversos": 62,
            
            # Resultado del análisis
            # "resultado del análisis": 56
            "ha indicado que la presente actividad presenta riesgo escaso o nulo": 64,
            "evidencias relativas al resultado del análisis: con riesgo escaso o nulo": 65,
            "ha indicado que la presente actividad presenta riesgo alto": 66,
            "evidencias relativas al resultado del análisis: riesgo alto": 67,
            "otra información relevante de procedimientos anteriores": 68,
            "evidencias de otro tipo de información relevante adicional": 69
        }

        # 1. Preparar contenedores
        vars_cols = {i: nombre for i, nombre in enumerate(self.COLUMNAS_ORDENADAS_AARR)}
        resultado: dict[str, list[str]] = {col: [] for col in self.COLUMNAS_ORDENADAS_AARR}
        # resultado["Nombre del Archivo"] = [nombre_archivo]
        resultado[vars_cols[0]].append(nombre_archivo)

        # if resultado["Tratamiento / Finalidad"][0] == "Prueba":
        #     print("DEBUG: Archivo de prueba detectado. Verificando extracción...")
        # elif resultado["Tratamiento / Finalidad"][0] == "SERVICIOS MÉDICOS - CENTRO MÉDICO CARACAS":
        #     print("DEBUG: Archivo de prueba detectado. Verificando extracción...")
        
        # --- 1. EXTRACCIÓN PROYECTO / TRATAMIENTO (Cabecera) ---
        col_titulo = "Tratamiento / Finalidad" 
        
        # Prioridad: Buscamos el H1 con la clase específica que indicas
        h1_titulo = soup.find('h1', class_='type_form_title')
        
        if h1_titulo:
            texto_raw = h1_titulo.get_text(strip=True)
            # re.sub(r'^\d+\s*', '', ...) quita los dígitos (^\d+) y el espacio (\s*) al inicio
            titulo_final = re.sub(r'^\d+\s*', '', texto_raw)
            # resultado[col_titulo].append(titulo_final)
            resultado[vars_cols[1]].append(titulo_final)
        else:
            # Prioridad 2: Buscar el div con id "report_cover_title_h2"
            div_titulo = soup.find('div', id='report_cover_title_h2')
            if div_titulo:
                texto_raw = div_titulo.get_text(strip=True)
                
                # 1. Nos quedamos solo con lo que hay después de "GRUPO ASISA"
                if "GRUPO ASISA" in texto_raw:
                    texto_raw = texto_raw.split("GRUPO ASISA")[-1]
                
                # 2. Reemplazamos ">" y "->" por guiones "-"
                # BeautifulSoup ya convierte los &gt; en > automáticamente al hacer get_text()
                texto_procesado = texto_raw.replace('->', '-').replace('>', '-')
                
                # 3. Limpieza final: quitar números iniciales, espacios extra y guiones repetidos
                # Quitamos número inicial
                texto_procesado = re.sub(r'^\d+\s*', '', texto_procesado)
                # Quitamos espacios y guiones sobrantes a los lados
                titulo_final = texto_procesado.strip("- ").replace('  ', ' ')
                
                # resultado[col_titulo].append(titulo_final)
                resultado[vars_cols[1]].append(titulo_final)

        # --- PROCESAMIENTO DE PREGUNTAS Y JUSTIFICACIONES ---
        # Buscamos todos los bloques _0 y _1
        bloques = soup.find_all('div', class_=re.compile(r'levelQuestionClass_[01]'))

        for bloque in bloques:
            preg_el = bloque.find(class_="pregunta")
            # resp_el = bloque.find(class_="respuesta")
            resp_els = bloque.find_all(class_="respuesta")
            
            if not preg_el or not resp_els:
                continue
                
            # Limpiamos el texto de la pregunta para buscar en el mapeo
            t_preg_raw = preg_el.get_text(" ", strip=True)
            t_preg_low = t_preg_raw.lower()
            
            # Extraemos la respuesta (puede ser texto simple o tabla)
            # Usamos "\n" para preservar saltos de línea en justificaciones largas
            # texto_respuesta = resp_el.get_text("\n", strip=True)
            # textos_resp_lista = [r.get_text(strip=True) for r in resp_els if r.get_text(strip=True)]
            # textos_resp_lista = [r.get_text("\n", strip=True) for r in resp_els if r.get_text(strip=True)]
            # texto_respuesta = ", ".join(textos_resp_lista)
            textos_resp_lista = []
            for r in resp_els:
                # Capa 1: Sacamos el texto respetando los <br> y <p> como saltos de línea (\n)
                texto_raw = r.get_text("\n", strip=True)
                if not texto_raw:
                    continue
                
                # Capa 2: Dividimos el texto usando una Regex que detecta saltos de línea O puntos y coma
                # Esto separará "GREENCUBE; PAPEL" en ["GREENCUBE", " PAPEL"]
                sub_elementos = re.split(r'[\n;]+', texto_raw)
                
                lineas_limpias = []
                for el in sub_elementos:
                    # Capa 3: Quitamos bullets (- * •) y espacios sobrantes
                    el_sin_bullet = re.sub(r'^\s*[-*•]\s*', '', el)
                    texto_final = el_sin_bullet.strip()
                    
                    if texto_final:
                        lineas_limpias.append(texto_final)
                
                # Guardamos este bloque ya procesado
                if lineas_limpias:
                    textos_resp_lista.append("\n".join(lineas_limpias))
            texto_respuesta = "\n".join(textos_resp_lista)

            # Buscamos en el mapeo_indices
            for kw, idx in mapeo_indices.items():
                if kw.lower() in t_preg_low:
                    # col_nombre = vars_cols[idx-1]
                    col_nombre = vars_cols[idx]
                    
                    # Guardamos la respuesta tal cual, sin el ruido de la cabecera
                    if texto_respuesta and (not resultado[col_nombre] or resultado[col_nombre][-1] != texto_respuesta):
                        resultado[col_nombre].append(texto_respuesta)
                    
                    break

        # 3. APLANAR RESULTADO
        fila_final = {}
        for col in self.COLUMNAS_ORDENADAS_AARR:
            fila_final[col] = "\n".join(resultado[col]).strip()
        
        parsed_file = ParsedFile()
        parsed_file.raw = str(soup)
        parsed_file.title = nombre_archivo
        parsed_file.sections = []

        for col in self.COLUMNAS_ORDENADAS_AARR:
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
                    # column=self.COLUMNAS_ORDENADAS_AARR.index(section.title) + 1
                    column=self.COLUMNAS_ORDENADAS_AARR.index(section.title)
                )
            )
            entries.append(entry)

        file = ExcelWriter.WritableExcelFile(self.target, entries)

        return file


