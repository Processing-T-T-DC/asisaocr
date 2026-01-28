
from typing import cast
from src.model.model import Model, ParsedFile
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell import Cell

from src.model.writers.excel_writer import ExcelWriter

class AARR_Model(Model):

    FONT_SIZES_MAPPING = {
        18: "title",
        14: "H1",
        12: "H2",
        11: "H3",
        9: "H4",
    }


    videostreaming_columns = {
        "evaluacion_objetiva": "Evaluación objetiva",
        "descripcion": "Descripción",
        "ubicacion_actividad": "Ubicación de la actividad en la organización (Categoría, tratamiento y finalidad)",
        "figuras_implicadas": "Figuras implicadas en la realización de la EIPD",
        "notificar_a": "Notificar a",
        "numero_de_sujetos_afectados": "Número de sujetos afectados",
        "categoria_datos_tratados": "Categoría de datos tratados",
        "como_recopila_datos": "¿Cómo se recopilan los datos personales?",
        "origen_de_datos": "Indique los orígenes de datos",
        "duracion_tratamiento": "Duración del tratamiento",
        "extension_geografica": "Extensión geográfica",
        "legitimacion_del_tratamiendo": "Legitimación del tratamiento (Indicación de la base legitimadora)",
        "datos_personales_perfiles": "¿Se van a tratar datos personales para elaborar perfiles, categorizar/segmentar, hacer ratings/scoring o para la toma de decisiones? <span class=\"green\">LOPDGDD</span>",
        "tratamiento_de_datos_implica": "¿El tratamiento de los datos implica una toma de decisiones automatizada sin que haya ninguna persona que intervenga en la decisión o valore los resultados?",
        "recogida_datos_finalidad_monitorizacion": "¿La recogida de los datos tiene como finalidad la monitorización o evaluación sistemática y exhaustiva de aspectos personales? <span class=\"green\">LOPDGDD</span>",
        "recogida_datos_finalidad_protegidos": "¿La recogida de los datos tiene como finalidad el tratamiento de datos especialmente protegidos? <span class=\"green\">LOPDGDD</span>",
        "recogida_datos_finalidad_gran_escala": "¿La recogida de los datos tiene como finalidad el tratamiento a gran escala? <span class=\"green\">LOPDGDD</span>",
        "combina_datos": "Para llevar a cabo este tratamiento, ¿se combinan conjuntos de datos utilizados por otros responsables de tratamiento cuya finalidad diste en exceso de las expectativas del interesado?",
        "implica_uso_especifico": "¿La finalidad del tratamiento implica el uso específico de datos de personas con discapacidad o cualquier otro colectivo en situación de especial vulnerabilidad? <span class=\"green\">LOPDGDD</span>",
        "tecnologias_inmaduras": "¿Se prevé el uso de tecnologías que se pueden percibir como inmaduras, de reciente creación o salida al mercado, cuyo alcance no puede ser previsto por el interesado de forma clara o razonable e implique elevado riesgo para el acceso no autorizado?",
        "detalle_tecnologias_inmaduras": "Detalle de las tecnologías que puedan ser percibidas como inmaduras utilizadas",
        "involucra_contacto_interesados": "¿El tratamiento involucra contacto con los interesados de manera que, dicho contacto, pueda resultar intrusivo o se prevé el uso de tecnologías que se pueden percibir como especialmente intrusivas en la privacidad?",
        "enriquece_informacion": "¿Se enriquece la información de los interesados mediante la recogida de nuevas categorías de datos o se usan las existentes con nuevas finalidades que antes no se contemplaban, en particular, si estas finalidades son más intrusivas o inesperadas para los afectados, o incluso pueda llegar a bloquear el disfrute de algún servicio?",
        "tratamiento_acceso_datos_personales": "¿El tratamiento implica que un elevado número de personas (más allá de las necesarias para llevar a cabo el mismo) tenga acceso a los datos personales tratados?",
        "datos_relativos_acceso_publico": "¿Se van a tratar datos relativos a la observación de zonas de acceso público?",
        "datos_personales_no_anonimados": "¿Se utilizan datos de carácter personal no disociados o no anonimizados de forma irreversible con fines estadísticos, históricos o de investigación científica?",
        "puede_impedir_ejercer_derecho": "¿Puede el tratamiento impedir ejercer un derecho, utilizar un servicio o ejecutar un contrato? <span class=\"green\">LOPDGDD</span>",
        "cesiones_otras_entidades_privadas": "¿Se realizan cesiones de datos a otras entidades privadas u otras organizaciones, ya sean del mismo grupo o proveedores externos al mismo?",
        "detalle_cesiones_otras_entidades_privadas": "Detalle de las cesiones realizadas",
        "transferencias_internacionales": "¿Se realizan transferencias internacionales de datos a países fuera de la Unión Europea y que no cuenten con medidas de protección de datos de carácter personal similares a las establecidas por la Autoridad de Control ?",
        "detalle_transferencias_internacionales": "Detalle de las transferencias realizadas",
        "tratamiento_similar_EPID": "¿Es este tratamiento similar a otro para el que haya sido necesario realizar un EIPD?",
        "justificacion_tratamiento_similar_EPID": "Justificación de la percepción por parte del responsable de la actividad de tratamiento respecto de la similitud a otro tratamiento para el que haya sido necesario realizar un EIPD",
        "tratamiento_conlleva_perdida_informacion": "¿Este tratamiento puede conllevar una pérdida o alteración de la información? <span class=\"green\">LOPDGDD</span>",
        "justificacion_tratamiento_conlleva_perdida_informacion": "Justificación de la percepción por parte del responsable de la actividad de tratamiento respecto de la posibilidad de pérdida o alteración de la información",
        "utiliza_papel_datos_personales": "¿Se utiliza documentación en papel para tratar datos personales?",
        "medidas_utiliza_papel_datos_personales": "Indique las medidas aplicadas a la documentación en papel",
        "justificacion_utiliza_papel_datos_personales": " Justificación por parte del responsable de la actividad de tratamiento de las medidas aplicadas a la documentación en papel",
        "interviene_proveedor_en_proceso": "¿Interviene algún proveedor en el proceso?",
        "justificacion_interviene_proveedor_en_proceso": "Justificación de los proveedores que intervienen en el proceso",
        "sistemas_tratando_datos": "Indique los sistemas en los que se tratarán los datos (Medios electrónicos (físicos o en la nube)/Papel)",
        "fines_secundarios_tratamiento_datos": "¿Existen fines secundarios/intermedios con el tratamiento de los datos? ",
        "especificacion_fines_secundarios_tratamiento_datos": "Especificar cuáles son los fines secundarios/intermedios",
        "frecuencia_recabacion_datos": "¿Con qué frecuencia se recaban los datos?",
        "recaban_todos_datos_conjunta": "¿Se recaban todos los datos afectados de forma conjunta o en distintos momentos?",
        "descripcion_ciclo_vida": "Descripción del ciclo de vida",
        "opciones_tratamiento": "Operaciones de tratamiento",
        "interfieren_datos_adicionales_de_tratamiento": "¿Se infieren u obtienen datos adicionales a partir del tratamiento de datos original?",
        "especificacion_interfieren_datos_adicionales_de_tratamiento": "Especificar qué datos adicionales se infieren u obtienen",
        "contexto_del_tratamiento": "CONTEXTO DEL TRATAMIENTO",
        "mercado_sector_actividad": "Mercado y sector de actividad",
        "preven_efectos_colaterales_interesados": "¿Se prevén efectos colaterales o adversos para los interesados?",
        "especificacion_preven_efectos_colaterales_interesados": "Especificar qué efectos colaterales o adversos se prevén",
        "es_riesgo_alto": "¿En base a las respuestas realizadas en el cuestionario de evaluación objetiva, debe calificarse la presente actividad de tratamiento como de \"RIESGO ALTO\" para los derechos y libertades de las personas?",
        "justificacion_riesgo_escaso_PIA": "Ha indicado que la presente actividad presenta riesgo ESCASO o NULO, por lo tanto no existe la obligación de realizar un PIA. Por favor, indique los motivos que justifican esta decisión como resultado del análisis ",
        "evidencias_relativas_resultado": "Evidencias relativas al resultado del análisis: con riesgo ESCASO o NULO",
        "justificacion_riesgo_alto_PIA": "Ha indicado que la presente actividad presenta RIESGO ALTO, por lo tanto existe la obligación de realizar un PIA. Por favor, indique los motivos que justifican esta decisión como resultado del análisis ",
        "evidencias_relativas_riesgo_alto": "Evidencias relativas al resultado del análisis: RIESGO ALTO",
        "otra_informacion_relevante": "Otra información relevante de procedimientos anteriores",
        "evidencias_otra_informacion_relevante": "Evidencias de otro tipo de información relevante adicional",
    }

    entries: list[ExcelWriter.WritableExcelFile.WritableExcelEntry] = []
    target: str | None

    def _get_text_from_cell_or_empty_string(self, cell: Cell) -> str:
        return cast(str, cell.value)

    def process(self, workbook: Workbook, target: str):
        self.grab_first_row_of_data(workbook)
        self.target = target

    def grab_first_row_of_data(self, workbook: Workbook):
        sheet = workbook["Evaluacion objetiva"]
        try:
            cells = sheet["A4:BH4"][0]
            cells: tuple[Cell] = cells

            for cell in cells:
                params = ExcelWriter.WritableExcelFile.WritableExcelParameters(cell.row, cell.column)
                self.entries.append(ExcelWriter.WritableExcelFile.WritableExcelEntry(self._get_text_from_cell_or_empty_string(cell), params))

        except IndexError as e:
            raise e

    def create_writable_file(self, parsed_file: ParsedFile) -> ExcelWriter.WritableExcelFile:
        
        # entries = []

        # for cell in self.entries:
        #     parameters = ExcelWriter.WritableExcelFile.WritableExcelParameters(ce)
        #     entry = ExcelWriter.WritableExcelFile.WritableExcelEntry(cell, parameters)
        #     entries.append(entry)

        assert self.target is not None

        file = ExcelWriter.WritableExcelFile(self.target, self.entries)
        
        return file
        
        




