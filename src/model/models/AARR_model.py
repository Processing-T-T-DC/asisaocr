
from functools import cached_property
from typing import Literal, cast

from openpyxl import Workbook
from openpyxl.cell import Cell, ReadOnlyCell
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

from src.model.model import Field, Model, ParsedFile
from src.model.writers.excel_writer import ExcelWriter

type ColumnKey = Literal[
"evaluacion_objetiva",
"descripcion",
"ubicacion_actividad",
"figuras_implicadas",
"notificar_a",
"numero_de_sujetos_afectados",
"categoria_datos_tratados",
"como_recopila_datos",
"origen_de_datos",
"duracion_tratamiento",
"extension_geografica",
"legitimacion_del_tratamiendo",
"datos_personales_perfiles",
"tratamiento_de_datos_implica",
"recogida_datos_finalidad_monitorizacion",
"recogida_datos_finalidad_protegidos",
"recogida_datos_finalidad_gran_escala",
"combina_datos",
"implica_uso_especifico",
"tecnologias_inmaduras",
"detalle_tecnologias_inmaduras",
"involucra_contacto_interesados",
"enriquece_informacion",
"tratamiento_acceso_datos_personales",
"datos_relativos_acceso_publico",
"datos_personales_no_anonimados",
"puede_impedir_ejercer_derecho",
"cesiones_otras_entidades_privadas",
"detalle_cesiones_otras_entidades_privadas",
"transferencias_internacionales",
"detalle_transferencias_internacionales",
"tratamiento_similar_EPID",
"justificacion_tratamiento_similar_EPID",
"tratamiento_conlleva_perdida_informacion",
"justificacion_tratamiento_conlleva_perdida_informacion",
"utiliza_papel_datos_personales",
"medidas_utiliza_papel_datos_personales",
"justificacion_utiliza_papel_datos_personales",
"interviene_proveedor_en_proceso",
"justificacion_interviene_proveedor_en_proceso",
"sistemas_tratando_datos",
"fines_secundarios_tratamiento_datos",
"especificacion_fines_secundarios_tratamiento_datos",
"frecuencia_recabacion_datos",
"recaban_todos_datos_conjunta",
"descripcion_ciclo_vida",
"opciones_tratamiento",
"interfieren_datos_adicionales_de_tratamiento",
"especificacion_interfieren_datos_adicionales_de_tratamiento",
"contexto_del_tratamiento",
"mercado_sector_actividad",
"preven_efectos_colaterales_interesados",
"especificacion_preven_efectos_colaterales_interesados",
"es_riesgo_alto",
"justificacion_riesgo_escaso_PIA",
"evidencias_relativas_resultado",
"justificacion_riesgo_alto_PIA",
"evidencias_relativas_riesgo_alto",
"otra_informacion_relevante",
"evidencias_otra_informacion_relevante",

]

class AARR_Model(Model):

    FONT_SIZES_MAPPING = {
        18: "title",
        14: "H1",
        12: "H2",
        11: "H3",
        9: "H4",
    }
    
    @cached_property
    def fields(self) -> list[Field]:
        return [
            Field("Evaluación objetiva", "B2", "text", True),
            Field("Descripción", "C2", "text", False),
            Field("Figuras implicadas en la EIPD?", "D2", "text", False),
            Field("Notificar a ", "E2", "text", False),
            Field("Nº sujetos afectados", "F2", "fixed_number_range", True),
            Field("Categoría de datos tratados", "G2", "text", True),
            Field("Orígenes de datos", "H2", "text", True),
            Field("Duración del tratamiento", "I2", "treatment_duration", True),
            Field("Extensión geográfica", "J2", "geographic_extension", True),
            Field("Legitimación del tratamiento", "K2", "text", True),
            Field("¿Se van a tratar datos personales para hacer ratings/scoring o toma de decisiones?", "L2", "yes_no", True),
            Field("Detalle/Justificación", "M2", "text", False),
            Field("¿El tratamiento de los datos implica una toma de decisiones automatizada sin que haya ninguna persona que intervenga en la decisión o valore los resultados?", "N2", "yes_no", True),
            Field("Detalle/Justificación (si existe)", "O2", "text", False),
            Field("La recogida de los datos tiene como finalidad la monitorización o evaluación sistemática y exhaustiva de aspectos personales?", "P2", "yes_no", False),
            Field("Detalle/Justificación", "Q2", "text", False),
            Field("¿La recogida de los datos tiene como finalidad el tratamiento de datos especialmente protegidos?", "R2", "yes_no", True),
            Field("Detalle/Justificación", "S2", "text", False),
            Field("¿La recogida de los datos tiene como finalidad el tratamiento a gran escala?", "T2", "yes_no", True),
            Field("Detalle/Justificación", "U2", "text", False),
            Field("Para llevar a cabo este tratamiento, ¿se combinan conjuntos de datos utilizados por otros responsables de tratamiento cuya finalidad diste en exceso de las expectativas del interesado?", "V2", "yes_no", True),
            Field("Detalle/Justificación", "W2", "text", False),
            Field("¿La finalidad del tratamiento implica el uso específico de datos de personas con discapacidad o cualquier otro colectivo en situación de especial vulnerabilidad?", "X2", "yes_no", True),
            Field("Detalle/Justificación", "Y2", "text", False),
            Field("¿Se prevé el uso de tecnologías que se pueden percibir como inmaduras, de reciente creación o salida al mercado, cuyo alcance no puede ser previsto por el interesado de forma clara o razonable e implique elevado riesgo para el acceso no autorizado?", "Z2", "yes_no", True),
            Field("Detalle/Justificación", "AA2", "text", False),
            Field("¿El tratamiento involucra contacto con los interesados de manera que, dicho contacto, pueda resultar intrusivo o se prevé el uso de tecnologías que se pueden percibir como especialmente intrusivas en la privacidad?", "AB2", "yes_no", True),
            Field("Detalle/Justificación", "AC2", "text", False),
            Field("¿Se enriquece la información de los interesados mediante la recogida de nuevas categorías de datos o se usen las existentes con nuevas finalidades que antes no se contemplaban, en particular, si estas finalidades son más intrusivas o inesperadas para los afectados, o incluso pueda llegar a bloquear el disfrute de algún servicio?", "AD2", "yes_no", True),
            Field("Detalle/justificación", "AE2", "text", False),
            Field("¿El tratamiento implica que un elevado número de personas (más allá de las necesarias para llevar a cabo el mismo) tenga acceso a los datos personales tratados?", "AF2", "yes_no", True),
            Field("Detalle/Justificación", "AG2", "text", False),
            Field("¿Se van a tratar datos relativos a la observación de zonas de acceso público?", "AH2", "yes_no", True),
            Field("Detalle/Justificación", "AI2", "text", False),
            Field("¿Se utilizan datos de carácter personal no disociados o no anonimizados de forma irreversible con fines estadísticos, históricos o de investigación científica?", "AJ2", "yes_no", True),
            Field("Detalle/Justificación", "AK2", "text", False),
            Field("¿Puede el tratamiento impedir ejercer un derecho, utilizar un servicio o ejecutar un contrato?", "AL2", "yes_no", True),
            Field("Detalle/Justificación", "AM2", "text", False),
            Field("¿Se realizan cesiones de datos a otras entidades privadas u otras organizaciones, ya sean del mismo grupo o proveedores externos al mismo?", "AN2", "yes_no", True),
            Field("Detalle de las cesiones realizadas", "AO2", "text", False),
            Field("¿Se realizan transferencias internacionales de datos a países fuera de la Unión Europea y que no cuenten con medidas de protección de datos de carácter personal similares a las establecidas por la Autoridad de Control ?", "AP2", "yes_no", True),
            Field("Detalle/Justifiación", "AQ2", "text", False),
            Field("¿Es este tratamiento similar a otro para el que haya sido necesario realizar un EIPD?", "AR2", "yes_no", True),
            Field("Detalle/Justificacación", "AS2", "text", False),
            Field("¿Este tratamiento puede conllevar una pérdida o alteración de la información?", "AT2", "yes_no", True),
            Field("Justificación", "AU2", "text", False),
            Field("¿Es utilizada documentación en papel para tratar datos personales?", "AV2", "yes_no", True),
            Field("Detalle/Justificación", "AW2", "text", False),
            Field("¿Interviene algún proveedor en el proceso?", "AX2", "yes_no", True),
            Field("Justificación", "AY2", "text", False),
            Field("Operaciones de tratamiento", "AZ2", "text", False),
            Field("¿Se infieren u obtienen datos adicionales a partir del tratamiento de datos original?", "BA2", "yes_no", True),
            Field("Detalle/justificación", "BB2", "text", False),
            Field("¿Se prevén efectos colaterales o adversos para los interesados?", "BC2", "yes_no", False),
            Field("Detalle/justificación", "BD2", "text", False),
            Field("¿En base a las respuestas realizadas en el cuestionario de evaluación objetiva, debe calificarse la presente actividad de tratamiento como de \"RIESGO ALTO\" para los derechos y libertades de las personas?", "BE2", "yes_no", False),
            Field("Ha indicado que la presente actividad presenta riesgo ESCASO o NULO, por lo tanto no existe la obligación de realizar un PIA. Por favor, indique los motivos que justifican esta decisión como resultado del análisis ", "BF2", "text", False),
            Field("Evidencias relativas al resultado del análisis: con riesgo ESCASO o NULO", "BG2", "text", False),
            Field("Ha indicado que la presente actividad presenta RIESGO ALTO, por lo tanto existe la obligación de realizar un PIA. Por favor, indique los motivos que justifican esta decisión como resultado del análisis ", "BH2", "text", False),
            Field("Evidencias relativas al resultado del análisis: RIESGO ALTO", "BI2", "text", False),
            Field("Otra información relevante de procedimientos anteriores", "BJ2", "text", False),
            Field("Evidencias de otro tipo de información relevante adicional", "BK2", "text", False)
        ]

    output_mapping: dict[ColumnKey, str] = {
        "evaluacion_objetiva": "B",
        "descripcion": "C",
        "figuras_implicadas": "D",
        "notificar_a": "E",
        "numero_de_sujetos_afectados": "F",
        "categoria_datos_tratados": "G",
        "origen_de_datos": "H",
        "duracion_tratamiento": "I",
        "ubicacion_actividad": "",
        "extension_geografica": "J",
        "legitimacion_del_tratamiendo": "K",
        "datos_personales_perfiles": "L",
        "tratamiento_de_datos_implica": "N",
        "como_recopila_datos": "",
        "recogida_datos_finalidad_monitorizacion": "P",
        "recogida_datos_finalidad_protegidos": "R",
        "recogida_datos_finalidad_gran_escala": "T",
        "combina_datos": "V",
        "implica_uso_especifico": "X",
        "tecnologias_inmaduras": "Z",
        "detalle_tecnologias_inmaduras": "AA",
        "involucra_contacto_interesados": "AB",
        "enriquece_informacion": "AD",
        "tratamiento_acceso_datos_personales": "AF",
        "datos_relativos_acceso_publico": "AH",
        "datos_personales_no_anonimados": "AJ",
        "puede_impedir_ejercer_derecho": "AL",
        "cesiones_otras_entidades_privadas": "AN",
        "detalle_cesiones_otras_entidades_privadas": "AO",
        "transferencias_internacionales": "AP",
        "detalle_transferencias_internacionales": "AQ",
        "tratamiento_similar_EPID": "AR",
        "justificacion_tratamiento_similar_EPID": "AS",
        "tratamiento_conlleva_perdida_informacion": "AT",
        "justificacion_tratamiento_conlleva_perdida_informacion": "AU",
        "utiliza_papel_datos_personales": "AV",
        "medidas_utiliza_papel_datos_personales": "",
        "justificacion_utiliza_papel_datos_personales": "AW",
        "interviene_proveedor_en_proceso": "AX",
        "justificacion_interviene_proveedor_en_proceso": "AY",
        "sistemas_tratando_datos": "",
        "fines_secundarios_tratamiento_datos": "",
        "especificacion_fines_secundarios_tratamiento_datos": "",
        "frecuencia_recabacion_datos": "",
        "recaban_todos_datos_conjunta": "",
        "descripcion_ciclo_vida": "",
        "opciones_tratamiento": "AZ",
        "interfieren_datos_adicionales_de_tratamiento": "BA",
        "especificacion_interfieren_datos_adicionales_de_tratamiento": "BB",
        "contexto_del_tratamiento": "",
        "mercado_sector_actividad": "",
        "preven_efectos_colaterales_interesados": "BC",
        "especificacion_preven_efectos_colaterales_interesados": "BD",
        "es_riesgo_alto": "",
        "justificacion_riesgo_escaso_PIA": "BE",
        "evidencias_relativas_resultado": "BF",
        "justificacion_riesgo_alto_PIA": "BG",
        "evidencias_relativas_riesgo_alto": "BH",
        "otra_informacion_relevante": "BI",
        "evidencias_otra_informacion_relevante": "BK",
    }

    output_row_start = 4

    class Evaluacion:
        evaluacion_objetiva: str
        descripcion: str
        ubicacion_actividad: str
        figuras_implicadas: str
        notificar_a: str
        numero_de_sujetos_afectados: str
        categoria_datos_tratados: str
        como_recopila_datos: str
        origen_de_datos: str
        duracion_tratamiento: str
        extension_geografica: str
        legitimacion_del_tratamiendo: str
        datos_personales_perfiles: str
        tratamiento_de_datos_implica: str
        recogida_datos_finalidad_monitorizacion: str
        recogida_datos_finalidad_protegidos: str
        recogida_datos_finalidad_gran_escala: str
        combina_datos: str
        implica_uso_especifico: str
        tecnologias_inmaduras: str
        detalle_tecnologias_inmaduras: str
        involucra_contacto_interesados: str
        enriquece_informacion: str
        tratamiento_acceso_datos_personales: str
        datos_relativos_acceso_publico: str
        datos_personales_no_anonimados: str
        puede_impedir_ejercer_derecho: str
        cesiones_otras_entidades_privadas: str
        detalle_cesiones_otras_entidades_privadas: str
        transferencias_internacionales: str
        detalle_transferencias_internacionales: str
        tratamiento_similar_EPID: str
        justificacion_tratamiento_similar_EPID: str
        tratamiento_conlleva_perdida_informacion: str
        justificacion_tratamiento_conlleva_perdida_informacion: str
        utiliza_papel_datos_personales: str
        medidas_utiliza_papel_datos_personales: str
        justificacion_utiliza_papel_datos_personales: str
        interviene_proveedor_en_proceso: str
        justificacion_interviene_proveedor_en_proceso: str
        sistemas_tratando_datos: str
        fines_secundarios_tratamiento_datos: str
        especificacion_fines_secundarios_tratamiento_datos: str
        frecuencia_recabacion_datos: str
        recaban_todos_datos_conjunta: str
        descripcion_ciclo_vida: str
        opciones_tratamiento: str
        interfieren_datos_adicionales_de_tratamiento: str
        especificacion_interfieren_datos_adicionales_de_tratamiento: str
        contexto_del_tratamiento: str
        mercado_sector_actividad: str
        preven_efectos_colaterales_interesados: str
        especificacion_preven_efectos_colaterales_interesados: str
        es_riesgo_alto: str
        justificacion_riesgo_escaso_PIA: str
        evidencias_relativas_resultado: str
        justificacion_riesgo_alto_PIA: str
        evidencias_relativas_riesgo_alto: str
        otra_informacion_relevante: str
        evidencias_otra_informacion_relevante: str

    
    entries: list[ExcelWriter.WritableExcelFile.WritableExcelEntry] = []
    target: str | None
    
    _current_cell: Cell | ReadOnlyCell
    _current_sheet: Worksheet

    def _get_text_from_cell_or_empty_string(self, cell: Cell) -> str:
        return cast(str, cell.value or '') 
    
    def _get_next_column(self) -> str:
        next_cell = self._current_sheet.cell(self._current_cell.row, self._current_cell.column + 1)
        assert isinstance(next_cell, (Cell))

        self._current_cell = next_cell

        return next_cell.column_letter


    def process(self, workbook: Workbook, target: str):
        self.grab_first_row_of_data(workbook)
        self.target = target

    def grab_first_row_of_data(self, workbook: Workbook):
        sheet = workbook.worksheets[0]
        self._current_sheet = sheet
        
        self.evaluacion = AARR_Model.Evaluacion()    

        self._current_cell = sheet["A4"]

        self.evaluacion.evaluacion_objetiva = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._current_cell.column_letter}{self._current_cell.row}"]) # A
        self.evaluacion.descripcion = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # B
        self.evaluacion.ubicacion_actividad = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # C
        self.evaluacion.figuras_implicadas = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # D
        self.evaluacion.notificar_a = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # E
        self.evaluacion.numero_de_sujetos_afectados = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # F
        self.evaluacion.categoria_datos_tratados = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.como_recopila_datos = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.origen_de_datos = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.duracion_tratamiento = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.extension_geografica = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.legitimacion_del_tratamiendo = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.datos_personales_perfiles = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.tratamiento_de_datos_implica = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.recogida_datos_finalidad_monitorizacion = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # O
        self.evaluacion.recogida_datos_finalidad_protegidos = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # P
        self.evaluacion.recogida_datos_finalidad_gran_escala = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # Q
        self.evaluacion.combina_datos = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"]) # R
        self.evaluacion.implica_uso_especifico = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.tecnologias_inmaduras = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.detalle_tecnologias_inmaduras = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.involucra_contacto_interesados = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.enriquece_informacion = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.tratamiento_acceso_datos_personales = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.datos_relativos_acceso_publico = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.datos_personales_no_anonimados = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.puede_impedir_ejercer_derecho = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.cesiones_otras_entidades_privadas = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.detalle_cesiones_otras_entidades_privadas = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.transferencias_internacionales = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.detalle_transferencias_internacionales = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.tratamiento_similar_EPID = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.justificacion_tratamiento_similar_EPID = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.tratamiento_conlleva_perdida_informacion = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.justificacion_tratamiento_conlleva_perdida_informacion = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.utiliza_papel_datos_personales = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.medidas_utiliza_papel_datos_personales = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.justificacion_utiliza_papel_datos_personales = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.interviene_proveedor_en_proceso = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.justificacion_interviene_proveedor_en_proceso = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.sistemas_tratando_datos = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.fines_secundarios_tratamiento_datos = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.especificacion_fines_secundarios_tratamiento_datos = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.frecuencia_recabacion_datos = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.recaban_todos_datos_conjunta = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.descripcion_ciclo_vida = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.opciones_tratamiento = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.interfieren_datos_adicionales_de_tratamiento = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.especificacion_interfieren_datos_adicionales_de_tratamiento = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.contexto_del_tratamiento = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.mercado_sector_actividad = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.preven_efectos_colaterales_interesados = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.especificacion_preven_efectos_colaterales_interesados = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.es_riesgo_alto = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.justificacion_riesgo_escaso_PIA = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.evidencias_relativas_resultado = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.justificacion_riesgo_alto_PIA = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.evidencias_relativas_riesgo_alto = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.otra_informacion_relevante = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])
        self.evaluacion.evidencias_otra_informacion_relevante = self._get_text_from_cell_or_empty_string(self._current_sheet[f"{self._get_next_column()}{self._current_cell.row}"])



    
    def create_writable_file(self, parsed_file: ParsedFile) -> ExcelWriter.WritableExcelFile:
        entries = []
    
        for name, entry in vars(self.evaluacion).items():

            name = cast(ColumnKey, name)
            mapped_output_column = self.output_mapping[name]

            if mapped_output_column is None or len(mapped_output_column) == 0:
                continue

            params = ExcelWriter.WritableExcelFile.WritableExcelParameters(1, column_index_from_string(mapped_output_column))
            new_entry = ExcelWriter.WritableExcelFile.WritableExcelEntry(entry, params)

            # setattr(self.responsables, name, cell)
            entries.append(new_entry)
                                        
        assert self.target is not None

        file = ExcelWriter.WritableExcelFile(self.target, entries)
        
        return file

