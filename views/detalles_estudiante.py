import re
from models.registro_base import RegistroBase
from ui_compiled.ficha_estu_ui import Ui_ficha_estu
from PySide6.QtWidgets import QDialog, QMessageBox, QMenu, QToolButton, QComboBox, QVBoxLayout, QHBoxLayout, QPushButton, QLabel
from PySide6.QtCore import QDate, Signal, Qt
from PySide6.QtGui import QStandardItemModel, QStandardItem
from models.repre_model import RepresentanteModel
from models.estu_model import EstudianteModel
from models.institucion_model import InstitucionModel
from models.notas_model import NotasModel
from models.colaboracion_model import ColaboracionModel
from models.secciones_model import SeccionesModel
from models.anio_model import AnioEscolarModel
from utils.db import get_connection
from utils.widgets import Switch
from utils.exportar import (
    generar_constancia_estudios, generar_constancia_estudios_docx, generar_buena_conducta,
    generar_buena_conducta_docx,
    generar_constancia_inscripcion, generar_constancia_prosecucion_inicial,
    generar_constancia_prosecucion_inicial_docx,
    generar_constancia_aceptacion,
    generar_constancia_inscripcion_docx, generar_constancia_aceptacion_docx,
    generar_constancia_retiro_docx,
    generar_buena_conducta_retirado,
    generar_constancia_retiro, generar_historial_estudiante_pdf,
    generar_constancia_retiro_retirado,
    generar_historial_notas_pdf, generar_certificado_promocion_sexto,
    generar_certificado_promocion_sexto_docx, generar_certificado_prosecucion_primaria,
    generar_certificado_prosecucion_primaria_docx
)
from utils.sombras import crear_sombra_flotante
from utils.logo_manager import aplicar_logo_a_label
from utils.forms import ajustar_columnas_tabla
from utils.forms import set_campos_editables
from utils.dialogs import crear_msgbox
from utils.archivos import abrir_archivo
from datetime import date
from utils.edad import calcular_edad


class DetallesEstudiante(QDialog, Ui_ficha_estu):
    """Ventana de detalles de un estudiante."""
    
    # Señal emitida cuando se modifican datos (para refrescar tablas padre)
    datos_actualizados = Signal()

    def __init__(self, id_estudiante, usuario_actual, anio_escolar, es_egresado=False, parent=None):
        super().__init__(parent)
        self.usuario_actual = usuario_actual
        self.anio_escolar = anio_escolar
        self.es_egresado = es_egresado
        self.setupUi(self)

        self.setWindowTitle("Ficha de estudiante")
        self.id = id_estudiante
        self.id_estudiante = id_estudiante
        self.stackFicha_estu.setCurrentIndex(0)
        
        # Inicializar lista de delegates para tooltips
        self.tooltip_delegates = []
        
        # Variable para evitar bucles en las señales del switch
        self.actualizando_switch = False
        
        # Inicializar diccionarios vacíos para evitar errores
        self.grados_por_nivel = {}
        self.secciones_por_grado = {}
        
        # Cargar secciones desde la BD (solo si NO es egresado)
        if not self.es_egresado:
            self.cargar_secciones_en_combos()
            # Conectar señales de combos en cascada (nivel -> grado -> sección)
            self.cbxTipoEdu_ficha_estu.currentIndexChanged.connect(self.actualizar_grado)
            self.cbxGrado_ficha_estu.currentTextChanged.connect(self.actualizar_seccion)
        
        # Cargar datos del estudiante y su historial
        self.cargar_datos()
        self.cargar_historial()
        self.cargar_historial_notas()
        
        # Inicializar el switch de estado después de cargar datos
        self.switchActivo = Switch()
        self.switchActivo.setFixedSize(50, 25)
        self.contenedorSwitch.layout().addWidget(self.switchActivo)

        # Establecer el estado inicial del switch sin disparar eventos
        self.actualizando_switch = True
        # Estado=1 (activo) -> Checked(True) verde
        # Estado=0 (inactivo) -> Checked(False) gris
        self.switchActivo.setChecked(bool(self.estudiante_actual.get("Estado", 1)))
        self.actualizando_switch = False

        # Inicializar switch de colaboración
        self.switchColaboracion = Switch()
        self.switchColaboracion.setFixedSize(50, 25)
        self.contenedorSwitchColaboracion.layout().addWidget(self.switchColaboracion)

        self.actualizando_switch = True
        colaboro = ColaboracionModel.obtener_estado_actual(self.id_estudiante)
        self.switchColaboracion.setChecked(bool(colaboro))
        self.actualizando_switch = False

        # Configurar visibilidad según tipo (egresado vs regular)
        self.configurar_visibilidad_campos()

        # Conectar señales de botones de acción
        self.btnModificar_ficha_estu.clicked.connect(self.toggle_edicion)
        self.btnDevolver_grado.clicked.connect(self.devolver_estudiante)
        self.btnEliminar_ficha_estu.clicked.connect(self.eliminar_estudiante)

        self.configurar_menu_exportacion()
        
        # Conectar señales de fechas para cálculo automático de edad
        self.lneFechaNac_ficha_estu.dateChanged.connect(self.actualizar_edad_estudiante)
        self.lneFechaNac_repre_ficha_estu.dateChanged.connect(self.actualizar_edad_representante)
              
        # Bloquear campos y configurar estado inicial (modo lectura)
        self.set_campos_editables(False)
        self.lneCedula_repre_ficha_estu.setReadOnly(True)
        self.lneCedula_ficha_estu.setReadOnly(True)
        
        # Conectar la señal del switch después de establecer el estado inicial
        self.switchActivo.stateChanged.connect(self.cambiar_estado_estudiante)
        self.switchColaboracion.stateChanged.connect(self.cambiar_colaboracion)

        # Aplicar efectos visuales (sombras flotantes)
        self.aplicar_sombras()
    
    def configurar_menu_exportacion(self):
        """Configura el menú de exportación."""
        self.btnExportar_ficha_estu.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        menu_exportar_estu = QMenu(self.btnExportar_ficha_estu)
        
        # Agregar opciones de exportación
        menu_exportar_estu.addAction("Constancia de estudios (PDF)", self.exportar_constancia_estudios)
        menu_exportar_estu.addAction("Constancia de estudios (DOCX)", self.exportar_constancia_estudios_docx)
        menu_exportar_estu.addAction("Constancia de buena conducta", self.exportar_buena_conducta)
        menu_exportar_estu.addAction("Constancia de buena conducta (DOCX)", self.exportar_buena_conducta_docx)
        menu_exportar_estu.addAction("Constancia de buena conducta (retirado)", self.exportar_buena_conducta_retirado)
        menu_exportar_estu.addAction("Constancia de inscripción", self.exportar_constancia_inscripcion)
        menu_exportar_estu.addAction("Constancia de inscripción (DOCX)", self.exportar_constancia_inscripcion_docx)
        menu_exportar_estu.addAction("Constancia de aceptación", self.exportar_constancia_aceptacion)
        menu_exportar_estu.addAction("Constancia de aceptación (DOCX)", self.exportar_constancia_aceptacion_docx)
        menu_exportar_estu.addAction("Constancia prosecución Educación Inicial", 
                                     self.exportar_constancia_prosecucion_inicial)
        menu_exportar_estu.addAction("Constancia prosecución Educación Inicial (DOCX)", 
                                     self.exportar_constancia_prosecucion_inicial_docx)
        menu_exportar_estu.addAction("Certif. prosecución primaria", self.exportar_certificado_prosecucion_primaria)
        menu_exportar_estu.addAction("Certif. prosecución primaria (DOCX)", 
                                     self.exportar_certificado_prosecucion_primaria_docx)
        menu_exportar_estu.addAction("Certificado promoción 6to a Secundaria",
                                      self.exportar_certificado_promocion_sexto)
        menu_exportar_estu.addAction("Certificado promoción 6to a Secundaria (DOCX)",
                                      self.exportar_certificado_promocion_sexto_docx)
        menu_exportar_estu.addAction("Constancia de retiro (normal)", self.exportar_constancia_retiro)
        menu_exportar_estu.addAction("Constancia de retiro (normal) (DOCX)", self.exportar_constancia_retiro_docx)
        menu_exportar_estu.addAction("Constancia de retiro (grado anterior)", self.exportar_constancia_retiro_retirado)
        menu_exportar_estu.addSeparator()
        menu_exportar_estu.addAction("Exportar historial académico (PDF)", self.exportar_historial_pdf)
        menu_exportar_estu.addAction("Exportar historial de notas (PDF)", self.exportar_historial_notas_pdf)
        
        self.btnExportar_ficha_estu.setMenu(menu_exportar_estu)
    
    def obtener_estudiante_actual_dict(self):
        """Convierte los datos del formulario en dict."""
        return {
            "ID": str(self.id_estudiante),
            "Cédula": self.lneCedula_ficha_estu.text().strip(),
            "Nombres": self.lneNombre_ficha_estu.text().strip(),
            "Apellidos": self.lneApellido_ficha_estu.text().strip(),
            "Fecha Nac.": self.lneFechaNac_ficha_estu.date().toPython(),
            "Edad": self.lneEdad_ficha_estu.text().strip(),
            "Ciudad": self.lneCity_ficha_estu.text().strip(),
            "Género": self.cbxGenero_ficha_estu.currentText().strip(),
            "Dirección": self.lneDir_ficha_estu.text().strip(),
            "Fecha Ingreso": self.lneFechaIng_ficha_estu.date().toPython(),
            "Tipo Educ.": self.cbxTipoEdu_ficha_estu.currentText().strip() if not self.es_egresado else "",
            "Grado": self.cbxGrado_ficha_estu.currentText().strip() if not self.es_egresado else "",
            "Sección": self.cbxSeccion_ficha_estu.currentText().split("  (")[0].strip() if not self.es_egresado else "",
            "Docente": self.lneDocente_ficha_estu.text().strip(),
        }

    def exportar_constancia_estudios(self):
        """Genera constancia de estudios."""
        try:
            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_estudios(estudiante, institucion)
            crear_msgbox(self, "Éxito", f"Constancia generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_constancia_estudios_docx(self):
        """Genera constancia de estudios en formato DOCX."""
        try:
            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_estudios_docx(estudiante, institucion)
            crear_msgbox(self, "Éxito", f"Constancia (DOCX) generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar (DOCX):\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_constancia_inscripcion(self):
        """Genera constancia de inscripción."""
        try:
            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_inscripcion(estudiante, institucion)
            crear_msgbox(self, "Éxito", f"Constancia generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_constancia_inscripcion_docx(self):
        """Genera constancia de inscripción en formato DOCX."""
        try:
            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_inscripcion_docx(estudiante, institucion)
            crear_msgbox(self, "Éxito", f"Constancia (DOCX) generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar (DOCX):\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_constancia_aceptacion(self):
        """Genera constancia de aceptación."""
        try:
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_aceptacion(institucion, self.anio_escolar)
            crear_msgbox(self, "Éxito", f"Constancia generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_constancia_aceptacion_docx(self):
        """Genera constancia de aceptación en formato DOCX."""
        try:
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_aceptacion_docx(institucion, self.anio_escolar)
            crear_msgbox(self, "Éxito", f"Constancia (DOCX) generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar (DOCX):\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_certificado_prosecucion_primaria(self):
        """Genera certificado prosecucion primaria"""
        try:
            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_certificado_prosecucion_primaria(estudiante, institucion, self.anio_escolar)
            crear_msgbox(self, "Éxito", f"Certificado generado:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_certificado_prosecucion_primaria_docx(self):
        """Genera certificado de prosecución primaria en formato DOCX."""
        try:
            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_certificado_prosecucion_primaria_docx(estudiante, institucion, self.anio_escolar)
            crear_msgbox(self, "Éxito", f"Certificado (DOCX) generado:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar (DOCX):\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_buena_conducta(self):
        """Genera constancia de buena conducta"""
        try:
            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_buena_conducta(estudiante, institucion, self.anio_escolar)
            crear_msgbox(self, "Éxito", f"Constancia generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_buena_conducta_docx(self):
        """Genera constancia de buena conducta en formato DOCX."""
        try:
            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_buena_conducta_docx(estudiante, institucion, self.anio_escolar)
            crear_msgbox(self, "Éxito", f"Constancia (DOCX) generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar (DOCX):\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_buena_conducta_retirado(self):
        """Genera constancia de buena conducta para estudiante retirado (año anterior)."""
        try:
            # Verificar que el estudiante esté retirado
            datos_bd = EstudianteModel.obtener_por_id(self.id_estudiante)
            if not datos_bd:
                crear_msgbox(self, "Error", "No se encontró el estudiante.", QMessageBox.Icon.Critical).exec()
                return
            estatus_acad = str(datos_bd.get("estatus_academico", "")).strip()
            estado_est = datos_bd.get("estado", 1)
            if estatus_acad != "Retirado" and estado_est != 0:
                crear_msgbox(
                    self, "Estudiante no elegible",
                    "La constancia de buena conducta (retirado) solo se puede generar "
                    "para estudiantes con estatus académico 'Retirado'.",
                    QMessageBox.Icon.Warning
                ).exec()
                return

            # Obtener último grado cursado del historial (año anterior)
            historial = EstudianteModel.obtener_historial_estudiante(self.id_estudiante)
            if not historial:
                crear_msgbox(self, "Sin historial",
                             "No se encontró historial académico para este estudiante.",
                             QMessageBox.Icon.Warning).exec()
                return
            # Buscar el último curso anterior al año escolar activo
            # (el año que realmente cursó y aprobó, no el recién aperturado)
            anio_activo_inicio = int(self.anio_escolar['año_inicio'])
            ultimo_curso = None
            for registro in historial:
                if registro['año_inicio'] < anio_activo_inicio:
                    ultimo_curso = registro
                    break
            if not ultimo_curso:
                crear_msgbox(
                    self, "Sin historial previo",
                    "No se encontró un grado cursado y aprobado antes del año escolar activo.",
                    QMessageBox.Icon.Warning
                ).exec()
                return

            estudiante = self.obtener_estudiante_actual_dict()
            estudiante['Grado'] = ultimo_curso['grado']
            estudiante['Sección'] = ultimo_curso['letra']

            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_buena_conducta_retirado(
                estudiante, institucion, ultimo_curso['año_escolar']
            )
            crear_msgbox(self, "Éxito", f"Constancia generada:\n{archivo}",
                         QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_constancia_prosecucion_inicial(self):
        """Genera constancia de prosecución"""
        try:
            if self.es_egresado:
                crear_msgbox(
                    self,
                    "No disponible",
                    "Esta constancia solo aplica para estudiantes regulares desde 3er nivel en adelante.",
                    QMessageBox.Icon.Warning
                ).exec()
                return

            tipo_actual = self.cbxTipoEdu_ficha_estu.currentText().strip().lower()
            grado_actual = self.cbxGrado_ficha_estu.currentText().strip().lower()
            anio_inicio_actual = int(self.anio_escolar['año_inicio'])

            anio_escolar_inicial = None

            if tipo_actual in ['inicial', 'preescolar'] and '3' in grado_actual:
                anio_escolar_inicial = {
                    'año_inicio': anio_inicio_actual,
                    'año_fin': anio_inicio_actual + 1
                }
            elif tipo_actual == 'primaria':
                historial = EstudianteModel.obtener_historial_estudiante(self.id_estudiante)
                if not historial:
                    crear_msgbox(self, "Sin historial", "No hay historial.", QMessageBox.Icon.Warning).exec()
                    return

                curso_tercer_nivel = next(
                    (
                        r for r in historial
                        if '3' in str(r.get('grado', '')).lower()
                        and str(r.get('nivel', '')).lower() in ['inicial', 'preescolar']
                    ),
                    None
                )

                if not curso_tercer_nivel:
                    crear_msgbox(
                        self,
                        "No elegible",
                        "No se encontró registro de 3er nivel de educación inicial para este estudiante.",
                        QMessageBox.Icon.Warning
                    ).exec()
                    return

                anio_inicio_tercer_nivel = int(curso_tercer_nivel['año_inicio'])
                anio_escolar_inicial = {
                    'año_inicio': anio_inicio_tercer_nivel,
                    'año_fin': anio_inicio_tercer_nivel + 1
                }
            else:
                crear_msgbox(
                    self,
                    "Estudiante no elegible",
                    "Esta constancia está disponible desde 3er nivel de inicial en adelante.",
                    QMessageBox.Icon.Warning
                ).exec()
                return

            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_prosecucion_inicial(estudiante, institucion, anio_escolar_inicial)
            crear_msgbox(self, "Éxito", f"Constancia generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_constancia_prosecucion_inicial_docx(self):
        """Genera constancia de prosecución inicial en formato DOCX."""
        try:
            if self.es_egresado:
                crear_msgbox(
                    self,
                    "No disponible",
                    "Esta constancia solo aplica para estudiantes regulares desde 3er nivel en adelante.",
                    QMessageBox.Icon.Warning
                ).exec()
                return

            tipo_actual = self.cbxTipoEdu_ficha_estu.currentText().strip().lower()
            grado_actual = self.cbxGrado_ficha_estu.currentText().strip().lower()
            anio_inicio_actual = int(self.anio_escolar['año_inicio'])

            anio_escolar_inicial = None

            if tipo_actual in ['inicial', 'preescolar'] and '3' in grado_actual:
                anio_escolar_inicial = {
                    'año_inicio': anio_inicio_actual,
                    'año_fin': anio_inicio_actual + 1
                }
            elif tipo_actual == 'primaria':
                historial = EstudianteModel.obtener_historial_estudiante(self.id_estudiante)
                if not historial:
                    crear_msgbox(self, "Sin historial", "No hay historial.", QMessageBox.Icon.Warning).exec()
                    return

                curso_tercer_nivel = next(
                    (
                        r for r in historial
                        if '3' in str(r.get('grado', '')).lower()
                        and str(r.get('nivel', '')).lower() in ['inicial', 'preescolar']
                    ),
                    None
                )

                if not curso_tercer_nivel:
                    crear_msgbox(
                        self,
                        "No elegible",
                        "No se encontró registro de 3er nivel de educación inicial para este estudiante.",
                        QMessageBox.Icon.Warning
                    ).exec()
                    return

                anio_inicio_tercer_nivel = int(curso_tercer_nivel['año_inicio'])
                anio_escolar_inicial = {
                    'año_inicio': anio_inicio_tercer_nivel,
                    'año_fin': anio_inicio_tercer_nivel + 1
                }
            else:
                crear_msgbox(
                    self,
                    "Estudiante no elegible",
                    "Esta constancia está disponible desde 3er nivel de inicial en adelante.",
                    QMessageBox.Icon.Warning
                ).exec()
                return

            estudiante = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_prosecucion_inicial_docx(estudiante, institucion, anio_escolar_inicial)
            crear_msgbox(self, "Éxito", f"Constancia (DOCX) generada:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar (DOCX):\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_certificado_promocion_sexto(self):
        """Genera certificado de promoción de 6to a 1er año de secundaria."""
        try:
            anio_escolar_certificado = None
            ultima_seccion = None

            if self.es_egresado:
                historial = EstudianteModel.obtener_historial_estudiante(self.id_estudiante)
                if not historial:
                    crear_msgbox(self, "Sin historial", "No hay historial.", QMessageBox.Icon.Warning).exec()
                    return

                curso_sexto = next(
                    (
                        r for r in historial
                        if 'primaria' in str(r.get('nivel', '')).lower()
                        and '6' in str(r.get('grado', '')).lower()
                    ),
                    None
                )

                if not curso_sexto:
                    crear_msgbox(
                        self,
                        "No elegible",
                        "Este estudiante no cursó 6to grado en esta institución.",
                        QMessageBox.Icon.Warning
                    ).exec()
                    return

                anio_escolar_certificado = curso_sexto['año_escolar']
                ultima_seccion = curso_sexto.get('letra')
            else:
                tipo_actual = self.cbxTipoEdu_ficha_estu.currentText().strip().lower()
                grado_actual = self.cbxGrado_ficha_estu.currentText().strip().lower()

                if tipo_actual != 'primaria' or '6' not in grado_actual:
                    crear_msgbox(
                        self,
                        "Estudiante no elegible",
                        "Este certificado solo se puede generar para estudiantes cursando 6to grado o egresados.",
                        QMessageBox.Icon.Warning
                    ).exec()
                    return

                anio_inicio = int(self.anio_escolar['año_inicio'])
                anio_escolar_certificado = f"{anio_inicio}/{anio_inicio + 1}"
                ultima_seccion = self.cbxSeccion_ficha_estu.currentText().split("  (")[0].strip()

            estudiante = self.obtener_estudiante_actual_dict()
            estudiante['ultima_seccion'] = ultima_seccion or 'N/A'

            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_certificado_promocion_sexto(
                estudiante,
                institucion,
                anio_escolar_certificado
            )

            crear_msgbox(self, "Éxito", f"Certificado generado:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_certificado_promocion_sexto_docx(self):
        """Genera certificado de promoción de 6to en formato DOCX."""
        try:
            anio_escolar_certificado = None
            ultima_seccion = None

            if self.es_egresado:
                historial = EstudianteModel.obtener_historial_estudiante(self.id_estudiante)
                if not historial:
                    crear_msgbox(self, "Sin historial", "No hay historial.", QMessageBox.Icon.Warning).exec()
                    return

                curso_sexto = next(
                    (
                        r for r in historial
                        if 'primaria' in str(r.get('nivel', '')).lower()
                        and '6' in str(r.get('grado', '')).lower()
                    ),
                    None
                )

                if not curso_sexto:
                    crear_msgbox(
                        self,
                        "No elegible",
                        "Este estudiante no cursó 6to grado en esta institución.",
                        QMessageBox.Icon.Warning
                    ).exec()
                    return

                anio_escolar_certificado = curso_sexto['año_escolar']
                ultima_seccion = curso_sexto.get('letra')
            else:
                tipo_actual = self.cbxTipoEdu_ficha_estu.currentText().strip().lower()
                grado_actual = self.cbxGrado_ficha_estu.currentText().strip().lower()

                if tipo_actual != 'primaria' or '6' not in grado_actual:
                    crear_msgbox(
                        self,
                        "Estudiante no elegible",
                        "Este certificado solo se puede generar para estudiantes cursando 6to grado o egresados.",
                        QMessageBox.Icon.Warning
                    ).exec()
                    return

                anio_inicio = int(self.anio_escolar['año_inicio'])
                anio_escolar_certificado = f"{anio_inicio}/{anio_inicio + 1}"
                ultima_seccion = self.cbxSeccion_ficha_estu.currentText().split(" (")[0].strip()

            estudiante = self.obtener_estudiante_actual_dict()
            estudiante['ultima_seccion'] = ultima_seccion or 'N/A'

            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_certificado_promocion_sexto_docx(
                estudiante,
                institucion,
                anio_escolar_certificado
            )

            crear_msgbox(self, "Éxito", f"Certificado DOCX generado:\n{archivo}", QMessageBox.Icon.Information).exec()
            abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"No se pudo generar:\n{e}", QMessageBox.Icon.Critical).exec()

    def exportar_constancia_retiro(self):
        """Genera constancia de retiro."""
        # Verificar si el estudiante está inactivo
        if self.estudiante_actual.get("Estado", 1) == 1:
            crear_msgbox(
                self,
                "Estudiante activo",
                "La constancia de retiro solo se puede generar para estudiantes retirados (inactivos).\n\n"
                "Use el switch para marcar al estudiante como retirado primero.",
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        try:
            # Obtener motivo de retiro desde la BD
            datos = EstudianteModel.obtener_por_id(self.id_estudiante)
            motivo_retiro = datos.get("motivo_retiro") if datos else None
            
            estudiante_dict = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)
            
            archivo = generar_constancia_retiro(
                estudiante_dict,
                institucion,
                self.anio_escolar,
                motivo_retiro
            )
            
            crear_msgbox(
                self,
                "Éxito",
                f"Constancia de retiro generada:\n{archivo}",
                QMessageBox.Icon.Information
            ).exec()
            
            abrir_archivo(archivo)
            
        except Exception as e:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo generar la constancia: {e}",
                QMessageBox.Icon.Critical
            ).exec()

    def exportar_constancia_retiro_docx(self):
        """Genera constancia de retiro en formato DOCX."""
        # Verificar si el estudiante está inactivo
        if self.estudiante_actual.get("Estado", 1) == 1:
            crear_msgbox(
                self,
                "Estudiante activo",
                "La constancia de retiro solo se puede generar para estudiantes retirados (inactivos).\n\n"
                "Use el switch para marcar al estudiante como retirado primero.",
                QMessageBox.Icon.Warning
            ).exec()
            return

        try:
            # Obtener motivo de retiro desde la BD
            datos = EstudianteModel.obtener_por_id(self.id_estudiante)
            motivo_retiro = datos.get("motivo_retiro") if datos else None

            estudiante_dict = self.obtener_estudiante_actual_dict()
            institucion = InstitucionModel.obtener_por_id(1)

            archivo = generar_constancia_retiro_docx(
                estudiante_dict,
                institucion,
                self.anio_escolar,
                motivo_retiro
            )

            crear_msgbox(
                self,
                "Éxito",
                f"Constancia de retiro (DOCX) generada:\n{archivo}",
                QMessageBox.Icon.Information
            ).exec()

            abrir_archivo(archivo)

        except Exception as e:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo generar la constancia (DOCX): {e}",
                QMessageBox.Icon.Critical
            ).exec()

    def exportar_constancia_retiro_retirado(self):
        """Genera constancia de retiro mostrando el grado y año anterior al activo."""
        # Verificar que el estudiante esté retirado
        try:
            datos_bd = EstudianteModel.obtener_por_id(self.id_estudiante)
            if not datos_bd:
                crear_msgbox(self, "Error", "No se encontró el estudiante.",
                             QMessageBox.Icon.Critical).exec()
                return
            estatus_acad = str(datos_bd.get("estatus_academico", "")).strip()
            estado_est = datos_bd.get("estado", 1)
            if estatus_acad != "Retirado" and estado_est != 0:
                crear_msgbox(
                    self, "Estudiante no elegible",
                    "La constancia de retiro (grado anterior) solo se puede generar "
                    "para estudiantes con estatus académico 'Retirado'.",
                    QMessageBox.Icon.Warning
                ).exec()
                return

            # Obtener último grado cursado del historial (año anterior al activo)
            historial = EstudianteModel.obtener_historial_estudiante(self.id_estudiante)
            if not historial:
                crear_msgbox(self, "Sin historial",
                             "No se encontró historial académico para este estudiante.",
                             QMessageBox.Icon.Warning).exec()
                return
            anio_activo_inicio = int(self.anio_escolar['año_inicio'])
            ultimo_curso = None
            for registro in historial:
                if registro['año_inicio'] < anio_activo_inicio:
                    ultimo_curso = registro
                    break
            if not ultimo_curso:
                crear_msgbox(
                    self, "Sin historial previo",
                    "No se encontró un grado cursado y aprobado antes del año escolar activo.",
                    QMessageBox.Icon.Warning
                ).exec()
                return

            motivo_retiro = datos_bd.get("motivo_retiro")
            estudiante_dict = self.obtener_estudiante_actual_dict()
            estudiante_dict['Grado'] = ultimo_curso['grado']
            estudiante_dict['Sección'] = ultimo_curso['letra']
            estudiante_dict['Ciudad'] = datos_bd.get("ciudad", "")
            estudiante_dict['Fecha Nac.'] = datos_bd.get("fecha_nac", "")
            estudiante_dict['Género'] = datos_bd.get("genero", "")
            estudiante_dict['Tipo Educ.'] = datos_bd.get("tipo_educacion", "")

            institucion = InstitucionModel.obtener_por_id(1)
            archivo = generar_constancia_retiro_retirado(
                estudiante_dict, institucion, ultimo_curso['año_escolar'], motivo_retiro
            )

            crear_msgbox(
                self, "Éxito",
                f"Constancia de retiro (grado anterior) generada:\n{archivo}",
                QMessageBox.Icon.Information
            ).exec()
            abrir_archivo(archivo)

        except Exception as e:
            crear_msgbox(
                self, "Error",
                f"No se pudo generar la constancia: {e}",
                QMessageBox.Icon.Critical
            ).exec()

    def aplicar_sombras(self):
        """Aplica sombras a los elementos de la interfaz."""
        crear_sombra_flotante(self.btnModificar_ficha_estu)
        crear_sombra_flotante(self.btnExportar_ficha_estu)
        crear_sombra_flotante(self.btnEliminar_ficha_estu)
        crear_sombra_flotante(self.lneCedula_ficha_estu, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.stackFicha_estu, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.lblTitulo_ficha_estu, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.lblLogo_ficha_estu, blur_radius=8, y_offset=1)
        
        # Aplicar logo institucional dinámico
        aplicar_logo_a_label(self.lblLogo_ficha_estu)

    def cargar_secciones_en_combos(self):
        """Carga las secciones del año actual en los combos."""
        anio = self.anio_escolar['año_inicio']
        secciones = EstudianteModel.obtener_secciones_activas(anio)
        
        # Limpiar combos existentes
        self.cbxTipoEdu_ficha_estu.clear()
        self.cbxGrado_ficha_estu.clear()
        self.cbxSeccion_ficha_estu.clear()
        
        # Estructuras para organizar datos jerárquicamente
        niveles = set()
        self.grados_por_nivel = {}
        self.secciones_por_grado = {}
        
        # Procesar cada sección y organizarla
        for sec in secciones:
            nivel = sec["nivel"]
            grado = sec["grado"]
            letra = sec["letra"]
            
            # Agregar nivel único
            if nivel not in niveles:
                niveles.add(nivel)
                self.cbxTipoEdu_ficha_estu.addItem(nivel)
            
            # Agrupar grados por nivel
            if nivel not in self.grados_por_nivel:
                self.grados_por_nivel[nivel] = set()
            self.grados_por_nivel[nivel].add(grado)
            
            # Agrupar secciones por combinación nivel_grado
            clave = f"{nivel}_{grado}"
            if clave not in self.secciones_por_grado:
                self.secciones_por_grado[clave] = []
            self.secciones_por_grado[clave].append({
                "letra": letra,
                "id": sec["id"],
                "cupo_maximo": sec.get("cupo_maximo", 30) or 30,
                "estudiantes_actuales": sec.get("estudiantes_actuales", 0) or 0
            })
    
    def actualizar_grado(self):
        """Actualiza el combo de grados según el nivel seleccionado."""
        t_educacion = self.cbxTipoEdu_ficha_estu.currentText()
        
        # Limpiar combo de grado
        self.cbxGrado_ficha_estu.clear()
        
        if not t_educacion:
            self.cbxGrado_ficha_estu.setEnabled(False)
            self.cbxSeccion_ficha_estu.clear()
            self.cbxSeccion_ficha_estu.setEnabled(False)
            return
        
        # Cargar grados disponibles para este nivel
        grados = sorted(self.grados_por_nivel.get(t_educacion, set()))
        if grados:
            for g in grados:
                self.cbxGrado_ficha_estu.addItem(g)
            self.cbxGrado_ficha_estu.setEnabled(True)
            
            # Actualizar secciones si hay un grado seleccionado por defecto
            grado_actual = self.cbxGrado_ficha_estu.currentText()
            if grado_actual:
                self.actualizar_seccion(grado_actual)
        else:
            self.cbxGrado_ficha_estu.setEnabled(False)
            self.cbxSeccion_ficha_estu.clear()
            self.cbxSeccion_ficha_estu.setEnabled(False)
    
    def actualizar_seccion(self, grado):
        """Actualiza el combo de secciones según nivel y grado."""
        nivel = self.cbxTipoEdu_ficha_estu.currentText()
        if not nivel or not grado:
            self.cbxSeccion_ficha_estu.clear()
            self.cbxSeccion_ficha_estu.setEnabled(False)
            return
        
        # Buscar secciones para esta combinación nivel-grado
        clave = f"{nivel}_{grado}"
        self.cbxSeccion_ficha_estu.clear()
        opciones = self.secciones_por_grado.get(clave, [])
        
        if opciones:
            for opt in opciones:
                actuales = opt.get("estudiantes_actuales", 0)
                maximo = opt.get("cupo_maximo", 30)
                disponible = maximo - actuales
                
                if disponible > 0:
                    texto = f"{opt['letra']}  ({actuales}/{maximo})"
                    self.cbxSeccion_ficha_estu.addItem(texto, opt["id"])
                else:
                    texto = f"{opt['letra']}  (LLENA {actuales}/{maximo})"
                    self.cbxSeccion_ficha_estu.addItem(texto, opt["id"])
                    idx = self.cbxSeccion_ficha_estu.count() - 1
                    model = self.cbxSeccion_ficha_estu.model()
                    item = model.item(idx)
                    if item:
                        item.setEnabled(False)
                        item.setForeground(Qt.GlobalColor.gray)
            self.cbxSeccion_ficha_estu.setEnabled(True)
        else:
            self.cbxSeccion_ficha_estu.setEnabled(False)
    
    def cambiar_estado_estudiante(self, state):
        """Maneja el cambio de estado del estudiante."""
        # Evitar bucles infinitos durante actualizaciones
        if self.actualizando_switch:
            return
        
        self.actualizando_switch = True

        # Convertir estado del switch a valor de BD
        # Checked=2 (verde) -> Estado 1 (activo)
        # Unchecked=0 (gris) -> Estado 0 (inactivo)
        nuevo_estado = 1 if state == 2 else 0
        estado_actual = int(self.estudiante_actual.get("Estado", 1))
        
        # No hacer nada si el estado no cambió
        if nuevo_estado == estado_actual:
            self.actualizando_switch = False
            return

        # Texto para mensajes
        texto = "activar" if nuevo_estado else "desactivar"

        # Solicitar confirmación
        msg = crear_msgbox(
            self,
            "Confirmar acción",
            f"¿Seguro que deseas {texto} a este estudiante?",
            QMessageBox.Icon.Question,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if msg.exec() == QMessageBox.StandardButton.Yes:
            motivo_retiro = None
            
            # Si se está inactivando (retirando), pedir motivo
            if nuevo_estado == 0:
                from PySide6.QtWidgets import QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit as QLineEditDialog
                dialogo_motivo = QDialog(self)
                dialogo_motivo.setWindowTitle("Motivo de retiro")
                dialogo_motivo.setFixedSize(450, 160)
                dialogo_motivo.setStyleSheet("background-color: #f5f6fa;")
                
                layout_m = QVBoxLayout(dialogo_motivo)
                layout_m.setContentsMargins(20, 15, 20, 15)
                layout_m.setSpacing(10)
                
                lbl_m = QLabel("Ingrese el motivo del retiro del estudiante:\n(Dejar vacío para usar motivo predeterminado)")
                lbl_m.setWordWrap(True)
                layout_m.addWidget(lbl_m)
                
                txt_motivo = QLineEditDialog()
                txt_motivo.setFixedHeight(30)
                txt_motivo.setText("es retirado de la institución a solicitud de su representante siendo Promovido al siguiente grado")
                txt_motivo.setStyleSheet("""
                    QLineEdit {
                        border: 2px solid #2980b9;
                        border-radius: 10px;
                        padding: 5px 8px;
                        background-color: white;
                    }
                """)
                layout_m.addWidget(txt_motivo)
                
                h_btns_m = QHBoxLayout()
                btn_cancelar_m = QPushButton("Cancelar")
                btn_cancelar_m.setFixedHeight(32)
                btn_cancelar_m.setCursor(Qt.PointingHandCursor)
                btn_cancelar_m.setStyleSheet("""
                    QPushButton {
                        background-color: white;
                        color: #2980b9;
                        border: 1.5px solid #2980b9;
                        padding: 5px 15px;
                        border-radius: 10px;
                        font-weight: bold;
                    }
                    QPushButton:hover { background-color: #e3f2fd; }
                """)
                btn_cancelar_m.clicked.connect(dialogo_motivo.reject)
                
                btn_aceptar_m = QPushButton("Aceptar")
                btn_aceptar_m.setFixedHeight(32)
                btn_aceptar_m.setCursor(Qt.PointingHandCursor)
                btn_aceptar_m.setStyleSheet("""
                    QPushButton {
                        background-color: #2980b9;
                        color: white;
                        border: none;
                        padding: 5px 15px;
                        border-radius: 10px;
                        font-weight: bold;
                    }
                    QPushButton:hover { background-color: #0D47A1; }
                """)
                btn_aceptar_m.clicked.connect(dialogo_motivo.accept)
                
                h_btns_m.addWidget(btn_cancelar_m)
                h_btns_m.addWidget(btn_aceptar_m)
                layout_m.addLayout(h_btns_m)
                
                if dialogo_motivo.exec() == QDialog.DialogCode.Accepted:
                    motivo_retiro = txt_motivo.text().strip()
                    if not motivo_retiro:
                        motivo_retiro = "es retirado de la institución a solicitud de su representante siendo Promovido al siguiente grado"
                else:
                    motivo_retiro = "es retirado de la institución a solicitud de su representante siendo Promovido al siguiente grado"
            
            try:
                # Ejecutar cambio en BD
                base = RegistroBase()
                ok, mensaje = base.cambiar_estado(
                    "estudiantes", 
                    self.id, 
                    nuevo_estado, 
                    self.usuario_actual
                )

                if ok:
                    # Si fue retiro, guardar motivo y fecha
                    if nuevo_estado == 0 and motivo_retiro:
                        from utils.db import get_connection
                        from datetime import date
                        
                        conn = get_connection()
                        if conn:
                            try:
                                cursor = conn.cursor()
                                cursor.execute(
                                    "UPDATE estudiantes SET motivo_retiro = %s, fecha_retiro = %s WHERE id = %s",
                                    (motivo_retiro, date.today(), self.id)
                                )
                                conn.commit()
                                cursor.close()
                            except Exception as e:
                                print(f"Error guardando motivo de retiro: {e}")
                            finally:
                                if conn.is_connected():
                                    conn.close()
                    
                    # Actualizar estado local
                    self.estudiante_actual["Estado"] = nuevo_estado
                    self.lblEstado_ficha_estu.setText("Activo" if nuevo_estado else "Inactivo")
                    
                    # Emitir señal para actualizar tablas padre
                    self.datos_actualizados.emit()
                    
                    dlg = crear_msgbox(
                        self,
                        "Éxito",
                        f"Estudiante {texto}do correctamente.",
                        QMessageBox.Icon.Information,
                    )
                    dlg.exec()
                else:
                    dlg = crear_msgbox(
                        self,
                        "Error",
                        f"No se pudo {texto} al estudiante: {mensaje}",
                        QMessageBox.Icon.Critical,
                    )
                    dlg.exec()
                    self.revertir_switch()

            except Exception as e:
                print(f"Excepción al cambiar estado: {str(e)}")
                dlg = crear_msgbox(
                    self,
                    "Error",
                    f"Error inesperado: {str(e)}",
                    QMessageBox.Icon.Critical,
                )
                dlg.exec()
                self.revertir_switch()
        else:
            # Usuario canceló: revertir switch
            self.revertir_switch()
        
        self.actualizando_switch = False
    
    def revertir_switch(self):
        """Revierte el switch al estado anterior."""
        self.actualizando_switch = True
        estado = self.estudiante_actual.get("Estado", 1) == 1
        self.switchActivo.setChecked(estado)
        self.lblEstado_ficha_estu.setText("Activo" if estado else "Inactivo")
        self.actualizando_switch = False

    def cambiar_colaboracion(self, state):
        if self.actualizando_switch:
            return
        self.actualizando_switch = True
        colaboro = state == 2
        anio_actual = AnioEscolarModel.obtener_actual()
        if not anio_actual:
            crear_msgbox(self, "Error", "No se pudo obtener el año escolar actual.",
                         QMessageBox.Icon.Warning).exec()
            self.switchColaboracion.setChecked(not colaboro)
            self.actualizando_switch = False
            return

        ok, msg = ColaboracionModel.registrar(
            self.id_estudiante, anio_actual['id'], colaboro, self.usuario_actual
        )

        if not ok:
            crear_msgbox(self, "Error", msg, QMessageBox.Icon.Critical).exec()
            self.switchColaboracion.setChecked(not colaboro)

        self.actualizando_switch = False

    def actualizar_edad_estudiante(self):
        """Calcula y muestra la edad del estudiante."""
        fecha_nac = self.lneFechaNac_ficha_estu.date().toPython()
        
        # Validar que no sea fecha futura
        if fecha_nac > date.today():
            self.lneEdad_ficha_estu.setText("0")
            return
        
        edad = calcular_edad(fecha_nac)
        self.lneEdad_ficha_estu.setText(str(edad))

    def actualizar_edad_representante(self):
        """Calcula y muestra la edad del representante."""
        fecha_nac = self.lneFechaNac_repre_ficha_estu.date().toPython()
        
        # Validar que no sea fecha futura
        if fecha_nac > date.today():
            self.lneEdad_repre_ficha_estu.setText("0")
            return
        
        edad = calcular_edad(fecha_nac)
        self.lneEdad_repre_ficha_estu.setText(str(edad))

    def set_campos_editables(self, estado: bool):
        """Habilita/deshabilita la edición de campos."""
        # Campos editables del estudiante
        campos = [
            self.lneNombre_ficha_estu, self.lneApellido_ficha_estu, self.lneFechaNac_ficha_estu,
            self.lneCity_ficha_estu, self.cbxGenero_ficha_estu, self.lneDir_ficha_estu,
            self.cbxTipoEdu_ficha_estu, self.cbxGrado_ficha_estu,
            self.cbxSeccion_ficha_estu, self.lneTallaC_ficha_estu,
            self.lneTallaP_ficha_estu, self.lneTallaZ_ficha_estu,
            self.lneMadre_ficha_estu, self.lneOcup_madre_ficha_estu, self.lnePadre_ficha_estu,
            self.lneCedula_padre_ficha_estu, self.lneOcup_padre_ficha_estu, 
            self.lneApellidos_repre_ficha_estu, self.lneNombres_repre_ficha_estu, 
            self.lneFechaNac_repre_ficha_estu, self.cbxGenero_repre_ficha_estu, 
            self.lneDir_repre_ficha_estu, self.lneNum_repre_ficha_estu, self.lneFechaIng_ficha_estu,
            self.lneCorreo_repre_ficha_estu, self.lneObser_ficha_estu_repre,
        ]
        
        # Campos de solo lectura (calculados o inmutables)
        campos_solo_lectura = [
            self.lneEdad_ficha_estu, 
            self.lneEdad_repre_ficha_estu, 
            self.lneCedula_madre_ficha_estu,
            self.lneDocente_ficha_estu
        ]
        
        set_campos_editables(campos, estado, campos_solo_lectura)

    def cargar_datos(self):
        """Carga los datos del estudiante desde la BD."""
        datos = EstudianteModel.obtener_por_id(self.id)

        if not datos:
            msg = crear_msgbox(
                self,
                "Error",
                f"No se encontró el estudiante con ID {self.id}",
                QMessageBox.Icon.Critical
            )
            msg.exec()
            self.reject()
            return

        # Guardar estado actual del estudiante en memoria
        self.estudiante_actual = {
            "ID": self.id,
            "Cédula": str(datos["cedula"]) if datos else "",
            "Estado": int(datos.get("estado", 1)) if datos else 1
        }
        
        # Mostrar estado en el label
        estado_texto = "Activo" if self.estudiante_actual["Estado"] else "Inactivo"
        self.lblEstado_ficha_estu.setText(estado_texto)
        
        # --- DATOS PERSONALES DEL ESTUDIANTE ---
        self.lneCedula_ficha_estu.setText(str(datos["cedula"]))
        self.lneNombre_ficha_estu.setText(str(datos["nombres"]))
        self.lneApellido_ficha_estu.setText(str(datos["apellidos"]))

        # Fecha de nacimiento y edad
        fecha_est = datos["fecha_nac"]
        qdate_est = QDate(fecha_est.year, fecha_est.month, fecha_est.day)
        self.lneFechaNac_ficha_estu.setDate(qdate_est)
        self.lneEdad_ficha_estu.setText(str(calcular_edad(fecha_est)))
        
        # Otros datos personales
        self.lneCity_ficha_estu.setText(str(datos["ciudad"]))
        index_genero = self.cbxGenero_ficha_estu.findText(str(datos["genero"]))
        if index_genero >= 0:
            self.cbxGenero_ficha_estu.setCurrentIndex(index_genero)
        self.lneDir_ficha_estu.setText(str(datos["direccion"]))
        
        # Fecha de ingreso
        fecha_ing = datos["fecha_ingreso"]
        if fecha_ing:
            qdate_ing = QDate(fecha_ing.year, fecha_ing.month, fecha_ing.day)
            self.lneFechaIng_ficha_estu.setDate(qdate_ing)
        
        # --- DATOS ACADÉMICOS ---
        # Cargar según si es egresado o estudiante regular
        if self.es_egresado:
            # Datos de egresado (históricos)
            estatus = datos.get("estatus_academico", "Egresado")
            self.lneEstatus_egresado.setText(estatus)
            
            # Último grado cursado (incluye letra de sección)
            ultimo_grado = datos.get("ultimo_grado", "N/A")
            self.lneUltimoGrado.setText(ultimo_grado if ultimo_grado else "N/A")

            # Año escolar de egreso
            anio_egreso = datos.get("año_egreso", "N/A")
            self.lneAnioEgreso.setText(anio_egreso)
        else:
            # Estudiante regular: cargar sección actual
            tipo_edu = datos.get("tipo_educacion", "")
            if tipo_edu and tipo_edu != "Sin asignar":
                index_edu = self.cbxTipoEdu_ficha_estu.findText(tipo_edu)
                if index_edu >= 0:
                    self.cbxTipoEdu_ficha_estu.setCurrentIndex(index_edu)
                    self.actualizar_grado()
                    
                    grado = datos.get("grado", "")
                    if grado and grado != "Sin asignar":
                        index_grado = self.cbxGrado_ficha_estu.findText(grado)
                        if index_grado >= 0:
                            self.cbxGrado_ficha_estu.setCurrentIndex(index_grado)
                            self.actualizar_seccion(grado)
                            
                            seccion = datos.get("seccion", "")
                            if seccion and seccion != "Sin asignar":
                                # Buscar por inicio de texto (la letra) ya que el combo incluye info de cupo
                                index_seccion = self.cbxSeccion_ficha_estu.findText(
                                    seccion, Qt.MatchStartsWith
                                )
                                if index_seccion >= 0:
                                    self.cbxSeccion_ficha_estu.setCurrentIndex(index_seccion)
        
        docente_seccion = datos.get("docente_seccion", "Sin docente asignado")
        self.lneDocente_ficha_estu.setText(str(docente_seccion))
        
        # Tallas
        self.lneTallaC_ficha_estu.setText(str(datos["tallaC"]))
        self.lneTallaP_ficha_estu.setText(str(datos["tallaP"]))
        self.lneTallaZ_ficha_estu.setText(str(datos["tallaZ"]))
        
        # --- DATOS DE PADRES ---
        self.lnePadre_ficha_estu.setText(str(datos["padre"]))
        self.lneCedula_padre_ficha_estu.setText(str(datos["padre_ci"]))
        self.lneOcup_padre_ficha_estu.setText(str(datos["ocupacion_padre"]))
        self.lneMadre_ficha_estu.setText(str(datos["madre"]))
        self.lneCedula_madre_ficha_estu.setText(str(datos["madre_ci"]))
        self.lneOcup_madre_ficha_estu.setText(str(datos["ocupacion_madre"]))

        # --- DATOS DEL REPRESENTANTE ---
        representante_id = datos["representante_id"]
        if representante_id:
            datos_repre = RepresentanteModel.obtener_representante(representante_id)
            if datos_repre:
                self.lneCedula_repre_ficha_estu.setText(str(datos_repre["cedula"]))
                self.lneNombres_repre_ficha_estu.setText(str(datos_repre["nombres"]))
                self.lneApellidos_repre_ficha_estu.setText(str(datos_repre["apellidos"]))

                # Fecha de nacimiento del representante
                fecha_repre = datos_repre["fecha_nac"]
                qdate_repre = QDate(fecha_repre.year, fecha_repre.month, fecha_repre.day)
                self.lneFechaNac_repre_ficha_estu.setDate(qdate_repre)
                self.lneEdad_repre_ficha_estu.setText(str(calcular_edad(fecha_repre)))

                # Género del representante
                index_genero_repre = self.cbxGenero_repre_ficha_estu.findText(
                    str(datos_repre["genero"])
                )
                if index_genero_repre >= 0:
                    self.cbxGenero_repre_ficha_estu.setCurrentIndex(index_genero_repre)
                
                # Resto de datos del representante
                self.lneDir_repre_ficha_estu.setText(str(datos_repre["direccion"]))
                self.lneNum_repre_ficha_estu.setText(str(datos_repre["num_contact"]))
                self.lneCorreo_repre_ficha_estu.setText(str(datos_repre["email"]))
                self.lneObser_ficha_estu_repre.setText(str(datos_repre["observacion"]))

    def crear_y_vincular_representante(self, representante_data: dict) -> tuple:
        """Crea un nuevo representante y lo vincula al estudiante actual."""
        conexion = None
        cursor = None
        try:
            conexion = get_connection()
            if not conexion:
                return False, "Error de conexión a BD"
            cursor = conexion.cursor(dictionary=True)

            # Verificar si ya existe por cédula
            cursor.execute(
                "SELECT id FROM representantes WHERE cedula = %s",
                (representante_data["cedula"],)
            )
            row = cursor.fetchone()
            if row:
                representante_id = row["id"]
            else:
                cursor.execute("""
                    INSERT INTO representantes (
                        cedula, nombres, apellidos, fecha_nac,
                        genero, direccion, num_contact, email, observacion
                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    representante_data["cedula"],
                    representante_data["nombres"],
                    representante_data["apellidos"],
                    representante_data["fecha_nac"],
                    representante_data["genero"],
                    representante_data.get("direccion") or None,
                    representante_data.get("num_contact") or None,
                    representante_data.get("email") or None,
                    representante_data.get("observacion") or None,
                ))
                representante_id = cursor.lastrowid

            # Vincular al estudiante
            cursor.execute(
                "UPDATE estudiantes SET representante_id = %s WHERE id = %s",
                (representante_id, self.id)
            )
            conexion.commit()
            return True, "Representante creado y vinculado"
        except Exception as e:
            if conexion:
                conexion.rollback()
            return False, str(e)
        finally:
            if cursor:
                cursor.close()
            if conexion and conexion.is_connected():
                conexion.close()

    def _validar_texto_solo_letras(self, texto, nombre_campo):
        """Valida que el texto contenga solo letras."""
        if not texto:
            return False, ""

        # Validar patrón: solo letras (incluyendo acentos) y espacios
        if not re.match(r'^[A-Za-zÁÉÍÓÚÑáéíóúñ\s]+$', texto):
            crear_msgbox(
                self,
                "Formato inválido",
                f"El campo '{nombre_campo}' solo puede contener letras y espacios.",
                QMessageBox.Icon.Warning,
            ).exec()
            return False, ""

        # Normalizar: capitalizar cada palabra
        texto_normalizado = " ".join(p.capitalize() for p in texto.split())
        return True, texto_normalizado

    def _validar_email(self, email):
        """Valida formato de email."""
        if not email:
            return True  # Email opcional

        # Patrón básico de validación de email
        patron = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        if not re.match(patron, email):
            crear_msgbox(
                self,
                "Email inválido",
                "El formato del correo electrónico no es válido.",
                QMessageBox.Icon.Warning,
            ).exec()
            return False

        return True

    def _validar_telefono(self, telefono):
        """Valida formato de teléfono."""
        if not telefono:
            return True  # Teléfono opcional

        # Solo números y guiones permitidos
        if not re.match(r'^[\d\-]+$', telefono):
            crear_msgbox(
                self,
                "Teléfono inválido",
                "El teléfono solo puede contener números y guiones.",
                QMessageBox.Icon.Warning,
            ).exec()
            return False

        return True

    def guardar_datos(self):
        """Guarda los cambios en la BD."""
        try:
            # --- VALIDACIONES DE DATOS DEL ESTUDIANTE ---
            
            # Validar nombres y apellidos
            nombres = self.lneNombre_ficha_estu.text().strip()
            apellidos = self.lneApellido_ficha_estu.text().strip()
            
            valido_nombres, nombres_norm = self._validar_texto_solo_letras(
                nombres, "Nombres del estudiante"
            )
            valido_apellidos, apellidos_norm = self._validar_texto_solo_letras(
                apellidos, "Apellidos del estudiante"
            )
            
            if not valido_nombres or not valido_apellidos:
                return
            
            # Validar nombre de la madre (obligatorio)
            madre = self.lneMadre_ficha_estu.text().strip()
            valido_madre, madre_norm = self._validar_texto_solo_letras(
                madre, "Nombre de la madre"
            )
            if not valido_madre:
                return
            
            # Validar nombre del padre (opcional)
            padre = self.lnePadre_ficha_estu.text().strip()
            if padre:
                valido_padre, padre_norm = self._validar_texto_solo_letras(
                    padre, "Nombre del padre"
                )
                if not valido_padre:
                    return
            else:
                padre_norm = ""
            
            # Validar fecha de nacimiento
            fecha_nac = self.lneFechaNac_ficha_estu.date().toPython()
            if fecha_nac > date.today():
                crear_msgbox(
                    self,
                    "Fecha inválida",
                    "La fecha de nacimiento del estudiante no puede ser futura.",
                    QMessageBox.Icon.Warning,
                ).exec()
                return
            
            # --- RECOLECTAR DATOS DEL ESTUDIANTE ---
            estudiante_data = {
                "nombres": nombres_norm,
                "apellidos": apellidos_norm,
                "fecha_nac": fecha_nac,
                "ciudad": self.lneCity_ficha_estu.text().strip(),
                "genero": self.cbxGenero_ficha_estu.currentText().strip(),
                "direccion": self.lneDir_ficha_estu.text().strip(),
                "fecha_ingreso": self.lneFechaIng_ficha_estu.date().toPython(),
                "tipo_educacion": self.cbxTipoEdu_ficha_estu.currentText().strip(),
                "grado": self.cbxGrado_ficha_estu.currentText().strip(),
                "seccion": self.cbxSeccion_ficha_estu.currentText().split("  (")[0].strip(),
                "docente": self.lneDocente_ficha_estu.text().strip(),
                "tallaC": self.lneTallaC_ficha_estu.text().strip(),
                "tallaP": self.lneTallaP_ficha_estu.text().strip(),
                "tallaZ": self.lneTallaZ_ficha_estu.text().strip(),
                "padre": padre_norm,
                "padre_ci": self.lneCedula_padre_ficha_estu.text().strip(),
                "ocupacion_padre": self.lneOcup_padre_ficha_estu.text().strip(),
                "madre": madre_norm,
                "madre_ci": self.lneCedula_madre_ficha_estu.text().strip(),
                "ocupacion_madre": self.lneOcup_madre_ficha_estu.text().strip(),
            }
            
            # Obtener ID de sección si cambió
            seccion_id = self.cbxSeccion_ficha_estu.currentData()
            
            # --- VALIDACIONES DEL REPRESENTANTE (solo si hay datos) ---
            
            nombres_repre = self.lneNombres_repre_ficha_estu.text().strip()
            apellidos_repre = self.lneApellidos_repre_ficha_estu.text().strip()
            cedula_repre = self.lneCedula_repre_ficha_estu.text().strip()
            telefono = self.lneNum_repre_ficha_estu.text().strip()
            email = self.lneCorreo_repre_ficha_estu.text().strip()
            fecha_nac_repre = self.lneFechaNac_repre_ficha_estu.date().toPython()

            # Determinar si hay datos de representante para validar
            tiene_datos_repre = bool(nombres_repre or apellidos_repre or cedula_repre)
            
            nombres_repre_norm = ""
            apellidos_repre_norm = ""

            if tiene_datos_repre:
                valido_nombres_repre, nombres_repre_norm = self._validar_texto_solo_letras(
                    nombres_repre, "Nombres del representante"
                )
                valido_apellidos_repre, apellidos_repre_norm = self._validar_texto_solo_letras(
                    apellidos_repre, "Apellidos del representante"
                )
                
                if not valido_nombres_repre or not valido_apellidos_repre:
                    return
                
                if not cedula_repre or not cedula_repre.isdigit():
                    crear_msgbox(
                        self,
                        "Cédula inválida",
                        "La cédula del representante debe contener solo números.",
                        QMessageBox.Icon.Warning,
                    ).exec()
                    return

                if not self._validar_telefono(telefono):
                    return
                
                if not self._validar_email(email):
                    return
                
                if fecha_nac_repre > date.today():
                    crear_msgbox(
                        self,
                        "Fecha inválida",
                        "La fecha de nacimiento del representante no puede ser futura.",
                        QMessageBox.Icon.Warning,
                    ).exec()
                    return
            
            # --- ACTUALIZAR ESTUDIANTE ---
            ok_estudiante, msg_estudiante = EstudianteModel.actualizar(
                self.id, 
                estudiante_data, 
                self.usuario_actual,
                seccion_id
            )
            
            if not ok_estudiante:
                crear_msgbox(
                    self,
                    "Error",
                    f"No se pudo actualizar el estudiante:\n{msg_estudiante}",
                    QMessageBox.Icon.Critical,
                ).exec()
                return

            # --- ACTUALIZAR O CREAR REPRESENTANTE ---
            representante_id = RepresentanteModel.obtener_representante_id(self.id)

            if tiene_datos_repre:
                representante_data = {
                    "nombres": nombres_repre_norm,
                    "apellidos": apellidos_repre_norm,
                    "fecha_nac": fecha_nac_repre,
                    "genero": self.cbxGenero_repre_ficha_estu.currentText().strip(),
                    "direccion": self.lneDir_repre_ficha_estu.text().strip(),
                    "num_contact": telefono,
                    "email": email,
                    "observacion": self.lneObser_ficha_estu_repre.text().strip(),
                }

                if representante_id:
                    # Actualizar representante existente
                    ok_repre, msg_repre = RepresentanteModel.actualizar_representante(
                        representante_id, 
                        representante_data,
                        self.usuario_actual
                    )
                    
                    if not ok_repre:
                        crear_msgbox(
                            self,
                            "Advertencia",
                            f"Estudiante actualizado, pero hubo un problema con el representante:\n{msg_repre}",
                            QMessageBox.Icon.Warning,
                        ).exec()
                else:
                    # Crear nuevo representante y vincularlo al estudiante
                    representante_data["cedula"] = cedula_repre
                    ok_repre, msg_repre = self.crear_y_vincular_representante(
                        representante_data
                    )
                    if not ok_repre:
                        crear_msgbox(
                            self,
                            "Advertencia",
                            f"Estudiante actualizado, pero no se pudo crear el representante:\n{msg_repre}",
                            QMessageBox.Icon.Warning,
                        ).exec()

            # Éxito total
            crear_msgbox(
                self,
                "Éxito",
                "Datos actualizados correctamente.",
                QMessageBox.Icon.Information,
            ).exec()
            
            # Emitir señal para actualizar tablas padre
            self.datos_actualizados.emit()

        except Exception as err:
            msg = crear_msgbox(
                self,
                "Error",
                f"No se pudo guardar cambios: {err}",
                QMessageBox.Icon.Critical,
            )
            msg.exec()
    
    def toggle_edicion(self):
        """
        Alterna entre modo edición y guardado.
        En modo edición: habilita campos y cambia botón a "Guardar"
        En modo guardado: guarda cambios, deshabilita campos y cambia botón a "Modificar"
        """
        if self.btnModificar_ficha_estu.text() == "Modificar":
            # Entrar en modo edición
            self.set_campos_editables(True)
            self.btnModificar_ficha_estu.setText("Guardar")
        else:
            # Guardar cambios y salir de modo edición
            self.guardar_datos()
            self.set_campos_editables(False)
            self.btnModificar_ficha_estu.setText("Modificar")

    def eliminar_estudiante(self):
        """Elimina el registro completo del estudiante."""
        msg = crear_msgbox(
            self,
            "Confirmar eliminación",
            "¿Está seguro de eliminar este estudiante?\n\n"
            "Esta acción eliminará:\n"
            "- El registro del estudiante\n"
            "- Sus asignaciones a secciones\n"
            "- Su historial académico\n\n"
            "Esta acción NO se puede deshacer.",
            QMessageBox.Icon.Question,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if msg.exec() != QMessageBox.StandardButton.Yes:
            return

        try:
            ok, mensaje = EstudianteModel.eliminar(self.id_estudiante, self.usuario_actual)

            if ok:
                msg = crear_msgbox(
                    self,
                    "Éxito",
                    mensaje,
                    QMessageBox.Icon.Information,
                )
                msg.exec()
                
                # Emitir señal y cerrar ventana
                self.datos_actualizados.emit()
                self.accept()
            else:
                msg = crear_msgbox(
                    self,
                    "Error",
                    mensaje,
                    QMessageBox.Icon.Warning,
                )
                msg.exec()

        except Exception as err:
            msg = crear_msgbox(
                self,
                "Error",
                f"Error en la BD: {err}",
                QMessageBox.Icon.Critical,
            )
            msg.exec()

    def configurar_visibilidad_campos(self):
        """Configura qué campos mostrar según si es egresado o estudiante regular."""
        if self.es_egresado:
            # --- MODO EGRESADO (SOLO LECTURA) ---
            
            # Ocultar campos de estudiante regular
            self.lblTipoEdu.setVisible(False)
            self.frTipoEdu_reg_estu.setVisible(False)
            self.cbxTipoEdu_ficha_estu.setVisible(False)
            self.lblGrado.setVisible(False)
            self.frGrado_reg_estu.setVisible(False)
            self.cbxGrado_ficha_estu.setVisible(False)
            self.lblSeccion.setVisible(False)
            self.frseccion.setVisible(False)
            self.cbxSeccion_ficha_estu.setVisible(False)
            
            # Mostrar campos de egresado
            self.lblEstatus.setVisible(True)
            self.lneEstatus_egresado.setVisible(True)
            self.lblUltimoGrado.setVisible(True)
            self.lneUltimoGrado.setVisible(True)
            self.lblAnioEgreso.setVisible(True)
            self.lneAnioEgreso.setVisible(True)
            
            # Hacer campos de solo lectura
            self.lneEstatus_egresado.setReadOnly(True)
            self.lneUltimoGrado.setReadOnly(True)
            self.lneAnioEgreso.setReadOnly(True)
            
            # Ocultar controles de edición y estado
            self.btnModificar_ficha_estu.setVisible(False)
            self.btnEliminar_ficha_estu.setVisible(False)
            self.btnDevolver_grado.setVisible(False)
            self.switchActivo.setVisible(False)
            self.lblEstado_ficha_estu.setVisible(False)
            self.contenedorSwitchColaboracion.setVisible(False)
            if hasattr(self, 'label_50'):
                self.label_50.setVisible(False)
        else:
            # --- MODO REGULAR (EDITABLE) ---
            
            # Ocultar campos de egresado
            self.lblEstatus.setVisible(False)
            self.lneEstatus_egresado.setVisible(False)
            self.lblUltimoGrado.setVisible(False)
            self.lneUltimoGrado.setVisible(False)
            self.lblAnioEgreso.setVisible(False)
            self.lneAnioEgreso.setVisible(False)
            
            # Mostrar campos normales
            self.lblTipoEdu.setVisible(True)
            self.frTipoEdu_reg_estu.setVisible(True)
            self.cbxTipoEdu_ficha_estu.setVisible(True)
            self.lblGrado.setVisible(True)
            self.frGrado_reg_estu.setVisible(True)
            self.cbxGrado_ficha_estu.setVisible(True)
            self.lblSeccion.setVisible(True)
            self.frseccion.setVisible(True)
            self.cbxSeccion_ficha_estu.setVisible(True)
            
            # Mostrar controles normales
            self.btnModificar_ficha_estu.setVisible(True)
            self.btnEliminar_ficha_estu.setVisible(True)
            self.btnDevolver_grado.setVisible(True)
            self.switchActivo.setVisible(True) 
            self.lblEstado_ficha_estu.setVisible(True)
            self.contenedorSwitchColaboracion.setVisible(True)
            if hasattr(self, 'label_50'):
                self.label_50.setVisible(True)

    def _obtener_indice_grado(self, nivel, grado):
        """Retorna el índice numérico del grado para comparación.
        
        Inicial: 1er Nivel=0, 2do Nivel=1, 3er Nivel=2
        Primaria: 1ero=3, 2do=4, 3ero=5, 4to=6, 5to=7, 6to=8
        """
        if nivel == "Inicial":
            grados = SeccionesModel.GRADOS_INICIAL
            offset = 0
        else:
            grados = SeccionesModel.GRADOS_PRIMARIA
            offset = len(SeccionesModel.GRADOS_INICIAL)
        
        if grado in grados:
            return offset + grados.index(grado)
        return -1

    def devolver_estudiante(self):
        """
        Permite devolver al estudiante a un grado/sección anterior por repitencia.
        
        - Solo secciones del año actual
        - Solo grados iguales o inferiores al actual
        """
        # Validar que el estudiante esté activo
        if self.estudiante_actual.get("Estado") == 0:
            crear_msgbox(
                self,
                "Estudiante inactivo",
                "No se puede devolver un estudiante inactivo.\n"
                "Reactive al estudiante primero.",
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        # 1. Obtener secciones disponibles del año actual
        anio_actual = self.anio_escolar['año_inicio']
        secciones = EstudianteModel.obtener_secciones_activas(anio_actual)
        
        if not secciones:
            msg = crear_msgbox(
                self,
                "Sin secciones",
                "No hay secciones disponibles para devolver al estudiante.",
                QMessageBox.Icon.Warning
            )
            msg.exec()
            return
        
        # 2. Filtrar solo grados válidos (mismo grado o inferior al actual)
        nivel_actual = self.cbxTipoEdu_ficha_estu.currentText().strip()
        grado_actual = self.cbxGrado_ficha_estu.currentText().strip()
        indice_actual = self._obtener_indice_grado(nivel_actual, grado_actual)
        
        if indice_actual < 0:
            crear_msgbox(
                self,
                "Error",
                "No se pudo determinar el grado actual del estudiante.",
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        # 3. Crear lista de opciones filtradas (solo grados <= actual)
        opciones = []
        mapa_secciones = {}
        for sec in secciones:
            indice_sec = self._obtener_indice_grado(sec['nivel'], sec['grado'])
            if indice_sec >= 0 and indice_sec <= indice_actual:
                texto = f"{sec['grado']} {sec['letra']} - {sec['nivel']}"
                opciones.append(texto)
                mapa_secciones[texto] = sec['id']
        
        if not opciones:
            crear_msgbox(
                self,
                "Sin secciones",
                "No hay secciones de grado igual o inferior disponibles.",
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        # 4. Mostrar diálogo de selección
        dialogo = QDialog(self)
        dialogo.setWindowTitle("Devolver estudiante")
        dialogo.setFixedSize(400, 180)
        dialogo.setStyleSheet("background-color: #f5f6fa;")
        
        layout = QVBoxLayout(dialogo)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)
        
        lbl = QLabel("Seleccione la sección a la que desea devolver al estudiante\n(por repitencia o cambio de grado):")
        lbl.setWordWrap(True)
        layout.addWidget(lbl)
        
        cbx = QComboBox()
        cbx.setFixedHeight(30)
        cbx.setStyleSheet("""
            QComboBox {
                border: 2px solid #2980b9;
                border-radius: 10px;
                padding: 5px 8px;
                background-color: white;
            }
        """)
        cbx.addItems(opciones)
        layout.addWidget(cbx)
        
        layout.addStretch()
        
        h_btns = QHBoxLayout()
        btn_cancelar = QPushButton("Cancelar")
        btn_cancelar.setFixedHeight(32)
        btn_cancelar.setCursor(Qt.PointingHandCursor)
        btn_cancelar.setStyleSheet("""
            QPushButton {
                background-color: white;
                color: #2980b9;
                border: 1.5px solid #2980b9;
                padding: 5px 15px;
                border-radius: 10px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #e3f2fd; }
        """)
        btn_cancelar.clicked.connect(dialogo.reject)
        
        btn_aceptar = QPushButton("Aceptar")
        btn_aceptar.setFixedHeight(32)
        btn_aceptar.setCursor(Qt.PointingHandCursor)
        btn_aceptar.setStyleSheet("""
            QPushButton {
                background-color: #2980b9;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 10px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #0D47A1; }
        """)
        btn_aceptar.clicked.connect(dialogo.accept)
        
        h_btns.addWidget(btn_cancelar)
        h_btns.addWidget(btn_aceptar)
        layout.addLayout(h_btns)
        
        if dialogo.exec() != QDialog.DialogCode.Accepted:
            return
        
        seleccion = cbx.currentText()
        
        # 5. Confirmación explícita
        msg_confirm = crear_msgbox(
            self,
            "Confirmar devolución",
            f"¿Está seguro de devolver al estudiante a {seleccion}?\n\n"
            "Esto registrará en el historial que el estudiante repitió grado "
            "o fue reasignado a un nivel inferior.",
            QMessageBox.Icon.Question,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if msg_confirm.exec() != QMessageBox.StandardButton.Yes:
            return
        
        # 6. Ejecutar devolución
        try:
            seccion_id = mapa_secciones[seleccion]
            ok, mensaje = EstudianteModel.devolver_estudiante(
                self.id_estudiante,
                seccion_id,
                anio_actual,
                self.usuario_actual
            )
            
            if ok:
                msg = crear_msgbox(
                    self,
                    "Éxito",
                    mensaje,
                    QMessageBox.Icon.Information
                )
                msg.exec()
                
                # Recargar datos completos
                self.cargar_datos()
                self.cargar_historial()
                self.cargar_historial_notas()
                
                # Emitir señal para actualizar tablas padre
                self.datos_actualizados.emit()
            else:
                msg = crear_msgbox(
                    self,
                    "Error",
                    mensaje,
                    QMessageBox.Icon.Critical
                )
                msg.exec()

        except Exception as e:
            msg = crear_msgbox(
                self,
                "Error",
                f"Error inesperado: {e}",
                QMessageBox.Icon.Critical
            )
            msg.exec()

    def exportar_historial_pdf(self):
        """Exporta el historial académico del estudiante a PDF"""
        try:
            # Obtener historial desde el modelo
            historial = EstudianteModel.obtener_historial_estudiante(self.id_estudiante)
            
            # Obtener datos del estudiante
            estudiante_data = self.obtener_estudiante_actual_dict()
            
            # Obtener datos de la institución
            institucion = InstitucionModel.obtener_por_id(1)
            
            if not institucion:
                crear_msgbox(
                    self,
                    "Error",
                    "No se pudieron cargar los datos de la institución.",
                    QMessageBox.Icon.Warning
                ).exec()
                return
            
            # Generar PDF
            archivo = generar_historial_estudiante_pdf(estudiante_data, historial)
            
            crear_msgbox(
                self,
                "Éxito",
                f"Historial académico generado correctamente:\n{archivo}",
                QMessageBox.Icon.Information
            ).exec()
            
            abrir_archivo(archivo)
            
        except Exception as e:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo generar el historial académico:\n{e}",
                QMessageBox.Icon.Critical
            ).exec()
    
    def cargar_historial(self):
        """
        Carga y muestra el historial académico completo del estudiante.
        Muestra todos los años escolares cursados con sus respectivas secciones.
        """
        try:
            # Obtener historial desde el modelo
            historial = EstudianteModel.obtener_historial_estudiante(self.id_estudiante)
            
            if not historial:
                # Si no hay historial, mostrar tabla vacía
                model = QStandardItemModel(0, 5)
                model.setHorizontalHeaderLabels(["Año Escolar", "Nivel", "Grado", "Sección", "Docente"])
                self.tableW_historial.setModel(model)
                return
            
            # Configurar modelo con datos
            columnas = ["Año Escolar", "Nivel", "Grado", "Sección", "Docente"]
            model = QStandardItemModel(len(historial), len(columnas))
            model.setHorizontalHeaderLabels(columnas)
            
            # Llenar tabla con datos del historial
            for fila, registro in enumerate(historial):
                # Año escolar (formato: "2023-2024")
                anio_escolar = f"{registro['año_inicio']}-{registro['año_inicio']+1}"
                item_anio = QStandardItem(anio_escolar)
                item_anio.setEditable(False)
                model.setItem(fila, 0, item_anio)
                
                # Nivel educativo
                item_nivel = QStandardItem(str(registro['nivel']))
                item_nivel.setEditable(False)
                model.setItem(fila, 1, item_nivel)
                
                # Grado
                item_grado = QStandardItem(str(registro['grado']))
                item_grado.setEditable(False)
                model.setItem(fila, 2, item_grado)
                
                # Sección (letra)
                item_seccion = QStandardItem(str(registro['letra']))
                item_seccion.setEditable(False)
                model.setItem(fila, 3, item_seccion)
                
                # Docente
                item_docente = QStandardItem(str(registro['docente']))
                item_docente.setEditable(False)
                model.setItem(fila, 4, item_docente)
            
            # Asignar modelo a la tabla
            self.tableW_historial.setModel(model)
            self.tableW_historial.setSortingEnabled(True)
            self.tableW_historial.setAlternatingRowColors(True)
            
            # Ajustar anchos de columnas con tooltips si las columnas son muy estrechas
            anchos_historial = {
                0: 120,  # Año Escolar
                1: 150,  # Nivel
                2: 80,   # Grado
                3: 80,   # Sección
                4: 200   # Docente
            }
            ajustar_columnas_tabla(self, self.tableW_historial, anchos_historial)
            
        except Exception as e:
            print(f"Error al cargar historial: {e}")
            msg = crear_msgbox(
                self,
                "Error",
                f"No se pudo cargar el historial académico:\n{e}",
                QMessageBox.Icon.Warning
            )
            msg.exec()
    
    def cargar_historial_notas(self):
        """Carga y muestra el historial de notas por año escolar y lapsos."""
        
        try:
            # Obtener notas del estudiante
            notas = NotasModel.obtener_notas_estudiante(self.id_estudiante)
            
            if not notas:
                # Si no hay notas, mostrar tabla vacía
                model = QStandardItemModel(0, 9)
                model.setHorizontalHeaderLabels([
                    "Año Escolar", "Nivel", "Grado", "Sección", "Materia",
                    "L1", "L2", "L3", "Nota Final"
                ])
                self.tableW_historial_notas.setModel(model)
                return
            
            # Agrupar notas por año escolar
            notas_por_anio = {}
            for nota in notas:
                anio_key = f"{nota['año_inicio']}-{nota['año_inicio']+1}"
                if anio_key not in notas_por_anio:
                    notas_por_anio[anio_key] = []
                notas_por_anio[anio_key].append(nota)
            
            # Crear tabla con todas las notas
            columnas = ["Año Escolar", "Nivel", "Grado", "Sección", "Materia",
                       "Lapso 1", "Lapso 2", "Lapso 3", "Nota Final"]
            
            # Contar total de filas
            total_filas = sum(len(notas_anio) for notas_anio in notas_por_anio.values())
            
            model = QStandardItemModel(total_filas, len(columnas))
            model.setHorizontalHeaderLabels(columnas)
            
            # Llenar tabla
            fila_actual = 0
            for anio_escolar in sorted(notas_por_anio.keys()):
                notas_anio = notas_por_anio[anio_escolar]
                
                for nota in notas_anio:
                    # Año escolar
                    item_anio = QStandardItem(str(anio_escolar))
                    item_anio.setEditable(False)
                    model.setItem(fila_actual, 0, item_anio)
                    
                    # Nivel
                    item_nivel = QStandardItem(str(nota.get('nivel', '-')))
                    item_nivel.setEditable(False)
                    model.setItem(fila_actual, 1, item_nivel)
                    
                    # Grado
                    item_grado = QStandardItem(str(nota.get('grado', '-')))
                    item_grado.setEditable(False)
                    model.setItem(fila_actual, 2, item_grado)
                    
                    # Sección
                    item_seccion = QStandardItem(str(nota.get('letra', '-')))
                    item_seccion.setEditable(False)
                    model.setItem(fila_actual, 3, item_seccion)
                    
                    # Materia
                    item_materia = QStandardItem(str(nota.get('materia', '-')))
                    item_materia.setEditable(False)
                    model.setItem(fila_actual, 4, item_materia)
                    
                    # Lapso 1 (priorizar literal si existe)
                    lapso1_text = "-"
                    if nota.get('lapso_1_lit'):
                        lapso1_text = str(nota['lapso_1_lit'])
                    elif nota.get('lapso_1') is not None:
                        lapso1_text = str(nota['lapso_1'])
                    item_lapso1 = QStandardItem(lapso1_text)
                    item_lapso1.setEditable(False)
                    model.setItem(fila_actual, 5, item_lapso1)
                    
                    # Lapso 2
                    lapso2_text = "-"
                    if nota.get('lapso_2_lit'):
                        lapso2_text = str(nota['lapso_2_lit'])
                    elif nota.get('lapso_2') is not None:
                        lapso2_text = str(nota['lapso_2'])
                    item_lapso2 = QStandardItem(lapso2_text)
                    item_lapso2.setEditable(False)
                    model.setItem(fila_actual, 6, item_lapso2)
                    
                    # Lapso 3
                    lapso3_text = "-"
                    if nota.get('lapso_3_lit'):
                        lapso3_text = str(nota['lapso_3_lit'])
                    elif nota.get('lapso_3') is not None:
                        lapso3_text = str(nota['lapso_3'])
                    item_lapso3 = QStandardItem(lapso3_text)
                    item_lapso3.setEditable(False)
                    model.setItem(fila_actual, 7, item_lapso3)
                    
                    # Nota Final (priorizar literal)
                    nota_final_text = "-"
                    if nota.get('nota_final_literal'):
                        nota_final_text = str(nota['nota_final_literal'])
                    elif nota.get('nota_final') is not None:
                        nota_final_text = str(nota['nota_final'])
                    item_nota_final = QStandardItem(nota_final_text)
                    item_nota_final.setEditable(False)
                    model.setItem(fila_actual, 8, item_nota_final)
                    
                    fila_actual += 1
            
            # Asignar modelo a tabla
            self.tableW_historial_notas.setModel(model)
            self.tableW_historial_notas.setSortingEnabled(True)
            self.tableW_historial_notas.setAlternatingRowColors(True)
            
            # Ajustar anchos de columnas
            anchos_notas = {
                0: 120,  # Año Escolar
                1: 100,  # Nivel
                2: 80,   # Grado
                3: 80,   # Sección
                4: 120,  # Materia
                5: 70,   # Lapso 1
                6: 70,   # Lapso 2
                7: 70,   # Lapso 3
                8: 90    # Nota Final
            }
            ajustar_columnas_tabla(self, self.tableW_historial_notas, anchos_notas)
            
        except Exception as e:
            print(f"Error al cargar historial de notas: {e}")
            msg = crear_msgbox(
                self,
                "Error",
                f"No se pudo cargar el historial de notas:\n{e}",
                QMessageBox.Icon.Warning
            )
            msg.exec()
    
    def exportar_historial_notas_pdf(self):
        """Genera PDF con historial de notas del estudiante."""
        try:
            # Obtener datos del estudiante
            estudiante_data = {
                "Nombres": self.lneNombre_ficha_estu.text(),
                "Apellidos": self.lneApellido_ficha_estu.text(),
                "Cédula": self.lneCedula_ficha_estu.text()
            }
            
            # Obtener notas
            notas = NotasModel.obtener_notas_estudiante(self.id_estudiante)
            
            # Obtener datos de institución
            institucion = InstitucionModel.obtener_por_id(1)
            
            if not institucion:
                raise Exception("No se pudieron obtener los datos de la institución")
            
            # Generar PDF
            archivo = generar_historial_notas_pdf(estudiante_data, notas)
            
            # Abrir archivo
            msg = crear_msgbox(
                self,
                "Éxito",
                f"Historial de notas generado:\n{archivo}",
                QMessageBox.Icon.Information,
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Open,
                QMessageBox.StandardButton.Ok
            )
            
            if msg.exec() == QMessageBox.StandardButton.Open:
                abrir_archivo(archivo)
                
        except Exception as e:
            msg = crear_msgbox(
                self,
                "Error",
                f"No se pudo generar el historial de notas:\n{e}",
                QMessageBox.Icon.Critical
            )
            msg.exec()