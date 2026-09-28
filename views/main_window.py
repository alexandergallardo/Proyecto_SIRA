import os

from PySide6.QtWidgets import (
    QMainWindow, QToolButton, QMenu, QMessageBox,
    QSizePolicy, QLabel, QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QFrame, QFileDialog, QCompleter, QApplication
)
from PySide6.QtCore import QTimer, Qt, QSortFilterProxyModel, QSize, QStringListModel
from PySide6.QtGui import QStandardItem, QStandardItemModel, QAction, QIcon, QScreen
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtPdf import QPdfDocument

from models.dashboard_model import DashboardModel
from utils.exportar import (
    exportar_reporte_pdf,
    generar_constancia_estudios, generar_constancia_estudios_docx, generar_buena_conducta,
    generar_constancia_inscripcion, generar_constancia_prosecucion_inicial,
    generar_constancia_aceptacion,
    generar_buena_conducta_retirado,
    generar_constancia_trabajo, generar_constancia_retiro,
    generar_constancia_retiro_retirado,
    generar_historial_estudiante_pdf, generar_historial_notas_pdf,
    generar_certificado_promocion_sexto,
    generar_certificado_promocion_sexto_docx, generar_certificado_prosecucion_primaria,
    generar_certificado_prosecucion_primaria_docx,
    generar_buena_conducta_docx,
    generar_constancia_prosecucion_inicial_docx,
    generar_constancia_inscripcion_docx,
    generar_constancia_aceptacion_docx,
    generar_constancia_retiro_docx,
    generar_listado_colaboracion_seccion
)
from utils.backup import BackupManager
from models.estu_model import EstudianteModel
from models.emple_model import EmpleadoModel
from models.notas_model import NotasModel
from models.colaboracion_model import ColaboracionModel

from datetime import datetime

from views.delegates import UsuarioDelegate
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from utils.reportes_config import CriteriosReportes, criterios_por_poblacion
from views.registro_usuario import RegistroUsuario
from models.user_model import UsuarioModel
from views.actualizar_usuario import ActualizarUsuario
from models.auditoria_model import AuditoriaModel
from utils.forms import set_campos_editables, ajustar_columnas_tabla
from models.institucion_model import InstitucionModel
from models.anio_model import AnioEscolarModel
from views.gestion_estudiantes import GestionEstudiantesPage
from views.gestion_empleados import GestionEmpleadosPage
from views.gestion_secciones import GestionSeccionesPage
from views.gestion_anio import GestionAniosPage
from views.egresados import Egresados
from views.gestion_notas import GestionNotasPage
from views.gestion_materias import GestionMateriasPage
from views.acerca_de import AcercaDe
from utils.dialogs import crear_msgbox
from utils.sombras import crear_sombra_flotante
from utils.logo_manager import (
    aplicar_logo_a_label, obtener_logo_pixmap,
    procesar_imagen, invalidar_cache, LOGO_FILTRO_DIALOGO
)
from utils.archivos import abrir_archivo, abrir_carpeta
from utils.fecha_validacion import obtener_texto_advertencia
from paths import resource_path
from PySide6.QtWidgets import QListWidget, QListWidgetItem


def _seleccionar_ui_mainwindow():
    """Selecciona la UI principal y devuelve (clase_ui, variante)."""
    app = QApplication.instance()
    altura_disponible = None
    if app is not None:
        pantalla = app.primaryScreen()
        if pantalla is not None:
            altura_disponible = pantalla.availableGeometry().height()
            print("Altura disponible: ", altura_disponible)

    if altura_disponible is not None and altura_disponible <= 70:
        from ui_compiled.main_1024x600_ui import Ui_MainWindow
        print("UI 1024x600 cargada")
        return Ui_MainWindow, "1024x600"

    from ui_compiled.main_ui import Ui_MainWindow
    return Ui_MainWindow, "normal"


UiMainWindowBase, UI_MAINWINDOW_VARIANTE = _seleccionar_ui_mainwindow()


class MainWindow(QMainWindow, UiMainWindowBase):
    def __init__(self, usuario_actual, parent=None, resultado_fecha=None):
        super().__init__(parent)
        self.setupUi(self)
        self.ui_mainwindow_variante = UI_MAINWINDOW_VARIANTE
        
        self.usuario_actual = usuario_actual
        self.logout = False
        self.resultado_fecha = resultado_fecha
        
        # Inicializar lista de delegates para tooltips
        self.tooltip_delegates = []

        # Inicializar proxy de usuarios (se configura en database_usuarios)
        self.proxy_usuarios = None

        # Cargar año escolar actual
        self.anio_escolar = AnioEscolarModel.obtener_actual()
        if not self.anio_escolar:
            # Intentar crear año escolar por defecto
            ok, mensaje = AnioEscolarModel.inicializar_si_no_existe()
            if ok:
                self.anio_escolar = AnioEscolarModel.obtener_actual()
            else:
                crear_msgbox(
                    self,
                    "Error crítico",
                    f"No se pudo inicializar año escolar: {mensaje}\n\n"
                    "La aplicación puede no funcionar correctamente.",
                    QMessageBox.Icon.Critical
                ).exec()
                # Crear año ficticio para evitar crashes
                self.anio_escolar = {"id": 0, "nombre": "Sin año escolar"}
        
        self.setWindowTitle("SIRA - Sistema de Información para el Registro Académico")
        
        # Aplicar logo institucional dinámico a todos los labels del MainWindow
        for lbl in [
            self.lblLogo_dashboard_escuela,
            self.lblLogo_reportes,
            self.lblLogo_usuarios,
            self.lblLogo_auditoria,
            self.lblLogo_datos_insti,
            self.lblLogo_backup,
        ]:
            aplicar_logo_a_label(lbl)
        
        self.configurar_permisos()

        self.aplicar_sombras()

        self.lblBienvenida.setText(f"Bienvenido, {self.usuario_actual['username']}!")
        
        # Actualizar todos los labels de "Conectado como" en las diferentes páginas
        texto_conectado = f"Conectado como: {self.usuario_actual['username']}"
        self.lblConectado_como.setText(texto_conectado)
        self.lblConectado_como_2.setText(texto_conectado)
        self.lblConectado_como_3.setText(texto_conectado)
        self.lblConectado_como_4.setText(texto_conectado)
        self.lblConectado_como_5.setText(texto_conectado)
        
        self.btnUsuario_home.setText(f"{self.usuario_actual['username']}")

        # Obtener widgets placeholder
        placeholder_1 = self.stackMain.widget(1)
        placeholder_2 = self.stackMain.widget(2)
        placeholder_3 = self.stackMain.widget(3)
        placeholder_4 = self.stackMain.widget(4)
        placeholder_9 = self.stackMain.widget(9)

        # Crear páginas
        self.page_gestion_estudiantes = GestionEstudiantesPage(self.usuario_actual,self.anio_escolar,self,
                                                               ui_variant=self.ui_mainwindow_variante)
        self.page_egresados = Egresados(self.usuario_actual, self.anio_escolar, self, ui_variant=self.ui_mainwindow_variante)
        self.page_gestion_secciones = GestionSeccionesPage(self.usuario_actual, self.anio_escolar, self,
                                                           ui_variant=self.ui_mainwindow_variante)
        self.page_gestion_empleados = GestionEmpleadosPage(self.usuario_actual, self,
                                                           ui_variant=self.ui_mainwindow_variante)
        self.page_gestion_anios = GestionAniosPage(self.usuario_actual, self,
                                                   ui_variant=self.ui_mainwindow_variante)
        self.page_gestion_notas = GestionNotasPage(self.usuario_actual, self.anio_escolar, self,
                                                   ui_variant=self.ui_mainwindow_variante)
        self.page_gestion_materias = GestionMateriasPage(self.usuario_actual, self,
                                                         ui_variant=self.ui_mainwindow_variante)

        # Reemplazar placeholders
        self.stackMain.removeWidget(placeholder_1)
        self.stackMain.removeWidget(placeholder_2)
        self.stackMain.removeWidget(placeholder_3)
        self.stackMain.removeWidget(placeholder_4)
        self.stackMain.removeWidget(placeholder_9)
        self.stackMain.insertWidget(1, self.page_gestion_estudiantes)
        self.stackMain.insertWidget(2, self.page_gestion_secciones)
        self.stackMain.insertWidget(3, self.page_egresados)
        self.stackMain.insertWidget(4, self.page_gestion_empleados)
        self.stackMain.insertWidget(9, self.page_gestion_anios)
        
        # Agregar nuevas páginas de Notas y Materias
        self.stackMain.addWidget(self.page_gestion_notas)      # índice 11
        self.stackMain.addWidget(self.page_gestion_materias)   # índice 12

        # Configurar timer global
        self.timer_global = QTimer(self)
        self.timer_global.timeout.connect(self.actualizar_dashboard)
        self.timer_global.timeout.connect(self.cargar_auditoria)
        self.timer_global.timeout.connect(self.actualizar_widget_notificaciones)
        self.timer_global.start(60000)  # cada 60 segundos
        self.actualizar_dashboard()
        
        # Configurar timer para backups automáticos (cada 3 días)
        self.timer_backup = QTimer(self)
        self.timer_backup.timeout.connect(self.realizar_backup_automatico)
        # 3 días = 259200000 milisegundos (3 * 24 * 60 * 60 * 1000)
        self.timer_backup.start(259200000)
        # Cargar info del último backup
        self.cargar_info_backup()

        self.stackBarra_lateral.setCurrentIndex(0)
        self.stackMain.setCurrentIndex(0)
        
        # Widget de notificaciones
        if hasattr(self, 'frNotificaciones_home'):
            self.actualizar_widget_notificaciones()

        ## Botones barra lateral ##
        self.btnHome.clicked.connect(lambda: self.cambiar_pagina_main(0))
        self.btnEstudiantes.clicked.connect(lambda: self.cambiar_pagina_barra_lateral(1))
        self.btnGestion_estudiantes.clicked.connect(lambda: self.cambiar_pagina_main(1))
        self.btnSecciones.clicked.connect(lambda: self.cambiar_pagina_main(2))
        self.btnEgresados.clicked.connect(lambda: self.cambiar_pagina_main(3))
        self.btnNotas.clicked.connect(lambda: self.cambiar_pagina_main(11))
        self.btnEmpleados.clicked.connect(lambda: self.cambiar_pagina_main(4))
        self.btnReportes.clicked.connect(lambda: self.cambiar_pagina_main(5))
        self.btnAdmin.clicked.connect(lambda: self.cambiar_pagina_barra_lateral(2))
        self.btnRegresar_estudiantes.clicked.connect(lambda: self.cambiar_pagina_barra_lateral(0))
        self.btnRegresar_admin.clicked.connect(lambda: self.cambiar_pagina_barra_lateral(0))
        
        ## Botones de acceso directo ##
        self.btnAccesoDirecto_reg_estu.clicked.connect(self.acceso_directo_registro_estudiante)
        self.btnAccesoDirecto_reg_emple.clicked.connect(self.acceso_directo_registro_empleado)
        self.btnAccesoDirecto_secciones.clicked.connect(self.acceso_directo_crear_seccion)

        menu_usuario = QMenu(self)
        menu_usuario.setStyleSheet("""
            QMenu {
                background-color: white;
                color: black;
                border: 1px solid #c0c0c0;
            }
            QMenu::item {
                padding: 5px 20px 5px 10px;
            }
            QMenu::item:selected {
                background-color: #0078d7;
                color: white;
            }
            QMenu::icon {
                padding-left: 8px;
            }
        """)
        # Crear acciones con iconos personalizados
        accion_cerrar = QAction(QIcon(resource_path("resources/icons/logout.png")), "Cerrar sesión", self)
        accion_acerca_de = QAction(QIcon(resource_path("resources/icons/acerca_de.png")), "Acerca de SIRA", self)
        accion_manual_usuario = QAction(QIcon(resource_path("resources/icons/manual_usuario.png")), "Manual de usuario", self)
        accion_cerrar.triggered.connect(self.cerrar_sesion)
        accion_acerca_de.triggered.connect(self.mostrar_acerca_de)
        accion_manual_usuario.triggered.connect(self.mostrar_manual_usuario)
        menu_usuario.addAction(accion_cerrar)
        menu_usuario.addAction(accion_acerca_de)
        menu_usuario.addAction(accion_manual_usuario)
        self.btnUsuario_home.setMenu(menu_usuario)
        self.btnUsuario_home.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)

        ## MODULO REPORTES ##
        self.ultima_consulta = ([], [])
        self.figure = Figure()
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.canvas.updateGeometry()
        self.frGrafica_reportes.layout().addWidget(self.canvas)
        self.cbxPoblacion.currentIndexChanged.connect(self.actualizar_criterios)
        self.cbxCriterio.currentIndexChanged.connect(self.on_criterio_changed)
        self.btnGenerarGrafica.clicked.connect(self.on_generar)
        self.btnExportar_reporte.clicked.connect(self.on_exportar)

        # Estado inicial reportes
        self.lblMin.setVisible(False)
        self.lblMax.setVisible(False)
        self.spnMin.setVisible(False)
        self.spnMax.setVisible(False)
        self.cbxPoblacion.setEnabled(False)
        self.cbxCriterio.setEnabled(False)
        self.cbxTipoGrafica.setEnabled(False)

        # --- Configuración de constancias ---
        self._modo_reporte = "Estadístico"  # Modo activo: "Constancia", "RAC", "Estadístico"
        self._ruta_pdf_temporal = None
        self._persona_seleccionada = None  # dict con datos de la persona
        self._cache_personas = []  # Cache de personas para búsqueda

        # Guardar ítems originales de cbxPoblacion
        self._poblacion_items_original = [
            self.cbxPoblacion.itemText(i) for i in range(self.cbxPoblacion.count())
        ]

        # Constancias disponibles por población
        self._constancias_estudiantes = [
            "Constancia de estudios",
            "Constancia de estudios (DOCX)",
            "Constancia de inscripción",
            "Constancia de inscripción (DOCX)",
            "Constancia de aceptación",
            "Constancia de aceptación (DOCX)",
            "Constancia de buena conducta",
            "Constancia de buena conducta (DOCX)",
            "Constancia de buena conducta (retirado)",
            "Constancia de prosecución inicial",
            "Constancia de prosecución inicial (DOCX)",
            "Certif. de prosecución primaria",
            "Certif. de prosecución primaria (DOCX)",
            "Certif. promoción 6to a Secundaria",
            "Certif. promoción 6to a Secundaria (DOCX)",
            "Constancia de retiro (normal)",
            "Constancia de retiro (normal) (DOCX)",
            "Constancia de retiro (grado anterior)",
            "Historial académico",
            "Historial de notas",
        ]
        self._constancias_empleados = [
            "Constancia de trabajo",
        ]

        # Configurar QPdfView en la página de Constancias
        self.pdf_document = QPdfDocument(self)

        # Frame contenedor para el visor PDF
        self._pdf_frame = QFrame(self.Constancias)
        self._pdf_frame.setGeometry(10, 65, 931, 456)
        self._pdf_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border: 1.2px solid #2980b9;
                border-radius: 12px;
            }
        """)
        pdf_layout = QVBoxLayout(self._pdf_frame)
        pdf_layout.setContentsMargins(4, 4, 4, 4)

        self.pdf_viewer = QPdfView(self._pdf_frame)
        self.pdf_viewer.setObjectName("pdf_viewer")
        self.pdf_viewer.setDocument(self.pdf_document)
        self.pdf_viewer.setPageMode(QPdfView.PageMode.MultiPage)
        self.pdf_viewer.setZoomMode(QPdfView.ZoomMode.FitInView)
        # Fondo blanco en el área de visualización
        self.pdf_viewer.setStyleSheet("""
            QPdfView {
                background-color: #f0f0f0;
                border: none;
            }
            QPdfView > QWidget {
                background-color: #f0f0f0;
            }
        """)
        self.pdf_viewer.viewport().setStyleSheet("background-color: #f0f0f0;")
        pdf_layout.addWidget(self.pdf_viewer)

        # Configurar autocompletado para búsqueda de constancias
        self._completer_model = QStringListModel()
        self._completer = QCompleter(self._completer_model, self)
        self._completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self._completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self._completer.setMaxVisibleItems(10)
        self.lneBuscar_constancia.setCompleter(self._completer)
        self._completer.activated.connect(self._on_persona_seleccionada_completer)

        # Conectar señal de tipo de reporte
        self.cbxTipo_reporte.currentIndexChanged.connect(self.on_tipo_reporte_changed)

        # Ocultar controles de constancia inicialmente y el stacked hasta elegir tipo
        self.lneBuscar_constancia.setVisible(False)
        self.stackedReportes.setVisible(False)  # Oculto hasta que se elija tipo de reporte

        ## Botones Admin ##
        #--Gestion Usuarios--#
        self.btnGestion_usuarios.clicked.connect(lambda: self.cambiar_pagina_main(6))
        self.database_usuarios()
        self.btnCrear_usuario.clicked.connect(self.registro_usuario)
        self.btnActualizar_usuario.clicked.connect(self.actualizar_usuario)
        self.btnDisable_usuario.clicked.connect(self.cambiar_estado_usuario)
        
        self.chkMostrar_inactivos_user.stateChanged.connect(self.database_usuarios)
        
        #--Auditoria--#
        self.btnAuditoria.clicked.connect(lambda: self.cambiar_pagina_main(7))
        self.cargar_auditoria()
        
        #--Datos Institucionales--#
        self.btnDatos_institucion.clicked.connect(lambda: self.cambiar_pagina_main(8))
        self.set_campos_editables(False)
        self.cargar_datos_institucion()
        self.btnModificar_institucion.clicked.connect(self.toggle_edicion)
        
        # --- Logo institucional: preview y botones ---
        self.configurar_logo_institucional()

        #--Años escolares--#
        self.btnAnio_escolar.clicked.connect(lambda: self.cambiar_pagina_main(9))
        
        #--Gestión de Materias--#
        self.btnGestion_materias.clicked.connect(lambda: self.cambiar_pagina_main(12))
        
        #--Copia seguridad--#
        self.btnCopia_seguridad.clicked.connect(lambda: self.cambiar_pagina_main(10))
        self.btnBackup_manual.clicked.connect(self.realizar_backup_manual)
        self.btnRestablecer_backup.clicked.connect(self.restaurar_backup)

    def aplicar_sombras(self):
        crear_sombra_flotante(self.btnBackup_manual)
        crear_sombra_flotante(self.btnRestablecer_backup)
        crear_sombra_flotante(self.lblTitulo_backup, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.lblLogo_backup, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.lblTitulo_datos_insti, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.lblLogo_datos_insti, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.btnModificar_institucion)
        crear_sombra_flotante(self.frameInstitucion)
        crear_sombra_flotante(self.btnActualizar_tabla_auditoria)
        crear_sombra_flotante(self.frameTabla_auditoria)
        crear_sombra_flotante(self.lblTitulo_auditoria, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.lblLogo_auditoria, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.btnCrear_usuario)
        crear_sombra_flotante(self.btnActualizar_usuario)
        crear_sombra_flotante(self.btnDisable_usuario)
        crear_sombra_flotante(self.btnActualizar_tabla_user)
        crear_sombra_flotante(self.frameTabla_usuarios)
        crear_sombra_flotante(self.lblTitulo_usuarios, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.lblLogo_usuarios, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.frGrafica_border)
        crear_sombra_flotante(self.lblTitulo_reportes, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.lblLogo_reportes, blur_radius=5, y_offset=1)
        crear_sombra_flotante(self.btnGenerarGrafica)
        crear_sombra_flotante(self.btnExportar_reporte)
        crear_sombra_flotante(self.frameCriterio, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.framePoblacion, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.frameTipoGrafica, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.btnAccesoDirecto_reg_estu)
        crear_sombra_flotante(self.btnAccesoDirecto_reg_emple)
        crear_sombra_flotante(self.btnAccesoDirecto_secciones)
        crear_sombra_flotante(self.frMatricula_home)
        crear_sombra_flotante(self.frRepresentantes_home)
        crear_sombra_flotante(self.frTrabajadores_home)
        crear_sombra_flotante(self.frSeccion_home)
        crear_sombra_flotante(self.frNotificaciones_home)
        crear_sombra_flotante(self.frameSaludo, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.lblLogo_dashboard, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.lblLogo_dashboard_escuela, blur_radius=8, y_offset=1)
           
    def configurar_ventana_adaptable(self):
        """Configura la ventana para ser redimensionable."""
        # Obtener información de la pantalla
        screen = QScreen.availableGeometry(self.screen())
        screen_width = screen.width()
        screen_height = screen.height()
        
        # Calcular tamaño inicial (80% de la pantalla)
        initial_width = int(screen_width * 0.8)
        initial_height = int(screen_height * 0.8)
        
        # Establecer tamaño mínimo
        min_width = 1200
        min_height = 700
        self.setMinimumSize(QSize(min_width, min_height))
        
        # Establecer tamaño máximo (tamaño de la pantalla)
        self.setMaximumSize(QSize(screen_width, screen_height))
        
        # Establecer tamaño inicial
        self.resize(initial_width, initial_height)
        
        # Centrar la ventana en la pantalla
        self.move(
            screen.x() + (screen_width - initial_width) // 2,
            screen.y() + (screen_height - initial_height) // 2
        )
        
        # Permitir redimensionar
        self.setWindowFlags(self.windowFlags())
        
        print(f"Ventana configurada: {initial_width}x{initial_height} (min: {min_width}x{min_height})")

    def configurar_permisos(self):
        """Configura visibilidad según rol del usuario."""
        rol = self.usuario_actual.get("rol", "")
        if rol in "Administrador":
            self.btnAdmin.setVisible(True)
        else:
            self.btnAdmin.setVisible(False)

    def actualizar_dashboard(self):
        """Actualiza estadísticas del dashboard usando una sola conexión a BD."""
        try:
            # Una sola llamada para obtener todas las estadisticas
            datos = DashboardModel.obtener_todo_dashboard()
            
            self.lblMatricula_home.setText(str(datos['estudiantes'].get('activos', 0) or 0))
            self.lblRepresentantes_home.setText(str(datos['representantes_total']))
            self.lblEmpleados_home.setText(str(datos['empleados'].get('activos', 0) or 0))

            resultado = datos['seccion_mas_numerosa']
            if resultado:
                self.lblSeccion_home.setText(f"{resultado['grado']} {resultado['letra']}")
            else:
                self.lblSeccion_home.setText("Sin datos")
            
            # Actualizar conteos de usuarios
            if hasattr(self, 'lblActivos_usuarios'):
                self.lblActivos_usuarios.setText(str(datos['usuarios'].get('activos', 0) or 0))
            if hasattr(self, 'lblInactivos_usuarios'):
                self.lblInactivos_usuarios.setText(str(datos['usuarios'].get('inactivos', 0) or 0))
            
            # Actualizar conteos en módulos
            if hasattr(self, 'page_gestion_estudiantes'):
                est = datos['estudiantes']
                self.page_gestion_estudiantes.actualizar_conteo_desde_cache(
                    int(est.get('activos', 0) or 0),
                    int(est.get('inactivos', 0) or 0)
                )
            
            if hasattr(self, 'page_gestion_empleados'):
                emp = datos['empleados']
                self.page_gestion_empleados.actualizar_conteo_desde_cache(
                    int(emp.get('activos', 0) or 0),
                    int(emp.get('inactivos', 0) or 0)
                )
            
            # Guardar datos para notificaciones
            self._dashboard_cache = datos

        except Exception as err:
            print(f"Error en dashboard: {err}")
    
    def actualizar_widget_notificaciones(self):
        """Actualiza el widget de notificaciones usando datos cacheados del dashboard."""
        if not hasattr(self, 'lblNotificaciones_home'):
            return
        
        try:
            # Reutilizar datos del dashboard si están disponibles
            datos = getattr(self, '_dashboard_cache', None)
            if not datos:
                datos = DashboardModel.obtener_todo_dashboard()
            
            notificaciones = []
            
            # Estudiantes sin sección
            sin_seccion = datos.get('estudiantes_sin_seccion', 0)
            if sin_seccion > 0:
                icono = "⚠️" if sin_seccion > 5 else "📋"
                notificaciones.append(f"{icono} {sin_seccion} estudiante{'s' if sin_seccion != 1 else ''} sin sección")
            
            # Empleados sin código RAC
            sin_rac = datos.get('empleados_sin_rac', 0)
            if sin_rac > 0:
                notificaciones.append(f"📝 {sin_rac} empleado{'s' if sin_rac != 1 else ''} sin código RAC")
            
            # Secciones sin docente
            sin_docente = datos.get('secciones_sin_docente', 0)
            if sin_docente > 0:
                notificaciones.append(f"👨‍🏫 {sin_docente} sección{'es' if sin_docente != 1 else ''} sin docente asignado")
            
            # Secciones con cupo superado (prioridad alta)
            secciones_excedidas = datos.get('secciones_cupo_superado', [])
            if secciones_excedidas:
                n = len(secciones_excedidas)
                detalles = ", ".join(
                    f"{s['grado']} {s['letra']} ({s['actuales']}/{s['cupo_maximo']})"
                    for s in secciones_excedidas[:3]
                )
                if n > 3:
                    detalles += f" y {n - 3} más"
                notificaciones.append(f"🚨 Cupo superado: {detalles}")
            
            # Secciones de Primaria sin materias asignadas
            sin_materias = datos.get('secciones_sin_materias', 0)
            if sin_materias > 0:
                notificaciones.append(f"📖 {sin_materias} sección{'es' if sin_materias != 1 else ''} de Primaria sin materias asignadas")
            
            # Secciones vacías (activas sin estudiantes)
            secciones_vacias = datos.get('secciones_vacias', 0)
            if secciones_vacias > 0:
                notificaciones.append(f"🏫 {secciones_vacias} sección{'es' if secciones_vacias != 1 else ''} sin estudiantes")
            
            # Estudiantes retirados en el año actual
            retirados = datos.get('estudiantes_retirados_recientes', 0)
            if retirados > 0:
                notificaciones.append(f"🔴 {retirados} estudiante{'s' if retirados != 1 else ''} retirado{'s' if retirados != 1 else ''} este año")
            
            # Secciones con cupo disponible
            con_cupo = datos.get('secciones_con_cupo', 0)
            if con_cupo > 0:
                notificaciones.append(f"✅ {con_cupo} sección{'es' if con_cupo != 1 else ''} con cupo disponible")
            
            # Información del sistema
            info_sistema = []
            
            # Último backup
            try:
                ultimo_backup = BackupManager.obtener_ultimo_backup()
                if ultimo_backup:
                    dias_desde_backup = (datetime.now() - ultimo_backup['fecha']).days
                    if dias_desde_backup == 0:
                        info_sistema.append("💾 Backup: Hoy")
                    elif dias_desde_backup == 1:
                        info_sistema.append("💾 Backup: Ayer")
                    else:
                        icono_backup = "⚠️" if dias_desde_backup > 7 else "💾"
                        info_sistema.append(f"{icono_backup} Backup: Hace {dias_desde_backup} días")
                else:
                    info_sistema.append("⚠️ Sin backups")
            except Exception:
                pass
            
            # Año escolar actual
            if self.anio_escolar and self.anio_escolar.get('id', 0) > 0:
                info_sistema.append(f"📚 Año escolar activo: {self.anio_escolar['nombre']}")
            
            # Datos institucionales incompletos
            campos_incompletos = datos.get('datos_institucion_incompletos', [])
            if campos_incompletos:
                if len(campos_incompletos) <= 2:
                    detalle = ", ".join(campos_incompletos)
                    info_sistema.append(f"⚠️ Datos institucionales faltantes: {detalle}")
                else:
                    info_sistema.append(f"⚠️ {len(campos_incompletos)} datos institucionales incompletos")
            
            # Advertencia de fecha/hora del sistema
            advertencia_fecha = None
            if hasattr(self, 'resultado_fecha') and self.resultado_fecha:
                advertencia_fecha = obtener_texto_advertencia(self.resultado_fecha)
            
            # Construir texto final
            texto_final = ""
            
            # Mostrar advertencia de fecha primero (prioridad alta)
            if advertencia_fecha:
                texto_final = advertencia_fecha + "\n\n"
            
            if notificaciones:
                texto_final += "\n".join(notificaciones[:6])  # Máximo 6 notificaciones
            elif not advertencia_fecha:
                texto_final = "✅ Todo al día"
            
            # Agregar separador e info del sistema
            if info_sistema:
                texto_final += "\n\n" + "\n".join(info_sistema)
            
            self.lblNotificaciones_home.setText(texto_final)
            
        except Exception as e:
            print(f"Error actualizando notificaciones: {e}")
            if hasattr(self, 'lblNotificaciones_home'):
                self.lblNotificaciones_home.setText("❌ Error al cargar notificaciones")

    def actualizar_anio_escolar(self):
        """Actualiza el año escolar después de aperturar uno nuevo."""
        try:
            self.anio_escolar = AnioEscolarModel.obtener_actual()
            
            if self.anio_escolar:
                
                # Actualizar año en páginas hijas
                if hasattr(self, 'page_gestion_estudiantes'):
                    self.page_gestion_estudiantes.anio_escolar = self.anio_escolar
                    self.page_gestion_estudiantes.database_estudiantes()
                
                if hasattr(self, 'page_gestion_secciones'):
                    self.page_gestion_secciones.anio_escolar = self.anio_escolar
                    self.page_gestion_secciones.cargar_secciones()
                
                if hasattr(self, 'page_egresados'):
                    self.page_egresados.anio_escolar = self.anio_escolar
                
                # Actualizar página de notas
                if hasattr(self, 'page_gestion_notas'):
                    self.page_gestion_notas.anio_escolar = self.anio_escolar
                    self.page_gestion_notas.cargar_secciones()
                
                # Actualizar dashboard
                self.actualizar_dashboard()
                
        except Exception as e:
            print(f"Error actualizando año escolar: {e}")

    def cambiar_pagina_main(self, indice):
        """Cambia la página principal sin animación."""
        self.stackMain.setCurrentIndexInstant(indice)
        
    def cambiar_pagina_barra_lateral(self, indice):
        """Cambia la página de la barra lateral con efecto slide"""
        self.stackBarra_lateral.setCurrentIndexSlide(indice)
    
    ### MODULO REPORTES ###

    def on_tipo_reporte_changed(self):
        """Maneja el cambio de tipo de reporte (Constancia/RAC/Estadístico)."""
        tipo = self.cbxTipo_reporte.currentText()
        self._limpiar_pdf_temporal()
        self._persona_seleccionada = None
        self._cache_personas = []
        self.lneBuscar_constancia.clear()

        # Si está en el placeholder, ocultar y bloquear
        if self.cbxTipo_reporte.currentIndex() == 0:
            self._modo_reporte = ""
            self.stackedReportes.setVisible(False)
            self.cbxPoblacion.setEnabled(False)
            self.cbxCriterio.setEnabled(False)
            self.cbxTipoGrafica.setEnabled(False)
            self.lneBuscar_constancia.setVisible(False)
            self.frameCriterio.setVisible(True)
            self.lblCriterio.setVisible(True)
            return

        if tipo == "Constancia":
            self._modo_reporte = "Constancia"
            self.stackedReportes.setVisible(True)
            self.stackedReportes.setCurrentIndex(0)

            # Habilitar cbxPoblacion y adaptar a Estudiantes/Empleados
            self.cbxPoblacion.blockSignals(True)
            self.cbxPoblacion.clear()
            self.cbxPoblacion.addItems(["Seleccione población", "Estudiantes", "Empleados"])
            model_pob = self.cbxPoblacion.model()
            item0 = model_pob.item(0)
            if item0:
                item0.setEnabled(False)
                item0.setForeground(Qt.GlobalColor.gray)
            self.cbxPoblacion.setCurrentIndex(0)
            self.cbxPoblacion.setEnabled(True)
            self.cbxPoblacion.blockSignals(False)

            # Cambiar label de criterio y mostrar frameCriterio
            self.lblCriterio.setText("Constancia/Certificado")
            self.lblCriterio.setVisible(True)
            self.frameCriterio.setVisible(True)

            # Ocultar controles de estadísticos
            self.lblTipoGrafica.setVisible(False)
            self.frameTipoGrafica.setVisible(False)
            self.lblMin.setVisible(False)
            self.lblMax.setVisible(False)
            self.spnMin.setVisible(False)
            self.spnMax.setVisible(False)

            # Mostrar controles de constancia
            self.lneBuscar_constancia.setVisible(True)

            # Resetear criterio
            self.cbxCriterio.clear()
            self.cbxCriterio.setEnabled(False)

            # Forzar actualizar criterios
            self.actualizar_criterios()

        elif tipo == "Estadístico":
            self._modo_reporte = "Estadístico"
            self.stackedReportes.setVisible(True)
            self.stackedReportes.setCurrentIndex(1)

            # Restaurar cbxPoblacion original y habilitarlo
            self.cbxPoblacion.blockSignals(True)
            self.cbxPoblacion.clear()
            self.cbxPoblacion.addItems(self._poblacion_items_original)
            # Asegurar que Colaboración esté incluida
            if "Colaboración" not in self._poblacion_items_original:
                self.cbxPoblacion.addItem("Colaboración")
            model_pob = self.cbxPoblacion.model()
            item0 = model_pob.item(0)
            if item0:
                item0.setEnabled(False)
                item0.setForeground(Qt.GlobalColor.gray)
            self.cbxPoblacion.setCurrentIndex(0)
            self.cbxPoblacion.setEnabled(True)
            self.cbxPoblacion.blockSignals(False)

            # Restaurar labels y mostrar frameCriterio
            self.lblCriterio.setText("Criterio")
            self.lblCriterio.setVisible(True)
            self.frameCriterio.setVisible(True)
            self.lblTipoGrafica.setVisible(True)
            self.frameTipoGrafica.setVisible(True)

            # Ocultar controles de constancia
            self.lneBuscar_constancia.setVisible(False)

            # Resetear criterios
            self.cbxCriterio.clear()
            self.cbxCriterio.setEnabled(False)
            self.cbxTipoGrafica.clear()
            self.cbxTipoGrafica.setEnabled(False)

            # Limpiar gráfica
            self.figure.clear()
            self.canvas.draw()

        elif tipo == "RAC":
            self._modo_reporte = "RAC"
            # Mantener stacked oculto para RAC
            self.stackedReportes.setVisible(False)

            # cbxPoblacion fijo en Empleados
            self.cbxPoblacion.blockSignals(True)
            self.cbxPoblacion.clear()
            self.cbxPoblacion.addItems(["Empleados"])
            self.cbxPoblacion.setCurrentIndex(0)
            self.cbxPoblacion.setEnabled(False)
            self.cbxPoblacion.blockSignals(False)

            # Ocultar criterio
            self.frameCriterio.setVisible(False)
            self.lblCriterio.setVisible(False)
            self.cbxCriterio.clear()
            self.cbxCriterio.setEnabled(False)

            # Ocultar resto de controles
            self.lblTipoGrafica.setVisible(False)
            self.frameTipoGrafica.setVisible(False)
            self.lneBuscar_constancia.setVisible(False)
            self.lblMin.setVisible(False)
            self.lblMax.setVisible(False)
            self.spnMin.setVisible(False)
            self.spnMax.setVisible(False)

    def actualizar_criterios(self):
        """Actualiza los criterios disponibles según la población seleccionada"""
        poblacion = self.cbxPoblacion.currentText()

        # Limpiar y agregar placeholder
        self.cbxCriterio.clear()
        self.cbxCriterio.addItem("Seleccione un criterio" if self._modo_reporte == "Estadístico" else "Seleccione constancia")
        model = self.cbxCriterio.model()
        item0 = model.item(0)
        if item0 is not None:
            item0.setEnabled(False)
            item0.setForeground(Qt.GlobalColor.gray)

        if self._modo_reporte == "Constancia":
            # Modo constancia: cargar constancias según población
            if poblacion == "Estudiantes":
                self.cbxCriterio.addItems(self._constancias_estudiantes)
                self.cbxCriterio.setEnabled(True)
            elif poblacion == "Empleados":
                self.cbxCriterio.addItems(self._constancias_empleados)
                self.cbxCriterio.setEnabled(True)
            else:
                self.cbxCriterio.setEnabled(False)

            self.cbxCriterio.setCurrentIndex(0)

            # Limpiar búsqueda y persona seleccionada
            self.lneBuscar_constancia.clear()
            self._persona_seleccionada = None
            self._cache_personas = []
            self._limpiar_pdf_temporal()

            # Cargar cache de personas para búsqueda
            self._cargar_cache_personas()
            return

        # Modo estadístico: cargar criterios originales
        if poblacion in criterios_por_poblacion:
            self.cbxCriterio.addItems(criterios_por_poblacion[poblacion])
            self.cbxCriterio.setEnabled(True)
            self.cbxCriterio.setCurrentIndex(0)
        else:
            self.cbxCriterio.setEnabled(False)

        # Ocultar controles extra y limpiar gráfica
        self.lblMin.setVisible(False)
        self.lblMax.setVisible(False)
        self.spnMin.setVisible(False)
        self.spnMax.setVisible(False)
        self.cbxTipoGrafica.setEnabled(False)
        
        # Limpiar gráfica
        self.figure.clear()
        self.canvas.draw()

    def _cargar_cache_personas(self):
        """Carga la lista de personas según la población seleccionada para búsqueda."""
        poblacion = self.cbxPoblacion.currentText()
        self._cache_personas = []

        try:
            if poblacion == "Estudiantes":
                datos = EstudianteModel.listar(self.anio_escolar['id'])
                for d in datos:
                    self._cache_personas.append({
                        "id": d["id"],
                        "cedula": d.get("cedula", ""),
                        "nombres": d.get("nombres", ""),
                        "apellidos": d.get("apellidos", ""),
                        "display": f"{d.get('cedula', '')} - {d.get('nombres', '')} {d.get('apellidos', '')}",
                    })
            elif poblacion == "Empleados":
                datos = EmpleadoModel.listar_activos()
                for d in datos:
                    self._cache_personas.append({
                        "id": d["id"],
                        "cedula": d.get("cedula", ""),
                        "nombres": d.get("nombres", ""),
                        "apellidos": d.get("apellidos", ""),
                        "display": f"{d.get('cedula', '')} - {d.get('nombres', '')} {d.get('apellidos', '')}",
                    })

            # Actualizar completer
            items = [p["display"] for p in self._cache_personas]
            self._completer_model.setStringList(items)
        except Exception as e:
            print(f"Error cargando cache de personas: {e}")

    def _on_persona_seleccionada_completer(self, texto):
        """Maneja la selección de una persona desde el autocompletado."""
        for persona in self._cache_personas:
            if persona["display"] == texto:
                self._persona_seleccionada = persona
                return
        self._persona_seleccionada = None

    def on_criterio_changed(self):
        """Maneja el cambio de criterio mostrando controles específicos"""
        if self._modo_reporte == "Constancia":
            # En modo constancia no se necesitan controles extras
            self._limpiar_pdf_temporal()
            self._persona_seleccionada = None
            self.lneBuscar_constancia.clear()
            return

        idx = self.cbxCriterio.currentIndex()
        criterio = self.cbxCriterio.currentText() if idx > 0 else ""

        # Ocultar todos los controles primero
        self.lblMin.setVisible(False)
        self.lblMax.setVisible(False)
        self.spnMin.setVisible(False)
        self.spnMax.setVisible(False)
        
        if hasattr(self, 'cbxSeccion_reporte'):
            self.cbxSeccion_reporte.setVisible(False)
            if hasattr(self, 'frameSeccion_reporte'):
                self.frameSeccion_reporte.setVisible(False)
            if hasattr(self, 'lblSeccion_reporte'):
                self.lblSeccion_reporte.setVisible(False)

        # Mostrar controles según criterio
        if criterio == "Rango de edad":
            self.lblMin.setText("Edad mínima")
            self.lblMax.setText("Edad máxima")
            self.lblMin.setVisible(True)
            self.lblMax.setVisible(True)
            self.spnMin.setVisible(True)
            self.spnMax.setVisible(True)
            self.spnMin.setEnabled(True)
            self.spnMax.setEnabled(True)
            self.spnMin.setMinimum(0)
            self.spnMin.setMaximum(100)
            self.spnMax.setMinimum(0)
            self.spnMax.setMaximum(100)
            self.spnMin.setValue(3)
            self.spnMax.setValue(18)

        elif criterio == "Rango de salario":
            self.lblMin.setText("Salario mínimo")
            self.lblMax.setText("Salario máximo")
            self.lblMin.setVisible(True)
            self.lblMax.setVisible(True)
            self.spnMin.setVisible(True)
            self.spnMax.setVisible(True)
            self.spnMin.setEnabled(True)
            self.spnMax.setEnabled(True)
            self.spnMin.setMinimum(0)
            self.spnMin.setMaximum(999999)
            self.spnMax.setMinimum(0)
            self.spnMax.setMaximum(999999)
            self.spnMin.setValue(100)
            self.spnMax.setValue(5000)

        elif criterio == "Matricula por año escolar":
            self.lblMin.setText("Año inicio")
            self.lblMax.setText("Año fin")
            self.lblMin.setVisible(True)
            self.lblMax.setVisible(True)
            self.spnMin.setVisible(True)
            self.spnMax.setVisible(True)
            self.spnMin.setEnabled(True)
            self.spnMax.setEnabled(True)
            self.spnMin.setMinimum(2000)
            self.spnMin.setMaximum(2100)
            self.spnMax.setMinimum(2000)
            self.spnMax.setMaximum(2100)
            anio_actual = datetime.now().year
            self.spnMin.setValue(anio_actual - 5)
            self.spnMax.setValue(anio_actual)

        elif criterio == "Género por sección específica":
            if hasattr(self, 'cbxSeccion_reporte'):
                secciones = CriteriosReportes.obtener_secciones_activas()
                self.cbxSeccion_reporte.clear()
                self.cbxSeccion_reporte.addItem("Seleccione una sección")
                self.cbxSeccion_reporte.addItems(secciones)
                self.cbxSeccion_reporte.setVisible(True)
                if hasattr(self, 'frameSeccion_reporte'):
                    self.frameSeccion_reporte.setVisible(True)
                if hasattr(self, 'lblSeccion_reporte'):
                    self.lblSeccion_reporte.setVisible(True)

        elif criterio == "Listado por sección":
            if hasattr(self, 'cbxSeccion_reporte'):
                secciones = CriteriosReportes.obtener_secciones_activas()
                self.cbxSeccion_reporte.clear()
                self.cbxSeccion_reporte.addItem("Seleccione una sección")
                self.cbxSeccion_reporte.addItems(secciones)
                self.cbxSeccion_reporte.setVisible(True)
                self.cbxTipoGrafica.setEnabled(False)
                self.cbxTipoGrafica.setVisible(False)
                self.frameTipoGrafica.setVisible(False)
                self.lblTipoGrafica.setVisible(False)
                if hasattr(self, 'frameSeccion_reporte'):
                    self.frameSeccion_reporte.setVisible(True)
                if hasattr(self, 'lblSeccion_reporte'):
                    self.lblSeccion_reporte.setVisible(True)

        # Configurar tipos de gráfica
        if idx > 0:
            self.actualizar_tipos_grafica()
        else:
            self.cbxTipoGrafica.clear()
            self.cbxTipoGrafica.setEnabled(False)
        
        # Limpiar gráfica al cambiar criterio
        self.figure.clear()
        self.canvas.draw()

    def actualizar_tipos_grafica(self):
        """Actualiza los tipos de gráfica disponibles"""
        self.cbxTipoGrafica.clear()
        self.cbxTipoGrafica.addItem("Seleccione un tipo de gráfica")
        model = self.cbxTipoGrafica.model()
        item0 = model.item(0)
        if item0 is not None:
            item0.setEnabled(False)
            item0.setForeground(Qt.GlobalColor.gray)

        tipos = list(CriteriosReportes.GRAFICAS.keys())
        self.cbxTipoGrafica.addItems(tipos)
        self.cbxTipoGrafica.setEnabled(True)
        self.cbxTipoGrafica.setCurrentIndex(0)

    # --- Dispatcher de Generar / Exportar ---

    def on_generar(self):
        """Dispatcher: genera según el modo activo."""
        if self._modo_reporte == "Constancia":
            self._generar_constancia()
        elif self._modo_reporte == "RAC":
            self._generar_rac()
        elif self.cbxPoblacion.currentText() == "Colaboración" and self.cbxCriterio.currentText() == "Listado por sección":
            self._generar_listado_colaboracion()
        else:
            self.actualizar_reporte()

    def on_exportar(self):
        """Dispatcher: exporta según el modo activo."""
        if self._modo_reporte == "Constancia":
            self._exportar_constancia()
        elif self._modo_reporte == "RAC":
            self._generar_rac()
        elif self.cbxPoblacion.currentText() == "Colaboración" and self.cbxCriterio.currentText() == "Listado por sección":
            self._exportar_listado_colaboracion()
        else:
            self.on_exportar_reporte()

    # --- Constancias ---

    def _construir_dict_estudiante(self, datos_bd):
        """Construye el dict con las claves esperadas por las funciones de exportar.py"""
        return {
            "ID": str(datos_bd.get("id", "")),
            "Cédula": datos_bd.get("cedula", ""),
            "Nombres": datos_bd.get("nombres", ""),
            "Apellidos": datos_bd.get("apellidos", ""),
            "Fecha Nac.": datos_bd.get("fecha_nac", ""),
            "Edad": str(datos_bd.get("edad", "")),
            "Ciudad": datos_bd.get("ciudad", ""),
            "Género": datos_bd.get("genero", ""),
            "Dirección": datos_bd.get("direccion", ""),
            "Fecha Ingreso": datos_bd.get("fecha_ingreso", ""),
            "Tipo Educ.": datos_bd.get("tipo_educacion", ""),
            "Grado": datos_bd.get("grado", ""),
            "Sección": datos_bd.get("seccion", ""),
            "Docente": datos_bd.get("docente_seccion", ""),
        }

    def _construir_dict_empleado(self, datos_bd):
        """Construye el dict con las claves esperadas por generar_constancia_trabajo."""
        return {
            "ID": str(datos_bd.get("id", "")),
            "Cédula": datos_bd.get("cedula", ""),
            "Nombres": datos_bd.get("nombres", ""),
            "Apellidos": datos_bd.get("apellidos", ""),
            "Cargo": datos_bd.get("cargo", ""),
            "Fecha Ingreso": datos_bd.get("fecha_ingreso", ""),
        }

    def _generar_constancia(self):
        """Genera la constancia seleccionada y la muestra en el QPdfView."""
        poblacion = self.cbxPoblacion.currentText()
        constancia = self.cbxCriterio.currentText()
        idx_constancia = self.cbxCriterio.currentIndex()

        if not poblacion or poblacion == "Seleccione población":
            crear_msgbox(self, "Atención", "Debe seleccionar una población.", QMessageBox.Icon.Warning).exec()
            return

        if idx_constancia <= 0:
            crear_msgbox(self, "Atención", "Debe seleccionar un tipo de constancia.", QMessageBox.Icon.Warning).exec()
            return

        if not self._persona_seleccionada:
            crear_msgbox(self, "Atención", "Debe buscar y seleccionar una persona.", QMessageBox.Icon.Warning).exec()
            return

        persona_id = self._persona_seleccionada["id"]

        try:
            institucion = InstitucionModel.obtener_por_id(1)
            if not institucion:
                crear_msgbox(self, "Error", "No se pudieron cargar los datos de la institución.", QMessageBox.Icon.Critical).exec()
                return

            archivo = None

            if poblacion == "Estudiantes":
                datos_bd = EstudianteModel.obtener_por_id(persona_id)
                if not datos_bd:
                    crear_msgbox(self, "Error", "No se encontró el estudiante.", QMessageBox.Icon.Critical).exec()
                    return

                estudiante = self._construir_dict_estudiante(datos_bd)

                if constancia == "Constancia de estudios":
                    archivo = generar_constancia_estudios(estudiante, institucion)

                elif constancia == "Constancia de estudios (DOCX)":
                    archivo = generar_constancia_estudios_docx(estudiante, institucion)

                elif constancia == "Constancia de inscripción":
                    archivo = generar_constancia_inscripcion(estudiante, institucion)

                elif constancia == "Constancia de inscripción (DOCX)":
                    archivo = generar_constancia_inscripcion_docx(estudiante, institucion)

                elif constancia == "Constancia de aceptación":
                    archivo = generar_constancia_aceptacion(institucion ,self.anio_escolar)

                elif constancia == "Constancia de aceptación (DOCX)":
                    archivo = generar_constancia_aceptacion_docx(institucion, self.anio_escolar)

                elif constancia == "Constancia de buena conducta":
                    archivo = generar_buena_conducta(estudiante, institucion, self.anio_escolar)

                elif constancia == "Constancia de buena conducta (DOCX)":
                    archivo = generar_buena_conducta_docx(estudiante, institucion, self.anio_escolar)

                elif constancia == "Constancia de buena conducta (retirado)":
                    # Solo para estudiantes retirados
                    estatus_acad = str(datos_bd.get("estatus_academico", "")).strip()
                    estado_est = datos_bd.get("estado", 1)
                    if estatus_acad != "Retirado" and estado_est != 0:
                        crear_msgbox(
                            self,
                            "Estudiante no elegible",
                            "La constancia de buena conducta (retirado) solo se puede generar "
                            "para estudiantes con estatus académico 'Retirado'.",
                            QMessageBox.Icon.Warning
                        ).exec()
                        return
                    historial = EstudianteModel.obtener_historial_estudiante(persona_id)
                    if not historial:
                        crear_msgbox(
                            self, "Sin historial",
                            "No se encontró historial académico para este estudiante.",
                            QMessageBox.Icon.Warning
                        ).exec()
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
                    estudiante['Grado'] = ultimo_curso['grado']
                    estudiante['Sección'] = ultimo_curso['letra']
                    archivo = generar_buena_conducta_retirado(
                        estudiante, institucion, ultimo_curso['año_escolar']
                    )
                
                elif constancia == "Certif. de prosecución primaria":
                    archivo = generar_certificado_prosecucion_primaria(estudiante, institucion, self.anio_escolar)

                elif constancia == "Certif. de prosecución primaria (DOCX)":
                    archivo = generar_certificado_prosecucion_primaria_docx(estudiante, institucion, self.anio_escolar)

                elif constancia == "Constancia de prosecución inicial":
                    tipo_actual = str(datos_bd.get("tipo_educacion", "")).strip().lower()
                    grado_actual = str(datos_bd.get("grado", "")).strip().lower()
                    anio_inicio_actual = int(self.anio_escolar['año_inicio'])

                    anio_escolar_inicial = None

                    if tipo_actual in ['inicial', 'preescolar'] and '3' in grado_actual:
                        anio_escolar_inicial = {
                            'año_inicio': anio_inicio_actual,
                            'año_fin': anio_inicio_actual + 1
                        }
                    elif tipo_actual == 'primaria':
                        historial = EstudianteModel.obtener_historial_estudiante(persona_id)
                        if not historial:
                            crear_msgbox(self, "Sin historial", "No hay historial académico.", QMessageBox.Icon.Warning).exec()
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

                    archivo = generar_constancia_prosecucion_inicial(estudiante, institucion, anio_escolar_inicial)

                elif constancia == "Constancia de prosecución inicial (DOCX)":
                    tipo_actual = str(datos_bd.get("tipo_educacion", "")).strip().lower()
                    grado_actual = str(datos_bd.get("grado", "")).strip().lower()
                    anio_inicio_actual = int(self.anio_escolar['año_inicio'])

                    anio_escolar_inicial = None

                    if tipo_actual in ['inicial', 'preescolar'] and '3' in grado_actual:
                        anio_escolar_inicial = {
                            'año_inicio': anio_inicio_actual,
                            'año_fin': anio_inicio_actual + 1
                        }
                    elif tipo_actual == 'primaria':
                        historial = EstudianteModel.obtener_historial_estudiante(persona_id)
                        if not historial:
                            crear_msgbox(self, "Sin historial", "No hay historial académico.", QMessageBox.Icon.Warning).exec()
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

                    archivo = generar_constancia_prosecucion_inicial_docx(estudiante, institucion, anio_escolar_inicial)

                elif constancia == "Certif. promoción 6to a Secundaria":
                    tipo_actual = str(datos_bd.get("tipo_educacion", "")).strip().lower()
                    grado_actual = str(datos_bd.get("grado", "")).strip().lower()

                    if tipo_actual != 'primaria' or '6' not in grado_actual:
                        crear_msgbox(
                            self,
                            "Estudiante no elegible",
                            "Este certificado solo se puede generar para estudiantes que cursan 6to grado.",
                            QMessageBox.Icon.Warning
                        ).exec()
                        return

                    anio_inicio = int(self.anio_escolar['año_inicio'])
                    estudiante['ultima_seccion'] = estudiante.get('Sección', 'N/A')
                    archivo = generar_certificado_promocion_sexto(
                        estudiante,
                        institucion,
                        f"{anio_inicio}/{anio_inicio + 1}"
                    )

                elif constancia == "Certif. promoción 6to a Secundaria (DOCX)":
                    tipo_actual = str(datos_bd.get("tipo_educacion", "")).strip().lower()
                    grado_actual = str(datos_bd.get("grado", "")).strip().lower()

                    if tipo_actual != 'primaria' or '6' not in grado_actual:
                        crear_msgbox(
                            self,
                            "Estudiante no elegible",
                            "Este certificado solo se puede generar para estudiantes que cursan 6to grado.",
                            QMessageBox.Icon.Warning
                        ).exec()
                        return

                    anio_inicio = int(self.anio_escolar['año_inicio'])
                    estudiante['ultima_seccion'] = estudiante.get('Sección', 'N/A')
                    archivo = generar_certificado_promocion_sexto_docx(
                        estudiante,
                        institucion,
                        f"{anio_inicio}/{anio_inicio + 1}"
                    )

                elif constancia == "Constancia de retiro (normal)":
                    if datos_bd.get("estado", 1) == 1:
                        crear_msgbox(self, "Estudiante activo",
                                     "La constancia de retiro solo se puede generar para estudiantes retirados (inactivos).",
                                     QMessageBox.Icon.Warning).exec()
                        return
                    motivo_retiro = datos_bd.get("motivo_retiro")
                    archivo = generar_constancia_retiro(estudiante, institucion, self.anio_escolar, motivo_retiro)

                elif constancia == "Constancia de retiro (normal) (DOCX)":
                    if datos_bd.get("estado", 1) == 1:
                        crear_msgbox(self, "Estudiante activo",
                                     "La constancia de retiro solo se puede generar para estudiantes retirados (inactivos).",
                                     QMessageBox.Icon.Warning).exec()
                        return
                    motivo_retiro = datos_bd.get("motivo_retiro")
                    archivo = generar_constancia_retiro_docx(estudiante, institucion, self.anio_escolar, motivo_retiro)

                elif constancia == "Constancia de retiro (grado anterior)":
                    # Solo para estudiantes retirados
                    estatus_acad = str(datos_bd.get("estatus_academico", "")).strip()
                    estado_est = datos_bd.get("estado", 1)
                    if estatus_acad != "Retirado" and estado_est != 0:
                        crear_msgbox(
                            self,
                            "Estudiante no elegible",
                            "La constancia de retiro (grado anterior) solo se puede generar "
                            "para estudiantes con estatus académico 'Retirado'.",
                            QMessageBox.Icon.Warning
                        ).exec()
                        return
                    historial = EstudianteModel.obtener_historial_estudiante(persona_id)
                    if not historial:
                        crear_msgbox(
                            self, "Sin historial",
                            "No se encontró historial académico para este estudiante.",
                            QMessageBox.Icon.Warning
                        ).exec()
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
                    motivo_retiro = datos_bd.get("motivo_retiro")
                    estudiante['Grado'] = ultimo_curso['grado']
                    estudiante['Sección'] = ultimo_curso['letra']
                    estudiante['Ciudad'] = datos_bd.get("ciudad", "")
                    estudiante['Fecha Nac.'] = datos_bd.get("fecha_nac", "")
                    estudiante['Género'] = datos_bd.get("genero", "")
                    estudiante['Tipo Educ.'] = datos_bd.get("tipo_educacion", "")
                    archivo = generar_constancia_retiro_retirado(
                        estudiante, institucion, ultimo_curso['año_escolar'], motivo_retiro
                    )

                elif constancia == "Historial académico":
                    historial = EstudianteModel.obtener_historial_estudiante(persona_id)
                    archivo = generar_historial_estudiante_pdf(estudiante, historial)

                elif constancia == "Historial de notas":
                    notas = NotasModel.obtener_notas_estudiante(persona_id)
                    archivo = generar_historial_notas_pdf(estudiante, notas)

            elif poblacion == "Empleados":
                datos_bd = EmpleadoModel.obtener_por_id(persona_id)
                if not datos_bd:
                    crear_msgbox(self, "Error", "No se encontró el empleado.", QMessageBox.Icon.Critical).exec()
                    return

                empleado = self._construir_dict_empleado(datos_bd)

                if constancia == "Constancia de trabajo":
                    archivo = generar_constancia_trabajo(empleado, institucion)

            if archivo:
                # Guardar ruta del PDF generado para exportar después
                self._ruta_pdf_temporal = archivo
                # Cargar en el visor
                self.pdf_document.close()
                self.pdf_document.load(archivo)
                self.pdf_viewer.setDocument(self.pdf_document)
            else:
                crear_msgbox(self, "Error", "No se pudo generar la constancia.", QMessageBox.Icon.Critical).exec()

        except Exception as e:
            crear_msgbox(self, "Error", f"Error generando constancia:\n{e}", QMessageBox.Icon.Critical).exec()

    def _exportar_constancia(self):
        """Exporta la constancia previamente generada (abre el archivo)."""
        if not self._ruta_pdf_temporal or not os.path.exists(self._ruta_pdf_temporal):
            crear_msgbox(self, "Sin documento",
                "Debe generar una constancia primero.",
                QMessageBox.Icon.Warning).exec()
            return

        crear_msgbox(self, "Éxito",
            f"Constancia exportada correctamente:\n{self._ruta_pdf_temporal}",
            QMessageBox.Icon.Information).exec()

        abrir_archivo(self._ruta_pdf_temporal)

    def _generar_rac(self):
        """Genera el reporte RAC de empleados."""
        from utils.exportar import generar_reporte_rac
        try:
            empleados = EmpleadoModel.listar_activos()
            institucion = InstitucionModel.obtener_por_id(1)
            if not institucion:
                crear_msgbox(self, "Error", "No se pudieron cargar los datos de la institución.", QMessageBox.Icon.Critical).exec()
                return

            archivo = generar_reporte_rac(self, empleados, institucion)
            if archivo:
                crear_msgbox(self, "Éxito",
                    f"Reporte RAC exportado:\n{archivo}",
                    QMessageBox.Icon.Information).exec()
                abrir_archivo(archivo)
        except Exception as e:
            crear_msgbox(self, "Error", f"Error generando RAC:\n{e}", QMessageBox.Icon.Critical).exec()

    def _limpiar_pdf_temporal(self):
        """Limpia el visor de PDF."""
        self.pdf_document.close()
        self._ruta_pdf_temporal = None

    # --- Reportes estadísticos ---

    def actualizar_reporte(self):
        """Genera y muestra el reporte según criterios seleccionados"""
        poblacion = self.cbxPoblacion.currentText()
        idx_criterio = self.cbxCriterio.currentIndex()
        idx_tipo = self.cbxTipoGrafica.currentIndex()

        if not poblacion or idx_criterio <= 0 or idx_tipo <= 0:
            self.figure.clear()
            self.canvas.draw()
            return

        criterio = self.cbxCriterio.currentText()
        tipo = self.cbxTipoGrafica.currentText()

        # Validar rangos
        if criterio in ("Rango de edad", "Rango de salario", "Matricula por año escolar"):
            min_val = self.spnMin.value()
            max_val = self.spnMax.value()
            
            if min_val > max_val:
                crear_msgbox(
                    self,
                    "Rango inválido",
                    f"El valor mínimo ({min_val}) no puede ser mayor que el máximo ({max_val}).",
                    QMessageBox.Icon.Warning
                ).exec()
                return

        consulta_info = CriteriosReportes.CONSULTAS.get((poblacion, criterio))
        grafica = CriteriosReportes.GRAFICAS.get(tipo)

        # Limpiar completamente la figura antes de crear nueva gráfica
        self.figure.clear()
        
        ax = self.figure.add_subplot(111)
        
        if consulta_info and grafica:
            consulta, params = consulta_info
            args = []
            titulo = f"{criterio} ({poblacion})"

            if "edad_min" in params and "edad_max" in params:
                args = [self.spnMin.value(), self.spnMax.value()]
                titulo += f" {args[0]}-{args[1]} años"

            elif "salario_min" in params and "salario_max" in params:
                args = [self.spnMin.value(), self.spnMax.value()]
                titulo += f" ${args[0]}-${args[1]}"

            elif "año_inicio" in params and "año_fin" in params:
                args = [self.spnMin.value(), self.spnMax.value()]
                titulo += f" ({args[0]}-{args[1]})"

            elif "seccion" in params:
                if hasattr(self, 'cbxSeccion_reporte') and self.cbxSeccion_reporte.currentIndex() > 0:
                    seccion = self.cbxSeccion_reporte.currentText()
                    args = [seccion]
                    titulo += f" - {seccion}"
                else:
                    ax.axis("off")
                    ax.text(0.5, 0.5, "Debe seleccionar una sección", ha="center", va="center", fontsize=12)
                    self.canvas.draw()
                    return

            try:
                etiquetas, valores = consulta(*args)
                
                if not etiquetas or not valores:
                    ax.axis("off")
                    ax.text(0.5, 0.5, "No hay datos disponibles para este criterio", 
                           ha="center", va="center", fontsize=12, color="gray")
                    self.canvas.draw()
                    return
                
                self.ultima_consulta = (etiquetas, valores)
                
                # Advertencia si se limitan datos
                if len(etiquetas) > 15 and tipo in ["Torta", "Barras"] and (poblacion, criterio) != ("Estudiantes", "Por sección"):
                    titulo += "\n(Mostrando top 15)"
                
                grafica(ax, etiquetas, valores, titulo)
                
            except Exception as e:
                ax.axis("off")
                ax.text(0.5, 0.5, f"Error generando reporte:\n{str(e)}", 
                       ha="center", va="center", fontsize=10, color="red")
        else:
            ax.axis("off")
            ax.text(0.5, 0.5, "Combinación no soportada", ha="center", va="center", fontsize=12)

        self.canvas.draw()

    def on_exportar_reporte(self):
        """Exporta el reporte a PDF."""
        poblacion = self.cbxPoblacion.currentText()
        idx_criterio = self.cbxCriterio.currentIndex()
        idx_tipo = self.cbxTipoGrafica.currentIndex()
        
        if not poblacion or idx_criterio <= 0 or idx_tipo <= 0:
            crear_msgbox(
                self,
                "Sin datos",
                "Debe generar un reporte antes de exportar.",
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        if not hasattr(self, "ultima_consulta") or not self.ultima_consulta[0]:
            crear_msgbox(
                self,
                "Sin datos",
                "No hay datos para exportar. Genere un reporte primero.",
                QMessageBox.Icon.Warning
            ).exec()
            return

        try:
            criterio = self.cbxCriterio.currentText()
            tipo = self.cbxTipoGrafica.currentText()
            etiquetas, valores = self.ultima_consulta
            total = sum(valores)
            titulo = f"{criterio} ({poblacion}) - {tipo}"

            archivo = exportar_reporte_pdf(self, self.figure, titulo, criterio, etiquetas, valores, total)
            
            # Si el usuario canceló el diálogo, no hacer nada
            if not archivo:
                return
            
            crear_msgbox(
                self,
                "Éxito",
                f"Reporte exportado correctamente:\n{archivo}",
                QMessageBox.Icon.Information
            ).exec()
            
            # Abrir archivo
            abrir_archivo(archivo)
            
        except Exception as e:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo exportar el reporte: {e}",
                QMessageBox.Icon.Critical
            ).exec()

    def _generar_listado_colaboracion(self):
        """Genera un listado PDF de colaboración de inscripción por sección."""
        if not hasattr(self, 'cbxSeccion_reporte') or self.cbxSeccion_reporte.currentIndex() <= 0:
            crear_msgbox(
                self,
                "Sin sección",
                "Debe seleccionar una sección para generar el listado.",
                QMessageBox.Icon.Warning
            ).exec()
            return

        seccion_nombre = self.cbxSeccion_reporte.currentText().strip()
        partes = seccion_nombre.rsplit(' ', 1)
        if len(partes) < 2:
            crear_msgbox(
                self,
                "Sección inválida",
                "Formato de sección no reconocido.",
                QMessageBox.Icon.Warning
            ).exec()
            return

        grado = partes[0]
        letra = partes[1]

        anio = AnioEscolarModel.obtener_actual()
        if not anio:
            crear_msgbox(
                self,
                "Error",
                "No hay año escolar activo.",
                QMessageBox.Icon.Critical
            ).exec()
            return

        estudiantes = ColaboracionModel.listar_colaboracion_por_seccion(grado, letra, anio['id'])
        if not estudiantes:
            crear_msgbox(
                self,
                "Sin datos",
                "No hay estudiantes en la sección seleccionada.",
                QMessageBox.Icon.Information
            ).exec()
            return

        seccion_dict = {"nivel": "", "grado": grado, "letra": letra}
        conn = None
        cursor = None
        try:
            from utils.db import get_connection
            conn = get_connection()
            if conn:
                cursor = conn.cursor(dictionary=True)
                cursor.execute(
                    "SELECT nivel FROM secciones WHERE grado = %s AND letra = %s AND año_escolar_id = %s AND activo = 1 LIMIT 1",
                    (grado, letra, anio['id'])
                )
                row = cursor.fetchone()
                if row:
                    seccion_dict['nivel'] = row['nivel']
        except Exception:
            pass
        finally:
            if cursor:
                cursor.close()
            if conn and conn.is_connected():
                conn.close()

        institucion = InstitucionModel.obtener_por_id(1)
        resultado = generar_listado_colaboracion_seccion(seccion_dict, estudiantes, institucion)

        if resultado and not resultado.startswith("Error"):
            self._ruta_pdf_temporal = resultado
            self.stackedReportes.setCurrentIndex(0)
            self.pdf_document.close()
            self.pdf_document.load(resultado)
            self.pdf_viewer.setDocument(self.pdf_document)
        else:
            crear_msgbox(
                self,
                "Error",
                resultado,
                QMessageBox.Icon.Critical
            ).exec()

    def _exportar_listado_colaboracion(self):
        """Exporta el listado de colaboración previamente generado."""
        if not self._ruta_pdf_temporal or not os.path.exists(self._ruta_pdf_temporal):
            crear_msgbox(self, "Sin documento",
                "Debe generar el listado primero.",
                QMessageBox.Icon.Warning).exec()
            return

        crear_msgbox(self, "Éxito",
            f"Listado exportado correctamente:\n{self._ruta_pdf_temporal}",
            QMessageBox.Icon.Information).exec()

        abrir_archivo(self._ruta_pdf_temporal)
    
    ### MODULO ADMIN ###
    
    def registro_usuario(self):
        """Abre ventana de registro de usuario."""
        ventana = RegistroUsuario(self.usuario_actual, self)
        if ventana.exec() == QDialog.DialogCode.Accepted:
            self.database_usuarios()
    
    def actualizar_usuario(self):
        """Abre ventana de actualización del usuario."""
        index = self.tableW_usuarios.currentIndex()
        
        if not index.isValid():
            crear_msgbox(
                self,
                "Selección requerida",
                "Debe seleccionar un usuario de la tabla.",
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        try:
            index_source = self.tableW_usuarios.model().mapToSource(index)
            fila = index_source.row()
            model = index_source.model()
            id_usuario = int(model.item(fila, 0).text())

            ventana = ActualizarUsuario(id_usuario, self.usuario_actual, self)
            ventana.datos_actualizados.connect(self.database_usuarios)
            ventana.exec()
            
        except Exception as e:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo abrir actualización: {e}",
                QMessageBox.Icon.Critical
            ).exec()
    
    def cambiar_estado_usuario(self):
        """Cambia el estado activo/inactivo del usuario seleccionado"""
        index = self.tableW_usuarios.currentIndex()
        
        if not index.isValid():
            crear_msgbox(
                self,
                "Selección requerida",
                "Debe seleccionar un usuario de la tabla.",
                QMessageBox.Icon.Warning
            ).exec()
            return

        try:
            index_source = self.tableW_usuarios.model().mapToSource(index)
            fila = index_source.row()
            model = index_source.model()

            id_usuario = int(model.item(fila, 0).text())
            username = model.item(fila, 1).text()
            estado_actual_texto = model.item(fila, 3).text()

            # Convertir a booleano
            estado_actual = 1 if estado_actual_texto.lower() == "activo" else 0
            nuevo_estado = 0 if estado_actual == 1 else 1
            nuevo_estado_texto = "Activo" if nuevo_estado == 1 else "Inactivo"

            # Confirmación
            reply = crear_msgbox(
                self,
                "Confirmar cambio de estado",
                f"¿Está seguro de cambiar el estado del usuario '{username}' "
                f"de {estado_actual_texto} a {nuevo_estado_texto}?",
                QMessageBox.Icon.Question,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )

            if reply.exec() != QMessageBox.StandardButton.Yes:
                return

            ok, mensaje = UsuarioModel.cambiar_estado(id_usuario, nuevo_estado, self.usuario_actual)

            if ok:
                crear_msgbox(
                    self,
                    "Éxito",
                    mensaje,
                    QMessageBox.Icon.Information
                ).exec()
                self.database_usuarios()
            else:
                crear_msgbox(
                    self,
                    "Error",
                    mensaje,
                    QMessageBox.Icon.Warning
                ).exec()

        except Exception as err:
            crear_msgbox(
                self,
                "Error",
                f"Error al cambiar estado: {err}",
                QMessageBox.Icon.Critical
            ).exec()

    def database_usuarios(self):
        """Carga la tabla de usuarios."""
        try:
            datos = UsuarioModel.listar()
            
            if not datos:
                # Crear modelo vacío
                model_vacio = QStandardItemModel(0, 7)
                model_vacio.setHorizontalHeaderLabels([
                    "ID", "Usuario", "Rol", "Estado", "Nombre Completo",
                    "Fecha de creación", "Ultima actualización"
                ])
                self.proxy_usuarios = QSortFilterProxyModel(self)
                self.proxy_usuarios.setSourceModel(model_vacio)
                self.tableW_usuarios.setModel(self.proxy_usuarios)
                return
            
            columnas = [
                "ID", "Usuario", "Rol", "Estado", "Nombre Completo",
                "Fecha de creación", "Ultima actualización"
            ]

            model = QStandardItemModel(len(datos), len(columnas))
            model.setHorizontalHeaderLabels(columnas)

            for fila, registro in enumerate(datos):
                for col, valor in enumerate(registro):
                    item = QStandardItem(str(valor))
                    item.setEditable(False)
                    model.setItem(fila, col, item)

            # Proxy
            self.proxy_usuarios = QSortFilterProxyModel(self)
            self.proxy_usuarios.setSourceModel(model)
            self.proxy_usuarios.setSortCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            self.proxy_usuarios.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

            # Filtrar inactivos si checkbox está desmarcado
            if hasattr(self, 'chkMostrar_inactivos_user') and not self.chkMostrar_inactivos_user.isChecked():
                self.proxy_usuarios.setFilterRegularExpression("^Activo$")
                self.proxy_usuarios.setFilterKeyColumn(3)

            delegate = UsuarioDelegate(self.tableW_usuarios)
            self.tableW_usuarios.setItemDelegate(delegate)

            self.tableW_usuarios.setModel(self.proxy_usuarios)
            self.tableW_usuarios.setSortingEnabled(True)
            self.tableW_usuarios.setAlternatingRowColors(True)
            self.tableW_usuarios.setColumnHidden(0, True)

            # Anchos personalizados
            anchos_usuarios = {
                0: 50, 1: 120, 2: 100, 3: 80, 4: 200, 5: 170, 6: 170
            }
            ajustar_columnas_tabla(self, self.tableW_usuarios, anchos_usuarios)

            # Numeración vertical
            row_count = self.proxy_usuarios.rowCount()
            for fila in range(row_count):
                self.proxy_usuarios.setHeaderData(fila, Qt.Vertical, str(fila + 1))

        except Exception as err:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo cargar la tabla de usuarios: {err}",
                QMessageBox.Icon.Critical
            ).exec()
    
    ### MODULO AUDITORIA ###
    
    def cargar_auditoria(self, limit=50):
        """Carga los registros de auditoría."""
        try:
            datos = AuditoriaModel.listar(limit)
            
            if not datos:
                model_vacio = QStandardItemModel(0, 8)
                model_vacio.setHorizontalHeaderLabels([
                    "ID", "Usuario", "Acción", "Entidad", "Entidad ID",
                    "Referencia", "Descripción", "Fecha"
                ])
                self.tableW_auditoria.setModel(model_vacio)
                return
            
            columnas = ["ID", "Usuario", "Acción", "Entidad", "Entidad ID", "Referencia", "Descripción", "Fecha"]

            model = QStandardItemModel(len(datos), len(columnas))
            model.setHorizontalHeaderLabels(columnas)

            for fila, registro in enumerate(datos):
                items = [
                    QStandardItem(str(registro["id"])),
                    QStandardItem(registro["usuario"]),
                    QStandardItem(registro["accion"]),
                    QStandardItem(registro["entidad"]),
                    QStandardItem(str(registro["entidad_id"])),
                    QStandardItem(registro["referencia"] or ""),
                    QStandardItem(registro["descripcion"] or ""),
                    QStandardItem(str(registro["fecha"]))
                ]
                
                for col, item in enumerate(items):
                    item.setEditable(False)
                    model.setItem(fila, col, item)

            self.tableW_auditoria.setModel(model)
            self.tableW_auditoria.setSortingEnabled(True)
            self.tableW_auditoria.setAlternatingRowColors(True)

            # Ordenar por columna Fecha (índice 7) en orden descendente
            self.tableW_auditoria.sortByColumn(7, Qt.SortOrder.DescendingOrder)

            # Anchos personalizados
            anchos_auditoria = {
                0: 50, 1: 120, 2: 100, 3: 120, 4: 80, 5: 150, 6: 300, 7: 150
            }
            ajustar_columnas_tabla(self, self.tableW_auditoria, anchos_auditoria)

        except Exception as err:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo cargar la auditoría: {err}",
                QMessageBox.Icon.Critical
            ).exec()
    
    ## DATOS INSTITUCION ##
    
    def set_campos_editables(self, estado: bool):
        """Habilita/deshabilita edición de campos."""
        campos = [
            self.lneNombreInstitucion_admin, self.lneCodigoDEA_admin, self.lneCodigoDEP_admin,
            self.lneCodigoEST_admin, self.lneRIF_admin, self.lneDirInstitucion_admin,
            self.lneTlfInstitucion_admin, self.lneCorreoInstitucion_admin,
            self.lneDirector_institucion, self.lneCedula_director_institucion
        ]
        campos_solo_lectura = [self.lneUltimaActualizacion_admin]
        set_campos_editables(campos, estado, campos_solo_lectura)

    def configurar_logo_institucional(self):
        """Crea los widgets para gestión del logo en la página de datos institucionales."""
        # Crear frame contenedor para el logo
        # Buscar el contenedor padre del frameInstitucion
        parent_widget = self.frameInstitucion.parent()
        self.frameLogo_institucion = QFrame(parent_widget)
        self.frameLogo_institucion.setStyleSheet("""
            QFrame#frameLogo_institucion {
                background-color: #f8f9fa;
                border: 2px dashed #c0c0c0;
                border-radius: 10px;
            }
        """)
        self.frameLogo_institucion.setObjectName("frameLogo_institucion")
        self.frameLogo_institucion.setFixedSize(240, 220)  # Ancho y alto fijo
        
        layout_logo = QVBoxLayout(self.frameLogo_institucion)
        layout_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Label de título
        lbl_titulo_logo = QLabel("Logo de la Institución")
        lbl_titulo_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_titulo_logo.setStyleSheet("font-weight: bold; font-size: 12px; color: #333; border: none;")
        layout_logo.addWidget(lbl_titulo_logo)
        
        # Label para preview del logo
        self.lblPreview_logo = QLabel()
        self.lblPreview_logo.setFixedSize(90, 90)
        self.lblPreview_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lblPreview_logo.setStyleSheet("""
            background-color: white;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 5px;
        """)
        self.lblPreview_logo.setScaledContents(False)
        layout_logo.addWidget(self.lblPreview_logo, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # Botones
        layout_btns = QHBoxLayout()
        layout_btns.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.btnSubir_logo = QPushButton("Cambiar logo")
        self.btnSubir_logo.setFixedWidth(80)
        self.btnSubir_logo.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 5px 5px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #005fa3;
            }
        """)
        self.btnSubir_logo.clicked.connect(self._subir_logo)
        crear_sombra_flotante(self.btnSubir_logo)
        
        self.btnEliminar_logo = QPushButton("Quitar logo")
        self.btnEliminar_logo.setFixedWidth(80)
        self.btnEliminar_logo.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 5px 5px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #b02a37;
            }
        """)
        self.btnEliminar_logo.clicked.connect(self._eliminar_logo)
        crear_sombra_flotante(self.btnEliminar_logo)
        
        layout_btns.addWidget(self.btnSubir_logo)
        layout_btns.addWidget(self.btnEliminar_logo)
        layout_logo.addLayout(layout_btns)
        
        # Calcular posición: a la derecha del frameInstitucion con un margen
        x_pos = self.frameInstitucion.x() + self.frameInstitucion.width() + 15
        y_pos = self.frameInstitucion.y() + 300
        self.frameLogo_institucion.move(x_pos, y_pos)
        
        # Asegurar visibilidad
        self.frameLogo_institucion.raise_()
        self.frameLogo_institucion.show()
        
        # Aplicar sombra flotante
        crear_sombra_flotante(self.frameLogo_institucion, blur_radius=10, y_offset=2)
        
        # Cargar preview inicial
        self._cargar_preview_logo()
    
    def _cargar_preview_logo(self):
        """Carga la vista previa del logo institucional."""
        pixmap = obtener_logo_pixmap()
        if pixmap and not pixmap.isNull():
            scaled = pixmap.scaled(
                80, 80,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            self.lblPreview_logo.setPixmap(scaled)
            self.btnEliminar_logo.setEnabled(True)
        else:
            # Mostrar placeholder
            self.lblPreview_logo.setText("Sin logo")
            self.lblPreview_logo.setStyleSheet("""
                background-color: #f0f0f0;
                border: 1px dashed #aaa;
                border-radius: 8px;
                padding: 5px;
                color: #999;
                font-size: 11px;
            """)
            self.btnEliminar_logo.setEnabled(False)
    
    def _subir_logo(self):
        """Permite al usuario seleccionar y subir un logo institucional."""
        ruta, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar logo de la institución",
            "",
            LOGO_FILTRO_DIALOGO
        )
        
        if not ruta:
            return
        
        # Procesar imagen (validar y redimensionar)
        datos, mensaje = procesar_imagen(ruta)
        
        if not datos:
            crear_msgbox(
                self,
                "Error al procesar imagen",
                mensaje,
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        # Guardar en BD
        ok, msg_bd = InstitucionModel.guardar_logo(1, datos, self.usuario_actual)
        
        if not ok:
            crear_msgbox(
                self,
                "Error al guardar",
                msg_bd,
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        # Invalidar caché y recargar en toda la UI
        invalidar_cache()
        self._cargar_preview_logo()
        self._actualizar_logos_globales()
        
        crear_msgbox(
            self,
            "Éxito",
            "El logo de la institución se actualizó correctamente.\n"
            "Se aplicará en todas las ventanas y documentos PDF.",
            QMessageBox.Icon.Information
        ).exec()
    
    def _eliminar_logo(self):
        """Elimina el logo institucional de la BD."""
        msg = crear_msgbox(
            self,
            "Confirmar eliminación",
            "¿Está seguro de que desea quitar el logo de la institución?\n\n"
            "Se volverá al logo por defecto del sistema.",
            QMessageBox.Icon.Question,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if msg.exec() != QMessageBox.StandardButton.Yes:
            return
        
        ok, mensaje = InstitucionModel.eliminar_logo(1, self.usuario_actual)
        
        if ok:
            invalidar_cache()
            self._cargar_preview_logo()
            self._actualizar_logos_globales()
            
            crear_msgbox(
                self,
                "Éxito",
                "El logo fue eliminado. Se usará el logo por defecto.",
                QMessageBox.Icon.Information
            ).exec()
        else:
            crear_msgbox(
                self,
                "Error",
                mensaje,
                QMessageBox.Icon.Warning
            ).exec()
    
    def _actualizar_logos_globales(self):
        """Actualiza el logo en todos los labels visibles del MainWindow."""
        # Labels del MainWindow
        for lbl in [
            self.lblLogo_dashboard_escuela,
            self.lblLogo_reportes,
            self.lblLogo_usuarios,
            self.lblLogo_auditoria,
            self.lblLogo_datos_insti,
            self.lblLogo_backup,
        ]:
            aplicar_logo_a_label(lbl)
        
        # Actualizar páginas hijas que tengan el logo
        paginas_con_logo = [
            (self.page_gestion_estudiantes, 'lblLogo_estu'),
            (self.page_gestion_secciones, 'lblLogo_secciones'),
            (self.page_gestion_empleados, 'lblLogo_emple'),
            (self.page_egresados, 'lblLogo_egresados'),
            (self.page_gestion_notas, 'lblLogo_notas'),
            (self.page_gestion_materias, 'lblLogo_materias'),
        ]
        for pagina, attr in paginas_con_logo:
            if hasattr(pagina, attr):
                lbl = getattr(pagina, attr)
                if attr == 'lblLogo_notas':
                    aplicar_logo_a_label(lbl, ancho=45, alto=45)
                else:
                    aplicar_logo_a_label(lbl)

    def cargar_datos_institucion(self):
        """Carga datos de la institución."""
        try:
            datos = InstitucionModel.obtener_por_id(1)
            
            if not datos:
                ok, mensaje = InstitucionModel.inicializar_si_no_existe()
                if ok:
                    datos = InstitucionModel.obtener_por_id(1)
                else:
                    crear_msgbox(
                        self,
                        "Error",
                        f"No se pudo inicializar datos de institución: {mensaje}",
                        QMessageBox.Icon.Critical
                    ).exec()
                    return
            
            self.lneNombreInstitucion_admin.setText(str(datos.get("nombre", "")))
            self.lneCodigoDEA_admin.setText(str(datos.get("codigo_dea", "")))
            self.lneCodigoDEP_admin.setText(str(datos.get("codigo_dependencia", "")))
            self.lneCodigoEST_admin.setText(str(datos.get("codigo_estadistico", "")))
            self.lneRIF_admin.setText(str(datos.get("rif", "")))
            self.lneDirInstitucion_admin.setText(str(datos.get("direccion", "")))
            self.lneTlfInstitucion_admin.setText(str(datos.get("telefono", "")))
            self.lneCorreoInstitucion_admin.setText(str(datos.get("correo", "")))
            self.lneDirector_institucion.setText(str(datos.get("director", "")))
            self.lneCedula_director_institucion.setText(str(datos.get("director_ci", "")))
            
            fecha_act = datos.get("actualizado_en")
            if fecha_act:
                if hasattr(fecha_act, 'strftime'):
                    self.lneUltimaActualizacion_admin.setText(fecha_act.strftime("%d-%m-%Y %H:%M:%S"))
                else:
                    self.lneUltimaActualizacion_admin.setText(str(fecha_act))
            else:
                self.lneUltimaActualizacion_admin.setText("Sin actualizar")
            
            # Actualizar preview del logo si existe el widget
            if hasattr(self, 'lblPreview_logo'):
                self._cargar_preview_logo()
                
        except Exception as err:
            crear_msgbox(
                self,
                "Error",
                f"No se pudieron cargar los datos: {err}",
                QMessageBox.Icon.Critical
            ).exec()

    def guardar_datos_institucion(self):
        """Guarda cambios en los datos de la institución."""
        try:
            institucion_data = {
                "nombre": self.lneNombreInstitucion_admin.text().strip(),
                "codigo_dea": self.lneCodigoDEA_admin.text().strip(),
                "codigo_dependencia": self.lneCodigoDEP_admin.text().strip(),
                "codigo_estadistico": self.lneCodigoEST_admin.text().strip(),
                "rif": self.lneRIF_admin.text().strip(),
                "direccion": self.lneDirInstitucion_admin.text().strip(),
                "telefono": self.lneTlfInstitucion_admin.text().strip(),
                "correo": self.lneCorreoInstitucion_admin.text().strip(),
                "director": self.lneDirector_institucion.text().strip(),
                "director_ci": self.lneCedula_director_institucion.text().strip(),
            }

            ok, mensaje = InstitucionModel.actualizar(1, institucion_data, self.usuario_actual)

            if ok:
                crear_msgbox(
                    self,
                    "Éxito",
                    mensaje,
                    QMessageBox.Icon.Information
                ).exec()
                self.cargar_datos_institucion()
            else:
                crear_msgbox(
                    self,
                    "Error",
                    mensaje,
                    QMessageBox.Icon.Warning
                ).exec()

        except Exception as err:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo guardar cambios: {err}",
                QMessageBox.Icon.Critical
            ).exec()

    def toggle_edicion(self):
        """Alterna entre edición y guardado."""
        if self.btnModificar_institucion.text() == "Modificar datos":
            self.set_campos_editables(True)
            self.btnModificar_institucion.setText("Guardar")
        else:
            self.guardar_datos_institucion()
            self.set_campos_editables(False)
            self.btnModificar_institucion.setText("Modificar datos")
   
    ### MODULO BACKUP ###
    
    def cargar_info_backup(self):
        """Carga información del último backup."""
        try:
            ultimo = BackupManager.obtener_ultimo_backup()
            total = BackupManager.contar_backups()
            
            if hasattr(self, 'lblUltimo_backup'):
                if ultimo:
                    fecha_str = ultimo['fecha'].strftime("%d-%m-%Y %H:%M:%S")
                    tipo_str = ultimo['tipo'].capitalize()
                    self.lblUltimo_backup.setText(
                        f"Último backup: {fecha_str}\n"
                        f"Tipo: {tipo_str}\n"
                        f"Tamaño: {ultimo['tamaño_mb']:.2f} MB\n"
                        f"Total de backups: {total}"
                    )
                else:
                    self.lblUltimo_backup.setText("No hay backups disponibles")
                    
        except Exception as e:
            print(f"Error cargando info de backup: {e}")
    
    def realizar_backup_manual(self):
        """Ejecuta backup manual."""
        try:
            # Confirmación
            reply = crear_msgbox(
                self,
                "Confirmar backup",
                "¿Desea crear un backup manual de la base de datos?\n\n"
                "Este proceso puede tardar varios segundos dependiendo del tamaño de la base de datos.",
                QMessageBox.Icon.Question,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            
            if reply.exec() != QMessageBox.StandardButton.Yes:
                return
            
            # Mostrar mensaje de progreso
            self.statusBar().showMessage("Creando backup manual...")
            
            # Crear backup
            ok, mensaje = BackupManager.crear_backup_manual()
            
            # Limpiar status bar
            self.statusBar().clearMessage()
            
            if ok:
                crear_msgbox(
                    self,
                    "Éxito",
                    mensaje,
                    QMessageBox.Icon.Information
                ).exec()
                
                # Actualizar información
                self.cargar_info_backup()
                
                # Abrir carpeta de backups
                reply_abrir = crear_msgbox(
                    self,
                    "Backup creado",
                    "¿Desea abrir la carpeta de backups?",
                    QMessageBox.Icon.Question,
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                
                if reply_abrir.exec() == QMessageBox.StandardButton.Yes:
                    ruta_backups = os.path.abspath("backups")
                    abrir_carpeta(ruta_backups)
            else:
                crear_msgbox(
                    self,
                    "Error",
                    f"No se pudo crear el backup:\n{mensaje}",
                    QMessageBox.Icon.Critical
                ).exec()
                
        except Exception as e:
            crear_msgbox(
                self,
                "Error",
                f"Error inesperado creando backup: {e}",
                QMessageBox.Icon.Critical
            ).exec()
    
    def realizar_backup_automatico(self):
        """Ejecuta backup automático."""
        try:
            print("Iniciando backup automático...")
            ok, mensaje = BackupManager.crear_backup_automatico()
            
            if ok:
                print(f"Backup automático exitoso: {mensaje}")
                # Actualizar info si la página está visible
                if hasattr(self, 'stackMain') and self.stackMain.currentIndex() == 10:
                    self.cargar_info_backup()
            else:
                print(f"Error en backup automático: {mensaje}")
                
        except Exception as e:
            print(f"Error en backup automático: {e}")

    def restaurar_backup(self):
        """Permite al usuario seleccionar y restaurar un backup."""
        try:
            # Listar backups disponibles
            backups = BackupManager.listar_backups()

            if not backups:
                # Si no hay backups en la carpeta, permitir seleccionar archivo externo
                reply_externo = crear_msgbox(
                    self,
                    "Sin backups",
                    "No hay backups disponibles en la carpeta de backups.\n\n"
                    "¿Desea seleccionar un archivo .sql externo?",
                    QMessageBox.Icon.Question,
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                if reply_externo.exec() != QMessageBox.StandardButton.Yes:
                    return
                self._restaurar_desde_archivo_externo()
                return

            # Crear diálogo de selección de backup
            dialog = QDialog(self)
            dialog.setWindowTitle("Restaurar copia de seguridad")
            dialog.setMinimumSize(550, 400)
            layout = QVBoxLayout(dialog)

            # Instrucciones
            lbl_instrucciones = QLabel(
                "Seleccione el backup que desea restaurar.\n"
                "⚠️ Esta acción reemplazará todos los datos actuales de la base de datos."
            )
            lbl_instrucciones.setWordWrap(True)
            lbl_instrucciones.setStyleSheet("color: #c0392b; font-weight: bold; padding: 8px;")
            layout.addWidget(lbl_instrucciones)

            # Lista de backups
            lista = QListWidget()
            lista.setStyleSheet("""
                QListWidget {
                    font-size: 13px;
                    padding: 4px;
                }
                QListWidget::item {
                    padding: 8px;
                    border-bottom: 1px solid #e0e0e0;
                }
                QListWidget::item:selected {
                    background-color: #0078d7;
                    color: white;
                }
            """)

            for backup in backups:
                fecha_str = backup['fecha'].strftime("%d-%m-%Y %H:%M:%S")
                tipo_str = backup['tipo'].capitalize()
                texto = (
                    f"{backup['nombre']}\n"
                    f"   📅 {fecha_str}  |  📦 {tipo_str}  |  💾 {backup['tamaño_mb']:.2f} MB"
                )
                item = QListWidgetItem(texto)
                item.setData(Qt.ItemDataRole.UserRole, backup['ruta'])
                lista.addItem(item)

            lista.setCurrentRow(0)
            layout.addWidget(lista)

            # Botones
            frame_botones = QFrame()
            layout_botones = QHBoxLayout(frame_botones)

            btn_externo = QPushButton("Seleccionar archivo externo...")
            btn_externo.setStyleSheet("padding: 8px 16px;")
            btn_restaurar = QPushButton("Restaurar seleccionado")
            btn_restaurar.setStyleSheet(
                "padding: 8px 16px; background-color: #e74c3c; color: white; font-weight: bold;"
            )
            btn_cancelar = QPushButton("Cancelar")
            btn_cancelar.setStyleSheet("padding: 8px 16px;")

            layout_botones.addWidget(btn_externo)
            layout_botones.addStretch()
            layout_botones.addWidget(btn_restaurar)
            layout_botones.addWidget(btn_cancelar)
            layout.addWidget(frame_botones)

            ruta_seleccionada = [None]

            def on_restaurar():
                item_actual = lista.currentItem()
                if item_actual:
                    ruta_seleccionada[0] = item_actual.data(Qt.ItemDataRole.UserRole)
                    dialog.accept()

            def on_externo():
                dialog.reject()
                self._restaurar_desde_archivo_externo()

            btn_restaurar.clicked.connect(on_restaurar)
            btn_cancelar.clicked.connect(dialog.reject)
            btn_externo.clicked.connect(on_externo)

            if dialog.exec() != QDialog.DialogCode.Accepted or not ruta_seleccionada[0]:
                return

            self._ejecutar_restauracion(ruta_seleccionada[0])

        except Exception as e:
            crear_msgbox(
                self,
                "Error",
                f"Error inesperado: {e}",
                QMessageBox.Icon.Critical
            ).exec()

    def _restaurar_desde_archivo_externo(self):
        """Permite seleccionar un archivo .sql externo para restaurar."""
        archivo, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo de backup",
            "",
            "Archivos SQL (*.sql);;Todos los archivos (*)"
        )
        if archivo:
            self._ejecutar_restauracion(archivo)

    def _ejecutar_restauracion(self, ruta_archivo: str):
        """Ejecuta la restauración de un backup con confirmación de seguridad."""
        nombre_archivo = os.path.basename(ruta_archivo)

        # Primera confirmación
        reply = crear_msgbox(
            self,
            "⚠️ Confirmar restauración",
            f"Está a punto de restaurar el backup:\n\n"
            f"📄 {nombre_archivo}\n\n"
            f"Esta acción REEMPLAZARÁ todos los datos actuales de la base de datos.\n"
            f"Se recomienda crear un backup antes de continuar.\n\n"
            f"¿Desea continuar?",
            QMessageBox.Icon.Warning,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply.exec() != QMessageBox.StandardButton.Yes:
            return

        # Segunda confirmación (seguridad extra)
        reply2 = crear_msgbox(
            self,
            "Última confirmación",
            "¿Está COMPLETAMENTE seguro?\n\n"
            "Los datos actuales serán reemplazados y esta acción no se puede deshacer.",
            QMessageBox.Icon.Critical,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply2.exec() != QMessageBox.StandardButton.Yes:
            return

        # Mostrar progreso
        self.statusBar().showMessage("Restaurando backup... Por favor espere.")

        # Ejecutar restauración
        ok, mensaje = BackupManager.restaurar_backup(ruta_archivo)

        # Limpiar status bar
        self.statusBar().clearMessage()

        if ok:
            crear_msgbox(
                self,
                "Éxito",
                f"{mensaje}\n\n"
                f"La aplicación se cerrará para aplicar los cambios.\n"
                f"Por favor, inicie sesión nuevamente.",
                QMessageBox.Icon.Information
            ).exec()

            # Actualizar info de backup
            self.cargar_info_backup()

            # Cerrar sesión para que los datos se recarguen
            self.cerrar_sesion()
        else:
            crear_msgbox(
                self,
                "Error",
                f"No se pudo restaurar el backup:\n{mensaje}",
                QMessageBox.Icon.Critical
            ).exec()

    def acceso_directo_registro_estudiante(self):
        """Abre el formulario de registro de estudiante."""
        if hasattr(self, 'page_gestion_estudiantes'):
            self.page_gestion_estudiantes.registro_estudiante()
    
    def acceso_directo_registro_empleado(self):
        """Abre el formulario de registro de empleado."""
        if hasattr(self, 'page_gestion_empleados'):
            self.page_gestion_empleados.registro_empleados()
    
    def acceso_directo_crear_seccion(self):
        """Abre el formulario de crear sección."""
        if hasattr(self, 'page_gestion_secciones'):
            self.page_gestion_secciones.nueva_seccion()
    
    def cerrar_sesion(self):
        """Cierra la sesión actual."""
        self.logout = True
        self.close()
    
    def mostrar_acerca_de(self):
        """Abre ventana 'Acerca de'."""
        ventana = AcercaDe(self)
        ventana.exec()

    def mostrar_manual_usuario(self):
        """Abre el manual de usuario en PDF"""
        ruta_manual = resource_path("resources/icons/Manual_de_Usuario_SIRA.pdf")
        if os.path.exists(ruta_manual):
            abrir_archivo(ruta_manual)
        else:
            crear_msgbox(
                self,
                "Manual no encontrado",
                "No se encontró el archivo del manual de usuario.",
                QMessageBox.Icon.Warning
            ).exec()
    
    def closeEvent(self, event):
        """Limpia recursos al cerrar la ventana."""
        # Limpiar tooltips de delegates
        for delegate in self.tooltip_delegates:
            if hasattr(delegate, 'close_tooltip'):
                delegate.close_tooltip()
        
        # Detener timers
        if hasattr(self, 'timer_global'):
            self.timer_global.stop()
        if hasattr(self, 'timer_backup'):
            self.timer_backup.stop()
        
        super().closeEvent(event)
