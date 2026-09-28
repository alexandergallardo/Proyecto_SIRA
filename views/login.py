from PySide6.QtWidgets import QDialog, QMessageBox
from PySide6.QtCore import QTimer

from utils.db import get_user_by_username
from utils.security import check_password
from utils.dialogs import crear_msgbox
from utils.sombras import crear_sombra_flotante
from utils.logo_manager import aplicar_logo_a_label
from models.auditoria_model import AuditoriaModel

from ui_compiled.login_ui import Ui_login


class LoginDialog(QDialog, Ui_login):
    """Diálogo de autenticación"""
    
    MAX_INTENTOS = 3
    TIEMPO_BLOQUEO_MS = 30000  # 30 segundos
    
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("Inicio de sesión")
        self.lblVersion.setText("Versión SIRA: v-1.2.6")
        crear_sombra_flotante(self.lblLogo_SIRA, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.lblLogo_escuela, blur_radius=8, y_offset=1)
        
        # Aplicar logo institucional dinámico
        aplicar_logo_a_label(self.lblLogo_escuela, ancho=71, alto=80)
        
        # Control de intentos fallidos
        self.intentos_fallidos = 0
        self.bloqueado = False
        self.usuario = None

        # Conectar señales
        self.btnLogin.clicked.connect(self.on_login_clicked)
        self.inputPassword.returnPressed.connect(self.on_login_clicked)
        self.inputUser.returnPressed.connect(self.inputPassword.setFocus)
        
        # Aplicar efectos visuales
        crear_sombra_flotante(self.btnLogin)
        crear_sombra_flotante(self.inputUser, blur_radius=8, y_offset=1)
        crear_sombra_flotante(self.inputPassword, blur_radius=8, y_offset=1)

    def on_login_clicked(self):
        """Procesa el intento de login."""
        
        # Verificar si está bloqueado
        if self.bloqueado:
            crear_msgbox(
                self,
                "Acceso bloqueado",
                "Demasiados intentos fallidos.\n"
                "Por favor, espere antes de intentar nuevamente.",
                QMessageBox.Icon.Warning
            ).exec()
            return
        
        # Validar campos vacíos
        username = self.inputUser.text().strip()
        password = self.inputPassword.text()
        
        if not username:
            crear_msgbox(
                self,
                "Campo requerido",
                "Debe ingresar un nombre de usuario.",
                QMessageBox.Icon.Warning
            ).exec()
            self.inputUser.setFocus()
            return
        
        if not password:
            crear_msgbox(
                self,
                "Campo requerido",
                "Debe ingresar una contraseña.",
                QMessageBox.Icon.Warning
            ).exec()
            self.inputPassword.setFocus()
            return

        try:
            # Buscar usuario en la base de datos
            user = get_user_by_username(username)
            
            if not user:
                self._manejar_login_fallido("Usuario no encontrado.")
                return

            # Verificar estado del usuario
            if user.get("estado") == 0:
                crear_msgbox(
                    self,
                    "Cuenta deshabilitada",
                    "Su cuenta ha sido deshabilitada.\n"
                    "Contacte con el administrador del sistema.",
                    QMessageBox.Icon.Warning
                ).exec()
                self._limpiar_campos()
                return

            # Verificar contraseña
            if not check_password(password, user.get("password_hash", "")):
                self._manejar_login_fallido("Contraseña incorrecta.")
                return

            # Login exitoso
            self.usuario = {
                "id": user["id"],
                "username": user["username"],
                "rol": user["rol"],
                "estado": user["estado"]
            }
            
            # Registrar acceso en auditoría
            try:
                AuditoriaModel.registrar(
                    usuario_id=user["id"],
                    accion="LOGIN",
                    entidad="sistema",
                    entidad_id=user["id"],
                    referencia=user["username"],
                    descripcion=f"Inicio de sesión exitoso - Rol: {user['rol']}"
                )
            except Exception as e:
                print(f"Error registrando auditoría de login: {e}")
            
            # Aceptar diálogo
            self.accept()

        except Exception as e:
            crear_msgbox(
                self,
                "Error de conexión",
                f"No se pudo conectar con el servidor:\n{str(e)}\n\n"
                "Verifique su conexión a la base de datos.",
                QMessageBox.Icon.Critical
            ).exec()
            self._limpiar_campos()

    def _manejar_login_fallido(self, mensaje: str):
        """Maneja intentos fallidos con bloqueo progresivo."""
        self.intentos_fallidos += 1
        
        intentos_restantes = self.MAX_INTENTOS - self.intentos_fallidos
        
        if intentos_restantes > 0:
            crear_msgbox(
                self,
                "Error de autenticación",
                f"{mensaje}\n\n"
                f"Intentos restantes: {intentos_restantes}",
                QMessageBox.Icon.Warning
            ).exec()
        else:
            # Bloquear temporalmente
            self.bloqueado = True
            self.btnLogin.setEnabled(False)
            self.inputUser.setEnabled(False)
            self.inputPassword.setEnabled(False)
            
            crear_msgbox(
                self,
                "Acceso bloqueado",
                f"Ha excedido el número máximo de intentos ({self.MAX_INTENTOS}).\n\n"
                f"El acceso será bloqueado durante {self.TIEMPO_BLOQUEO_MS // 1000} segundos.",
                QMessageBox.Icon.Critical
            ).exec()
            
            # Desbloquear después del tiempo establecido
            QTimer.singleShot(self.TIEMPO_BLOQUEO_MS, self._desbloquear_login)
        
        self._limpiar_campos()

    def _desbloquear_login(self):
        """Desbloquea el formulario después del tiempo de penalización."""
        self.bloqueado = False
        self.intentos_fallidos = 0
        self.btnLogin.setEnabled(True)
        self.inputUser.setEnabled(True)
        self.inputPassword.setEnabled(True)
        
        crear_msgbox(
            self,
            "Acceso restaurado",
            "Puede intentar iniciar sesión nuevamente.",
            QMessageBox.Icon.Information
        ).exec()
        
        self.inputUser.setFocus()

    def _limpiar_campos(self):
        """Limpia los campos de entrada."""
        self.inputPassword.clear()
        self.inputUser.setFocus()
