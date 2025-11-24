from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, QFormLayout, QLineEdit, QPushButton,
    QFileDialog, QListWidget, QListWidgetItem, QHBoxLayout, QInputDialog,
    QDialog, QLabel, QComboBox, QDialogButtonBox
)
import os, json

class ConfiguracoesWidget(QWidget):
    def __init__(self, settings_file, parent=None):
        super().__init__(parent)
        self.settings_file = settings_file
        self.log_path = ""
        self.users = []

        self._load_settings()
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)

        # Grupo 1: Caminho da pasta de logs
        log_group = QGroupBox("Pasta Principal de Logs")
        log_layout = QFormLayout(log_group)

        self.log_path_display = QLineEdit(self.log_path)
        self.log_path_display.setReadOnly(True)
        log_layout.addRow("Caminho Atual:", self.log_path_display)

        select_btn = QPushButton("Selecionar Pasta")
        select_btn.clicked.connect(self._selecionar_pasta)
        log_layout.addRow(select_btn)
        layout.addWidget(log_group)

        # Grupo 2: Usuários
        user_group = QGroupBox("Usuários Autorizados")
        user_layout = QVBoxLayout(user_group)

        self.user_list = QListWidget()
        for user in self.users:
            self.user_list.addItem(QListWidgetItem(user))

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("Adicionar Usuário")
        remove_btn = QPushButton("Remover Selecionado")

        add_btn.clicked.connect(self._adicionar_usuario)
        remove_btn.clicked.connect(self._remover_usuario)

        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)

        user_layout.addWidget(self.user_list)
        user_layout.addLayout(btn_layout)

        layout.addWidget(user_group)

    def _selecionar_pasta(self):
        path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Logs")
        if path:
            self.log_path = path
            self.log_path_display.setText(path)
            self._salvar_settings()

    def _adicionar_usuario(self):
        nome, ok = QInputDialog.getText(self, "Adicionar Usuário", "Nome:")
        if ok and nome.strip():
            nome = nome.strip()
            if nome not in self.users:
                self.users.append(nome)
                self.user_list.addItem(QListWidgetItem(nome))
                self._salvar_settings()

    def _remover_usuario(self):
        item = self.user_list.currentItem()
        if item:
            nome = item.text()
            self.users.remove(nome)
            self.user_list.takeItem(self.user_list.row(item))
            self._salvar_settings()

    def _load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.log_path = data.get("log_path", "")
                    self.users = data.get("users", [])
            except Exception:
                pass

    def _salvar_settings(self):
        data = {
            "log_path": self.log_path,
            "users": self.users
        }
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"Erro ao salvar settings: {e}")

    def get_log_path(self):
        return self.log_path or os.path.expanduser("~/Documents")

    def get_users(self):
        return self.users

    def selecionar_usuario_popup(self):
        from PyQt6.QtWidgets import QDialog, QVBoxLayout, QLabel, QComboBox, QDialogButtonBox
        import os, json

        dlg = QDialog(self)
        dlg.setWindowTitle("Selecionar Usuário")
        layout = QVBoxLayout(dlg)

        layout.addWidget(QLabel("Selecione o operador do teste:"))

        # Caminho do settings local
        local_settings_path = os.path.join(os.path.expanduser("~"), "Documents", "EmbTechSerial", "local_settings.json")
        ultimo_usuario_local = ""

        try:
            with open(local_settings_path, "r", encoding="utf-8") as f:
                local_data = json.load(f)
                ultimo_usuario_local = local_data.get("last_user", "")
        except:
            pass

        combo = QComboBox()
        combo.addItems(self.users)

        if ultimo_usuario_local in self.users:
            combo.setCurrentText(ultimo_usuario_local)

        layout.addWidget(combo)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        layout.addWidget(buttons)

        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)

        if dlg.exec():
            selecionado = combo.currentText()
            try:
                with open(local_settings_path, "w", encoding="utf-8") as f:
                    json.dump({"last_user": selecionado}, f, indent=4)
            except Exception as e:
                print(f"Erro ao salvar usuário local: {e}")
            return selecionado

        return None
