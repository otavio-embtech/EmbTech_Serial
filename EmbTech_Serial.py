import sys
import serial
import json
import time
import re
import os
import shutil
import configparser
import serial.tools.list_ports
from configuracoes_widget import ConfiguracoesWidget
import threading
import tempfile
import shlex
from relatorio_eficiencia_widget import RelatorioEficienciaWidget
from oled_timer_display import OledTimerDisplay
from collections import deque
from datetime import datetime


from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QComboBox, QLabel, QLineEdit, QPushButton, QTextEdit, QGroupBox,
    QFileDialog, QMessageBox, QCheckBox, QDialog, QDialogButtonBox,
    QTabWidget, QFormLayout, QSpinBox, QDoubleSpinBox, QListWidget, QListWidgetItem,
    QToolButton, QMenu, QRadioButton, QSizePolicy, QTableWidget, QTableWidgetItem,
    QHeaderView, QScrollArea
)
from PyQt6.QtCore import QThread, pyqtSignal, QTimer, Qt, QEvent, QPoint, QRegularExpression
from PyQt6.QtGui import QCursor, QAction, QIcon, QBrush, QColor, QPixmap, QRegularExpressionValidator, QPalette
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import requests
import winsound

# Importa a nova biblioteca Modbus (assumindo que está no mesmo diretório ou acessível no PATH)
import modbus_lib


class _FolderHandler(FileSystemEventHandler):
    def __init__(self, on_new_dir_callback):
        super().__init__()
        self._cb = on_new_dir_callback

    def on_created(self, event):
        if event.is_directory:
            name = os.path.basename(event.src_path)
            self._cb(name)

    def on_moved(self, event):
        # Algumas criações aparecem como 'moved' no Windows (ex.: criar e renomear)
        try:
            if event.is_directory:
                name = os.path.basename(getattr(event, 'dest_path', event.src_path))
                self._cb(name)
        except Exception:
            pass


class ClickableLabel(QLabel):
    """
    Um QLabel que emite um sinal 'clicked' quando clicado.
    Útil para criar áreas clicáveis que se comportam como botões invisíveis.
    """
    clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True) # Habilita o rastreamento do mouse para feedback visual

    def mousePressEvent(self, event):
        """
        Sobrescreve o evento de clique do mouse para emitir o sinal 'clicked'.
        """
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event) # Chama o método da classe base para manter o comportamento padrão


class ZoomableImageViewer(QScrollArea):
    """
    Visualizador de imagem com zoom pelo scroll do mouse e duplo clique para alternar
    entre 'ajustar à janela' e 100%.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self._pixmap = QPixmap()
        self._scale = 1.0
        self._fit = True
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setWidget(self.label)
        self.setWidgetResizable(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)

    def set_image(self, pixmap: QPixmap):
        self._pixmap = pixmap
        self._scale = 1.0
        self._fit = True
        self._update_pixmap()

    def wheelEvent(self, event):
        if self._pixmap.isNull():
            return
        delta = event.angleDelta().y()
        if delta == 0:
            return
        factor = 1.15 if delta > 0 else 1 / 1.15
        self._fit = False
        self._scale = max(0.05, min(20.0, self._scale * factor))
        self._update_pixmap()

    def mouseDoubleClickEvent(self, event):
        if self._pixmap.isNull():
            return
        # Alterna entre ajustar e 100%
        self._fit = not self._fit
        if self._fit:
            self._scale = 1.0
        else:
            self._scale = 1.0
        self._update_pixmap()
        super().mouseDoubleClickEvent(event)

    def resizeEvent(self, event):
        self._update_pixmap()
        super().resizeEvent(event)

    def _update_pixmap(self):
        if self._pixmap.isNull():
            self.label.clear()
            return
        if self._fit:
            area_w = max(1, self.viewport().width())
            area_h = max(1, self.viewport().height())
            scaled = self._pixmap.scaled(area_w, area_h, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.label.setPixmap(scaled)
        else:
            w = max(1, int(self._pixmap.width() * self._scale))
            h = max(1, int(self._pixmap.height() * self._scale))
            scaled = self._pixmap.scaled(w, h, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            self.label.setPixmap(scaled)

class RefreshableComboBox(QComboBox):
    """
    Um QComboBox que atualiza sua lista de itens antes de exibir o popup,
    garantindo que as portas COM estejam sempre atualizadas.
    """
    def __init__(self, refresh_function, parent=None):
        super().__init__(parent)
        self.refresh_function = refresh_function

    def showPopup(self):
        """
        Chamado antes de exibir a lista de opções.
        Chama a função de atualização de portas antes de mostrar o popup.
        """
        self.refresh_function() # Chama a função para atualizar a lista de portas
        super().showPopup() # Exibe o popup com a lista atualizada


class SerialReaderThread(QThread):
    """
    Thread dedicada para leitura de dados de uma porta serial.
    Evita que a interface do usuário congele durante a espera por dados.
    Emite 'data_received' quando dados são recebidos e 'connection_lost' em caso de erro.
    """
    # data_received agora emite bytes para a porta Modbus e string para a porta serial principal
    data_received = pyqtSignal(object, str) # Sinal para dados recebidos (dados, nome da porta)
    connection_lost = pyqtSignal(str)    # Sinal para conexão perdida (nome da porta)

    def __init__(self, ser_instance, port_name="", is_modbus_port=False):
        super().__init__()
        self.ser = ser_instance
        self._running = True # Flag para controlar o loop da thread
        # Buffer para armazenar as últimas respostas (bytes para Modbus, string para serial)
        self._response_buffer = deque(maxlen=100) 
        self._clear_response_buffer_flag = False # Flag para limpar o buffer antes de um novo passo de teste
        self.port_name = port_name # Nome da porta para identificação em logs
        self.is_modbus_port = is_modbus_port # Flag para indicar se é uma porta Modbus

    def run(self):
        """
        Método principal da thread que lê a porta serial continuamente.
        """
        while self._running and self.ser.is_open:
            try:
                if self._clear_response_buffer_flag:
                    self._response_buffer.clear()
                    self._clear_response_buffer_flag = False

                if self.is_modbus_port:
                    # Para Modbus, leia bytes diretamente
                    # A leitura de Modbus deve ser mais controlada (ex: por um timeout de inter-caracteres)
                    # Aqui, faremos uma leitura simples com timeout da porta
                    # Você pode precisar de uma lógica de leitura mais robusta para Modbus (ex: espera por X bytes ou timeout de silêncio)
                    line_bytes = self.ser.read_all() # Lê todos os bytes disponíveis até o timeout
                    if line_bytes:
                        self.data_received.emit(line_bytes, self.port_name)
                        self._response_buffer.append(line_bytes)
                else:
                    # Para porta serial principal, decodifique para string
                    line_bytes = self.ser.readline() # Lê uma linha da serial
                    if line_bytes:
                        # Decodifica e remove espaços em branco (incluindo o '\n' final)
                        line_str = line_bytes.decode('utf-8', errors='ignore').strip() 
                        # Emite o sinal mesmo se a linha for vazia (representa uma linha em branco do dispositivo)
                        self.data_received.emit(line_str, self.port_name) 
                        self._response_buffer.append(line_str) # Adiciona ao buffer

            except serial.SerialException:
                # Erro de comunicação serial (ex: cabo desconectado)
                self.connection_lost.emit(self.port_name)
                self._running = False
            except Exception as e:
                # Outros erros inesperados
                print(f"Erro inesperado na thread de leitura serial ({self.port_name}): {e}")
                self._running = False
                self.connection_lost.emit(self.port_name)

    def stop(self):
        """
        Para a execução da thread de forma segura.
        """
        self._running = False
        self.wait() # Espera a thread terminar sua execução

    def get_buffered_response(self):
        """
        Retorna o conteúdo atual do buffer de respostas e o limpa.
        Retorna bytes para Modbus e string para serial principal.
        """
        if self.is_modbus_port:
            # Para Modbus, concatena os bytes
            response_bytes = b''.join(self._response_buffer)
            self._response_buffer.clear()
            return response_bytes
        else:
            # Para serial principal, concatena as strings
            response_str = "\n".join(self._response_buffer)
            self._response_buffer.clear()
            return response_str
    
    def clear_response_buffer_for_next_step(self):
        """
        Define uma flag para limpar o buffer na próxima iteração do loop de leitura.
        Usado para garantir que a resposta de um novo passo não contenha dados antigos.
        """
        self._clear_response_buffer_flag = True


class TimerConfigDialog(QDialog):
    """
    Diálogo para configurar o intervalo de tempo para o envio automático de comandos.
    """
    def __init__(self, current_interval_seconds, parent=None):
        super().__init__(parent)
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "assets", "icone.ico")
        self.setWindowIcon(QIcon(icon_path))

        self.setWindowTitle("Configurar Intervalo do Timer")
        self.setFixedSize(250, 100)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False) # Remove o botão de ajuda

        self.interval_seconds = current_interval_seconds

        layout = QVBoxLayout(self)

        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("Intervalo (segundos):"))
        self.interval_line_edit = QLineEdit(str(self.interval_seconds))
        self.interval_line_edit.setPlaceholderText("Ex: 1, 0.5")
        input_layout.addWidget(self.interval_line_edit)
        layout.addLayout(input_layout)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept) # Conecta o botão OK ao método accept do diálogo
        self.button_box.rejected.connect(self.reject) # Conecta o botão Cancel ao método reject do diálogo
        layout.addWidget(self.button_box)

    def get_interval(self):
        """
        Retorna o intervalo de tempo configurado pelo usuário.
        Realiza validação para garantir que o valor é um número positivo.
        """
        try:
            interval = float(self.interval_line_edit.text())
            if interval <= 0:
                QMessageBox.warning(self, "Erro de Validação", "O intervalo deve ser um número positivo maior que zero.")
                return None
            return interval
        except ValueError:
            QMessageBox.warning(self, "Erro de Validação", "Por favor, insira um número válido para o intervalo (ex: 1, 0.5, 2.5).")
            return None


class TestIdDialog(QDialog):
    """
    Diálogo para coletar o número do PR e o número de série da placa antes de iniciar um teste.
    """
    def __init__(self, last_pr_number="", last_serial_number="", parent=None):
        super().__init__(parent)
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "assets", "icone.ico")
        self.setWindowIcon(QIcon(icon_path))

        self.setWindowTitle("Iniciar Teste - Informações da Placa")
        self.setFixedSize(350, 150)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False) # Remove o botão de ajuda

        layout = QFormLayout(self)

        self.pr_number_input = QLineEdit(last_pr_number)
        self.pr_number_input.setPlaceholderText("Ex: PR04237")
        pr_regex = QRegularExpression(r"^(PR)?\d{5}$")
        self.pr_number_input.setValidator(QRegularExpressionValidator(pr_regex))
        self.pr_number_input.setToolTip("Formato exigido: PR + 5 dígitos (aceita também apenas 5 dígitos).")
        layout.addRow("Número do PR:", self.pr_number_input)

        self.serial_number_input = QLineEdit(last_serial_number)
        self.serial_number_input.setPlaceholderText("Ex: 16913/155")
        sn_regex = QRegularExpression(r"^\d+\/\d+$")
        self.serial_number_input.setValidator(QRegularExpressionValidator(sn_regex))
        self.serial_number_input.setToolTip("Formato exigido: números no formato NNNNN/NN (somente dígitos e uma barra).")
        layout.addRow("Número de Série da Placa:", self.serial_number_input)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addRow(self.button_box)

    def get_info(self):
        """
        Retorna o número do PR e o número de série inseridos pelo usuário.
        """
        pr = self.pr_number_input.text().strip()
        # Normaliza PR removendo prefixo 'PR' se presente
        if pr.upper().startswith("PR"):
            pr = pr[2:]
        serial = self.serial_number_input.text().strip()
        return pr, serial



class RetestStepDialog(QDialog):
    """
    Diálogo exibido quando um passo de teste falha, permitindo ao usuário re-testar
    o passo ou finalizar o teste.
    """
    def __init__(self, step_name, error_description, parent=None):
        super().__init__(parent)
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "assets", "icone.ico")
        self.setWindowIcon(QIcon(icon_path))

        self.setWindowTitle(f"Re-testar Passo: {step_name}")
        self.setFixedSize(400, 200)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False)

        layout = QVBoxLayout(self)

        info_label = QLabel(f"Erro no passo '{step_name}':")
        info_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(info_label)

        error_text_edit = QTextEdit()
        error_text_edit.setPlainText(error_description)
        error_text_edit.setReadOnly(True)
        error_text_edit.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        layout.addWidget(error_text_edit)

        button_layout = QHBoxLayout()

        self.retest_button = QPushButton("Re-testar Passo")
        self.retest_button.clicked.connect(self.accept)
        button_layout.addWidget(self.retest_button)

        self.finish_test_button = QPushButton("Finalizar Teste")
        self.finish_test_button.clicked.connect(self.reject)
        button_layout.addWidget(self.finish_test_button)

        layout.addLayout(button_layout)
class TestPortConfigDialog(QDialog):
    """
    Diálogo para configurar as portas seriais específicas para um arquivo de teste.
    Permite definir baud rate, bits de dados, paridade, controle de fluxo e modo.
    """
    def __init__(self, current_serial_settings, current_modbus_settings, parent=None):
        super().__init__(parent)
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "assets", "icone.ico")
        self.setWindowIcon(QIcon(icon_path))

        self.setWindowTitle("Configurar Portas para este Teste")
        self.setFixedSize(450, 400)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False) # Remove o botão de ajuda

        main_layout = QVBoxLayout(self)

        # Configuração da Porta Principal
        serial_group = QGroupBox("Porta Principal (Dados e Comandos)")
        serial_layout = QFormLayout(serial_group)

        self.serial_baud_combo = QComboBox()
        self.serial_baud_combo.addItems(["9600", "19200", "38400", "57600", "115200", "230400", "460800", "921600"])
        self.serial_baud_combo.setCurrentText(current_serial_settings.get("baud", "9600"))
        serial_layout.addRow("Baud Rate:", self.serial_baud_combo)

        self.serial_data_bits_combo = QComboBox()
        self.serial_data_bits_combo.addItems(["5", "6", "7", "8"])
        self.serial_data_bits_combo.setCurrentText(current_serial_settings.get("data_bits", "8"))
        serial_layout.addRow("Bits de Dados:", self.serial_data_bits_combo)

        self.serial_parity_combo = QComboBox()
        self.serial_parity_combo.addItems(["Nenhuma", "Ímpar", "Par", "Marca", "Espaço"])
        self.serial_parity_combo.setCurrentText(current_serial_settings.get("parity", "Nenhuma"))
        serial_layout.addRow("Paridade:", self.serial_parity_combo)

        self.serial_handshake_combo = QComboBox()
        self.serial_handshake_combo.addItems(["Nenhum", "RTS/CTS", "XON/XOFF"])
        self.serial_handshake_combo.setCurrentText(current_serial_settings.get("handshake", "Nenhum"))
        serial_layout.addRow("Controle de Fluxo:", self.serial_handshake_combo)
        
        self.serial_mode_combo = QComboBox()
        self.serial_mode_combo.addItems(["Free", "PortStore test", "Data", "Setup"])
        self.serial_mode_combo.setCurrentText(current_serial_settings.get("mode", "Free"))
        serial_layout.addRow("Modo:", self.serial_mode_combo)

        main_layout.addWidget(serial_group)

        # Configuração da Porta Modbus
        modbus_group = QGroupBox("Porta Modbus (Leitura e Escrita)")
        modbus_layout = QFormLayout(modbus_group)

        self.modbus_baud_combo = QComboBox()
        self.modbus_baud_combo.addItems(["9600", "19200", "38400", "57600", "115200", "230400", "460800", "921600"])
        self.modbus_baud_combo.setCurrentText(current_modbus_settings.get("baud", "9600"))
        modbus_layout.addRow("Baud Rate:", self.modbus_baud_combo)

        self.modbus_data_bits_combo = QComboBox()
        self.modbus_data_bits_combo.addItems(["5", "6", "7", "8"])
        self.modbus_data_bits_combo.setCurrentText(current_modbus_settings.get("data_bits", "8"))
        modbus_layout.addRow("Bits de Dados:", self.modbus_data_bits_combo)

        self.modbus_parity_combo = QComboBox()
        self.modbus_parity_combo.addItems(["Nenhuma", "Ímpar", "Par", "Marca", "Espaço"])
        self.modbus_parity_combo.setCurrentText("Nenhuma")
        self.modbus_parity_combo.setCurrentText(current_modbus_settings.get("parity", "Nenhuma"))
        modbus_layout.addRow("Paridade:", self.modbus_parity_combo)

        self.modbus_handshake_combo = QComboBox()
        self.modbus_handshake_combo.addItems(["Nenhum", "RTS/CTS", "XON/XOFF"])
        self.modbus_handshake_combo.setCurrentText(current_modbus_settings.get("handshake", "Nenhum"))
        modbus_layout.addRow("Controle de Fluxo:", self.modbus_handshake_combo)
        
        self.modbus_mode_combo = QComboBox()
        self.modbus_mode_combo.addItems(["Free", "PortStore test", "Data", "Setup"])
        self.modbus_mode_combo.setCurrentText(current_modbus_settings.get("mode", "Free"))
        modbus_layout.addRow("Modo:", self.modbus_mode_combo)

        main_layout.addWidget(modbus_group)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        main_layout.addWidget(self.button_box)
        # Flag para identificar fechamento pelo X
        self.closed_via_titlebar = False

    def closeEvent(self, event):
        # Marca que foi fechado pelo X da janela (titlebar); exec() retornará Rejected
        self.closed_via_titlebar = True
        super().closeEvent(event)

        # Flag para identificar fechamento pelo X da janela
        self.closed_via_titlebar = False

    def get_settings(self):
        """
        Retorna as configurações de porta serial e Modbus selecionadas no diálogo.
        """
        serial_settings = {
            "baud": self.serial_baud_combo.currentText(),
            "data_bits": self.serial_data_bits_combo.currentText(),
            "parity": self.serial_parity_combo.currentText(),
            "handshake": self.serial_handshake_combo.currentText(),
            "mode": self.serial_mode_combo.currentText()
        }
        modbus_settings = {
            "baud": self.modbus_baud_combo.currentText(),
            "data_bits": self.modbus_data_bits_combo.currentText(),
            "parity": self.modbus_parity_combo.currentText(),
            "handshake": self.modbus_handshake_combo.currentText(),
            "mode": self.modbus_mode_combo.currentText()
        }
        return serial_settings, modbus_settings


class ManualInstructionDialog(QDialog):
    """
    Diálogo para exibir instruções manuais ao usuário, opcionalmente com uma imagem.
    """
    def __init__(self, step_number, total_steps, step_name, instruction_message, image_path, parent=None):
        super().__init__(parent)
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "assets", "icone.ico")
        self.setWindowIcon(QIcon(icon_path))

        self.setWindowTitle(f"Instrução Manual - Passo {step_number}/{total_steps}: {step_name}")
        self.setMinimumSize(400, 300)
        self.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False) # Remove o botão de ajuda

        main_layout = QVBoxLayout(self)

        instruction_label = QLabel("Instrução:")
        instruction_label.setStyleSheet("font-weight: bold;")
        main_layout.addWidget(instruction_label)

        self.instruction_text_edit = QTextEdit()
        self.instruction_text_edit.setPlainText(instruction_message)
        self.instruction_text_edit.setReadOnly(True)
        self.instruction_text_edit.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.instruction_text_edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        main_layout.addWidget(self.instruction_text_edit)

        if image_path and os.path.exists(image_path):
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                self.image_viewer = ZoomableImageViewer()
                self.image_viewer.setToolTip("Use a roda do mouse para dar zoom. Duplo clique alterna ajustar/100%.")
                self.image_viewer.setMinimumHeight(250)
                self.image_viewer.set_image(pixmap)
                main_layout.addWidget(self.image_viewer)
            else:
                self.image_label = QLabel("Erro: Não foi possível carregar a imagem.")
                self.image_label.setStyleSheet("color: red;")
                main_layout.addWidget(self.image_label)
        else:
            self.image_label = QLabel("Nenhuma imagem fornecida ou caminho inválido.")
            self.image_label.setStyleSheet("color: gray;")
            main_layout.addWidget(self.image_label)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        main_layout.addWidget(self.button_box)

        # Única flag e único closeEvent para detectar fechamento pelo X
        self.closed_via_titlebar = False

    def closeEvent(self, event):
        self.closed_via_titlebar = True
        super().closeEvent(event)

# Novo widget para exibir um passo de teste com um checkbox para o modo fast
class TestStepListItemWidget(QWidget):
    def __init__(self, step_data, parent=None):
        super().__init__(parent)
        self.step_data = step_data
        self.init_ui()

    def init_ui(self):
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)

        self.checkbox = QCheckBox()
        self.checkbox.setChecked(self.step_data.get('checked_for_fast_mode', False))
        self.checkbox.stateChanged.connect(self._update_checked_state)
        layout.addWidget(self.checkbox)

        step_type_label = ""
        if self.step_data.get("tipo_passo") == "instrucao_manual":
            step_type_label = "[Instrução Manual]"
        elif self.step_data.get("tipo_passo") == "tempo_espera":
            step_type_label = "[Tempo de Espera]"
        elif self.step_data.get("tipo_passo") == "modbus_comando": # Novo tipo de passo
            step_type_label = "[Comando Modbus]"
        elif self.step_data.get("tipo_passo") == "gravar_numero_serie":
            step_type_label = "[Gravar NS]"
        elif self.step_data.get("tipo_passo") == "gravar_placa":
            step_type_label = "[Gravar Placa]"
        else:
            port_label = "Principal" if self.step_data.get("port_type", "serial") == "serial" else "Modbus"
            step_type_label = f"[Comando Auto - Porta {port_label}]"

        self.name_label = QLabel(f"{step_type_label} {self.step_data.get('nome', 'Sem Nome')}")
        layout.addWidget(self.name_label)
        layout.addStretch(1) # Empurra os elementos para a esquerda

        self.setLayout(layout)

    def _update_checked_state(self, state):
        """Atualiza o estado 'checked_for_fast_mode' no dicionário step_data."""
        self.step_data['checked_for_fast_mode'] = (state == Qt.CheckState.Checked.value)

    def get_step_data(self):
        """Retorna os dados do passo de teste, incluindo o estado do checkbox."""
        return self.step_data


class PlacaTesterApp(QMainWindow):
    """
    Aplicação principal para teste de placas via comunicação serial e Modbus.
    Oferece terminal, envio automático/manual de comandos e um criador de testes.
    """
    VERSION = "3.9.6" # Versão atual do aplicativo (incrementada para tema)

    CONFIG_FILE_TEST_OPERATOR = 'test_operator_config.ini' # Use um nome de arquivo diferente para esta configuração
    CONFIG_SECTION_TEST_OPERATOR = 'TestOperator'

    def __init__(self):
        super().__init__()
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "assets", "icone.ico")
        self.setWindowIcon(QIcon(icon_path))
        self.config = configparser.ConfigParser() # Inicialize o configparser aqui

        # Inicializa um atributo para guardar o último operador de teste
        self.last_test_operator = "" 

        # Carrega o último operador de teste salvo logo na inicialização da aplicação
        self._load_last_test_operator() 
        self.setWindowTitle("EmbTech Serial")
        self.setGeometry(100, 100, 1050, 780)
        # Define limites de tamanho para evitar janela fora da tela
        # Ajustado o tamanho mínimo para garantir que o conteúdo seja visível
        # e o máximo para evitar expansão excessiva, mas permitindo flexibilidade.
        self.setMinimumSize(900, 700) 
        # Removido o setMaximumSize para permitir que a janela seja maximizada
        # sem ser limitada por um tamanho fixo, mas ainda respeitando a barra de tarefas.
        # self.setMaximumSize(1920, 1080)  # Limite máximo para telas Full HD

        # Impede o auto-redimensionamento vertical forçado
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Preferred)
        # Define a posição e o tamanho inicial da janela
        
        # --- NOVO: Inicialização do timer para monitorar portas COM ---
        self.port_monitor_timer = QTimer(self)
        self.port_monitor_timer.setInterval(2000)  # Verifica a cada 2 segundos (ajuste conforme necessário)
        self.port_monitor_timer.timeout.connect(self._update_available_ports)
        self.port_monitor_timer.start() # Inicia o timer

        # Armazenar as portas atualmente listadas para evitar atualizações desnecessárias da UI
        self.current_serial_ports = []
        self.current_modbus_ports = []
        # Instâncias de portas seriais e threads de leitura
        self.serial_command_ser = None
        self.modbus_ser = None
        self.serial_command_reader_thread = None
        self.modbus_serial_reader_thread = None

        # Variáveis de controle de teste
        self.current_test_steps = [] # Lista de passos do teste carregado/criado
        self.current_test_index = -1 # Índice do passo atual em execução
        self.test_in_progress = False # Flag para indicar se um teste está em andamento
        self.test_timer = QTimer(self) # Timer para controlar timeouts de resposta em testes
        self.passed_steps_count = 0 # Contador de passos aprovados
        self.failed_steps_count = 0 # Contador de passos reprovados

        self.current_pr_number = "" # Número do PR (Product Request) do teste atual
        self.current_serial_number = "" # Número de série da placa em teste
        self.test_log_entries = [] # Entradas de log para o teste atual
        self.modbus_required_for_test = False # Indica se o teste atual requer a porta Modbus

        self.test_serial_command_settings = {} # Configurações da porta principal para o teste
        self.test_modbus_settings = {} # Configurações da porta Modbus para o teste
        # Estados de auto-reconexão Modbus
        self.modbus_reconnect_timer = None
        self.modbus_reconnect_attempts_remaining = 0
        self.modbus_target_port = ""

        # Listas para gerenciar os campos de envio automático/manual
        self.send_command_inputs = []
        self.send_buttons = []
        self.auto_send_timers = []
        self.auto_send_checkboxes = []
        self.auto_send_config_buttons = []
        self.auto_send_intervals_s = [1.0] * 4 # Intervalos padrão para os 4 timers de auto-envio

        # Define o caminho base para arquivos de configuração
        if getattr(sys, 'frozen', False):
            self.base_path = os.path.dirname(sys.executable)
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))
        self.settings_file = os.path.join(self.base_path, "settings.json") # Caminho do arquivo de configurações

        self.hidden_button_click_count = 0 # Contador para liberar a aba de criador de teste
        self.test_creator_tab_index = -1 # Índice da aba do criador de teste
        self.editing_step_index = -1 # Índice do passo sendo editado no criador de teste
        
        self.log_interaction_count = 0 # Contador para limpar o log automaticamente

        # Variáveis para o Modo Fast
        self.fast_mode_active = False # Estado atual do modo fast
        self.fast_mode_secret_code = "" # Código secreto para ativar/desativar o modo fast

        self._apply_stylesheet() # Aplica o stylesheet antes de inicializar a UI
        self._init_ui() # Inicializa a interface do usuário do terminal
        self._init_test_creator_ui() # Inicializa a interface do usuário do criador de testes
        
        # Conecta os sinais dos rádio botões para alternar os campos do criador de teste
        self.radio_comando_auto.toggled.connect(self._toggle_step_type_fields)
        self.radio_instrucao_manual.toggled.connect(self._toggle_step_type_fields)
        self.radio_tempo_espera.toggled.connect(self._toggle_step_type_fields)
        self.radio_modbus_comando.toggled.connect(self._toggle_step_type_fields) # Conecta o novo rádio botão Modbus
        # Novo: rádio para gravar número de série
        try:
            self.radio_gravar_ns.toggled.connect(self._toggle_step_type_fields)
        except AttributeError:
            pass
        # Novo: rádio para Gravar Placa
        try:
            self.radio_gravar_placa.toggled.connect(self._toggle_step_type_fields)
        except AttributeError:
            pass
        self.radio_comando_auto.setChecked(True) # Define o rádio botão de comando automático como padrão

        self._list_serial_ports() # Lista as portas seriais disponíveis
        self._load_settings() # Carrega as configurações salvas
        self._update_port_config_visibility(True) # Inicia expandido
        self.setWindowTitle("EmbTech Serial")
        self.show()

        from oled_timer_display import OledTimerDisplay
        self.oled = OledTimerDisplay()
        self.oled.hide()

    def _load_last_test_operator(self):
        self.config.read(self.CONFIG_FILE_TEST_OPERATOR)
        if self.config.has_section(self.CONFIG_SECTION_TEST_OPERATOR):
            last_operator = self.config.get(self.CONFIG_SECTION_TEST_OPERATOR, 'last_operator_name', fallback="")
            self.last_test_operator = last_operator
        else:
            self.last_test_operator = "" # Nenhum operador salvo
    
    def _save_last_test_operator(self, operator_name):
        if not self.config.has_section(self.CONFIG_SECTION_TEST_OPERATOR):
            self.config.add_section(self.CONFIG_SECTION_TEST_OPERATOR)
        self.config.set(self.CONFIG_SECTION_TEST_OPERATOR, 'last_operator_name', operator_name)
        try:
            with open(self.CONFIG_FILE_TEST_OPERATOR, 'w') as configfile:
                self.config.write(configfile)
        except Exception as e:
            # Em um aplicativo real, você pode querer logar isso ou mostrar uma QMessageBox
            print(f"Erro ao salvar o nome do operador do teste: {e}")
            self.log_message(f"Erro ao salvar o nome do operador do teste: {e}", "erro")

    def _prompt_for_operator_name(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Nome do Operador")
        layout = QVBoxLayout()

        label = QLabel("Por favor, insira o nome do operador para iniciar o teste:")
        layout.addWidget(label)

        operator_input = QLineEdit()
        # Pré-preenche o QLineEdit com o último operador salvo
        operator_input.setText(self.last_test_operator) 
        layout.addWidget(operator_input)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        dialog.setLayout(layout)

        # Executa o diálogo e verifica o resultado
        if dialog.exec() == QDialog.DialogCode.Accepted:
            operator_name = operator_input.text().strip()
            if operator_name:
                self.current_operator = operator_name # Armazena o nome do operador no seu atributo de instância
                self.log_message(f"Operador '{self.current_operator}' definido para o teste.", "info")
                self._save_last_test_operator(self.current_operator) # Salva o nome do operador
                return True # Retorna True se o nome foi confirmado e é válido
            else:
                QMessageBox.warning(self, "Nome do Operador Inválido", "O nome do operador não pode ser vazio.")
                return False # Retorna False se o nome estiver vazio
        else:
            self.log_message("Início do teste cancelado (nome do operador não fornecido).", "aviso")
            return False # Retorna False se o diálogo foi cancelado

    def _calculate_modbus_crc(self, data):
        """
        Calcula o CRC-16 para dados Modbus RTU.
        """
        crc = 0xFFFF
        for byte in data:
            crc ^= byte
            for _ in range(8):
                if crc & 0x0001:
                    crc >>= 1
                    crc ^= 0xA001
                else:
                    crc >>= 1
        return crc.to_bytes(2, 'little') # CRC é LSB primeiro (low byte, high byte)

    def _create_port_configuration_layout(self, title, config_key):
        # Esta função é a que cria os QComboBoxes de portas, você a usa para Serial Command e Modbus.
        # Precisamos garantir que ela retorne o QComboBox para que possamos atualizar
        # Exemplo de como você pode estar usando ela:
        self.serial_command_port_combo = self._create_port_configuration_layout("Comando Serial", "serial_command")
        self.modbus_port_combo = self._create_port_configuration_layout("Modbus", "modbus")

        # ... (seu código existente para _create_port_configuration_layout) ...
        # Certifique-se de que os QComboBoxes de portas são atributos da classe (self.)
        # para que possam ser acessados pela função _update_available_ports.

        port_layout = QFormLayout()
        
        # O QComboBox da porta
        # Para garantir que é acessível em _update_available_ports, atribua-o a self
        if config_key == "serial_command":
            # Usa a nova classe RefreshableComboBox
            self.serial_command_port_combobox = RefreshableComboBox(self._update_available_ports)
            self.serial_command_port_combobox.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            port_layout.addRow("Porta:", self.serial_command_port_combobox)
            return self.serial_command_port_combobox # Retorne o combo box para que possa ser referenciado
        elif config_key == "modbus":
            # Usa a nova classe RefreshableComboBox
            self.modbus_port_combobox = RefreshableComboBox(self._update_available_ports)
            self.modbus_port_combobox.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            port_layout.addRow("Porta:", self.modbus_port_combobox)
            return self.modbus_port_combobox # Retorne o combo box para que possa ser referenciado
        # ... (resto da função) ...

        # Uma chamada inicial para popular as listas ao iniciar o app
        self._update_available_ports()

    def _update_available_ports(self):
        """
        Verifica as portas seriais disponíveis e atualiza os QComboBoxes.
        """

        available_ports = [port.device for port in serial.tools.list_ports.comports()]

        # Atualiza o QComboBox da porta Serial de Comando
        if hasattr(self, 'serial_command_port_combobox') and self.serial_command_port_combobox is not None:
            # Salva a seleção atual antes de limpar
            selected_port = self.serial_command_port_combobox.currentText()
            
            # Limpa e adiciona os novos itens apenas se houver uma mudança real nas portas
            if set(available_ports) != set(self.current_serial_ports):
                self.serial_command_port_combobox.clear()
                self.serial_command_port_combobox.addItems(available_ports)
                
                # Tenta re-selecionar a porta que estava selecionada, se ainda existir
                if selected_port and selected_port in available_ports:
                    self.serial_command_port_combobox.setCurrentText(selected_port)
                elif available_ports: # Se a porta selecionada não existe mais, seleciona a primeira disponível
                    self.serial_command_port_combobox.setCurrentIndex(0)
                
                self.current_serial_ports = available_ports[:] # Atualiza o cache de portas
            
            # Se não houver portas disponíveis, garante que o combobox esteja vazio
            if not available_ports and self.serial_command_port_combobox.count() > 0:
                self.serial_command_port_combobox.clear()
                self.serial_command_port_combobox.addItem("Nenhuma porta COM encontrada")
                self.serial_command_port_combobox.setCurrentText("Nenhuma porta COM encontrada")


        # Atualiza o QComboBox da porta Modbus
        if hasattr(self, 'modbus_port_combobox') and self.modbus_port_combobox is not None:
            # Salva a seleção atual antes de limpar
            selected_modbus_port = self.modbus_port_combobox.currentText()

            # Limpa e adiciona os novos itens apenas se houver uma mudança real nas portas
            if set(available_ports) != set(self.current_modbus_ports):
                self.modbus_port_combobox.clear()
                self.modbus_port_combobox.addItems(available_ports)

                # Tenta re-selecionar a porta que estava selecionada, se ainda existir
                if selected_modbus_port and selected_modbus_port in available_ports:
                    self.modbus_port_combobox.setCurrentText(selected_modbus_port)
                elif available_ports: # Se a porta selecionada não existe mais, seleciona a primeira disponível
                    self.modbus_port_combobox.setCurrentIndex(0)

                self.current_modbus_ports = available_ports[:] # Atualiza o cache de portas

            # Se não houver portas disponíveis, garante que o combobox esteja vazio
            if not available_ports and self.modbus_port_combobox.count() > 0:
                self.modbus_port_combobox.clear()
                self.modbus_port_combobox.addItem("Nenhuma porta COM encontrada")
                self.modbus_port_combobox.setCurrentText("Nenhuma porta COM encontrada")

        # Habilita/desabilita botões de conexão com base na disponibilidade de portas
        self.connect_serial_command_button.setEnabled(len(available_ports) > 0)
        self.serial_command_port_combobox.setEnabled(len(available_ports) > 0)
        self.connect_modbus_button.setEnabled(len(available_ports) > 0 and self.modbus_serial_group.isVisible())
        self.modbus_port_combobox.setEnabled(len(available_ports) > 0 and self.modbus_serial_group.isVisible())

        # Você pode adicionar um log ou print para ver as portas que ele detecta
        #self.log_message(f"Portas COM atualizadas: {', '.join(available_ports) if available_ports else 'Nenhuma'}", "informacao")

    def _apply_stylesheet(self):
        """
        Aplica um stylesheet QSS para estilizar a aplicação, adaptando-se ao tema do sistema.
        Define cores de log para serem visíveis em modos claro e escuro.
        """
        # Detecta o tema do sistema
        palette = QApplication.instance().palette()
        # Verifica se a cor de fundo padrão é escura (indicando um tema escuro)
        is_dark_mode = palette.color(QPalette.ColorRole.Window).lightnessF() < 0.5

        if is_dark_mode:
            # Cores para o log - Otimizadas para modo escuro
            self.LOG_COLORS = {
                "enviado": '#87CEEB',    # Azul Céu (mais claro)
                "recebido": '#90EE90',   # Verde Claro (mais vibrante)
                "sistema": '#FFD700',    # Dourado (vibrante)
                "erro": '#FF6347',       # Tomate (vibrante)
                "test_pass": '#32CD32',  # Verde Limão (vibrante)
                "test_fail": '#FF4500',  # Laranja Avermelhado (vibrante)
                "informacao": '#E0E0E0', # Cinza Claro (para texto padrão)
                "alerta": '#FF0000'      # Vermelho para alertas
            }
            # Estilos para o modo escuro
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #2b2b2b; /* Fundo escuro para a janela principal */
                    color: #e0e0e0; /* Texto claro */
                    font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
                    font-size: 13px;
                }
                QTabWidget::pane {
                    border: 1px solid #505050;
                    background-color: #3c3c3c; /* Fundo escuro para o conteúdo das abas */
                }
                QTabWidget::tab-bar {
                    left: 5px;
                }
                QTabBar::tab {
                    background: #4a4a4a;
                    border: 1px solid #505050;
                    border-bottom-color: #3c3c3c;
                    border-top-left-radius: 4px;
                    border-top-right-radius: 4px;
                    min-width: 8ex;
                    padding: 6px 12px;
                    margin-right: 2px;
                    color: #e0e0e0;
                }
                QTabBar::tab:selected {
                    background: #3c3c3c;
                    border-color: #505050;
                    border-bottom-color: #3c3c3c;
                }
                QTabBar::tab:hover {
                    background: #5a5a5a;
                }
                QGroupBox {
                    background-color: #3c3c3c;
                    border: 1px solid #505050;
                    border-radius: 6px;
                    margin-top: 1ex;
                    padding-top: 10px;
                    padding-bottom: 5px;
                    color: #e0e0e0;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top left;
                    padding: 0 3px;
                    background-color: #4a4a4a;
                    border: 1px solid #505050;
                    border-radius: 3px;
                    color: #e0e0e0;
                    font-weight: bold;
                }
                QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {
                    border: 1px solid #606060;
                    border-radius: 4px;
                    padding: 4px;
                    background-color: #4a4a4a;
                    color: #e0e0e0;
                }
                QTextEdit#log_text_edit { /* Estilo específico para o terminal no modo escuro */
                    background-color: #1e1e1e; /* Fundo bem escuro para o terminal */
                    color: #e0e0e0; /* Cor do texto padrão do terminal */
                    border: 1px solid #606060;
                }
                QPushButton {
                    background-color: #007bff;
                    color: white;
                    border: 1px solid #007bff;
                    border-radius: 5px;
                    padding: 6px 12px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #0056b3;
                    border-color: #004085;
                }
                QPushButton:pressed {
                    background-color: #004085;
                }
                QPushButton:disabled {
                    background-color: #505050;
                    border-color: #404040;
                    color: #999999;
                }
                QCheckBox, QRadioButton {
                    padding: 3px;
                    color: #e0e0e0;
                }
                QLabel {
                    color: #e0e0e0;
                }
                QToolButton {
                    border: none;
                    background-color: transparent;
                    padding: 5px;
                }
                QToolButton::menu-indicator {
                    image: none;
                }
                QListWidget {
                    border: 1px solid #606060;
                    border-radius: 4px;
                    background-color: #4a4a4a;
                    color: #e0e0e0;
                    padding: 5px;
                }
                QListWidget::item {
                    padding: 3px;
                }
                QListWidget::item:selected {
                    background-color: #5a5a5a;
                    color: #ffffff;
                }
                /* Estilos específicos para os botões de conexão de porta */
                #connect_serial_command_button[text="Abrir Porta Principal"] {
                    background-color: #28a745;
                    border-color: #218838;
                }
                #connect_serial_command_button[text="Abrir Porta Principal"]:hover {
                    background-color: #218838;
                }
                #connect_serial_command_button[text="Fechar Porta Principal"] {
                    background-color: #dc3545;
                    border-color: #c82333;
                }
                #connect_serial_command_button[text="Fechar Porta Principal"]:hover {
                    background-color: #c82333;
                }
                #connect_modbus_button[text="Abrir Porta Modbus"] {
                    background-color: #28a745;
                    border-color: #218838;
                }
                #connect_modbus_button[text="Abrir Porta Modbus"]:hover {
                    background-color: #218838;
                }
                #connect_modbus_button[text="Fechar Porta Modbus"] {
                    background-color: #dc3545;
                    border-color: #c82333;
                }
                #connect_modbus_button[text="Fechar Porta Modbus"]:hover {
                    background-color: #c82333;
                }
                /* Estilos para a QTableWidget Modbus */
                QTableWidget {
                    border: 1px solid #606060;
                    border-radius: 4px;
                    background-color: #4a4a4a;
                    gridline-color: #505050;
                    color: #e0e0e0;
                }
                QHeaderView::section {
                    background-color: #007bff;
                    color: white;
                    padding: 4px;
                    border: 1px solid #0056b3;
                    font-weight: bold;
                }
                QTableWidget::item {
                    padding: 4px;
                }
                QTableWidget::item:selected {
                    background-color: #5a5a5a;
                    color: #ffffff;
                }
            """)
        else:
            # Cores para o log - Otimizadas para modo claro (levemente beges para não cansar os olhos)
            self.LOG_COLORS = {
                "enviado": '#0000CD',    # Azul Médio
                "recebido": '#006400',   # Verde Escuro
                "sistema": '#4B0082',    # Índigo
                "erro": '#FF0000',       # Vermelho Puro
                "test_pass": '#008000',  # Verde
                "test_fail": '#FF4500',  # Laranja Avermelhado
                "informacao": '#333333', # Cinza Escuro
                "alerta": '#FF0000'      # Vermelho para alertas
            }
            # Estilos para o modo claro (com beges suaves)
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #f0f0f0; /* Fundo claro para a janela principal */
                    font-family: "Segoe UI", "Helvetica Neue", Arial, sans-serif;
                    font-size: 13px;
                }
                QTabWidget::pane {
                    border: 1px solid #c0c0c0;
                    background-color: #fdfdf5; /* Bege muito claro para o conteúdo das abas */
                }
                QTabWidget::tab-bar {
                    left: 5px;
                }
                QTabBar::tab {
                    background: #e0e0e0;
                    border: 1px solid #c0c0c0;
                    border-bottom-color: #c0c0c0;
                    border-top-left-radius: 4px;
                    border-top-right-radius: 4px;
                    min-width: 8ex;
                    padding: 6px 12px;
                    margin-right: 2px;
                }
                QTabBar::tab:selected {
                    background: #fdfdf5; /* Bege muito claro para a aba selecionada */
                    border-color: #c0c0c0;
                    border-bottom-color: #fdfdf5;
                }
                QTabBar::tab:hover {
                    background: #d0d0d0;
                }
                QGroupBox {
                    background-color: #fdfdf5; /* Bege muito claro para os GroupBoxes */
                    border: 1px solid #d0d0d0;
                    border-radius: 6px;
                    margin-top: 1ex;
                    padding-top: 10px;
                    padding-bottom: 5px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top left;
                    padding: 0 3px;
                    background-color: #e8e8e8;
                    border: 1px solid #d0d0d0;
                    border-radius: 3px;
                    color: #333333;
                    font-weight: bold;
                }
                QLineEdit, QTextEdit, QComboBox, QSpinBox, QDoubleSpinBox {
                    border: 1px solid #a0a0a0;
                    border-radius: 4px;
                    padding: 4px;
                    background-color: #fefefe; /* Branco quase bege para inputs */
                }
                QTextEdit#log_text_edit { /* Estilo específico para o terminal no modo claro */
                    background-color: #f8f8f0; /* Bege suave para o fundo do terminal */
                    color: #333333; /* Cor do texto padrão do terminal */
                    border: 1px solid #d0d0d0;
                }
                QPushButton {
                    background-color: #007bff;
                    color: white;
                    border: 1px solid #007bff;
                    border-radius: 5px;
                    padding: 6px 12px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #0056b3;
                    border-color: #004085;
                }
                QPushButton:pressed {
                    background-color: #004085;
                }
                QPushButton:disabled {
                    background-color: #cccccc;
                    border-color: #aaaaaa;
                    color: #666666;
                }
                QCheckBox, QRadioButton {
                    padding: 3px;
                }
                QLabel {
                    color: #333333;
                }
                QToolButton {
                    border: none;
                    background-color: transparent;
                    padding: 5px;
                }
                QToolButton::menu-indicator {
                    image: none;
                }
                QListWidget {
                    border: 1px solid #a0a0a0;
                    border-radius: 4px;
                    background-color: #fefefe; /* Branco quase bege para listas */
                    padding: 5px;
                }
                QListWidget::item {
                    padding: 3px;
                }
                QListWidget::item:selected {
                    background-color: #e0e0e0;
                    color: #333333;
                }
                /* Estilos para a QTableWidget Modbus */
                QTableWidget {
                    border: 1px solid #a0a0a0;
                    border-radius: 4px;
                    background-color: #fefefe; /* Branco quase bege para tabelas */
                    gridline-color: #d0d0d0;
                }
                QHeaderView::section {
                    background-color: #007bff;
                    color: white;
                    padding: 4px;
                    border: 1px solid #0056b3;
                    font-weight: bold;
                }
                QTableWidget::item {
                    padding: 4px;
                }
                QTableWidget::item:selected {
                    background-color: #e0e0e0;
                    color: #333333;
                }
            """)

    def _init_ui(self):
        """
        Inicializa os componentes da interface do usuário para a aba "Terminal".
        """
        self.tab_widget = QTabWidget(self)
        self.setCentralWidget(self.tab_widget) # Define o QTabWidget como o widget central
  
        main_tab_widget = QWidget()
        main_layout = QHBoxLayout(main_tab_widget)
        self.tab_widget.addTab(main_tab_widget, "Terminal") # Adiciona a aba "Terminal"
        self.main_tab_widget = main_tab_widget

        # Painel esquerdo (log de comunicação e envio de comandos)
        left_panel = QVBoxLayout()
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setObjectName("log_text_edit") # Adiciona objectName para QSS
        self.log_text_edit.setReadOnly(True) # Torna o log somente leitura
        
        # Define a cor de fundo do terminal para se adaptar ao tema do sistema
        # A cor de fundo será definida pelo stylesheet, então removemos a definição via palette aqui
        # palette = self.log_text_edit.palette()
        # palette.setColor(QPalette.ColorRole.Base, QApplication.palette().color(QPalette.ColorRole.Base))
        # self.log_text_edit.setPalette(palette)
        
        # Remove a cor de fundo hardcoded do stylesheet para que a cor da paleta seja aplicada
        self.log_text_edit.setStyleSheet("font-size: 14px; font-family: 'Consolas', 'Courier New', monospace; border: 1px solid #d0d0d0; border-radius: 4px;")
        
        left_panel.addWidget(QLabel("Dados Recebidos/Enviados:"))
        left_panel.addWidget(self.log_text_edit)

        # Configura o menu de contexto para o log
        self.log_text_edit.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.log_text_edit.customContextMenuRequested.connect(self._show_log_context_menu)


        # Grupo para envio de comando direto
        direct_send_group = QGroupBox("Enviar Comando Direto")
        direct_send_layout = QHBoxLayout(direct_send_group)
        
        self.direct_command_input = QLineEdit()
        self.direct_command_input.setPlaceholderText("Digite seu comando aqui...")
        self.direct_command_input.setEnabled(False) # Desabilitado até a porta ser conectada
        self.direct_command_input.returnPressed.connect(self._send_direct_command) # Envia ao pressionar Enter
        direct_send_layout.addWidget(self.direct_command_input)

        # Novo: seletor de porta para envio direto e pré-configurados
        self.send_target_port_combo = QComboBox()
        self.send_target_port_combo.addItems(["Principal", "Modbus"])  # Porta alvo do envio
        self.send_target_port_combo.setCurrentText("Principal")
        self.send_target_port_combo.setEnabled(False)
        direct_send_layout.addWidget(self.send_target_port_combo)
        # Ao selecionar Modbus, exibe/expande a configuração Modbus para conectar
        self.send_target_port_combo.currentTextChanged.connect(self._on_send_target_changed)

        # Opção: exibir resposta Modbus como texto (para testes unitários)
        self.modbus_display_text_cb = QCheckBox("Texto")
        self.modbus_display_text_cb.setToolTip("Exibir respostas Modbus como texto (ASCII)")
        self.modbus_display_text_cb.setChecked(True)
        self.modbus_display_text_cb.setEnabled(False)
        direct_send_layout.addWidget(self.modbus_display_text_cb)

        self.direct_send_button = QPushButton("Enviar")
        self.direct_send_button.setEnabled(False) # Desabilitado até a porta ser conectada
        self.direct_send_button.clicked.connect(self._send_direct_command)
        direct_send_layout.addWidget(self.direct_send_button)

        left_panel.addWidget(direct_send_group)

        # Grupo para comandos pré-configurados e envio automático
        send_group = QGroupBox("Comandos Pré-configurados (Envio Automático/Manual)")
        send_layout = QVBoxLayout(send_group)

        for i in range(4): # Cria 4 linhas para comandos pré-configurados
            single_command_line_layout = QHBoxLayout()

            send_input = QLineEdit()
            send_input.setPlaceholderText(f"Comando {i+1}")
            send_input.setEnabled(False)
            
            send_button = QPushButton("Enviar")
            send_button.setEnabled(False)
            send_button.clicked.connect(lambda _, idx=i: self._send_command_button_clicked(idx)) # Usa lambda para passar o índice
            
            auto_send_cb = QCheckBox("Auto Enviar")
            auto_send_cb.setChecked(False)
            auto_send_cb.setEnabled(False)
            
            config_timer_btn = QPushButton("Configurar Timer")
            config_timer_btn.setEnabled(False)
            config_timer_btn.clicked.connect(lambda _, idx=i: self._open_timer_config_dialog(idx))

            single_command_line_layout.addWidget(send_input, 1)
            single_command_line_layout.addWidget(send_button)
            single_command_line_layout.addWidget(auto_send_cb)
            single_command_line_layout.addWidget(config_timer_btn)
            
            send_layout.addLayout(single_command_line_layout)
            
            # Armazena os widgets em listas para acesso posterior
            self.send_command_inputs.append(send_input)
            self.send_buttons.append(send_button)
            self.auto_send_checkboxes.append(auto_send_cb)
            self.auto_send_config_buttons.append(config_timer_btn)

            # Configura o timer para auto-envio
            auto_timer = QTimer(self)
            auto_timer.timeout.connect(lambda checked=False, idx=i: self._send_command_for_auto_send(idx))
            self.auto_send_timers.append(auto_timer)
            
            auto_send_cb.stateChanged.connect(lambda state, idx=i: self._toggle_auto_send_timer(state, idx))
            
        left_panel.addWidget(send_group)
        main_layout.addLayout(left_panel, 2) # Adiciona o painel esquerdo ao layout principal

        # Painel direito (configuração de portas e controle de teste)
        right_panel = QVBoxLayout()

        # Grupo para configuração da porta serial principal
        self.serial_group = QGroupBox("Configuração Porta Serial Principal")
        serial_layout = QVBoxLayout(self.serial_group)

        port_layout_serial = QHBoxLayout()
        port_layout_serial.addWidget(QLabel("Nome da Porta:"))
        # Usa RefreshableComboBox para a porta serial principal
        self.serial_command_port_combobox = RefreshableComboBox(self._update_available_ports)
        self.serial_command_port_combobox.setMinimumWidth(150)
        port_layout_serial.addWidget(self.serial_command_port_combobox)
        port_layout_serial.addStretch()

        # Botão para expandir/recolher detalhes da configuração da porta
        self.serial_config_toggle_button = QToolButton(self)
        self.serial_config_toggle_button.setArrowType(Qt.ArrowType.DownArrow) # Começa expandido
        self.serial_config_toggle_button.setCheckable(True)
        self.serial_config_toggle_button.setChecked(True) # Começa expandido
        self.serial_config_toggle_button.clicked.connect(lambda: self._toggle_port_config_details("serial_command"))
        port_layout_serial.addWidget(self.serial_config_toggle_button)
        
        serial_layout.addLayout(port_layout_serial)

        # Widget para os detalhes da configuração da porta principal
        self.serial_details_widget = QWidget()
        self.serial_details_layout = QFormLayout(self.serial_details_widget)

        self.serial_baud_label = QLabel("Baud Rate:")
        self.serial_baud_combo = QComboBox()
        self.serial_baud_combo.addItems(["9600", "19200", "38400", "57600", "115200", "230400", "460800", "921600"])
        self.serial_baud_combo.setCurrentText("115200")
        self.serial_details_layout.addRow(self.serial_baud_label, self.serial_baud_combo)

        self.serial_data_bits_label = QLabel("Bits de Dados:")
        self.serial_data_bits_combo = QComboBox()
        self.serial_data_bits_combo.addItems(["5", "6", "7", "8"])
        self.serial_data_bits_combo.setCurrentText("8")
        self.serial_details_layout.addRow(self.serial_data_bits_label, self.serial_data_bits_combo)

        self.serial_parity_label = QLabel("Paridade:")
        self.serial_parity_combo = QComboBox()
        self.serial_parity_combo.addItems(["Nenhuma", "Ímpar", "Par", "Marca", "Espaço"])
        self.serial_parity_combo.setCurrentText("Nenhuma")
        self.serial_details_layout.addRow(self.serial_parity_label, self.serial_parity_combo)

        self.serial_handshake_label = QLabel("Controle de Fluxo:")
        self.serial_handshake_combo = QComboBox()
        self.serial_handshake_combo.addItems(["Nenhum", "RTS/CTS", "XON/XOFF"])
        self.serial_handshake_combo.setCurrentText("Nenhum")
        self.serial_details_layout.addRow(self.serial_handshake_label, self.serial_handshake_combo)
        
        self.serial_mode_label = QLabel("Modo:")
        self.serial_mode_combo = QComboBox()
        self.serial_mode_combo.addItems(["Free", "PortStore test", "Data", "Setup"])
        self.serial_mode_combo.setCurrentText("Free")
        self.serial_details_layout.addRow(self.serial_mode_label, self.serial_mode_combo)

        # Checkboxes DTR e RTS alinhados horizontalmente
        dtr_rts_layout_serial = QHBoxLayout()
        self.serial_dtr_checkbox = QCheckBox("DTR")
        self.serial_dtr_checkbox.setChecked(False) # Padrão desmarcado
        dtr_rts_layout_serial.addWidget(self.serial_dtr_checkbox)

        self.serial_rts_checkbox = QCheckBox("RTS")
        self.serial_rts_checkbox.setChecked(False) # Padrão desmarcado
        dtr_rts_layout_serial.addWidget(self.serial_rts_checkbox)
        dtr_rts_layout_serial.addStretch() # Empurra os checkboxes para a esquerda
        self.serial_details_layout.addRow("", dtr_rts_layout_serial) # Adiciona o layout ao formulário

        serial_layout.addWidget(self.serial_details_widget)
        self.serial_details_widget.setVisible(True) # Começa expandido

        self.connect_serial_command_button = QPushButton("Abrir Porta Principal")
        self.connect_serial_command_button.setObjectName("connect_serial_command_button") # Adiciona objectName para QSS
        self.connect_serial_command_button.clicked.connect(lambda: self._toggle_serial_connection("serial_command"))
        serial_layout.addWidget(self.connect_serial_command_button)
        self.serial_group.setLayout(serial_layout)
        right_panel.addWidget(self.serial_group)

        # Grupo para configuração da porta Modbus (opcional)
        self.modbus_serial_group = QGroupBox("Configuração Porta Modbus (Opcional)")
        modbus_serial_layout = QVBoxLayout(self.modbus_serial_group)

        port_layout_modbus = QHBoxLayout()
        port_layout_modbus.addWidget(QLabel("Nome da Porta:"))
        # Usa RefreshableComboBox para a porta Modbus
        self.modbus_port_combobox = RefreshableComboBox(self._update_available_ports)
        self.modbus_port_combobox.setMinimumWidth(150)
        port_layout_modbus.addWidget(self.modbus_port_combobox)
        port_layout_modbus.addStretch()

        # Botão para expandir/recolher detalhes da configuração da porta Modbus
        self.modbus_config_toggle_button = QToolButton(self)
        self.modbus_config_toggle_button.setArrowType(Qt.ArrowType.RightArrow) # Começa recolhido
        self.modbus_config_toggle_button.setCheckable(True)
        self.modbus_config_toggle_button.setChecked(False) # Começa recolhido
        self.modbus_config_toggle_button.clicked.connect(lambda: self._toggle_port_config_details("modbus"))
        port_layout_modbus.addWidget(self.modbus_config_toggle_button)

        modbus_serial_layout.addLayout(port_layout_modbus)

        # Widget para os detalhes da configuração da porta Modbus
        self.modbus_details_widget = QWidget()
        self.modbus_details_layout = QFormLayout(self.modbus_details_widget)

        self.modbus_baud_combo = QComboBox()
        self.modbus_baud_combo.addItems(["9600", "19200", "38400", "57600", "115200", "230400", "460800", "921600"])
        self.modbus_baud_combo.setCurrentText("9600")
        self.modbus_details_layout.addRow("Baud Rate:", self.modbus_baud_combo)

        self.modbus_data_bits_combo = QComboBox()
        self.modbus_data_bits_combo.addItems(["5", "6", "7", "8"])
        self.modbus_data_bits_combo.setCurrentText("8")
        self.modbus_details_layout.addRow("Bits de Dados:", self.modbus_data_bits_combo)

        self.modbus_parity_combo = QComboBox()
        self.modbus_parity_combo.addItems(["Nenhuma", "Ímpar", "Par", "Marca", "Espaço"])
        self.modbus_parity_combo.setCurrentText("Nenhuma")
        self.modbus_parity_combo.setCurrentText(self.test_modbus_settings.get("parity", "Nenhuma"))
        self.modbus_details_layout.addRow("Paridade:", self.modbus_parity_combo) # type: ignore

        self.modbus_handshake_combo = QComboBox()
        self.modbus_handshake_combo.addItems(["Nenhum", "RTS/CTS", "XON/XOFF"])
        self.modbus_handshake_combo.setCurrentText("Nenhum")
        self.modbus_handshake_combo.setCurrentText(self.test_modbus_settings.get("handshake", "Nenhum"))
        self.modbus_details_layout.addRow("Controle de Fluxo:", self.modbus_handshake_combo)
        
        self.modbus_mode_combo = QComboBox()
        self.modbus_mode_combo.addItems(["Free", "PortStore test", "Data", "Setup"])
        self.modbus_mode_combo.setCurrentText("Free")
        self.modbus_mode_combo.setCurrentText(self.test_modbus_settings.get("mode", "Free"))
        self.modbus_details_layout.addRow("Modo:", self.modbus_mode_combo)

        # Checkboxes DTR e RTS para Modbus alinhados horizontalmente
        dtr_rts_layout_modbus = QHBoxLayout()
        self.modbus_dtr_checkbox = QCheckBox("DTR")
        self.modbus_dtr_checkbox.setChecked(False) # Padrão desmarcado
        dtr_rts_layout_modbus.addWidget(self.modbus_dtr_checkbox)

        self.modbus_rts_checkbox = QCheckBox("RTS")
        self.modbus_rts_checkbox.setChecked(False) # Padrão desmarcado
        dtr_rts_layout_modbus.addWidget(self.modbus_rts_checkbox)
        dtr_rts_layout_modbus.addStretch() # Empurra os checkboxes para a esquerda
        self.modbus_details_layout.addRow("", dtr_rts_layout_modbus) # Adiciona o layout ao formulário

        modbus_serial_layout.addWidget(self.modbus_details_widget)
        self.modbus_details_widget.setVisible(False) # Começa recolhido

        self.connect_modbus_button = QPushButton("Abrir Porta Modbus")
        self.connect_modbus_button.setObjectName("connect_modbus_button") # Adiciona objectName para QSS
        self.connect_modbus_button.clicked.connect(lambda: self._toggle_serial_connection("modbus"))
        self.modbus_serial_group.setLayout(modbus_serial_layout)
        self.modbus_serial_group.setVisible(False) # Oculta o grupo Modbus por padrão
        modbus_serial_layout.addWidget(self.connect_modbus_button)
        right_panel.addWidget(self.modbus_serial_group)

        # Grupo para controle de teste
        test_control_group = QGroupBox("Controle de Teste")
        test_control_layout = QVBoxLayout(test_control_group)

        self.load_test_button = QPushButton("Carregar Arquivo de Teste")
        self.load_test_button.clicked.connect(self._load_test_file)
        test_control_layout.addWidget(self.load_test_button)

        self.start_test_button = QPushButton("Iniciar Teste")
        self.start_test_button.clicked.connect(self._prompt_for_test_info) # Abre o diálogo de informações da placa
        self.start_test_button.setEnabled(False) # Desabilitado até um teste ser carregado e portas conectadas
        test_control_layout.addWidget(self.start_test_button)

        self.stop_test_button = QPushButton("Parar Teste")
        self.stop_test_button.clicked.connect(self._stop_test)
        self.stop_test_button.setEnabled(False) # Desabilitado até um teste ser iniciado
        test_control_layout.addWidget(self.stop_test_button)

        self.test_status_label = QLabel("Status do Teste: Ocioso")
        
        self.test_status_label.setWordWrap(True)
        self.test_status_label.setMaximumWidth(400)
        test_control_layout.addWidget(self.test_status_label)
        
        # Rótulo para o status do Modo Fast - Removido da exibição
        #self.fast_mode_status_label = QLabel("Modo Fast: DESATIVADO")
        #        # test_control_layout.addWidget(self.fast_mode_status_label) # LINHA REMOVIDA

        # Grupo para exibir o progresso do teste
        self.test_progress_group = QGroupBox("Progresso do Teste")
        test_progress_layout = QVBoxLayout(self.test_progress_group)

        self.test_progress_label = QLabel("Status Geral: Nenhum teste em execução")
        self.test_progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.test_progress_label.setStyleSheet("font-weight: bold; padding: 5px; ")
        test_progress_layout.addWidget(self.test_progress_label)

        self.test_progress_list = QListWidget()
        self.test_progress_list.setMinimumHeight(150)
        self.test_progress_list.itemClicked.connect(self._handle_test_progress_item_click) # Permite re-testar passos falhos
        test_progress_layout.addWidget(self.test_progress_list)
        
        self.test_progress_group.setLayout(test_progress_layout)
        self.test_progress_group.setVisible(False) # Oculta o progresso do teste por padrão
        test_control_layout.addWidget(self.test_progress_group)

        test_control_group.setLayout(test_control_layout)
        right_panel.addWidget(test_control_group)

        right_panel.addStretch() # Adiciona um espaço flexível para empurrar os widgets para cima

        # Ajusta o fator de estiramento do painel direito para torná-lo mais compacto
        main_layout.addLayout(right_panel, 0) # Fator 0 significa que ele ocupará o mínimo de espaço necessário

        # Rótulo de versão e área de clique oculta
        self.version_label = QLabel(f"Versão: {self.VERSION}")
        self.version_label.setParent(self)
        self.version_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.version_label.setStyleSheet("color: gray; font-size: 10px;")
        
        self.hidden_click_area = ClickableLabel(self)
        self.hidden_click_area.setStyleSheet("background-color: rgba(0,0,0,0);") # Totalmente transparente
        self.hidden_click_area.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.hidden_click_area.clicked.connect(self._handle_hidden_button_click)

        self.tab_widget.currentChanged.connect(self._handle_tab_change) # Conecta para gerenciar visibilidade de elementos

        self.resizeEvent = self._custom_resize_event # Sobrescreve o evento de redimensionamento

        self._custom_resize_event(None)

        from configuracoes_widget import ConfiguracoesWidget

        # Define caminho seguro para salvar configurações
        import os
        settings_path = os.path.join(os.path.expanduser("~"), "Documents", "EmbTechSerial", "settings.json")
        os.makedirs(os.path.dirname(settings_path), exist_ok=True)

        from configuracoes_widget import ConfiguracoesWidget
        import os
        import json

        # Define caminho da rede e local
        rede_config_path = r"\\192.168.0.100\EmbTech\Config\settings.json"
        local_config_dir = os.path.join(os.path.expanduser("~"), "Documents", "EmbTechSerial")
        local_config_path = os.path.join(local_config_dir, "settings.json")

        # Verifica se o arquivo da rede está disponível
        if os.path.exists(rede_config_path):
            settings_path = rede_config_path
        else:
            os.makedirs(local_config_dir, exist_ok=True)
            settings_path = local_config_path

        # Criação da aba de configurações
        self.configuracoes_tab = ConfiguracoesWidget(settings_path)
        self.tab_widget.addTab(self.configuracoes_tab, "Configurações")
        self.tab_widget.setTabVisible(self.tab_widget.indexOf(self.configuracoes_tab), False)
        # Monitor de pastas de Logs Do Teste
        self._folder_observer = None
        self._folder_watch_path = ""
        self._folder_check_timer = QTimer(self)
        self._folder_check_timer.setInterval(5000)
        self._folder_check_timer.timeout.connect(self._ensure_folder_monitor_running)
        self._ensure_folder_monitor_running()
        self._folder_check_timer.start()
        # Lista para manter referências de diálogos de alerta
        self._active_alert_dialogs = []

        from relatorio_eficiencia_widget import RelatorioEficienciaWidget

        self.relatorio_widget = RelatorioEficienciaWidget(parent=self)
        self.tab_widget.addTab(self.relatorio_widget, "Relatórios")
        self.tab_widget.setTabVisible(self.tab_widget.indexOf(self.relatorio_widget), False)

    def _custom_resize_event(self, event):
        """
        Ajusta a posição do rótulo de versão e da área de clique oculta
        quando a janela é redimensionada.
        """
        self.version_label.setGeometry(self.width() - 150, self.height() - 30, 140, 20)
        self.hidden_click_area.setGeometry(self.width() - 150, self.height() - 30, 140, 20)
        super().resizeEvent(event) # Chama o método da classe base

    def _handle_tab_change(self, index):
        """
        Gerencia a visibilidade do rótulo de versão e da área de clique oculta
        com base na aba selecionada.
        """
        aba_oculta_versao = self.tab_widget.tabText(index) in ["Criador de Teste", "Configurações", "Relatórios"]
        self.version_label.setVisible(not aba_oculta_versao)
        self.hidden_click_area.setVisible(not aba_oculta_versao)

        if self.tab_widget.tabText(index) == "Criador de Teste":
            self._update_move_buttons_state() # Atualiza o estado dos botões de mover passos

    def _handle_hidden_button_click(self):
        """
        Incrementa um contador e, se atingir um limite, libera a aba "Criador de Teste".
        """
        self.hidden_button_click_count += 1

        if self.hidden_button_click_count >= 10: # Clicks necessários para liberar a aba
            if self.test_creator_tab_index != -1:
                self.tab_widget.setTabVisible(self.test_creator_tab_index, True)
                self.tab_widget.setTabVisible(self.tab_widget.indexOf(self.configuracoes_tab), True)
                self.tab_widget.setTabVisible(self.tab_widget.indexOf(self.relatorio_widget), True)

    def _ensure_folder_monitor_running(self):
        documents_path = self.configuracoes_tab.get_log_path()
        # Mesma regra de _generate_test_log_file: base = <get_log_path()>\"Logs Do Teste"
        base_log_dir = os.path.join(documents_path, "Logs Do Teste")

        if self._folder_watch_path != base_log_dir:
            self._start_folder_monitor(base_log_dir)

    def _start_folder_monitor(self, base_log_dir):
        try:
            self._stop_folder_monitor()
        except Exception:
            pass
        os.makedirs(base_log_dir, exist_ok=True)
        handler = _FolderHandler(self._on_new_dir_detected)
        self._folder_observer = Observer()
        self._folder_observer.schedule(handler, base_log_dir, recursive=False)
        self._folder_observer.start()
        self._folder_watch_path = base_log_dir

    def _stop_folder_monitor(self):
        if self._folder_observer is not None:
            try:
                self._folder_observer.stop()
                self._folder_observer.join(timeout=2.0)
            except Exception:
                pass
        self._folder_observer = None
        self._folder_watch_path = ""

    def _on_new_dir_detected(self, folder_name: str):
        operador = self._get_last_operator()
        maquina = os.environ.get("COMPUTERNAME", "") or os.environ.get("HOSTNAME", "")
        extra = []
        if operador:
            extra.append(f"Operador: {operador}")
        if maquina:
            extra.append(f"Máquina: {maquina}")
        suffix = f" | {' | '.join(extra)}" if extra else ""
        msg = f"Nova pasta detectada: {folder_name}{suffix}"
        try:
            self.log_message(msg, "alerta")
        except Exception:
            pass
        try:
            self._send_folder_whatsapp(msg)
        except Exception:
            pass

    def _get_last_operator(self) -> str:
        # Prioriza operador atual em memória
        nome = getattr(self, 'current_tester_name', '') or ''
        if nome:
            return nome
        # Tenta ler do arquivo local_settings.json usado pelo ConfiguracoesWidget
        try:
            local_settings_path = os.path.join(os.path.expanduser("~"), "Documents", "EmbTechSerial", "local_settings.json")
            if os.path.exists(local_settings_path):
                with open(local_settings_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    return str(data.get("last_user", "")).strip()
        except Exception:
            pass
        return ""

    def _send_folder_whatsapp(self, text: str):
        # Carrega do settings.json da aba Configurações, com fallback em variáveis de ambiente
        phone, apikey, url = self._load_callmebot_settings()
        if not phone:
            phone = os.environ.get("CALLMEBOT_PHONE", "").strip()
        if not apikey:
            apikey = os.environ.get("CALLMEBOT_APIKEY", "").strip()
        if not url:
            url = os.environ.get("CALLMEBOT_URL", "https://api.callmebot.com/whatsapp.php").strip()

        if not phone or not apikey:
            self.log_message("CALLMEBOT_PHONE/APIKEY não configurados. Ignorando envio WhatsApp.", "informacao")
            return

        params = {"phone": "".join(ch for ch in phone if ch.isdigit()), "text": text, "apikey": apikey}
        try:
            r = requests.get(url, params=params, timeout=15)
            if r.status_code == 200:
                pass
            else:
                self.log_message(f"Falha ao enviar WhatsApp: {r.status_code} {r.text}", "erro")
        except Exception as e:
            try:
                self.log_message(f"Erro ao enviar WhatsApp: {e}", "erro")
            except Exception:
                pass

    def _open_callmebot_settings_dialog(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Configurações do CallMeBot")
        layout = QFormLayout(dlg)

        phone, apikey, url = self._load_callmebot_settings()
        if not url:
            url = "https://api.callmebot.com/whatsapp.php"

        phone_edit = QLineEdit(phone)
        phone_edit.setPlaceholderText("Ex.: 553591089082 (apenas dígitos)")
        apikey_edit = QLineEdit(apikey)
        apikey_edit.setPlaceholderText("Sua API Key")
        url_edit = QLineEdit(url)

        layout.addRow("Telefone:", phone_edit)
        layout.addRow("API Key:", apikey_edit)
        layout.addRow("URL:", url_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        if dlg.exec():
            self._save_callmebot_settings(phone_edit.text().strip(), apikey_edit.text().strip(), url_edit.text().strip())
            self.log_message("Configurações do CallMeBot salvas.", "sistema")

    def _load_callmebot_settings(self):
        phone = ""
        apikey = ""
        url = ""
        try:
            settings_path = getattr(self.configuracoes_tab, 'settings_file', None)
            if settings_path and os.path.exists(settings_path):
                with open(settings_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    phone = str(data.get("callmebot_phone", "")).strip()
                    apikey = str(data.get("callmebot_api_key", "")).strip()
                    url = str(data.get("callmebot_url", "")).strip()
        except Exception:
            pass
        # Defaults se ausentes
        if not phone:
            phone = "553591089082"
        if not apikey:
            apikey = "9259918"
        if not url:
            url = "https://api.callmebot.com/whatsapp.php"
        return phone, apikey, url

    def _save_callmebot_settings(self, phone: str, apikey: str, url: str):
        try:
            settings_path = getattr(self.configuracoes_tab, 'settings_file', None)
            if not settings_path:
                return
            data = {}
            if os.path.exists(settings_path):
                try:
                    with open(settings_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                except Exception:
                    data = {}
            data["callmebot_phone"] = phone
            data["callmebot_api_key"] = apikey
            data["callmebot_url"] = url
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            try:
                self.log_message(f"Erro ao salvar CallMeBot: {e}", "erro")
            except Exception:
                pass

    def closeEvent(self, event):
        try:
            self._stop_folder_monitor()
        except Exception:
            pass
        super().closeEvent(event)
        

    def _check_fast_mode_activation(self):
        """Verifica se o código secreto digitado ativa o modo Fast."""
        if not self.fast_mode_active and self.fast_mode_code_input.text().strip() == self.fast_mode_secret_code:
            self.fast_mode_active = False
            self.fast_mode_code_input.setEnabled(False)
            #self.log_message("Modo Fast ATIVADO com sucesso. Agora você pode usar testes otimizados.", "sistema")
            self.fast_mode_code_input.setText("")
            self._update_fast_mode_status_label()
                #self.fast_mode_active = True
                #self.fast_mode_code_input.setEnabled(False)
                #self.log_message("Modo Fast ATIVADO com sucesso. Agora você pode usar testes otimizados.", "sistema")

    def _init_test_creator_ui(self):
        """
        Inicializa os componentes da interface do usuário para a aba "Criador de Teste".
        """
        self.test_creator_tab = QWidget()
        test_creator_layout = QVBoxLayout(self.test_creator_tab)
        
        # Grupo para configurações de comunicação do teste
        port_config_group = QGroupBox("Configurações de Comunicação para o Teste")
        port_config_layout = QVBoxLayout(port_config_group)

        self.configure_test_ports_button = QPushButton("Configurar Portas do Teste")
        self.configure_test_ports_button.clicked.connect(self._open_test_port_config_dialog)
        port_config_layout.addWidget(self.configure_test_ports_button)

        self.test_serial_settings_label = QLabel("Porta Principal: Não configurada")
        self.test_modbus_settings_label = QLabel("Porta Modbus: Não configurada")
        port_config_layout.addWidget(self.test_serial_settings_label)
        port_config_layout.addWidget(self.test_modbus_settings_label)

        test_creator_layout.addWidget(port_config_group)

        # Configurações do Modo Fast (agora como um campo direto, sem QGroupBox)
        fast_mode_layout = QHBoxLayout()
        fast_mode_layout.addWidget(QLabel("Código Secreto (4 dígitos):"))
        self.fast_mode_code_input = QLineEdit()
        self.fast_mode_code_input.setPlaceholderText("Digite 4 dígitos para o código secreto")
        self.fast_mode_code_input.setMaxLength(4)
        self.fast_mode_code_input.setFixedWidth(200) # Ajusta a largura
        # Validador para garantir apenas 4 dígitos numéricos
        self.fast_mode_code_input.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d{0,4}$")))
        self.fast_mode_code_input.textChanged.connect(self._check_fast_mode_activation)
        fast_mode_layout.addWidget(self.fast_mode_code_input)
        fast_mode_layout.addStretch(1) # Empurra o campo para a esquerda
        self.fast_mode_code_input.textChanged.connect(self._check_fast_mode_activation)
        test_creator_layout.addLayout(fast_mode_layout)

        # Grupo para adicionar/editar passos de teste
        edit_step_group = QGroupBox("Adicionar/Editar Passo de Teste")
        edit_step_layout = QFormLayout(edit_step_group)

        self.step_name_input = QLineEdit()
        edit_step_layout.addRow("Nome do Passo:", self.step_name_input)

        # Rádio botões para selecionar o tipo de passo (automático, manual, tempo de espera, Modbus)
        step_type_layout = QHBoxLayout()
        self.radio_comando_auto = QRadioButton("Comando/Validação Automática")
        self.radio_instrucao_manual = QRadioButton("Instrução Manual")
        self.radio_tempo_espera = QRadioButton("Tempo de Espera")
        self.radio_modbus_comando = QRadioButton("Comando Modbus") # Novo rádio botão para Modbus
        self.radio_gravar_ns = QRadioButton("Gravar Número de Série") # Novo rádio botão para Gravação de NS
        self.radio_gravar_placa = QRadioButton("Gravar Placa")
        step_type_layout.addWidget(self.radio_comando_auto)
        step_type_layout.addWidget(self.radio_instrucao_manual)
        step_type_layout.addWidget(self.radio_tempo_espera)
        step_type_layout.addWidget(self.radio_modbus_comando) # Adiciona o novo rádio botão
        step_type_layout.addWidget(self.radio_gravar_ns)
        step_type_layout.addWidget(self.radio_gravar_placa)
        step_type_layout.addStretch()
        edit_step_layout.addRow("Tipo de Passo:", step_type_layout)
        
        # Grupo de campos para passos de comando/validação automática
        self.command_validation_group = QWidget()
        self.command_validation_layout = QFormLayout(self.command_validation_group)

        self.step_port_type_combo = QComboBox()
        self.step_port_type_combo.addItems(["Porta Principal (Dados e Comandos)", "Porta Modbus (Texto e Validação)"]) # Removido Modbus daqui
        self.step_port_type_combo.setCurrentText("Porta Principal (Dados e Comandos)")
        self.command_validation_layout.addRow("Usar Porta:", self.step_port_type_combo)

        self.step_command_input = QLineEdit()
        self.command_validation_layout.addRow("Comando para Enviar:", self.step_command_input)

        # RTC (Data/Hora) feature: enable + mode + pattern
        self.rtc_feature_cb = QCheckBox("RTC Automático (Data/Hora)")
        self.rtc_feature_cb.stateChanged.connect(lambda _state: self._toggle_rtc_fields())
        self.command_validation_layout.addRow("RTC Automático:", self.rtc_feature_cb)

        rtc_mode_layout = QHBoxLayout()
        self.rtc_mode_combo = QComboBox()
        self.rtc_mode_combo.addItems(["Escrita", "Leitura"])  # Escrita => SET_RTC=..., Leitura => READ_RTC
        self.rtc_mode_combo.currentIndexChanged.connect(lambda _i: self._toggle_rtc_fields())
        rtc_mode_layout.addWidget(self.rtc_mode_combo)

        self.rtc_pattern_input = QLineEdit()
        self.rtc_pattern_input.setPlaceholderText("Padrão RTC, ex: SET_RTC=YYYY-MM-DD HH:MM:SS; ou READ_RTC;")
        rtc_mode_layout.addWidget(self.rtc_pattern_input)

        self.command_validation_layout.addRow("Modo/Padrão RTC:", rtc_mode_layout)
        # Inicializa estados
        self._toggle_rtc_fields()

        self.step_expect_response_cb = QCheckBox("Aguardar Resposta?")
        self.command_validation_layout.addRow("Aguardar Resposta?:", self.step_expect_response_cb)
        self.step_expect_response_cb.stateChanged.connect(self._toggle_validation_fields)

        self.step_timeout_input = QSpinBox()
        self.step_timeout_input.setMinimum(10)
        self.step_timeout_input.setMaximum(60000)
        self.step_timeout_input.setValue(1000)
        self.step_timeout_input.setSuffix(" ms")
        self.command_validation_layout.addRow("Tempo Limite para Resposta (ms):", self.step_timeout_input)
        self.step_timeout_input.setEnabled(False)

        self.step_validation_type_combo = QComboBox()
        # Removido "Modbus" daqui
        self.step_validation_type_combo.addItems(["Nenhuma", "Texto Exato", "Número em Faixa", "Texto Simples com Número", "Texto com Vários Números", "Número de Série", "Data/Hora"])
        self.command_validation_layout.addRow("Como Validar a Resposta?:", self.step_validation_type_combo)
        self.step_validation_type_combo.setEnabled(False)
        self.step_validation_type_combo.currentIndexChanged.connect(self._toggle_validation_params_fields)

        # Campos para validação de "Texto Exato"
        self.exact_string_label = QLabel("Texto Exato:")
        self.exact_string_input = QLineEdit()
        self.exact_string_input.setPlaceholderText(r"Ex: OK\r\n")
        self.command_validation_layout.addRow(self.exact_string_label, self.exact_string_input)
        
        # Campos para validação de "Texto Simples com Número"
        self.simplified_regex_label = QLabel("Texto com o valor a ser capturado (use [VALOR]):")
        self.simplified_regex_input = QLineEdit()
        self.simplified_regex_input.setPlaceholderText("Ex: Temperatura: [VALOR] °C")
        self.command_validation_layout.addRow(self.simplified_regex_label, self.simplified_regex_input)

        # Campos para validação de "Número em Faixa" (e também para texto simples com número)
        self.numeric_range_label = QLabel("Faixa do Valor Esperado (Mín/Máx):")
        
        self.numeric_min_label = QLabel("Mín:")
        self.numeric_max_label = QLabel("Máx:")

        self.numeric_min_input = QDoubleSpinBox()
        self.numeric_min_input.setMinimum(-999999.999)
        self.numeric_min_input.setMaximum(999999.999)
        self.numeric_min_input.setDecimals(3)
        self.numeric_min_input.setSingleStep(0.1)
        self.numeric_min_input.setValue(0.0)

        self.numeric_max_input = QDoubleSpinBox()
        self.numeric_max_input.setMinimum(-999999.999)
        self.numeric_max_input.setMaximum(999999.999)
        self.numeric_max_input.setDecimals(3)
        self.numeric_max_input.setSingleStep(0.1)
        self.numeric_max_input.setValue(0.0)

        numeric_layout = QHBoxLayout()
        numeric_layout.addWidget(self.numeric_min_label)
        numeric_layout.addWidget(self.numeric_min_input)
        numeric_layout.addWidget(self.numeric_max_label)
        numeric_layout.addWidget(self.numeric_max_input)
        self.command_validation_layout.addRow(self.numeric_range_label, numeric_layout)

        self._hide_all_validation_fields() # Oculta todos os campos de validação por padrão

        edit_step_layout.addRow(self.command_validation_group)
        self.command_validation_group.setVisible(True) # Visível por padrão ao iniciar

        # Grupo de campos para passos de instrução manual
        self.instruction_group = QWidget()
        self.instruction_layout = QFormLayout(self.instruction_group)
        self.step_instruction_text = QTextEdit()
        self.step_instruction_text.setPlaceholderText("Digite a instrução que aparecerá para o usuário. (Ex: 'Conecte o cabo USB na placa.')")
        self.instruction_layout.addRow("Mensagem de Instrução:", self.step_instruction_text)
        
        image_path_layout = QHBoxLayout()
        self.step_image_path_input = QLineEdit()
        self.step_image_path_input.setPlaceholderText("Caminho para a imagem (opcional)")
        image_path_layout.addWidget(self.step_image_path_input)
        self.browse_image_button = QPushButton("Procurar Imagem...")
        self.browse_image_button.clicked.connect(self._browse_image_file)
        image_path_layout.addWidget(self.browse_image_button)
        self.instruction_layout.addRow("Imagem (Opcional):", image_path_layout)

        edit_step_layout.addRow(self.instruction_group)
        self.instruction_group.setVisible(False) # Oculta o grupo de instrução manual por padrão

        # Grupo de campos para passo 'Gravar Placa'
        self.flash_board_group = QWidget()
        self.flash_board_layout = QFormLayout(self.flash_board_group)
        self.flash_question_input = QLineEdit()
        self.flash_question_input.setPlaceholderText("Pergunta ao operador. Ex: A placa está gravada?")
        self.flash_board_layout.addRow("Pergunta:", self.flash_question_input)

        self.flash_instruction_text = QTextEdit()
        self.flash_instruction_text.setPlaceholderText("Explique como gravar a placa. Este texto aparecerá no popup com o botão Gravar.")
        self.flash_board_layout.addRow("Instruções:", self.flash_instruction_text)

        flash_cmd_layout = QHBoxLayout()
        self.flash_cmd_path_input = QLineEdit()
        self.flash_cmd_path_input.setPlaceholderText("Caminho do arquivo .cmd a executar")
        flash_cmd_layout.addWidget(self.flash_cmd_path_input)
        self.flash_browse_cmd_button = QPushButton("Procurar .cmd...")
        self.flash_browse_cmd_button.clicked.connect(self._browse_flash_cmd_file)
        flash_cmd_layout.addWidget(self.flash_browse_cmd_button)
        self.flash_board_layout.addRow("Arquivo de Gravação (.cmd):", flash_cmd_layout)

        edit_step_layout.addRow(self.flash_board_group)
        self.flash_board_group.setVisible(False)

        # Segundo arquivo opcional de gravação (.cmd)
        second_cmd_layout = QHBoxLayout()
        self.flash_cmd_path_input2 = QLineEdit()
        self.flash_cmd_path_input2.setPlaceholderText("Caminho do segundo arquivo .cmd (opcional)")
        second_cmd_layout.addWidget(self.flash_cmd_path_input2)
        self.flash_browse_cmd_button2 = QPushButton("Procurar 2º .cmd...")
        self.flash_browse_cmd_button2.clicked.connect(self._browse_flash_cmd_file2)
        second_cmd_layout.addWidget(self.flash_browse_cmd_button2)
        self.flash_board_layout.addRow("Segundo Arquivo (.cmd) (Opcional):", second_cmd_layout)

        # Novo grupo de campos para passos de tempo de espera
        self.wait_time_group = QWidget()
        self.wait_time_layout = QFormLayout(self.wait_time_group)
        self.wait_duration_input = QSpinBox()
        self.wait_duration_input.setMinimum(1)
        self.wait_duration_input.setMaximum(3600) # Máximo de 1 hora (3600 segundos)
        self.wait_duration_input.setValue(10) # Padrão de 10 segundos
        self.wait_duration_input.setSuffix(" segundos")
        self.wait_time_layout.addRow("Duração da Espera:", self.wait_duration_input)

        edit_step_layout.addRow(self.wait_time_group)
        self.wait_time_group.setVisible(False) # Oculta o grupo de tempo de espera por padrão

        # NOVO: Grupo de campos para passos de Comando Modbus
        self.modbus_command_group = QWidget()
        modbus_command_layout = QVBoxLayout(self.modbus_command_group)

        self.modbus_table_widget = QTableWidget()
        self.modbus_table_widget.setColumnCount(9) # ID Escravo, Função, Endereço Reg., Qtd., Tipo, Valor Escrita, Valor Esperado, Limite Mín., Limite Máx.
        self.modbus_table_widget.setHorizontalHeaderLabels([
            "ID Escravo", "Função", "Endereço Reg.", "Qtd.", "Tipo", 
            "Valor Escrita", "Valor Esperado", "Limite Mín.", "Limite Máx."
        ])
        self.modbus_table_widget.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.modbus_table_widget.horizontalHeader().setStretchLastSection(True)
        self.modbus_table_widget.setRowCount(1)
        self._init_modbus_table_row(0)

        self.add_modbus_row_button = QPushButton("Adicionar Linha Modbus")
        self.add_modbus_row_button.clicked.connect(self._add_modbus_table_row)
        self.remove_modbus_row_button = QPushButton("Remover Linha Modbus")
        self.remove_modbus_row_button.clicked.connect(self._remove_modbus_table_row)

        modbus_table_buttons_layout = QHBoxLayout()
        modbus_table_buttons_layout.addWidget(self.add_modbus_row_button)
        modbus_table_buttons_layout.addWidget(self.remove_modbus_row_button)
        
        modbus_command_layout.addWidget(self.modbus_table_widget)
        modbus_command_layout.addLayout(modbus_table_buttons_layout)

        edit_step_layout.addRow(self.modbus_command_group)
        self.modbus_command_group.setVisible(False) # Oculta o grupo Modbus por padrão


        # Botões de ação (adicionar/atualizar passo)
        action_buttons_layout = QHBoxLayout()
        self.add_step_button = QPushButton("Adicionar Novo Passo")
        self.add_step_button.clicked.connect(self._add_new_test_step)
        action_buttons_layout.addWidget(self.add_step_button)

        self.update_step_button = QPushButton("Atualizar Passo Selecionado")
        self.update_step_button.clicked.connect(self._update_selected_test_step)
        self.update_step_button.setEnabled(False) # Desabilitado até um passo ser selecionado para edição
        action_buttons_layout.addWidget(self.update_step_button)

        edit_step_layout.addRow(action_buttons_layout)

        test_creator_layout.addWidget(edit_step_group)

        # Botões para mover passos na lista
        move_step_buttons_layout = QHBoxLayout()
        self.move_step_up_button = QPushButton("Mover para Cima")
        self.move_step_up_button.clicked.connect(self._move_step_up)
        self.move_step_up_button.setEnabled(False)
        move_step_buttons_layout.addWidget(self.move_step_up_button)

        self.move_step_down_button = QPushButton("Mover para Baixo")
        self.move_step_down_button.clicked.connect(self._move_step_down)
        self.move_step_down_button.setEnabled(False)
        move_step_buttons_layout.addWidget(self.move_step_down_button)
        
        test_creator_layout.addLayout(move_step_buttons_layout)

        # Lista de passos de teste atuais (usará TestStepListItemWidget)
        self.test_steps_list = QListWidget()
        self.test_steps_list.itemClicked.connect(self._load_step_for_editing) # Carrega o passo selecionado para edição
        self.test_steps_list.itemSelectionChanged.connect(self._update_move_buttons_state) # Atualiza o estado dos botões de mover
        test_creator_layout.addWidget(QLabel("Passos do Teste Atuais (Marque para Modo Fast):"))
        test_creator_layout.addWidget(self.test_steps_list)

        # Botões para gerenciar o arquivo de teste
        test_file_buttons_layout = QHBoxLayout()
        self.save_test_button = QPushButton("Salvar Teste Atual")
        self.save_test_button.clicked.connect(self._save_test_file)
        test_file_buttons_layout.addWidget(self.save_test_button)

        self.remove_step_button = QPushButton("Remover Passo Selecionado")
        self.remove_step_button.clicked.connect(self._remove_selected_test_step)
        self.remove_step_button.setEnabled(False)
        test_file_buttons_layout.addWidget(self.remove_step_button)

        self.clear_test_button = QPushButton("Limpar Todos os Passos")
        self.clear_test_button.clicked.connect(self._clear_test_steps)
        self.clear_test_button.setEnabled(False)
        test_file_buttons_layout.addWidget(self.clear_test_button)

        test_creator_layout.addLayout(test_file_buttons_layout)

        # Adiciona a aba "Criador de Teste" e a oculta por padrão
        self.test_creator_tab_index = self.tab_widget.addTab(self.test_creator_tab, "Criador de Teste")
        self.tab_widget.setTabVisible(self.test_creator_tab_index, False) 

        # Configurações iniciais dos campos de validação e lista de passos
        self._toggle_validation_params_fields(self.step_validation_type_combo.currentIndex())
        self._update_test_steps_list()
        self._update_test_port_settings_display()

    def _init_modbus_table_row(self, row):
        """Inicializa os widgets para uma nova linha na tabela Modbus."""
        # Coluna 0: ID Escravo
        slave_id_spinbox = QSpinBox()
        slave_id_spinbox.setMinimum(1)
        slave_id_spinbox.setMaximum(247)
        slave_id_spinbox.setValue(1)
        self.modbus_table_widget.setCellWidget(row, 0, slave_id_spinbox)

        # Coluna 1: Função
        function_combo = QComboBox()
        function_combo.addItems([
            "Read Coils (0x01)", 
            "Read Discrete Inputs (0x02)", 
            "Read Holding Registers (0x03)", 
            "Read Input Registers (0x04)",
            "Write Single Coil (0x05)",
            "Write Single Register (0x06)"
        ])
        function_combo.currentIndexChanged.connect(lambda index, r=row: self._toggle_modbus_table_fields(r, index))
        self.modbus_table_widget.setCellWidget(row, 1, function_combo)

        # Coluna 2: Endereço Reg.
        address_spinbox = QSpinBox()
        address_spinbox.setMinimum(0)
        address_spinbox.setMaximum(65535)
        address_spinbox.setValue(0)
        self.modbus_table_widget.setCellWidget(row, 2, address_spinbox)

        # Coluna 3: Qtd.
        quantity_spinbox = QSpinBox()
        quantity_spinbox.setMinimum(1)
        quantity_spinbox.setMaximum(2000) # Max coils
        quantity_spinbox.setValue(1)
        self.modbus_table_widget.setCellWidget(row, 3, quantity_spinbox)

        # Coluna 4: Tipo (para leitura/escrita de valor)
        type_combo = QComboBox()
        type_combo.addItems(["Coil (Boolean)", "Register (16-bit)", "Float (32-bit)", "Double (64-bit)"])
        type_combo.setCurrentText("Register (16-bit)")
        self.modbus_table_widget.setCellWidget(row, 4, type_combo)

        # Coluna 5: Valor Escrita (para funções de escrita)
        write_value_input = QLineEdit()
        write_value_input.setPlaceholderText("Ex: 0x01, 123, 0b1")
        self.modbus_table_widget.setCellWidget(row, 5, write_value_input)

        # Coluna 6: Valor Esperado (para validação) - pode ser uma string ou um número
        expected_value_input = QLineEdit()
        expected_value_input.setPlaceholderText("Ex: 25, true, 0x1A")
        self.modbus_table_widget.setCellWidget(row, 6, expected_value_input)

        # Coluna 7: Limite Mín.
        min_limit_spinbox = QDoubleSpinBox()
        min_limit_spinbox.setMinimum(-999999.999)
        min_limit_spinbox.setMaximum(999999.999)
        min_limit_spinbox.setDecimals(3)
        min_limit_spinbox.setSingleStep(0.1)
        min_limit_spinbox.setValue(0.0)
        self.modbus_table_widget.setCellWidget(row, 7, min_limit_spinbox)

        # Coluna 8: Limite Máx.
        max_limit_spinbox = QDoubleSpinBox()
        max_limit_spinbox.setMinimum(-999999.999)
        max_limit_spinbox.setMaximum(999999.999)
        max_limit_spinbox.setDecimals(3)
        max_limit_spinbox.setSingleStep(0.1)
        max_limit_spinbox.setValue(0.0)
        self.modbus_table_widget.setCellWidget(row, 8, max_limit_spinbox)

        # Inicializa a visibilidade dos campos com base na função padrão (Read Coils)
        self._toggle_modbus_table_fields(row, function_combo.currentIndex())

    def _add_modbus_table_row(self):
        """Adiciona uma nova linha à tabela de configuração Modbus."""
        row_count = self.modbus_table_widget.rowCount()
        self.modbus_table_widget.insertRow(row_count)
        self._init_modbus_table_row(row_count)
        self.modbus_table_widget.resizeColumnsToContents()

    def _remove_modbus_table_row(self):
        """Remove a linha selecionada da tabela de configuração Modbus."""
        current_row = self.modbus_table_widget.currentRow()
        if current_row >= 0:
            self.modbus_table_widget.removeRow(current_row)
            if self.modbus_table_widget.rowCount() == 0:
                self._add_modbus_table_row() # Garante que sempre haja pelo menos uma linha

    def _toggle_modbus_table_fields(self, row, function_index):
        """
        Alterna a visibilidade e habilitação de campos na linha da tabela Modbus
        com base na função Modbus selecionada.
        """
        function_code_text = self.modbus_table_widget.cellWidget(row, 1).itemText(function_index)
        is_write_function = "Write" in function_code_text
        is_read_function = "Read" in function_code_text

        # Coluna 3: Qtd. (visível apenas para funções de leitura)
        self.modbus_table_widget.cellWidget(row, 3).setVisible(is_read_function)
        self.modbus_table_widget.cellWidget(row, 3).setEnabled(is_read_function)
        if not is_read_function:
            self.modbus_table_widget.cellWidget(row, 3).setValue(1) # Reseta a quantidade

        # Coluna 5: Valor Escrita (visível apenas para funções de escrita)
        self.modbus_table_widget.cellWidget(row, 5).setVisible(is_write_function)
        self.modbus_table_widget.cellWidget(row, 5).setEnabled(is_write_function)
        if not is_write_function:
            self.modbus_table_widget.cellWidget(row, 5).clear()

        # Coluna 4: Tipo (sempre visível, mas a lógica de validação muda)
        # Coluna 6: Valor Esperado (sempre visível, mas o significado muda)
        # Coluna 7 e 8: Limite Mín. e Máx. (sempre visíveis, mas a lógica de validação muda)

        # Ajusta o máximo de quantidade com base no tipo de dado (Coil vs Register)
        if is_read_function:
            type_combo = self.modbus_table_widget.cellWidget(row, 4)
            value_type = type_combo.currentText()
            if "Coil" in value_type:
                self.modbus_table_widget.cellWidget(row, 3).setMaximum(2000) # Max coils
            else:
                self.modbus_table_widget.cellWidget(row, 3).setMaximum(125) # Max registers


    def _hide_all_validation_fields(self):
        """Oculta todos os campos de validação no criador de teste."""
        fields = [
            self.exact_string_label, self.exact_string_input,
            self.simplified_regex_label, self.simplified_regex_input,
            self.numeric_range_label, self.numeric_min_input, self.numeric_max_input,
            self.numeric_min_label, self.numeric_max_label,
            # self.modbus_table_container # Removido daqui
        ]
        for field in fields:
            field.setVisible(False)

    def _toggle_step_type_fields(self):
        """
        Alterna a visibilidade dos campos de entrada no criador de teste
        com base no tipo de passo selecionado (automático, manual, tempo de espera, Modbus).
        """
        is_command_auto = self.radio_comando_auto.isChecked()
        is_manual_instruction = self.radio_instrucao_manual.isChecked()
        is_wait_time = self.radio_tempo_espera.isChecked()
        is_modbus_command = self.radio_modbus_comando.isChecked() # Novo estado para Modbus
        is_write_serial = getattr(self, 'radio_gravar_ns', None) and self.radio_gravar_ns.isChecked()
        is_flash_board = getattr(self, 'radio_gravar_placa', None) and self.radio_gravar_placa.isChecked()

        self.command_validation_group.setVisible(is_command_auto)
        self.instruction_group.setVisible(is_manual_instruction)
        self.wait_time_group.setVisible(is_wait_time)
        self.modbus_command_group.setVisible(is_modbus_command) # Controla a visibilidade do novo grupo Modbus
        if hasattr(self, 'flash_board_group'):
            self.flash_board_group.setVisible(is_flash_board)
        # Novo tipo 'Gravar NS' não requer campos adicionais visíveis
        if is_write_serial:
            # Força uso da porta principal e pré-preenche o comando modelo na UI para referência
            self.step_port_type_combo.setCurrentText("Porta Principal (Dados e Comandos)")
            self.step_command_input.setText("SET_NS=XXXXX/XXX;")

        # Limpa os campos de outros tipos de passo quando a seleção muda
        if not is_command_auto:
            self._clear_command_validation_fields()
        if not is_manual_instruction:
            self.step_instruction_text.clear()
            self.step_image_path_input.clear()
        if not is_flash_board:
            if hasattr(self, 'flash_question_input'): self.flash_question_input.clear()
            if hasattr(self, 'flash_instruction_text'): self.flash_instruction_text.clear()
            if hasattr(self, 'flash_cmd_path_input'): self.flash_cmd_path_input.clear()
            if hasattr(self, 'flash_cmd_path_input2'): self.flash_cmd_path_input2.clear()
        if not is_wait_time:
            self.wait_duration_input.setValue(10) # Reseta para o valor padrão
        if not is_modbus_command:
            self._clear_modbus_command_fields() # Limpa os campos Modbus
        # Nenhuma limpeza específica necessária para 'Gravar NS'

    def _browse_image_file(self):
        """
        Abre um diálogo de seleção de arquivo para escolher uma imagem.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Imagem", "", "Arquivos de Imagem (*.png *.jpg *.jpeg *.bmp *.gif);;Todos os Arquivos (*)")
        if file_path:
            self.step_image_path_input.setText(file_path)

    def _browse_flash_cmd_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo de Gravação", "", "Scripts (*.cmd *.bat);;Executáveis (*.exe);;Todos os Arquivos (*)")
        if file_path:
            self.flash_cmd_path_input.setText(file_path)

    def _browse_flash_cmd_file2(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Segundo Arquivo de Gravação", "", "Scripts (*.cmd *.bat);;Executáveis (*.exe);;Todos os Arquivos (*)")
        if file_path:
            self.flash_cmd_path_input2.setText(file_path)

    def _toggle_validation_fields(self, state):
        """
        Alterna a visibilidade e habilitação dos campos de validação de resposta
        com base no estado do checkbox "Aguardar Resposta?".
        """
        enabled = (state == Qt.CheckState.Checked.value)
        self.step_timeout_input.setEnabled(enabled)
        self.step_validation_type_combo.setEnabled(enabled)
        if not enabled:
            self.step_validation_type_combo.setCurrentIndex(0) # Reseta o tipo de validação para "Nenhuma"
            self._toggle_validation_params_fields(0) # Oculta os campos de parâmetros de validação

    # Utilidades leves para evitar atualizações desnecessárias na UI
    def _set_combo_if_changed(self, combo, text):
        try:
            if text is None:
                return
            current = combo.currentText()
            new_text = str(text)
            if current != new_text and combo.findText(new_text) != -1:
                from PyQt6.QtCore import QSignalBlocker
                blocker = QSignalBlocker(combo)
                combo.setCurrentText(new_text)
                del blocker
        except Exception:
            pass

    def _set_lineedit_if_changed(self, line_edit, text):
        try:
            if text is None:
                return
            current = line_edit.text()
            new_text = str(text)
            if current != new_text:
                line_edit.setText(new_text)
        except Exception:
            pass

    def _set_label_if_changed(self, label, text):
        try:
            if text is None:
                return
            current = label.text()
            new_text = str(text)
            if current != new_text:
                label.setText(new_text)
        except Exception:
            pass

    def _rx_datetime(self):
        # Compila e cacheia regex para YYYY-MM-DD HH:MM:SS
        if not hasattr(self, "_rx_datetime_compiled") or self._rx_datetime_compiled is None:
            self._rx_datetime_compiled = re.compile(r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})")
        return self._rx_datetime_compiled

    def _toggle_rtc_fields(self):
        """Controla visibilidade e valores padrão dos campos de RTC (Data/Hora)."""
        enabled = self.rtc_feature_cb.isChecked() if hasattr(self, "rtc_feature_cb") else False
        if hasattr(self, "rtc_mode_combo"):
            self.rtc_mode_combo.setEnabled(enabled)
        if hasattr(self, "rtc_pattern_input"):
            self.rtc_pattern_input.setEnabled(enabled)
            if enabled:
                mode = self.rtc_mode_combo.currentText() if hasattr(self, "rtc_mode_combo") else "Escrita"
                # Preseta sempre o campo de padrão conforme o modo selecionado
                if mode == "Escrita":
                    self._set_lineedit_if_changed(self.rtc_pattern_input, "SET_RTC=YYYY-MM-DD HH:MM:SS;")
                else: # Leitura
                    self._set_lineedit_if_changed(self.rtc_pattern_input, "READ_RTC;")
            else:
                # mantém o texto, apenas desabilita
                pass

    def _build_command_with_rtc(self, base_command, use_rtc, rtc_mode, rtc_pattern):
        """Retorna o comando final considerando o recurso de RTC.
        Escrita: substitui 'YYYY-MM-DD HH:MM:SS' ou '[RTC]' pela data/hora atual.
        Leitura: retorna o padrão configurado (ex.: READ_RTC;).
        """
        try:
            if not use_rtc:
                return base_command
            pattern = rtc_pattern or ""
            if str(rtc_mode).lower().startswith("escr"):
                now_dt = datetime.now()
                now_str = now_dt.strftime("%Y-%m-%d %H:%M:%S")
                final_cmd = pattern.replace("[RTC]", now_str)
                if final_cmd == pattern:
                    final_cmd = pattern.replace("YYYY-MM-DD HH:MM:SS", now_str)
                # Guarda quando foi feito o último SET_RTC para cálculo de tolerância percentual
                try:
                    self.last_rtc_set_time = now_dt
                except Exception:
                    pass
                return final_cmd
            else:
                return pattern or base_command
        except Exception:
            return base_command or rtc_pattern

    def _toggle_validation_params_fields(self, index):
        """
        Alterna a visibilidade dos campos de parâmetros de validação
        com base no tipo de validação selecionado no combobox.
        """
        self._hide_all_validation_fields() # Oculta todos os campos primeiro

        validation_type = self.step_validation_type_combo.currentText()
        
        if validation_type == "Texto Exato":
            self.exact_string_label.setVisible(True)
            self.exact_string_input.setVisible(True)
        elif validation_type == "Número em Faixa":
            self.numeric_range_label.setVisible(True)
            self.numeric_min_input.setVisible(True)
            self.numeric_max_input.setVisible(True)
            self.numeric_min_label.setVisible(True) 
            self.numeric_max_label.setVisible(True) 
        elif validation_type == "Texto Simples com Número":
            self.simplified_regex_label.setVisible(True)
            self.simplified_regex_input.setVisible(True)
            self.numeric_range_label.setVisible(True)
            self.numeric_min_input.setVisible(True)
            self.numeric_max_input.setVisible(True)
            self.numeric_min_label.setVisible(True) 
            self.numeric_max_label.setVisible(True) 
        elif validation_type == "Texto com Vários Números":
            self.simplified_regex_label.setVisible(True)
            self.simplified_regex_input.setVisible(True)
            self.numeric_range_label.setVisible(True)
            self.numeric_min_input.setVisible(True)
            self.numeric_max_input.setVisible(True)
            self.numeric_min_label.setVisible(True)
            self.numeric_max_label.setVisible(True)
        elif validation_type == "Número de Série":
            # Não há parâmetros para este tipo; nada a mostrar
            pass
        elif validation_type == "Data/Hora":
            # Sem parâmetros adicionais
            pass
        # Removido o caso "Modbus" daqui

    def _update_test_steps_list(self, select_index=-1):
        """
        Atualiza a QListWidget que exibe os passos do teste.
        Controla também a habilitação de botões relacionados à lista.
        Agora usa TestStepListItemWidget para incluir o checkbox.
        """
        self.test_steps_list.clear()
        if not self.current_test_steps:
            self.test_steps_list.addItem("Nenhum passo de teste adicionado ainda.")
            self.test_steps_list.setEnabled(False)
            self.move_step_up_button.setEnabled(False)
            self.move_step_down_button.setEnabled(False)
            self.remove_step_button.setEnabled(False)
            self.clear_test_button.setEnabled(False)
            self.save_test_button.setEnabled(False)
            return
        
        self.test_steps_list.setEnabled(True)

        for i, step_data in enumerate(self.current_test_steps):
            item = QListWidgetItem(self.test_steps_list)
            widget = TestStepListItemWidget(step_data)
            item.setSizeHint(widget.sizeHint())
            self.test_steps_list.addItem(item)
            self.test_steps_list.setItemWidget(item, widget)
            # Atualiza a referência ao item da lista no dicionário do passo
            step_data['list_item'] = item 

        if select_index != -1 and select_index < self.test_steps_list.count():
            self.test_steps_list.setCurrentRow(select_index)
            self.editing_step_index = select_index
        else:
            self.test_steps_list.clearSelection()
            self.editing_step_index = -1
            
        self.remove_step_button.setEnabled(len(self.current_test_steps) > 0)
        self.clear_test_button.setEnabled(len(self.current_test_steps) > 0)
        self.save_test_button.setEnabled(len(self.current_test_steps) > 0)
        self._update_move_buttons_state()

    def _load_step_for_editing(self, item: QListWidgetItem):
        """
        Carrega os detalhes de um passo de teste selecionado na lista para os campos de edição.
        """
        self.editing_step_index = self.test_steps_list.row(item)
        
        if self.editing_step_index == -1 or self.editing_step_index >= len(self.current_test_steps):
            self._clear_step_input_fields()
            self.update_step_button.setEnabled(False)
            self.remove_step_button.setEnabled(False)
            self.editing_step_index = -1
            self._update_move_buttons_state()
            return

        step = self.current_test_steps[self.editing_step_index]

        self.step_name_input.setText(step.get("nome", ""))
        
        step_type = step.get("tipo_passo", "comando_validacao")
        if step_type == "comando_validacao":
            self.radio_comando_auto.setChecked(True)
            # Define a porta conforme o passo salvo
            port_type = step.get("port_type", "serial")
            if port_type == "modbus":
                self.step_port_type_combo.setCurrentText("Porta Modbus (Texto e Validação)")
            else:
                self.step_port_type_combo.setCurrentText("Porta Principal (Dados e Comandos)")
            self.step_command_input.setText(step.get("comando_enviar", ""))
            # Carrega campos RTC
            self.rtc_feature_cb.setChecked(step.get("use_rtc", False))
            self.rtc_mode_combo.setCurrentText(step.get("rtc_mode", "Escrita") or "Escrita")
            self.rtc_pattern_input.setText(step.get("rtc_pattern", ""))
            
            expect_response = step.get("esperar_resposta", False)
            self.step_expect_response_cb.setChecked(expect_response)
            
            if expect_response:
                self.step_timeout_input.setValue(step.get("timeout_ms", 1000))
                
                validation_type = step.get("tipo_validacao", "nenhuma")
                # Mapeia o nome interno para o nome de exibição no combobox
                validation_type_display_map = {
                    "nenhuma": "Nenhuma",
                    "string_exata": "Texto Exato",
                    "numerico_faixa": "Número em Faixa",
                    "texto_numerico_simples": "Texto Simples com Número",
                    "texto_numerico_multiplos": "Texto com Vários Números",
                    "serial_settings_match": "Número de Série (settings.json)",
                    "datetime_20s": "Data/Hora (±20s)",
                }
                display_text = validation_type_display_map.get(validation_type, "Nenhuma")
                self.step_validation_type_combo.setCurrentText(display_text)

                # Carrega os parâmetros de validação específicos
                if validation_type == "string_exata":
                    params = step.get("param_validacao", "")
                    self.exact_string_input.setText(params)
                elif validation_type == "numerico_faixa":
                    params = step.get("param_validacao", {})
                    self.numeric_min_input.setValue(params.get("min", 0.0))
                    self.numeric_max_input.setValue(params.get("max", 0.0))
                elif validation_type == "texto_numerico_simples":
                    params = step.get("param_validacao", {})
                    self.simplified_regex_input.setText(params.get("simplified_text", ""))
                    self.numeric_min_input.setValue(params.get("min", 0.0))
                    self.numeric_max_input.setValue(params.get("max", 0.0))
                elif validation_type == "texto_numerico_multiplos":
                    params = step.get("param_validacao", {})
                    self.simplified_regex_input.setText(params.get("simplified_text", ""))
                    self.numeric_min_input.setValue(params.get("min", 0.0))
                    self.numeric_max_input.setValue(params.get("max", 0.0))
                elif validation_type == "datetime_20s":
                    # Sem campos adicionais
                    pass
            else:
                self._clear_command_validation_fields()

        elif step_type == "instrucao_manual":
            self.radio_instrucao_manual.setChecked(True)
            self.step_instruction_text.setPlainText(step.get("mensagem_instrucao", ""))
            self.step_image_path_input.setText(step.get("caminho_imagem", ""))
        
        elif step_type == "tempo_espera":
            self.radio_tempo_espera.setChecked(True)
            self.wait_duration_input.setValue(step.get("duracao_espera_segundos", 10))
        
        elif step_type == "modbus_comando": # Carrega o novo tipo de passo Modbus
            self.radio_modbus_comando.setChecked(True)
            modbus_params_list = step.get("modbus_params", [])
            self.modbus_table_widget.setRowCount(len(modbus_params_list))
            for r, params in enumerate(modbus_params_list):
                self._init_modbus_table_row(r) # Garante que os widgets existam
                self.modbus_table_widget.cellWidget(r, 0).setValue(params.get("slave_id", 1))
                self.modbus_table_widget.cellWidget(r, 1).setCurrentText(params.get("function_code_display", "Read Holding Registers (0x03)"))
                self.modbus_table_widget.cellWidget(r, 2).setValue(params.get("address", 0))
                self.modbus_table_widget.cellWidget(r, 3).setValue(params.get("quantity", 1))
                self.modbus_table_widget.cellWidget(r, 4).setCurrentText(params.get("value_type", "Register (16-bit)"))
                self.modbus_table_widget.cellWidget(r, 5).setText(params.get("write_value", ""))
                self.modbus_table_widget.cellWidget(r, 6).setText(params.get("expected_value", ""))
                self.modbus_table_widget.cellWidget(r, 7).setValue(params.get("min_limit", 0.0))
                self.modbus_table_widget.cellWidget(r, 8).setValue(params.get("max_limit", 0.0))
                self._toggle_modbus_table_fields(r, self.modbus_table_widget.cellWidget(r, 1).currentIndex())
            self.modbus_table_widget.resizeColumnsToContents()

        elif step_type == "gravar_placa":
            # Carrega o novo tipo de passo 'Gravar Placa'
            if hasattr(self, 'radio_gravar_placa'):
                self.radio_gravar_placa.setChecked(True)
            if hasattr(self, 'flash_question_input'):
                self.flash_question_input.setText(step.get("pergunta_gravada", ""))
            if hasattr(self, 'flash_instruction_text'):
                self.flash_instruction_text.setPlainText(step.get("texto_como_gravar", ""))
            if hasattr(self, 'flash_cmd_path_input'):
                self.flash_cmd_path_input.setText(step.get("caminho_cmd", ""))
            if hasattr(self, 'flash_cmd_path_input2'):
                self.flash_cmd_path_input2.setText(step.get("caminho_cmd2", ""))

        elif step_type == "gravar_numero_serie":
            # Carrega o novo tipo de passo 'Gravar Número de Série'
            self.radio_gravar_ns.setChecked(True)
            # Apenas exibe o preset no campo de comando (para referência), sem usá-lo na execução
            self.step_command_input.setText("SET_NS=XXXXX/XXX;")


        self.update_step_button.setEnabled(True)
        self.remove_step_button.setEnabled(True)
        self._update_move_buttons_state()

    def _get_validation_params_from_ui(self, validation_type):
        """
        Coleta e valida os parâmetros de validação da UI com base no tipo de validação.
        Retorna um dicionário com os parâmetros ou levanta ValueError em caso de erro.
        """
        params = {}
        if validation_type == "string_exata":
            exact_string = self.exact_string_input.text()
            if not exact_string:
                raise ValueError("Para validação de Texto Exato, o campo não pode ser vazio.")
            params = exact_string
        
        elif validation_type == "numerico_faixa":
            min_val = self.numeric_min_input.value()
            max_val = self.numeric_max_input.value()
            if min_val > max_val:
                raise ValueError("O valor mínimo não pode ser maior que o valor máximo na validação numérica.")
            params = {"min": min_val, "max": max_val}
        
        elif validation_type == "texto_numerico_simples":
            min_val = self.numeric_min_input.value()
            max_val = self.numeric_max_input.value()
            if min_val > max_val:
                raise ValueError("O valor mínimo não pode ser maior que o valor máximo na validação de 'Texto Simples com Número'.")
            
            simplified_text = self.simplified_regex_input.text().strip()
            if not simplified_text:
                raise ValueError("Para 'Texto Simples com Número', o campo não pode ser vazio.")
            if "[VALOR]" not in simplified_text:
                raise ValueError("Para 'Texto Simples com Número', use '[VALOR]' para indicar a posição do número.")
            
            # Gera a expressão regular a partir do texto simples
            escaped_text_parts = [re.escape(part) for part in simplified_text.split("[VALOR]")]
            regex_pattern = r"".join([
                escaped_text_parts[0], 
                r"\s*([-+]?\d*\.?\d+)\s*", # Captura um número (inteiro ou flutuante, com sinal opcional)
                escaped_text_parts[1] if len(escaped_text_parts) > 1 else ""
            ])
            
            params = {
                "simplified_text": simplified_text,
                "regex": regex_pattern,
                "min": min_val,
                "max": max_val
            }
        elif validation_type == "texto_numerico_multiplos":
            min_val = self.numeric_min_input.value()
            max_val = self.numeric_max_input.value()
            if min_val > max_val:
                raise ValueError("O valor mínimo não pode ser maior que o valor máximo na validação de 'Texto com Vários Números'.")

            simplified_text = self.simplified_regex_input.text().strip()
            if not simplified_text:
                raise ValueError("Para 'Texto com Vários Números', o campo não pode ser vazio.")
            if "[VALOR]" not in simplified_text:
                raise ValueError("Para 'Texto com Vários Números', use '[VALOR]' para indicar as posições dos números.")

            parts = simplified_text.split("[VALOR]")
            expected_count = len(parts) - 1
            if expected_count < 2:
                raise ValueError("Para 'Texto com Vários Números', informe ao menos dois '[VALOR]'.")

            escaped_parts = [re.escape(p) for p in parts]
            regex_pattern = ""
            for i, p in enumerate(escaped_parts):
                regex_pattern += p
                if i < len(escaped_parts) - 1:
                    regex_pattern += r"\s*([-+]?\d*\.?\d+)\s*"

            params = {
                "simplified_text": simplified_text,
                "regex": regex_pattern,
                "min": min_val,
                "max": max_val,
                "expected_count": expected_count
            }
        elif validation_type == "serial_settings_match":
            # Não requer parâmetros adicionais; comparação usará settings.json em tempo de execução
            params = {}
        elif validation_type == "datetime_20s":
            # Não requer parâmetros adicionais
            params = {}
        # Removido o caso "modbus" daqui
        return params

    def _get_modbus_params_from_table(self):
        """
        Coleta e valida os parâmetros Modbus da tabela.
        Retorna uma lista de dicionários com os parâmetros ou levanta ValueError em caso de erro.
        """
        modbus_entries = []
        for r in range(self.modbus_table_widget.rowCount()):
            slave_id = self.modbus_table_widget.cellWidget(r, 0).value()
            function_combo = self.modbus_table_widget.cellWidget(r, 1)
            function_code_display = function_combo.currentText()
            function_code_hex = function_code_display.split('(')[1].strip(')').replace('0x', '')
            address = self.modbus_table_widget.cellWidget(r, 2).value()
            quantity = self.modbus_table_widget.cellWidget(r, 3).value()
            value_type = self.modbus_table_widget.cellWidget(r, 4).currentText()
            write_value = self.modbus_table_widget.cellWidget(r, 5).text().strip()
            expected_value = self.modbus_table_widget.cellWidget(r, 6).text().strip()
            min_limit = self.modbus_table_widget.cellWidget(r, 7).value()
            max_limit = self.modbus_table_widget.cellWidget(r, 8).value()

            if min_limit > max_limit:
                raise ValueError(f"Linha {r+1}: O limite mínimo não pode ser maior que o limite máximo.")

            entry = {
                "slave_id": slave_id,
                "function_code_display": function_code_display,
                "function_code": function_code_hex,
                "address": address,
                "quantity": quantity,
                "value_type": value_type,
                "write_value": write_value,
                "expected_value": expected_value,
                "min_limit": min_limit,
                "max_limit": max_limit
            }
            modbus_entries.append(entry)
        return modbus_entries

    def _add_new_test_step(self):
        """
        Adiciona um novo passo de teste à lista de passos.
        Valida os campos de entrada antes de adicionar.
        """
        step_name = self.step_name_input.text().strip()
        
        if not step_name:
            QMessageBox.warning(self, "Campo Vazio", "Por favor, insira um nome para o passo.")
            return

        step = {"nome": step_name, 'checked_for_fast_mode': False}

        if self.radio_comando_auto.isChecked():
            # Coleta dados para um passo de comando/validação automática
            port_type_display = self.step_port_type_combo.currentText()
            port_type_map = {
                "Porta Principal (Dados e Comandos)": "serial",
    "Porta Modbus (Texto e Validação)": "modbus",
            }
            port_type = port_type_map.get(port_type_display, "serial")
            
            command = self.step_command_input.text()
            expect_response = self.step_expect_response_cb.isChecked()
            timeout = self.step_timeout_input.value()
            
            validation_type_display = self.step_validation_type_combo.currentText()
            validation_type_map = {
                "Nenhuma": "nenhuma",
                "Texto Exato": "string_exata",
                "Número em Faixa": "numerico_faixa",
                "Texto Simples com Número": "texto_numerico_simples",
                "Texto com Vários Números": "texto_numerico_multiplos",
                "Número de Série": "serial_settings_match",
                "Data/Hora": "datetime_20s",
            }
            validation_type = validation_type_map.get(validation_type_display, "nenhuma")

            if not command:
                ns_ok = bool(getattr(self, 'ns_feature_cb', None) and self.ns_feature_cb.isChecked() and (getattr(self, 'ns_pattern_input', None) and self.ns_pattern_input.text().strip()))
                rtc_ok = bool(getattr(self, 'rtc_feature_cb', None) and self.rtc_feature_cb.isChecked() and (getattr(self, 'rtc_pattern_input', None) and self.rtc_pattern_input.text().strip()))
                if not (ns_ok or rtc_ok):
                    QMessageBox.warning(self, "Campo Vazio", "Para passo automático, informe um comando OU ative NS/RTC com um padrão válido.")
                    return

            # Salva o passo de comando/validação automática corretamente
            step["tipo_passo"] = "comando_validacao"
            step["port_type"] = port_type
            step["comando_enviar"] = command
            step["esperar_resposta"] = expect_response
            step["timeout_ms"] = timeout if expect_response else 0
            step["tipo_validacao"] = "nenhuma"
            # RTC fields
            step["use_rtc"] = bool(self.rtc_feature_cb.isChecked())
            step["rtc_mode"] = self.rtc_mode_combo.currentText() if self.rtc_feature_cb.isChecked() else ""
            step["rtc_pattern"] = self.rtc_pattern_input.text() if self.rtc_feature_cb.isChecked() else ""
            if expect_response:
                try:
                    step["tipo_validacao"] = validation_type
                    step["param_validacao"] = self._get_validation_params_from_ui(validation_type)
                except ValueError as e:
                    QMessageBox.warning(self, "Validação Inválida", str(e))
                    return

        elif getattr(self, 'radio_gravar_ns', None) and self.radio_gravar_ns.isChecked():
            # Novo: Passo para gravar número de série a partir do settings.json
            step["tipo_passo"] = "gravar_numero_serie"
            # Armazena um campo informativo opcional (apenas para exibição)
            step["comando_modelo"] = "SET_NS=XXXXX/XXX;"
            
        elif self.radio_instrucao_manual.isChecked():
            # Coleta dados para um passo de instrução manual
            instruction_text = self.step_instruction_text.toPlainText().strip()
            image_path = self.step_image_path_input.text().strip()
            
            if not instruction_text:
                QMessageBox.warning(self, "Campo Vazio", "Para passo de Instrução Manual, a mensagem de instrução não pode ser vazia.")
                return
            
            if image_path and not os.path.exists(image_path):
                QMessageBox.warning(self, "Caminho de Imagem Inválido", f"O caminho da imagem '{image_path}' não existe. Por favor, verifique.")
                return

            step["tipo_passo"] = "instrucao_manual"
            step["mensagem_instrucao"] = instruction_text
            step["caminho_imagem"] = image_path
        
        elif self.radio_tempo_espera.isChecked():
            duration = self.wait_duration_input.value()
            if duration <= 0:
                QMessageBox.warning(self, "Duração Inválida", "A duração da espera deve ser um número positivo.")
                return
            step["tipo_passo"] = "tempo_espera"
            step["duracao_espera_segundos"] = duration

        elif getattr(self, 'radio_gravar_placa', None) and self.radio_gravar_placa.isChecked():
            # Coleta dados para o passo 'Gravar Placa'
            question = self.flash_question_input.text().strip() if hasattr(self, 'flash_question_input') else ""
            instructions = self.flash_instruction_text.toPlainText().strip() if hasattr(self, 'flash_instruction_text') else ""
            cmd_path = self.flash_cmd_path_input.text().strip() if hasattr(self, 'flash_cmd_path_input') else ""
            cmd_path2 = self.flash_cmd_path_input2.text().strip() if hasattr(self, 'flash_cmd_path_input2') else ""

            if not question:
                QMessageBox.warning(self, "Campo Vazio", "Para 'Gravar Placa', o campo Pergunta não pode ser vazio.")
                return
            if not instructions:
                QMessageBox.warning(self, "Campo Vazio", "Para 'Gravar Placa', o campo Instruções não pode ser vazio.")
                return
            if not cmd_path:
                QMessageBox.warning(self, "Campo Vazio", "Para 'Gravar Placa', informe o caminho do arquivo .cmd.")
                return

            step["tipo_passo"] = "gravar_placa"
            step["pergunta_gravada"] = question
            step["texto_como_gravar"] = instructions
            step["caminho_cmd"] = cmd_path
            if cmd_path2:
                step["caminho_cmd2"] = cmd_path2

        
        elif self.radio_modbus_comando.isChecked(): # NOVO: Adiciona o passo de comando Modbus
            try:
                modbus_params = self._get_modbus_params_from_table()
                if not modbus_params:
                    QMessageBox.warning(self, "Configuração Modbus Vazia", "Adicione pelo menos uma linha à tabela Modbus.")
                    return
                step["tipo_passo"] = "modbus_comando"
                step["modbus_params"] = modbus_params
            except ValueError as e:
                QMessageBox.warning(self, "Erro de Configuração Modbus", str(e))
                return


        self.current_test_steps.append(step)
        self._update_test_steps_list() # Atualiza a lista exibida
        
        self._clear_step_input_fields() # Limpa os campos de entrada
        
        self.test_steps_list.clearSelection()
        self.update_step_button.setEnabled(False)
        self.remove_step_button.setEnabled(False)
        self.editing_step_index = -1
        self._update_move_buttons_state()

    def _update_selected_test_step(self):
        """
        Atualiza um passo de teste existente na lista com base nos campos de edição.
        Valida os campos de entrada antes de atualizar.
        """
        if self.editing_step_index == -1 or self.editing_step_index >= len(self.current_test_steps):
            QMessageBox.warning(self, "Nenhum Passo Selecionado", "Por favor, selecione um passo na lista para editar.")
            return

        step_name = self.step_name_input.text().strip()
        if not step_name:
            QMessageBox.warning(self, "Campo Vazio", "Por favor, insira um nome para o passo.")
            return

        # Recupera o estado 'checked_for_fast_mode' atual do widget para não perdê-lo
        current_list_item = self.test_steps_list.item(self.editing_step_index)
        current_widget = self.test_steps_list.itemWidget(current_list_item)
        if isinstance(current_widget, TestStepListItemWidget):
            checked_state = current_widget.checkbox.isChecked()
        else:
            checked_state = False # Fallback se não for um TestStepListItemWidget

        updated_step = {"nome": step_name, 'checked_for_fast_mode': checked_state}

        if self.radio_comando_auto.isChecked():
            # Coleta dados para um passo de comando/validação automática
            port_type_display = self.step_port_type_combo.currentText()
            port_type_map = {
                "Porta Principal (Dados e Comandos)": "serial",
    "Porta Modbus (Texto e Validação)": "modbus",
            }
            port_type = port_type_map.get(port_type_display, "serial")

            command = self.step_command_input.text()
            expect_response = self.step_expect_response_cb.isChecked()
            timeout = self.step_timeout_input.value()
            
            validation_type_display = self.step_validation_type_combo.currentText()
            validation_type_map = {
                "Nenhuma": "nenhuma",
                "Texto Exato": "string_exata",
                "Número em Faixa": "numerico_faixa",
                "Texto Simples com Número": "texto_numerico_simples",
                "Texto com Vários Números": "texto_numerico_multiplos",
                "Data/Hora": "datetime_20s",
            }
            validation_type = validation_type_map.get(validation_type_display, "nenhuma")

            if not command:
                ns_ok = bool(getattr(self, 'ns_feature_cb', None) and self.ns_feature_cb.isChecked() and (getattr(self, 'ns_pattern_input', None) and self.ns_pattern_input.text().strip()))
                rtc_ok = bool(getattr(self, 'rtc_feature_cb', None) and self.rtc_feature_cb.isChecked() and (getattr(self, 'rtc_pattern_input', None) and self.rtc_pattern_input.text().strip()))
                if not (ns_ok or rtc_ok):
                    QMessageBox.warning(self, "Campo Vazio", "Para passo automático, informe um comando OU ative NS/RTC com um padrão válido.")
                    return

            updated_step["tipo_passo"] = "comando_validacao"
            updated_step["port_type"] = port_type
            updated_step["comando_enviar"] = command
            updated_step["esperar_resposta"] = expect_response
            updated_step["timeout_ms"] = timeout if expect_response else 0
            updated_step["tipo_validacao"] = "nenhuma"
            # RTC fields
            updated_step["use_rtc"] = bool(self.rtc_feature_cb.isChecked())
            updated_step["rtc_mode"] = self.rtc_mode_combo.currentText() if self.rtc_feature_cb.isChecked() else ""
            updated_step["rtc_pattern"] = self.rtc_pattern_input.text() if self.rtc_feature_cb.isChecked() else ""

            if expect_response:
                try:
                    updated_step["tipo_validacao"] = validation_type
                    updated_step["param_validacao"] = self._get_validation_params_from_ui(validation_type)
                except ValueError as e:
                    QMessageBox.warning(self, "Validação Inválida", str(e))
                    return
        elif self.radio_instrucao_manual.isChecked():
            # Coleta dados para um passo de instrução manual
            instruction_text = self.step_instruction_text.toPlainText().strip()
            image_path = self.step_image_path_input.text().strip()
            
            if not instruction_text:
                QMessageBox.warning(self, "Campo Vazio", "Para passo de Instrução Manual, a mensagem de instrução não pode ser vazia.")
                return
            
            if image_path and not os.path.exists(image_path):
                QMessageBox.warning(self, "Caminho de Imagem Inválido", f"O caminho da imagem '{image_path}' não existe. Por favor, verifique.")
                return

            updated_step["tipo_passo"] = "instrucao_manual"
            updated_step["mensagem_instrucao"] = instruction_text
            updated_step["caminho_imagem"] = image_path
        
        elif self.radio_tempo_espera.isChecked():
            duration = self.wait_duration_input.value()
            if duration <= 0:
                QMessageBox.warning(self, "Duração Inválida", "A duração da espera deve ser um número positivo.")
                return
            updated_step["tipo_passo"] = "tempo_espera"
            updated_step["duracao_espera_segundos"] = duration
        
        elif getattr(self, 'radio_gravar_placa', None) and self.radio_gravar_placa.isChecked():
            # Atualiza dados para o passo 'Gravar Placa'
            question = self.flash_question_input.text().strip() if hasattr(self, 'flash_question_input') else ""
            instructions = self.flash_instruction_text.toPlainText().strip() if hasattr(self, 'flash_instruction_text') else ""
            cmd_path = self.flash_cmd_path_input.text().strip() if hasattr(self, 'flash_cmd_path_input') else ""
            cmd_path2 = self.flash_cmd_path_input2.text().strip() if hasattr(self, 'flash_cmd_path_input2') else ""

            if not question:
                QMessageBox.warning(self, "Campo Vazio", "Para 'Gravar Placa', o campo Pergunta não pode ser vazio.")
                return
            if not instructions:
                QMessageBox.warning(self, "Campo Vazio", "Para 'Gravar Placa', o campo Instruções não pode ser vazio.")
                return
            if not cmd_path:
                QMessageBox.warning(self, "Campo Vazio", "Para 'Gravar Placa', informe o caminho do arquivo .cmd.")
                return

            updated_step["tipo_passo"] = "gravar_placa"
            updated_step["pergunta_gravada"] = question
            updated_step["texto_como_gravar"] = instructions
            updated_step["caminho_cmd"] = cmd_path
            if cmd_path2:
                updated_step["caminho_cmd2"] = cmd_path2

        
        elif self.radio_modbus_comando.isChecked(): # NOVO: Atualiza o passo de comando Modbus
            try:
                modbus_params = self._get_modbus_params_from_table()
                if not modbus_params:
                    QMessageBox.warning(self, "Configuração Modbus Vazia", "Adicione pelo menos uma linha à tabela Modbus.")
                    return
                updated_step["tipo_passo"] = "modbus_comando"
                updated_step["modbus_params"] = modbus_params
            except ValueError as e:
                QMessageBox.warning(self, "Erro de Configuração Modbus", str(e))
                return
        elif getattr(self, 'radio_gravar_ns', None) and self.radio_gravar_ns.isChecked():
            # Atualiza para o passo 'gravar número de série'
            updated_step["tipo_passo"] = "gravar_numero_serie"
            updated_step["comando_modelo"] = "SET_NS=XXXXX/XXX;"
            
        self.current_test_steps[self.editing_step_index] = updated_step # Atualiza o passo na lista
        self._update_test_steps_list(select_index=self.editing_step_index) # Atualiza a exibição da lista
        QMessageBox.information(self, "Passo Atualizado", f"Passo '{step_name}' atualizado com sucesso!")
        
        self._clear_step_input_fields() # Limpa os campos de entrada
        self.test_steps_list.clearSelection()
        self.update_step_button.setEnabled(False)
        self.remove_step_button.setEnabled(False)
        self.editing_step_index = -1
        self._update_move_buttons_state()

    def _remove_selected_test_step(self):
        """
        Remove o passo de teste selecionado da lista.
        """
        if self.editing_step_index == -1 or self.editing_step_index >= len(self.current_test_steps):
            QMessageBox.warning(self, "Nenhum Passo Selecionado", "Por favor, selecione um passo na lista para remover.")
            return

        reply = QMessageBox.question(self, "Confirmar Remoção", 
                                    f"Tem certeza que deseja remover o passo '{self.current_test_steps[self.editing_step_index]['nome']}'?",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            removed_step_name = self.current_test_steps[self.editing_step_index]['nome']
            del self.current_test_steps[self.editing_step_index] # Remove o passo da lista
            self._update_test_steps_list() # Atualiza a exibição da lista
            QMessageBox.information(self, "Passo Removido", f"Passo '{removed_step_name}' removido com sucesso!")
            
            self._clear_step_input_fields()
            self.test_steps_list.clearSelection()
            self.update_step_button.setEnabled(False)
            self.remove_step_button.setEnabled(False)
            self.editing_step_index = -1
            self._update_move_buttons_state()

    def _clear_test_steps(self):
        """
        Limpa todos os passos de teste da lista atual.
        """
        reply = QMessageBox.question(self, "Limpar Passos", 
                                    "Tem certeza que deseja limpar todos os passos de teste atuais? Esta ação não pode ser desfeita.",
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.current_test_steps = []
            self.test_serial_command_settings = {}
            self.test_modbus_settings = {}
            self.fast_mode_secret_code = "" # Limpa o código secreto
            self.fast_mode_code_input.clear() # Limpa o campo na UI
            self._update_test_steps_list()
            self._clear_step_input_fields()
            self.editing_step_index = -1
            self.update_step_button.setEnabled(False)
            self.remove_step_button.setEnabled(False)
            self._update_move_buttons_state()
            self._update_test_port_settings_display()
            QMessageBox.information(self, "Passos Limpos", "Todos os passos de teste foram removidos.")

    def _clear_command_validation_fields(self):
        """Limpa todos os campos relacionados a comandos e validação."""
        self.step_port_type_combo.setCurrentText("Porta Principal (Dados e Comandos)")
        self.step_command_input.clear()
        self.step_expect_response_cb.setChecked(False)
        self.step_timeout_input.setValue(1000)
        self.step_validation_type_combo.setCurrentIndex(0) # Isso irá chamar _toggle_validation_params_fields(0)
        self.exact_string_input.clear()
        self.simplified_regex_input.clear()
        self.numeric_min_input.setValue(0.0)
        self.numeric_max_input.setValue(0.0)
        # Reset RTC controls
        if hasattr(self, "rtc_feature_cb"):
            self.rtc_feature_cb.setChecked(False)
        if hasattr(self, "rtc_mode_combo"):
            self.rtc_mode_combo.setCurrentIndex(0)
        if hasattr(self, "rtc_pattern_input"):
            self.rtc_pattern_input.clear()
        if hasattr(self, "_toggle_rtc_fields"):
            self._toggle_rtc_fields()

    def _clear_modbus_command_fields(self):
        """Limpa e reseta a tabela Modbus para o estado padrão."""
        self.modbus_table_widget.clearContents()
        self.modbus_table_widget.setRowCount(1)
        self._init_modbus_table_row(0)

    def _clear_step_input_fields(self):
        """Limpa todos os campos de entrada do criador de teste."""
        self.step_name_input.clear()
        # Definir o rádio botão de comando automático como True irá chamar _toggle_step_type_fields
        # que por sua vez limpará os campos de comando/validação e ocultará os de instrução manual.
        self.radio_comando_auto.setChecked(True) 
        self.step_instruction_text.clear()
        self.step_image_path_input.clear()
        self.wait_duration_input.setValue(10) # Reseta o valor do tempo de espera
        self._clear_modbus_command_fields() # Limpa os campos Modbus

    def _move_step_up(self):
        """
        Move o passo de teste selecionado para cima na lista.
        """
        current_row = self.test_steps_list.currentRow()
        if current_row > 0:
            item_data = self.current_test_steps.pop(current_row) # Remove o item da posição atual
            self.current_test_steps.insert(current_row - 1, item_data) # Insere na posição anterior
            self._update_test_steps_list(select_index=current_row - 1) # Atualiza a lista e seleciona o item movido

    def _move_step_down(self):
        """
        Move o passo de teste selecionado para baixo na lista.
        """
        current_row = self.test_steps_list.currentRow()
        if current_row < len(self.current_test_steps) - 1:
            item_data = self.current_test_steps.pop(current_row) # Remove o item da posição atual
            self.current_test_steps.insert(current_row + 1, item_data) # Insere na próxima posição
            self._update_test_steps_list(select_index=current_row + 1) # Atualiza a lista e seleciona o item movido

    def _update_move_buttons_state(self):
        """
        Atualiza o estado (habilitado/desabilitado) dos botões de mover passo.
        """
        current_row = self.test_steps_list.currentRow()
        num_steps = len(self.current_test_steps)

        is_item_selected = (current_row != -1)
        self.update_step_button.setEnabled(is_item_selected)
        self.remove_step_button.setEnabled(is_item_selected)
        
        if num_steps <= 1 or current_row == -1:
            self.move_step_up_button.setEnabled(False)
            self.move_step_down_button.setEnabled(False)
        else:
            self.move_step_up_button.setEnabled(current_row > 0)
            self.move_step_down_button.setEnabled(current_row < num_steps - 1)

    def _open_test_port_config_dialog(self):
        """
        Abre o diálogo para configurar as portas seriais específicas para o teste.
        """
        dialog = TestPortConfigDialog(self.test_serial_command_settings, self.test_modbus_settings, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            serial_settings, modbus_settings = dialog.get_settings()
            self.test_serial_command_settings = serial_settings
            self.test_modbus_settings = modbus_settings
            self._update_test_port_settings_display() # Atualiza os rótulos de exibição das configurações
            QMessageBox.information(self, "Configurações Salvas", "Configurações de porta para este teste salvas. Lembre-se de salvar o arquivo de teste.")
        else:
            self.log_message("Configuração de portas do teste cancelada.", "informacao")

    def _update_test_port_settings_display(self):
        """
        Atualiza os rótulos que exibem as configurações das portas para o teste.
        """
        if self.test_serial_command_settings:
            s = self.test_serial_command_settings
            self.test_serial_settings_label.setText(
                f"Porta Principal: Baud: {s.get('baud')}, Data: {s.get('data_bits')}, Parity: {s.get('parity')}, Flow: {s.get('handshake')}, Modo: {s.get('mode')}"
            )
        else:
            self.test_serial_settings_label.setText("Porta Principal: Não configurada")

        if self.test_modbus_settings:
            m = self.test_modbus_settings
            self.test_modbus_settings_label.setText(
                f"Porta Modbus: Baud: {m.get('baud')}, Data: {m.get('data_bits')}, Parity: {m.get('parity')}, Flow: {m.get('handshake')}, Modo: {m.get('mode')}"
            )
        else:
            self.test_modbus_settings_label.setText("Porta Modbus: Não configurada")

    def _save_test_file(self):
        """
        Salva os passos de teste atuais e suas configurações de porta em um arquivo JSON.
        Inclui o código secreto do modo fast e o status 'checked_for_fast_mode' de cada passo.
        """
        if not self.current_test_steps:
            QMessageBox.warning(self, "Nenhum Passo", "Não há passos para salvar. Adicione alguns primeiro.")
            return
        
        # Garante que os dados dos checkboxes estão atualizados nos self.current_test_steps
        for i in range(self.test_steps_list.count()):
            item = self.test_steps_list.item(i)
            widget = self.test_steps_list.itemWidget(item)
            if isinstance(widget, TestStepListItemWidget):
                # Atualiza o dicionário original do passo com o estado do checkbox do widget
                self.current_test_steps[i]['checked_for_fast_mode'] = widget.checkbox.isChecked()

        # Salva o código secreto do modo fast
        self.fast_mode_secret_code = self.fast_mode_code_input.text().strip()
        if not re.fullmatch(r"^\d{4}$", self.fast_mode_secret_code):
            QMessageBox.warning(self, "Código Secreto Inválido",
                                "O código secreto do Modo Fast deve conter exatamente 4 dígitos numéricos. Ele não será salvo se for inválido.")
            self.fast_mode_secret_code = "" # Limpa se for inválido

        # Verifica se algum passo Modbus está presente para definir 'modbus_required'
        modbus_needed_for_save = any(step.get("tipo_passo") == "modbus_comando" for step in self.current_test_steps)
        
        serializable_steps = []
        for step in self.current_test_steps:
            temp_step = step.copy()
            if 'list_item' in temp_step: # Remove a referência ao item da lista antes de serializar
                del temp_step['list_item']
            serializable_steps.append(temp_step)

        test_data_to_save = {
            "version": "1.4", # Versão do formato do arquivo de teste (atualizada para Modbus como passo exclusivo)
            "modbus_required": modbus_needed_for_save,
            "fast_mode_code": self.fast_mode_secret_code, # Salva o código secreto
            "steps": serializable_steps,
            "port_configurations": {
                "serial_command": self.test_serial_command_settings,
                "modbus": self.test_modbus_settings
            }
        }

        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo de Teste", "", "Arquivos JSON (*.json);;Todos os Arquivos (*)")
        if file_path:
            if not file_path.endswith(".json"):
                file_path += ".json"
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(test_data_to_save, f, indent=4, ensure_ascii=False)
                QMessageBox.information(self, "Salvar Teste", f"Teste salvo com sucesso em:\n{os.path.basename(file_path)}")
            except Exception as e:
                self.log_message(f"Erro ao salvar arquivo de teste: {e}", "erro")
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo:\n{e}")

    def _list_serial_ports(self):
        """
        Lista as portas seriais disponíveis no sistema e as popula nos comboboxes.
        Atualiza o estado dos botões de conexão.
        """
        # Salva as seleções atuais antes de limpar
        current_serial_command_port = self.serial_command_port_combobox.currentText()
        current_modbus_port = self.modbus_port_combobox.currentText()

        self.serial_command_port_combobox.clear()
        self.modbus_port_combobox.clear()

        try:
            from serial.tools.list_ports import comports
            ports = comports()
        except ImportError:
            self.log_message("Módulo 'pyserial' não encontrado. Por favor, instale-o com 'pip install pyserial'", "erro")
            ports = []

        if not ports:
            self.serial_command_port_combobox.addItem("Nenhuma porta COM encontrada")
            self.modbus_port_combobox.addItem("Nenhuma porta COM encontrada")
            self.connect_serial_command_button.setEnabled(False)
            self.connect_modbus_button.setEnabled(False)
            self.serial_command_port_combobox.setEnabled(False)
            self.modbus_port_combobox.setEnabled(False)
            return

        for p in ports:
            self.serial_command_port_combobox.addItem(p.device)
            self.modbus_port_combobox.addItem(p.device)

        # Tenta restaurar as seleções salvas
        if current_serial_command_port and self.serial_command_port_combobox.findText(current_serial_command_port) != -1:
            self.serial_command_port_combobox.setCurrentText(current_serial_command_port)
        else:
            # Se a porta salva não está mais disponível, limpa a seleção
            self.serial_command_port_combobox.setCurrentText("")

        if current_modbus_port and self.modbus_port_combobox.findText(current_modbus_port) != -1:
            self.modbus_port_combobox.setCurrentText(current_modbus_port)
        else:
            # Se a porta salva não está mais disponível, limpa a seleção
            self.modbus_port_combobox.setCurrentText("")


        self.connect_serial_command_button.setEnabled(True)
        self.serial_command_port_combobox.setEnabled(True)
        # Habilita o botão Modbus apenas se o grupo Modbus estiver visível (ou seja, se o teste carregado o requer)
        self.connect_modbus_button.setEnabled(self.modbus_serial_group.isVisible())
        self.modbus_port_combobox.setEnabled(self.modbus_serial_group.isVisible())

    # --- Auto-reconexão Modbus ---
    def _handle_connection_lost(self, port_name: str):
        try:
            self.log_message(f"Conexão perdida na porta {port_name}.", "erro")
            if port_name == "Modbus":
                # Limpa instâncias antigas
                try:
                    if self.modbus_serial_reader_thread:
                        self.modbus_serial_reader_thread.stop()
                except Exception:
                    pass
                self.modbus_serial_reader_thread = None
                try:
                    if self.modbus_ser and self.modbus_ser.is_open:
                        self.modbus_ser.close()
                except Exception:
                    pass
                self.modbus_ser = None

                # Armazena porta alvo se vazio
                if not self.modbus_target_port:
                    self.modbus_target_port = self.modbus_port_combobox.currentText()

                # Inicia tentativas
                self.modbus_reconnect_attempts_remaining = 10
                if self.modbus_reconnect_timer is None:
                    self.modbus_reconnect_timer = QTimer(self)
                    self.modbus_reconnect_timer.setSingleShot(True)
                    self.modbus_reconnect_timer.timeout.connect(self._try_reconnect_modbus)
                self.log_message(f"Tentando reconectar Modbus em '{self.modbus_target_port}' (10 tentativas, a cada 2s)...", "sistema")
                self.modbus_reconnect_timer.start(2000)
        except Exception as e:
            self.log_message(f"Erro no handler de conexão perdida: {e}", "erro")

    def _try_reconnect_modbus(self):
        # Se não há mais tentativas, aborta
        if self.modbus_reconnect_attempts_remaining <= 0:
            self.log_message("Falha ao reconectar Modbus após 10 tentativas.", "erro")
            return
        target = self.modbus_target_port or self.modbus_port_combobox.currentText()
        if not target:
            self.modbus_reconnect_attempts_remaining = 0
            self.log_message("Sem porta Modbus alvo para reconectar.", "erro")
            return

        # Evita tentar se o usuário já reconectou/fechou manualmente
        if self.modbus_ser is not None and getattr(self.modbus_ser, 'is_open', False):
            self.modbus_reconnect_attempts_remaining = 0
            return

        # Checa se a porta alvo está presente no sistema
        try:
            from serial.tools.list_ports import comports
            available = [p.device for p in comports()]
        except Exception:
            available = []
        if target not in available:
            # Aguarda porta reaparecer
            self.log_message(f"Porta '{target}' ainda não disponível. Tentativas restantes: {self.modbus_reconnect_attempts_remaining}", "informacao")
            self.modbus_reconnect_attempts_remaining -= 1
            if self.modbus_reconnect_attempts_remaining > 0 and self.modbus_reconnect_timer:
                self.modbus_reconnect_timer.start(2000)
            return

        # Define configurações (do teste se existir, senão da UI)
        if self.current_test_steps:
            port_settings = self.test_modbus_settings
        else:
            port_settings = {
                "baud": self.modbus_baud_combo.currentText(),
                "data_bits": self.modbus_data_bits_combo.currentText(),
                "parity": self.modbus_parity_combo.currentText(),
                "handshake": self.modbus_handshake_combo.currentText(),
                "mode": self.modbus_mode_combo.currentText()
            }

        parity_map = {
            "Nenhuma": serial.PARITY_NONE, "Ímpar": serial.PARITY_ODD, "Par": serial.PARITY_EVEN,
            "Marca": serial.PARITY_MARK, "Espaço": serial.PARITY_SPACE
        }
        baud = int(port_settings.get("baud", "9600"))
        data_bits = int(port_settings.get("data_bits", "8"))
        parity = parity_map.get(port_settings.get("parity", "Nenhuma"), serial.PARITY_NONE)
        flow = port_settings.get("handshake", "Nenhum")
        xonxoff = (flow == "XON/XOFF")
        rtscts = (flow == "RTS/CTS")

        # Estados DTR/RTS desejados (aplicar após abrir)
        dtr_state = self.modbus_dtr_checkbox.isChecked()
        rts_state = self.modbus_rts_checkbox.isChecked()

        try:
            # Estratégia segura: abrir com handshake desligado e DTR/RTS default
            ser_instance = serial.Serial(
                port=target, baudrate=baud, bytesize=data_bits, parity=parity,
                stopbits=serial.STOPBITS_ONE, xonxoff=False, rtscts=False, timeout=0.5
            )
            self.modbus_ser = ser_instance
            # aplica DTR/RTS após abrir
            try:
                self.modbus_ser.dtr = dtr_state
                self.modbus_ser.rts = rts_state
            except Exception:
                pass
            # Se o handshake configurado exige rtscts/xonxoff, aplica agora
            try:
                if xonxoff or rtscts:
                    self.modbus_ser.close()
                    self.modbus_ser = serial.Serial(
                        port=target, baudrate=baud, bytesize=data_bits, parity=parity,
                        stopbits=serial.STOPBITS_ONE, xonxoff=xonxoff, rtscts=rtscts, timeout=0.5
                    )
                    try:
                        self.modbus_ser.dtr = dtr_state
                        self.modbus_ser.rts = rts_state
                    except Exception:
                        pass
            except Exception:
                # Se falhar aplicar handshake, mantém conexão básica
                pass
            self.modbus_ser.flushInput()
            self.modbus_ser.flushOutput()

            # UI
            self.modbus_port_combobox.setCurrentText(target)
            self.modbus_port_combobox.setEnabled(False)
            self.connect_modbus_button.setText("Fechar Porta Modbus")

            # Thread leitora
            new_reader_thread = SerialReaderThread(self.modbus_ser, "Modbus", is_modbus_port=True)
            new_reader_thread.data_received.connect(self._display_received_data)
            new_reader_thread.connection_lost.connect(self._handle_connection_lost)
            self.modbus_serial_reader_thread = new_reader_thread
            new_reader_thread.start()

            self.log_message(f"Reconectado Modbus em '{target}'.", "sistema")
            # sucesso: para tentativas
            self.modbus_reconnect_attempts_remaining = 0
            if self.modbus_reconnect_timer:
                self.modbus_reconnect_timer.stop()
            return
        except Exception as e:
            self.modbus_reconnect_attempts_remaining -= 1
            self.log_message(f"Falha ao reconectar Modbus em '{target}': {e}. Tentativas restantes: {self.modbus_reconnect_attempts_remaining}", "erro")
            if self.modbus_reconnect_attempts_remaining > 0 and self.modbus_reconnect_timer:
                self.modbus_reconnect_timer.start(2000)

    def _toggle_serial_connection(self, port_type):
        """
        Alterna a conexão para a porta serial principal ou Modbus.
        Abre ou fecha a porta e inicia/para a thread de leitura correspondente.
        """
        current_ser = None
        port_combobox = None
        connect_button = None
        reader_thread_attr = ''
        ser_attr = ''
        button_text_open = ""
        button_text_close = ""
        log_prefix = ""
        is_modbus = False # Adiciona a flag para o SerialReaderThread
        
        port_settings = {}

        # Define as variáveis com base no tipo de porta
        if port_type == "serial_command":
            current_ser = self.serial_command_ser
            port_combobox = self.serial_command_port_combobox
            connect_button = self.connect_serial_command_button
            reader_thread_attr = 'serial_command_reader_thread'
            ser_attr = 'serial_command_ser'
            button_text_open = "Fechar Porta Principal"
            button_text_close = "Abrir Porta Principal"
            log_prefix = "Principal"
            is_modbus = False
            
            # Usa as configurações do teste se um teste estiver carregado, senão as da UI principal
            if self.current_test_steps:
                port_settings = self.test_serial_command_settings
            else:
                port_settings = {
                    "baud": self.serial_baud_combo.currentText(),
                    "data_bits": self.serial_data_bits_combo.currentText(),
                    "parity": self.serial_parity_combo.currentText(),
                    "handshake": self.serial_handshake_combo.currentText(),
                    "mode": self.serial_mode_combo.currentText()
                }
            # Obtém os estados de DTR e RTS dos checkboxes da UI
            dtr_state = self.serial_dtr_checkbox.isChecked()
            rts_state = self.serial_rts_checkbox.isChecked()

        elif port_type == "modbus":
            current_ser = self.modbus_ser
            port_combobox = self.modbus_port_combobox
            connect_button = self.connect_modbus_button
            reader_thread_attr = 'modbus_serial_reader_thread'
            ser_attr = 'modbus_ser'
            button_text_open = "Fechar Porta Modbus"
            button_text_close = "Abrir Porta Modbus"
            log_prefix = "Modbus"
            is_modbus = True
            
            if self.current_test_steps:
                port_settings = self.test_modbus_settings
            else:
                port_settings = {
                    "baud": self.modbus_baud_combo.currentText(),
                    "data_bits": self.modbus_data_bits_combo.currentText(),
                    "parity": self.modbus_parity_combo.currentText(),
                    "handshake": self.modbus_handshake_combo.currentText(),
                    "mode": self.modbus_mode_combo.currentText()
                }
            # Obtém os estados de DTR e RTS dos checkboxes da UI
            dtr_state = self.modbus_dtr_checkbox.isChecked()
            rts_state = self.modbus_rts_checkbox.isChecked()
        else:
            return

        if current_ser is None or not current_ser.is_open:
            # Tenta abrir a porta
            port = port_combobox.currentText()
            if port == "Nenhuma porta COM encontrada" or not port:
                QMessageBox.warning(self, "Erro de Conexão", f"Nenhuma porta serial selecionada ou disponível para a porta {log_prefix}.")
                return
            
            baud = int(port_settings.get("baud", "9600"))
            data_bits = int(port_settings.get("data_bits", "8"))
            
            parity_map = {
                "Nenhuma": serial.PARITY_NONE, "Ímpar": serial.PARITY_ODD, "Par": serial.PARITY_EVEN,
                "Marca": serial.PARITY_MARK, "Espaço": serial.PARITY_SPACE
            }
            parity = parity_map.get(port_settings.get("parity", "Nenhuma"), serial.PARITY_NONE)

            flow_control_type = port_settings.get("handshake", "Nenhum")
            xonxoff = False
            rtscts = False
            if flow_control_type == "RTS/CTS": rtscts = True
            elif flow_control_type == "XON/XOFF": xonxoff = True

            try:
                # Tenta criar a instância da porta serial
                # REMOVIDO: Força DTR e RTS para False na abertura para evitar resets
                ser_instance = serial.Serial(
                    port=port, baudrate=baud, bytesize=data_bits, parity=parity,
                    stopbits=serial.STOPBITS_ONE, xonxoff=xonxoff, rtscts=rtscts, 
                    timeout=0.5
                )
                setattr(self, ser_attr, ser_instance)
                current_ser = getattr(self, ser_attr)
                # Se Modbus, guarda porta alvo para auto-reconectar
                if port_type == "modbus":
                    self.modbus_target_port = port
                
                # ADICIONADO: Após a abertura, aplica o estado desejado pelos checkboxes
                try:
                    current_ser.dtr = dtr_state
                    current_ser.rts = rts_state
                except Exception as dtr_rts_e:
                    self.log_message(f"AVISO: Não foi possível definir DTR/RTS para a porta {log_prefix}: {dtr_rts_e}. Verifique a compatibilidade do hardware/driver.", "erro")


                current_ser.flushInput() # Limpa buffers de entrada e saída
                current_ser.flushOutput()
                self.log_message(f"Conectado à porta {port}.", "sistema")
                connect_button.setText(button_text_open)
                
                port_combobox.setEnabled(False) # Desabilita o combobox da porta após a conexão
                
                if port_type == "serial_command":
                    # Conecta os sinais dos checkboxes DTR/RTS para controle dinâmico
                    self.serial_dtr_checkbox.stateChanged.connect(lambda state: self._toggle_dtr(state, "serial_command"))
                    self.serial_rts_checkbox.stateChanged.connect(lambda state: self._toggle_rts(state, "serial_command"))

                    # Habilita campos de envio direto e automático para a porta principal
                    self.direct_command_input.setEnabled(True)
                    self.direct_send_button.setEnabled(True)
                    self.send_target_port_combo.setEnabled(True)
                    self.modbus_display_text_cb.setEnabled(True)
                    self.direct_command_input.setFocus()

                    for i in range(len(self.send_command_inputs)):
                        self.send_command_inputs[i].setEnabled(True)
                        self.send_buttons[i].setEnabled(True)
                        self.auto_send_checkboxes[i].setEnabled(True)
                        self.auto_send_config_buttons[i].setEnabled(True)
                        if self.auto_send_checkboxes[i].isChecked():
                            self._toggle_auto_send_timer(Qt.CheckState.Checked.value, i)
                           
                elif port_type == "modbus":
                    # Conecta os sinais dos checkboxes DTR/RTS para controle dinâmico
                    self.modbus_dtr_checkbox.stateChanged.connect(lambda state: self._toggle_dtr(state, "modbus"))
                    self.modbus_rts_checkbox.stateChanged.connect(lambda state: self._toggle_rts(state, "modbus"))

                    # Habilita campos de envio manual/auto quando somente Modbus estiver aberto
                    self.direct_command_input.setEnabled(True)
                    self.direct_send_button.setEnabled(True)
                    self.send_target_port_combo.setEnabled(True)
                    self.modbus_display_text_cb.setEnabled(True)
                    for i in range(len(self.send_command_inputs)):
                        self.send_command_inputs[i].setEnabled(True)
                        self.send_buttons[i].setEnabled(True)
                        self.auto_send_checkboxes[i].setEnabled(True)
                        self.auto_send_config_buttons[i].setEnabled(True)

                self._update_start_test_button_state() # Atualiza o estado do botão "Iniciar Teste"

                # Inicia a thread de leitura para a porta
                new_reader_thread = SerialReaderThread(current_ser, log_prefix, is_modbus_port=is_modbus) # Passa a flag is_modbus_port
                new_reader_thread.data_received.connect(self._display_received_data)
                new_reader_thread.connection_lost.connect(self._handle_connection_lost)
                setattr(self, reader_thread_attr, new_reader_thread)
                new_reader_thread.start()

            except serial.SerialException as e:
                QMessageBox.critical(self, "Erro de Conexão", f"Não foi possível conectar à porta serial ({log_prefix}):\n{e}")
                self._disconnect_port_ui(port_type, reset_selection=True) # Garante que a UI seja resetada em caso de falha
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Ocorreu um erro inesperado ao conectar ({log_prefix}):\n{e}")
        else:
            # Se a porta já está aberta, tenta fechá-la (fechamento manual)
            self._disconnect_port_ui(port_type, reset_selection=False) # Não reseta a seleção
            self.log_message(f"Conexão serial ({log_prefix}) fechada.", "sistema")
            connect_button.setText(button_text_close)
            
            port_combobox.setEnabled(True) # Habilita o combobox da porta novamente (mantém a seleção)

            if port_type == "serial_command":
                # Desabilita campos de envio direto e automático para a porta principal
                if not (self.modbus_ser and self.modbus_ser.is_open):
                    self.direct_command_input.setEnabled(False)
                    self.direct_send_button.setEnabled(False)
                    self.send_target_port_combo.setEnabled(False)
                    self.modbus_display_text_cb.setEnabled(False)
                    for i in range(len(self.send_command_inputs)):
                        self.send_command_inputs[i].setEnabled(False)
                        self.send_buttons[i].setEnabled(False)
                        self.auto_send_checkboxes[i].setEnabled(False)
                        self.auto_send_checkboxes[i].setChecked(False) # Desativa o auto-envio
                        self.auto_send_config_buttons[i].setEnabled(False)
            
            self._update_start_test_button_state()
            self._stop_test_if_connection_lost() # Verifica se o teste deve ser parado

    def _disconnect_port_ui(self, port_type, reset_selection=True):
        """
        Desconecta a porta serial e para a thread de leitura associada.
        Usado internamente para garantir uma desconexão limpa.
        Adicionado 'reset_selection' para controlar se o combobox deve ser limpo.
        """
        reader_thread_attr = 'serial_command_reader_thread' if port_type == "serial_command" else 'modbus_serial_reader_thread'
        ser_attr = 'serial_command_ser' if port_type == "serial_command" else 'modbus_ser'
        port_combobox = self.serial_command_port_combobox if port_type == "serial_command" else self.modbus_port_combobox
        connect_button = self.connect_serial_command_button if port_type == "serial_command" else self.connect_modbus_button

        reader_thread = getattr(self, reader_thread_attr)
        if reader_thread:
            reader_thread.stop() # Para a thread de leitura
            setattr(self, reader_thread_attr, None) # Remove a referência à thread
        
        current_ser = getattr(self, ser_attr)
        if current_ser and current_ser.is_open:
            try:
                current_ser.close() # Fecha a porta serial
            except Exception as e:
                self.log_message(f"Erro ao fechar a porta serial ({port_type}) em disconnect_port_ui: {e}", "erro")
        setattr(self, ser_attr, None) # Remove a referência à instância serial

        # Desconecta os sinais dos checkboxes DTR/RTS para evitar erros
        if port_type == "serial_command":
            try:
                self.serial_dtr_checkbox.stateChanged.disconnect()
            except TypeError: # Sinal não conectado
                pass
            try:
                self.serial_rts_checkbox.stateChanged.disconnect()
            except TypeError: # Sinal não conectado
                pass
        elif port_type == "modbus":
            try:
                self.modbus_dtr_checkbox.stateChanged.disconnect()
            except TypeError: # Sinal não conectado
                pass
            try:
                self.modbus_rts_checkbox.stateChanged.disconnect()
            except TypeError: # Sinal não conectado
                pass
        
        if reset_selection:
            # Se a porta foi perdida, limpa a seleção e adiciona a mensagem de "Nenhuma porta..."
            if port_combobox.count() > 0:
                port_combobox.clear()
            port_combobox.addItem("Nenhuma porta COM encontrada")
            port_combobox.setCurrentText("Nenhuma porta COM encontrada")
        # Garante que os controles voltem ao estado de porta fechada
        port_combobox.setEnabled(True)
        if port_type == "serial_command":
            # Atualiza texto do botão para refletir estado de porta fechada
            connect_button.setText("Abrir Porta Principal")
            # Desabilita campos de envio direto/linhas quando a conexão cai
            self.direct_command_input.setEnabled(False)
            self.direct_send_button.setEnabled(False)
            for i in range(len(self.send_command_inputs)):
                self.send_command_inputs[i].setEnabled(False)
                self.send_buttons[i].setEnabled(False)
                self.auto_send_checkboxes[i].setEnabled(False)
                self.auto_send_checkboxes[i].setChecked(False)
                self.auto_send_config_buttons[i].setEnabled(False)
                self.auto_send_timers[i].stop()
        else:
            connect_button.setText("Abrir Porta Modbus")


    def _toggle_dtr(self, state, port_type):
        """
        Alterna a linha DTR para a porta serial especificada.
        """
        ser_instance = None
        log_prefix = ""
        if port_type == "serial_command":
            ser_instance = self.serial_command_ser
            log_prefix = "Principal"
        elif port_type == "modbus":
            ser_instance = self.modbus_ser
            log_prefix = "Modbus"

        if ser_instance and ser_instance.is_open:
            try:
                new_state = (state == Qt.CheckState.Checked.value)
                ser_instance.dtr = new_state
                # self.log_message(f"DTR da porta {log_prefix} definido para: {new_state}", "informacao") # Removido para reduzir verbosidade
            except Exception as e:
                self.log_message(f"Erro ao definir DTR para a porta {log_prefix}: {e}", "erro")
                # Reverte o estado do checkbox se a definição falhou
                if port_type == "serial_command":
                    self.serial_dtr_checkbox.blockSignals(True)
                    self.serial_dtr_checkbox.setChecked(not new_state)
                    self.serial_dtr_checkbox.blockSignals(False)
                elif port_type == "modbus":
                    self.modbus_dtr_checkbox.blockSignals(True)
                    self.modbus_dtr_checkbox.setChecked(not new_state)
                    self.modbus_dtr_checkbox.blockSignals(False)
        else:
            self.log_message(f"Porta {log_prefix} não está aberta para definir DTR.", "erro")
            # Garante que o checkbox esteja desmarcado se a porta não estiver aberta
            if port_type == "serial_command":
                self.serial_dtr_checkbox.blockSignals(True)
                self.serial_dtr_checkbox.setChecked(False)
                self.serial_dtr_checkbox.blockSignals(False)
            elif port_type == "modbus":
                self.modbus_dtr_checkbox.blockSignals(True)
                self.modbus_dtr_checkbox.setChecked(False)
                self.modbus_dtr_checkbox.blockSignals(False)

    def _toggle_rts(self, state, port_type):
        """
        Alterna a linha RTS para a porta serial especificada.
        """
        ser_instance = None
        log_prefix = ""
        if port_type == "serial_command":
            ser_instance = self.serial_command_ser
            log_prefix = "Principal"
        elif port_type == "modbus":
            ser_instance = self.modbus_ser
            log_prefix = "Modbus"

        if ser_instance and ser_instance.is_open:
            try:
                new_state = (state == Qt.CheckState.Checked.value)
                ser_instance.rts = new_state
                # self.log_message(f"RTS da porta {log_prefix} definido para: {new_state}", "informacao") # Removido para reduzir verbosidade
            except Exception as e:
                self.log_message(f"Erro ao definir RTS para a porta {log_prefix}: {e}", "erro")
                # Reverte o estado do checkbox se a definição falhou
                if port_type == "serial_command":
                    self.serial_rts_checkbox.blockSignals(True)
                    self.serial_rts_checkbox.setChecked(not new_state)
                    self.serial_rts_checkbox.blockSignals(False)
                elif port_type == "modbus":
                    self.modbus_rts_checkbox.blockSignals(True)
                    self.modbus_rts_checkbox.setChecked(not new_state)
                    self.modbus_rts_checkbox.blockSignals(False)
        else:
            self.log_message(f"Porta {log_prefix} não está aberta para definir RTS.", "erro")
            # Garante que o checkbox esteja desmarcado se a porta não estiver aberta
            if port_type == "serial_command":
                self.serial_rts_checkbox.blockSignals(True)
                self.serial_rts_checkbox.setChecked(False)
                self.serial_rts_checkbox.blockSignals(False)
            elif port_type == "modbus":
                self.modbus_rts_checkbox.blockSignals(True)
                self.modbus_rts_checkbox.setChecked(False)
                self.modbus_rts_checkbox.blockSignals(False)


    def _update_start_test_button_state(self):
        """
        Atualiza o estado (habilitado/desabilitado) do botão 'Iniciar Teste'
        com base nas conexões das portas e se um teste foi carregado.
        """
        can_start = False
        if self.current_test_steps: # Verifica se há passos de teste carregados
            serial_command_connected = (self.serial_command_ser is not None and self.serial_command_ser.is_open)
            modbus_connected = (self.modbus_ser is not None and self.modbus_ser.is_open)

            # Um teste pode requerer Modbus se houver qualquer passo do tipo "modbus_comando"
            modbus_required_by_steps = any(step.get("tipo_passo") == "modbus_comando" for step in self.current_test_steps)
            self.modbus_required_for_test = modbus_required_by_steps # Atualiza a flag global

            if self.modbus_required_for_test:
                can_start = serial_command_connected and modbus_connected # Requer ambas as portas
            else:
                can_start = serial_command_connected # Requer apenas a porta principal
        
        self.start_test_button.setEnabled(can_start and not self.test_in_progress)

    def _handle_connection_lost(self, port_name):
        """
        Lida com a perda de conexão de uma porta serial.
        Exibe um aviso e tenta desconectar a porta afetada.
        """
        # Ignora perda da Principal durante fechamento temporário para gravação
        try:
            if getattr(self, '_flash_temporarily_closing_principal', False) and port_name == "Principal":
                return
        except Exception:
            pass
        self.log_message(f"AVISO: A conexão com a porta serial '{port_name}' foi perdida. Desconectando.", "erro")
        
        # Determina qual porta foi perdida e desconecta sua UI, resetando a seleção
        if port_name == "Principal":
            self._disconnect_port_ui("serial_command", reset_selection=True)
        elif port_name == "Modbus":
            self._disconnect_port_ui("modbus", reset_selection=True)
        
        QMessageBox.warning(self, "Conexão Perdida", f"A conexão com a porta serial '{port_name}' foi perdida. Porta será fechada.")
        
        # Re-lista as portas seriais para atualizar as portas disponíveis nos comboboxes
        self._list_serial_ports() 
        
        self._update_start_test_button_state()
        self._stop_test_if_connection_lost()
        self.fast_mode_active = False # Desativa o modo fast se a conexão principal for perdida
        self._update_fast_mode_status_label() # Atualiza o rótulo do modo fast

    def _stop_test_if_connection_lost(self):
        """
        Verifica se o teste em andamento deve ser interrompido devido à perda de conexão
        de uma porta serial necessária.
        """
        if self.test_in_progress:
            # Não interrompe se o fechamento da Principal é temporário para gravação
            if getattr(self, '_flash_temporarily_closing_principal', False):
                return
            if (self.modbus_required_for_test and (self.modbus_ser is None or not self.modbus_ser.is_open)) or \
               (self.serial_command_ser is None or not self.serial_command_ser.is_open):
                self._stop_test() # Interrompe o teste se uma porta necessária for perdida

    def _send_direct_command(self):
        """
        Envia um comando digitado diretamente no campo de entrada do terminal.
        O Modo Fast é ativado/desativado se o comando for o código secreto,
        seja digitando e pressionando Enter ou clicando no botão 'Enviar'.
        """
        # Se a configuração da porta estiver recolhida, expande-a
        if not self.serial_details_widget.isVisible():
            self._update_port_config_visibility(True)

        # Se o grupo de progresso do teste estiver visible e não houver teste em andamento, limpa-o
        if self.test_progress_group.isVisible() and not self.test_in_progress:
            self.test_progress_group.setVisible(False)
            self.test_progress_list.clear() 
            self.test_progress_label.setText("Status Geral: Nenhum teste em execução")

        command = self.direct_command_input.text().strip()
        if not command:
            return # Não envia comando vazio

        # Segredo para abrir popup de configurações do CallMeBot
        if command.upper() == "OTAVIO2020":
            try:
                self._open_callmebot_settings_dialog()
            finally:
                self.direct_command_input.clear()
            return

        # Verifica se o comando é o código secreto do modo fast
        if self.fast_mode_secret_code and command == self.fast_mode_secret_code:
            self.fast_mode_active = not self.fast_mode_active
            self._update_fast_mode_status_label()
            self.direct_command_input.clear() # Limpa o input após alternar o modo
            # Não envia o código secreto para a porta serial se ele for usado apenas para alternar o modo interno
            return 

        # Seleciona a porta alvo de acordo com o seletor
        target_name = self.send_target_port_combo.currentText() if hasattr(self, 'send_target_port_combo') else "Principal"
        target_ser = self.serial_command_ser if target_name == "Principal" else self.modbus_ser

        if target_ser and target_ser.is_open:
            try:
                target_ser.reset_input_buffer() # Limpa o buffer de entrada antes de enviar
                
                # Adiciona uma linha vazia antes do envio
                self.log_message("\n", "informacao") 
                target_ser.write(command.encode()) # Envia o comando codificado
                self.log_message(f"Enviado ({target_name}): {command.strip()}", "enviado")
                self.direct_command_input.clear() # Limpa o campo de entrada
            except serial.SerialException as e:
                QMessageBox.critical(self, "Erro de Envio", f"Erro ao enviar comando:\n{e}\nConexão pode ter sido perdida.")
                self._handle_connection_lost(target_name)
            except Exception as e:
                QMessageBox.critical(self, "Erro de Envio", f"Ocorreu um erro inesperado ao enviar o comando:\n{e}")
        else:
            QMessageBox.warning(self, "Erro de Envio", f"Conecte-se à porta {target_name} primeiro.")

    def _send_command_for_auto_send(self, index):
        """
        Envia um comando de uma das linhas de auto-envio.
        Chamado pelo QTimer de auto-envio.
        """
        command_input = self.send_command_inputs[index]
        command = command_input.text().strip()

        if not command:
            self.auto_send_timers[index].stop() # Para o timer se o comando estiver vazio
            self.auto_send_checkboxes[index].setChecked(False)
            self.log_message(f"Auto-envio {index+1} parado: comando vazio.", "informacao")
            return

        target_name = self.send_target_port_combo.currentText() if hasattr(self, 'send_target_port_combo') else "Principal"
        target_ser = self.serial_command_ser if target_name == "Principal" else self.modbus_ser
        if target_ser and target_ser.is_open:
            try:
                target_ser.reset_input_buffer()
                
                # Adiciona uma linha vazia antes do envio
                self.log_message("\n", "informacao")
                target_ser.write(command.encode())
                self.log_message(f"Enviado ({target_name}): {command.strip()}", "enviado")
            except serial.SerialException as e:
                self.log_message(f"Erro no auto-envio {index+1}: {e}", "erro")
                self.auto_send_timers[index].stop()
                self.auto_send_checkboxes[index].setChecked(False)
                self._handle_connection_lost(target_name)
            except Exception as e:
                self.log_message(f"Erro inesperado no auto-envio {index+1}: {e}", "erro")
                self.auto_send_timers[index].stop()
                self.auto_send_checkboxes[index].setChecked(False)
        else:
            self.log_message(f"Porta {target_name} não aberta para auto-envio {index+1}.", "erro")
            self.auto_send_timers[index].stop()
            self.auto_send_checkboxes[index].setChecked(False)

    def _send_command_button_clicked(self, index):
        """
        Envia um comando de uma das linhas de comando pré-configuradas (acionado por botão).
        """
        # Se a configuração da porta estiver recolhida, expande-a
        if not self.serial_details_widget.isVisible():
            self._update_port_config_visibility(True)

        # Se o grupo de progresso do teste estiver visible e não houver teste em andamento, limpa-o
        if self.test_progress_group.isVisible() and not self.test_in_progress:
            self.test_progress_group.setVisible(False)
            self.test_progress_list.clear()
            self.test_progress_label.setText("Status Geral: Nenhum teste em execução")

        command_input = self.send_command_inputs[index]
        command = command_input.text().strip()

        if not command:
            return

        target_name = self.send_target_port_combo.currentText() if hasattr(self, 'send_target_port_combo') else "Principal"
        target_ser = self.serial_command_ser if target_name == "Principal" else self.modbus_ser
        if target_ser and target_ser.is_open:
            try:
                target_ser.reset_input_buffer()
                
                # Adiciona uma linha vazia antes do envio
                self.log_message("\n", "informacao")
                target_ser.write(command.encode())
                self.log_message(f"Enviado ({target_name}): {command.strip()}", "enviado")
            except serial.SerialException as e:
                QMessageBox.critical(self, "Erro de Envio", f"Erro ao enviar comando:\n{e}\nConexão pode ter sido perdida.")
                self._handle_connection_lost(target_name)
            except Exception as e:
                QMessageBox.critical(self, "Erro de Envio", f"Ocorreu um erro inesperado ao enviar o comando:\n{e}")
        else:
            QMessageBox.warning(self, "Erro de Envio", f"Conecte-se à porta {target_name} primeiro.")

    def _open_timer_config_dialog(self, index):
        """
        Abre o diálogo de configuração do timer para uma linha de auto-envio específica.
        """
        current_interval = self.auto_send_intervals_s[index]
        dialog = TimerConfigDialog(current_interval, self)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_interval = dialog.get_interval()
            if new_interval is not None:
                self.auto_send_intervals_s[index] = new_interval
                
                if self.auto_send_checkboxes[index].isChecked():
                    self._toggle_auto_send_timer(Qt.CheckState.Checked.value, index) # Reinicia o timer com o novo intervalo

    def _toggle_auto_send_timer(self, state, index):
        """
        Inicia ou para o timer de auto-envio para uma linha específica.
        """
        timer = self.auto_send_timers[index]
        command_input = self.send_command_inputs[index]
        checkbox = self.auto_send_checkboxes[index]
        
        interval_seconds = self.auto_send_intervals_s[index]

        if state == Qt.CheckState.Checked.value:
            target_name = self.send_target_port_combo.currentText() if hasattr(self, 'send_target_port_combo') else "Principal"
            target_ser = self.serial_command_ser if target_name == "Principal" else self.modbus_ser
            if not target_ser or not target_ser.is_open:
                QMessageBox.warning(self, "Conexão Necessária", f"Conecte-se à porta {target_name} antes de ativar o envio automático.")
                checkbox.setChecked(False)
                return

            if not command_input.text().strip():
                QMessageBox.warning(self, "Comando Vazio", "Preencha o campo de comando antes de ativar o envio automático.")
                checkbox.setChecked(False)
                return

            try:
                if interval_seconds <= 0:
                    interval_seconds = 1.0 # Garante um intervalo mínimo
                    self.auto_send_intervals_s[index] = interval_seconds
                
                interval_ms = int(round(interval_seconds * 1000))
                
                if interval_ms < 10: # Garante um mínimo de 10ms para o timer
                    interval_ms = 10
                
                timer.stop() # Para o timer antes de iniciá-lo novamente
                timer.start(interval_ms)
                self.log_message(f"Auto-envio {index+1} iniciado com intervalo de {interval_seconds}s.", "informacao")
            except Exception as e:
                QMessageBox.critical(self, "Erro Inesperado", f"Ocorreu um erro inesperado ao configurar o timer: {e}")
                checkbox.setChecked(False)
        else:
            timer.stop()
            self.log_message(f"Auto-envio {index+1} parado.", "informacao")

    def _display_received_data(self, data, source_port_name):
        """
        Exibe os dados recebidos no log do terminal.
        Se a 'data' for uma string vazia, isso representa uma linha em branco
        enviada pelo dispositivo, e será logada como tal.
        """
        # Se for Modbus (bytes), exibe como texto somente se for majoritariamente imprimível; senão, usa HEX
        if source_port_name == "Modbus" and isinstance(data, bytes):
            try:
                if hasattr(self, "modbus_display_text_cb") and self.modbus_display_text_cb.isChecked():
                    # Heurística: considera texto se >=90% dos bytes forem imprimíveis (ASCII) ou controles \t\r\n
                    total = max(1, len(data))
                    printable = sum(1 for x in data if (32 <= x <= 126) or x in (9, 10, 13))
                    if printable / total >= 0.9 and b"\x00" not in data:
                        txt = data.decode("latin-1", errors="replace")
                        self.log_message(f"Recebido (Modbus Texto): {txt}", "recebido")
                    else:
                        display_data = data.hex().upper()
                        self.log_message(f"Recebido (Modbus): {display_data}", "recebido")
                else:
                    display_data = data.hex().upper()
                    self.log_message(f"Recebido (Modbus): {display_data}", "recebido")
            except Exception:
                pass
        else: # Se for serial principal (string)
            self.log_message(f"Recebido: {data}", "recebido") 

        # Verifica o código secreto do modo fast
        # A validação do código secreto só faz sentido para a porta principal
        if source_port_name == "Principal" and self.fast_mode_secret_code and isinstance(data, str) and data.strip() == self.fast_mode_secret_code:
            self.fast_mode_active = not self.fast_mode_active # Alterna o estado
            self._update_fast_mode_status_label()
            return # Não processa como dado normal se for o código secreto

    def _update_fast_mode_status_label(self):
        """
        Atualiza o rótulo de status do Modo Fast na interface.
        Este método ainda é chamado, mas o rótulo não está mais no layout.
        """
        status = "ATIVADO" if self.fast_mode_active else "DESATIVADO"        
        self.log_message(f"Modo Fast {'ATIVADO' if self.fast_mode_active else 'DESATIVADO'}!", "sistema")

    def log_message(self, message, msg_type="informacao"):
        # Ignorar mensagens técnicas no terminal
        for prefix in ("Recebido:", "Enviado:", "Comando Enviado:", "Resposta Coletada para Validação:"):
            if message.startswith(prefix):
                message = message.replace(prefix, "").strip()
                break

        color = self.LOG_COLORS.get(msg_type, self.LOG_COLORS["informacao"])

        self.log_interaction_count += 1
        if self.log_interaction_count >= 1500: # Limite de interações para limpar o log
            self.log_text_edit.clear()
            self.log_text_edit.append(f"<font color='#AAAAAA'>--- Log Limpo ---</font>")
            self.log_interaction_count = 0
        # Evita inserir mensagens duplicadas consecutivas
        try:
            if not hasattr(self, "_last_log_line"):
                self._last_log_line = ""
            full_line = f"<font color='{color}'>{message}</font>"
            if full_line == getattr(self, "_last_log_line", ""):
                return
            self._last_log_line = full_line
        except Exception:
            pass

        self.log_text_edit.setUpdatesEnabled(False)
        self.log_text_edit.append(full_line)
        self.log_text_edit.setUpdatesEnabled(True)
        self.log_text_edit.verticalScrollBar().setValue(self.log_text_edit.verticalScrollBar().maximum())
        
        # Removido: o gatilho baseado no texto do log foi substituído por uma chamada direta em _on_new_dir_detected
        
    def _show_log_context_menu(self, pos):
        """
        Exibe o menu de contexto para o QTextEdit do log.
        """
        context_menu = QMenu(self)
        clear_action = QAction("Limpar Terminal", self)
        clear_action.triggered.connect(self._clear_terminal_log)
        context_menu.addAction(clear_action)
        context_menu.exec(self.log_text_edit.mapToGlobal(pos))

    def _clear_terminal_log(self):
        """
        Limpa o conteúdo do QTextEdit do log.
        """
        self.log_text_edit.clear()
        self.log_text_edit.append(f"<font color='#AAAAAA'>--- Terminal Limpo ---</font>")
        self.log_interaction_count = 0 # Reseta o contador de interações


    def _load_test_file(self):
        """        self.fast_mode_code_input.setEnabled(True)
        self.fast_mode_code_input.setText("")
        Carrega um arquivo de teste JSON, validando sua estrutura.
        Popula os passos de teste e as configurações de porta.
        """
        file_path, _ = QFileDialog.getOpenFileName(self, "Carregar Arquivo de Teste", "", "Arquivos JSON (*.json);;Todos os Arquivos (*)")
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    loaded_data = json.load(f)
                    
                    file_version = loaded_data.get("version", "0.0")
                    if file_version == "0.0": # Compatibilidade com versões antigas do arquivo
                        loaded_steps = loaded_data
                        # Re-avalia modbus_required_for_test para arquivos antigos
                        self.modbus_required_for_test = any(step.get("tipo_validacao") == "modbus" for step in loaded_steps if step.get("tipo_passo") == "comando_validacao")
                        self.test_serial_command_settings = {}
                        self.test_modbus_settings = {}
                        self.fast_mode_secret_code = "" # Não existia em versões antigas
                    else: # Formato de arquivo mais recente (1.1 para incluir modo fast, 1.2 para tempo de espera, 1.3 para tabela modbus, 1.4 para modbus exclusivo)
                        loaded_steps = loaded_data.get("steps", [])
                        self.modbus_required_for_test = loaded_data.get("modbus_required", False)
                        self.fast_mode_secret_code = loaded_data.get("fast_mode_code", "") # Carrega o código secreto
                        port_configs = loaded_data.get("port_configurations", {})
                        self.test_serial_command_settings = port_configs.get("serial_command", {})
                        self.test_modbus_settings = port_configs.get("modbus", {})

                    if not isinstance(loaded_steps, list):
                        raise ValueError("O conteúdo do arquivo não é uma lista de passos de teste.")
                    
                    # Validação da estrutura de cada passo
                    for step in loaded_steps:
                        if 'list_item' in step: # Remove a referência ao item da lista, se presente
                            del step['list_item']
                        # Garante que 'checked_for_fast_mode' exista, mesmo para arquivos antigos
                        step['checked_for_fast_mode'] = step.get('checked_for_fast_mode', False)

                        # Tentativa de correção: inferir tipo_passo se ausente para passos conhecidos
                        if "tipo_passo" not in step:
                            try:
                                if all(k in step for k in ["pergunta_gravada", "texto_como_gravar", "caminho_cmd"]):
                                    step["tipo_passo"] = "gravar_placa"
                            except Exception:
                                pass

                        if not all(k in step for k in ["nome", "tipo_passo"]):
                            raise ValueError("Um ou mais passos no arquivo estão mal formatados (faltando 'nome' ou 'tipo_passo').")
                        
                        if step["tipo_passo"] == "comando_validacao":
                            step["port_type"] = step.get("port_type", "serial") 
                            if not all(k in step for k in ["comando_enviar", "esperar_resposta", "timeout_ms", "tipo_validacao"]):
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'comando_validacao' está mal formatado.")
                            
                            val_type = step["tipo_validacao"]
                            # Valida parâmetros específicos de cada tipo de validação
                            if val_type == "string_exata" and "param_validacao" not in step:
                                raise ValueError(f"Validação '{val_type}' no passo '{step.get('nome', 'N/A')}' faltando 'param_validacao'.")
                            elif val_type == "numerico_faixa":
                                if not all(k in step.get("param_validacao", {}) for k in ["min", "max"]):
                                    raise ValueError(f"Validação '{val_type}' no passo '{step.get('nome', 'N/A')}' faltando 'min' ou 'max' nos parâmetros.")
                            elif val_type == "texto_numerico_simples":
                                if not all(k in step.get("param_validacao", {}) for k in ["simplified_text", "regex", "min", "max"]):
                                    if "regex" in step.get("param_validacao", {}) and "min" in step.get("param_validacao", {}) and "max" in step.get("param_validacao", {}):
                                        if "[VALOR]" in step["param_validacao"]["regex"]:
                                            step["param_validacao"]["simplified_text"] = step["param_validacao"]["regex"].replace(r"\s*([-+]?\d*\.?\d+)\s*", "[VALOR]")
                                        else:
                                            step["param_validacao"]["simplified_text"] = ""
                                    else:
                                        raise ValueError(f"Validação '{val_type}' no passo '{step.get('nome', 'N/A')}' faltando parâmetros essenciais.")
                            # Removido o caso 'modbus' daqui para compatibilidade com a nova estrutura
                            if val_type == "modbus": # Se for um arquivo antigo com modbus como validação
                                self.log_message(f"AVISO: Passo '{step.get('nome', 'N/A')}' tem validação Modbus antiga. Converta-o para o novo tipo de passo 'Comando Modbus' para melhor funcionalidade.", "informacao")
                                # O sistema continuará a carregar, mas o usuário será avisado para atualizar.
                                # Para evitar erros, vamos garantir que param_validacao seja uma lista vazia se não for um tipo de validação esperado.
                                if val_type not in ["nenhuma", "string_exata", "numerico_faixa", "texto_numerico_simples"]:
                                    step["param_validacao"] = []
                                    step["tipo_validacao"] = "nenhuma" # Reseta para nenhuma validação

                        elif step["tipo_passo"] == "instrucao_manual":
                            if "mensagem_instrucao" not in step:
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'instrucao_manual' faltando 'mensagem_instrucao'.")
                            step["caminho_imagem"] = step.get("caminho_imagem", "")

                        elif step["tipo_passo"] == "tempo_espera":
                            if "duracao_espera_segundos" not in step:
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'tempo_espera' faltando 'duracao_espera_segundos'.")
                            if not isinstance(step["duracao_espera_segundos"], (int, float)) or step["duracao_espera_segundos"] <= 0:
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'tempo_espera' tem duração inválida.")
                        
                        elif step["tipo_passo"] == "modbus_comando": # NOVO: Validação para o passo Modbus
                            if "modbus_params" not in step or not isinstance(step["modbus_params"], list):
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'modbus_comando' faltando 'modbus_params' ou formato inválido.")
                            for entry in step["modbus_params"]:
                                if not all(k in entry for k in ["slave_id", "function_code", "address", "quantity", "value_type", "write_value", "expected_value", "min_limit", "max_limit"]):
                                    raise ValueError(f"Passo '{entry.get('function_code_display', 'N/A')}' do tipo 'modbus_comando' faltando parâmetros Modbus na entrada da tabela.")
                                if "function_code_display" not in entry:
                                    entry["function_code_display"] = {
                                        "01": "Read Coils (0x01)", "02": "Read Discrete Inputs (0x02)",
                                        "03": "Read Holding Registers (0x03)", "04": "Read Input Registers (0x04)",
                                        "05": "Write Single Coil (0x05)", "06": "Write Single Register (0x06)"
                                    }.get(entry["function_code"], "Read Holding Registers (0x03)")
                        elif step["tipo_passo"] == "gravar_placa":
                            # Validação do novo tipo 'Gravar Placa'
                            if not all(k in step for k in ["pergunta_gravada", "texto_como_gravar", "caminho_cmd"]):
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'gravar_placa' faltando campos obrigatórios.")
                            if not isinstance(step.get("pergunta_gravada"), str) or not step.get("pergunta_gravada").strip():
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'gravar_placa' com 'pergunta_gravada' inválida.")
                            if not isinstance(step.get("texto_como_gravar"), str) or not step.get("texto_como_gravar").strip():
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'gravar_placa' com 'texto_como_gravar' inválido.")
                            if not isinstance(step.get("caminho_cmd"), str) or not step.get("caminho_cmd").strip():
                                raise ValueError(f"Passo '{step.get('nome', 'N/A')}' do tipo 'gravar_placa' com 'caminho_cmd' inválido.")

                    # Determina dinamicamente se o teste requer Modbus (passos novos e antigos)
                    requires_modbus_dynamic = any(
                        (step.get("tipo_passo") == "modbus_comando") or
                        (step.get("tipo_passo") == "comando_validacao" and step.get("port_type", "serial") == "modbus")
                        for step in loaded_steps
                    )
                    # Mantém retrocompatibilidade com a flag gravada no arquivo, mas prioriza a detecção dinâmica
                    self.modbus_required_for_test = bool(self.modbus_required_for_test or requires_modbus_dynamic)

                    self.current_test_steps = loaded_steps # Atribui os passos carregados
                    self.fast_mode_code_input.setText(self.fast_mode_secret_code) # Atualiza o campo na UI

                # Ajusta a visibilidade do grupo Modbus com base na necessidade do teste
                # Agora, a visibilidade do grupo Modbus é controlada pela flag modbus_required_for_test
                self.modbus_serial_group.setVisible(self.modbus_required_for_test)
                self.connect_modbus_button.setEnabled(self.modbus_serial_group.isVisible())
                self.modbus_port_combobox.setEnabled(self.modbus_serial_group.isVisible())
                if self.modbus_serial_group.isVisible():
                    self.modbus_details_widget.setVisible(True)
                    self.modbus_config_toggle_button.setArrowType(Qt.ArrowType.DownArrow)
                    self.modbus_config_toggle_button.setChecked(True)

                self._update_test_port_settings_display() # Atualiza a exibição das configurações de porta
                self._apply_test_port_settings_to_main_ui() # Aplica as configs carregadas nos controles da interface
                self._update_start_test_button_state() # Atualiza o estado do botão "Iniciar Teste"
                self.test_status_label.setText(f"Status do Teste: Carregado '{os.path.basename(file_path)}'")
                self._update_test_steps_list() # Atualiza a lista de passos na UI
                self._clear_step_input_fields() # Limpa os campos de edição
                self.editing_step_index = -1
                self.update_step_button.setEnabled(False)
                self.remove_step_button.setEnabled(False)
                self._update_move_buttons_state()
                
                self._update_port_config_visibility(False) # Recolhe as configurações de porta na aba terminal

            except FileNotFoundError:
                QMessageBox.critical(self, "Erro de Arquivo", "Arquivo de teste não encontrado.")
            except json.JSONDecodeError as e:
                QMessageBox.critical(self, "Erro de JSON", f"Arquivo de teste não é um JSON válido. Verifique a sintaxe.\nDetalhes: {e}")
            except ValueError as e:
                    QMessageBox.critical(self, "Erro de Validação de Teste", f"Erro na estrutura do arquivo de teste: {e}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao carregar o arquivo de teste:\n{e}")

    def _apply_test_port_settings_to_main_ui(self):
        """Aplica na interface (comboboxes principais) as configs de porta carregadas do teste."""
        # Porta Principal
        if hasattr(self, "serial_baud_combo") and self.test_serial_command_settings:
            s = self.test_serial_command_settings
            def _set_if_present(combo, value):
                if value is None:
                    return
                try:
                    text = str(value)
                    if combo.findText(text) != -1:
                        combo.setCurrentText(text)
                except Exception:
                    pass
            _set_if_present(self.serial_baud_combo, s.get("baud"))
            _set_if_present(self.serial_data_bits_combo, s.get("data_bits"))
            _set_if_present(self.serial_parity_combo, s.get("parity"))
            _set_if_present(self.serial_handshake_combo, s.get("handshake"))
            _set_if_present(self.serial_mode_combo, s.get("mode"))

        # Porta Modbus
        if hasattr(self, "modbus_baud_combo") and self.test_modbus_settings:
            m = self.test_modbus_settings
            def _set_if_present_m(combo, value):
                if value is None:
                    return
                try:
                    text = str(value)
                    if combo.findText(text) != -1:
                        combo.setCurrentText(text)
                except Exception:
                    pass
            _set_if_present_m(self.modbus_baud_combo, m.get("baud"))
            _set_if_present_m(self.modbus_data_bits_combo, m.get("data_bits"))
            _set_if_present_m(self.modbus_parity_combo, m.get("parity"))
            _set_if_present_m(self.modbus_handshake_combo, m.get("handshake"))
            _set_if_present_m(self.modbus_mode_combo, m.get("mode"))

        # Atualiza a visibilidade/estado dos grupos e botões conforme necessário
        self.modbus_serial_group.setVisible(self.modbus_required_for_test)
        self.connect_modbus_button.setEnabled(self.modbus_serial_group.isVisible())
        self.modbus_port_combobox.setEnabled(self.modbus_serial_group.isVisible())

    def _prompt_for_test_info(self):
        self._update_port_config_visibility(False)
        """
        Abre um diálogo para o usuário inserir o número do PR e o número de série
        antes de iniciar a execução do teste.
        """
        serial_command_connected = (self.serial_command_ser is not None and self.serial_command_ser.is_open)
        modbus_connected = (self.modbus_ser is not None and self.modbus_ser.is_open)
        self.oled._reposicionar_no_topo()
        self.oled.show()
        # Verifica se as portas necessárias estão conectadas
        if not serial_command_connected:
            QMessageBox.warning(self, "Conexão Necessária", "A Porta Serial Principal não está conectada. Por favor, abra a porta antes de iniciar o teste.")
            return

        # Verifica se o teste requer Modbus e se a porta Modbus está conectada
        modbus_required_by_steps = any(step.get("tipo_passo") == "modbus_comando" for step in self.current_test_steps)
        if modbus_required_by_steps and not modbus_connected:
            self.modbus_serial_group.setVisible(True)
            self.modbus_details_widget.setVisible(True)
            self.modbus_config_toggle_button.setArrowType(Qt.ArrowType.DownArrow)
            self.modbus_config_toggle_button.setChecked(True)
            QMessageBox.warning(self, "Conexão Necessária", "Este teste requer a Porta Modbus, que não está conectada. Por favor, abra a porta antes de iniciar o teste.")
            return

        if not self.current_test_steps:
            QMessageBox.warning(self, "Nenhum Teste Carregado", "Carregue um arquivo de teste antes de iniciar.")
            return

        dialog = TestIdDialog(self.current_pr_number, self.current_serial_number, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            pr, serial = dialog.get_info()
            if not pr or not serial:
                QMessageBox.warning(self, "Informações Faltando", "Número do PR e Número de Série são obrigatórios para iniciar o teste.")
                return

            # Validação de formato
            pr_ok = re.fullmatch(r"\d{5}", pr) is not None
            sn_ok = re.fullmatch(r"\d+/\d+", serial) is not None
            if not pr_ok:
                QMessageBox.warning(self, "PR inválido", "Formato do PR inválido. Use PR + 5 dígitos (ex.: PR04237) ou apenas 5 dígitos (ex.: 04237).")
                return
            if not sn_ok:
                QMessageBox.warning(self, "Número de Série inválido", "Formato do Número de Série inválido. Use apenas dígitos no formato NNNNN/NN (ex.: 16913/155).")
                return

            self.current_pr_number = pr
            self.current_serial_number = serial
            self._save_settings() # Salva as últimas informações de PR/Série
            # Seleciona operador
            selected_user = self.configuracoes_tab.selecionar_usuario_popup()
            if not selected_user:
                QMessageBox.warning(self, "Usuário não selecionado", "É necessário selecionar o operador que está realizando o teste.")
                return
            self.current_tester_name = selected_user
            self._start_new_test_full_run() # Inicia a execução do teste
        else:
            self.log_message("Início do teste cancelado pelo usuário.", "informacao")

    def _start_new_test_full_run(self):
        """
        Prepara todos os passos do teste para uma nova execução completa,
        redefinindo seus status para "Pendente".
        """
        for i, step in enumerate(self.current_test_steps):
            step['status'] = "Pendente"
            step['last_error_detail'] = ""

        self.test_start_time = datetime.now()
        self.oled.iniciar_teste()
        self._start_test_execution() # Inicia a execução do primeiro passo

    def _start_test_execution(self):
        """
        Inicia a execução do teste, configurando o estado da UI e registrando os dados no log.
        """
        from datetime import datetime
        from socket import gethostname
        import platform

        self.test_log_entries = []  # Limpa o log do teste atual
        self.test_start_time = datetime.now()  # Registra o tempo de início

        # Captura informações da máquina e do operador
        machine_name = platform.node()

        if hasattr(self, "configuracoes_tab") and self.configuracoes_tab:
            usuario = self.current_tester_name if hasattr(self, "current_tester_name") else "Desconhecido"
        else:
            usuario = "Desconhecido"

        self.current_test_operator = usuario

        # Cabeçalho do log
        self.test_log_entries.append("--- INÍCIO DO TESTE ---")
        self.test_log_entries.append(f"Data/Hora Início: {self.test_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        self.test_log_entries.append(f"Número do PR: {self.current_pr_number}")
        self.test_log_entries.append(f"Número de Série da Placa: {self.current_serial_number}")
        self.test_log_entries.append(f"Operador do Teste: {usuario}")
        self.test_log_entries.append(f"Máquina de Teste: {machine_name}")
        self.test_log_entries.append("")  # Linha em branco

        # Mensagem de início no painel
        self.log_message("INICIANDO EXECUÇÃO DO TESTE AUTOMÁTICO", "sistema")

        self.test_in_progress = True
        self.current_test_index = 0 # Começa do primeiro passo
        self.passed_steps_count = 0
        self.failed_steps_count = 0
        self.start_test_button.setEnabled(False) # Desabilita o botão Iniciar
        self.stop_test_button.setEnabled(True) # Habilita o botão Parar
        self.load_test_button.setEnabled(False) # Desabilita o botão Carregar
        
        self.direct_command_input.setEnabled(False) # Desabilita envio direto
        self.direct_send_button.setEnabled(False)

        # Desabilita e para todos os auto-envios
        for i in range(len(self.auto_send_checkboxes)):
            self.auto_send_checkboxes[i].setEnabled(False)
            self.auto_send_checkboxes[i].setChecked(False)
            self.auto_send_config_buttons[i].setEnabled(False)
            self.auto_send_timers[i].stop()

        for i in range(len(self.send_command_inputs)):
            self.send_command_inputs[i].setEnabled(False)
            self.send_buttons[i].setEnabled(False)

        # Limpa os buffers de resposta das threads de leitura
        if self.serial_command_reader_thread:
            self.serial_command_reader_thread.clear_response_buffer_for_next_step()
        if self.modbus_serial_reader_thread:
            self.modbus_serial_reader_thread.clear_response_buffer_for_next_step()
        
        self.test_progress_list.clear() # Limpa a lista de progresso
        self.test_progress_group.setVisible(True) # Exibe o grupo de progresso
        
        # Cria os itens na lista de progresso, respeitando o modo fast
        for i, step in enumerate(self.current_test_steps):
            # Se o modo fast está ativo e o passo NÃO está marcado para inclusão, ele não é adicionado à lista de progresso
            if self.fast_mode_active and not step.get('checked_for_fast_mode', False):
                step['status'] = "PULADO" # Marca como pulado no modo fast
                step['list_item'] = None # Garante que não haja item na lista para este passo
                continue 

            item_text = f"Passo {i+1}: {step.get('nome', 'Sem Nome')} - Pendente"
            item = QListWidgetItem(item_text)
            item.setForeground(QBrush(QColor("#AAAAAA"))) # Cor cinza para pendente
            self.test_progress_list.addItem(item)
            step['list_item'] = item # Associa o item da lista ao dicionário do passo
            step['status'] = "Pendente" # Garante que o status inicial seja Pendente

        self._execute_next_test_step() # Inicia a execução do primeiro passo

    def _stop_test(self):
        """
        Interrompe a execução do teste em andamento.
        """
        if self.test_in_progress:
            self.test_timer.stop() # Para qualquer timer de timeout ativo
            self.test_in_progress = False
            self.log_message("EXECUÇÃO DO TESTE INTERROMPIDA PELO USUÁRIO", "sistema")
            self.test_status_label.setText("Status do Teste: Interrompido")
            self.test_progress_label.setText("Status Geral: Teste Interrompido")
            self.test_log_entries.append(f"--- TESTE INTERROMPIDO PELO USUÁRIO ({datetime.now().strftime('%H:%M:%S')}) ---")
            # Marca os passos restantes como "INTERROMPIDO"
            for i in range(self.current_test_index, len(self.current_test_steps)):
                step = self.current_test_steps[i]
                # Apenas atualiza o status se o passo não foi pulado pelo modo fast
                if not (self.fast_mode_active and not step.get('checked_for_fast_mode', False)):
                    if 'list_item' in step and step['list_item'] is not None and step.get('status') not in ['APROVADO', 'REPROVADO']: 
                        self._update_list_item_status(step, "INTERROMPIDO", QColor("#FFA07A"), "Teste interrompido.")
                        self.test_log_entries.append(f"[{datetime.now().strftime('%H:%M:%S')}] PASSO {i+1}: {step['nome']} - Status: INTERROMPIDO (Interrompido pelo Usuário)")
                        step['status'] = "INTERROMPIDO"

            self.oled.cancelar_teste()
            self._finish_test() # Finaliza o teste (gera log, etc.)
    
    def _execute_next_test_step(self):
        """
        Executa o próximo passo na sequência do teste.
        Gerencia diferentes tipos de passos (comando/validação, instrução manual, tempo de espera, Modbus).
        Inclui lógica para pular passos no modo fast.
        """
        # Verifica se as portas necessárias ainda estão abertas
        serial_command_connected = (self.serial_command_ser is not None and self.serial_command_ser.is_open)
        modbus_connected = (self.modbus_ser is not None and self.modbus_ser.is_open)

        if not serial_command_connected:
            self.log_message("AVISO: Porta 'Principal' não está aberta. Finalizando teste.", "erro")
            self.test_log_entries.append(f"  ERRO: Porta 'Principal' não está aberta. Teste finalizado inesperadamente.")
            self._finish_test()
            return
        
        # Pula passos que já foram aprovados (em caso de re-teste) ou que estão ocultos no modo fast
        while self.test_in_progress and self.current_test_index < len(self.current_test_steps):
            step = self.current_test_steps[self.current_test_index]
            if step.get('status') == "APROVADO":
                self.log_message(f"Passo {self.current_test_index + 1}: '{step['nome']}' já APROVADO, pulando.", "informacao")
                self.current_test_index += 1
                continue
            
            # Lógica para pular passos no modo fast
            if self.fast_mode_active and not step.get('checked_for_fast_mode', False):
                self.log_message(f"Modo Fast: Pulando passo oculto {self.current_test_index + 1}: '{step['nome']}'", "informacao")
                self.test_log_entries.append(f"[{datetime.now().strftime('%H:%M:%S')}] PASSO {self.current_test_index+1}: {step['nome']} - Status: PULADO (Modo Fast)")
                step['status'] = "PULADO" # Marca como pulado no modo fast
                self.current_test_index += 1
                continue
            else:
                break # Encontrou um passo para executar

        # Se não há mais passos ou o teste foi interrompido
        if not self.test_in_progress or self.current_test_index >= len(self.current_test_steps):
            self._finish_test()
            return

        step = self.current_test_steps[self.current_test_index]
        self.test_status_label.setText(f"Status do Teste: Executando passo {self.current_test_index + 1}/{len(self.current_test_steps)} - {step['nome']}")
        
        # Atualiza o item da lista de progresso para "Em Execução"
        if 'list_item' in step and step['list_item'] is not None:
            item = step['list_item']
            item_text = f"Passo {self.current_test_index + 1}: {step.get('nome', 'Sem Nome')} - Em Execução..."
            item.setText(item_text)
            item.setForeground(QBrush(QColor("#FFD700"))) # Cor amarela para "Em Execução"
            self.test_progress_list.setCurrentItem(item) # Rola até o item atual

        self.log_message(f"\n--- PASSO {self.current_test_index + 1}/{len(self.current_test_steps)}: {step['nome']} ---", "sistema")
        self.test_log_entries.append("")
        self.test_log_entries.append(f"[{datetime.now().strftime('%H:%M:%S')}] PASSO {self.current_test_index+1}: {step['nome']} - Em Execução")
        
        step_type = step.get("tipo_passo", "comando_validacao")
        step['status'] = "Em Execução" # Marca o passo como em execução

        if step_type == "comando_validacao":
            command_to_send = step.get("comando_enviar", "")
            port_type = step.get("port_type", "serial")
            # Aplica recurso de NS (se habilitado para o passo)
            use_ns = step.get("use_ns", False)
            ns_mode = step.get("ns_mode", "")
            ns_pattern = step.get("ns_pattern", "")
            if hasattr(self, "_build_command_with_ns"):
                command_to_send = self._build_command_with_ns(command_to_send, use_ns, ns_mode, ns_pattern)
            # Aplica recurso de RTC (se habilitado para o passo)
            use_rtc = step.get("use_rtc", False)
            rtc_mode = step.get("rtc_mode", "")
            rtc_pattern = step.get("rtc_pattern", "")
            if hasattr(self, "_build_command_with_rtc"):
                command_to_send = self._build_command_with_rtc(command_to_send, use_rtc, rtc_mode, rtc_pattern)
            
            # Seleciona porta e thread de leitura conforme 'port_type'
            target_ser = self.serial_command_ser
            target_reader = self.serial_command_reader_thread
            target_port_name = "Principal"
            if port_type == "modbus":
                target_ser = self.modbus_ser
                target_reader = self.modbus_serial_reader_thread
                target_port_name = "Modbus"

            # Verifica se a porta selecionada está conectada para este passo
            if target_ser is None or not target_ser.is_open:
                error_msg = f"Porta '{target_port_name}' não está conectada para o passo '{step['nome']}'."
                self.log_message(f"ERRO: {error_msg} Teste falhou neste passo.", "erro")
                self.test_log_entries.append(f"  ERRO: {error_msg}")
                self.failed_steps_count += 1
                self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), error_msg)
                self.current_test_index += 1
                QTimer.singleShot(100, self._execute_next_test_step)
                return

            try:
                # Substitui sequências de escape e codifica o comando
                command_to_send_encoded = command_to_send.replace("\\n", "\n").replace("\\r", "\r").replace("\\t", "\t").encode()
                target_ser.write(command_to_send_encoded)
                self.log_message(f"Comando Enviado: '{command_to_send.strip()}'", "enviado")
                self.test_log_entries.append(f"  Comando Enviado ({target_port_name}): '{command_to_send.strip()}'")
            except serial.SerialException as e:
                error_msg = f"Erro de comunicação serial (Principal): {e}"
                self.log_message(f"ERRO: Falha ao enviar comando para '{step['nome']}': {e}. Teste falhou neste passo.", "erro")
                self.test_log_entries.append(f"  ERRO: Falha ao enviar comando: {e}")
                self.failed_steps_count += 1
                self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), error_msg)
                self.current_test_index += 1
                QTimer.singleShot(100, self._execute_next_test_step)
                return
            except Exception as e:
                error_msg = f"Erro inesperado: {e}"
                self.log_message(f"ERRO INESPERADO: Falha ao enviar comando para '{step['nome']}': {e}. Teste falhou neste passo.", "erro")
                self.test_log_entries.append(f"  ERRO INESPERADO: Falha ao enviar comando: {e}")
                self.failed_steps_count += 1
                self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), error_msg)
                self.current_test_index += 1
                QTimer.singleShot(100, self._execute_next_test_step)
                return

            if step.get("esperar_resposta", False):
                timeout_ms = int(step.get("timeout_ms", 1000))
                if target_reader:
                    target_reader.clear_response_buffer_for_next_step() # Limpa o buffer antes de esperar nova resposta
                # Usa QTimer.singleShot estático para não conflitar com outros usos de self.test_timer
                QTimer.singleShot(max(1, timeout_ms), lambda: self._process_test_response(step, target_reader))
            else:
                # Se não espera resposta, o passo é considerado aprovado imediatamente
                self.log_message(f"APROVADO: '{step['nome']}' (Nenhuma resposta esperada)", "test_pass")
                step_index = self.current_test_steps.index(step) + 1
                self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO")
                self.passed_steps_count += 1
                self._update_list_item_status(step, "APROVADO", QColor("#32CD32"))
                self.current_test_index += 1
                QTimer.singleShot(100, self._execute_next_test_step) # Avança para o próximo passo

        elif step_type == "instrucao_manual":
            instruction_message = step.get("mensagem_instrucao", "Nenhuma instrução fornecida.")
            image_path = step.get("caminho_imagem", "")
            
            self.test_log_entries.append(f"  Tipo: Instrução Manual")
            self.test_log_entries.append(f"  Mensagem de Instrução: '{instruction_message}'")
            if image_path:
                self.test_log_entries.append(f"  Caminho da Imagem: '{image_path}'")
            
            self.test_timer.stop() # Para qualquer timer de timeout anterior
            self.log_message(f"INSTRUÇÃO MANUAL: '{step['nome']}'", "informacao")

            # Abre o diálogo de instrução manual
            manual_dialog = ManualInstructionDialog(
                self.current_test_index + 1, 
                len(self.current_test_steps), 
                step['nome'], 
                instruction_message, 
                image_path, 
                self
            )
            
            result = manual_dialog.exec()
            if result == QDialog.DialogCode.Accepted:
                self.log_message(f"APROVADO: INSTRUÇÃO MANUAL - '{step['nome']}' (Usuário Confirmou)", "test_pass")
                step_index = self.current_test_steps.index(step) + 1
                self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO (Usuário Confirmou)")
                self.passed_steps_count += 1
                self._update_list_item_status(step, "APROVADO", QColor("#32CD32"))
            else:
                # Se o usuário fechou pelo X, interrompe o teste como "Parar Teste"
                if getattr(manual_dialog, 'closed_via_titlebar', False):
                    self.log_message("Teste interrompido pelo usuário ao fechar a janela de Instrução Manual.", "sistema")
                    self._stop_test()
                    return
                self.log_message(f"REPROVADO: INSTRUÇÃO MANUAL - '{step['nome']}' (Usuário Cancelou/Rejeitou)", "test_fail")
                step_index = self.current_test_steps.index(step) + 1
                self.test_log_entries.append(f"  Status: PASSO {step_index}: REPROVADO (Usuário Cancelou)")
                self.failed_steps_count += 1
                self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), "Instrução manual não confirmada pelo usuário.")

            self.current_test_index += 1
            QTimer.singleShot(100, self._execute_next_test_step) # Avança para o próximo passo
        
        elif step_type == "tempo_espera":
            duration_seconds = step.get("duracao_espera_segundos", 1)
            self.log_message(f"TEMPO DE ESPERA: Aguardando {duration_seconds} segundos para o passo '{step['nome']}'...", "informacao")
            self.test_log_entries.append(f"  Tipo: Tempo de Espera")
            self.test_log_entries.append(f"  Duração da Espera: {duration_seconds} segundos")
            
            self.test_timer.singleShot(duration_seconds * 1000, lambda: self._continue_after_wait(step))

        elif step_type == "modbus_comando": # NOVO: Execução do passo de comando Modbus (assíncrono)
            modbus_params_list = step.get("modbus_params", [])

            if not modbus_connected:
                error_msg = f"Porta 'Modbus' não está conectada para o passo '{step['nome']}'."
                self.log_message(f"ERRO: {error_msg} Teste falhou neste passo.", "erro")
                self.test_log_entries.append(f"  ERRO: {error_msg}")
                self.failed_steps_count += 1
                self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), error_msg)
                self.current_test_index += 1
                QTimer.singleShot(100, self._execute_next_test_step)
                return

            if not modbus_params_list:
                error_msg = "Nenhuma configuração Modbus encontrada para este passo."
                self.log_message(f"ERRO: {error_msg}. Teste falhou neste passo.", "erro")
                self.test_log_entries.append(f"  ERRO: {error_msg}")
                self.failed_steps_count += 1
                self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), error_msg)
                self.current_test_index += 1
                QTimer.singleShot(100, self._execute_next_test_step)
                return

            self.test_log_entries.append(f"  Tipo: Comando Modbus")

            all_passed = True
            overall_error_msg = ""

            def finish_step():
                if all_passed:
                    self.log_message(f"APROVADO: '{step['nome']}' (Comando Modbus e validação concluídos com sucesso)", "test_pass")
                    step_index = self.current_test_steps.index(step) + 1
                    self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO")
                    self.passed_steps_count += 1
                    self._update_list_item_status(step, "APROVADO", QColor("#32CD32"))
                else:
                    self.log_message(f"REPROVADO: '{step['nome']}' (Falha no comando Modbus ou validação: {overall_error_msg})", "test_fail")
                    step_index = self.current_test_steps.index(step) + 1
                    self.test_log_entries.append(f"  Status: PASSO {step_index}: REPROVADO")
                    self.test_log_entries.append(f"  Detalhe do Erro: {overall_error_msg}")
                    self.failed_steps_count += 1
                    self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), overall_error_msg)

                self.current_test_index += 1
                QTimer.singleShot(100, self._execute_next_test_step)

            def process_entry(i: int):
                nonlocal all_passed, overall_error_msg
                if not self.test_in_progress:
                    return
                if i >= len(modbus_params_list):
                    finish_step()
                    return

                entry = modbus_params_list[i]
                slave_id = entry.get("slave_id", 1)
                function_code_display = entry.get("function_code_display", "Read Holding Registers (0x03)")
                function_code_hex = entry.get("function_code", "03")
                address = entry.get("address", 0)
                quantity = entry.get("quantity", 1)
                write_value = entry.get("write_value")
                value_type = entry.get("value_type")
                expected_value = entry.get("expected_value")
                min_limit = entry.get("min_limit")
                max_limit = entry.get("max_limit")

                self.log_message(f"  Executando Modbus (Linha {i+1}): {function_code_display} End: {address}, Qtd: {quantity}", "informacao")
                self.test_log_entries.append(f"    Modbus (Linha {i+1}): {function_code_display} End: {address}, Qtd: {quantity}")

                try:
                    determined_value_format = None
                    if write_value:
                        if write_value.lower().startswith("0x"):
                            determined_value_format = "HEX"
                        elif write_value.lower().startswith("0b"):
                            determined_value_format = "BIN"
                        else:
                            determined_value_format = "DEC Unsigned"

                    if "Write" in function_code_display:
                        if not write_value:
                            raise ValueError("Valor para escrita é obrigatório para funções de escrita.")
                        request_bytes = modbus_lib.build_modbus_rtu_request(
                            slave_id=slave_id,
                            function_code_hex=function_code_hex,
                            address=address,
                            quantity_or_value_str=write_value,
                            value_type=value_type,
                            value_format=determined_value_format
                        )
                    else:
                        request_bytes = modbus_lib.build_modbus_rtu_request(
                            slave_id=slave_id,
                            function_code_hex=function_code_hex,
                            address=address,
                            quantity_or_value_str=str(quantity),
                            value_type=value_type
                        )

                    self.modbus_ser.write(request_bytes)
                    self.log_message(f"    Comando Modbus Enviado: {request_bytes.hex().upper()}", "enviado")
                    self.test_log_entries.append(f"      Comando Enviado: {request_bytes.hex().upper()}")

                except Exception as e:
                    all_passed = False
                    overall_error_msg = f"Modbus (Linha {i+1}): Erro na preparação/envio: {e}"
                    finish_step()
                    return

                # Agenda leitura/validação sem bloquear a UI
                def on_response_timeout():
                    nonlocal all_passed, overall_error_msg
                    try:
                        response_data = self.modbus_serial_reader_thread.get_buffered_response()
                        if response_data:
                            self.log_message(f"    Resposta Modbus Recebida:\n'{response_data.hex().upper()}'", "recebido")
                            self.test_log_entries.append(f"      Resposta Recebida: '{response_data.hex().upper()}'")
                        else:
                            self.log_message(f"    Nenhuma resposta Modbus recebida (Linha {i+1}).", "informacao")
                            self.test_log_entries.append(f"      Resposta Recebida: Nenhuma (Linha {i+1})")
                            all_passed = False
                            overall_error_msg = f"Modbus (Linha {i+1}): Nenhuma resposta recebida."
                            finish_step()
                            return

                        success, message, extracted_value = modbus_lib.parse_modbus_rtu_response(
                            response_data,
                            function_code_hex,
                            quantity,
                            value_type,
                            "HEX"
                        )

                        if not success:
                            all_passed = False
                            overall_error_msg = f"Modbus (Linha {i+1}): Erro de parsing da resposta: {message}"
                            finish_step()
                            return

                        if "Read" in function_code_display:
                            if extracted_value is not None:
                                if expected_value:
                                    if isinstance(extracted_value, list):
                                        extracted_value_str = "[" + ", ".join(map(str, extracted_value)) + "]"
                                        if extracted_value_str != expected_value:
                                            all_passed = False
                                            overall_error_msg = f"Modbus (Linha {i+1}): Valor(es) esperado(s) '{expected_value}', mas recebeu '{extracted_value_str}'."
                                            finish_step()
                                            return
                                    else:
                                        if str(extracted_value) != expected_value:
                                            all_passed = False
                                            overall_error_msg = f"Modbus (Linha {i+1}): Valor esperado '{expected_value}', mas recebeu '{extracted_value}'."
                                            finish_step()
                                            return

                                if min_limit is not None and max_limit is not None:
                                    values_to_check = extracted_value if isinstance(extracted_value, list) else [extracted_value]
                                    for val in values_to_check:
                                        try:
                                            num_val = float(val)
                                            if not (min_limit <= num_val <= max_limit):
                                                all_passed = False
                                                overall_error_msg = f"Modbus (Linha {i+1}): Valor '{num_val}' fora da faixa esperada [{min_limit}, {max_limit}]."
                                                finish_step()
                                                return
                                        except ValueError:
                                            all_passed = False
                                            overall_error_msg = f"Modbus (Linha {i+1}): Não foi possível converter '{val}' para número para validação de faixa."
                                            finish_step()
                                            return
                            else:
                                all_passed = False
                                overall_error_msg = f"Modbus (Linha {i+1}): Nenhum valor extraído para validação."
                                finish_step()
                                return

                        self.log_message(f"    Modbus (Linha {i+1}): APROVADO.", "test_pass")
                        self.test_log_entries.append(f"      Status Linha {i+1}: APROVADO")
                        # Avança para próxima linha
                        process_entry(i + 1)
                    except Exception as e:
                        all_passed = False
                        overall_error_msg = f"Modbus (Linha {i+1}): Erro inesperado: {e}"
                        finish_step()

                # Tempo de espera para resposta: usa timeout da porta ou 5000ms como padrão
                try:
                    delay_ms = int(getattr(self.modbus_ser, 'timeout', 5) * 1000) if self.modbus_ser else 5000
                    delay_ms = max(50, min(delay_ms, 10000))
                except Exception:
                    delay_ms = 5000
                QTimer.singleShot(delay_ms, on_response_timeout)

            # Inicia processamento assíncrono
            process_entry(0)

        elif step_type == "gravar_numero_serie":
            # Lê o número de série do settings.json e envia o comando SET_NS=<serial>;
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings_data = json.load(f)
                serial_number = settings_data.get("last_serial_number", "").strip()
                if not serial_number:
                    raise ValueError("last_serial_number vazio no settings.json")
                command_to_send = f"SET_NS={serial_number};"
                # Envia pela porta principal, sem adicionar nova linha
                self.serial_command_ser.write(command_to_send.encode())
                self.log_message(f"Comando Enviado: '{command_to_send}'", "enviado")
                self.test_log_entries.append(f"  Comando Enviado (Gravar NS): '{command_to_send}'")
                # Aprova imediatamente, sem esperar resposta
                self.log_message(f"APROVADO: '{step['nome']}' (Gravação do NS enviada)", "test_pass")
                step_index = self.current_test_steps.index(step) + 1
                self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO")
                self.passed_steps_count += 1
                self._update_list_item_status(step, "APROVADO", QColor("#32CD32"))
            except Exception as e:
                error_msg = f"Falha ao gravar número de série: {e}"
                self.log_message(f"ERRO: {error_msg}", "erro")
                self.test_log_entries.append(f"  ERRO: {error_msg}")
                self.failed_steps_count += 1
                self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), str(e))
            self.current_test_index += 1
            QTimer.singleShot(100, self._execute_next_test_step)

        elif step_type == "gravar_placa":
            # Fluxo: pergunta -> se 'Sim', aprova; se 'Não', mostra instruções com botão 'Gravar' que executa o .cmd e avança
            question = step.get("pergunta_gravada", "A placa está gravada?")
            instructions = step.get("texto_como_gravar", "")
            cmd_path = step.get("caminho_cmd", "")

            self.test_log_entries.append("  Tipo: Gravar Placa")
            self.test_log_entries.append(f"  Pergunta: '{question}'")
            self.test_log_entries.append(f"  CMD: '{cmd_path}'")

            reply = QMessageBox.question(self, "Gravar Placa", question, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.log_message(f"APROVADO: '{step['nome']}' (Operador informou que já está gravada)", "test_pass")
                step_index = self.current_test_steps.index(step) + 1
                self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO (Já estava gravada)")
                self.passed_steps_count += 1
                self._update_list_item_status(step, "APROVADO", QColor("#32CD32"))
            else:
                # Constrói popup com instruções e dois botões
                dlg = QDialog(self)
                dlg.setWindowTitle("Gravar Placa")
                dlg.setWindowFlag(Qt.WindowType.WindowContextHelpButtonHint, False)
                v = QVBoxLayout(dlg)
                label = QLabel("Como gravar a placa:")
                label.setStyleSheet("font-weight: bold;")
                v.addWidget(label)
                txt = QTextEdit()
                txt.setReadOnly(True)
                txt.setPlainText(instructions)
                v.addWidget(txt)

                # Seleção da porta COM a ser enviada ao .cmd (pode ser diferente da Principal)
                port_sel_layout = QHBoxLayout()
                port_sel_label = QLabel("Porta para gravar:")
                port_sel_combo = QComboBox()
                try:
                    ports = [p.device for p in serial.tools.list_ports.comports()]
                except Exception:
                    ports = []
                for d in ports:
                    port_sel_combo.addItem(d)
                # Pré-seleciona a porta Principal se existir na lista
                try:
                    principal_port = ""
                    if self.serial_command_ser is not None:
                        principal_port = getattr(self.serial_command_ser, 'port', '') or getattr(self.serial_command_ser, 'portstr', '')
                    if principal_port:
                        idx = port_sel_combo.findText(principal_port)
                        if idx >= 0:
                            port_sel_combo.setCurrentIndex(idx)
                except Exception:
                    pass
                port_sel_layout.addWidget(port_sel_label)
                port_sel_layout.addWidget(port_sel_combo)
                v.addLayout(port_sel_layout)
                btns = QDialogButtonBox()
                btn_gravar = QPushButton("Gravar")
                btn_cancel = QPushButton("Cancelar")
                btns.addButton(btn_gravar, QDialogButtonBox.ButtonRole.AcceptRole)
                btns.addButton(btn_cancel, QDialogButtonBox.ButtonRole.RejectRole)
                v.addWidget(btns)

                def on_gravar():
                    try:
                        # Se a COM selecionada para gravar é a mesma da Principal, fecha temporariamente a Principal
                        closed_principal = False
                        principal_saved_port = ""
                        try:
                            sel_port_text = ""
                            try:
                                sel_port_text = port_sel_combo.currentText()
                            except Exception:
                                sel_port_text = ""
                            current_principal_text = self.serial_command_port_combobox.currentText() if hasattr(self, 'serial_command_port_combobox') else ""
                            if sel_port_text and current_principal_text and (sel_port_text.upper() == current_principal_text.upper()):
                                principal_saved_port = current_principal_text
                                self._flash_temporarily_closing_principal = True
                                # Fecha temporariamente SEM mexer na UI
                                try:
                                    if getattr(self, 'serial_command_reader_thread', None):
                                        try:
                                            self.serial_command_reader_thread.stop()
                                        except Exception:
                                            pass
                                        self.serial_command_reader_thread = None
                                    if self.serial_command_ser and getattr(self.serial_command_ser, 'is_open', False):
                                        try:
                                            self.serial_command_ser.close()
                                        except Exception:
                                            pass
                                except Exception:
                                    pass
                                closed_principal = True
                                # pequeno atraso para SO liberar a COM
                                time.sleep(0.3)
                        except Exception:
                            pass
                        def run_cmd_file(path: str, send_port: bool, com_number: str):
                            if not path or not os.path.isfile(path):
                                raise FileNotFoundError(path)
                            workdir = os.path.dirname(path) or None
                            # Caminho direto (STMFlashLoader.exe) temporariamente desabilitado para estabilidade

                            # Caminho 2: fallback, cria cópia temporária do .cmd substituindo 'start' por 'call'
                            patched_path = None
                            try:
                                with open(path, 'r', encoding='cp1252', errors='ignore') as f:
                                    src = f.read()
                                patched = re.sub(r"(?im)^(\s*)start\b", r"\1call", src)
                                patched = re.sub(r"(?im)^(\s*)pause\b", r"\1rem pause", patched)
                                tmp = tempfile.NamedTemporaryFile(prefix="flash_", suffix=".cmd", delete=False)
                                patched_path = tmp.name
                                tmp.write(patched.encode('cp1252', errors='ignore'))
                                tmp.close()
                            except Exception:
                                patched_path = None
                            exec_path = patched_path or path
                            p = subprocess.Popen(["cmd", "/c", exec_path], cwd=workdir, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=False)
                            try:
                                if send_port and p.stdin:
                                    com_num = com_number
                                    if not com_num:
                                        # Extrai do texto selecionado (ex.: COM7)
                                        sel = port_sel_combo.currentText()
                                        m = re.search(r"COM(\d+)", sel.upper())
                                        com_num = m.group(1) if m else ""
                                    if com_num:
                                        # Dá um pequeno tempo para o prompt aparecer
                                        time.sleep(0.3)
                                        try:
                                            p.stdin.write((com_num + "\r\n").encode())
                                            p.stdin.flush()
                                            try:
                                                self.log_message(f"[CMD] >> Enviado COM: {com_num}", "informacao")
                                            except Exception:
                                                pass
                                        except Exception:
                                            pass
                            except Exception:
                                pass
                            return p, exec_path if patched_path else None

                        if not cmd_path or not os.path.isfile(cmd_path):
                            QMessageBox.warning(dlg, "Arquivo inválido", f"Selecione um arquivo .cmd válido (não uma pasta):\n{cmd_path}")
                            return

                        cmd2 = step.get("caminho_cmd2", "").strip()

                        # Flags de sucesso
                        success_patterns = ["sucesso", "conclu", "ok"]

                        def monitor_and_finish(proc, after=None, tag="CMD", cleanup_path=None, com_number_for_prompt: str = ""):
                            success_flag = False
                            stop_requested = False
                            sent_press_any = False
                            sent_retry_n = False
                            sent_com = False
                            buf = ""
                            try:
                                while True:
                                    if proc.poll() is not None and not proc.stdout:
                                        break
                                    chunk = b""
                                    try:
                                        chunk = proc.stdout.read(64) if proc.stdout else b""
                                    except Exception:
                                        chunk = b""
                                    if not chunk:
                                        if proc.poll() is not None:
                                            break
                                        time.sleep(0.02)
                                        continue
                                    # Decode with fallbacks
                                    try:
                                        text = chunk.decode('cp850', errors='ignore')
                                    except Exception:
                                        try:
                                            text = chunk.decode('cp1252', errors='ignore')
                                        except Exception:
                                            text = chunk.decode('utf-8', errors='ignore')
                                    buf += text
                                    # Split lines by CR/LF for logging (some tools use only CR)
                                    while True:
                                        split_idx_n = buf.find('\n')
                                        split_idx_r = buf.find('\r')
                                        split_idx = -1
                                        if split_idx_n != -1 and split_idx_r != -1:
                                            split_idx = min(split_idx_n, split_idx_r)
                                        elif split_idx_n != -1:
                                            split_idx = split_idx_n
                                        elif split_idx_r != -1:
                                            split_idx = split_idx_r
                                        if split_idx == -1:
                                            break
                                        line = buf[:split_idx].strip('\r\n')
                                        buf = buf[split_idx+1:]
                                        lower = line.lower()
                                        if any(pat in lower for pat in success_patterns):
                                            success_flag = True
                                        try:
                                            self.log_message(f"[{tag}] {line}", "informacao")
                                        except Exception:
                                            pass
                                        # Error/open port KO handling
                                        if ("opening port [ko]" in lower or "cannot open the com port" in lower) and not stop_requested:
                                            stop_requested = True
                                            try:
                                                if proc.stdin:
                                                    proc.stdin.write("n\r\n".encode())
                                                    proc.stdin.flush()
                                            except Exception:
                                                pass
                                            try:
                                                self.log_message(f"[{tag}] ERRO: Gravador não conseguiu abrir a porta. Verifique a COM selecionada e se não está em uso.", "erro")
                                            except Exception:
                                                pass
                                        # Prompt responses without blocking
                                        try:
                                            if proc.stdin:
                                                # Se o batch pedir COM(x), envia o número uma única vez
                                                if (("com(x):" in lower) or ("digite o numero da porta serial" in lower) or ("digite o numero da porta" in lower)) and not sent_com and com_number_for_prompt:
                                                    proc.stdin.write((com_number_for_prompt + "\r\n").encode())
                                                    proc.stdin.flush()
                                                    sent_com = True
                                                    try:
                                                        self.log_message(f"[{tag}] >> COM digitada: {com_number_for_prompt}", "informacao")
                                                    except Exception:
                                                        pass
                                                if ("press any key to continue" in lower or "pressione qualquer tecla" in lower) and not sent_press_any:
                                                    proc.stdin.write("\r\n".encode())
                                                    proc.stdin.flush()
                                                    sent_press_any = True
                                                if ("deseja gravar novamente" in lower or "(s/n):" in lower) and not sent_retry_n:
                                                    success_flag = True
                                                    proc.stdin.write("n\r\n".encode())
                                                    proc.stdin.flush()
                                                    sent_retry_n = True
                                                    # Fecha stdin e força término do processo para evitar novo ciclo
                                                    try:
                                                        proc.stdin.close()
                                                    except Exception:
                                                        pass
                                                    # Aguarda até 5s pela saída limpa
                                                    t0 = time.time()
                                                    while proc.poll() is None and (time.time() - t0) < 5.0:
                                                        time.sleep(0.1)
                                                    if proc.poll() is None:
                                                        # Tenta terminar/kill
                                                        try:
                                                            proc.terminate()
                                                        except Exception:
                                                            pass
                                                        time.sleep(0.2)
                                                        if proc.poll() is None:
                                                            try:
                                                                proc.kill()
                                                            except Exception:
                                                                pass
                                        except Exception:
                                            pass
                                    # Also check buffer for prompts without newline
                                    lower_buf = buf.lower()
                                    try:
                                        if proc.stdin:
                                            if (("com(x):" in lower_buf) or ("digite o numero da porta serial" in lower_buf) or ("digite o numero da porta" in lower_buf)) and not sent_com and com_number_for_prompt:
                                                proc.stdin.write((com_number_for_prompt + "\r\n").encode())
                                                proc.stdin.flush()
                                                sent_com = True
                                                try:
                                                    self.log_message(f"[{tag}] >> COM digitada: {com_number_for_prompt}", "informacao")
                                                except Exception:
                                                    pass
                                            if ("press any key to continue" in lower_buf or "pressione qualquer tecla" in lower_buf) and not sent_press_any:
                                                proc.stdin.write("\r\n".encode())
                                                proc.stdin.flush()
                                                sent_press_any = True
                                            if ("deseja gravar novamente" in lower_buf or "(s/n):" in lower_buf) and not sent_retry_n:
                                                success_flag = True
                                                proc.stdin.write("n\r\n".encode())
                                                proc.stdin.flush()
                                                sent_retry_n = True
                                                try:
                                                    proc.stdin.close()
                                                except Exception:
                                                    pass
                                                t0 = time.time()
                                                while proc.poll() is None and (time.time() - t0) < 5.0:
                                                    time.sleep(0.1)
                                                if proc.poll() is None:
                                                    try:
                                                        proc.terminate()
                                                    except Exception:
                                                        pass
                                                    time.sleep(0.2)
                                                    if proc.poll() is None:
                                                        try:
                                                            proc.kill()
                                                        except Exception:
                                                            pass
                                    except Exception:
                                        pass
                                # Flush remaining buffer as one line
                                if buf.strip():
                                    try:
                                        self.log_message(f"[{tag}] {buf.strip()}", "informacao")
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                            finally:
                                try:
                                    proc.wait()
                                except Exception:
                                    pass
                                # Cleanup temporary patched file
                                if cleanup_path:
                                    try:
                                        os.remove(cleanup_path)
                                    except Exception:
                                        pass
                                if after:
                                    after(success_flag)

                        def finish_step(success):
                            def do_finish():
                                # Marca status e avança ao próximo passo
                                if success:
                                    self.log_message(f"APROVADO: '{step['nome']}' (Gravação concluída)", "test_pass")
                                    step_index = self.current_test_steps.index(step) + 1
                                    self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO (Gravação concluída)")
                                    self.passed_steps_count += 1
                                    self._update_list_item_status(step, "APROVADO", QColor("#32CD32"))
                                else:
                                    self.log_message(f"APROVADO: '{step['nome']}' (Gravação finalizada)", "test_pass")
                                    step_index = self.current_test_steps.index(step) + 1
                                    self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO (Processo finalizado)")
                                    self.passed_steps_count += 1
                                    self._update_list_item_status(step, "APROVADO", QColor("#32CD32"))

                                # Reabre porta principal se foi fechada temporariamente
                                try:
                                    if closed_principal and principal_saved_port:
                                        try:
                                            self.serial_command_port_combobox.setCurrentText(principal_saved_port)
                                        except Exception:
                                            pass
                                        self._flash_temporarily_closing_principal = False
                                        try:
                                            self._toggle_serial_connection("serial_command")
                                        except Exception:
                                            pass
                                    else:
                                        self._flash_temporarily_closing_principal = False
                                except Exception:
                                    self._flash_temporarily_closing_principal = False

                                self.current_test_index += 1
                                QTimer.singleShot(100, self._execute_next_test_step)

                            # Garante que toda a lógica rode no thread principal do Qt
                            QTimer.singleShot(0, do_finish)

                        if cmd2:
                            def run_sequence():
                                try:
                                    sel = port_sel_combo.currentText()
                                    m1 = re.search(r"COM(\d+)", sel.upper())
                                    com_num_sel = m1.group(1) if m1 else ""
                                    p1, clean1 = run_cmd_file(cmd_path, send_port=True, com_number=com_num_sel)
                                    def after_p1(_s1):
                                        try:
                                            if self.serial_command_ser and self.serial_command_ser.is_open:
                                                self.serial_command_ser.rts = True
                                                time.sleep(0.2)
                                                self.serial_command_ser.rts = False
                                        except Exception:
                                            pass
                                        try:
                                            p2, clean2 = run_cmd_file(cmd2, send_port=True, com_number=com_num_sel)
                                            threading.Thread(target=lambda: monitor_and_finish(p2, finish_step, tag="CMD2", cleanup_path=clean2, com_number_for_prompt=com_num_sel), daemon=True).start()
                                        except Exception:
                                            finish_step(False)
                                    threading.Thread(target=lambda: monitor_and_finish(p1, after_p1, tag="CMD1", cleanup_path=clean1, com_number_for_prompt=com_num_sel), daemon=True).start()
                                except Exception:
                                    finish_step(False)
                            threading.Thread(target=run_sequence, daemon=True).start()
                        else:
                            # Executa sem bloquear; envia COM uma vez e monitora até terminar para só então aprovar
                            try:
                                sel = port_sel_combo.currentText()
                                m1 = re.search(r"COM(\d+)", sel.upper())
                                com_num_sel = m1.group(1) if m1 else ""
                                p, clean = run_cmd_file(cmd_path, send_port=True, com_number=com_num_sel)
                                threading.Thread(target=lambda: monitor_and_finish(p, finish_step, tag="CMD", cleanup_path=clean, com_number_for_prompt=com_num_sel), daemon=True).start()
                            except Exception:
                                finish_step(False)

                        # Não mostrar popup bloqueante; apenas iniciar e acompanhar no terminal
                    except Exception as e:
                        QMessageBox.critical(dlg, "Erro ao Executar", f"Não foi possível executar o arquivo:\n{e}")
                    finally:
                        dlg.accept()

                def on_cancel():
                    dlg.reject()

                btn_gravar.clicked.connect(on_gravar)
                btn_cancel.clicked.connect(on_cancel)

                dlg.exec()
                # Não avança aqui. O avanço ocorrerá quando a gravação terminar (no monitor).

        else:
            # Caso o tipo de passo seja desconhecido
            error_msg = f"Tipo de passo desconhecido: '{step_type}'"
            self.log_message(f"ERRO: Tipo de passo desconhecido '{step_type}' para '{step['nome']}'. Teste falhou neste passo.", "erro")
            self.test_log_entries.append(f"  ERRO: {error_msg}")
            step_index = self.current_test_steps.index(step) + 1
            self.test_log_entries.append(f"  Status: PASSO {step_index}: REPROVADO")
            self.failed_steps_count += 1
            self._update_list_item_status(step, "REPROVADO", QColor("#FF4500"), error_msg)
            self.current_test_index += 1
            QTimer.singleShot(100, self._execute_next_test_step)

    def _continue_after_wait(self, step_config):
        """
        Chamado após o término do tempo de espera de um passo.
        Aprova o passo e avança para o próximo.
        """
        if not self.test_in_progress:
            return # Não faz nada se o teste foi interrompido

        self.log_message(f"APROVADO: TEMPO DE ESPERA - '{step_config['nome']}' (Tempo de espera concluído)", "test_pass")
        step_index = self.current_test_steps.index(step_config) + 1
        self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO (Tempo de espera concluído)")
        self.passed_steps_count += 1
        self._update_list_item_status(step_config, "APROVADO", QColor("#32CD32"))
        try:
            # Guarda a última duração de espera para cálculo de tolerância percentual
            self.last_wait_duration_s = float(step_config.get("duracao_espera_segundos", 0))
        except Exception:
            pass
        
        self.current_test_index += 1
        QTimer.singleShot(100, self._execute_next_test_step) # Avança para o próximo passo

    def _process_test_response(self, step_config, reader_thread):
        """
        Processa a resposta recebida da porta serial para um passo de teste.
        Realiza a validação da resposta e atualiza o status do passo.
        """
        if not self.test_in_progress:
            return

        # Verifica se a porta ainda está aberta
        if reader_thread.ser is None or not reader_thread.ser.is_open:
            self.log_message(f"AVISO: Porta '{reader_thread.port_name}' não está aberta. Finalizando teste.", "erro")
            self.test_log_entries.append(f"  ERRO: Porta '{reader_thread.port_name}' não está aberta. Teste finalizado inesperadamente.")
            self._finish_test()
            return

        response_data = ""
        if reader_thread:
            response_data = reader_thread.get_buffered_response() # Obtém os dados do buffer da thread
            # Caso venha de porta Modbus mas o conteúdo seja texto (ASCII sobre RS485), converte para string
            if isinstance(response_data, (bytes, bytearray)):
                try:
                    decoded = response_data.decode('utf-8', errors='ignore').strip()
                    # Usa a string decodificada para validação textual
                    response_data = decoded
                except Exception:
                    # Mantém bytes se não decodificar
                    pass
        
        if response_data:
            self.log_message(f"Resposta Coletada para Validação:\n'{response_data}'", "recebido")
            # Garante string no log
            try:
                display_resp = response_data.strip() if isinstance(response_data, str) else response_data
            except Exception:
                display_resp = response_data
            self.test_log_entries.append(f"  Resposta Recebida ({reader_thread.port_name}): '{display_resp}'")
        else:
            self.log_message(f"Nenhuma resposta recebida dentro do timeout para validação.", "informacao")
            self.test_log_entries.append(f"  Resposta Recebida ({reader_thread.port_name}): Nenhuma (Timeout)")

        passed, error_msg = self._validate_response(response_data, step_config) # Valida a resposta
        if passed:
            self.log_message(f"APROVADO: '{step_config.get('nome', 'Passo sem nome')}': Resposta validada com sucesso.", "test_pass")
            step_index = self.current_test_steps.index(step_config) + 1
            self.test_log_entries.append(f"  Status: PASSO {step_index}: APROVADO")
            self.passed_steps_count += 1
            self._update_list_item_status(step_config, "APROVADO", QColor("#32CD32"))
        else:
            self.log_message(f"REPROVADO: '{step_config.get('nome', 'Passo sem nome')}': Resposta não validada. Resposta Recebida:\n'{response_data}'", "test_fail")
            step_index = self.current_test_steps.index(step_config) + 1
            self.test_log_entries.append(f"  Status: PASSO {step_index}: REPROVADO")
            if error_msg:
                self.test_log_entries.append(f"  Detalhe do Erro: {error_msg}")
            self.failed_steps_count += 1
            self._update_list_item_status(step_config, "REPROVADO", QColor("#FF4500"), error_msg)            
        self.current_test_index += 1
        QTimer.singleShot(100, self._execute_next_test_step) # Avança para o próximo passo

    def _update_list_item_status(self, step, status_text, color, error_detail=""):
        """
        Atualiza o texto e a cor de um item na QListWidget de progresso do teste.
        """
        # Verifica se o passo tem um 'list_item' associado (ou seja, não foi pulado pelo modo fast)
        if 'list_item' in step and step['list_item'] is not None:
            item = step['list_item']
            # O item_index aqui deve ser o índice do passo na lista original (current_test_steps)
            # para fins de numeração do passo no texto, não o índice na lista de progresso filtrada.
            item_index = self.current_test_steps.index(step) 
            display_text = f"Passo {item_index + 1}: {step.get('nome', 'Sem Nome')} - {status_text}"
            
            step['status'] = status_text 
            step['last_error_detail'] = error_detail

            if status_text == "REPROVADO" and error_detail:
                display_text += f" (Erro: {error_detail})"
            item.setText(display_text)
            item.setForeground(QBrush(color))
            self.test_progress_list.scrollToItem(item) # Rola para o item atual

    def _handle_test_progress_item_click(self, item: QListWidgetItem):
        """
        Lida com o clique em um item da lista de progresso do teste.
        Permite re-testar um passo que falhou.
        """
        if self.test_in_progress:
            return # Não permite interação se um teste já estiver em andamento

        # Encontra o passo correspondente no self.current_test_steps
        clicked_step = None
        for step in self.current_test_steps:
            if 'list_item' in step and step['list_item'] == item:
                clicked_step = step
                break

        if clicked_step and clicked_step.get('status') == "REPROVADO":
            step_name = clicked_step.get('nome', 'Passo sem nome')
            error_description = clicked_step.get('last_error_detail', 'Descrição de erro não disponível.')
            
            dialog = RetestStepDialog(step_name, error_description, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                # Encontra o índice do passo no self.current_test_steps
                failed_step_index = self.current_test_steps.index(clicked_step)
                self._retest_failed_step(failed_step_index) # Inicia o re-teste
            else:
                self.log_message(f"Teste finalizado pelo usuário após passo reprovado {self.current_test_steps.index(clicked_step) + 1}.", "sistema")
                self._stop_test() # O usuário escolheu finalizar o teste

    def _retest_failed_step(self, failed_step_index):
        """
        Prepara e inicia o re-teste de um passo específico que falhou.
        """
        if self.test_in_progress:
            QMessageBox.warning(self, "Teste em Andamento", "Não é possível re-testar um passo enquanto outro teste está em execução.")
            return

        self.log_message(f"Re-testando passo {failed_step_index + 1}: '{self.current_test_steps[failed_step_index]['nome']}'", "sistema")
        
        step_to_retest = self.current_test_steps[failed_step_index]

        # Ajusta contadores de passos
        if step_to_retest.get('status') == 'APROVADO':
            self.passed_steps_count -= 1
        elif step_to_retest.get('status') == 'REPROVADO':
            self.failed_steps_count -= 1
        
        step_to_retest['status'] = "Pendente" # Reseta o status do passo
        step_to_retest['last_error_detail'] = ""

        self.current_test_index = failed_step_index # Define o índice para o passo a ser re-testado
        
        self.test_in_progress = True
        self.start_test_button.setEnabled(False)
        self.stop_test_button.setEnabled(True)
        self.load_test_button.setEnabled(False)

        # Desabilita controles de envio manual/automático durante o re-teste
        self.direct_command_input.setEnabled(False)
        self.direct_send_button.setEnabled(False)
        for i in range(len(self.send_command_inputs)):
            self.send_command_inputs[i].setEnabled(False)
            self.send_buttons[i].setEnabled(False)
            self.auto_send_checkboxes[i].setEnabled(False)
            self.auto_send_checkboxes[i].setChecked(False)
            self.auto_send_config_buttons[i].setEnabled(False)
            self.auto_send_timers[i].stop()

        if self.serial_command_reader_thread:
            self.serial_command_reader_thread.clear_response_buffer_for_next_step()
        if self.modbus_serial_reader_thread:
            self.modbus_serial_reader_thread.clear_response_buffer_for_next_step()

        self.test_progress_label.setText(f"Status Geral: Re-testando passo {failed_step_index + 1}...")
        
        self._execute_next_test_step() # Inicia a execução do passo re-testado

    def _validate_response(self, response, step_config):
        """
        Valida a resposta recebida com base no tipo de validação configurado para o passo.
        Retorna True/False para aprovação/reprovação e uma mensagem de erro, se houver.
        """
        tipo_validacao = step_config.get("tipo_validacao")
        param_validacao = step_config.get("param_validacao")

        normalized_response = response.strip() # Remove espaços em branco do início/fim

        if tipo_validacao == "string_exata":
            expected_string = str(param_validacao).strip()
            if normalized_response == expected_string:
                return True, ""
            else:
                return False, f"Resposta esperada: '{expected_string}', recebida: '{normalized_response}'"
        
        elif tipo_validacao == "numerico_faixa":
            try:
                # Tenta encontrar um número na resposta
                match = re.search(r"[-+]?\d*\.?\d+", normalized_response)
                if not match:
                    return False, f"Nenhum número encontrado na resposta: '{normalized_response}'"
                
                num_response = float(match.group(0)) # Converte o número encontrado para float

                if param_validacao and "min" in param_validacao and "max" in param_validacao:
                    min_val = param_validacao["min"]
                    max_val = param_validacao["max"]
                    if min_val <= num_response <= max_val:
                        return True, ""
                    else:
                        return False, f"Valor '{num_response}' fora da faixa esperada [{min_val}, {max_val}]"
                else:
                    return False, "Parâmetros min/max ausentes ou inválidos para validação numérica."
            except ValueError:
                return False, f"Não foi possível converter o valor extraído '{match.group(0) if match else 'N/A'}' para número."
        
        elif tipo_validacao == "texto_numerico_simples":
            if not param_validacao or "regex" not in param_validacao or "min" not in param_validacao or "max" not in param_validacao:
                return False, "Parâmetros de validação incompletos ou ausentes para 'Texto Simples com Número'."
            
            regex_pattern = param_validacao["regex"]
            min_val = param_validacao["min"]
            max_val = param_validacao["max"]

            extracted_value_str = ""
            try:
                match = re.search(regex_pattern, normalized_response)
                if not match or len(match.groups()) == 0:
                    return False, f"Padrão de Texto '{regex_pattern}' não encontrou correspondência ou grupo de captura na resposta."
                
                extracted_value_str = match.group(1) # Captura o primeiro grupo (o número)
                num_response = float(extracted_value_str)
                
                if min_val <= num_response <= max_val:
                    return True, ""
                else:
                    return False, f"Valor '{num_response}' fora da faixa esperada [{min_val}, {max_val}] com Padrão de Texto."
            except re.error as e:
                return False, f"Expressão Regular inválida: {e}"
            except ValueError:
                return False, f"Não foi possível converter o valor extraído '{extracted_value_str}' para número."
            except IndexError:
                return False, f"Padrão de Texto '{regex_pattern}' não possui grupo de captura. Verifique se o regex possui parênteses para capturar o valor."
        
        elif tipo_validacao == "texto_numerico_multiplos":
            if not param_validacao or "regex" not in param_validacao or "min" not in param_validacao or "max" not in param_validacao:
                return False, "Parâmetros de validação incompletos ou ausentes para 'Texto com Vários Números'."

            regex_pattern = param_validacao["regex"]
            min_val = param_validacao["min"]
            max_val = param_validacao["max"]
            expected_count = int(param_validacao.get("expected_count", 0))

            try:
                m = re.search(regex_pattern, normalized_response)
                if not m:
                    return False, f"Padrão de Texto '{regex_pattern}' não encontrou correspondência na resposta."
                groups = m.groups()
                if expected_count and len(groups) != expected_count:
                    return False, f"Foram capturados {len(groups)} valores, mas eram esperados {expected_count}."

                for idx, g in enumerate(groups, start=1):
                    try:
                        v = float(g)
                    except ValueError:
                        return False, f"Não foi possível converter o valor '{g}' (posição {idx}) para número."
                    if not (min_val <= v <= max_val):
                        return False, f"Valor '{v}' na posição {idx} fora da faixa esperada [{min_val}, {max_val}]."
                return True, ""
            except re.error as e:
                return False, f"Expressão Regular inválida: {e}"
        
        elif tipo_validacao == "datetime_20s":
            # Valida com tolerância percentual baseada no tempo decorrido ou última espera
            try:
                # Remove aspas simples/dobras envolventes e espaços extras
                resp = normalized_response.strip().strip("'").strip('"').strip()
                # Procura qualquer ocorrência de YYYY-MM-DD HH:MM:SS
                match = re.search(r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})", resp)
                if not match:
                    return False, "Resposta não contém data/hora no formato 'YYYY-MM-DD HH:MM:SS'."
                device_dt_str = match.group(1)
                device_dt = datetime.strptime(device_dt_str, "%Y-%m-%d %H:%M:%S")
                now_dt = datetime.now()
                delta_s = abs((now_dt - device_dt).total_seconds())

                # Calcula tolerância: 10% do tempo decorrido desde o último SET_RTC se existir,
                # senão 10% da última espera. Define pisos/tetos razoáveis.
                tol_s = 20.0
                try:
                    if hasattr(self, 'last_rtc_set_time') and self.last_rtc_set_time:
                        elapsed = max(0.0, (now_dt - self.last_rtc_set_time).total_seconds())
                        tol_s = 0.10 * elapsed
                    elif hasattr(self, 'last_wait_duration_s') and self.last_wait_duration_s:
                        tol_s = 0.10 * float(self.last_wait_duration_s)
                except Exception:
                    pass
                # Aplica limites mínimos/máximos para evitar tolerância irrisória ou exagerada
                tol_s = max(2.0, min(tol_s, 60.0))

                if delta_s <= tol_s:
                    return True, ""
                return False, f"Data/Hora fora da tolerância (diferença {int(delta_s)}s > {int(tol_s)}s)."
            except Exception as e:
                return False, f"Erro ao validar Data/Hora: {e}"
        
        elif tipo_validacao == "serial_settings_match":
            # Compara a resposta com o last_serial_number do settings.json
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings_data = json.load(f)
                expected_serial = str(settings_data.get("last_serial_number", "")).strip()
                if not expected_serial:
                    return False, "last_serial_number ausente no settings.json"
                # Normaliza a resposta e verifica igualdade ou ocorrência na resposta
                resp = normalized_response
                if resp == expected_serial or expected_serial in resp:
                    return True, ""
                # Tenta extrair padrão de série (ex.: 16586/3) da resposta
                m = re.search(r"\b(\d+[/-]\d+)\b", resp)
                if m and m.group(1) == expected_serial:
                    return True, ""
                return False, f"Número de série esperado '{expected_serial}', recebido: '{resp}'"
            except Exception as e:
                return False, f"Erro ao ler settings.json: {e}"
        
        elif tipo_validacao == "nenhuma":
            return True, "" # Sempre passa se não houver validação

        return False, f"Tipo de validação desconhecido: '{tipo_validacao}'"

    def _finish_test(self):
        """
        Finaliza o teste, exibe os resultados, gera o arquivo de log e redefine os controles.
        """
        self.test_in_progress = False
        self.test_timer.stop()
        self.log_message("TESTE CONCLUÍDO", "sistema")
        total_steps = len(self.current_test_steps)
        self.log_message(f"Total de Passos: {total_steps}", "sistema")
        self.log_message(f"Passos Aprovados: {self.passed_steps_count}", "sistema")
        self.log_message(f"Passos Reprovados: {self.failed_steps_count}", "sistema")

        final_status_text = ""
        passos_validos = [s for s in self.current_test_steps if s.get("status") not in ["PULADO", "INTERROMPIDO"]]
        aprovados_reais = sum(1 for s in passos_validos if s.get("status") == "APROVADO")
        all_approved = all(s.get("status") == "APROVADO" for s in passos_validos)

        if all_approved:
            QMessageBox.information(self, "Teste Concluído", f"Todos os {aprovados_reais} passos foram aprovados!")
            self.test_status_label.setText("Status do Teste: APROVADO!")
            final_status_text = "APROVADO!"
            self.test_progress_label.setStyleSheet(f"font-weight: bold; padding: 5px; color: {self.LOG_COLORS['test_pass']};")
            self.test_log_entries.append(f"\n--- STATUS FINAL: APROVADO ({aprovados_reais}/{len(passos_validos)} Passos Aprovados) ---")
        else:
            QMessageBox.warning(self, "Teste Concluído", f"O teste falhou em {len(passos_validos) - aprovados_reais} passo(s).")
            self.test_status_label.setText("Status do Teste: REPROVADO!")
            final_status_text = "REPROVADO!"
            self.test_progress_label.setStyleSheet(f"font-weight: bold; padding: 5px; color: {self.LOG_COLORS['test_fail']};")
            self.test_log_entries.append(f"\n--- STATUS FINAL: REPROVADO ({len(passos_validos) - aprovados_reais}/{len(passos_validos)} Passos Reprovados) ---")
        
        self.test_progress_label.setText(f"Status Geral: Teste {final_status_text} ({self.passed_steps_count}/{total_steps} Passos Aprovados)")
        self.test_log_entries.append(f"Data/Hora Término: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        tempo_total = int((datetime.now() - self.test_start_time).total_seconds())
        houve_erro = any(s.get("status") == "REPROVADO" for s in self.current_test_steps)
        self.oled.finalizar_teste(duracao_segundos=tempo_total, houve_erro=houve_erro)

        self._generate_test_log_file() # Gera o arquivo de log do teste
        self._reset_test_controls() # Redefine os controles da UI

    def _generate_test_log_file(self):
        """
        Gera um arquivo de log detalhado do teste na pasta selecionada.
        Se não for possível, salva em ~/Documents/Logs Do Teste.
        Quando a pasta selecionada voltar, move os logs pendentes.
        """
        if not self.current_pr_number or not self.current_serial_number:
            self.log_message("Não é possível gerar log: Número do PR ou de Série ausente.", "erro")
            return

        # Caminho principal (selecionado pelo usuário)
        documents_path = self.configuracoes_tab.get_log_path()
        base_log_dir = os.path.join(documents_path, "Logs Do Teste")
        pr_folder_name = f"PR{self.current_pr_number}"
        pr_log_dir = os.path.join(base_log_dir, pr_folder_name)

        # Caminho alternativo (fallback)
        fallback_base = os.path.join(os.path.expanduser("~"), "Documents", "Logs Do Teste")
        fallback_pr_dir = os.path.join(fallback_base, pr_folder_name)

        clean_serial_number = self.current_serial_number.replace('/', '-')
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        file_name = f"PR{self.current_pr_number}_{clean_serial_number}_{timestamp}.txt"
        log_content = "\n".join(self.test_log_entries)

        # Tenta salvar na pasta selecionada
        try:
            os.makedirs(pr_log_dir, exist_ok=True)
            full_file_path = os.path.join(pr_log_dir, file_name)
            with open(full_file_path, 'w', encoding='utf-8') as f:
                f.write(log_content)
            self.log_message(f"Log do teste salvo como: '{full_file_path}'", "sistema")
            # Após salvar, tenta mover logs pendentes do fallback
            self._move_pending_logs_to_selected_folder()
        except Exception as e:
            self.log_message(f"Erro ao salvar o log na pasta selecionada: {e}", "erro")
            # Salva no fallback
            try:
                os.makedirs(fallback_pr_dir, exist_ok=True)
                fallback_file_path = os.path.join(fallback_pr_dir, file_name)
                with open(fallback_file_path, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                self.log_message(f"Log do teste salvo em fallback: '{fallback_file_path}'", "sistema")
            except Exception as e2:
                self.log_message(f"Erro ao salvar o log no fallback: {e2}", "erro")
                QMessageBox.critical(self, "Erro ao Salvar Log", f"Não foi possível salvar o arquivo de log:\n{e2}")

    def _move_pending_logs_to_selected_folder(self):
        import shutil
        fallback_base = os.path.join(os.path.expanduser("~"), "Documents", "Logs Do Teste")
        documents_path = self.configuracoes_tab.get_log_path()
        if documents_path.strip().endswith("Logs Do Teste"):
            base_log_dir = documents_path
        else:
            base_log_dir = os.path.join(documents_path, "Logs Do Teste")

        # Só move se os caminhos forem diferentes!
        if os.path.abspath(fallback_base) == os.path.abspath(base_log_dir):
            return

        if not os.path.exists(fallback_base):
            return

        arquivos = []
        for root, dirs, files in os.walk(fallback_base):
            for file in files:
                arquivos.append(os.path.join(root, file))
        if not arquivos:
            return

        for root, dirs, files in os.walk(fallback_base):
            for file in files:
                src_file = os.path.join(root, file)
                rel_path = os.path.relpath(root, fallback_base)
                dest_dir = os.path.join(base_log_dir, rel_path)
                os.makedirs(dest_dir, exist_ok=True)
                dest_file = os.path.join(dest_dir, file)
                try:
                    shutil.move(src_file, dest_file)
                    self.log_message(f"Log pendente movido para pasta selecionada: {dest_file}", "sistema")
                except Exception as e:
                    self.log_message(f"Erro ao mover log pendente: {e}", "erro")
        for root, dirs, files in os.walk(fallback_base, topdown=False):
            if not os.listdir(root):
                os.rmdir(root)
            

    def _reset_test_controls(self):
        """
        Redefine o estado dos controles da UI após a conclusão ou interrupção de um teste.
        """
        self._update_start_test_button_state()
        self.stop_test_button.setEnabled(False)
        self.load_test_button.setEnabled(True)
        
        is_serial_command_connected = (self.serial_command_ser is not None and self.serial_command_ser.is_open)
        self.direct_command_input.setEnabled(is_serial_command_connected)
        self.direct_send_button.setEnabled(is_serial_command_connected)
        if is_serial_command_connected:
            self.direct_command_input.setFocus()

        for i in range(len(self.send_command_inputs)):
            self.send_command_inputs[i].setEnabled(is_serial_command_connected)
            self.send_buttons[i].setEnabled(is_serial_command_connected)
            self.auto_send_checkboxes[i].setEnabled(is_serial_command_connected)
            self.auto_send_config_buttons[i].setEnabled(is_serial_command_connected)
            
            if not is_serial_command_connected:
                self.auto_send_checkboxes[i].setChecked(False)
            elif self.auto_send_checkboxes[i].isChecked() and not self.auto_send_timers[i].isActive():
                self._toggle_auto_send_timer(Qt.CheckState.Checked.value, i) # Reinicia auto-envio se estava ativo

        self.current_test_index = -1 # Reseta o índice do teste
        
        self._update_port_config_visibility(False) # Recolhe as configurações de porta na aba terminal
        #self.fast_mode_active = False # Garante que o modo fast seja desativado ao final do teste
        self._update_fast_mode_status_label() # Atualiza o rótulo do modo fast

    def _save_default_settings(self):
        """
        Salva um novo conjunto de configurações padrão no arquivo 'settings.json'.
        Usado para inicializar ou corrigir um arquivo de configurações corrompido.
        """
        default_settings = {
            "serial_command_port": "",
            "modbus_port": "",
            "serial_baud": "115200",
            "serial_data_bits": "8",
            "serial_parity": "Nenhuma",
            "serial_handshake": "Nenhum",
            "serial_mode": "Free",
            "serial_dtr": False, # Adicionado DTR padrão
            "serial_rts": False, # Adicionado RTS padrão
            "modbus_baud": "9600",
            "modbus_data_bits": "8",
            "modbus_parity": "Nenhuma",
            "modbus_handshake": "Nenhum",
            "modbus_mode": "Free",
            "modbus_dtr": False, # Adicionado DTR padrão
            "modbus_rts": False, # Adicionado RTS padrão
            "auto_send_lines": [{"command": "", "auto_send_checked": False, "interval_seconds": 1.0}] * 4,
            "last_pr_number": "",
            "last_serial_number": ""
        }
        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(default_settings, f, indent=4, ensure_ascii=False)
            self.log_message("Arquivo de configurações padrão salvo para corrigir erro.", "sistema")
        except Exception as e:
            self.log_message(f"Erro ao salvar arquivo de configurações padrão: {e}", "erro")

    def _save_settings(self):
        """
        Salva as configurações atuais da UI (portas, baud rates, comandos de auto-envio, etc.)
        no arquivo 'settings.json'.
        """
        settings = {
            "serial_command_port": self.serial_command_port_combobox.currentText(),
            "modbus_port": self.modbus_port_combobox.currentText(),
            "serial_baud": self.serial_baud_combo.currentText(),
            "serial_data_bits": self.serial_data_bits_combo.currentText(),
            "serial_parity": self.serial_parity_combo.currentText(),
            "serial_handshake": self.serial_handshake_combo.currentText(),
            "serial_mode": self.serial_mode_combo.currentText(),
            "serial_dtr": self.serial_dtr_checkbox.isChecked(), # Salva estado do DTR
            "serial_rts": self.serial_rts_checkbox.isChecked(), # Salva estado do RTS
            "modbus_baud": self.modbus_baud_combo.currentText(),
            "modbus_data_bits": self.modbus_data_bits_combo.currentText(),
            "modbus_parity": self.modbus_parity_combo.currentText(),
            "modbus_handshake": self.modbus_handshake_combo.currentText(),
            "modbus_mode": self.modbus_mode_combo.currentText(),
            "modbus_dtr": self.modbus_dtr_checkbox.isChecked(), # Salva estado do DTR Modbus
            "modbus_rts": self.modbus_rts_checkbox.isChecked(), # Salva estado do RTS Modbus
            "auto_send_lines": [],
            "last_pr_number": self.current_pr_number,
            "last_serial_number": self.current_serial_number
        }
        for i in range(len(self.send_command_inputs)):
            line_config = {
                "command": self.send_command_inputs[i].text(),
                "auto_send_checked": self.auto_send_checkboxes[i].isChecked(),
                "interval_seconds": self.auto_send_intervals_s[i]
            }
            settings["auto_send_lines"].append(line_config)

        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.log_message(f"Erro ao salvar configurações: {e}", "erro")

    def _load_settings(self):
        """
        Carrega as configurações salvas do arquivo 'settings.json' e as aplica à UI.
        Em caso de erro ou arquivo não encontrado, salva as configurações padrão.
        """
        if self.test_creator_tab_index != -1: 
            self.tab_widget.setTabVisible(self.test_creator_tab_index, False) # Garante que a aba do criador esteja oculta inicialmente

        settings = {}
        if not os.path.exists(self.settings_file):
            self.log_message("Arquivo de configurações não encontrado. Aplicando tema padrão e criando novo arquivo.", "informacao")
            self._save_default_settings()
            self._apply_stylesheet() # Reaplica o stylesheet para garantir que o tema padrão seja carregado
            return

        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                settings = json.load(f)
            
            # Aplica o stylesheet novamente para garantir que as cores de log sejam atualizadas
            self._apply_stylesheet() 

            # Carrega configurações da porta principal
            port = settings.get("serial_command_port")
            # A lógica de carregamento de portas foi movida para _update_available_ports
            # para evitar duplicidade e garantir que a lista esteja sempre atualizada.
            # Apenas definimos o texto atual, e _update_available_ports garantirá que a lista seja populada.
            if port:
                self.serial_command_port_combobox.setCurrentText(port)

            # Carrega configurações da porta Modbus
            modbus_port = settings.get("modbus_port")
            if modbus_port:
                self.modbus_port_combobox.setCurrentText(modbus_port)

            self.serial_baud_combo.setCurrentText(settings.get("serial_baud", "115200"))
            self.serial_data_bits_combo.setCurrentText(settings.get("serial_data_bits", "8"))
            self.serial_parity_combo.setCurrentText(settings.get("serial_parity", "Nenhuma"))
            self.serial_handshake_combo.setCurrentText(settings.get("serial_handshake", "Nenhum"))
            self.serial_mode_combo.setCurrentText(settings.get("serial_mode", "Free"))
            self.serial_dtr_checkbox.setChecked(settings.get("serial_dtr", False)) # Carrega estado do DTR
            self.serial_rts_checkbox.setChecked(settings.get("serial_rts", False)) # Carrega estado do RTS

            self.modbus_baud_combo.setCurrentText(settings.get("modbus_baud", "9600"))
            self.modbus_data_bits_combo.setCurrentText(settings.get("modbus_data_bits", "8"))
            self.modbus_parity_combo.setCurrentText(settings.get("modbus_parity", "Nenhuma"))
            self.modbus_handshake_combo.setCurrentText(settings.get("modbus_handshake", "Nenhum"))
            self.modbus_mode_combo.setCurrentText(settings.get("modbus_mode", "Free"))
            self.modbus_dtr_checkbox.setChecked(settings.get("modbus_dtr", False)) # Carrega estado do DTR Modbus
            self.modbus_rts_checkbox.setChecked(settings.get("modbus_rts", False)) # Carrega estado do RTS Modbus

            self.current_pr_number = settings.get("last_pr_number", "")
            self.current_serial_number = settings.get("last_serial_number", "")

            # Carrega configurações das linhas de auto-envio
            loaded_lines = settings.get("auto_send_lines", [])
            for i in range(min(len(loaded_lines), len(self.send_command_inputs))):
                line_config = loaded_lines[i]
                self.send_command_inputs[i].setText(line_config.get("command", ""))
                
                interval = line_config.get("interval_seconds", 1.0)
                if interval <= 0:
                    interval = 1.0
                self.auto_send_intervals_s[i] = interval
                
                self.auto_send_checkboxes[i].setChecked(line_config.get("auto_send_checked", False))
            
        except (json.JSONDecodeError, Exception) as e:
            self.log_message(f"Erro ao carregar configurações do arquivo: {e}. Revertendo para configurações padrão.", "erro")
            self._save_default_settings() # Salva configurações padrão em caso de erro
            self._apply_stylesheet() # Reaplica o stylesheet para garantir que o tema padrão seja carregado

    def _toggle_port_config_details(self, port_type, force_hide=False):
        """
        Alterna a visibilidade dos detalhes de configuração de uma porta serial
        (baud rate, data bits, etc.).
        """
        if port_type == "serial_command":
            details_widget = self.serial_details_widget
            toggle_button = self.serial_config_toggle_button
        elif port_type == "modbus":
            details_widget = self.modbus_details_widget
            toggle_button = self.modbus_config_toggle_button
        else:
            return

        if force_hide:
            details_widget.setVisible(False)
            toggle_button.setArrowType(Qt.ArrowType.RightArrow)
            toggle_button.setChecked(False)
        else:
            is_visible = details_widget.isVisible()
            details_widget.setVisible(not is_visible)
            if details_widget.isVisible():
                toggle_button.setArrowType(Qt.ArrowType.DownArrow)
            else:
                toggle_button.setArrowType(Qt.ArrowType.RightArrow)

    def _update_port_config_visibility(self, show_full_config):
        """
        Controla a visibilidade dos campos de configuração detalhada das portas
        na aba "Terminal".
        """
        self.serial_details_widget.setVisible(show_full_config)
        if show_full_config:
            self.serial_config_toggle_button.setArrowType(Qt.ArrowType.DownArrow)
            self.serial_config_toggle_button.setChecked(True)
        else:
            self.serial_config_toggle_button.setArrowType(Qt.ArrowType.RightArrow)
            self.serial_config_toggle_button.setChecked(False)

        if self.modbus_serial_group.isVisible(): # Só afeta o Modbus se ele já estiver visível
            self.modbus_details_widget.setVisible(show_full_config)
            if show_full_config:
                self.modbus_config_toggle_button.setArrowType(Qt.ArrowType.DownArrow)
                self.modbus_config_toggle_button.setChecked(True)
            else:
                self.modbus_config_toggle_button.setArrowType(Qt.ArrowType.RightArrow)
                self.modbus_config_toggle_button.setChecked(False)

    def _on_send_target_changed(self, target_name: str):
        """Quando o seletor de envio muda, exibe a configuração da porta correspondente.
        Se 'Modbus' for selecionado, garante que o grupo Modbus apareça e fique expandido
        para permitir a conexão da porta.
        """
        try:
            if target_name == "Modbus":
                # Mostra e expande a configuração Modbus para facilitar a conexão
                self.modbus_serial_group.setVisible(True)
                self.modbus_details_widget.setVisible(True)
                self.modbus_config_toggle_button.setArrowType(Qt.ArrowType.DownArrow)
                self.modbus_config_toggle_button.setChecked(True)
                self.connect_modbus_button.setEnabled(True)
                self.modbus_port_combobox.setEnabled(True)
                # Atualiza a lista de portas para garantir que o usuário veja as COM disponíveis
                self._list_serial_ports()
            else:
                # Nada a fazer especial ao voltar para Principal
                pass
        except Exception as e:
            self.log_message(f"Erro ao atualizar UI para alvo de envio '{target_name}': {e}", "erro")

    def closeEvent(self, event):
        """
        Sobrescreve o evento de fechamento da janela para salvar as configurações
        e fechar as portas seriais de forma segura.
        """
        self._save_settings()
        
         # Pare o timer quando a aplicação for fechada
        if self.port_monitor_timer.isActive():
            self.port_monitor_timer.stop()
        # Para e limpa as threads de leitura
        if self.serial_command_reader_thread:
            self.serial_command_reader_thread.stop()
            self.serial_command_reader_thread = None
        
        if self.modbus_serial_reader_thread:
            self.modbus_serial_reader_thread.stop()
            self.modbus_serial_reader_thread = None

        # Para todos os timers de auto-envio
        for timer in self.auto_send_timers:
            if timer.isActive():
                timer.stop()
        
        # Fecha as portas seriais abertas
        if self.serial_command_ser and self.serial_command_ser.is_open:
            try:
                self.serial_command_ser.close()
            except Exception as e:
                self.log_message(f"Erro ao fechar a porta serial principal em closeEvent: {e}", "erro")
        
        if self.modbus_ser and self.modbus_ser.is_open:
            try:
                self.modbus_ser.close()
            except Exception as e:
                self.log_message(f"Erro ao fechar a porta Modbus em closeEvent: {e}", "erro")
        
        event.accept() # Aceita o evento de fechamento da janela


if __name__ == "__main__":
    # Ponto de entrada da aplicação
    try:
        app = QApplication(sys.argv)
        # Caminho absoluto para o ícone
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "assets", "icone.ico")
        app.setWindowIcon(QIcon(icon_path))
        window = PlacaTesterApp()
        window.show() # Exibe a janela principal
        sys.exit(app.exec()) # Inicia o loop de eventos da aplicação
    except Exception as e:
        print(f"Erro fatal ao iniciar a aplicação: {e}")
        sys.exit(1) # Sai com código de erro
