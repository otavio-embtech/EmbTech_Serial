import os
import time
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime

import pandas as pd
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator
from PyQt6.QtCore import QDate, QEvent, QObject, QSignalBlocker, QThread, QTimer, Qt, pyqtSignal, pyqtSlot
from PyQt6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)


class DashboardFigureCanvas(FigureCanvas):
    def __init__(self, figure):
        super().__init__(figure)
        self._scroll_area = None
        self._hover_suppression_callback = None

    def set_scroll_area(self, scroll_area):
        self._scroll_area = scroll_area

    def set_hover_suppression_callback(self, callback):
        self._hover_suppression_callback = callback

    def wheelEvent(self, event):
        if callable(self._hover_suppression_callback):
            try:
                self._hover_suppression_callback()
            except Exception:
                pass
        if self._scroll_area is not None:
            vertical_bar = self._scroll_area.verticalScrollBar()
            horizontal_bar = self._scroll_area.horizontalScrollBar()
            delta_y = event.angleDelta().y()
            delta_x = event.angleDelta().x()

            if delta_y:
                step = max(vertical_bar.singleStep(), 20)
                vertical_bar.setValue(vertical_bar.value() - int((delta_y / 120) * step * 3))
                event.accept()
                return

            if delta_x:
                step = max(horizontal_bar.singleStep(), 20)
                horizontal_bar.setValue(horizontal_bar.value() - int((delta_x / 120) * step * 3))
                event.accept()
                return

        super().wheelEvent(event)


class LogAnalysisWorker(QObject):
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(object, object, int, int, str)
    failed = pyqtSignal(str)
    canceled = pyqtSignal()

    _RE_FINAL_STATUS = re.compile(r"--- STATUS FINAL: (APROVADO|REPROVADO)", re.IGNORECASE)
    _RE_STEP_STARTED = re.compile(r"PASSO\s+(\d+):\s*(.*?)\s*-\s*Em Exec", re.IGNORECASE)
    _RE_STEP_FAILED = re.compile(r"Status:\s*PASSO\s+(\d+):\s*REPROVADO", re.IGNORECASE)
    _RE_STEP_APPROVED = re.compile(r"Status:\s*PASSO\s+(\d+):\s*APROVADO", re.IGNORECASE)
    _RE_STEP_FAST_SKIPPED = re.compile(r"PASSO\s+\d+:.*?Status:\s*PULADO\s+\(Modo Fast\)", re.IGNORECASE)
    _RE_AUTO_RETRY = re.compile(r"RETENTATIVA AUTOM", re.IGNORECASE)
    _RE_TIMEOUT = re.compile(r"Nenhuma \(Timeout\)|Nenhuma resposta recebida", re.IGNORECASE)
    _RE_ERROR_DETAIL = re.compile(r"Detalhe do Erro:\s*(.+)", re.IGNORECASE)
    _RE_COMMAND_SENT = re.compile(r"Comando Enviado", re.IGNORECASE)
    _RE_MODBUS_NO_RESPONSE = re.compile(r"Modbus.*Nenhuma resposta recebida", re.IGNORECASE)

    def __init__(self, root_dir):
        super().__init__()
        self.root_dir = root_dir
        self._cancel_requested = False

    def request_cancel(self):
        self._cancel_requested = True

    @pyqtSlot()
    def run(self):
        try:
            arquivos = self._collect_log_files(self.root_dir)
            total_arquivos = len(arquivos)
            logs_por_placa = {}

            for index, caminho in enumerate(arquivos, start=1):
                if self._cancel_requested:
                    self.canceled.emit()
                    return

                registro = self._parse_log_file(caminho)
                if registro is not None:
                    chave = (registro["pr"], registro["numero_serie"])
                    existente = logs_por_placa.get(chave)
                    if existente is None or registro["fim"] >= existente["fim"]:
                        logs_por_placa[chave] = registro

                self.progress.emit(index, total_arquivos, os.path.basename(caminho))

            resultados = sorted(
                logs_por_placa.values(),
                key=lambda item: (item["fim"], item["inicio"]),
                reverse=True,
            )
            operadores = sorted({item["operador"] for item in resultados if item.get("operador")})
            self.finished.emit(resultados, operadores, total_arquivos, len(resultados), self.root_dir)
        except Exception as exc:
            self.failed.emit(str(exc))

    def _collect_log_files(self, root_dir):
        arquivos = []
        for current_root, _dirs, files in os.walk(root_dir):
            for nome in files:
                if nome.lower().endswith(".txt"):
                    arquivos.append(os.path.join(current_root, nome))
        arquivos.sort()
        return arquivos

    def _parse_log_file(self, caminho):
        conteudo = self._read_text(caminho)
        campos = self._extract_fields(conteudo)
        if not campos:
            return None

        pr_doc = campos.get("pr")
        numero_serie = campos.get("numero_serie")
        operador = campos.get("operador")
        data_inicio = campos.get("inicio")
        data_fim = campos.get("fim")

        if not pr_doc or not numero_serie or not operador or not data_inicio or not data_fim:
            return None

        try:
            dt_inicio = datetime.strptime(data_inicio, "%Y-%m-%d %H:%M:%S")
            dt_fim = datetime.strptime(data_fim, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None

        status_matches = self._RE_FINAL_STATUS.findall(conteudo)
        if not status_matches:
            return None
        resultado = status_matches[-1].upper()

        duracao_segundos = max((dt_fim - dt_inicio).total_seconds(), 0.0)
        passos_reprovados = self._extract_failed_steps(conteudo)
        detalhes_erro = self._extract_error_details(conteudo)
        erro_principal = self._summarize_primary_error(detalhes_erro)
        pr_formatado = pr_doc.strip()
        if pr_formatado.isdigit():
            pr_formatado = f"PR{pr_formatado}"

        return {
            "pr": pr_formatado,
            "numero_serie": numero_serie.strip(),
            "operador": operador.strip(),
            "inicio": dt_inicio,
            "fim": dt_fim,
            "duracao_segundos": duracao_segundos,
            "resultado": resultado,
            "maquina": (campos.get("maquina") or "").strip(),
            "passos_reprovados": passos_reprovados,
            "passos_aprovados": len(self._RE_STEP_APPROVED.findall(conteudo)),
            "passos_pulados_fast": len(self._RE_STEP_FAST_SKIPPED.findall(conteudo)),
            "retentativas_automaticas": len(self._RE_AUTO_RETRY.findall(conteudo)),
            "timeouts": len(self._RE_TIMEOUT.findall(conteudo)),
            "comandos_enviados": len(self._RE_COMMAND_SENT.findall(conteudo)),
            "falhas_modbus_sem_resposta": len(self._RE_MODBUS_NO_RESPONSE.findall(conteudo)),
            "usou_modo_fast": len(self._RE_STEP_FAST_SKIPPED.findall(conteudo)) > 0,
            "detalhes_erro": detalhes_erro,
            "erro_principal": erro_principal,
        }

    def _extract_failed_steps(self, conteudo):
        nomes_passos = {}
        for match in self._RE_STEP_STARTED.finditer(conteudo):
            try:
                numero = int(match.group(1))
            except ValueError:
                continue
            nomes_passos[numero] = match.group(2).strip()

        passos = []
        for match in self._RE_STEP_FAILED.finditer(conteudo):
            try:
                numero = int(match.group(1))
            except ValueError:
                continue
            descricao = nomes_passos.get(numero, "").strip()
            if descricao:
                passos.append(f"PASSO {numero}: {descricao}")
            else:
                passos.append(f"PASSO {numero}")
        return passos

    def _extract_error_details(self, conteudo):
        detalhes = []
        for match in self._RE_ERROR_DETAIL.finditer(conteudo):
            detalhe = match.group(1).strip()
            if detalhe:
                detalhes.append(detalhe)
        return detalhes

    def _summarize_primary_error(self, detalhes_erro):
        if not detalhes_erro:
            return ""
        categorias = Counter(self._categorize_error_detail(item) for item in detalhes_erro if item)
        if categorias:
            return categorias.most_common(1)[0][0]
        return detalhes_erro[0]

    def _categorize_error_detail(self, detalhe):
        texto = self._normalize_text(detalhe)
        if "timeout" in texto or "recebida: ''" in texto:
            return "Timeout / Sem resposta"
        if "nenhuma resposta recebida" in texto:
            return "Modbus sem resposta"
        if "padrao de texto" in texto:
            return "Falha de padrao de resposta"
        if "resposta esperada" in texto:
            return "Resposta diferente da esperada"
        if "nao encontrou correspondencia" in texto:
            return "Resposta sem correspondencia"
        return detalhe.strip()

    def _extract_fields(self, conteudo):
        campos = {}
        for linha in conteudo.splitlines():
            texto = linha.strip()
            if not texto:
                continue
            normalizado = self._normalize_text(texto)
            if ":" not in texto:
                continue
            valor = texto.split(":", 1)[1].strip()
            if normalizado.startswith("numero do pr:"):
                campos["pr"] = valor
            elif normalizado.startswith("numero de serie da placa:"):
                campos["numero_serie"] = valor
            elif normalizado.startswith("operador do teste:"):
                campos["operador"] = valor
            elif normalizado.startswith("maquina de teste:"):
                campos["maquina"] = valor
            elif normalizado.startswith("data/hora inicio:"):
                campos["inicio"] = valor
            elif normalizado.startswith("data/hora termino:"):
                campos["fim"] = valor
        return campos

    def _read_text(self, caminho):
        for encoding in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
            try:
                with open(caminho, "r", encoding=encoding) as arquivo:
                    return arquivo.read()
            except UnicodeDecodeError:
                continue
        with open(caminho, "r", encoding="utf-8", errors="ignore") as arquivo:
            return arquivo.read()

    def _normalize_text(self, texto):
        texto = unicodedata.normalize("NFKD", texto)
        texto = texto.encode("ascii", "ignore").decode("ascii")
        return texto.strip().lower()


class RelatorioEficienciaWidget(QWidget):
    def __init__(self, logs_dir=None, parent=None):
        super().__init__(parent)
        self.logs_dir = self._normalize_logs_dir(logs_dir or os.path.expanduser("~/Documents"))
        self.resultados_detalhados = []
        self._analysis_thread = None
        self._analysis_worker = None
        self._progress_dialog = None
        self._hover_targets = []
        self._hover_annotation = None
        self._hover_suppressed_until = 0.0
        self._init_ui()

    def _handle_runtime_error(self, etapa, exc):
        self._close_progress_dialog()
        self.analisar_button.setEnabled(True)
        self.exportar_button.setEnabled(bool(self.resultados_detalhados))
        self.status_label.setText(f"Erro durante {etapa}.")
        self.text_area.setPlainText(f"Erro durante {etapa}:\n{exc}")
        QMessageBox.critical(self, "Erro", f"Falha durante {etapa}:\n{exc}")

    def _init_ui(self):
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Filtros de Relatorio:"))

        filtro_layout = QHBoxLayout()
        self.data_inicio = QDateEdit()
        self.data_inicio.setCalendarPopup(True)
        self.data_fim = QDateEdit()
        self.data_fim.setCalendarPopup(True)
        hoje = QDate.currentDate()
        self.data_inicio.setDate(hoje.addMonths(-1))
        self.data_fim.setDate(hoje)
        self.data_inicio.dateChanged.connect(self._filtrar_e_mostrar)
        self.data_fim.dateChanged.connect(self._filtrar_e_mostrar)
        filtro_layout.addWidget(QLabel("De:"))
        filtro_layout.addWidget(self.data_inicio)
        filtro_layout.addWidget(QLabel("Ate:"))
        filtro_layout.addWidget(self.data_fim)

        self.operador_filtro = QComboBox()
        self.operador_filtro.addItem("Todos os Operadores")
        self.operador_filtro.currentIndexChanged.connect(self._filtrar_e_mostrar)
        filtro_layout.addWidget(QLabel("Operador:"))
        filtro_layout.addWidget(self.operador_filtro)

        self.resultado_filtro = QComboBox()
        self.resultado_filtro.addItems(["Todos os Resultados", "APROVADO", "REPROVADO"])
        self.resultado_filtro.currentIndexChanged.connect(self._filtrar_e_mostrar)
        filtro_layout.addWidget(QLabel("Resultado:"))
        filtro_layout.addWidget(self.resultado_filtro)

        self.placa_filtro = QLineEdit()
        self.placa_filtro.setPlaceholderText("Filtrar por lote ou lote/serie da placa")
        self.placa_filtro.textChanged.connect(self._filtrar_e_mostrar)
        filtro_layout.addWidget(QLabel("Lote/Serie da Placa:"))
        filtro_layout.addWidget(self.placa_filtro)
        layout.addLayout(filtro_layout)

        filtro_extra_layout = QHBoxLayout()
        self.maquina_filtro = QComboBox()
        self.maquina_filtro.addItem("Todas as Maquinas")
        self.maquina_filtro.currentIndexChanged.connect(self._filtrar_e_mostrar)
        filtro_extra_layout.addWidget(QLabel("Maquina:"))
        filtro_extra_layout.addWidget(self.maquina_filtro)

        self.pr_filtro = QLineEdit()
        self.pr_filtro.setPlaceholderText("Filtrar por PR, ex: PR04237")
        self.pr_filtro.textChanged.connect(self._filtrar_e_mostrar)
        filtro_extra_layout.addWidget(QLabel("PR:"))
        filtro_extra_layout.addWidget(self.pr_filtro)

        self.erro_filtro = QLineEdit()
        self.erro_filtro.setPlaceholderText("Filtrar por erro principal")
        self.erro_filtro.textChanged.connect(self._filtrar_e_mostrar)
        filtro_extra_layout.addWidget(QLabel("Erro:"))
        filtro_extra_layout.addWidget(self.erro_filtro)
        layout.addLayout(filtro_extra_layout)

        btns_layout = QHBoxLayout()
        self.analisar_button = QPushButton("Analisar Logs")
        self.analisar_button.clicked.connect(self._analisar_logs)
        btns_layout.addWidget(self.analisar_button)

        self.exportar_button = QPushButton("Exportar Excel")
        self.exportar_button.clicked.connect(self._exportar_excel)
        self.exportar_button.setEnabled(False)
        btns_layout.addWidget(self.exportar_button)
        layout.addLayout(btns_layout)

        self.pasta_label = QLabel(f"Pasta padrao: {self.logs_dir}")
        self.pasta_label.setWordWrap(True)
        layout.addWidget(self.pasta_label)

        self.status_label = QLabel("Nenhuma analise executada.")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        self.content_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.content_splitter.setChildrenCollapsible(False)

        resumo_panel = QWidget()
        resumo_layout = QVBoxLayout(resumo_panel)
        resumo_layout.setContentsMargins(0, 0, 0, 0)
        resumo_layout.setSpacing(6)
        resumo_layout.addWidget(QLabel("Resumo / Terminal"))

        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        self.text_area.setMinimumWidth(240)
        self.text_area.setMaximumWidth(16777215)
        self.text_area.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        resumo_layout.addWidget(self.text_area)

        dashboard_panel = QWidget()
        dashboard_layout = QVBoxLayout(dashboard_panel)
        dashboard_layout.setContentsMargins(0, 0, 0, 0)
        dashboard_layout.setSpacing(6)
        dashboard_layout.addWidget(QLabel("  Dashboard"))

        self.figura = Figure(figsize=(13.8, 34.0), constrained_layout=False)
        self.canvas = DashboardFigureCanvas(self.figura)
        self.canvas.setMinimumSize(900, 2900)
        self.canvas.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.canvas.set_hover_suppression_callback(self._suppress_hover_temporarily)
        self.canvas.mpl_connect("motion_notify_event", self._on_canvas_hover)
        self.canvas.mpl_connect("figure_leave_event", self._on_canvas_leave)

        self.dashboard_container = QWidget()
        self.dashboard_container.setObjectName("dashboard_container")
        self.dashboard_container.setStyleSheet("#dashboard_container { background-color: #2f2f2f; }")
        dashboard_container_layout = QVBoxLayout(self.dashboard_container)
        dashboard_container_layout.setContentsMargins(6, 6, 6, 6)
        dashboard_container_layout.setSpacing(0)
        dashboard_container_layout.addWidget(self.canvas, alignment=Qt.AlignmentFlag.AlignTop)

        self.dashboard_scroll_area = QScrollArea()
        self.dashboard_scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.dashboard_scroll_area.setWidget(self.dashboard_container)
        self.dashboard_scroll_area.setWidgetResizable(True)
        self.dashboard_scroll_area.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.dashboard_scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.dashboard_scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.dashboard_scroll_area.viewport().installEventFilter(self)
        self.canvas.set_scroll_area(self.dashboard_scroll_area)
        dashboard_layout.addWidget(self.dashboard_scroll_area)

        self.content_splitter.addWidget(resumo_panel)
        self.content_splitter.addWidget(dashboard_panel)
        self.content_splitter.setStretchFactor(0, 1)
        self.content_splitter.setStretchFactor(1, 3)
        self.content_splitter.setSizes([320, 960])

        layout.addWidget(self.content_splitter, 1)
        QTimer.singleShot(0, self._sync_dashboard_canvas_size)

    def set_logs_dir(self, logs_dir):
        self.logs_dir = self._normalize_logs_dir(logs_dir or os.path.expanduser("~/Documents"))
        if hasattr(self, "pasta_label"):
            self.pasta_label.setText(f"Pasta padrao: {self.logs_dir}")

    def eventFilter(self, obj, event):
        scroll_area = getattr(self, "dashboard_scroll_area", None)
        if scroll_area is not None and obj is scroll_area.viewport() and event.type() == QEvent.Type.Resize:
            QTimer.singleShot(0, self._sync_dashboard_canvas_size)
        return super().eventFilter(obj, event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        QTimer.singleShot(0, self._sync_dashboard_canvas_size)

    def _sync_dashboard_canvas_size(self):
        if not hasattr(self, "dashboard_scroll_area") or not hasattr(self, "canvas"):
            return
        viewport = self.dashboard_scroll_area.viewport()
        if viewport.width() <= 0:
            return
        horizontal_padding = 12
        viewport_width = max(920, viewport.width() - horizontal_padding)
        aspect_ratio = 34.0 / 13.8
        canvas_height = max(3000, int(viewport_width * aspect_ratio))
        self.canvas.setFixedSize(viewport_width, canvas_height)
        if hasattr(self, "dashboard_container"):
            self.dashboard_container.setMinimumWidth(viewport.width())
        dpi = self.figura.get_dpi() or 100
        self.figura.set_size_inches(viewport_width / dpi, canvas_height / dpi, forward=True)
        self.canvas.draw_idle()

    def _normalize_logs_dir(self, base_dir):
        base_dir = os.path.normpath(base_dir)
        if os.path.basename(base_dir).lower() == "logs do teste":
            return base_dir
        candidato = os.path.join(base_dir, "Logs Do Teste")
        if os.path.isdir(candidato):
            return candidato
        return base_dir

    def _analisar_logs(self):
        if self._analysis_thread is not None:
            QMessageBox.information(self, "Analise em andamento", "Aguarde a analise atual terminar.")
            return

        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Logs", self.logs_dir)
        if not pasta:
            return

        pasta = self._normalize_logs_dir(pasta)
        self.set_logs_dir(pasta)
        self._start_analysis(pasta)

    def _start_analysis(self, pasta):
        self.analisar_button.setEnabled(False)
        self.exportar_button.setEnabled(False)
        self.status_label.setText("Processando logs...")

        self._progress_dialog = QProgressDialog("Preparando leitura dos logs...", "Cancelar", 0, 0, self)
        self._progress_dialog.setWindowTitle("Analisando Logs")
        self._progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self._progress_dialog.setMinimumDuration(0)
        self._progress_dialog.setAutoClose(False)
        self._progress_dialog.setAutoReset(False)
        self._progress_dialog.canceled.connect(self._cancel_analysis)
        self._progress_dialog.show()

        self._analysis_thread = QThread(self)
        self._analysis_worker = LogAnalysisWorker(pasta)
        self._analysis_worker.moveToThread(self._analysis_thread)

        self._analysis_thread.started.connect(self._analysis_worker.run)
        self._analysis_worker.progress.connect(self._on_analysis_progress)
        self._analysis_worker.finished.connect(self._on_analysis_finished)
        self._analysis_worker.failed.connect(self._on_analysis_failed)
        self._analysis_worker.canceled.connect(self._on_analysis_canceled)

        self._analysis_worker.finished.connect(self._analysis_thread.quit)
        self._analysis_worker.failed.connect(self._analysis_thread.quit)
        self._analysis_worker.canceled.connect(self._analysis_thread.quit)
        self._analysis_worker.finished.connect(self._analysis_worker.deleteLater)
        self._analysis_worker.failed.connect(self._analysis_worker.deleteLater)
        self._analysis_worker.canceled.connect(self._analysis_worker.deleteLater)
        self._analysis_thread.finished.connect(self._clear_analysis_refs)
        self._analysis_thread.finished.connect(self._analysis_thread.deleteLater)
        self._analysis_thread.start()

    def _cancel_analysis(self):
        if self._analysis_worker is not None:
            self._analysis_worker.request_cancel()
        progress_dialog = self._progress_dialog
        if progress_dialog is not None:
            progress_dialog.setLabelText("Cancelando analise...")

    def _on_analysis_progress(self, atual, total, nome_arquivo):
        try:
            progress_dialog = self._progress_dialog
            if progress_dialog is not None:
                if progress_dialog.maximum() != max(total, 1):
                    progress_dialog.setRange(0, max(total, 1))
                progress_dialog.setValue(atual)
                progress_dialog.setLabelText(f"Processando {atual}/{total}: {nome_arquivo}")
            self.status_label.setText(f"Processando {atual}/{total} arquivo(s)...")
        except Exception as exc:
            self._handle_runtime_error("a atualizacao do progresso", exc)

    def _on_analysis_finished(self, resultados, operadores, total_arquivos, total_placas, pasta):
        try:
            self.resultados_detalhados = list(resultados)
            self._rebuild_operator_filter(operadores)
            self._rebuild_machine_filter(
                sorted({item["maquina"] for item in self.resultados_detalhados if item.get("maquina")})
            )
            self.set_logs_dir(pasta)
            self._close_progress_dialog()

            if not self.resultados_detalhados:
                self.exportar_button.setEnabled(False)
                self.status_label.setText(
                    f"Analise concluida. Nenhum log valido encontrado em {pasta}."
                )
                self.text_area.setPlainText(
                    f"Nenhum log valido encontrado.\n\nPasta analisada:\n{pasta}"
                )
                self._clear_chart()
                self.analisar_button.setEnabled(True)
                return

            self.status_label.setText(
                f"Analise concluida. {total_placas} placa(s) unicas encontradas em {total_arquivos} arquivo(s)."
            )
            self.exportar_button.setEnabled(True)
            self.analisar_button.setEnabled(True)
            self._filtrar_e_mostrar()
        except Exception as exc:
            self._handle_runtime_error("a finalizacao da analise", exc)

    def _on_analysis_failed(self, erro):
        self._close_progress_dialog()
        self.analisar_button.setEnabled(True)
        self.exportar_button.setEnabled(bool(self.resultados_detalhados))
        self.status_label.setText("Falha ao analisar os logs.")
        QMessageBox.critical(self, "Erro na Analise", f"Nao foi possivel analisar os logs:\n{erro}")

    def _on_analysis_canceled(self):
        self._close_progress_dialog()
        self.analisar_button.setEnabled(True)
        self.exportar_button.setEnabled(bool(self.resultados_detalhados))
        self.status_label.setText("Analise cancelada pelo usuario.")

    def _clear_analysis_refs(self):
        self._analysis_thread = None
        self._analysis_worker = None

    def _close_progress_dialog(self):
        if self._progress_dialog is not None:
            try:
                self._progress_dialog.canceled.disconnect(self._cancel_analysis)
            except Exception:
                pass
            self._progress_dialog.close()
            self._progress_dialog.deleteLater()
            self._progress_dialog = None

    def _rebuild_operator_filter(self, operadores):
        selecionado_atual = self.operador_filtro.currentText()
        with QSignalBlocker(self.operador_filtro):
            self.operador_filtro.clear()
            self.operador_filtro.addItem("Todos os Operadores")
            for operador in operadores:
                self.operador_filtro.addItem(operador)
            index = self.operador_filtro.findText(selecionado_atual)
            if index >= 0:
                self.operador_filtro.setCurrentIndex(index)
            else:
                self.operador_filtro.setCurrentIndex(0)

    def _rebuild_machine_filter(self, maquinas):
        selecionado_atual = self.maquina_filtro.currentText()
        with QSignalBlocker(self.maquina_filtro):
            self.maquina_filtro.clear()
            self.maquina_filtro.addItem("Todas as Maquinas")
            for maquina in maquinas:
                self.maquina_filtro.addItem(maquina)
            index = self.maquina_filtro.findText(selecionado_atual)
            if index >= 0:
                self.maquina_filtro.setCurrentIndex(index)
            else:
                self.maquina_filtro.setCurrentIndex(0)

    def _get_filtered_results(self):
        data_ini = self.data_inicio.date().toPyDate()
        data_fim = self.data_fim.date().toPyDate()
        operador_filtro = self.operador_filtro.currentText()
        resultado_filtro = self.resultado_filtro.currentText()
        maquina_filtro = self.maquina_filtro.currentText()
        placa_filtro = self.placa_filtro.text().strip()
        pr_filtro = self.pr_filtro.text().strip().upper()
        erro_filtro = self.erro_filtro.text().strip().upper()

        filtrados = []
        for registro in self.resultados_detalhados:
            if not (data_ini <= registro["inicio"].date() <= data_fim):
                continue
            if operador_filtro != "Todos os Operadores" and registro["operador"] != operador_filtro:
                continue
            if resultado_filtro != "Todos os Resultados" and registro["resultado"] != resultado_filtro:
                continue
            if maquina_filtro != "Todas as Maquinas" and (registro.get("maquina") or "") != maquina_filtro:
                continue
            if placa_filtro and not self._matches_placa_filter(registro["numero_serie"], placa_filtro):
                continue
            if pr_filtro and pr_filtro not in (registro.get("pr") or "").upper():
                continue
            if erro_filtro and erro_filtro not in (registro.get("erro_principal") or "").upper():
                continue
            filtrados.append(registro)
        return filtrados

    def _aggregate_filtered_data(self, resultados_filtrados):
        resultados = defaultdict(
            lambda: {
                "aprovado": 0,
                "reprovado": 0,
                "total": 0,
                "tempos": [],
                "passos": Counter(),
            }
        )
        dias = defaultdict(list)

        for registro in resultados_filtrados:
            bucket = resultados[registro["operador"]]
            bucket["total"] += 1
            if registro["resultado"] == "APROVADO":
                bucket["aprovado"] += 1
            elif registro["resultado"] == "REPROVADO":
                bucket["reprovado"] += 1
            bucket["tempos"].append(registro["duracao_segundos"])
            bucket["passos"].update(registro.get("passos_reprovados", []))
            dias[registro["inicio"].date()].append(registro["duracao_segundos"])

        return resultados, dias

    def _matches_placa_filter(self, numero_serie, placa_filtro):
        numero_serie_upper = (numero_serie or "").upper()
        filtro_upper = placa_filtro.upper()

        if "/" in placa_filtro and not placa_filtro.endswith("/"):
            return numero_serie_upper == filtro_upper

        prefixo = filtro_upper
        if not prefixo.endswith("/"):
            prefixo += "/"
        return numero_serie_upper.startswith(prefixo)

    def _filtrar_e_mostrar(self):
        try:
            if not self.resultados_detalhados:
                self.text_area.setPlainText("Nenhum dado carregado para analise.")
                self._clear_chart()
                self.exportar_button.setEnabled(False)
                return

            resultados_filtrados = self._get_filtered_results()
            if not resultados_filtrados:
                self.text_area.setPlainText("Nenhum dado encontrado para os filtros selecionados.")
                self._clear_chart()
                self.exportar_button.setEnabled(False)
                return

            self.exportar_button.setEnabled(True)
            resultados, dias = self._aggregate_filtered_data(resultados_filtrados)

            linhas_relatorio = self._build_dashboard_summary(resultados, dias, resultados_filtrados)

            for operador in sorted(resultados.keys()):
                dados = resultados[operador]
                media = sum(dados["tempos"]) / len(dados["tempos"]) if dados["tempos"] else 0.0
                aprovacao = 100 * dados["aprovado"] / dados["total"] if dados["total"] else 0.0
                top_passos = dados["passos"].most_common(3)

                linhas_relatorio.extend(
                    [
                        f"Operador: {operador}",
                        f"  Total de Testes: {dados['total']}",
                        f"  Aprovados: {dados['aprovado']}",
                        f"  Reprovados: {dados['reprovado']}",
                        f"  % Aprovacao: {aprovacao:.1f}%",
                        f"  Tempo Medio: {media:.1f} s ({media / 60:.1f} min)",
                    ]
                )

                if top_passos:
                    linhas_relatorio.append("  Passos Criticos:")
                    for passo, ocorrencias in top_passos:
                        linhas_relatorio.append(f"    - {passo} ({ocorrencias} ocorrencia(s))")
                else:
                    linhas_relatorio.append("  Passos Criticos: Nenhum")

                linhas_relatorio.append("-" * 30)

            self.text_area.setPlainText("\n".join(linhas_relatorio))
            self._plotar_graficos(resultados, dias, resultados_filtrados)
        except Exception as exc:
            self._handle_runtime_error("a exibicao do relatorio", exc)

    def _format_operator_label(self, nome, limite=14):
        nome = (nome or "").strip()
        if len(nome) <= limite:
            return nome
        partes = nome.split()
        if len(partes) >= 2:
            primeira = partes[0]
            restante = " ".join(partes[1:])
            if len(primeira) <= limite and len(restante) <= limite:
                return f"{primeira}\n{restante}"
        return nome[: max(limite - 3, 1)] + "..."

    def _format_step_label(self, nome, limite=18):
        nome = (nome or "").strip()
        if len(nome) <= limite:
            return nome
        return nome[: max(limite - 3, 1)] + "..."

    def _extract_step_number_label(self, nome):
        texto = (nome or "").strip()
        match = re.match(r"^(PASSO\s+\d+)", texto, re.IGNORECASE)
        if match:
            return match.group(1).upper()
        return self._format_step_label(texto, 10)

    def _extract_step_description(self, nome):
        texto = (nome or "").strip()
        match = re.match(r"^PASSO\s+\d+\s*:\s*(.+)$", texto, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return texto

    def _shorten_text(self, texto, limite=42):
        texto = (texto or "").strip()
        if len(texto) <= limite:
            return texto
        return texto[: max(limite - 3, 1)] + "..."

    def _clear_hover_targets(self):
        self._hover_targets = []
        annotation = getattr(self, "_hover_annotation", None)
        if annotation is not None:
            annotation.set_visible(False)

    def _ensure_hover_annotation(self):
        if self._hover_annotation is None:
            axis = self.figura.add_axes([0, 0, 1, 1], frameon=False)
            axis.set_axis_off()
            self._hover_annotation = axis.annotate(
                "",
                xy=(0, 0),
                xycoords="figure pixels",
                xytext=(16, 16),
                textcoords="offset points",
                ha="left",
                va="bottom",
                fontsize=8,
                color="#1f1f1f",
                bbox=dict(boxstyle="round,pad=0.35", fc="#FFFDEB", ec="#999999", alpha=0.96),
                annotation_clip=False,
            )
            self._hover_annotation.set_visible(False)
        return self._hover_annotation

    def _register_hover_artist(self, artist, tooltip):
        if artist is None or not tooltip:
            return
        self._hover_targets.append({"kind": "artist", "artist": artist, "tooltip": tooltip})

    def _register_hover_axes(self, axis, tooltip):
        if axis is None or not tooltip:
            return
        self._hover_targets.append({"kind": "axes", "axes": axis, "tooltip": tooltip})

    def _register_heatmap_hover(self, axis, matrix, datas, operadores):
        if axis is None or matrix is None or not datas or not operadores:
            return
        self._hover_targets.append(
            {
                "kind": "heatmap",
                "axes": axis,
                "matrix": matrix,
                "datas": datas,
                "operadores": operadores,
            }
        )

    def _resolve_hover_tooltip(self, event):
        for target in reversed(self._hover_targets):
            kind = target.get("kind")
            if kind == "artist":
                artist = target.get("artist")
                try:
                    contains, _details = artist.contains(event)
                except Exception:
                    contains = False
                if contains:
                    return target.get("tooltip", "")
            elif kind == "axes":
                axis = target.get("axes")
                if event.inaxes is axis:
                    return target.get("tooltip", "")
            elif kind == "heatmap":
                axis = target.get("axes")
                if event.inaxes is not axis or event.xdata is None or event.ydata is None:
                    continue
                col = int(round(event.xdata))
                row = int(round(event.ydata))
                matrix = target.get("matrix") or []
                if row < 0 or col < 0 or row >= len(matrix):
                    continue
                if not matrix[row] or col >= len(matrix[row]):
                    continue
                data = target["datas"][col]
                operador = target["operadores"][row]
                quantidade = matrix[row][col]
                return f"Operador: {operador}\nData: {data.strftime('%d/%m/%Y')}\nTestes: {quantidade}"
        return ""

    def _show_hover_annotation(self, event, texto):
        annotation = self._ensure_hover_annotation()
        annotation.xy = (event.x, event.y)
        annotation.set_text(texto)
        annotation.set_visible(True)
        self.canvas.draw_idle()

    def _hide_hover_annotation(self):
        annotation = getattr(self, "_hover_annotation", None)
        if annotation is not None and annotation.get_visible():
            annotation.set_visible(False)
            self.canvas.draw_idle()

    def _suppress_hover_temporarily(self, duration_s=0.25):
        self._hover_suppressed_until = time.monotonic() + max(duration_s, 0.0)
        self._hide_hover_annotation()

    def _on_canvas_hover(self, event):
        if time.monotonic() < getattr(self, "_hover_suppressed_until", 0.0):
            return
        if event is None or event.inaxes is None:
            self._hide_hover_annotation()
            return
        texto = self._resolve_hover_tooltip(event)
        if texto:
            self._show_hover_annotation(event, texto)
        else:
            self._hide_hover_annotation()

    def _on_canvas_leave(self, _event):
        self._hide_hover_annotation()

    def _safe_percent(self, numerador, denominador):
        if not denominador:
            return 0.0
        return 100.0 * numerador / denominador

    def _build_dashboard_metrics(self, resultados, dias, resultados_filtrados):
        total_placas = len(resultados_filtrados)
        aprovados_total = sum(1 for item in resultados_filtrados if item["resultado"] == "APROVADO")
        reprovados_total = sum(1 for item in resultados_filtrados if item["resultado"] == "REPROVADO")
        taxa_aprovacao = self._safe_percent(aprovados_total, total_placas)
        total_retentativas = sum(item.get("retentativas_automaticas", 0) for item in resultados_filtrados)
        total_timeouts = sum(item.get("timeouts", 0) for item in resultados_filtrados)
        tempo_medio_min = (
            sum(item["duracao_segundos"] for item in resultados_filtrados) / total_placas / 60
            if total_placas
            else 0.0
        )

        operador_destaque = "-"
        if resultados:
            ranking_operadores = sorted(
                resultados.items(),
                key=lambda item: (
                    self._safe_percent(item[1]["aprovado"], item[1]["total"]),
                    item[1]["total"],
                    -((sum(item[1]["tempos"]) / len(item[1]["tempos"])) if item[1]["tempos"] else 0.0),
                ),
                reverse=True,
            )
            operador_destaque = ranking_operadores[0][0]

        maquinas = Counter(item.get("maquina") or "Nao informado" for item in resultados_filtrados)
        maquina_top, maquina_total = maquinas.most_common(1)[0] if maquinas else ("-", 0)

        prs = Counter(item.get("pr") or "-" for item in resultados_filtrados)
        pr_top, pr_total = prs.most_common(1)[0] if prs else ("-", 0)

        dias_volume = Counter(item["inicio"].date() for item in resultados_filtrados)
        dia_pico, dia_pico_total = dias_volume.most_common(1)[0] if dias_volume else (None, 0)
        dia_pico_texto = dia_pico.strftime("%d/%m/%Y") if dia_pico else "-"

        return [
            {
                "title": "Placas Filtradas",
                "value": str(total_placas),
                "subtitle": f"{aprovados_total} aprovadas / {reprovados_total} reprovadas",
                "facecolor": "#EAF3FF",
                "accent": "#3D85C6",
            },
            {
                "title": "Aprovacao Geral",
                "value": f"{taxa_aprovacao:.1f}%",
                "subtitle": "Indicador geral do periodo filtrado",
                "facecolor": "#EDF8EC",
                "accent": "#6AA84F",
            },
            {
                "title": "Tempo Medio",
                "value": f"{tempo_medio_min:.2f} min",
                "subtitle": "Media por placa testada",
                "facecolor": "#FFF4E5",
                "accent": "#E69138",
            },
            {
                "title": "Retentativas",
                "value": str(total_retentativas),
                "subtitle": "Total de retentativas automaticas",
                "facecolor": "#FFF7DA",
                "accent": "#BF9000",
            },
            {
                "title": "Timeouts",
                "value": str(total_timeouts),
                "subtitle": "Ocorrencias de timeout / sem resposta",
                "facecolor": "#FCE5E5",
                "accent": "#CC0000",
            },
            {
                "title": "Operador Destaque",
                "value": operador_destaque,
                "subtitle": "Maior taxa de aprovacao no filtro",
                "facecolor": "#F2ECFA",
                "accent": "#8E7CC3",
            },
            {
                "title": "Maquina Lider",
                "value": maquina_top,
                "subtitle": f"{maquina_total} teste(s) no periodo",
                "facecolor": "#FDEDEC",
                "accent": "#C0504D",
            },
            {
                "title": "PR / Dia Pico",
                "value": pr_top,
                "subtitle": f"Dia pico: {dia_pico_texto} ({dia_pico_total} teste(s))",
                "facecolor": "#EEF7F7",
                "accent": "#45818E",
            },
        ]

    def _build_dashboard_summary(self, resultados, dias, resultados_filtrados):
        total_placas = len(resultados_filtrados)
        aprovados_total = sum(1 for item in resultados_filtrados if item["resultado"] == "APROVADO")
        reprovados_total = sum(1 for item in resultados_filtrados if item["resultado"] == "REPROVADO")
        taxa_aprovacao = self._safe_percent(aprovados_total, total_placas)
        total_retentativas = sum(item.get("retentativas_automaticas", 0) for item in resultados_filtrados)
        total_timeouts = sum(item.get("timeouts", 0) for item in resultados_filtrados)
        total_pulados_fast = sum(item.get("passos_pulados_fast", 0) for item in resultados_filtrados)
        tempo_medio_seg = (
            sum(item["duracao_segundos"] for item in resultados_filtrados) / total_placas
            if total_placas
            else 0.0
        )

        maquinas = Counter(item.get("maquina") or "Nao informado" for item in resultados_filtrados)
        maquina_top, maquina_total = maquinas.most_common(1)[0] if maquinas else ("-", 0)

        prs = Counter(item.get("pr") or "-" for item in resultados_filtrados)
        pr_top, pr_total = prs.most_common(1)[0] if prs else ("-", 0)

        dias_volume = Counter(item["inicio"].date() for item in resultados_filtrados)
        dia_pico, dia_pico_total = dias_volume.most_common(1)[0] if dias_volume else (None, 0)
        dia_pico_texto = dia_pico.strftime("%d/%m/%Y") if dia_pico else "-"

        operador_destaque = "-"
        if resultados:
            ranking_operadores = sorted(
                resultados.items(),
                key=lambda item: (
                    self._safe_percent(item[1]["aprovado"], item[1]["total"]),
                    item[1]["total"],
                ),
                reverse=True,
            )
            operador_destaque = ranking_operadores[0][0]
            ranking_qualidade = [
                f"{nome} ({self._safe_percent(dados['aprovado'], dados['total']):.1f}%)"
                for nome, dados in ranking_operadores[:3]
            ]
            ranking_rapidez = [
                (
                    nome,
                    (sum(dados["tempos"]) / len(dados["tempos"])) if dados["tempos"] else 0.0,
                )
                for nome, dados in resultados.items()
                if dados["tempos"]
            ]
            ranking_rapidez.sort(key=lambda item: item[1])
            ranking_rapidez = [f"{nome} ({tempo/60:.2f} min)" for nome, tempo in ranking_rapidez[:3]]
        else:
            ranking_qualidade = []
            ranking_rapidez = []

        prs_reprovados = Counter(
            item.get("pr") or "-"
            for item in resultados_filtrados
            if item["resultado"] == "REPROVADO"
        )
        prs_reprovados_top = [
            f"{pr} ({quantidade} reprov.)" for pr, quantidade in prs_reprovados.most_common(3)
        ]
        erros_principais = Counter(
            item.get("erro_principal") or "Sem erro detalhado"
            for item in resultados_filtrados
            if item.get("erro_principal")
        )
        erros_principais_top = [
            f"{erro} ({quantidade})" for erro, quantidade in erros_principais.most_common(3)
        ]

        return [
            "Resumo Executivo",
            "-" * 40,
            f"Total de Placas Testadas (Filtradas): {total_placas}",
            f"Aprovadas: {aprovados_total}",
            f"Reprovadas: {reprovados_total}",
            f"Taxa Geral de Aprovacao: {taxa_aprovacao:.1f}%",
            f"Tempo Medio Geral: {tempo_medio_seg:.1f} s ({tempo_medio_seg / 60:.2f} min)",
            f"Retentativas Automaticas Totais: {total_retentativas}",
            f"Timeouts / Sem Resposta: {total_timeouts}",
            f"Passos Pulados no Modo Fast: {total_pulados_fast}",
            f"Operador Destaque: {operador_destaque}",
            f"Maquina com Maior Volume: {maquina_top} ({maquina_total} teste(s))",
            f"PR com Maior Volume: {pr_top} ({pr_total} teste(s))",
            f"Dia com Maior Volume: {dia_pico_texto} ({dia_pico_total} teste(s))",
            f"Top 3 Operadores por Qualidade: {', '.join(ranking_qualidade) if ranking_qualidade else '-'}",
            f"Top 3 Operadores Mais Rapidos: {', '.join(ranking_rapidez) if ranking_rapidez else '-'}",
            f"PRs com Mais Reprovacoes: {', '.join(prs_reprovados_top) if prs_reprovados_top else '-'}",
            f"Erros Principais: {', '.join(erros_principais_top) if erros_principais_top else '-'}",
            "",
            "Eficiencia por Operador",
            "-" * 40,
        ]

    def _draw_metric_card(self, axis, title, value, subtitle, facecolor, accent):
        value = str(value)
        title_font = 8.0
        value_font = 13 if len(value) <= 10 else 11
        subtitle_font = 6.8
        axis.set_facecolor(facecolor)
        axis.set_xticks([])
        axis.set_yticks([])
        axis.set_xlim(0, 1)
        axis.set_ylim(0, 1)
        for spine in axis.spines.values():
            spine.set_visible(False)
        axis.axhline(0.98, color=accent, linewidth=4, xmin=0.03, xmax=0.97)
        axis.text(0.05, 0.84, title, fontsize=title_font, fontweight="bold", color="#444444", va="top", clip_on=True)
        axis.text(0.05, 0.48, value, fontsize=value_font, fontweight="bold", color=accent, va="center", clip_on=True)
        axis.text(0.05, 0.10, subtitle, fontsize=subtitle_font, color="#555555", va="bottom", wrap=True, clip_on=True)
        self._register_hover_artist(axis.patch, f"{title}\n{value}\n{subtitle}")

    def _apply_axis_style(self, axis):
        axis.set_facecolor("#F7F9FC")
        axis.grid(axis="y", linestyle="--", alpha=0.25)
        for spine in axis.spines.values():
            spine.set_alpha(0.15)

    def _plotar_graficos(self, resultados, dias, resultados_filtrados):
        self._clear_hover_targets()
        self._hover_annotation = None
        self.figura.clear()
        self.figura.subplots_adjust(top=0.985, bottom=0.03, left=0.08, right=0.985)
        if not resultados:
            self.canvas.draw_idle()
            return

        metricas = self._build_dashboard_metrics(resultados, dias, resultados_filtrados)
        operadores = sorted(resultados.keys())
        operadores_labels = [self._format_operator_label(operador) for operador in operadores]
        totais = [resultados[operador]["total"] for operador in operadores]
        aprovados = [resultados[operador]["aprovado"] for operador in operadores]
        reprovados = [resultados[operador]["reprovado"] for operador in operadores]
        aprovacao_pct = [
            (100 * resultados[operador]["aprovado"] / resultados[operador]["total"])
            if resultados[operador]["total"]
            else 0
            for operador in operadores
        ]

        posicoes = list(range(len(operadores)))
        grid = self.figura.add_gridspec(
            9,
            1,
            height_ratios=[1.35, 1.45, 1.2, 1.2, 1.25, 1.25, 1.3, 1.5, 1.5],
            hspace=0.95,
        )
        cards_grid = grid[0, 0].subgridspec(2, 4, hspace=0.42, wspace=0.18)

        for idx, metrica in enumerate(metricas):
            card_axis = self.figura.add_subplot(cards_grid[idx // 4, idx % 4])
            self._draw_metric_card(
                card_axis,
                metrica["title"],
                metrica["value"],
                metrica["subtitle"],
                metrica["facecolor"],
                metrica["accent"],
            )

        ax1 = self.figura.add_subplot(grid[1, 0])
        ax2 = self.figura.add_subplot(grid[2, 0])
        ax3 = self.figura.add_subplot(grid[3, 0])
        ax4 = self.figura.add_subplot(grid[4, 0])
        ax5 = self.figura.add_subplot(grid[5, 0])
        ax6 = self.figura.add_subplot(grid[6, 0])
        ax7 = self.figura.add_subplot(grid[7, 0])
        ax8 = self.figura.add_subplot(grid[8, 0])

        self._apply_axis_style(ax1)
        self._apply_axis_style(ax2)
        self._apply_axis_style(ax3)
        self._apply_axis_style(ax4)
        self._apply_axis_style(ax5)
        self._apply_axis_style(ax6)
        self._apply_axis_style(ax8)

        ax1.bar(posicoes, aprovados, color="#6AA84F", label="Aprovados")
        ax1.bar(posicoes, reprovados, bottom=aprovados, color="#E06666", label="Reprovados")
        ax1.set_title("Testes por Operador", fontsize=12, fontweight="bold", pad=16)
        ax1.set_xticks(posicoes)
        ax1.set_xticklabels(operadores_labels, fontsize=9, rotation=12, ha="right")
        ax1.yaxis.set_major_locator(MaxNLocator(integer=True))
        ax1.legend(loc="lower left", bbox_to_anchor=(0.0, 1.02), ncols=2, frameon=False, borderaxespad=0.0)
        max_total = max(totais) if totais else 0
        ax1.set_ylim(0, max(5, max_total * 1.28))
        ax1.margins(x=0.04, y=0.10)
        for index, total in enumerate(totais):
            ax1.text(
                index,
                total + max(max_total * 0.02, 2.5),
                str(total),
                ha="center",
                va="bottom",
                fontsize=8,
                color="#444444",
            )
        for operador, aprovado, reprovado, total, bar_aprovado, bar_reprovado in zip(
            operadores, aprovados, reprovados, totais, ax1.patches[: len(aprovados)], ax1.patches[len(aprovados):]
        ):
            self._register_hover_artist(
                bar_aprovado,
                f"Operador: {operador}\nAprovados: {aprovado}\nReprovados: {reprovado}\nTotal: {total}",
            )
            self._register_hover_artist(
                bar_reprovado,
                f"Operador: {operador}\nAprovados: {aprovado}\nReprovados: {reprovado}\nTotal: {total}",
            )

        pares_aprovacao = sorted(zip(operadores, aprovacao_pct), key=lambda item: item[1], reverse=True)
        operadores_ordenados = [item[0] for item in pares_aprovacao]
        aprovacao_ordenada = [item[1] for item in pares_aprovacao]
        posicoes_aprovacao = list(range(len(operadores_ordenados)))
        ax2.barh(posicoes_aprovacao, aprovacao_ordenada, color="#8E7CC3")
        ax2.set_title("Taxa de Aprovacao por Operador", fontsize=11, fontweight="bold", pad=12)
        ax2.set_yticks(posicoes_aprovacao)
        ax2.set_yticklabels([self._format_operator_label(nome, 18).replace("\n", " ") for nome in operadores_ordenados], fontsize=9)
        ax2.set_xlim(0, 106)
        ax2.set_xlabel("% de aprovacao")
        ax2.invert_yaxis()
        for index, valor in enumerate(aprovacao_ordenada):
            ax2.text(min(valor + 1.0, 104.0), index, f"{valor:.1f}%", va="center", fontsize=8)
        for operador, valor, barra in zip(operadores_ordenados, aprovacao_ordenada, ax2.patches):
            self._register_hover_artist(
                barra,
                f"Operador: {operador}\nTaxa de aprovacao: {valor:.1f}%",
            )

        maquinas = defaultdict(lambda: {"aprovado": 0, "reprovado": 0, "total": 0})
        for registro in resultados_filtrados:
            maquina = registro.get("maquina") or "Nao informado"
            maquinas[maquina]["total"] += 1
            if registro["resultado"] == "APROVADO":
                maquinas[maquina]["aprovado"] += 1
            elif registro["resultado"] == "REPROVADO":
                maquinas[maquina]["reprovado"] += 1

        maquinas_ordenadas = sorted(maquinas.items(), key=lambda item: item[1]["total"], reverse=True)[:6]
        nomes_maquinas = [item[0] for item in maquinas_ordenadas]
        aprovados_maquina = [item[1]["aprovado"] for item in maquinas_ordenadas]
        reprovados_maquina = [item[1]["reprovado"] for item in maquinas_ordenadas]
        posicoes_maquina = list(range(len(nomes_maquinas)))
        if nomes_maquinas:
            ax3.bar(posicoes_maquina, aprovados_maquina, color="#76A5AF", label="Aprovados")
            ax3.bar(posicoes_maquina, reprovados_maquina, bottom=aprovados_maquina, color="#C27BA0", label="Reprovados")
            ax3.set_title("Comparacao por Maquina", fontsize=11, fontweight="bold", pad=12)
            ax3.set_xticks(posicoes_maquina)
            ax3.set_xticklabels([self._format_operator_label(nome, 14) for nome in nomes_maquinas], fontsize=8, rotation=10, ha="right")
            ax3.yaxis.set_major_locator(MaxNLocator(integer=True))
            ax3.legend(loc="lower left", bbox_to_anchor=(0.0, 1.02), frameon=False, fontsize=8, borderaxespad=0.0)
            max_total_maquina = max((a + r) for a, r in zip(aprovados_maquina, reprovados_maquina)) if nomes_maquinas else 0
            ax3.set_ylim(0, max(5, max_total_maquina * 1.18))
            for maquina, aprovado, reprovado, bar_aprovado, bar_reprovado in zip(
                nomes_maquinas,
                aprovados_maquina,
                reprovados_maquina,
                ax3.patches[: len(aprovados_maquina)],
                ax3.patches[len(aprovados_maquina):],
            ):
                total_maquina = aprovado + reprovado
                tooltip_maquina = (
                    f"Maquina: {maquina}\nAprovados: {aprovado}\nReprovados: {reprovado}\nTotal: {total_maquina}"
                )
                self._register_hover_artist(bar_aprovado, tooltip_maquina)
                self._register_hover_artist(bar_reprovado, tooltip_maquina)
        else:
            ax3.text(0.5, 0.5, "Sem dados por maquina", ha="center", va="center", transform=ax3.transAxes)
            ax3.set_title("Comparacao por Maquina", fontsize=11, fontweight="bold", pad=12)

        prs_reprovados = Counter(
            registro.get("pr") or "-"
            for registro in resultados_filtrados
            if registro["resultado"] == "REPROVADO"
        )
        prs_top = prs_reprovados.most_common(6)
        if prs_top:
            pr_nomes = [item[0] for item in prs_top]
            pr_valores = [item[1] for item in prs_top]
            posicoes_pr = list(range(len(pr_nomes)))
            ax4.barh(posicoes_pr, pr_valores, color="#C0504D")
            ax4.set_title("PRs com Mais Reprovacoes", fontsize=11, fontweight="bold", pad=12)
            ax4.set_yticks(posicoes_pr)
            ax4.set_yticklabels(pr_nomes, fontsize=8)
            ax4.invert_yaxis()
            ax4.set_xlabel("Reprovacoes")
            ax4.xaxis.set_major_locator(MaxNLocator(integer=True))
            ax4.set_xlim(0, max(pr_valores) * 1.12 if pr_valores else 1)
            for index, valor in enumerate(pr_valores):
                ax4.text(valor + max(max(pr_valores) * 0.01, 0.2), index, str(valor), va="center", fontsize=8)
            for pr_nome, pr_valor, barra in zip(pr_nomes, pr_valores, ax4.patches):
                self._register_hover_artist(
                    barra,
                    f"PR: {pr_nome}\nReprovacoes: {pr_valor}",
                )
        else:
            ax4.text(0.5, 0.5, "Nenhuma reprovacao no filtro atual", ha="center", va="center", transform=ax4.transAxes)
            ax4.set_title("PRs com Mais Reprovacoes", fontsize=11, fontweight="bold", pad=12)

        datas_filtradas = sorted(dias.keys())
        medias_minutos = [
            (sum(dias[data]) / len(dias[data])) / 60
            for data in datas_filtradas
            if dias[data]
        ]
        datas_plot = [data for data in datas_filtradas if dias[data]]
        if datas_plot:
            ax5.plot(datas_plot, medias_minutos, marker="o", color="#E69138", linewidth=2.0)
            ax5.fill_between(datas_plot, medias_minutos, color="#F9CB9C", alpha=0.35)
            ax5.set_title("Tempo Medio por Dia", fontsize=11, fontweight="bold", pad=12)
            ax5.set_ylabel("Minutos")
            ax5.tick_params(axis="x", labelrotation=25, labelsize=8)
            ax5.margins(x=0.03, y=0.12)
            if ax5.lines:
                self._register_hover_artist(
                    ax5.lines[0],
                    "Passe o mouse sobre os pontos para ver o tempo medio diario.",
                )
            for data, valor in zip(datas_plot, medias_minutos):
                marcador = ax5.scatter([data], [valor], s=55, alpha=0.0)
                self._register_hover_artist(
                    marcador,
                    f"Data: {data.strftime('%d/%m/%Y')}\nTempo medio: {valor:.2f} min",
                )
        else:
            ax5.text(0.5, 0.5, "Sem dados diarios", ha="center", va="center", transform=ax5.transAxes)
            ax5.set_title("Tempo Medio por Dia", fontsize=11, fontweight="bold", pad=12)

        status_por_dia = defaultdict(lambda: {"aprovado": 0, "reprovado": 0})
        for registro in resultados_filtrados:
            bucket = status_por_dia[registro["inicio"].date()]
            if registro["resultado"] == "APROVADO":
                bucket["aprovado"] += 1
            elif registro["resultado"] == "REPROVADO":
                bucket["reprovado"] += 1

        datas_status = sorted(status_por_dia.keys())
        reprovacao_diaria = []
        for data in datas_status:
            aprovado = status_por_dia[data]["aprovado"]
            reprovado = status_por_dia[data]["reprovado"]
            reprovacao_diaria.append(self._safe_percent(reprovado, aprovado + reprovado))

        if datas_status:
            ax6.plot(datas_status, reprovacao_diaria, marker="o", color="#CC0000", linewidth=2.0)
            ax6.fill_between(datas_status, reprovacao_diaria, color="#F4CCCC", alpha=0.4)
            ax6.set_title("Tendencia de Reprovacao por Dia", fontsize=11, fontweight="bold", pad=12)
            ax6.set_ylabel("% Reprovacao")
            ax6.set_ylim(-4, max(100, (max(reprovacao_diaria) + 14) if reprovacao_diaria else 100))
            ax6.tick_params(axis="x", labelrotation=25, labelsize=8)
            ax6.margins(x=0.03, y=0.08)
            if ax6.lines:
                self._register_hover_artist(
                    ax6.lines[0],
                    "Passe o mouse sobre os pontos para ver a reprovacao diaria.",
                )
            for data, valor in zip(datas_status, reprovacao_diaria):
                marcador = ax6.scatter([data], [valor], s=55, alpha=0.0)
                self._register_hover_artist(
                    marcador,
                    f"Data: {data.strftime('%d/%m/%Y')}\nReprovacao: {valor:.1f}%",
                )
        else:
            ax6.text(0.5, 0.5, "Sem dados diarios", ha="center", va="center", transform=ax6.transAxes)
            ax6.set_title("Tendencia de Reprovacao por Dia", fontsize=11, fontweight="bold", pad=12)

        heatmap_operadores = sorted(resultados.keys())
        heatmap_datas = sorted({registro["inicio"].date() for registro in resultados_filtrados})
        if len(heatmap_datas) > 14:
            heatmap_datas = heatmap_datas[-14:]
        heatmap_matrix = []
        for operador in heatmap_operadores:
            linha = []
            for data in heatmap_datas:
                quantidade = sum(
                    1
                    for registro in resultados_filtrados
                    if registro["operador"] == operador and registro["inicio"].date() == data
                )
                linha.append(quantidade)
            heatmap_matrix.append(linha)

        if heatmap_operadores and heatmap_datas:
            ax7.set_facecolor("#F7F9FC")
            heatmap = ax7.imshow(heatmap_matrix, aspect="auto", cmap="YlGnBu")
            ax7.set_title("Heatmap de Testes por Dia x Operador", fontsize=11, fontweight="bold", pad=12)
            ax7.set_xticks(list(range(len(heatmap_datas))))
            ax7.set_xticklabels([data.strftime("%d/%m") for data in heatmap_datas], rotation=30, ha="right", fontsize=8)
            ax7.set_yticks(list(range(len(heatmap_operadores))))
            ax7.set_yticklabels([self._format_operator_label(nome, 16).replace("\n", " ") for nome in heatmap_operadores], fontsize=8)
            ax7.tick_params(axis="y", pad=4)
            self.figura.colorbar(heatmap, ax=ax7, fraction=0.03, pad=0.02)
            self._register_heatmap_hover(ax7, heatmap_matrix, heatmap_datas, heatmap_operadores)
        else:
            ax7.axis("off")
            ax7.text(0.5, 0.5, "Sem dados suficientes para heatmap", ha="center", va="center", transform=ax7.transAxes)

        passos_counter = Counter()
        for registro in resultados_filtrados:
            passos_counter.update(registro.get("passos_reprovados", []))

        passos_criticos = passos_counter.most_common(5)
        if passos_criticos:
            nomes_passos = [item[0] for item in passos_criticos]
            valores_passos = [item[1] for item in passos_criticos]
            descricoes_passos = [self._extract_step_description(item) for item in nomes_passos]
            rotulos_passos = [self._extract_step_number_label(item) for item in nomes_passos]
            posicoes_passos = list(range(len(nomes_passos)))
            ax8.barh(posicoes_passos, valores_passos, color="#CC8F48")
            ax8.set_yticks(posicoes_passos)
            ax8.set_yticklabels(rotulos_passos, fontsize=8)
            ax8.invert_yaxis()
            ax8.set_title("Passos Criticos Mais Frequentes", fontsize=11, fontweight="bold", pad=12)
            ax8.set_xlabel("Ocorrencias")
            ax8.xaxis.set_major_locator(MaxNLocator(integer=True))
            ax8.tick_params(axis="y", pad=4)
            ax8.set_xlim(0, max(valores_passos) * 1.12 if valores_passos else 1)
            for index, (valor, descricao, nome_completo, barra) in enumerate(zip(valores_passos, descricoes_passos, nomes_passos, ax8.patches)):
                ax8.text(valor + max(max(valores_passos) * 0.01, 0.2), index, str(valor), va="center", fontsize=8)
                self._register_hover_artist(
                    barra,
                    f"{nome_completo}\nDescricao: {descricao}\nOcorrencias: {valor}",
                )
        else:
            ax8.text(0.5, 0.5, "Nenhum passo critico no filtro atual", ha="center", va="center", transform=ax8.transAxes)
            ax8.set_title("Passos Criticos Mais Frequentes", fontsize=11, fontweight="bold", pad=12)

        self.canvas.draw_idle()

    def _clear_chart(self):
        self.figura.clear()
        self.canvas.draw_idle()

    def _build_export_dataframes(self, dados_para_exportar):
        resultados, dias = self._aggregate_filtered_data(dados_para_exportar)

        df_base = pd.DataFrame(dados_para_exportar).copy()
        if df_base.empty:
            return {}

        df_base = df_base[
            [
                "pr",
                "numero_serie",
                "operador",
                "maquina",
                "inicio",
                "fim",
                "resultado",
                "duracao_segundos",
                "passos_aprovados",
                "passos_pulados_fast",
                "retentativas_automaticas",
                "timeouts",
                "comandos_enviados",
                "falhas_modbus_sem_resposta",
                "usou_modo_fast",
                "erro_principal",
                "passos_reprovados",
                "detalhes_erro",
            ]
        ]
        df_base.rename(
            columns={
                "pr": "PR",
                "numero_serie": "Numero de Serie",
                "operador": "Operador",
                "maquina": "Maquina",
                "inicio": "Inicio",
                "fim": "Fim",
                "resultado": "Resultado",
                "duracao_segundos": "Duracao (minutos)",
                "passos_aprovados": "Passos Aprovados",
                "passos_pulados_fast": "Passos Pulados Fast",
                "retentativas_automaticas": "Retentativas Automaticas",
                "timeouts": "Timeouts",
                "comandos_enviados": "Comandos Enviados",
                "falhas_modbus_sem_resposta": "Falhas Modbus Sem Resposta",
                "usou_modo_fast": "Usou Modo Fast",
                "erro_principal": "Erro Principal",
                "passos_reprovados": "Passos Reprovados",
                "detalhes_erro": "Detalhes de Erro",
            },
            inplace=True,
        )
        df_base["Inicio"] = pd.to_datetime(df_base["Inicio"]).dt.strftime("%Y-%m-%d %H:%M:%S")
        df_base["Fim"] = pd.to_datetime(df_base["Fim"]).dt.strftime("%Y-%m-%d %H:%M:%S")
        df_base["Duracao (minutos)"] = (df_base["Duracao (minutos)"] / 60).round(2)
        df_base["Usou Modo Fast"] = df_base["Usou Modo Fast"].apply(lambda valor: "Sim" if valor else "Nao")
        df_base["Passos Reprovados"] = df_base["Passos Reprovados"].apply(self._format_failed_steps_for_excel)
        df_base["Detalhes de Erro"] = df_base["Detalhes de Erro"].apply(self._format_failed_steps_for_excel)

        df_kpis = pd.DataFrame(
            [
                {
                    "Indicador": item["title"],
                    "Valor": item["value"],
                    "Descricao": item["subtitle"],
                }
                for item in self._build_dashboard_metrics(resultados, dias, dados_para_exportar)
            ]
        )

        linhas_operadores = []
        for operador in sorted(resultados.keys()):
            dados = resultados[operador]
            linhas_operadores.append(
                {
                    "Operador": operador,
                    "Total": dados["total"],
                    "Aprovados": dados["aprovado"],
                    "Reprovados": dados["reprovado"],
                    "% Aprovacao": round(self._safe_percent(dados["aprovado"], dados["total"]), 2),
                    "Tempo Medio (min)": round((sum(dados["tempos"]) / len(dados["tempos"]) / 60) if dados["tempos"] else 0.0, 2),
                    "Passos Criticos": self._format_failed_steps_for_excel([texto for texto, _ in dados["passos"].most_common(5)]),
                }
            )
        df_operadores = pd.DataFrame(linhas_operadores)

        maquinas = Counter(item.get("maquina") or "Nao informado" for item in dados_para_exportar)
        maquinas_aprov = Counter(
            item.get("maquina") or "Nao informado"
            for item in dados_para_exportar
            if item["resultado"] == "APROVADO"
        )
        maquinas_repr = Counter(
            item.get("maquina") or "Nao informado"
            for item in dados_para_exportar
            if item["resultado"] == "REPROVADO"
        )
        df_maquinas = pd.DataFrame(
            [
                {
                    "Maquina": maquina,
                    "Total": total,
                    "Aprovados": maquinas_aprov.get(maquina, 0),
                    "Reprovados": maquinas_repr.get(maquina, 0),
                    "% Aprovacao": round(self._safe_percent(maquinas_aprov.get(maquina, 0), total), 2),
                }
                for maquina, total in maquinas.most_common()
            ]
        )

        erros = Counter(item.get("erro_principal") or "Sem erro detalhado" for item in dados_para_exportar if item.get("erro_principal"))
        prs_reprov = Counter(item.get("pr") or "-" for item in dados_para_exportar if item["resultado"] == "REPROVADO")
        df_falhas = pd.DataFrame(
            [
                {"Tipo": "Erro Principal", "Item": erro, "Ocorrencias": quantidade}
                for erro, quantidade in erros.most_common(20)
            ]
            + [
                {"Tipo": "PR Reprovado", "Item": pr, "Ocorrencias": quantidade}
                for pr, quantidade in prs_reprov.most_common(20)
            ]
        )

        return {
            "Base_Detalhada": df_base,
            "Resumo_KPIs": df_kpis,
            "Operadores": df_operadores,
            "Maquinas": df_maquinas,
            "Falhas": df_falhas,
        }

    def _write_dataframe_sheet(self, writer, sheet_name, df, header_format, even_row_format, odd_row_format):
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            row_format = even_row_format if row_num % 2 == 1 else odd_row_format
            for col_num, value in enumerate(row):
                worksheet.write(row_num, col_num, value, row_format)
        for col_num, coluna in enumerate(df.columns):
            comprimento = max(df[coluna].astype(str).map(len).max() if not df.empty else 0, len(coluna)) + 2
            worksheet.set_column(col_num, col_num, min(comprimento, 60))

    def _exportar_excel(self):
        try:
            dados_para_exportar = self._get_filtered_results()
            if not dados_para_exportar:
                QMessageBox.warning(self, "Sem Dados", "Nenhum dado disponivel para exportar com os filtros atuais.")
                return

            salvar_em, _ = QFileDialog.getSaveFileName(
                self,
                "Salvar Relatorio Excel",
                "relatorio_teste.xlsx",
                "Excel (*.xlsx)",
            )
            if not salvar_em:
                return

            planilhas = self._build_export_dataframes(dados_para_exportar)
            if not planilhas:
                QMessageBox.warning(self, "Sem Dados", "Nenhum dado disponivel para exportacao.")
                return

            with pd.ExcelWriter(salvar_em, engine="xlsxwriter") as writer:
                workbook = writer.book
                header_format = workbook.add_format(
                    {
                        "bold": True,
                        "text_wrap": True,
                        "valign": "vcenter",
                        "align": "center",
                        "fg_color": "#2C4B7A",
                        "font_color": "#FFFFFF",
                        "border": 1,
                    }
                )
                even_row_format = workbook.add_format({"fg_color": "#D7E4BC", "text_wrap": True})
                odd_row_format = workbook.add_format({"text_wrap": True})
                for nome_planilha, df_sheet in planilhas.items():
                    self._write_dataframe_sheet(
                        writer,
                        nome_planilha,
                        df_sheet,
                        header_format,
                        even_row_format,
                        odd_row_format,
                    )

            QMessageBox.information(self, "Sucesso", f"Relatorio exportado com sucesso:\n{salvar_em}")
        except Exception as exc:
            QMessageBox.critical(self, "Erro", f"Falha ao exportar:\n{exc}")

    def closeEvent(self, event):
        if self._analysis_worker is not None:
            try:
                self._analysis_worker.request_cancel()
            except Exception:
                pass
        if self._analysis_thread is not None:
            try:
                self._analysis_thread.quit()
            except Exception:
                pass
            try:
                self._analysis_thread.wait(5000)
            except Exception:
                pass
        self._close_progress_dialog()
        super().closeEvent(event)

    def _format_failed_steps_for_excel(self, passos):
        if not isinstance(passos, list):
            return ""
        limite = 5
        if len(passos) <= limite:
            return "\n".join(passos)
        return "\n".join(passos[:limite]) + f"\n... (+{len(passos) - limite} mais)"
