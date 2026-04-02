from PyQt6.QtCore import QCoreApplication, QPoint, QTime, QTimer, Qt
from PyQt6.QtWidgets import QApplication, QHBoxLayout, QLabel, QPushButton, QWidget


class OledTimerDisplay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet(
            """
            background-color: #002244;
            color: #CCFF00;
            font: bold 20pt 'Courier New';
            border: 4px solid #555;
            border-radius: 15px;
            padding: 8px;
            """
        )

        layout = QHBoxLayout()
        layout.setContentsMargins(15, 5, 10, 5)

        self.label_tempo = QLabel("0:00")
        self.label_media = QLabel("Media 0:00")
        self.label_erros = QLabel("Erro 0/0")

        for label in (self.label_tempo, self.label_media, self.label_erros):
            label.setStyleSheet("color: #CCFF00;")
            layout.addWidget(label)

        self.btn_minimizar = QPushButton("X")
        self.btn_minimizar.setFixedSize(22, 22)
        self.btn_minimizar.setStyleSheet(
            """
            QPushButton {
                background: transparent;
                color: #CCFF00;
                font: bold 16pt 'Courier New';
                border: none;
                padding-bottom: 2px;
            }
            QPushButton:hover {
                color: #FF4444;
            }
            """
        )
        self.btn_minimizar.clicked.connect(self.minimizar_display)
        layout.addWidget(
            self.btn_minimizar,
            alignment=Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight,
        )

        self.setLayout(layout)
        self.resize(365, 55)

        self.tempo_inicial = None
        self.tempo_teste = 0
        self.congelado = True

        # Estatisticas consolidadas por numero de serie.
        # Repetir um teste da mesma placa atualiza os dados; nao cria uma placa nova.
        self.placas_stats = {}
        self.placa_em_teste = ""

        self.timer = QTimer(self)
        self.timer.timeout.connect(self._atualizar_display)
        self.timer.start(1000)

        self.anchor_widget = None
        self.hide()

    def _get_target_geometry(self, anchor_widget=None):
        widget = anchor_widget or self.anchor_widget
        try:
            if widget is not None:
                if hasattr(widget, "windowHandle") and widget.windowHandle() is not None:
                    screen = widget.windowHandle().screen()
                    if screen is not None:
                        return screen.availableGeometry()
                if hasattr(widget, "screen"):
                    screen = widget.screen()
                    if screen is not None:
                        return screen.availableGeometry()
                center = widget.mapToGlobal(widget.rect().center())
                screen = QApplication.screenAt(center)
                if screen is not None:
                    return screen.availableGeometry()
        except Exception:
            pass
        return QCoreApplication.instance().primaryScreen().availableGeometry()

    def _reposicionar_no_topo(self, anchor_widget=None):
        if anchor_widget is not None:
            self.anchor_widget = anchor_widget
        desktop = self._get_target_geometry(anchor_widget)
        x = desktop.left() + (desktop.width() - self.width()) // 2
        y = desktop.top() + 9
        self.move(QPoint(x, y))

    def iniciar_teste(self, numero_serie="", anchor_widget=None):
        self.tempo_inicial = QTime.currentTime()
        self.tempo_teste = 0
        self.placa_em_teste = (numero_serie or "").strip()
        self.congelado = False
        self._reposicionar_no_topo(anchor_widget)
        self._forcar_display()
        self.show()

    def finalizar_teste(self, duracao_segundos, houve_erro=False, contar=True, numero_serie=""):
        serial = (numero_serie or self.placa_em_teste or "").strip()
        if contar and serial:
            self.placas_stats[serial] = {
                "duracao_segundos": max(int(duracao_segundos), 0),
                "houve_erro": bool(houve_erro),
            }
        self.congelado = True
        self._forcar_display()

    def cancelar_teste(self):
        # Cancelamento nao conta placa, nao afeta media e nao afeta erros.
        self.placa_em_teste = ""
        self.congelado = True
        self._forcar_display()

    def _forcar_display(self):
        tempo_str = QTime(0, 0).addSecs(max(int(self.tempo_teste), 0)).toString("m:ss")
        media_str = QTime(0, 0).addSecs(self._calcular_media_segundos()).toString("m:ss")

        self.label_tempo.setText(tempo_str)
        self.label_media.setText(f"Media {media_str}")
        self.label_erros.setText(f"Erro {self._contar_erros()}/{self._contar_placas()}")

    def _atualizar_display(self):
        if self.congelado or not self.tempo_inicial:
            return

        self.tempo_teste = self.tempo_inicial.secsTo(QTime.currentTime())
        self._forcar_display()

    def _calcular_media_segundos(self):
        if not self.placas_stats:
            return 0
        total = sum(item["duracao_segundos"] for item in self.placas_stats.values())
        return int(total / len(self.placas_stats))

    def _contar_placas(self):
        return len(self.placas_stats)

    def _contar_erros(self):
        return sum(1 for item in self.placas_stats.values() if item.get("houve_erro"))

    def minimizar_display(self):
        self.hide()

    def restaurar_display(self, anchor_widget=None):
        self._reposicionar_no_topo(anchor_widget)
        self.show()
