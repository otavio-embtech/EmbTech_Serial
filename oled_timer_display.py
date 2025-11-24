from PyQt6.QtWidgets import QWidget, QLabel, QHBoxLayout, QPushButton
from PyQt6.QtCore import Qt, QTimer, QTime, QPoint, QCoreApplication

class OledTimerDisplay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet("""
            background-color: #002244;
            color: #CCFF00;
            font: bold 20pt 'Courier New';
            border: 4px solid #555;
            border-radius: 15px;
            padding: 8px;
        """)
        layout = QHBoxLayout()
        layout.setContentsMargins(15, 5, 10, 5)  # margem direita menor

        self.label_tempo = QLabel("0:00")
        self.label_media = QLabel("Média 0:00")
        self.label_erros = QLabel("Erro 0/0")

        for label in [self.label_tempo, self.label_media, self.label_erros]:
            label.setStyleSheet("color: #CCFF00;")
            layout.addWidget(label)

        # Botão X para minimizar
        self.btn_minimizar = QPushButton("✕")
        self.btn_minimizar.setFixedSize(22, 22)  # menor e mais centralizado
        self.btn_minimizar.setStyleSheet("""
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
        """)
        self.btn_minimizar.clicked.connect(self.minimizar_display)
        layout.addWidget(self.btn_minimizar, alignment=Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight)

        self.setLayout(layout)
        self.resize(365, 55)  # aumente um pouco para caber o botão

        self.tempo_inicial = None
        self.tempo_teste = 0
        self.erros_total = 0
        self.testes_total = 0
        self.medias = []
        self.congelado = True

        self.timer = QTimer()
        self.timer.timeout.connect(self._atualizar_display)
        self.timer.start(1000)

        self.hide()  # Só aparece ao iniciar teste

    def _reposicionar_no_topo(self):
        desktop = QCoreApplication.instance().primaryScreen().availableGeometry()
        x = (desktop.width() - self.width()) // 2
        y = 9
        self.move(QPoint(x, y))

    def iniciar_teste(self):
        self.tempo_inicial = QTime.currentTime()
        self.tempo_teste = 0
        self.congelado = False
        self._reposicionar_no_topo()
        self.show()

    def finalizar_teste(self, duracao_segundos, houve_erro=False, contar=True):
        self.medias.append(duracao_segundos)
        if contar:
            self.testes_total += 1
            if houve_erro:
                self.erros_total += 1
        self.congelado = True
        self._forcar_display()
    
    def _forcar_display(self):
        tempo_str = QTime(0, 0).addSecs(self.tempo_teste).toString("m:ss")
        media_str = QTime(0, 0).addSecs(int(sum(self.medias) / len(self.medias))) if self.medias else QTime(0, 0)
        media_str = media_str.toString("m:ss")

        self.label_tempo.setText(f"{tempo_str}")
        self.label_media.setText(f"Média {media_str}")
        self.label_erros.setText(f"Erro {self.erros_total}/{self.testes_total}")

    def cancelar_teste(self):
        self.testes_total += 1
        self.congelado = True

    def _atualizar_display(self):
        if self.congelado or not self.tempo_inicial:
            return

        self.tempo_teste = self.tempo_inicial.secsTo(QTime.currentTime())
        tempo_str = QTime(0, 0).addSecs(self.tempo_teste).toString("m:ss")
        media_str = QTime(0, 0).addSecs(int(sum(self.medias) / len(self.medias))) if self.medias else QTime(0, 0)
        media_str = media_str.toString("m:ss")

        self.label_tempo.setText(f"{tempo_str}")
        self.label_media.setText(f"Média {media_str}")
        self.label_erros.setText(f"Erro {self.erros_total}/{self.testes_total}")

    def minimizar_display(self):
        """Oculta o display sem perder os dados."""
        self.hide()

    def restaurar_display(self):
        """Mostra novamente o display."""
        self.show()
        