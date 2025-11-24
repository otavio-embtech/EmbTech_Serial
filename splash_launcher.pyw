import sys
import os
from PyQt6.QtWidgets import QApplication, QWidget, QLabel
from PyQt6.QtGui import QPixmap
from PyQt6.QtCore import Qt, QTimer

def main():
    # Adiciona o diretório atual ao sys.path para garantir que os módulos sejam encontrados
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller cria um atributo _MEIPASS que aponta para o diretório temporário
        sys.path.insert(0, sys._MEIPASS)
    else:
        sys.path.insert(0, os.path.dirname(__file__))

    # Desativa a criação de arquivos .pyc para implantações mais limpas
    sys.dont_write_bytecode = True

    app = QApplication(sys.argv)

    # Caminho absoluto para splash.png
    caminho_imagem = os.path.join(os.path.dirname(__file__), "assets/splash.png")
    splash_pix = QPixmap(caminho_imagem)

    if splash_pix.isNull():
        print("Erro: a imagem não foi carregada. Verifique o caminho e o nome do arquivo.")
        return

    # Usar QWidget personalizado com QLabel para controle completo
    splash = QWidget(flags=Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
    splash.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
    splash.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)
    splash.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

    label = QLabel(splash)
    label.setPixmap(splash_pix)
    label.setAlignment(Qt.AlignmentFlag.AlignCenter)

    splash.resize(splash_pix.size())
    label.resize(splash_pix.size())
    splash.show()
    app.processEvents()

    def carregar_app():
        try:
            from EmbTech_Serial import PlacaTesterApp
            window = PlacaTesterApp()
            window.show()
            splash.close()
        except Exception as e:
            import traceback
            print("Erro ao iniciar a aplicação:", e)
            traceback.print_exc()
            splash.close()

    QTimer.singleShot(500, carregar_app)

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
