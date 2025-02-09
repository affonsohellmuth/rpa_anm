import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QTextEdit, QHBoxLayout, QMessageBox
from PyQt5.QtGui import QFont, QIcon, QPixmap
from PyQt5.QtCore import Qt
from controle_execucao import iniciar_execucao, parar_execucao  # Importa as funções do controle de execução

class App(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Maiores Arrecadadores CFEM - ANM")
        self.setWindowIcon(QIcon('anm.ico'))  # Coloque o caminho do seu ícone
        self.setGeometry(100, 100, 500, 400)
        self.setStyleSheet("background-color: #f0f5ff;")
        self.setFixedSize(650, 450)  # Define largura e altura fixas
        # Fundo claro com tom azulado

        # Layout principal
        main_layout = QVBoxLayout(self)

        # Título com ícone
        title_layout = QHBoxLayout()
        title_label = QLabel("Maiores Arrecadadores CFEM - ANM")
        title_label.setFont(QFont("Arial", 14, QFont.Bold))
        title_label.setStyleSheet("color: #003366;")  # Azul escuro
        title_layout.setAlignment(Qt.AlignCenter)

        icon_label = QLabel()
        icon_pixmap = QPixmap('anm.ico').scaled(24, 24, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        icon_label.setPixmap(icon_pixmap)

        title_layout.addWidget(icon_label)
        title_layout.addWidget(title_label)
        main_layout.addLayout(title_layout)

        # Botões de controle
        buttons_layout = QHBoxLayout()

        self.nova_execucao_button = QPushButton("Iniciar Execução")
        self.nova_execucao_button.setStyleSheet("""
                    QPushButton {
                        background-color: #007acc; /* Azul médio */
                        color: white;
                        border-radius: 5px;
                        padding: 10px;
                        font-size: 14px;
                    }
                    QPushButton:hover {
                        background-color: #005fa3; /* Azul mais escuro no hover */
                    }
                """)

        self.nova_execucao_button.clicked.connect(self.iniciar_execucao)  # Conectando ao método iniciar_execucao

        self.parar_execucao_button = QPushButton("Parar Execução")
        self.parar_execucao_button.setStyleSheet("""
                    QPushButton {
                        background-color: #005fa3; /* Azul escuro */
                        color: white;
                        border-radius: 5px;
                        padding: 10px;
                        font-size: 14px;
                    }
                    QPushButton:hover {
                        background-color: #004080;
                    }
                """)

        self.parar_execucao_button.clicked.connect(self.parar_execucao)  # Conectando ao método parar_execucao

        buttons_layout.addWidget(self.nova_execucao_button)
        buttons_layout.addWidget(self.parar_execucao_button)
        main_layout.addLayout(buttons_layout)

        # Caixa de logs (reduzida)
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        self.log_text.setFixedHeight(265)  # Reduz o tamanho da área de logs
        self.log_text.setStyleSheet("""
                    QTextEdit {
                        background-color: white;
                        border: 1px solid #ccc;
                        padding: 5px;
                        font-size: 12px;
                    }
                """)
        main_layout.addWidget(self.log_text)

        # Instrução discreta
        instrucao_label = QLabel(
            "1. O robô será executado em segundo plano, sendo aberto um terminal que não deverá ser fechado.\n"
            "2. Se necessário encerrar o robô antes do término (em caso de erro e etc), clique em 'Parar Execução'. \n"
            "3. A planilha será salva no final da execução em C:\\Users\\seu_usuario\\Documents\\Planilha ANM'. \n"
            "4. Qualquer caso de melhoria, suporte ou incidente, entrar em contato com o setor de Desenvolvimento de RPA."
        )
        instrucao_label.setFont(QFont("Arial", 9))
        instrucao_label.setStyleSheet("color: #555;")  # Tom de cinza discreto
        instrucao_label.setAlignment(Qt.AlignLeft)
        main_layout.addWidget(instrucao_label)

        # Ajustes de estilo gerais (QSS)
        self.setStyleSheet("""
            QWidget {
                background-color: #f8f8f8;
            }
            QLabel {
                font-size: 14px;
                color: #333;
            }
            QStatusBar {
                background-color: #222;
                color: white;
            }
        """)


    def atualizar_log(self, mensagem):
        """Atualiza o log na interface."""
        self.log_text.append(mensagem)

    def iniciar_execucao(self):
        try:
            self.atualizar_log("Iniciando nova execução...")
            # Chama a função para iniciar a execução
            iniciar_execucao()
        except Exception as e:
            self.atualizar_log(f"Erro ao iniciar nova execução: {str(e)}")
            QMessageBox.critical(self, "Erro", f"Erro ao iniciar nova execução: {str(e)}")

    def parar_execucao(self):
        """Para a execução em andamento."""
        try:
            self.atualizar_log("Interrompendo execução...")
            parar_execucao()
            self.atualizar_log("Execução interrompida com sucesso!")
        except Exception as e:
            self.atualizar_log(f"Erro ao interromper execução: {str(e)}")
            QMessageBox.critical(self, "Erro", f"Erro ao interromper execução: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec_())
