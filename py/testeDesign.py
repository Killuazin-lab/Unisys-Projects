from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox, QTableWidget, QTableWidgetItem, QSizePolicy
import sys

class TelaInicial(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Seleção de Base de Fitas")
        self.setGeometry(100, 100, 600, 300)
        
        layout = QVBoxLayout()
        button_layout = QHBoxLayout()
        
        self.btn_legato = QPushButton("Base Fitas Legato")
        self.btn_mcp = QPushButton("Base Fitas MCP")
        
        for btn in [self.btn_legato, self.btn_mcp]:
            btn.setStyleSheet(
                "font-size: 18px; padding: 20px; border-radius: 15px; "
                "background-color: #4CAF50; color: white; min-width: 200px;"
            )
        
        self.btn_legato.clicked.connect(self.abrir_legato)
        self.btn_mcp.clicked.connect(self.abrir_mcp)
        
        button_layout.addStretch()
        button_layout.addWidget(self.btn_legato)
        button_layout.addSpacing(20)
        button_layout.addWidget(self.btn_mcp)
        button_layout.addStretch()
        
        layout.addStretch()
        layout.addLayout(button_layout)
        layout.addStretch()
        
        self.setLayout(layout)
    
    def abrir_legato(self):
        self.janela_fitas = ControleFitasApp("BASE FITAS LEGATO")
        self.janela_fitas.show()
        self.close()
    
    def abrir_mcp(self):
        self.janela_fitas = ControleFitasApp("BASE FITAS MCP")
        self.janela_fitas.show()
        self.close()

class ControleFitasApp(QWidget):
    def __init__(self, base_fitas):
        super().__init__()
        self.base_fitas = base_fitas
        self.initUI()

    def initUI(self):
        self.setWindowTitle(f"Controle de Fitas - {self.base_fitas}")
        self.setGeometry(100, 100, 1000, 600)
        
        layout = QVBoxLayout()
        search_layout = QHBoxLayout()
        
        search_layout.addStretch()
        self.search_dropdown = QComboBox()
        self.search_dropdown.setEditable(True)
        self.search_dropdown.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        search_layout.addWidget(QLabel("Buscar:"))
        search_layout.addWidget(self.search_dropdown)
        search_layout.addStretch()
        
        self.table = QTableWidget(6, 2)
        self.table.setHorizontalHeaderLabels(["Campo", "Valor"])
        self.table.setColumnWidth(1, 400)
        campos = ["Data Expiração", "Label", "Local Físico", "Status Exportação", "Local Digital", "Status Atual"]
        for row, campo in enumerate(campos):
            self.table.setItem(row, 0, QTableWidgetItem(campo))
        
        button_layout = QVBoxLayout()
        botoes = ["Alterar Data Expiração", "Alterar Label", "Alterar Local Físico", 
                  "Alterar Status Exportação", "Alterar Local Digital", "Alterar Status Atual"]
        for nome in botoes:
            btn = QPushButton(nome)
            button_layout.addWidget(btn)
        
        main_layout = QHBoxLayout()
        main_layout.addWidget(self.table)
        main_layout.addLayout(button_layout)
        
        self.btn_voltar = QPushButton("Voltar")
        self.btn_voltar.setStyleSheet(
            "font-size: 12px; padding: 8px; border-radius: 8px; "
            "background-color: #FF5733; color: white; min-width: 80px;"
        )
        self.btn_voltar.clicked.connect(self.voltar_tela_inicial)
        
        layout.addLayout(search_layout)
        layout.addLayout(main_layout)
        layout.addWidget(self.btn_voltar)
        
        self.setLayout(layout)
    
    def voltar_tela_inicial(self):
        self.tela_inicial = TelaInicial()
        self.tela_inicial.show()
        self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TelaInicial()
    window.show()
    sys.exit(app.exec())
