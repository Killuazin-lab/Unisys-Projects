from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox, QTableWidget, QTableWidgetItem, QSizePolicy
import sys
import pandas as pd
from datetime import datetime

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
                "background-color: #4CAF50; color: white; min-width: 320px;"
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
        self.carregar_dados()
        
    def filtrar_tabela(self):
        filtro = self.search_dropdown.currentText().strip().lower()
        
        if not filtro:
            print("Filtro vazio. Exibindo todos os dados.")
            self.carregar_dados()
            return

        print(f"Filtrando por: {filtro}")

        try:
            with pd.ExcelFile("CADASTRO FITAS RIO.xlsx", engine='openpyxl') as xls:
                df = pd.read_excel(xls, sheet_name=self.base_fitas)

            df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.strip()

            df_filtrado = df[df.iloc[:, 0].str.lower().str.contains(filtro, na=False)]

            if df_filtrado.empty:
                print("Nenhuma fita encontrada!")
                self.table.setRowCount(0)
                return

            self.table.blockSignals(True)
            self.table.clearContents()
            self.table.setRowCount(len(df_filtrado))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(df.columns)

            for row_idx, row in df_filtrado.iterrows():
                for col_idx, value in enumerate(row):
                    if pd.isna(value):
                        value = ""
                    self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))

            self.table.resizeColumnsToContents()
            self.table.blockSignals(False)

        except Exception as e:
            print(f"Erro ao filtrar os dados: {e}")

    def initUI(self):
        self.setWindowTitle(f"Controle de Fitas - {self.base_fitas}")
        self.setGeometry(100, 100, 1500, 680)
        
        layout = QVBoxLayout()
        search_layout = QHBoxLayout()
        
        search_layout.addStretch()
        self.search_dropdown = QComboBox()
        self.search_dropdown.setEditable(True)
        self.search_dropdown.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        search_layout.addWidget(QLabel("Buscar:"))
        search_layout.addWidget(self.search_dropdown)
        search_layout.addStretch()
        
        self.search_dropdown.currentTextChanged.connect(self.filtrar_tabela)

        self.table = QTableWidget()
        
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
    
    def carregar_dados(self):
        try:
            with pd.ExcelFile("CADASTRO FITAS RIO.xlsx", engine='openpyxl') as xls:
                df = pd.read_excel(xls, sheet_name=self.base_fitas)

            if 'DATA EXPIRACAO(IR)' in df.columns:
                df['DATA EXPIRACAO(IR)'] = pd.to_datetime(df['DATA EXPIRACAO(IR)'], errors='coerce').dt.strftime('%d/%m/%Y')
        
            df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.strip()
            
            self.table.setRowCount(len(df))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(df.columns)

            for row_idx, row in df.iterrows():
                for col_idx, value in enumerate(row):
                    self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(value) if pd.notna(value) else ""))

            fitas_legato = df.iloc[:, 0].astype(str).tolist()
            fitas_legato = [fita for fita in fitas_legato if fita.startswith("RO")]
            
            self.search_dropdown.blockSignals(True)  # 🚨 Bloqueia eventos
            self.search_dropdown.clear()
            self.search_dropdown.addItems(fitas_legato)
            self.search_dropdown.setCurrentText("")  # Mantém vazio ao iniciar
            self.search_dropdown.blockSignals(False)  #
            
            self.table.resizeColumnsToContents()

        except Exception as e:
            print(f"Erro ao carregar os dados: {e}")

    def voltar_tela_inicial(self):
        self.tela_inicial = TelaInicial()
        self.tela_inicial.show()
        self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TelaInicial()
    window.show()
    sys.exit(app.exec())
