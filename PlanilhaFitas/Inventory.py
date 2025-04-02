import os
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


OPCOES_LOCALIDADE = ["IRON", "EMBRATEL"]
OPCOES_STATUS_EXP = ["EXPIRADA", "CUMPRINDO RETENÇÃO", "FITA S/ CONTEÚDO"]
OPCOES_MAIN_LEGATO = ["MAIN", "LEGATO"]
OPCOES_STATUS = [
    "EM USO", "RETIDA NA IRON", "FITA NA MALETA", "RETIDA NO ROBO DIARIA",
    "RETIDA NO ROBO SEMANAL", "LIBERADA P/DIARIA", "LIBERADA P/SEMANAL"
]

COLUNAS_EDITAVEIS = {
    "LOCALIDADE": OPCOES_LOCALIDADE,
    "STATUS/EXP": OPCOES_STATUS_EXP,
    "MAIN OU LEGATO": OPCOES_MAIN_LEGATO,
    "STATUS": OPCOES_STATUS
}

class TelaInicial(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Seleção de Base de Fitas")
        self.geometry("600x300")
        
        frame = tk.Frame(self)
        frame.pack(expand=True)
        
        self.btn_legato = tk.Button(frame, text="Base Fitas Legato", font=("Arial", 14), bg="#4CAF50", fg="white", 
                            padx=20, pady=10, width=20, command=lambda: self.abrir_janela("LEGATO"))
        self.btn_legato.pack(pady=10)

        self.btn_mcp = tk.Button(frame, text="Base Fitas MCP", font=("Arial", 14), bg="#4CAF50", fg="white", 
                          padx=20, pady=10, width=20, command=lambda: self.abrir_janela("MCP"))
        self.btn_mcp.pack(pady=10)

        
    def abrir_janela(self, base_fitas):
        self.withdraw()
        ControleFitasApp(self, base_fitas)

class ControleFitasApp(tk.Toplevel):
    def __init__(self, master, base_fitas):
        super().__init__(master)
        self.master = master
        self.base_fitas = base_fitas
        self.title(f"Controle de Fitas - {base_fitas}")
        self.geometry("1200x600")

        frame_filtro = tk.Frame(self)
        frame_filtro.pack(pady=10, fill=tk.X, padx=20)

        frame_busca_serial = tk.Frame(frame_filtro)
        frame_busca_serial.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(frame_busca_serial, text="Busca por serial:", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        self.entry_busca_serial = tk.Entry(frame_busca_serial, font=("Arial", 12))
        self.entry_busca_serial.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.btn_buscar_serial = tk.Button(frame_busca_serial, text="Buscar", command=self.buscar_serial, bg="#4CAF50", fg="white")
        self.btn_buscar_serial.pack(side=tk.LEFT, padx=5)

        frame_filtro_opcoes = tk.Frame(frame_filtro)
        frame_filtro_opcoes.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(frame_filtro_opcoes, text="Filtrar por:", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)

        self.combo_colunas = ttk.Combobox(frame_filtro_opcoes, state="readonly")
        self.combo_colunas.pack(side=tk.LEFT, padx=5)

        self.entry_filtro = tk.Entry(frame_filtro_opcoes, font=("Arial", 12))
        self.entry_filtro.pack(side=tk.LEFT, padx=5)

        self.btn_filtrar = tk.Button(frame_filtro_opcoes, text="Filtrar", command=self.filtrar_tabela, bg="#4CAF50", fg="white")
        self.btn_filtrar.pack(side=tk.LEFT, padx=5)

        self.btn_limpar = tk.Button(frame_filtro_opcoes, text="Limpar", command=self.limpar_filtro, bg="#FF5733", fg="white")
        self.btn_limpar.pack(side=tk.LEFT, padx=5)

        frame_tabela = tk.Frame(self)
        frame_tabela.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        self.scroll_y = tk.Scrollbar(frame_tabela, orient=tk.VERTICAL)
        self.scroll_x = tk.Scrollbar(frame_tabela, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(frame_tabela, yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)

        self.scroll_y.config(command=self.tree.yview)
        self.scroll_x.config(command=self.tree.xview)

        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(expand=True, fill=tk.BOTH)

        self.btn_salvar = tk.Button(self, text="Salvar Alterações", command=self.salvar_alteracoes, bg="#4CAF50", fg="white")
        self.btn_salvar.pack(pady=5)

        self.btn_voltar = tk.Button(self, text="Voltar", bg="#FF5733", fg="white", command=self.voltar_tela_inicial)
        self.btn_voltar.pack(pady=5)

        self.btn_editar = tk.Button(self, text="Editar Selecionado", command=self.abrir_tela_edicao, bg="#FFC107", fg="black")
        self.btn_editar.pack(pady=5)

        self.carregar_dados()

    def carregar_dados(self):
        if self.base_fitas == "MCP":
            excel_file = "CADASTROFITASRIOMCP.xlsx"
            sheet_name = "BASE FITAS MCP"
        else:
            excel_file = "CADASTROFITASRIOLEGATO.xlsx"
            sheet_name = "BASE FITAS LEGATO"

        if not os.path.exists(excel_file):
            messagebox.showerror("Erro", f"Arquivo não encontrado: {excel_file}")
            self.voltar_tela_inicial()
            return

        try:
            self.df_original = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
            
            # Verificar e formatar a coluna 'DATA EXPIRACAO(IR)'
            if 'DATA EXPIRACAO(IR)' in self.df_original.columns:
                self.df_original['DATA EXPIRACAO(IR)'] = pd.to_datetime(self.df_original['DATA EXPIRACAO(IR)'], errors='coerce').dt.strftime('%d/%m/%Y')

            self.tree["columns"] = list(self.df_original.columns)
            self.combo_colunas["values"] = list(self.df_original.columns) 
            self.tree.heading("#0", text="#")
            for col in self.df_original.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=150, anchor="center")

            self.atualizar_tabela(self.df_original)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados:\n{e}")
            self.voltar_tela_inicial()


    def atualizar_tabela(self, df):
        """Atualiza a TreeView com os dados do DataFrame."""
        self.tree.delete(*self.tree.get_children()) 
        for i, row in df.iterrows():
            self.tree.insert("", "end", iid=i, text=str(i), values=list(row))

    def filtrar_tabela(self):
        """Filtra os dados baseado na coluna e no valor fornecido."""
        coluna = self.combo_colunas.get().strip()
        filtro = self.entry_filtro.get().strip().lower()

        if not coluna or not filtro:
            messagebox.showwarning("Aviso", "Selecione uma coluna e digite um valor para filtrar.")
            return
        df_filtrado = self.df_original[self.df_original[coluna].astype(str).str.lower().str.contains(filtro, na=False)]
        self.atualizar_tabela(df_filtrado)

    def limpar_filtro(self):
        """Restaura a tabela para exibir todos os dados."""
        self.entry_filtro.delete(0, tk.END)
        self.atualizar_tabela(self.df_original)

    def buscar_serial(self):
        """Busca por serial na coluna 'SERIAL'."""
        serial = self.entry_busca_serial.get().strip().lower()

        if not serial:
            messagebox.showwarning("Aviso", "Digite um serial para buscar.")
            return

        df_filtrado = self.df_original[self.df_original["SERIAL DA FITA"].astype(str).str.lower().str.contains(serial, na=False)]
        self.atualizar_tabela(df_filtrado)

    def abrir_tela_edicao(self):
        """Abre uma nova tela para editar a célula selecionada."""
        item = self.tree.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecione um item para editar.")
            return

        item = item[0]
        coluna_nome = self.tree["columns"]
        valores = self.tree.item(item, "values")

        popup = tk.Toplevel(self)
        popup.title("Editar Valores")
        popup.geometry("400x400")

        tk.Label(popup, text="Editar Valores", font=("Arial", 14)).pack(pady=10)

        entries = {}
        campos_editaveis = ["STATUS/EXP", "STATUS", "LOCALIDADE", "LABEL", "DATA EXPIRAÇÃO"]
        for i, col in enumerate(coluna_nome):
            if col in campos_editaveis:
                tk.Label(popup, text=col).pack()
                if col in COLUNAS_EDITAVEIS:
                    combo = ttk.Combobox(popup, values=COLUNAS_EDITAVEIS[col], state="readonly")
                    combo.set(valores[i])
                    combo.pack(pady=5)
                    entries[col] = combo
                else:
                    entry = tk.Entry(popup)
                    entry.insert(0, valores[i])
                    entry.pack(pady=5)
                    entries[col] = entry

        def salvar():
            for col, widget in entries.items():
                self.tree.set(item, column=col, value=widget.get())
            popup.destroy()

        tk.Button(popup, text="Salvar", command=salvar, bg="#4CAF50", fg="white").pack(pady=10)

    def salvar_alteracoes(self):
        """Salva as alterações no arquivo Excel."""
        for item in self.tree.get_children():
            valores = self.tree.item(item, "values")
            self.df_original.iloc[int(item)] = list(valores)

        self.df_original.to_excel(EXCEL_FILE, sheet_name=self.base_fitas, index=False, engine='openpyxl')
        messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!")

    def voltar_tela_inicial(self):
        self.master.deiconify()
        self.destroy()

if __name__ == "__main__":
    app = TelaInicial()
    app.mainloop()
