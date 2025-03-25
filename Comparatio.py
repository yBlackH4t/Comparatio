import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import string

class ComparadorExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Arquivos Excel")
        self.root.geometry("700x500")
        self.root.resizable(False, False)

        self.style = ttk.Style("darkly")

        self.arquivo1 = None
        self.arquivo2 = None

        frame = ttk.Frame(self.root, padding=20)
        frame.pack(expand=True, fill="both")

        self.btn_abrir_arquivo1 = ttk.Button(
            frame, 
            text="\U0001F4C2 Abrir Arquivo 1", 
            bootstyle="primary-outline", 
            command=self.abrir_arquivo1
        )
        self.btn_abrir_arquivo2 = ttk.Button(
            frame, 
            text="\U0001F4C2 Abrir Arquivo 2", 
            bootstyle="primary-outline", 
            command=self.abrir_arquivo2
        )
        self.btn_abrir_arquivo1.pack(pady=10)
        self.btn_abrir_arquivo2.pack(pady=10)

        self.lbl_coluna = ttk.Label(
            frame, 
            text="Selecione a coluna para comparação:", 
            font=("Verdana", 12), 
            bootstyle="primary"
        )
        self.lbl_coluna.pack(pady=10)

        self.coluna_combobox = ttk.StringVar()
        self.combobox_coluna = ttk.Combobox(
            frame, 
            textvariable=self.coluna_combobox, 
            values=list(string.ascii_uppercase[:8]), 
            state="readonly", 
            width=10
        )
        self.combobox_coluna.pack(pady=10)

        self.btn_comparar = ttk.Button(
            frame, 
            text="\U0001F50D Comparar Arquivos", 
            bootstyle="success-outline", 
            command=self.comparar_arquivos
        )
        self.btn_comparar.pack(pady=20)

        self.btn_limpar = ttk.Button(
            frame, 
            text="\U0001F5D1 Limpar Resultado", 
            bootstyle="danger-outline", 
            command=self.limpar_resultado
        )
        self.btn_limpar.pack(pady=10)

        self.resultado_texto = tk.Text(
            frame, 
            height=10, 
            width=60, 
            wrap=tk.WORD, 
            font=("Verdana", 10), 
            bg="#f5f5f5"
        )
        self.resultado_texto.pack(pady=10)

    def abrir_arquivo1(self):
        self.arquivo1 = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
        if self.arquivo1:
            messagebox.showinfo("Arquivo Selecionado", f"📂 Arquivo 1 carregado:\n{self.arquivo1}")

    def abrir_arquivo2(self):
        self.arquivo2 = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])
        if self.arquivo2:
            messagebox.showinfo("Arquivo Selecionado", f"📂 Arquivo 2 carregado:\n{self.arquivo2}")

    def coluna_para_indice(self, coluna):
        """Converte a letra da coluna do Excel para um índice numérico."""
        return string.ascii_uppercase.index(coluna)

    def comparar_arquivos(self):
        if not self.arquivo1 or not self.arquivo2:
            messagebox.showerror("Erro", "❌ Por favor, carregue ambos os arquivos antes de comparar.")
            return

        if not self.coluna_combobox.get():
            messagebox.showwarning("Atenção", "⚠️ Selecione uma coluna para comparação.")
            return

        try:
            col_idx = self.coluna_para_indice(self.coluna_combobox.get())

            df1 = pd.read_excel(self.arquivo1, usecols=[col_idx])
            df2 = pd.read_excel(self.arquivo2, usecols=[col_idx])

            df1.columns = ["QTD_1"]
            df2.columns = ["QTD_2"]

            df1["Linha"] = df1.index + 2
            df2["Linha"] = df2.index + 2

            df_comparado = pd.merge(df1, df2, on="Linha", how="outer")
            df_diferencas = df_comparado[df_comparado["QTD_1"] != df_comparado["QTD_2"]]

            # Filtra apenas as linhas com diferenças
            df_diferencas = df_diferencas.dropna(subset=["QTD_1", "QTD_2"])

            self.resultado_texto.delete(1.0, tk.END)
            if not df_diferencas.empty:
                resultado = df_diferencas.to_string(index=False, header=True)
                self.resultado_texto.insert(tk.END, f"⚠️ **Diferenças Encontradas**:\n\n{resultado}")
            else:
                self.resultado_texto.insert(tk.END, "✅ Os arquivos são idênticos!")

        except Exception as e:
            messagebox.showerror("Erro", f"⚠️ Ocorreu um erro ao comparar: {e}")

    def limpar_resultado(self):
        """Limpa o texto exibido no campo de resultado."""
        self.resultado_texto.delete(1.0, tk.END)

if __name__ == "__main__":
    root = ttk.Window(themename="darkly")
    app = ComparadorExcelApp(root)
    root.mainloop()
