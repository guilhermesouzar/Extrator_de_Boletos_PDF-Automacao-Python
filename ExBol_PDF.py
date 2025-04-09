import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import fitz
import pandas as pd
import os
import re

class ExtratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Boletos PDF")
        self.root.geometry("600x400")

        self.caminho_planilha = None
        self.pasta_destino = None

        # Botão: Selecionar Planilha
        self.btn_planilha = tk.Button(root, text="Selecionar Planilha", command=self.selecionar_planilha)
        self.btn_planilha.pack(pady=10)

        # Botão: Selecionar PDF
        self.btn_pdf = tk.Button(root, text="Selecionar PDF", command=self.selecionar_pdf)
        self.btn_pdf.pack(pady=10)

        # Botão: Selecionar Pasta de Destino
        self.btn_destino = tk.Button(root, text="Selecionar Pasta de Destino", command=self.selecionar_pasta_destino)
        self.btn_destino.pack(pady=10)

        # Área de mensagens
        self.text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=15)
        self.text_area.pack(pady=10)

    def log(self, msg):
        self.text_area.insert(tk.END, msg + "\n")
        self.text_area.see(tk.END)
        self.root.update()

    def selecionar_planilha(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if caminho:
            self.caminho_planilha = caminho
            self.log(f"Planilha selecionada: {caminho}")

    def selecionar_pasta_destino(self):
        caminho = filedialog.askdirectory()
        if caminho:
            self.pasta_destino = caminho
            self.log(f"Pasta de destino selecionada: {caminho}")

    def selecionar_pdf(self):
        if not self.caminho_planilha:
            messagebox.showerror("Erro", "Selecione uma planilha primeiro!")
            return
        if not self.pasta_destino:
            messagebox.showerror("Erro", "Selecione uma pasta de destino!")
            return

        pdf = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if pdf:
            self.log(f"PDF selecionado: {pdf}")
            self.extrair_dados(pdf)

    def extrair_dados(self, pdf_file_name):
        def extrair_nome_pagador(lines):
            if len(lines) >= 22:
                return lines[21].strip()
            return None

        def extrair_nome_aluno(lines):
            for linha in lines:
                if re.search(r"\b(MENSALIDADE |TRANSPORTE|MATERIAL)\b", linha, re.IGNORECASE):
                    match = re.search(r"^(.*?)\s*(MENSALIDADE |TRANSPORTE|MATERIAL)", linha, re.IGNORECASE)
                    if match:
                        return match.group(1).strip()
            return None

        try:
            df = pd.read_excel(self.caminho_planilha)
            if not all(col in df.columns for col in ['nomealuno', 'serieturma', 'nomerespfinanceiro']):
                raise ValueError("Colunas necessárias não encontradas.")
        except Exception as e:
            self.log(f"[ERRO] Não foi possível acessar a planilha: {e}")
            return

        doc = fitz.open(pdf_file_name)
        pasta_base = self.pasta_destino
        pasta_nao_encontrado = os.path.join(pasta_base, "Não encontrado")
        pasta_incorreto = os.path.join(pasta_base, "Pagador incorreto")
        os.makedirs(pasta_nao_encontrado, exist_ok=True)
        os.makedirs(pasta_incorreto, exist_ok=True)

        contagem_salvos = {}
        contagem_incorretos = 0

        for page_number in range(len(doc)):
            page = doc[page_number]
            lines = page.get_text("text").splitlines()

            nome_pagador = extrair_nome_pagador(lines)
            nome_aluno = extrair_nome_aluno(lines)

            if not nome_pagador or not nome_aluno:
                self.log(f"[ERRO] Dados faltando na página {page_number + 1}.")
                continue

            if "__" in nome_aluno:
                nome_aluno = nome_aluno.split("__", 1)[-1].strip()
            nome_aluno = re.sub(r'[^\w\s]', '', nome_aluno)[:50]

            registros = df[(df['nomerespfinanceiro'] == nome_pagador) & (df['nomealuno'] == nome_aluno)]
            if registros.empty:
                pasta_path = pasta_incorreto
                contagem_incorretos += 1
            else:
                turma = registros.iloc[0]['serieturma']
                pasta_path = os.path.join(pasta_base, turma)
                os.makedirs(pasta_path, exist_ok=True)
                contagem_salvos[turma] = contagem_salvos.get(turma, 0) + 1

            output_file = os.path.join(pasta_path, f"{nome_aluno}-{page_number + 1}.pdf")

            if not os.path.exists(output_file):
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=page_number, to_page=page_number)
                new_doc.save(output_file)
                new_doc.close()
                self.log(f"Página {page_number + 1} salva: {output_file}")
                df.loc[(df['nomerespfinanceiro'] == nome_pagador) & 
                       (df['nomealuno'] == nome_aluno), 'caminhoarquivo'] = output_file
            else:
                self.log(f"[AVISO] Já existe: {output_file}")

        doc.close()
        saida_excel = os.path.join(pasta_base, "saida.xlsx")
        df.to_excel(saida_excel, index=False)

        self.log("Processamento concluído.")
        self.log(f"Saída salva em: {saida_excel}")
        self.log("Resumo:")
        for turma, qtd in contagem_salvos.items():
            self.log(f"  - {turma}: {qtd}")
        self.log(f"Total incorretos: {contagem_incorretos}")

# Iniciar a interface
if __name__ == "__main__":
    root = tk.Tk()
    app = ExtratorGUI(root)
    root.mainloop()
