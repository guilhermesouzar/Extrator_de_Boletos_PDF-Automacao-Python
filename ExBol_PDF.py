# -*- coding: utf-8 -*-

print("Extrator de boletos em PDF pelo nome.")
import fitz  # PyMuPDF
import re
import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd  # Para manipulação de planilhas

# Função para extrair o nome do pagador baseado na linha 22
def extrair_nome_pagador(lines):
    if len(lines) >= 22:
        linha_22 = lines[21]  # Índice 21 corresponde à linha 22
        return linha_22.strip()  # Retornar o texto completo da linha 22
    return None

# Função para extrair o nome do aluno baseado na primeira ocorrência de palavras-chave
def extrair_nome_aluno(lines):
    for linha in lines:
        # Verifica se a linha contém uma das palavras-chave
        if re.search(r"\b(MENSALIDADE |TRANSPORTE|MATERIAL)\b", linha, re.IGNORECASE):
            # Extrair o texto (nome do aluno) antes da palavra-chave encontrada
            match = re.search(r"^(.*?)\s*(MENSALIDADE |TRANSPORTE|MATERIAL)", linha, re.IGNORECASE)
            if match:
                return match.group(1).strip()  # Retorna o texto antes da palavra-chave
    return None

# Função para ler a planilha e retornar um DataFrame com os dados
def ler_planilha(caminho_planilha):
    try:
        df = pd.read_excel(caminho_planilha)
        # Verifica se as colunas necessárias existem
        required_columns = ['nomealuno', 'serieturma', 'nomerespfinanceiro']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"A coluna '{col}' não foi encontrada na planilha.")
        # Adiciona a coluna 'caminhoarquivo' se não existir
        if 'caminhoarquivo' not in df.columns:
            df['caminhoarquivo'] = None  # Inicializa a coluna com None
        return df  # Retorna o DataFrame completo
    except Exception as e:
        print(f"[ERRO] Não foi possível acessar a planilha: {e}")
        return None

# Função para reiniciar o processo
def reset_program():
    # Solicita ao usuário o caminho da planilha
    root = tk.Tk()
    root.withdraw()  # Ocultar a janela principal do tkinter
    print("Selecione a planilha:")
    caminho_planilha = filedialog.askopenfilename(
        title="Selecione a planilha",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )

    if not caminho_planilha:
        print("Nenhuma planilha selecionada. Encerrando o processo.")
        exit()

    # Ler os dados da planilha
    dados_planilha = ler_planilha(caminho_planilha)
    if dados_planilha is None:
        print("Erro ao ler a planilha. Encerrando o processo.")
        exit()

    # Criar a pasta "Não encontrado" se não existir
    pasta_nao_encontrado = os.path.join(os.getcwd(), "Não encontrado")
    if not os.path.exists(pasta_nao_encontrado):
        os.makedirs(pasta_nao_encontrado)

    # Criar a pasta "Pagador incorreto" se não existir
    pasta_pagador_incorreto = os.path.join(os.getcwd(), "Pagador incorreto")
    if not os.path.exists(pasta_pagador_incorreto):
        os.makedirs(pasta_pagador_incorreto)

    # Inicializar contadores
    contagem_salvos = {}
    contagem_incorretos = 0
    total_linhas_planilha = len(dados_planilha)
    total_linhas_analisadas = 0

    while True:
        # Abrir a janela de diálogo para selecionar o arquivo PDF
        print("Selecione o arquivo PDF:")
        pdf_file_name = filedialog.askopenfilename(
            title="Selecione o arquivo PDF",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )

        if not pdf_file_name:
            print("Nenhum arquivo selecionado. Encerrando o processo.")
            exit()

        # Abrir o PDF
        doc = fitz.open(pdf_file_name)

        # Processar cada página
        for page_number in range(len(doc)):
            page = doc[page_number]
            text = page.get_text("text")  # Extrair texto com estrutura em linhas

            # Dividir texto em linhas
            lines = text.splitlines()
            total_linhas_analisadas += 1  # Contar a linha analisada

            # Extrair o nome do pagador
            nome_pagador = extrair_nome_pagador(lines)

            if not nome_pagador:
                print(f"[ERRO] Nome do pagador não encontrado na página {page_number + 1}.")
                continue

            # Extrair o nome do aluno
            nome_aluno = extrair_nome_aluno(lines)

            if not nome_aluno:
                print(f"[ERRO] Nome do aluno não encontrado na página {page_number + 1}.")
                continue

            # Tratamento do nome do aluno
            if "__" in nome_aluno:
                nome_aluno = nome_aluno.split("__", 1)[-1].strip()  # Remove texto antes de "__"

            # Limitar o tamanho do nome e corrigir caracteres inválidos
            nome_aluno = re.sub(r'[^\w\s]', '', nome_aluno)
            nome_aluno = nome_aluno[:50]  # Limitar o tamanho do nome

            # Buscar os registros do pagador na planilha
            registros = dados_planilha[(dados_planilha['nomerespfinanceiro'] == nome_pagador) & (dados_planilha['nomealuno'] == nome_aluno)]

            if registros.empty:
                print(f"[AVISO] Nome do pagador '{nome_pagador}' e nome do aluno '{nome_aluno}' não encontrados na planilha.")
                # Salvar na pasta "Pagador incorreto"
                pasta_path = pasta_pagador_incorreto
                contagem_incorretos += 1
            else:
                # Usar a turma do primeiro registro
                turma = registros.iloc[0]['serieturma']
                pasta_path = os.path.join(os.getcwd(), turma)
                if not os.path.exists(pasta_path):
                    os.makedirs(pasta_path)

                # Contar arquivos salvos na pasta
                                # Contar arquivos salvos na pasta
                if turma not in contagem_salvos:
                    contagem_salvos[turma] = 0
                contagem_salvos[turma] += 1

            # Formatar o nome do arquivo
            output_file = os.path.join(pasta_path, f"{page_number + 1}-{nome_aluno}.pdf")

            # Verificar se o arquivo já existe para evitar duplicação
            if not os.path.exists(output_file):
                # Salvar a página como um novo PDF na pasta especificada
                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=page_number, to_page=page_number)
                new_doc.save(output_file)
                new_doc.close()
                print(f"Página {page_number + 1} salva como: {output_file}")

                # Salvar o caminho do arquivo na coluna 'caminhoarquivo'
                dados_planilha.loc[(dados_planilha['nomerespfinanceiro'] == nome_pagador) & 
                                   (dados_planilha['nomealuno'] == nome_aluno), 'caminhoarquivo'] = output_file
            else:
                print(f"[AVISO] O arquivo {output_file} já existe. Não foi salvo novamente.")

        # Finalizar o processamento
        doc.close()
        print("Processamento concluído. Os arquivos PDF foram salvos na pasta.")

        # Informar o caminho da pasta onde os arquivos foram salvos
        print(f"Os arquivos PDF foram salvos na pasta: {os.path.abspath(pasta_path)}")

        # Salvar o DataFrame atualizado de volta na planilha
        dados_planilha.to_excel(caminho_planilha, index=False)

        # Gerar log dos resultados
        log_file_path = os.path.join(os.getcwd(), "log_resultados.txt")
        with open(log_file_path, 'w') as log_file:
            log_file.write("Resultados do Processamento:\n")
            log_file.write(f"Total de linhas na planilha: {total_linhas_planilha}\n")
            log_file.write(f"Total de linhas analisadas: {total_linhas_analisadas}\n")
            log_file.write(f"Total de arquivos salvos:\n")
            for turma, quantidade in contagem_salvos.items():
                log_file.write(f"  - {turma}: {quantidade}\n")
            log_file.write(f"Total de arquivos com pagador incorreto: {contagem_incorretos}\n")

        # Exibir log no console
        print("Resultados do Processamento:")
        print(f"Total de linhas na planilha: {total_linhas_planilha}")
        print(f"Total de linhas analisadas: {total_linhas_analisadas}")
        print("Total de arquivos salvos:")
        for turma, quantidade in contagem_salvos.items():
            print(f"  - {turma}: {quantidade}")
        print(f"Total de arquivos com pagador incorreto: {contagem_incorretos}")

        # Perguntar ao usuário se deseja reiniciar o processo
        print("Deseja reiniciar o processo?")
        print("1: Reiniciar")
        print("2: Fechar")
        n = input("Digite 1 ou 2: ")
        if n == '2':
            break  # Encerra o programa
        elif n == '1':
            continue  # Reinicia o processo

# Chama a função para rodar o programa
reset_program()