# ExBolPDF

**Automatize a extração, renomeação e organização de boletos escolares a partir de arquivos PDF com centenas de páginas.**

## Sobre o projeto

O **ExBolPDF** foi criado para resolver um problema recorrente em secretarias escolares: extrair manualmente boletos de um único arquivo PDF com mais de 300 páginas,
renomeando cada página com base em informações contidas no conteúdo do PDF e organizando em pastas por turma.
Esse processo tomava cerca de **5 horas de trabalho manual** com possibilidade de erros e distrações, e com o ExBolPDF, foi reduzido para **apenas 15 segundos**.

## Funcionalidades

- Leitura automática de um arquivo PDF contendo boletos em massa;
- Extração do nome do aluno e do responsável financeiro diretamente do conteúdo da página PDF;
- Verificação das informações extraídas com uma planilha Excel fornecida pelo software de gestão escolar;
- Criação automática de pastas por turma;
- Renomeação e salvamento individual das páginas em suas respectivas pastas;
- Geração de um novo arquivo Excel com os caminhos dos arquivos salvos para utilização posterior.

## Interface Gráfica (GUI)

Desenvolvido com `tkinter`, o programa oferece uma interface simples e intuitiva para:

1. Selecionar a **pasta de destino**;
2. Selecionar a **planilha Excel** com os dados dos alunos;
3. Selecionar o **PDF** com os boletos.

## Estrutura esperada da planilha

A planilha deve conter pelo menos as seguintes colunas:

- `nomealuno`: nome do aluno conforme aparece no boleto;
- `serieturma`: nome da turma (usado para criar pastas);
- `nomerespfinanceiro`: nome do responsável financeiro (pagador).

> Exemplo de cabeçalho da planilha:

| nomealuno        | serieturma | nomerespfinanceiro |
|------------------|------------|---------------------|
| João da Silva    | 3º A       | Maria da Silva      |
| Ana Oliveira     | 2º B       | João Oliveira       |

## Requisitos

- Python 3.x
  
- tkinter: 
Biblioteca padrão do Python para criação de interfaces gráficas (GUI).
Usada para criar os botões, janelas e campos de texto da aplicação. Permite ao usuário selecionar arquivos e pastas pelo explorador do sistema.

- fitz (PyMuPDF): 
Biblioteca para manipulação de arquivos PDF.
Utilizada para abrir o PDF, extrair texto de cada página e salvar páginas individualmente.

- pandas: 
Biblioteca para manipulação e análise de dados em forma de tabelas (DataFrames).
Usada para ler a planilha Excel com os dados dos alunos e realizar filtros para identificar correspondências com os boletos.

- os: 
Biblioteca padrão do Python para manipulação de arquivos e diretórios.
Utilizada para criar pastas, verificar caminhos e salvar arquivos.

- re: 
Biblioteca padrão para expressões regulares.
Usada para extrair nomes específicos e limpar textos de forma mais precisa.

https://github.com/user-attachments/assets/e24d5f15-b245-4ec4-83ec-c74daed07ebf
