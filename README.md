# Preenchedor Automático de documentos de texto em Python.

Um programa simples para preencher documentos do Microsoft Word em lotes usando como base dados de planilhas do Excel.
Para rodar o programa é necessário verificar se as bibliotecas estão devidamente instaladas em seu ambiente python

Biblioteca usada para leitura de arquivos (.docx)
>pip install python-docx

Biblioteca usada para leitura de arquivos (.xlsx)
>pip install openpyxl

--------------------------------------------------------

# Preenchedor do Histórico Escolar

Descrição do código:

1. Importação de Bibliotecas:

- os: Para interagir com o sistema operacional, como manipulação de caminhos de arquivo;
- Document do módulo docx: Para criar, modificar e salvar documentos DOCX;
- load_workbook do módulo openpyxl: Para carregar dados de planilhas Excel;

2. Função carregar_dados(planilha):

- Essa função carrega os dados de uma planilha Excel especificada e os retorna como uma lista de listas, 
onde cada lista interna representa uma linha da planilha.

3. Função substituir_tags(doc_modelo, tags, data):

- Esta função substitui as tags dentro de um documento DOCX pelos dados fornecidos. As tags são especificadas como 
strings e são substituídas pelos valores correspondentes em data.

4. Função preencher_notas(...):

- Esta é a função principal que coordena todo o processo de preenchimento dos documentos.
- Primeiro, ela carrega os dados de várias planilhas Excel que contêm informações sobre alunos, componentes, 
competências, notas, conceitos e pareceres.
- Em seguida, itera sobre os dados dos alunos, preenchendo o modelo de documento DOCX com os dados específicos 
de cada aluno.
- Finalmente, salva o documento preenchido no diretório especificado.

5. Definição do diretório de saída dos arquivos (saida_notas).

6. Chamada da função preencher_notas(...) com os nomes das planilhas e o modelo de documento fornecidos como argumentos.

O script segue uma abordagem procedural e executa operações em uma sequência definida para realizar o preenchimento 
dos documentos.