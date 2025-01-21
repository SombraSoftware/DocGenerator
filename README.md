DocGenerator

DescriçãoO DocGenerator é um programa em Python projetado para facilitar a criação de documentos personalizados, como certificados. Ele utiliza modelos de documentos Word (.docx) e dados fornecidos em uma planilha Excel para gerar documentos automaticamente, economizando tempo e reduzindo erros manuais.

Funcionalidades

Preenchimento Automático: Substitui TAGs predefinidas no modelo Word com os dados fornecidos na planilha.

Flexibilidade no Modelo: Aceita qualquer modelo de certificado no formato Word com TAGs predefinidas.

Geração em Massa: Processa vários registros em uma única execução.

Manutenção da Formatação: Preserva a formatação original do documento.

Fácil Configuração: Não requer conhecimentos avançados de programação para uso.

Requisitos do Sistema

Python: Versão 3.8 ou superior.

Bibliotecas Python:

openpyxl para manipulação de planilhas Excel.

python-docx para manipulação de documentos Word.

Outros:

Arquivo Excel com dados estruturados.

Modelo de documento Word contendo as TAGs predefinidas.

Exemplo de Uso

Prepare um modelo de certificado no Word com TAGs como {NOME_ALUNO}, {NASCIMENTO}, etc.

Crie uma planilha Excel contendo os dados correspondentes às TAGs.

Execute o script Python para gerar os certificados:

python src/preencher_certificados.py

Os documentos gerados serão salvos no diretório output/.

Estrutura do Projeto

data/: Contém arquivos de entrada, como o modelo de certificado e a planilha.

output/: Diretório para os arquivos gerados.

src/: Código-fonte do programa.

preencher_certificados.py: Script principal.

utils/: Funções auxiliares.

tests/: Casos de teste para validar o funcionamento do programa.

README.md: Documentação do projeto.

Licença

Este projeto é disponibilizado sob a MIT License, permitindo o uso, modificação e distribuição do software conforme os termos definidos.

Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests para melhorias no projeto.

