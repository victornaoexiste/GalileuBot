# GalileuBot

Automação em Python (Selenium) que cadastra o acervo de uma biblioteca escolar
no sistema Galileu, a partir de uma planilha. Feito para o Colégio Politécnico
Dom Luciano.

Projeto concluído: 194 de 194 títulos disponíveis no acervo.

## Como funciona

O cadastro no Galileu é lento e repetitivo quando feito à mão. O bot faz isso
em dois passos, lendo os dados de uma planilha (planilhaBase.xlsx):

- 1_cadastrar_livros.py: cadastra os títulos.
- 2_cadastrar_exemplares.py: cadastra os exemplares de cada título.

O login é feito manualmente pelo operador numa janela do Chrome aberta com
depuração remota; o script se conecta a essa sessão, então nenhuma senha fica
gravada no código.

## Requisitos

    pip install -r requirements.txt

(selenium, webdriver-manager, beautifulsoup4, openpyxl, requests, lxml)

## Uso

Veja o passo a passo completo em LEIA-ME.txt. Em resumo:

1. Abrir o Chrome com depuração remota e fazer login no Galileu.
2. Rodar 1_cadastrar_livros.py e depois 2_cadastrar_exemplares.py.

Os PDFs, a planilha e os logs não são versionados (ver .gitignore), por conterem
dados operacionais e arquivos grandes.
