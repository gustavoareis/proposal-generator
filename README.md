# Gerador de Propostas Comerciais a partir de Notas Fiscais

Script em Python que lê notas fiscais em PDF (NF-e de produto, NFS-e "antiga" e DANFSe v2.0), extrai os dados relevantes (número, data, cliente, CNPJ, itens e valores) e gera automaticamente propostas comerciais em Word (`.docx`), aplicando reajustes percentuais sobre os valores originais.

## O que o script faz

1. Lê todos os PDFs da pasta `notas_entrada/`.
2. Identifica o tipo de documento (NF-e de produto, NFS-e ou DANFSe) e extrai:
   - Número da nota
   - Data de emissão
   - Nome do cliente/tomador
   - CNPJ
   - Itens (descrição, unidade, quantidade, valor unitário e total)
   - Observações, quando existirem
3. Aplica um reajuste percentual sobre os valores dos itens, gerando dois conjuntos de valores diferentes (ver [Modelos gerados](#modelos-gerados)).
4. Preenche os templates Word (`template_lc.docx` / `template.docx` e `template_jw.docx`), substituindo texto e preenchendo a tabela de itens.
5. Salva o resultado em `propostas_saida/NF <numero>/`.

## Modelos gerados

Para cada nota processada, o script pode gerar até dois documentos, dependendo de quais templates existem na pasta do projeto:

| Template | Arquivo gerado | Reajuste aplicado |
|---|---|---|
| `template_lc.docx` (ou `template.docx` como alternativa) | `LC COMERCIAL NF <numero>.docx` | +3,0% fixo |
| `template_jw.docx` | `JW COMERCIAL NF <numero>.docx` | aleatório entre +5% e +8% |

- O reajuste é aplicado item a item com uma pequena variação aleatória entre eles, e o último item é ajustado para que o total bata exatamente com o valor-alvo (evita erros de arredondamento).
- No modelo "JW", a tabela de itens é formatada inteiramente em negrito; nos demais, apenas o cabeçalho.

## Tipos de nota suportados

- **NF-e de produto**: extração de itens via tabelas do PDF (pdfplumber) e dos campos de cabeçalho por regex.
- **NFS-e (modelo antigo, ex. Fortaleza)**: extração a partir do bloco "DISCRIMINAÇÃO DOS SERVIÇOS" e "DADOS DO TOMADOR DE SERVIÇOS".
- **DANFSe v2.0 (modelo nacional)**: extração a partir das seções "TOMADOR/ADQUIRENTE" e "SERVIÇO PRESTADO".

A escolha do parser é automática, com base em palavras-chave encontradas no texto do PDF (`DANFSE`, `NOTA FISCAL DE SERVIÇO`, `NFS-e`); caso nenhuma seja encontrada, assume-se nota de produto.

## Estrutura de pastas esperada

```
projeto/
├── main.py
├── template.docx          # ou template_lc.docx
├── template_jw.docx
├── notas_entrada/         # PDFs de entrada (criada automaticamente se não existir)
│   ├── nota1.pdf
│   └── nota2.pdf
└── propostas_saida/        # criada automaticamente
    └── NF 123/
        ├── LC COMERCIAL NF 123.docx
        └── JW COMERCIAL NF 123.docx
```

> Se a pasta `notas_entrada/` não existir, o script a cria e encerra, pedindo que os PDFs sejam adicionados antes de rodar novamente.

## Requisitos dos templates Word

Para que a substituição de texto funcione, os templates precisam conter parágrafos com os seguintes padrões (case-insensitive quanto ao conteúdo, mas o script verifica o texto em maiúsculas):

- Um parágrafo iniciando com `Fortaleza` seguido de `... DE ...` (linha de data/local) — substituído por `Fortaleza, <dia> DE <MÊS> DE <ano>`, com a data recuada em 3 dias úteis (considerando fins de semana e feriados do Ceará + aniversário de Fortaleza + Nossa Senhora).
- Um parágrafo contendo `CLIENTE:` — substituído pelo nome do cliente (e observação/CNPJ, no caso do template JW).
- Um parágrafo contendo `CNPJ:` (apenas nos templates que não são JW) — substituído pelo CNPJ extraído.
- No template JW: um parágrafo contendo `VALOR DA PROPOSTA`.
- Nos demais templates: um parágrafo contendo `VALOR TOTAL DA PROPOSTA:`.
- Uma tabela (a primeira do documento) com colunas identificáveis pelo cabeçalho (aceita variações como `ITEM`/`ORDEM`, `DESC`/`PRODUTO`/`SERVIÇO`, `UND`/`UNID`, `QTD`/`QUANT`, `UNIT`/`PR. UNIT`, `TOTAL`/`PR. TOTAL`).

O texto é substituído preservando a formatação do primeiro "run" do parágrafo. Toda a tabela é reformatada para a fonte **Mongolian Baiti, 9pt** (negrito no cabeçalho, ou em toda a tabela no caso do template JW).

## Dependências

```
python-docx
holidays
num2words
pdfplumber
```

Instale com:

```bash
pip install python-docx holidays num2words pdfplumber
```

## Como usar

1. Coloque os templates (`template.docx`/`template_lc.docx` e/ou `template_jw.docx`) na mesma pasta do script.
2. Coloque os PDFs das notas fiscais em `notas_entrada/`.
3. Execute:

```bash
python main.py
```

4. Os documentos gerados aparecerão em `propostas_saida/NF <numero>/`.

## Observações e limitações

- Os valores monetários são interpretados no formato brasileiro (`1.234,56`).
- Quando o pdfplumber mescla várias colunas numéricas em uma célula só, o script assume que o primeiro valor monetário encontrado é o correto.
- Se nenhum template (`template*.docx` ou `template_jw.docx`) for encontrado, nenhum documento é gerado para aquela nota — mas o script não avisa isso explicitamente, apenas não imprime a confirmação de criação.
- Os percentuais de reajuste (3% fixo e faixa de 5–8%) estão fixos no código; para alterá-los, edite os parâmetros passados para `reajustar_valores` em `processar_notas()`.
