# Auditoria de Impressoras

> **Automação da conferência de contadores de impressão entre dados NDD e faturamento.**

## Sobre o projeto

O **Auditoria de Impressoras** é uma aplicação desenvolvida em Python para automatizar a conferência dos contadores de impressoras utilizados na fiscalização de contratos de impressão.

A proposta do projeto é comparar os dados brutos de leitura de contadores fornecidos pelo **NDD** com os dados de faturamento da **Tecprinters**, identificando automaticamente diferenças que poderiam passar despercebidas em uma conferência manual.

Além de realizar as comparações, a aplicação gera um **relatório em Excel com as divergências encontradas**, facilitando a análise pelo fiscal do contrato.

O sistema foi pensado para que a utilização não dependa de conhecimento de programação: os arquivos podem ser selecionados pela interface gráfica e o processo de auditoria é executado pelo próprio programa.

---

## Objetivo

O objetivo principal é **reduzir o trabalho manual na fiscalização dos contadores de impressoras**, aumentando a velocidade da conferência e diminuindo a possibilidade de erros durante a análise.

A aplicação transforma uma tarefa que depende de conferências repetitivas em um processo automatizado:

```text
Arquivo NDD
     +
Arquivo de Faturamento
     ↓
Leitura e organização dos dados
     ↓
Cruzamento por número de série
     ↓
Comparação dos contadores
     ↓
Identificação das divergências
     ↓
Relatório Excel
```

---

## Solução entregue

A aplicação permite:

- selecionar o arquivo **NDD** pela interface;
- selecionar o arquivo de **Faturamento** pela interface;
- identificar automaticamente o período presente no faturamento;
- cruzar os registros pelo **número de série** da impressora;
- comparar os contadores de acordo com o tipo de equipamento;
- registrar somente os campos que apresentam diferença;
- gerar um relatório Excel organizado;
- manter relatórios anteriores sem sobrescrevê-los;
- abrir o relatório gerado diretamente pela interface;
- apresentar mensagens de erro para situações como arquivo inexistente, dados fora do formato esperado ou impossibilidade de salvar o relatório.

---

## Comparações realizadas

As regras de comparação são baseadas no tipo da impressora informado no faturamento.

| Tipo | Comparação realizada |
|---|---|
| **1 e 2** | Contadores **A4** do NDD × contadores **Mono** do faturamento |
| **3** | Contadores **A3 e A4** do NDD × **Mono A3 e Mono** do faturamento |
| **4** | Contadores **Mono e Color** de `Print` no NDD × **Mono e Color** do faturamento |
| **5** | Contadores **A3/A4 + Mono/Color** do NDD × campos correspondentes do faturamento |

Quando existe diferença, ela é registrada individualmente no relatório.

---

## Relatório gerado

O resultado da auditoria é salvo em Excel com a aba **Divergências**.

A estrutura principal contém:

| Campo | Descrição |
|---|---|
| `SerialNumber` | Número de série da impressora |
| `Tipo/Item` | Tipo da impressora |
| `Campo Divergente` | Contador que apresentou diferença |
| `Valor NDD` | Valor encontrado no NDD |
| `Valor Tecprinters` | Valor encontrado no faturamento |

O relatório também apresenta no cabeçalho:

- período analisado;
- quantidade total de divergências;
- formatação para facilitar a leitura;
- filtro automático;
- congelamento do cabeçalho.

### Organização dos arquivos

Os relatórios não são mais sobrescritos.

O nome é gerado automaticamente usando o período analisado:

```text
relatorios/
├── Relatorio_Divergencias_30-06-2026_a_31-07-2026.xlsx
├── Relatorio_Divergencias_30-06-2026_a_31-07-2026_2.xlsx
└── Relatorio_Divergencias_31-07-2026_a_31-08-2026.xlsx
```

Isso permite manter o histórico das execuções.

---

## Interface gráfica

A aplicação possui uma interface gráfica desenvolvida com **Tkinter**.

O fluxo principal é:

```text
1. Selecionar arquivo NDD
2. Selecionar arquivo de Faturamento
3. Executar auditoria
4. Conferir o resultado
5. Abrir o relatório
```

A interface foi criada para esconder a complexidade do processamento e deixar a utilização mais simples para o usuário final.

---

## Tecnologias utilizadas

- **Python**
- **Pandas** — leitura, tratamento e comparação dos dados
- **OpenPyXL** — leitura e geração dos arquivos Excel
- **Tkinter** — interface gráfica
- **PyInstaller** — previsto para a criação da versão executável

---

## Estrutura do projeto

```text
Auditorialmpressores/
│
├── dados/
│   ├── NDD.xlsx
│   └── Faturamento.xlsx
│
├── relatorios/
│   └── Relatorio_Divergencias_*.xlsx
│
├── Auditorialmpressores.py
└── .gitignore
```

Os arquivos de dados utilizados na empresa não fazem parte do repositório público.

---

## Como utilizar

### 1. Preparar os arquivos

O arquivo NDD precisa conter os dados de leitura dos contadores, incluindo informações como:

```text
SerialNumber
CounterTypeDescription
StartCounterMono
EndCounterMono
StartCounterColor
EndCounterColor
StartCounterTotal
EndCounterTotal
```

O arquivo de faturamento precisa possuir os dados do dispositivo e os contadores de início e fim necessários para as comparações.

### 2. Executar a aplicação

Abra o projeto em um ambiente Python e execute o arquivo principal:

```bash
python Auditorialmpressores.py
```

### 3. Selecionar os arquivos

Na interface:

```text
Arquivo NDD
→ Procurar
→ selecionar o arquivo

Arquivo de Faturamento
→ Procurar
→ selecionar o arquivo
```

Depois, clique em:

```text
EXECUTAR AUDITORIA
```

### 4. Consultar o resultado

Ao finalizar, o programa gera o relatório na pasta:

```text
relatorios/
```

O botão **ABRIR RELATÓRIO** permite abrir diretamente o arquivo gerado na última execução.

---

## Tratamento de erros

A aplicação possui tratamento para alguns problemas comuns:

- arquivo não encontrado;
- arquivo ou relatório bloqueado por outro programa;
- coluna esperada ausente;
- dados em formato diferente do esperado;
- erros inesperados durante a execução.

Um exemplo prático é quando o relatório está aberto no Excel. Nesse caso, o programa informa que o arquivo não pôde ser salvo para que o usuário possa fechar a planilha e tentar novamente.

---

## Histórico de versões

### `v6.0` — Primeira versão com interface gráfica

Marco do projeto em que a auditoria passou a contar com uma interface gráfica para seleção dos arquivos e execução do processo.

O desenvolvimento continuou após esse marco com melhorias na organização dos relatórios e no fluxo da aplicação.

---

## Situação atual

O projeto já possui:

- automação das comparações dos tipos 1 a 5;
- geração do relatório em Excel;
- identificação automática do período do faturamento;
- interface gráfica;
- tratamento de erros;
- organização histórica dos relatórios;
- versionamento com Git e GitHub.

### Próximas evoluções

Entre as melhorias planejadas para as próximas versões estão:

- validar automaticamente se os períodos do NDD e do faturamento são compatíveis;
- continuar refinando a interface gráfica;
- melhorar a apresentação dos resultados;
- criar uma versão executável (`.exe`) para facilitar a distribuição e utilização.

---

## Por que este projeto foi criado?

A conferência de contadores de impressão pode envolver muitos dispositivos e vários tipos de contador. Quando essa verificação é feita manualmente, comparar linha por linha se torna uma tarefa repetitiva e sujeita a falhas.

A proposta do projeto é transformar essa conferência em uma rotina automatizada, mantendo o resultado **rastreável e fácil de revisar**.

Em vez de procurar manualmente por diferenças entre duas grandes planilhas:

```text
NDD ───────────────┐
                   ├──→ Auditoria automática ──→ Divergências
Faturamento ───────┘
```

O fiscal recebe um relatório já concentrado nos pontos que precisam ser analisados.

---

## Autor

**Luiz Felipe de Andrade**

Projeto desenvolvido em Python com foco em automação, tratamento de dados, auditoria e geração de relatórios.

---

## Licença

Este projeto foi desenvolvido para fins acadêmicos e profissionais relacionados ao processo de auditoria de dados de impressão.
