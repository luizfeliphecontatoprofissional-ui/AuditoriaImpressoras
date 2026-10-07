# Auditoria de Impressoras

> **Automação da conferência de contadores de impressão entre dados NDD e faturamento.**

## Sobre o projeto

O **Auditoria de Impressoras** é uma aplicação desenvolvida em Python para automatizar a conferência dos contadores de impressoras utilizados na fiscalização de contratos de impressão.

A proposta do projeto é comparar os dados brutos de leitura de contadores fornecidos pelo NDD com os dados do faturamento, identificando automaticamente diferenças que poderiam passar despercebidas em uma conferência manual.

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
| **4** | Contadores **Mono e Color** de `A4` no NDD × **Mono e Color** do faturamento |
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
| `Valor TP` | Valor encontrado no faturamento |

O relatório também apresenta no cabeçalho:

- período analisado;
- quantidade total de divergências;
- formatação para facilitar a leitura;
- filtro automático;
- congelamento do cabeçalho.

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


PyInstaller — previsto para a criação da versão executável

---

## Situação atual

O projeto já possui:

- automação das comparações dos tipos 1 a 5;
- geração do relatório em Excel;
- identificação automática do período do faturamento;
- interface gráfica;
- tratamento de erros;
- organização histórica dos relatórios;

### Próximas evoluções

Entre as melhorias planejadas para as próximas versões estão:

- validar automaticamente se os períodos do NDD e do faturamento são compatíveis;
- continuar refinando a interface gráfica;
- criar uma versão executável (`.exe`) para facilitar a distribuição e utilização.

---

## Autor

**Luiz Felipe de Andrade**

Projeto desenvolvido em Python com foco em automação, tratamento de dados, auditoria e geração de relatórios.

---

## Licença

Este projeto foi desenvolvido para profissionais relacionados ao processo de auditoria de dados de impressão.
