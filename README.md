# Um-Organizador-de-Declara-o-de-Imposto-de-Renda
Um Organizador de Declaração de Imposto de Renda

link:
https://onedrive.live.com/:x:/g/personal/1b59a2b2e67d44ba/IQB9wz28tNhBTZhBjJ9j5nG9ARuMgLbjNYUyts79WSRiyk0?rtime=5h2iWrMV30g&nav=MTVfezAwMDAwMDAwLTAwMDEtMDAwMC0wMDAwLTAwMDAwMDAwMDAwMH0&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy8xYjU5YTJiMmU2N2Q0NGJhL0lRQjl3ejI4dE5oQlRaaEJqSjlqNW5HOUFSdU1nTGJqTllVeXRzNzlXU1JpeWswP2U9dEZTNVIzJm5hdj1NVFZmZXpBd01EQXdNREF3TFRBd01ERXRNREF3TUMwd01EQXdMVEF3TURBd01EQXdNREF3TUgw

# 📊 Organizador de Declaração de Imposto de Renda

## 📌 Sobre o projeto

Este projeto consiste em uma ferramenta desenvolvida em **Microsoft Excel** para auxiliar na organização de informações necessárias para a preparação da Declaração de Imposto de Renda.

A planilha foi estruturada para centralizar informações pessoais, informes de rendimentos bancários, entradas financeiras e tabelas auxiliares, tornando o processo de organização dos dados mais simples, padronizado e rastreável.

O projeto também demonstra a utilização do **Excel como ferramenta de organização e automação de processos**, aplicando fórmulas, validações e estruturas de dados para reduzir tarefas manuais.

> **Objetivo:** criar uma estrutura organizada para facilitar o levantamento e a conferência das informações financeiras utilizadas na preparação da declaração.

---

## 🎯 Objetivos do projeto

* Organizar dados pessoais do titular;
* Centralizar informações de contas e instituições financeiras;
* Registrar informes de rendimentos;
* Organizar entradas financeiras mês a mês;
* Manter documentos e comprovantes associados aos registros;
* Utilizar fórmulas para automatizar cálculos;
* Utilizar tabelas auxiliares para padronização dos dados;
* Reduzir erros durante a organização das informações;
* Criar uma estrutura reutilizável para futuros períodos de declaração.

---

## 🛠️ Tecnologias e ferramentas utilizadas

* **Microsoft Excel**
* Fórmulas e funções do Excel
* Validação de dados
* Tabelas auxiliares
* Organização estruturada de dados
* Formatação condicional e visual
* Referenciamento de documentos e comprovantes

---

## 📁 Estrutura da planilha

O arquivo está dividido em quatro abas principais:

### 1. `TITULAR`

A aba **TITULAR** concentra os dados cadastrais da pessoa física.

Entre as informações estruturadas estão:

* Nome;
* CPF;
* Data de nascimento;
* Título de eleitor;
* Cônjuge;
* Endereço;
* CEP;
* Telefone;
* Celular.

Essa separação permite manter os dados cadastrais centralizados e facilita sua utilização durante o processo de organização da declaração.

---

### 2. `INFORMES`

A aba **INFORMES** foi criada para organizar os informes de rendimentos relacionados às instituições financeiras.

Os registros podem conter:

* Banco;
* Valor atual;
* Documento ou informe relacionado;
* Total consolidado das instituições.

A planilha utiliza fórmulas para realizar a consolidação dos valores.

Exemplo de fórmula utilizada:

```excel
=SUM(D11,D16,D21)
```

Essa fórmula realiza a soma dos valores registrados nos diferentes bancos, permitindo obter um total consolidado.

---

### 3. `NOTAS`

A aba **NOTAS** é utilizada para registrar as entradas financeiras ao longo do período.

A estrutura possui campos como:

| Campo     | Descrição                   |
| --------- | --------------------------- |
| DATA      | Data da movimentação        |
| CATEGORIA | Classificação da entrada    |
| VALOR     | Valor financeiro registrado |

Exemplo de categoria utilizada:

```text
HOLERITE
```

Essa estrutura permite registrar diferentes entradas financeiras e posteriormente realizar análises e conferências.

---

### 4. `TABELAS`

A aba **TABELAS** funciona como uma base auxiliar para padronização das informações.

Ela contém uma relação de instituições financeiras e seus respectivos códigos, por exemplo:

```text
1 - Banco do Brasil
104 - Caixa Econômica Federal
184 - Banco Itaú BBA S.A.
208 - Banco BTG Pactual
237 - Banco Bradesco
```

Essa estrutura pode ser utilizada como fonte para listas suspensas e validações de dados dentro da planilha.

---

# ⚙️ Processo de desenvolvimento

O desenvolvimento da ferramenta foi realizado considerando a necessidade de transformar informações financeiras dispersas em uma estrutura organizada dentro do Excel.

## 1. Levantamento das informações

Primeiramente foram identificadas as principais categorias de informações necessárias para organização dos dados:

* Informações pessoais;
* Instituições financeiras;
* Valores financeiros;
* Entradas mensais;
* Documentos e comprovantes;
* Tabelas auxiliares.

---

## 2. Estruturação das abas

As informações foram separadas em diferentes abas para evitar a concentração de todos os dados em uma única tabela.

A divisão adotada foi:

```text
TITULAR
   ↓
INFORMES
   ↓
NOTAS
   ↓
TABELAS
```

Essa organização facilita a navegação e permite que cada categoria de informação tenha sua própria estrutura.

---

## 3. Aplicação de fórmulas

Foram utilizadas fórmulas do Excel para automatizar cálculos e reduzir operações manuais.

Um exemplo é o cálculo do total dos valores registrados na aba de informes:

```excel
=SUM(D11,D16,D21)
```

Dessa forma, quando os valores individuais são alterados, o total pode ser atualizado automaticamente.

---

## 4. Padronização dos dados

A utilização da aba `TABELAS` permite manter uma relação padronizada das instituições financeiras.

Esse modelo pode ser utilizado em conjunto com recursos como:

* Validação de dados;
* Listas suspensas;
* Referências entre células;
* Tabelas auxiliares.

O objetivo é diminuir a quantidade de informações digitadas manualmente e melhorar a consistência dos registros.

---

## 5. Organização dos documentos

A estrutura também permite relacionar os registros financeiros aos respectivos documentos.

Por exemplo:

```text
Banco → Valor → Informe/Documento
```

Essa abordagem facilita a localização dos comprovantes utilizados durante a conferência das informações.

---

# 🧮 Cálculos e automações

Diferentemente de uma calculadora financeira tradicional, esta ferramenta tem como foco principal a **organização e consolidação de informações**.

Entre as automações utilizadas estão:

### Totalização de valores

Os valores registrados nas instituições financeiras podem ser consolidados automaticamente utilizando funções de soma.

Exemplo:

```excel
=SUM(D11,D16,D21)
```

### Organização das entradas

Os registros da aba `NOTAS` seguem uma estrutura baseada em:

```text
Data + Categoria + Valor
```

Isso possibilita posteriormente utilizar funções de soma, filtros, tabelas dinâmicas ou gráficos para análise dos dados.

---

# 🔄 Fluxo de utilização

O fluxo sugerido para utilização da ferramenta é:

```text
1. Preencher dados do titular
           ↓
2. Registrar instituições financeiras
           ↓
3. Informar valores dos informes
           ↓
4. Associar documentos/comprovantes
           ↓
5. Registrar entradas financeiras
           ↓
6. Conferir os valores
           ↓
7. Utilizar as informações organizadas
```

---

# 📋 Exemplo de estrutura dos dados

### Dados do titular

```text
Nome
CPF
Nascimento
Título de eleitor
Cônjuge
Endereço
CEP
Telefone
Celular
```

### Informes bancários

```text
Banco
Valor atual
Documento
```

### Entradas financeiras

```text
Data
Categoria
Valor
```

---

# 📈 Possíveis melhorias futuras

O projeto pode ser expandido com novas funcionalidades, como:

* Dashboard financeiro;
* Gráficos de entradas por mês;
* Indicadores de valores por instituição;
* Controle de documentos pendentes;
* Status de conferência dos documentos;
* Automatização de relatórios;
* Filtros por período;
* Consolidação anual automática;
* Proteção de células com fórmulas;
* Melhorias de usabilidade e interface.

---

# 📚 Aprendizados técnicos

Este projeto permite demonstrar conhecimentos relacionados a:

* Desenvolvimento de planilhas profissionais;
* Organização e modelagem de dados;
* Excel intermediário/avançado;
* Utilização de fórmulas;
* Automatização de cálculos;
* Validação e padronização de informações;
* Estruturação de bases auxiliares;
* Documentação de processos técnicos.

---
