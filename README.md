# GVBUS Comparator — Líder Limpe

App Streamlit que compara o **TXT comercial** (matrícula;nome;valor;obs) com o
**relatório de Saldo Estimado dos Cartões GVBUS** (PDF, XLSX, XLS ou CSV) e
gera automaticamente o TXT final com os valores a depositar:

| Caso | Regra |
|------|------|
| Saldo do cartão **≥** valor do TXT | resultado fica `0,00` |
| Saldo do cartão **<** valor do TXT | deposita só a **diferença** (complemento) |
| Matrícula **não encontrada** na planilha | mantém o valor cheio |
| Linha do TXT já estava em `0,00` (férias/afastamento) | preserva como está |

Negativos nunca aparecem no TXT final.

> Toda observação que vinha após o valor (**ATS**, **FLT**, **FERIAS …**, etc.) é **preservada**
> no TXT final.

## ✨ Funcionalidades

- 📤 Upload de TXT + relatório de saldo nos formatos **PDF / XLSX / XLS / CSV**
  (o PDF é o exportado direto do sistema GVBUS — formato preferido)
- 🔍 Detecção automática das colunas `Matrícula`, `Nome`, `Saldo` em qualquer formato
- 🧠 **Validação de matrículas divergentes**: se um nome no TXT bate com
  um nome na planilha mas a matrícula é diferente, o app exibe os pares
  lado a lado e você confirma com um checkbox se é a mesma pessoa;
  a correção é aplicada automaticamente no TXT final.
- 📊 Métricas (total a depositar, complementos, zerados, sem saldo, corrigidos)
- 🔎 Filtros por status e busca por nome/matrícula
- ⬇️ Download do **TXT de depósito** + relatório CSV detalhado para conferência

## 🚀 Como rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Deploy no Streamlit Community Cloud

1. Suba este repositório no GitHub.
2. Em [share.streamlit.io](https://share.streamlit.io) → **New app** →
   selecione o repo e o branch e aponte para `app.py`.
3. Pronto: o `requirements.txt` é instalado automaticamente.

## 📂 Estrutura

```
.
├── app.py                 # interface Streamlit
├── requirements.txt
├── README.md
├── .streamlit/
│   └── config.toml        # tema (cores Líder Limpe)
├── assets/
│   └── logo.png           # logo da empresa (exibido no header)
└── utils/
    ├── parser.py          # parser do TXT e da planilha
    └── comparator.py      # regra de negócio + detecção de conflitos
```

## 🧪 Formato esperado do TXT

```
46149;ADELSON COELHO PINHEIRO;81,60;
226;ADENILDO LEMOS RIBEIRO;142,80;ATS
4396;AILTON MARTINS DA SILVA;0,00;FLT MES TD
4675;ALMERITA PEREIRA DA SILVA;0,00;FERIAS 04/05 A 02/06
```

Separador: `;` · Casas decimais: `,` · Encoding: UTF-8 ou Latin-1 (auto-detectado).

## 🧪 Formato esperado do relatório de saldo

O app aceita **PDF** (preferido), **XLSX**, **XLS** ou **CSV** — detecta
automaticamente o layout.

### 📄 PDF (recomendado)

Relatório direto do sistema do cartão GVBUS
(*"Consulta do Saldo Estimado dos Cartões"*). Cada linha de dados tem o
formato:

```
06850000773592 ADELSON COELHO PINHEIRO 46149 VT Funcionário Ativo 29,60
   cartão            funcionário        matr.  tipo   status   saldo
```

O parser ignora automaticamente:
- cabeçalho de cada página (`Consulta`, `Hora:`, `Titular:`, `CNPJ:`, etc.),
- rodapé (`Página: N`, `Total de Cartões: N`),
- cartões bloqueados que não têm matrícula associada.

### 📊 XLSX (planilha)

A planilha exportada pelo sistema do cartão tem este layout (o app
**detecta automaticamente**, você não precisa preparar nada):

```
L1  | CONSULTA DO SALDO ESTIMADO DOS CARTÕES                       ← título
L2  | Titular: LIDER LIMPE LIMPEZA COMERCIAL    Data: 2025-05-18    ← metadata
L3  |                                            Hora: 15:12:10
L4  | (vazia)
L5  | Cartão | Funcionário | Matrícula | Tipo | Status | Saldo       ← cabeçalho real
L6+ |   06...| ROSANGELA  |    119    |  VT  | Ativo  |  0          ← dados
... | ...
LN  | Total de Cartões: 2413                                       ← rodapé ignorado
```

O app procura por:
- linha que contenha **'Matrícula' E 'Saldo'** (em qualquer das 30 primeiras linhas),
- ignora qualquer texto antes do cabeçalho,
- ignora qualquer linha sem matrícula válida (cartões bloqueados),
- ignora o rodapé `Total de Cartões: ...`,
- aceita saldos no formato BR (`96,90`) ou US (`96.9`).

### 💡 Dica para o `.xls` antigo do sistema do cartão

Se você exportar como **“Excel 97-2003 — Página da Web”**, o Excel gera um
`Arquivo.xls` **e uma pasta `Arquivo_arquivos/`**. Sem essa pasta o `.xls`
fica vazio. **Solução simples**: abra no Excel e **Salvar como → `.xlsx`**
antes de enviar para o app.

---

© Líder Limpe — Limpeza e Conservação
