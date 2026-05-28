# Gestão de Manifestações NF-e

App Streamlit para cruzamento fiscal SAP × Qive e controle de manifestações de destinatário.

---

## Estrutura do projeto

```
manifestacao/
├── app.py                    ← aplicativo principal
├── requirements.txt          ← dependências Python
├── .gitignore                ← protege o arquivo de secrets
├── .streamlit/
│   └── secrets.toml          ← configuração da planilha (NÃO subir para Git)
└── README.md
```

---

## Pré-requisitos

- Python 3.10 ou superior
- VSCode com a extensão **Python** instalada (Microsoft)

---

## Passo a passo para rodar

### 1. Abrir o projeto no VSCode

Abra a pasta `manifestacao/` no VSCode:
```
File → Open Folder → selecione a pasta manifestacao
```

### 2. Criar um ambiente virtual

No terminal integrado do VSCode (`Ctrl + '`):

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS / Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar as dependências

```bash
pip install -r requirements.txt
```

### 4. Configurar a planilha

Abra o arquivo `.streamlit/secrets.toml` e substitua a URL:

```toml
[google_sheets]
url = "https://docs.google.com/spreadsheets/d/SEU_ID_AQUI/edit"
```

> **Planilha pública:** basta trocar a URL acima. Compartilhe a planilha como
> "Qualquer pessoa com o link pode visualizar".
>
> **Planilha privada:** descomente e preencha o bloco `[gcp_service_account]`
> com as credenciais do seu Service Account do GCP e conceda acesso à planilha
> para o e-mail do Service Account.

### 5. Rodar o app

```bash
streamlit run app.py
```

O navegador abrirá automaticamente em `http://localhost:8501`.

---

## Formato esperado da planilha

A planilha deve ter **duas abas** com os nomes exatos:

| Aba | Colunas obrigatórias |
|-----|----------------------|
| `SAP` | `Chave de Acesso`, `Situação Manifestação`, `Nº NF` |
| `Qive` | `Chave de acesso` |

### Filtro base (aba SAP)

O app processa apenas linhas onde:
- `Situação Manifestação` é um dos valores: `Não informado`, `Uso autorizado`, `Doc não gerado`
- **ou** `Tipo Doc. Fiscal` = `NF-e`

### Exclusões automáticas

Linhas onde a coluna `Observação` contém as palavras `comodato`, `retorno`,
`devolução` ou `devolucao` são removidas do fluxo principal e exibidas na aba
**Excluídos**.

---

## Abas do app

| Aba | Descrição |
|-----|-----------|
| 📄 Cruzamento SAP × Qive | Merge completo com status de cruzamento |
| ⚠️ Pendências de Manifestação | Notas com "Não informado" ou "Sem Manifestação" |
| 📂 Fora do Escopo | Documentos SAP fora dos critérios de filtro |
| 📦 Excluídos | Retornos, comodatos e devoluções |

Todas as abas têm **busca por Nº NF / emitente / chave**, **filtro por status**
e **exportação para `.xlsx`**.

---

## Dúvidas comuns

**O app não carrega e mostra erro de URL**
→ Verifique se o valor em `secrets.toml` é uma URL válida do Google Sheets
  e se a planilha está compartilhada corretamente.

**"Aba SAP está faltando colunas obrigatórias"**
→ Verifique se os nomes das colunas na planilha batem exatamente com os
  listados acima (maiúsculas, acentos e espaços incluídos).

**As checkboxes "Tratado" se perdem ao atualizar**
→ Esperado — a marcação é visual e temporária dentro da sessão.
  Use o botão **Exportar** para salvar o estado filtrado.
