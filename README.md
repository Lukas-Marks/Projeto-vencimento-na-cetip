# 🔎 B3 ISIN Expiration Scraper (Headless)

Automação em Python que consulta códigos **ISIN** no site da B3 e extrai a **data de expiração** diretamente da página oficial, utilizando Selenium em modo **headless** (execução invisível).

Projeto ideal para rotinas operacionais, reconciliação e validação de ativos no contexto de mercado financeiro.

---

## 🌐 Fonte dos Dados

Consulta realizada no site oficial da B3:

https://www.b3.com.br/pt_br/market-data-e-indices/servicos-de-dados/market-data/consultas/mercado-a-vista/codigo-isin/pesquisa/

---

## 🚀 Funcionalidades

- 📥 Leitura automática de ISINs a partir de um arquivo Excel (`isins.xlsx`)
- 🌐 Acesso automático ao site da B3
- 🧠 Tratamento de iframe (`bvmf_iframe`)
- 🔎 Consulta individual de cada ISIN
- 📅 Extração da data de expiração
- ⚙️ Execução em modo headless (sem abrir navegador visualmente)
- 📊 Geração automática de arquivo `resultado.xlsx`
- 🛡️ Tratamento de erros e ISINs não encontrados

---

## 🛠️ Tecnologias Utilizadas

- Python 3.x  
- `pandas` – Manipulação de dados e Excel  
- `selenium` – Automação de navegador  
- `webdriver-manager` – Gerenciamento automático do ChromeDriver  

---

## 📦 Instalação

### 1️⃣ Clone o repositório

```bash
git clone https://github.com/seuusuario/seu-repositorio.git
cd seu-repositorio
