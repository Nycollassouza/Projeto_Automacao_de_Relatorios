# 🤖 Automação de Relatórios Financeiros

Automação completa para **extração, tratamento e atualização de relatórios financeiros** a partir de um portal corporativo.

O projeto utiliza **Selenium, Pandas e automação de Excel** para reduzir o processo manual de geração de relatórios de **10–15 minutos para aproximadamente 1–2 minutos**, minimizando erros humanos.

---

# 🎯 Objetivo

Automatizar o processo de:

1. Login em portal corporativo
2. Download de comprovantes financeiros em Excel
3. Tratamento e limpeza dos dados
4. Atualização automática de planilha de controle
5. Recalculo de fórmulas e geração do relatório final

---

# ⚙️ Tecnologias utilizadas

* **Python**
* **Selenium** – Automação de navegação web
* **Pandas** – Manipulação de dados
* **OpenPyXL** – Manipulação de planilhas Excel
* **PyAutoGUI** – Automação de interface gráfica
* **Win32COM** – Controle do Excel via COM API

---

# 📂 Estrutura do Projeto

```
Automacao_Relatorios/

automacao_relatorio.py   # Script principal da automação
relatorio_atualizacao.xlsx  # Planilha de destino
README.md
```

---

# 🔄 Fluxo da Automação

### 1️⃣ Login no portal

O script utiliza **Selenium** para acessar o portal corporativo e realizar login automaticamente.

Fluxo:

* Acessa a URL do portal
* Preenche usuário e senha
* Aguarda autenticação **2FA manual**

---

### 2️⃣ Extração de dados

Após autenticação:

* Aplica filtros no sistema
* Exporta os dados para Excel
* Baixa automaticamente o arquivo

---

### 3️⃣ Processamento dos dados

Os dados são processados com **Pandas**, incluindo:

* limpeza de colunas
* remoção de dados irrelevantes
* extração de datas
* criação de colunas auxiliares
* mapeamento de conferentes

Exemplo de transformação:

```
Data -> Data_full -> Mes_Ano
```

---

### 4️⃣ Atualização da planilha de controle

O script então:

* limpa os dados antigos da planilha
* importa os novos dados
* adiciona fórmulas automáticas

Exemplos de fórmulas adicionadas:

* Identificação de **estornos**
* Extração de **ano da transação**

---

### 5️⃣ Recálculo automático do Excel

Usando **Win32COM**, o script:

* abre o Excel
* recalcula todas as fórmulas
* salva o arquivo final

---

# 📦 Instalação

Clone o repositório:

```
git clone https://github.com/nycollassouza/Projeto_Automacao_de_Relatorios.git
```

Entre na pasta:

```
cd Projeto_Automacao_de_Relatorios
```

Instale as dependências:

```
pip install selenium pandas openpyxl pyautogui pywin32
```

---

# ⚠️ Pré-requisitos

Antes de executar o script:

✔️ Instalar **Google Chrome**
✔️ Instalar **ChromeDriver** compatível com sua versão do Chrome
✔️ Adicionar ChromeDriver ao **PATH do sistema**

---

# ▶️ Como executar

Execute o script principal:

```
python automacao_relatorio.py
```

O script irá:

1. abrir o navegador
2. realizar login
3. baixar os dados
4. atualizar a planilha automaticamente

---

# 🔐 Configuração

Edite as variáveis no início do script:

```
URL_PORTAL
LOGIN
SENHA
COORD_CLICKS
```

Para produção, recomenda-se usar:

* **variáveis de ambiente**
* **arquivo .env**
* **arquivo YAML de configuração**

---

# ⚠️ Observações

Este projeto foi desenvolvido para **automatizar processos repetitivos de relatórios corporativos**.

Alguns pontos dependem do ambiente:

* resolução da tela (PyAutoGUI)
* layout do portal
* autenticação 2FA
* estrutura da planilha Excel

---

# 🚀 Melhorias futuras

Possíveis evoluções do projeto:

* Configuração via **arquivo YAML**
* Sistema de **logs estruturados**
* Tratamento de erros robusto
* Suporte a múltiplos períodos
* Automação de 2FA
* Containerização com **Docker**

---

# 👨‍💻 Autor

**Nycollas Faustino de Souza**

