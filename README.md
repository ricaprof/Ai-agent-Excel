# 🛒 Agente de Compras com Voz + LLM

Um assistente inteligente que cria e gerencia **lista de compras** usando **voz** e **LLM**.  
Você pode falar de forma natural (*“coloca 2kg de arroz e três tomates”*) e o sistema entende, salva numa planilha Excel e ainda gera link para enviar no **WhatsApp**.

---

## 🚀 Instalação e Uso

### 1. Clone o repositório
git clone https://github.com/seu-usuario/agente-compras.git
cd agente-compras

### 2. Instale as dependências
pip install -r requirements.txt

> No Linux, pode ser necessário:
sudo apt install portaudio19-dev libasound2-dev

### 3. Baixe o modelo de voz (Vosk PT-BR)
- Pegue em: https://alphacephei.com/vosk/models
- Extraia a pasta `vosk-model` na raiz do projeto.

### 4. Configure o provedor LLM

#### 🔹 Usando **Ollama** (offline e fácil)
- Instale: https://ollama.com/download
- Puxe um modelo leve:
  ollama pull llama3:8b-instruct-q4
- Configure:
  export LLM_PROVIDER=ollama
  export OLLAMA_MODEL=llama3:8b-instruct-q4

#### 🔹 Usando **OpenAI**
  export LLM_PROVIDER=openai
  export OPENAI_API_KEY="sua_chave_aqui"
  export OPENAI_MODEL=gpt-4o-mini

### 5. Rode o agente
python agente_compras_llm.py

---

## 📝 Exemplos de Comando

- ➕ “coloca 2kg de arroz e três tomates”
- ➖ “tira o leite”
- 📋 “o que tem na lista?”
- 📲 “manda no zap”
- ❌ “valeu, pode encerrar”

---

## 📂 Estrutura

agente-compras/
│── agente_compras_llm.py   # Código principal
│── requirements.txt        # Dependências
│── lista_compras.xlsx      # Gerado automaticamente
└── README.md               # Este arquivo

---

## 📜 Licença
MIT – livre para usar e modificar.
