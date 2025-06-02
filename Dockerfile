FROM python:3.11-slim

# Define o diretório de trabalho
WORKDIR /app

# Copia o requirements.txt primeiro (para melhor cache)
COPY requirements.txt .

# Instala as dependências
RUN pip install --no-cache-dir -r requirements.txt

# Copia apenas o app.py
COPY app.py .

# Expõe a porta do Streamlit
EXPOSE 8501

# Comando para rodar o Streamlit
CMD ["streamlit", "run", "app.py", "--server.address", "0.0.0.0", "--server.port", "8501"]