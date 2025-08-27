# Dockerfile
# -----------------------------------------------------------------------------
# Define a receita para construir a imagem da aplicação Flask.
# -----------------------------------------------------------------------------

# 1. Imagem Base: Começamos com uma imagem Python leve e oficial.
FROM python:3.11-slim

# 2. Diretório de Trabalho: Define o diretório padrão dentro do container.
WORKDIR /app

# 3. Instalação de Dependências: Copia e instala as bibliotecas necessárias.
# Copiar apenas o requirements.txt primeiro aproveita o cache do Docker.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. Copiar Código da Aplicação: Copia todos os outros arquivos do projeto.
COPY . .

# 5. Expor a Porta: Informa ao Docker que a aplicação rodará na porta 8000.
EXPOSE 8000

# 6. Comando de Execução: Define como iniciar a aplicação usando Gunicorn.
# O Gunicorn é um servidor WSGI robusto para produção.
# "--bind 0.0.0.0:8000" faz o servidor aceitar conexões de fora do container.
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "app:app"]
