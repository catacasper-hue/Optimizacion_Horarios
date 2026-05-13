# Usamos una imagen de Python ligera y reciente
FROM python:3.11-slim

# Establecemos el directorio de trabajo dentro del contenedor
WORKDIR /app

# Copiamos primero el archivo de dependencias para aprovechar la caché de capas de Docker
COPY requirements.txt .

# Instalamos las dependencias
RUN pip install --no-cache-dir -r requirements.txt

# Copiamos todo el código fuente al directorio de trabajo en el contenedor
COPY . .

# El puerto interno que documentamos (Render proveerá el env var PORT para la ejecución real)
EXPOSE 8000

# Comando de inicio usando uvicorn. 
# Render inyecta automáticamente la variable de entorno $PORT, así que la usamos. 
# Si $PORT no existe (e.g. en desarrollo local), usará el puerto 8000.
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]
