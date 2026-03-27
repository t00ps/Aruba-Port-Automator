# Używamy lekkiego Linuxa (Debian) z gotowym Pythonem 3.10
FROM python:3.10-slim

# Ustawiamy folder roboczy wewnątrz kontenera
WORKDIR /app

# Kopiujemy nasz plik z wymaganiami
COPY requirements.txt .

# Instalujemy biblioteki (np. netmiko)
RUN pip install --no-cache-dir -r requirements.txt

# Kopiujemy nasz kod Pythona do kontenera
COPY arubauth.py .

# Komenda, która wykona się przy starcie kontenera
CMD ["python", "arubauth.py"]