FROM python:3.10-slim

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY aruba port-security.py .

CMD ["python", "aruba port-security.py"]
