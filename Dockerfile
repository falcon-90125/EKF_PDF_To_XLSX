FROM python:3.10.4

WORKDIR /app

COPY . .

RUN pip install --no-cache-dir -r requirements.txt 

CMD [ "python", "run.py" ]