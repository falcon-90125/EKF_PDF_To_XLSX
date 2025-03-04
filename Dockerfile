FROM python:3.10.4

WORKDIR /app

COPY . .

RUN apt-get update && apt-get install -y tk libx11-6
RUN pip install --no-cache-dir -r requirements.txt 

CMD [ "python", "run.py" ]