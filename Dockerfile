FROM python:3

WORKDIR /app

COPY requirements.txt /app
RUN pip install --no-cache-dir -r requirements.txt

COPY index.html /app
COPY main.py /app

RUN chown -R nobody:nogroup /app

USER nobody

CMD [ "python", "./main.py" ]
