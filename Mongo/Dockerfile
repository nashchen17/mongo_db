FROM python:3.11-slim

# set workdir
WORKDIR /app

# install dependencies
COPY requirements.txt /app/requirements.txt
RUN apt-get update && apt-get install -y build-essential \
    && pip install --upgrade pip \
    && pip install -r /app/requirements.txt \
    && apt-get remove -y build-essential \
    && apt-get autoremove -y && rm -rf /var/lib/apt/lists/*

# copy app
COPY app /app

# expose port
EXPOSE 5000

# Use gunicorn for production
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]