# Dockerfile for Puralox Flask + eLabFTW integration

FROM python:3.11-slim

# set working dir
WORKDIR /usr/src/app

# system deps for pandas, reportlab, etc.
RUN apt-get update \
 && apt-get install -y --no-install-recommends build-essential \
 && rm -rf /var/lib/apt/lists/*

# copy & install Python dependencies
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# copy application code
COPY . .

# expose Flask port
EXPOSE 5000

# default command
CMD ["python", "-m", "puralox.app"]
