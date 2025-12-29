FROM ubuntu:24.04
WORKDIR /src
COPY . .
RUN apt-get update && apt-get install -y ffmpeg python3.12 python3-pip
RUN pip install openpyxl moviepy pychorus jsonschema ruamel.yaml --break-system-packages
CMD ["python3", "recap_generator.py"]