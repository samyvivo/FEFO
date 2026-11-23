FROM python:3.11-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
        python3-tk tk \
        xvfb x11vnc novnc websockify \
        libx11-6 libxext6 libxrender1 libsm6 libice6 \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY fefo.py .
RUN mkdir -p /data
VOLUME ["/data"]
EXPOSE 6080
CMD bash -lc "Xvfb :99 -screen 0 1280x900x24 & \
              x11vnc -display :99 -nopw -forever -shared -quiet & \
              websockify --web=/usr/share/novnc 6080 localhost:5900 & \
              DISPLAY=:99 python3 /app/fefo.py"