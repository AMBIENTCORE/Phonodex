FROM debian:bookworm-slim

ARG PHONODEX_UID=1000
ARG PHONODEX_GID=1000

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PATH="/opt/venv/bin:${PATH}" \
    HOME=/config \
    DISPLAY=:99 \
    PHONODEX_CONFIG_DIR=/config \
    PHONODEX_LIBRARY_DIR=/music \
    PHONODEX_RUNTIME_DIR=/config/.runtime \
    PHONODEX_BROWSER=1 \
    DBUS_SESSION_BUS_ADDRESS=unix:path=/config/.runtime/bus \
    PHONODEX_SCREEN=1920x1080x24

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        ca-certificates fontconfig fonts-dejavu-core \
        python3 python3-venv python3-tk \
        xvfb x11-utils xclip openbox x11vnc novnc websockify supervisor \
        thunar xdg-utils dbus-daemon \
    && rm -rf /var/lib/apt/lists/* \
    && test "$PHONODEX_UID" -gt 0 && test "$PHONODEX_GID" -gt 0 \
    && groupadd --gid "$PHONODEX_GID" phonodex \
    && useradd --no-log-init --uid "$PHONODEX_UID" --gid "$PHONODEX_GID" \
        --home-dir /config --no-create-home phonodex \
    && install -d -o "$PHONODEX_UID" -g "$PHONODEX_GID" /config /music /work

WORKDIR /opt/phonodex
COPY requirements.txt ./
RUN python3 -m venv /opt/venv \
    && /opt/venv/bin/python -m pip install --no-cache-dir -r requirements.txt

COPY main.py config.py genres.txt LICENSE ./
COPY services/ ./services/
COPY ui/ ./ui/
COPY utils/ ./utils/
COPY assets/ ./assets/
COPY docker/ ./docker/
RUN install -m 644 assets/I-pixel-u.ttf /usr/local/share/fonts/I-pixel-u.ttf \
    && install -D -m 644 docker/mimeapps.list /etc/xdg/mimeapps.list \
    && fc-cache -f \
    && ln -sf vnc.html /usr/share/novnc/index.html

USER ${PHONODEX_UID}:${PHONODEX_GID}
EXPOSE 6080
STOPSIGNAL SIGTERM
HEALTHCHECK --interval=15s --timeout=20s --start-period=60s --retries=3 \
    CMD ["/usr/bin/python3", "/opt/phonodex/docker/healthcheck.py"]
ENTRYPOINT ["/usr/bin/python3", "/opt/phonodex/docker/entrypoint.py"]
