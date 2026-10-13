FROM python:3.12-slim
WORKDIR /app
COPY figma_bridge/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY figma_bridge /app/figma_bridge
RUN useradd --uid 10001 --create-home bridge && mkdir /data && chown bridge:bridge /data
USER bridge
ENV FIGMA_BRIDGE_DB=/data/reports.sqlite3
EXPOSE 8765
CMD ["uvicorn", "figma_bridge.service:create_app", "--factory", "--host", "0.0.0.0", "--port", "8765", "--no-access-log"]
