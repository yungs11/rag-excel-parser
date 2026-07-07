# 7.excel-parser/Dockerfile
FROM python:3.12-slim AS base
WORKDIR /app
# Node(kordoc CLI용)
RUN apt-get update && apt-get install -y --no-install-recommends \
      curl build-essential nodejs npm && rm -rf /var/lib/apt/lists/*
# kordoc: public npm package (npm info kordoc 확인됨, MIT)
RUN npm install -g kordoc
COPY requirements*.txt pyproject.toml* setup.py* ./
COPY excel_parser_rag ./excel_parser_rag
COPY service ./service
RUN pip install --no-cache-dir uv && \
    (uv pip install --system --no-cache -r requirements.txt 2>/dev/null || uv pip install --system --no-cache -e .) && \
    uv pip install --system --no-cache gunicorn uvicorn
ENV EXCEL_PARSER_BACKEND=auto KORDOC_BIN=kordoc KORDOC_MD_OUT=/tmp/kordoc_md_out
RUN mkdir -p /tmp/kordoc_md_out
EXPOSE 18055
# -w 1: _job_store 는 in-process 전역 변수 → multi-worker 에서 cross-worker 404 발생
CMD ["gunicorn","-k","uvicorn.workers.UvicornWorker","-w","1","-b","0.0.0.0:18055","--timeout","600","service.main:app"]
