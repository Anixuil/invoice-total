FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

ARG DEBIAN_MIRROR=https://mirrors.aliyun.com/debian

RUN find /etc/apt -type f \( -name '*.list' -o -name '*.sources' \) -exec sed -i \
        -e "s|https\?://deb.debian.org/debian-security|${DEBIAN_MIRROR}-security|g" \
        -e "s|https\?://deb.debian.org/debian|${DEBIAN_MIRROR}|g" \
        {} + \
    && apt-get update \
    && apt-get install --no-install-recommends -y \
        libglib2.0-0 \
        libgl1 \
        libxcb1 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
ARG PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple
RUN python -m pip install --upgrade pip \
        --index-url ${PIP_INDEX_URL} \
    && python -m pip install -r requirements.txt \
        --index-url ${PIP_INDEX_URL}

COPY invoice_total.py jira_processor.py weekly_report_processor.py reimbursement_generator.py server.py ./
COPY static ./static
COPY templates ./templates

RUN useradd --create-home --uid 10001 appuser \
    && mkdir -p /tmp/invoice_total_preview_jobs /tmp/invoice_total_jira_jobs /tmp/invoice_total_weekly_jobs \
    && chown -R appuser:appuser /app /tmp/invoice_total_preview_jobs /tmp/invoice_total_jira_jobs /tmp/invoice_total_weekly_jobs

USER appuser

EXPOSE 8000

CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000"]
