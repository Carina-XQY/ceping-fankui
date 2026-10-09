FROM python:3.11-slim

WORKDIR /app

# 安装系统依赖（playwright 需要）
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    ca-certificates \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# 安装 Python 依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 安装 Playwright 浏览器
RUN python -m playwright install chromium \
    && python -m playwright install-deps chromium

# 复制项目文件
COPY . .

# 创建必要目录
RUN mkdir -p _uploads 反馈报告

# 暴露端口
EXPOSE 8080

# 启动
CMD ["python", "app.py"]
