# for DigiKey Auto Processing

## 项目架构概览
- **核心功能**：本项目用于批量查询 Digi-Key 产品状态，并将结果写入 Excel 文件，支持 Web 界面和命令行两种方式。
- **主要组件**：
  - `web.py`：Flask Web 服务，处理文件上传、任务启动、进度查询、结果下载等。
  - `main.py`：命令行批量处理入口，适合本地批量处理。
  - `digikey.py`：DigiKey API 客户端，负责鉴权和产品信息查询。
  - `write_excel.py`：Excel 读写工具，支持多列写入。
  - `product_details.json`：保存最近一次处理的产品详情结果。
  - `logs/`：日志目录，按日期分文件，便于追踪问题。
  - `uploads/`：上传文件存储目录。
  - `templates/`、`static/`：前端页面和样式。

## 关键开发与运行流程
- **Web 端启动**：
  - 运行 `python web.py` 启动 Flask 服务，访问主页上传 Excel 文件，配置参数后发起处理。
  - 处理进度通过 `/processing_status` 轮询获取，结果可下载。
- **命令行批量处理**：
  - 运行 `python main.py`，按提示输入文件路径、工作表名、产品编号列名、输出列名。
  - 处理结果写入原 Excel 文件和 `product_details.json`。
- **日志**：所有操作均详细记录在 `logs/`，便于调试和追踪。

## 约定与模式
- **日志记录**：所有主流程、异常、关键步骤均写日志，日志文件名含日期。
- **Excel 处理**：
  - 仅支持 `.xlsx`/`.xls` 文件。
  - 支持多列写入，列名可自定义（Web 端通过 `custom_headers`）。
- **API 调用**：
  - DigiKey API 凭证通过环境变量或代码默认值配置。
  - Token 自动缓存，过期自动刷新。
- **Web 端异步处理**：
  - 任务在新线程中执行，支持进度查询，防止阻塞主线程。
  - 处理状态通过全局 `processing_status` 字典维护。

## 依赖与环境
- 依赖见 `requirements.txt`，需提前 `pip install -r requirements.txt`。
- 需 Python 3.7+。

## 其他说明
- 关于如何使用，请查看wiki
