# VS Code Sales Report Generator

命令：`Generate Sales Report` — 在终端中运行 Python 脚本，提示输入 Excel 路径，生成完整的 BI HTML 报表 `index.html`。

安装依赖：

```powershell
python -m pip install -r requirements.txt
```

示例 Excel 可以是单表或多表：
- 若使用单表，请确保包含 `order_id`, `product_name`, `sales_amount` 等必要列。
- 若使用多表，可命名为 `overview`, `category`, `products`, `trend`, `traffic`, `funnels`。

生成后的文件：
- `index.html`：企业级 BI dashboard，可直接在浏览器打开显示。
