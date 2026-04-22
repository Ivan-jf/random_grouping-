# 随机分组工具 🔀

基于 Python / Flask 的随机分组 Web 应用，作者：**张济甫**

## 功能
- 上传 Excel 文件，自动读取列名
- 按指定变量列降序分块后随机分组
- 支持"要不要"列过滤（Y 参与，N 排除）
- 输出分组结果 + 各组统计摘要，可下载 Excel

## 本地运行
```bash
pip install -r requirements.txt
python app.py
```
然后打开 http://127.0.0.1:5000

## 部署到 Render（免费）
1. 将本项目推送到 GitHub
2. 在 [render.com](https://render.com) 新建 Web Service
3. 选择仓库，Build Command: `pip install -r requirements.txt`
4. Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT`
5. 部署完成后获得公网 URL
# random_grouping-
