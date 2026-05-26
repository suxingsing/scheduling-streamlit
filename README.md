# 智能排产系统

基于 Streamlit 的排产计划模型，支持制程参数配置、工厂日历、单日工时设置、物料交期 Excel 模板下载与上传，以及老班组/新班组排产逻辑计算。

## 本地运行

```bash
streamlit run app.py
```

## Streamlit Cloud 部署

1. 将 `app.py`、`requirements.txt`、`.gitignore`、`README.md` 上传到 GitHub 仓库。
2. 打开 Streamlit Community Cloud。
3. 选择 GitHub 仓库。
4. Main file path 填写 `app.py`。
5. 点击 Deploy。
