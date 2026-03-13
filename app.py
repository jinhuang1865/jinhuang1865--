import streamlit as st
import pandas as pd
import os
import json
import requests
import base64
from datetime import datetime
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# 页面配置
st.set_page_config(page_title="名单收集系统", page_icon="📋", layout="wide")

# 路径配置
DATA_DIR = "data"
TEMPLATES_DIR = "templates"
DROPDOWN_CONFIG_FILE = "templates/dropdown_config.json"
DATA_FILE = os.path.join(DATA_DIR, "submissions.csv")
BACKUP_CSV_FILE = "data/backup_submissions.csv"

EXPORT_PASSWORD = "907"

# GitHub仓库
GITHUB_REPO = "jinhuang1865/jinhuang1865--"

# 创建目录
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# 初始化CSV
if not os.path.exists(DATA_FILE):
    df = pd.DataFrame(columns=["提交时间", "模板名称"])
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

# -----------------------------
# 本地备份
# -----------------------------
def backup_to_local_csv(df):
    df.to_csv(BACKUP_CSV_FILE, index=False, encoding="utf-8-sig")
    return True


# -----------------------------
# GitHub API备份
# -----------------------------
def backup_to_github():

    try:
        token = os.getenv("GITHUB_TOKEN")

        if not token:
            st.warning("⚠️ 未设置 GITHUB_TOKEN")
            return False

        repo = GITHUB_REPO
        path = "data/submissions.csv"

        url = f"https://api.github.com/repos/{repo}/contents/{path}"

        headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github+json"
        }

        with open(DATA_FILE, "rb") as f:
            content = base64.b64encode(f.read()).decode()

        r = requests.get(url, headers=headers)

        sha = None
        if r.status_code == 200:
            sha = r.json()["sha"]

        data = {
            "message": f"Auto backup {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "content": content,
            "branch": "main"
        }

        if sha:
            data["sha"] = sha

        r = requests.put(url, headers=headers, json=data)

        if r.status_code in [200, 201]:
            return True
        else:
            st.error(f"GitHub API错误: {r.text}")
            return False

    except Exception as e:
        st.error(f"GitHub备份失败: {str(e)}")
        return False


# -----------------------------
# 配置加载
# -----------------------------
def load_dropdown_config():
    if os.path.exists(DROPDOWN_CONFIG_FILE):
        with open(DROPDOWN_CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_dropdown_config(config):
    with open(DROPDOWN_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# -----------------------------
# 模板函数
# -----------------------------
def get_template_files():
    templates = []
    for f in os.listdir(TEMPLATES_DIR):
        if f.endswith(('.xlsx', '.xls')) and not f.startswith('~'):
            templates.append(f)
    return sorted(templates)


def get_template_columns(template_name):

    template_path = os.path.join(TEMPLATES_DIR, template_name)

    try:
        df = pd.read_excel(template_path)
        return [col for col in df.columns if col not in ["提交时间", "模板名称"]]
    except:
        return []


def parse_dropdown_options(input_text):

    options = []

    for sep in ['\n', '\r\n', ',', '，', ';', '；', '|']:
        if sep in input_text:
            options = input_text.replace('\r', '').split(sep)
            break

    if not options:
        options = [input_text]

    options = [opt.strip() for opt in options if opt.strip()]

    return options


def create_excel_with_dropdown(columns, dropdown_config, filename):

    wb = Workbook()
    ws = wb.active
    ws.title = "模板"

    for idx, col in enumerate(columns, 1):
        ws.cell(1, idx, col)

    for idx, col in enumerate(columns, 1):
        if col in dropdown_config and dropdown_config[col]:
            ws.cell(2, idx, f"请选择{col}")

    for idx, col in enumerate(columns, 1):

        if col in dropdown_config and dropdown_config[col]:

            options = dropdown_config[col]

            dv = DataValidation(
                type="list",
                formula1=f'"{",".join(options)}"',
                allow_blank=True
            )

            ws.add_data_validation(dv)

            dv.add(f'{get_column_letter(idx)}2:{get_column_letter(idx)}1048576')

    wb.save(filename)


# -----------------------------
# 页面
# -----------------------------
st.title("📋 名单收集系统")
st.markdown("---")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "模板管理",
    "下拉配置",
    "下载模板",
    "上传名单",
    "查看导出"
])


# -----------------------------
# 模板管理
# -----------------------------
with tab1:

    st.header("模板管理")

    admin_password = st.text_input("密码", type="password")

    if admin_password != EXPORT_PASSWORD:

        st.warning("请输入正确密码")

    else:

        new_template_file = st.file_uploader(
            "上传模板",
            type=["xlsx", "xls"]
        )

        if new_template_file:

            template_name = st.text_input("模板名称")

            if st.button("保存模板"):

                template_path = os.path.join(
                    TEMPLATES_DIR,
                    f"{template_name}.xlsx"
                )

                with open(template_path, "wb") as f:
                    f.write(new_template_file.getbuffer())

                st.success("模板保存成功")
                st.rerun()

        st.markdown("---")

        templates = get_template_files()

        for t in templates:

            col1, col2 = st.columns([4, 1])

            col1.write(t)

            if col2.button("删除", key=t):

                os.remove(os.path.join(TEMPLATES_DIR, t))

                st.rerun()


# -----------------------------
# 下拉配置
# -----------------------------
with tab2:

    st.header("下拉选项配置")

    password = st.text_input("密码", type="password")

    if password == EXPORT_PASSWORD:

        templates = get_template_files()

        if templates:

            selected_template = st.selectbox("选择模板", templates)

            columns = get_template_columns(selected_template)

            config = load_dropdown_config()

            current_config = config.get(selected_template, {})

            col_to_add = st.selectbox("字段", columns)

            options_input = st.text_area("输入选项")

            if st.button("保存配置"):

                opts = parse_dropdown_options(options_input)

                current_config[col_to_add] = opts

                config[selected_template] = current_config

                save_dropdown_config(config)

                st.success("保存成功")

                st.rerun()


# -----------------------------
# 下载模板
# -----------------------------
with tab3:

    templates = get_template_files()

    if templates:

        selected_template = st.selectbox("模板", templates)

        columns = get_template_columns(selected_template)

        config = load_dropdown_config()

        dropdown_config = config.get(selected_template, {})

        if st.button("生成模板"):

            temp_file = f"temp_{selected_template}"

            create_excel_with_dropdown(
                columns,
                dropdown_config,
                temp_file
            )

            with open(temp_file, "rb") as f:

                st.download_button(
                    "下载模板",
                    f,
                    file_name=selected_template
                )

            os.remove(temp_file)


# -----------------------------
# 上传名单
# -----------------------------
with tab4:

    st.header("上传名单")

    templates = get_template_files()

    if templates:

        selected_template = st.selectbox(
            "选择模板",
            templates
        )

        uploaded_file = st.file_uploader(
            "上传Excel",
            type=["xlsx", "xls"]
        )

        if uploaded_file:

            df_upload = pd.read_excel(uploaded_file)

            df_upload["模板名称"] = selected_template

            df_upload["提交时间"] = datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )

            df_existing = pd.read_csv(
                DATA_FILE,
                encoding="utf-8-sig"
            )

            df_combined = pd.concat(
                [df_existing, df_upload],
                ignore_index=True
            )

            if "工号" in df_combined.columns:
                df_combined = df_combined.drop_duplicates(
                    subset=["工号"],
                    keep="last"
                )

            df_combined.to_csv(
                DATA_FILE,
                index=False,
                encoding="utf-8-sig"
            )

            backup_to_local_csv(df_combined)

            with st.spinner("正在备份到GitHub"):

                backup_to_github()

            st.success(f"成功提交 {len(df_upload)} 条数据")


# -----------------------------
# 查看导出
# -----------------------------
with tab5:

    password = st.text_input("查看密码", type="password")

    if password == EXPORT_PASSWORD:

        df_all = pd.read_csv(DATA_FILE, encoding="utf-8-sig")

        st.dataframe(df_all, use_container_width=True)

        if st.button("导出Excel"):

            name = f"名单_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"

            df_all.to_excel(name, index=False)

            with open(name, "rb") as f:

                st.download_button(
                    "下载",
                    f,
                    file_name=name
                )
