import streamlit as st
import pandas as pd
import os
import requests
import base64
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

st.set_page_config(page_title="名单收集系统", page_icon="📋", layout="wide")

DATA_DIR = "data"
TEMPLATES_DIR = "templates"
DATA_FILE = os.path.join(DATA_DIR, "submissions.csv")
BACKUP_CSV_FILE = "data/backup_submissions.csv"

EXPORT_PASSWORD = "907"

GITHUB_REPO = "jinhuang1865/jinhuang1865--"

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

if not os.path.exists(DATA_FILE):
    df = pd.DataFrame(columns=["提交时间", "模板名称"])
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")


def backup_to_local_csv(df):
    df.to_csv(BACKUP_CSV_FILE, index=False, encoding="utf-8-sig")
    return True


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


def get_template_files():
    """获取所有 .xlsx 模板文件"""
    templates = []
    for f in os.listdir(TEMPLATES_DIR):
        if f.endswith('.xlsx') and not f.startswith('~'):
            templates.append(f)
    return sorted(templates)


def get_template_columns(template_name):
    """获取模板的列名（排除提交时间和模板名称）"""
    template_path = os.path.join(TEMPLATES_DIR, template_name)
    try:
        df = pd.read_excel(template_path)
        return [col for col in df.columns if col not in ["提交时间", "模板名称"]]
    except Exception as e:
        st.error(f"读取模板失败: {e}")
        return []


def get_dropdown_options_from_template(template_path):
    """
    从模板文件中提取所有数据验证（下拉选项）
    返回字典：{列名: 选项列表}
    """
    options_dict = {}
    try:
        wb = load_workbook(template_path, data_only=True)
        ws = wb.active
        if ws.data_validations:
            for dv in ws.data_validations.dataValidation:
                # 获取范围
                cells_range = dv.cells
                min_col = cells_range.min_col
                max_col = cells_range.max_col
                # 假设下拉列表应用于某一列，取第一列
                col = min_col
                # 获取列标题（第一行）
                title_cell = ws.cell(row=1, column=col)
                col_name = title_cell.value
                if not col_name:
                    continue

                # 解析选项
                formula = dv.formula1
                options = []
                if formula and formula.startswith('"') and formula.endswith('"'):
                    # 直接列表，如 "选项1,选项2"
                    options_str = formula[1:-1]
                    options = [opt.strip() for opt in options_str.split(',') if opt.strip()]
                elif formula and '!' in formula:
                    # 引用区域，如 'Sheet1!$A$1:$A$10'
                    parts = formula.split('!')
                    sheet_name = parts[0].strip("'")
                    range_addr = parts[1]
                    if sheet_name in wb.sheetnames:
                        ref_ws = wb[sheet_name]
                    else:
                        ref_ws = ws
                    # 解析范围
                    min_c, min_r, max_c, max_r = range_boundaries(range_addr)
                    for row in ref_ws.iter_rows(min_col=min_c, max_col=max_c, min_row=min_r, max_row=max_r):
                        for cell in row:
                            if cell.value and str(cell.value).strip():
                                options.append(str(cell.value).strip())
                    options = list(dict.fromkeys(options))  # 去重
                if options:
                    options_dict[col_name] = options
    except Exception as e:
        st.error(f"读取模板下拉选项失败: {e}")
    return options_dict


st.title("📋 名单收集系统")
st.markdown("---")

# 重新定义标签页：模板管理、下载模板、上传名单、查看导出
tab1, tab2, tab3, tab4 = st.tabs([
    "📝 模板管理",
    "📥 下载模板",
    "📤 上传名单",
    "👀 查看导出"
])


# ---------- 模板管理 ----------
with tab1:
    st.header("📝 模板管理")

    admin_password = st.text_input("🔑 密码", type="password", key="tab1_password")

    if admin_password != EXPORT_PASSWORD:
        st.warning("🔐 请输入正确密码")
    else:
        st.success("✅ 验证通过")

        new_template_file = st.file_uploader("📤 上传Excel模板（仅支持 .xlsx）", type=["xlsx"], key="upload_template")

        if new_template_file:
            template_name = st.text_input("📝 模板名称（不含扩展名）", key="template_name", placeholder="如：2024年晋升名单")

            if st.button("💾 保存模板", key="save_template"):
                if template_name:
                    # 保留原扩展名（.xlsx）
                    original_ext = os.path.splitext(new_template_file.name)[1]
                    template_path = os.path.join(TEMPLATES_DIR, f"{template_name}{original_ext}")
                    with open(template_path, "wb") as f:
                        f.write(new_template_file.getbuffer())
                    st.success(f"✅ 模板 '{template_name}' 保存成功！")
                    st.rerun()
                else:
                    st.error("❌ 请输入模板名称")

        st.markdown("---")
        st.subheader("📋 已有的模板")

        templates = get_template_files()
        if templates:
            for t in templates:
                col1, col2 = st.columns([4, 1])
                col1.write(f"📄 {t}")
                if col2.button("🗑️ 删除", key=f"delete_{t}"):
                    os.remove(os.path.join(TEMPLATES_DIR, t))
                    st.success(f"✅ 删除 {t}")
                    st.rerun()
        else:
            st.info("暂无模板")


# ---------- 下载模板 ----------
with tab2:
    st.header("📥 下载模板")

    templates = get_template_files()

    if not templates:
        st.warning("⚠️ 暂无可用模板")
    else:
        selected_template = st.selectbox("📋 选择模板", templates, key="download_template")

        # 显示字段信息（可选）
        columns = get_template_columns(selected_template)
        if columns:
            st.info(f"📝 模板字段：{', '.join(columns)}")

        template_path = os.path.join(TEMPLATES_DIR, selected_template)
        with open(template_path, "rb") as f:
            st.download_button(
                label=f"⬇️ 下载 {selected_template}",
                data=f,
                file_name=selected_template,
                key="download_template_btn"
            )


# ---------- 上传名单 ----------
with tab3:
    st.header("📤 上传名单")

    templates = get_template_files()

    if not templates:
        st.warning("⚠️ 暂无可用模板")
    else:
        selected_template = st.selectbox("📋 选择模板", templates, key="upload_template_select")

        template_path = os.path.join(TEMPLATES_DIR, selected_template)
        template_cols = get_template_columns(selected_template)

        # 获取模板中的下拉选项
        dropdown_options = get_dropdown_options_from_template(template_path)
        if dropdown_options:
            st.info(f"📌 模板中包含下拉字段：{', '.join(dropdown_options.keys())}")

        uploaded_file = st.file_uploader("📤 上传Excel文件（已填写的数据）", type=["xlsx", "xls"], key="upload_excel")

        if uploaded_file:
            try:
                df_upload = pd.read_excel(uploaded_file)

                # 检查字段缺失
                if template_cols:
                    missing_cols = [col for col in template_cols if col not in df_upload.columns]
                    if missing_cols:
                        st.warning(f"⚠️ 文件缺少字段：{', '.join(missing_cols)}")

                # 验证下拉选项
                if dropdown_options:
                    invalid_options = []
                    for col, options in dropdown_options.items():
                        if col in df_upload.columns:
                            # 转换为字符串并去除空格
                            values = df_upload[col].dropna().astype(str).str.strip()
                            # 忽略包含“请选择”的值
                            values = values[~values.str.contains("请选择", na=False)]
                            invalid = [v for v in values if v not in options]
                            if invalid:
                                invalid_options.append(f"{col}: {invalid[:3]}...")

                    if invalid_options:
                        st.error(f"❌ 以下字段存在无效选项：{', '.join(invalid_options)}")
                        st.info("💡 请确保所有值都来自模板中定义的下拉列表。")
                    else:
                        st.success("✅ 选项验证通过！")

                # 添加元数据列
                df_upload["模板名称"] = selected_template
                df_upload["提交时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                st.success("✅ 文件上传成功！")
                st.dataframe(df_upload.head(10), use_container_width=True)

                # 保存数据
                df_existing = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
                df_combined = pd.concat([df_existing, df_upload], ignore_index=True)

                # 按工号或姓名去重（保留最新）
                if "工号" in df_combined.columns:
                    df_combined = df_combined.drop_duplicates(subset=["工号"], keep="last")
                elif "姓名" in df_combined.columns:
                    df_combined = df_combined.drop_duplicates(subset=["姓名"], keep="last")

                df_combined.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

                # 本地备份
                backup_to_local_csv(df_combined)

                # GitHub备份
                with st.spinner("正在备份到GitHub..."):
                    backup_to_github()

                st.balloons()
                st.success(f"✅ 成功提交 {len(df_upload)} 条记录！")

            except Exception as e:
                st.error(f"❌ 读取文件失败：{str(e)}")


# ---------- 查看导出 ----------
with tab4:
    st.header("👀 查看与导出")

    password = st.text_input("🔑 查看密码", type="password", key="tab4_password")

    if password == EXPORT_PASSWORD:
        df_all = pd.read_csv(DATA_FILE, encoding="utf-8-sig")

        if len(df_all) == 0:
            st.warning("⚠️ 暂无记录")
        else:
            # 按模板筛选
            if "模板名称" in df_all.columns:
                template_filter = st.selectbox("📋 按模板筛选", ["全部"] + list(df_all["模板名称"].unique()))
                if template_filter != "全部":
                    df_all = df_all[df_all["模板名称"] == template_filter]

            # 统计
            col1, col2, col3 = st.columns(3)
            col1.metric("总记录数", len(df_all))
            if "模板名称" in df_all.columns:
                col2.metric("模板数", df_all["模板名称"].nunique())
            col3.metric("最新提交", df_all["提交时间"].max() if "提交时间" in df_all.columns else "N/A")

            st.dataframe(df_all, use_container_width=True, height=400)

            st.markdown("---")

            # 导出
            if st.button("📥 导出Excel", key="export_excel"):
                name = f"名单_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
                df_all.to_excel(name, index=False)

                with open(name, "rb") as f:
                    st.download_button("⬇️ 下载", f, file_name=name, key="download_export")
    else:
        st.warning("🔐 请输入正确密码")


# ---------- 侧边栏说明 ----------
with st.sidebar:
    st.title("📌 使用说明")
    st.markdown("""
    **备份机制：**
    ✅ 本地CSV自动保存
    ✅ GitHub自动备份

    **步骤：**

    1️⃣ 模板管理 - 上传带下拉选项的模板（.xlsx）

    2️⃣ 下载模板 - 获取模板用于填写

    3️⃣ 上传名单 - 上传填写好的文件，系统自动验证下拉选项

    4️⃣ 查看导出 - 查看和导出已提交数据

    ---
    **密码：907**
    """)
