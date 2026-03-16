import streamlit as st
import pandas as pd
import os
import requests
import base64
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

st.set_page_config(page_title="名单收集系统", page_icon="📋", layout="wide")

# 目录与文件配置
DATA_DIR = "data"
TEMPLATES_DIR = "templates"
DATA_FILE = os.path.join(DATA_DIR, "submissions.csv")
BACKUP_CSV_FILE = "data/backup_submissions.csv"
EXPORT_PASSWORD = "907"
GITHUB_REPO = "jinhuang1865/jinhuang1865--"

# 创建目录（不存在则创建）
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# 初始化主数据文件（首次运行创建空CSV）
if not os.path.exists(DATA_FILE):
    df = pd.DataFrame(columns=["提交时间", "模板名称"])
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

# 本地CSV备份函数
def backup_to_local_csv(df):
    df.to_csv(BACKUP_CSV_FILE, index=False, encoding="utf-8-sig")
    return True

# GitHub备份函数（优化日志、Token校验，确保机制生效）
def backup_to_github():
    try:
        token = os.getenv("GITHUB_TOKEN")
        if not token:
            st.warning("⚠️ 未检测到 GITHUB_TOKEN 环境变量，请检查配置")
            return False
        # 校验Token有效性（简单GET请求验证）
        token_check_url = "https://api.github.com/user"
        check_headers = {"Authorization": f"token {token}"}
        check_r = requests.get(token_check_url, headers=check_headers, timeout=10)
        if check_r.status_code != 200:
            st.error(f"⚠️ GITHUB_TOKEN 无效或权限不足，校验返回码：{check_r.status_code}")
            return False
        
        repo = GITHUB_REPO
        path = "data/submissions.csv"
        url = f"https://api.github.com/repos/{repo}/contents/{path}"
        headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github+json"
        }
        # 读取本地数据文件（确保为最新数据）
        if not os.path.exists(DATA_FILE):
            st.error("⚠️ 本地主数据文件不存在，无法备份到GitHub")
            return False
        with open(DATA_FILE, "rb") as f:
            content = base64.b64encode(f.read()).decode()
        
        # 获取文件SHA（用于覆盖更新）
        sha = None
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha = r.json()["sha"]
        elif r.status_code != 404:
            st.error(f"⚠️ 获取GitHub文件信息失败：{r.text}")
            return False
        
        # 构造提交数据
        data = {
            "message": f"Auto backup {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "content": content,
            "branch": "main"
        }
        if sha:
            data["sha"] = sha
        
        # 提交到GitHub
        r = requests.put(url, headers=headers, json=data, timeout=15)
        if r.status_code in [200, 201]:
            st.success("✅ GitHub备份成功！")
            return True
        else:
            st.error(f"❌ GitHub API提交失败：{r.status_code}，详情：{r.text}")
            return False
    except requests.exceptions.Timeout:
        st.error("❌ GitHub备份请求超时，请检查网络连接")
        return False
    except Exception as e:
        st.error(f"❌ GitHub备份异常：{str(e)}")
        return False

# 获取模板文件列表（仅.xlsx，排除临时文件）
def get_template_files():
    templates = []
    for f in os.listdir(TEMPLATES_DIR):
        if f.endswith('.xlsx') and not f.startswith('~'):
            templates.append(f)
    return sorted(templates)

# 获取模板列名（排除系统默认列）
def get_template_columns(template_name):
    template_path = os.path.join(TEMPLATES_DIR, template_name)
    try:
        df = pd.read_excel(template_path)
        return [col for col in df.columns if col not in ["提交时间", "模板名称"]]
    except Exception as e:
        st.error(f"读取模板失败: {e}")
        return []

# 从Excel模板提取下拉选项（终极修复：兼容CellRange/MultiCellRange所有场景，解决不可下标问题）
def get_dropdown_options_from_template(template_path):
    options_dict = {}
    try:
        wb = load_workbook(template_path, data_only=True)
        ws = wb.active
        if not ws.data_validations:
            return options_dict  # 无数据验证规则直接返回
        
        for dv in ws.data_validations.dataValidation:
            col = None
            # 核心修复：统一解析CellRange/MultiCellRange，获取第一个单元格的列索引
            if isinstance(dv.cells, CellRange):
                # 处理单个CellRange：直接取起始列
                col = dv.cells.min_col
            elif isinstance(dv.cells, MultiCellRange):
                # 处理MultiCellRange：取第一个子范围的起始列
                if dv.cells.ranges:
                    col = dv.cells.ranges[0].min_col
            else:
                # 其他未知范围类型：跳过
                continue
            
            if not col:
                continue  # 列索引为空直接跳过
            
            # 获取列标题（第一行对应列）
            title_cell = ws.cell(row=1, column=col)
            col_name = title_cell.value
            if not col_name or col_name.strip() == "":
                continue  # 列名为空跳过
            
            # 解析下拉选项公式
            formula = dv.formula1
            if not formula:
                continue  # 无下拉选项公式跳过
            options = []
            
            # 解析直接定义的选项（如 "选项1,选项2,选项3"）
            if formula.startswith('"') and formula.endswith('"'):
                options_str = formula[1:-1]
                options = [opt.strip() for opt in options_str.split(',') if opt.strip()]
            # 解析引用单元格的选项（如 'Sheet1!$A$1:$A$5'）
            elif '!' in formula:
                try:
                    # 拆分工作表和单元格范围
                    sheet_part, range_part = formula.split('!', 1)
                    sheet_name = sheet_part.strip("'")
                    # 解析引用范围
                    min_c, min_r, max_c, max_r = range_boundaries(range_part)
                    # 定位引用工作表
                    ref_ws = wb[sheet_name] if sheet_name in wb.sheetnames else ws
                    # 遍历引用单元格提取选项
                    for row in ref_ws.iter_rows(min_col=min_c, max_col=max_c, min_row=min_r, max_row=max_r):
                        for cell in row:
                            val = cell.value
                            if val and str(val).strip() != "":
                                options.append(str(val).strip())
                    # 去重并保留顺序
                    options = list(dict.fromkeys(options))
                except Exception as e:
                    # 单个字段解析失败不影响整体，仅跳过
                    continue
        
            # 提取到有效选项则加入字典
            if options:
                options_dict[col_name] = options
    except Exception as e:
        # 仅捕获异常不报错，完全不影响系统使用
        pass
    return options_dict

# 过滤空文字/空行
def clean_empty_data(df):
    # 去除全空行
    df = df.dropna(how='all')
    # 去除单元格仅含空格的行
    df = df[df.apply(lambda row: row.astype(str).str.strip().any(), axis=1)]
    # 重置索引
    df = df.reset_index(drop=True)
    return df

# 主页面标题
st.title("📋 名单收集系统")
st.markdown("---")

# 标签页定义（模板管理、下载模板、上传名单、查看导出）
tab1, tab2, tab3, tab4 = st.tabs([
    "📝 模板管理",
    "📥 下载模板",
    "📤 上传名单",
    "👀 查看导出"
])

# ---------- 模板管理标签页 ----------
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

# ---------- 下载模板标签页 ----------
with tab2:
    st.header("📥 下载模板")
    templates = get_template_files()
    if not templates:
        st.warning("⚠️ 暂无可用模板")
    else:
        selected_template = st.selectbox("📋 选择模板", templates, key="download_template")
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

# ---------- 上传名单标签页（已强制每个字段非空，并修复下拉解析报错）----------
with tab3:
    st.header("📤 上传名单")
    templates = get_template_files()
    if not templates:
        st.warning("⚠️ 暂无可用模板")
    else:
        selected_template = st.selectbox("📋 选择模板", templates, key="upload_template_select")
        template_path = os.path.join(TEMPLATES_DIR, selected_template)
        template_cols = get_template_columns(selected_template)
        dropdown_options = get_dropdown_options_from_template(template_path)
        
        if dropdown_options:
            st.info(f"📌 模板中包含下拉字段：{', '.join(dropdown_options.keys())}")
        
        uploaded_file = st.file_uploader("📤 上传Excel文件（已填写的数据）", type=["xlsx", "xls"], key="upload_excel")
        # 解决重复提交：通过session_state标记是否已处理过当前文件
        if "uploaded_file_key" not in st.session_state:
            st.session_state.uploaded_file_key = None
        
        if uploaded_file:
            # 生成唯一文件标识（文件名+文件大小），避免重复处理
            current_file_key = f"{uploaded_file.name}_{uploaded_file.size}"
            if st.session_state.uploaded_file_key != current_file_key:
                try:
                    df_upload = pd.read_excel(uploaded_file)
                    # 过滤空文字/空行
                    df_upload = clean_empty_data(df_upload)
                    if len(df_upload) == 0:
                        st.warning("⚠️ 上传文件无有效数据（全为空白/空行），请检查后重新上传")
                        st.session_state.uploaded_file_key = current_file_key
                        st.stop()
                    
                    # 1. 检查缺失字段（模板中的字段在上传文件中是否存在）
                    missing_cols = [col for col in template_cols if col not in df_upload.columns]
                    if missing_cols:
                        st.error(f"❌ 文件缺少必要字段：{', '.join(missing_cols)}，请使用正确的模板填写")
                        st.session_state.uploaded_file_key = current_file_key
                        st.stop()
                    
                    # 2. 检查空值（每个字段都必须有内容）
                    empty_cols = []
                    for col in template_cols:
                        if col in df_upload.columns:
                            # 判断是否包含空值（NaN或空字符串）
                            if df_upload[col].isna().any() or (df_upload[col].astype(str).str.strip() == '').any():
                                empty_cols.append(col)
                    if empty_cols:
                        st.error(f"❌ 以下字段存在空值，请填写完整后再上传：{', '.join(empty_cols)}")
                        st.session_state.uploaded_file_key = current_file_key
                        st.stop()
                    
                    # 3. 验证下拉选项合法性
                    invalid_options = []
                    if dropdown_options:
                        for col, options in dropdown_options.items():
                            if col in df_upload.columns:
                                values = df_upload[col].dropna().astype(str).str.strip()
                                values = values[~values.str.contains("请选择", na=False)]
                                invalid = [v for v in values if v not in options]
                                if invalid:
                                    invalid_options.append(f"{col}：{','.join(invalid[:3])}（共{len(invalid)}个无效值）")
                    if invalid_options:
                        st.error(f"❌ 下拉选项验证失败：{'; '.join(invalid_options)}")
                        st.info("💡 请确保所有值均来自模板定义的下拉列表")
                        st.session_state.uploaded_file_key = current_file_key
                        st.stop()
                    
                    # 所有验证通过，执行保存
                    # 添加系统元数据列
                    df_upload["模板名称"] = selected_template
                    df_upload["提交时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    
                    # 读取本地已有数据，合并并去重
                    df_existing = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
                    df_combined = pd.concat([df_existing, df_upload], ignore_index=True)
                    # 按工号/姓名去重（保留最新提交）
                    if "工号" in df_combined.columns:
                        df_combined = df_combined.drop_duplicates(subset=["工号"], keep="last")
                    elif "姓名" in df_combined.columns:
                        df_combined = df_combined.drop_duplicates(subset=["姓名"], keep="last")
                    
                    # 仅保存一次合并后的数据（解决重复写入）
                    df_combined.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
                    # 执行本地CSV备份
                    backup_to_local_csv(df_combined)
                    # 执行GitHub备份
                    with st.spinner("🔄 正在备份到GitHub..."):
                        backup_to_github()
                    
                    st.success(f"✅ 数据提交成功！共导入 {len(df_upload)} 条有效记录")
                    st.dataframe(df_upload.head(10), use_container_width=True)  # 仅展示本次上传数据
                    
                    # 更新session_state，标记当前文件已处理
                    st.session_state.uploaded_file_key = current_file_key
                    
                except Exception as e:
                    st.error(f"❌ 文件处理失败：{str(e)}")
            else:
                # 已处理过该文件，仅提示不重复执行
                st.info("ℹ️ 该文件已处理完成，无需重复上传")

# ---------- 查看导出标签页 ----------
with tab4:
    st.header("👀 查看与导出")
    password = st.text_input("🔑 查看密码", type="password", key="tab4_password")
    if password == EXPORT_PASSWORD:
        df_all = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
        if len(df_all) == 0:
            st.warning("⚠️ 暂无提交记录")
        else:
            # 按模板筛选数据
            if "模板名称" in df_all.columns:
                template_filter = st.selectbox("📋 按模板筛选", ["全部"] + list(df_all["模板名称"].unique()), key="template_filter")
                if template_filter != "全部":
                    df_all = df_all[df_all["模板名称"] == template_filter]
            # 数据统计指标
            col1, col2, col3 = st.columns(3)
            col1.metric("总记录数", len(df_all))
            col2.metric("模板数", df_all["模板名称"].nunique() if "模板名称" in df_all.columns else 0)
            col3.metric("最新提交", df_all["提交时间"].max() if "提交时间" in df_all.columns else "N/A")
            
            # 展示数据
            st.dataframe(df_all, use_container_width=True, height=400)
            st.markdown("---")
            
            # 导出Excel功能
            if st.button("📥 导出Excel", key="export_excel"):
                export_filename = f"名单收集_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
                df_all.to_excel(export_filename, index=False, encoding="utf-8-sig")
                with open(export_filename, "rb") as f:
                    st.download_button(
                        label="⬇️ 点击下载导出文件",
                        data=f,
                        file_name=export_filename,
                        key="download_export_btn"
                    )
                # 删除本地临时导出文件
                if os.path.exists(export_filename):
                    os.remove(export_filename)
    else:
        st.warning("🔐 请输入正确的查看密码")
