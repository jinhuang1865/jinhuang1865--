import streamlit as st
import pandas as pd
import os
import json
import subprocess
import shutil
from datetime import datetime
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# 配置页面
st.set_page_config(page_title="同事名单收集系统", page_icon="📋", layout="wide")

# 数据文件路径
DATA_DIR = "data"
TEMPLATES_DIR = "templates"
DROPDOWN_CONFIG_FILE = "templates/dropdown_config.json"
DATA_FILE = os.path.join(DATA_DIR, "submissions.csv")
BACKUP_CSV_FILE = "data/backup_submissions.csv"  # 本地备份CSV
EXPORT_PASSWORD = "907"  # 导出密码

# GitHub备份配置
GITHUB_REPO = "jinhuang1865/jinhuang1865--"  # 备份仓库

# 创建必要的目录
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# 初始化CSV文件
if not os.path.exists(DATA_FILE):
    df = pd.DataFrame(columns=["提交时间", "模板名称"])
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

# ========== 备份函数 ==========
def backup_to_local_csv(df):
    """本地备份CSV"""
    df.to_csv(BACKUP_CSV_FILE, index=False, encoding="utf-8-sig")
    return True

def backup_to_github():
    """推送到GitHub备份"""
    try:
        # 检查是否有Git配置
        if not os.path.exists(".git"):
            st.warning("⚠️ 本地不是Git仓库，跳过GitHub备份")
            return False
        
        # 添加文件到Git
        subprocess.run(["git", "add", "data/submissions.csv"], check=True, capture_output=True)
        
        # 检查是否有更改
        result = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)
        if not result.stdout.strip():
            return True  # 没有更改，不需要提交
        
        # 提交
        commit_msg = f"Backup: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        subprocess.run(["git", "commit", "-m", commit_msg], check=True, capture_output=True)
        
        # 推送
        subprocess.run(["git", "push", "origin", "main"], check=True, capture_output=True)
        
        return True
    except Exception as e:
        st.error(f"GitHub备份失败: {str(e)}")
        return False

# 加载/保存下拉配置
def load_dropdown_config():
    if os.path.exists(DROPDOWN_CONFIG_FILE):
        with open(DROPDOWN_CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_dropdown_config(config):
    with open(DROPDOWN_CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

# ========== 辅助函数 ==========
def get_template_files():
    """获取所有模板文件"""
    templates = []
    for f in os.listdir(TEMPLATES_DIR):
        if f.endswith(('.xlsx', '.xls')) and not f.startswith('~'):
            templates.append(f)
    return sorted(templates)

def get_template_columns(template_name):
    """获取模板的列名"""
    template_path = os.path.join(TEMPLATES_DIR, template_name)
    try:
        df = pd.read_excel(template_path)
        return [col for col in df.columns if col not in ["提交时间", "模板名称"]]
    except:
        return []

def parse_dropdown_options(input_text):
    """解析下拉选项，支持多种分隔符"""
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
    """创建带下拉验证的Excel模板"""
    wb = Workbook()
    ws = wb.active
    ws.title = "模板"
    
    # 写入表头
    for idx, col in enumerate(columns, 1):
        ws.cell(1, idx, col)
    
    # 添加示例数据行
    for idx, col in enumerate(columns, 1):
        if col in dropdown_config and dropdown_config[col]:
            ws.cell(2, idx, f"请选择{col}")
    
    # 为有下拉选项的列添加数据验证
    for idx, col in enumerate(columns, 1):
        if col in dropdown_config and dropdown_config[col]:
            options = dropdown_config[col]
            dv = DataValidation(type="list", formula1=f'"{",".join(options)}"', allow_blank=True)
            dv.error = "请从下拉列表中选择"
            dv.errorTitle = "无效的选项"
            dv.prompt = f"请选择{col}"
            dv.promptTitle = col
            ws.add_data_validation(dv)
            dv.add(f'{get_column_letter(idx)}2:{get_column_letter(idx)}1048576')
            
            if "下拉选项" not in wb.sheetnames:
                ws_options = wb.create_sheet("下拉选项参考")
                ws_options.append(["字段名", "可选值"])
                for c, opts in dropdown_config.items():
                    if opts:
                        ws_options.append([c, ", ".join(opts)])
    
    for idx in range(1, len(columns) + 1):
        ws.column_dimensions[get_column_letter(idx)].width = 15
    
    wb.save(filename)

# 页面标题
st.title("📋 同事名单收集系统（升级版）")
st.markdown("---")

# 标签页
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📝 模板管理", 
    "⚙️ 下拉配置",
    "📥 下载模板", 
    "📤 上传名单", 
    "👀 查看/导出"
])

# ========== 标签1：模板管理 ==========
with tab1:
    st.header("📝 模板管理")
    
    admin_password = st.text_input("🔑 请输入管理密码", type="password", key="admin_password")
    
    if admin_password != EXPORT_PASSWORD:
        st.warning("🔐 请输入正确密码进行模板管理")
    else:
        st.success("✅ 验证通过，可进行模板管理")
        
        st.subheader("➕ 上传新模板")
        new_template_file = st.file_uploader("选择Excel模板文件（首行为表头）", type=["xlsx", "xls"], key="new_template")
        
        if new_template_file:
            template_name = st.text_input("📝 输入模板名称", placeholder="如：2024年晋升名单")
            
            if st.button("💾 保存模板"):
                if template_name:
                    existing = get_template_files()
                    target_name = f"{template_name}.xlsx"
                    
                    if target_name in existing:
                        st.error("❌ 模板名称已存在！")
                    else:
                        template_path = os.path.join(TEMPLATES_DIR, target_name)
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
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.write(f"📄 {t}")
                with col2:
                    if st.button(f"🗑️ 删除", key=f"del_{t}"):
                        config = load_dropdown_config()
                        if t in config:
                            del config[t]
                            save_dropdown_config(config)
                        os.remove(os.path.join(TEMPLATES_DIR, t))
                        st.success(f"✅ 删除 {t}")
                        st.rerun()
        else:
            st.info("暂无模板，请先上传")

# ========== 标签2：下拉配置 ==========
with tab2:
    st.header("⚙️ 下拉选项配置")
    st.markdown("为指定字段配置下拉选项，生成带下拉验证的模板")
    
    dd_password = st.text_input("🔑 请输入管理密码", type="password", key="dd_password")
    
    if dd_password != EXPORT_PASSWORD:
        st.warning("🔐 请输入正确密码进行配置")
    else:
        templates = get_template_files()
        
        if not templates:
            st.warning("⚠️ 请先在「模板管理」中上传模板")
        else:
            selected_template = st.selectbox("📋 选择要配置的模板", options=templates)
            columns = get_template_columns(selected_template)
            
            if not columns:
                st.error("❌ 无法读取模板字段，请检查模板文件")
            else:
                st.success(f"✅ 模板包含 {len(columns)} 个字段")
                
                config = load_dropdown_config()
                current_config = config.get(selected_template, {})
                
                st.markdown("---")
                st.subheader(f"📝 配置「{selected_template}」的下拉字段")
                
                if current_config:
                    st.info(f"已配置下拉字段：{', '.join(current_config.keys())}")
                
                available_cols = [c for c in columns if c not in current_config]
                col_to_add = st.selectbox("➕ 选择要添加下拉选项的字段", options=available_cols)
                
                if col_to_add:
                    st.markdown("---")
                    st.subheader(f"为「{col_to_add}」设置选项")
                    
                    options_input = st.text_area(
                        "📋 输入下拉选项（每行一个，或用逗号分隔）",
                        placeholder=f"选项1\n选项2\n选项3",
                        height=150,
                        key=f"options_{col_to_add}"
                    )
                    
                    if options_input:
                        parsed_options = parse_dropdown_options(options_input)
                        st.write(f"预览选项（共{len(parsed_options)}个）：", parsed_options[:5], "..." if len(parsed_options) > 5 else "")
                        
                        col_save, col_clear = st.columns(2)
                        with col_save:
                            if st.button(f"💾 保存「{col_to_add}」的下拉配置"):
                                current_config[col_to_add] = parsed_options
                                config[selected_template] = current_config
                                save_dropdown_config(config)
                                st.success(f"✅ 已保存「{col_to_add}」的下拉配置！")
                                st.rerun()
                
                st.markdown("---")
                st.subheader("🗑️ 已配置的下拉字段")
                
                if current_config:
                    for field, opts in current_config.items():
                        col1, col2, col3 = st.columns([2, 3, 1])
                        with col1:
                            st.write(f"**{field}**")
                        with col2:
                            st.caption(f"可选值：{', '.join(opts[:3])}..." if len(opts) > 3 else f"可选值：{', '.join(opts)}")
                        with col3:
                            if st.button("删除", key=f"del_dd_{field}"):
                                del current_config[field]
                                config[selected_template] = current_config
                                save_dropdown_config(config)
                                st.rerun()
                else:
                    st.info("暂无下拉字段配置")

# ========== 标签3：下载模板 ==========
with tab3:
    st.header("📥 下载模板")
    
    templates = get_template_files()
    
    if not templates:
        st.warning("⚠️ 暂无可用模板，请先在「模板管理」中上传模板")
    else:
        selected_template = st.selectbox("📋 请选择模板", options=templates, key="download_template")
        
        columns = get_template_columns(selected_template)
        config = load_dropdown_config()
        dropdown_config = config.get(selected_template, {})
        
        st.markdown("---")
        
        if columns:
            col_info = []
            for col in columns:
                if col in dropdown_config:
                    col_info.append(f"📌 {col}（下拉）")
                else:
                    col_info.append(col)
            st.info(f"📝 该模板包含以下字段：{', '.join(col_info)}")
        
        if dropdown_config:
            st.success("⚠️ 此模板包含下拉字段，下载后可从下拉列表中选择")
        
        if st.button("🔄 生成模板文件"):
            temp_file = f"temp_{selected_template}"
            create_excel_with_dropdown(columns, dropdown_config, temp_file)
            
            with open(temp_file, "rb") as f:
                st.download_button(
                    label=f"⬇️ 下载 {selected_template}",
                    data=f,
                    file_name=selected_template,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            os.remove(temp_file)

# ========== 标签4：上传名单 ==========
with tab4:
    st.header("📤 上传名单")
    
    templates = get_template_files()
    
    if not templates:
        st.warning("⚠️ 暂无可用模板，请先在「模板管理」中上传模板")
    else:
        selected_template = st.selectbox("📋 第一步：选择使用的模板", options=templates, key="upload_template_v2")
        
        template_cols = get_template_columns(selected_template)
        config = load_dropdown_config()
        dropdown_config = config.get(selected_template, {})
        
        if dropdown_config:
            st.info(f"📌 此模板已配置下拉字段：{', '.join(dropdown_config.keys())}")
        
        st.markdown("---")
        st.subheader("📂 第二步：上传名单文件")
        
        uploaded_file = st.file_uploader(
            f"请选择 {selected_template} 对应的Excel文件", 
            type=["xlsx", "xls"],
            key="upload_file_v2"
        )
        
        if uploaded_file is not None:
            try:
                df_upload = pd.read_excel(uploaded_file)
                
                if template_cols:
                    missing_cols = [col for col in template_cols if col not in df_upload.columns]
                    if missing_cols:
                        st.warning(f"⚠️ 文件缺少以下字段：{', '.join(missing_cols)}")
                
                if dropdown_config:
                    invalid_options = []
                    for field, options in dropdown_config.items():
                        if field in df_upload.columns:
                            uploaded_values = df_upload[field].dropna().unique()
                            invalid = [v for v in uploaded_values if str(v) not in options and "请选择" not in str(v)]
                            if invalid:
                                invalid_options.append(f"{field}: {invalid[:3]}...")
                    
                    if invalid_options:
                        st.error(f"❌ 以下字段存在无效选项：{', '.join(invalid_options)}")
                        st.info("💡 请确保所有值都来自下拉列表，或重新下载最新模板")
                    else:
                        st.success("✅ 选项验证通过！")
                
                df_upload["模板名称"] = selected_template
                df_upload["提交时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                st.success("✅ 文件上传成功！以下是您提交的信息：")
                st.dataframe(df_upload.head(10), use_container_width=True)
                
                # ===== 双重备份机制 =====
                # 1. 保存到主CSV
                df_existing = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
                df_combined = pd.concat([df_existing, df_upload], ignore_index=True)
                
                if "工号" in df_combined.columns:
                    df_combined = df_combined.drop_duplicates(subset=["工号"], keep="last")
                elif "姓名" in df_combined.columns:
                    df_combined = df_combined.drop_duplicates(subset=["姓名"], keep="last")
                
                df_combined.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
                
                # 2. 本地备份CSV
                backup_to_local_csv(df_combined)
                
                # 3. GitHub备份（可选）
                if st.button("📤 备份到GitHub"):
                    with st.spinner("正在备份到GitHub..."):
                        if backup_to_github():
                            st.success("✅ 已备份到GitHub！")
                        else:
                            st.warning("⚠️ GitHub备份失败，但数据已保存在本地")
                
                st.balloons()
                st.success(f"✅ 成功提交 {len(df_upload)} 条记录！")
                st.info("💾 数据已自动备份到本地CSV，如需备份到GitHub请点击上方按钮")
                
            except Exception as e:
                st.error(f"❌ 读取文件失败：{str(e)}")

# ========== 标签5：查看/导出 ==========
with tab5:
    st.header("👀 查看与导出")
    
    view_password = st.text_input("🔑 请输入查看密码", type="password", key="view_password_v2")
    
    if view_password != EXPORT_PASSWORD:
        st.warning("🔐 请输入正确密码查看数据")
    else:
        df_all = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
        
        if len(df_all) == 0:
            st.warning("⚠️ 暂无提交记录")
        else:
            if "模板名称" in df_all.columns:
                template_filter = st.selectbox("📋 按模板筛选", ["全部"] + list(df_all["模板名称"].unique()))
                if template_filter != "全部":
                    df_all = df_all[df_all["模板名称"] == template_filter]
            
            col1, col2, col3 = st.columns(3)
            col1.metric("总记录数", len(df_all))
            if "模板名称" in df_all.columns:
                col2.metric("涉及模板数", df_all["模板名称"].nunique())
            col3.metric("最新提交", df_all["提交时间"].max() if "提交时间" in df_all.columns else "N/A")
            
            st.markdown("---")
            
            search_name = st.text_input("🔍 按姓名搜索", "")
            if search_name and "姓名" in df_all.columns:
                df_all = df_all[df_all["姓名"].str.contains(search_name, na=False)]
            
            st.dataframe(df_all, use_container_width=True, height=400)
            
            st.markdown("---")
            st.subheader("📦 导出数据")
            
            col_pass, col_btn = st.columns([2, 1])
            with col_pass:
                password_input = st.text_input("请输入导出密码", type="password", key="export_password_v2")
            with col_btn:
                st.write("")
                if st.button("🔓 导出 Excel"):
                    if password_input == EXPORT_PASSWORD:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        export_filename = f"名单汇总_{timestamp}.xlsx"
                        df_all.to_excel(export_filename, index=False)
                        with open(export_filename, "rb") as f:
                            st.download_button("⬇️ 下载Excel", f, file_name=export_filename)
                    else:
                        st.error("❌ 密码错误！")

# 侧边栏
with st.sidebar:
    st.title("📌 使用说明")
    st.markdown("""
    **备份机制：**
    ✅ 数据自动保存到本地CSV
    📤 可手动备份到GitHub
    
    **步骤：**
    
    1️⃣ **模板管理** - 上传Excel模板
    
    2️⃣ **下拉配置** - 为字段设置下拉选项
    
    3️⃣ **下载模板** - 生成带下拉的模板
    
    4️⃣ **上传名单** - 上传填写好的Excel
    
    5️⃣ **查看/导出** - 查看和导出数据
    
    ---
    
    **密码**：907
    """)
