import streamlit as st
import pandas as pd
import os
from datetime import datetime
import shutil

# 配置页面
st.set_page_config(page_title="同事名单收集", page_icon="📋", layout="wide")

# 数据文件路径
DATA_DIR = "data"
TEMPLATES_DIR = "templates"
DATA_FILE = os.path.join(DATA_DIR, "submissions.csv")
EXPORT_PASSWORD = "907"  # 导出密码

# 创建必要的目录
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

# 初始化CSV文件（如果不存在）
if not os.path.exists(DATA_FILE):
    df = pd.DataFrame(columns=["提交时间", "模板名称"])
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

# ========== 辅助函数 ==========
def get_template_files():
    """获取所有模板文件"""
    templates = []
    for f in os.listdir(TEMPLATES_DIR):
        if f.endswith(('.xlsx', '.xls')):
            templates.append(f)
    return templates

def get_template_columns(template_name):
    """获取模板的列名（排除系统字段）"""
    template_path = os.path.join(TEMPLATES_DIR, template_name)
    try:
        df = pd.read_excel(template_path)
        # 排除可能的系统字段
        return [col for col in df.columns if col not in ["提交时间", "模板名称"]]
    except:
        return []

# 页面标题
st.title("📋 同事名单收集系统")
st.markdown("---")

# 标签页
tab1, tab2, tab3, tab4 = st.tabs([
    "📝 模板管理", 
    "📥 下载模板", 
    "📤 上传名单", 
    "👀 查看/导出"
])

# ========== 标签1：模板管理 ==========
with tab1:
    st.header("📝 模板管理")
    st.markdown("上传新的Excel模板，命名后可被其他用户使用")
    
    # 密码验证
    admin_password = st.text_input("🔑 请输入管理密码", type="password", key="admin_password")
    
    if admin_password != EXPORT_PASSWORD:
        st.warning("🔐 请输入正确密码进行模板管理")
    else:
        st.success("✅ 验证通过，可进行模板管理")
        
        # 上传新模板
        st.subheader("➕ 上传新模板")
        new_template_file = st.file_uploader("选择Excel模板文件", type=["xlsx", "xls"], key="new_template")
        
        if new_template_file:
            template_name = st.text_input("📝 输入模板名称", placeholder="如：2024年晋升名单")
            
            if st.button("💾 保存模板"):
                if template_name:
                    # 检查是否已存在
                    existing = get_template_files()
                    target_name = f"{template_name}.xlsx"
                    
                    if target_name in existing:
                        st.error("❌ 模板名称已存在！")
                    else:
                        # 保存模板文件
                        template_path = os.path.join(TEMPLATES_DIR, target_name)
                        with open(template_path, "wb") as f:
                            f.write(new_template_file.getbuffer())
                        st.success(f"✅ 模板 '{template_name}' 保存成功！")
                        st.rerun()
                else:
                    st.error("❌ 请输入模板名称")
        
        # 查看已有模板
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
                        os.remove(os.path.join(TEMPLATES_DIR, t))
                        st.success(f"✅ 删除 {t}")
                        st.rerun()
        else:
            st.info("暂无模板，请先上传")

# ========== 标签2：下载模板 ==========
with tab2:
    st.header("📥 下载模板")
    
    templates = get_template_files()
    
    if not templates:
        st.warning("⚠️ 暂无可用模板，请先在「模板管理」中上传模板")
    else:
        # 选择模板
        selected_template = st.selectbox(
            "📋 请选择模板",
            options=templates,
            index=0
        )
        
        st.markdown("---")
        
        # 显示模板字段预览
        cols = get_template_columns(selected_template)
        if cols:
            st.info(f"📝 该模板包含以下字段：{', '.join(cols)}")
        
        # 下载按钮
        template_path = os.path.join(TEMPLATES_DIR, selected_template)
        with open(template_path, "rb") as f:
            st.download_button(
                label=f"⬇️ 下载 {selected_template}",
                data=f,
                file_name=selected_template,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ========== 标签3：上传名单 ==========
with tab3:
    st.header("📤 上传名单")
    
    templates = get_template_files()
    
    if not templates:
        st.warning("⚠️ 暂无可用模板，请先在「模板管理」中上传模板")
    else:
        # 第一步：选择模板
        selected_template = st.selectbox(
            "📋 第一步：选择使用的模板",
            options=templates,
            key="upload_template"
        )
        
        # 获取该模板的字段
        template_cols = get_template_columns(selected_template)
        
        # 第二步：上传文件
        st.markdown("---")
        st.subheader("📂 第二步：上传名单文件")
        
        uploaded_file = st.file_uploader(
            f"请选择 {selected_template} 对应的Excel文件", 
            type=["xlsx", "xls"],
            key="upload_file"
        )
        
        if uploaded_file is not None:
            try:
                # 读取上传的Excel
                df_upload = pd.read_excel(uploaded_file)
                
                # 检查必要的列（基于选定的模板）
                # 注意：上传的文件可能包含模板的所有字段或部分字段
                # 我们只要求必须有工号或姓名作为唯一标识
                
                # 添加系统字段
                df_upload["模板名称"] = selected_template
                df_upload["提交时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # 显示上传的数据预览
                st.success("✅ 文件上传成功！以下是您提交的信息：")
                st.dataframe(df_upload.head(10), use_container_width=True)
                
                # 保存到CSV（按工号去重，保留最新，如果没有工号则按姓名）
                df_existing = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
                
                # 合并数据
                df_combined = pd.concat([df_existing, df_upload], ignore_index=True)
                
                # 去重逻辑：优先按工号，其次按姓名
                if "工号" in df_combined.columns:
                    df_combined = df_combined.drop_duplicates(subset=["工号"], keep="last")
                elif "姓名" in df_combined.columns:
                    df_combined = df_combined.drop_duplicates(subset=["姓名"], keep="last")
                
                # 保存
                df_combined.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")
                
                st.balloons()
                st.success(f"✅ 成功提交 {len(df_upload)} 条记录！")
                
            except Exception as e:
                st.error(f"❌ 读取文件失败：{str(e)}")

# ========== 标签4：查看/导出 ==========
with tab4:
    st.header("👀 查看与导出")
    
    # 密码验证
    view_password = st.text_input("🔑 请输入查看密码", type="password", key="view_password")
    
    if view_password != EXPORT_PASSWORD:
        st.warning("🔐 请输入正确密码查看数据")
    else:
        # 读取所有数据
        df_all = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
        
        if len(df_all) == 0:
            st.warning("⚠️ 暂无提交记录")
        else:
            # 按模板筛选
            if "模板名称" in df_all.columns:
                template_filter = st.selectbox(
                    "📋 按模板筛选",
                    options=["全部"] + list(df_all["模板名称"].unique())
                )
                
                if template_filter != "全部":
                    df_all = df_all[df_all["模板名称"] == template_filter]
            
            # 统计信息
            col1, col2, col3 = st.columns(3)
            col1.metric("总记录数", len(df_all))
            if "模板名称" in df_all.columns:
                col2.metric("涉及模板数", df_all["模板名称"].nunique())
            else:
                col2.metric("涉及模板数", 1)
            col3.metric("最新提交", df_all["提交时间"].max() if "提交时间" in df_all.columns else "N/A")
            
            st.markdown("---")
            
            # 搜索过滤
            search_name = st.text_input("🔍 按姓名搜索", "")
            if search_name and "姓名" in df_all.columns:
                df_all = df_all[df_all["姓名"].str.contains(search_name, na=False)]
            
            # 展示数据
            st.dataframe(df_all, use_container_width=True, height=400)
            
            # 导出功能
            st.markdown("---")
            st.subheader("📦 导出数据")
            
            col_pass, col_btn = st.columns([2, 1])
            with col_pass:
                password_input = st.text_input("请输入导出密码", type="password", key="export_password")
            with col_btn:
                st.write("")
                st.write("")
                if st.button("🔓 导出 Excel"):
                    if password_input == EXPORT_PASSWORD:
                        # 添加时间戳到文件名
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        export_filename = f"名单汇总_{timestamp}.xlsx"
                        
                        df_all.to_excel(export_filename, index=False)
                        with open(export_filename, "rb") as f:
                            st.download_button(
                                label="⬇️ 下载Excel",
                                data=f,
                                file_name=export_filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error("❌ 密码错误！")

# 侧边栏
with st.sidebar:
    st.title("📌 使用说明")
    st.markdown("""
    **步骤：**
    
    1️⃣ **模板管理** - 管理员上传Excel模板
    
    2️⃣ **下载模板** - 选择需要的模板下载
    
    3️⃣ **上传名单** - 选择模板，上传填写好的Excel
    
    4️⃣ **查看/导出** - 输入密码查看和导出数据
    
    ---
    
    **密码**：907
    """)
