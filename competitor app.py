import os
import pandas as pd
import streamlit as st
import io

# 页面配置
st.set_page_config(page_title="竞品数据汇总工具", layout="wide")
st.title("📊 竞品数据自动汇总工具")

# 目标公司关键字映射
COMPANY_TARGETS = {
    "白马": "白马公司",
    "SR": "SR公司",
    "sr": "SR公司"
}

def find_header_row(df_raw):
    """动态寻找表头行"""
    for idx, row in df_raw.head(10).iterrows():
        non_nulls = row.dropna().tolist()
        if len(non_nulls) >= 2:
            return idx
    return 0

def process_uploaded_files(uploaded_files):
    summary_list = []
    
    for uploaded_file in uploaded_files:
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheet_names = excel_file.sheet_names
            
            for sheet in sheet_names:
                clean_sheet_name = sheet.strip()
                matched_company = None
                
                # 匹配公司关键字
                for keyword, mapped_name in COMPANY_TARGETS.items():
                    if keyword.lower() in clean_sheet_name.lower():
                        matched_company = mapped_name
                        break
                
                if not matched_company:
                    continue
                
                # 读取表头和数据
                df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                if df_raw.empty:
                    continue
                    
                header_idx = find_header_row(df_raw)
                df = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_idx)
                
                # 清理数据列
                df.columns = [str(col).strip() for col in df.columns]
                df = df.dropna(how="all")
                
                if df.empty:
                    continue
                    
                df["归属公司"] = matched_company
                df["来源Sheet"] = sheet
                df["来源文件"] = uploaded_file.name
                summary_list.append(df)
                
        except Exception as e:
            st.error(f"处理文件 {uploaded_file.name} 时出错: {e}")
            
    return summary_list

# 文件上传界面
uploaded_files = st.file_uploader(
    "请上传需要汇总的 Excel 文件（支持多选）", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 开始汇总解析"):
        with st.spinner("正在解析数据，请稍候..."):
            results = process_uploaded_files(uploaded_files)
            
            if results:
                final_df = pd.concat(results, ignore_index=True)
                st.success(f"解析成功！共汇总 {len(final_df)} 行数据。")
                
                # 数据预览
                st.subheader("📋 汇总预览")
                st.dataframe(final_df.head(50), use_container_width=True)
                
                # 提供 Excel 下载链接
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='汇总数据')
                
                st.download_button(
                    label="📥 下载汇总 Excel 文件",
                    data=output.getvalue(),
                    file_name="竞品数据汇总结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("未能匹配到包含目标公司关键字（如'白马'、'SR'）的 Sheet 页。")
else:
    st.info("请在上方上传 Excel 文件后开始使用。")