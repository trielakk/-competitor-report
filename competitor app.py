import os
import pandas as pd
import re
import io
import streamlit as st

# -----------------------------------------------------------------------------
# 1. Streamlit 页面初始化与配置
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="竞品数据处理与汇总工具",
    page_icon="📊",
    layout="wide"
)

st.title("📊 竞品数据自动清洗与汇总工具")
st.markdown("上传分众、白马、SR 或其他竞品公司的 Excel 导出文件，系统将自动清洗字段、补全省份并导出统一的标准竞品表。")

# -----------------------------------------------------------------------------
# 2. 内置城市与省份对照表
# -----------------------------------------------------------------------------
CITY_TO_PROVINCE = {
    # 直辖市
    "北京": "北京", "北京市": "北京",
    "上海": "上海", "上海市": "上海",
    "天津": "天津", "天津市": "天津",
    "重庆": "重庆", "重庆市": "重庆",
    
    # 广东省
    "广州": "广东省", "广州市": "广东省",
    "深圳": "深圳市", "深圳市": "广东省",
    "佛山": "广东省", "佛山市": "广东省",
    "东莞": "广东省", "东莞市": "广东省",
    "中山": "广东省", "中山市": "广东省",
    "珠海": "广东省", "珠海市": "广东省",
    "惠州": "广东省", "惠州市": "广东省",
    "江门": "广东省", "江门市": "广东省",
    "湛江": "广东省", "湛江市": "广东省",
    "汕头": "广东省", "汕头市": "广东省",
    
    # 浙江省
    "杭州": "浙江省", "杭州市": "浙江省",
    "宁波": "浙江省", "宁波市": "浙江省",
    "温州": "浙江省", "温州市": "浙江省",
    "嘉兴": "浙江省", "嘉兴市": "浙江省",
    "金华": "浙江省", "金华市": "浙江省",
    "台州": "浙江省", "台州市": "浙江省",
    "绍兴": "浙江省", "绍兴市": "浙江省",
    
    # 江苏省
    "南京": "江苏省", "南京市": "江苏省",
    "苏州": "江苏省", "苏州市": "江苏省",
    "无锡": "江苏省", "无锡市": "江苏省",
    "常州": "江苏省", "常州市": "江苏省",
    "南通": "江苏省", "南通市": "江苏省",
    "徐州": "江苏省", "徐州市": "江苏省",
    
    # 四川省
    "成都": "四川省", "成都市": "四川省",
    "绵阳": "四川省", "绵阳市": "四川省",
    
    # 湖北省
    "武汉": "湖北省", "武汉市": "湖北省",
    "宜昌": "湖北省", "宜昌市": "湖北省",
    
    # 湖南省
    "长沙": "湖南省", "长沙市": "湖南省",
    "株洲": "湖南省", "株洲市": "湖南省",
    
    # 福建省
    "福州": "福建省", "福州市": "福建省",
    "厦门": "福建省", "厦门市": "福建省",
    "泉州": "福建省", "泉州市": "福建省",
    
    # 山东省
    "济南": "山东省", "济南市": "山东省",
    "青岛": "山东省", "青岛市": "山东省",
    "烟台": "山东省", "烟台市": "山东省",
    "潍坊": "山东省", "潍坊市": "山东省",
    
    # 陕西省
    "西安": "陕西省", "西安市": "陕西省",
    
    # 河南省
    "郑州": "河南省", "郑州市": "河南省",
    "洛阳": "河南省", "洛阳市": "河南省",
    
    # 安徽省
    "合肥": "安徽省", "合肥市": "安徽省",
    "芜湖": "安徽省", "芜湖市": "安徽省",
    
    # 辽宁省
    "沈阳": "辽宁省", "沈阳市": "辽宁省",
    "大连": "辽宁省", "大连市": "辽宁省",
    
    # 吉林省
    "长春": "吉林省", "长春市": "吉林省",
    
    # 黑龙江省
    "哈尔滨": "黑龙江省", "哈尔滨市": "黑龙江省",
    
    # 江西省
    "南昌": "江西省", "南昌市": "江西省",
    
    # 云南省
    "昆明": "云南省", "昆明市": "云南省",
    
    # 广西壮族自治区
    "南宁": "广西壮族自治区", "南宁市": "广西壮族自治区",
    "桂林": "广西壮族自治区", "桂林市": "广西壮族自治区",
    
    # 贵州省
    "贵阳": "贵州省", "贵阳市": "贵州省",
    
    # 山西省
    "太原": "山西省", "太原市": "山西省",
    
    # 河北省
    "石家庄": "河北省", "石家庄市": "河北省",
    "唐山": "河北省", "唐山市": "河北省",
    
    # 海南省
    "海口": "海南省", "海口市": "海南省",
    "三亚": "海南省", "三亚市": "海南省",
}

# -----------------------------------------------------------------------------
# 3. 数据清洗辅助函数
# -----------------------------------------------------------------------------
def get_province_by_city(city_name):
    """根据城市名智能补全省份"""
    if not city_name or pd.isna(city_name):
        return ""
    city_str = str(city_name).strip()
    
    if city_str in CITY_TO_PROVINCE:
        return CITY_TO_PROVINCE[city_str]
    
    for key, prov in CITY_TO_PROVINCE.items():
        if key in city_str or city_str in key:
            return prov
            
    return ""

def find_header_row(df_raw):
    """动态寻找符合表头的行，防止表格上方有前导标题或空白行"""
    for idx, row in df_raw.head(10).iterrows():
        non_nulls = row.dropna().tolist()
        if len(non_nulls) >= 2:
            return idx
    return 0

def load_sheet_smart(file_obj, sheet_name):
    """智能读取Sheet，支持自动定位表头与清除空行"""
    df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
    if df_raw.empty:
        return pd.DataFrame()
        
    header_idx = find_header_row(df_raw)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df

def clean_dataframe_for_display(df):
    """
    核心修复：将 DataFrame 中所有混合数据类型转换为统一字符串
    彻底解决 PyArrow ArrowTypeError 导致 Streamlit 渲染崩溃的问题
    """
    df_clean = df.copy()
    for col in df_clean.columns:
        df_clean[col] = df_clean[col].fillna("").astype(str)
        df_clean[col] = df_clean[col].replace(["nan", "None", "<NA>"], "")
    return df_clean

# -----------------------------------------------------------------------------
# 4. 各媒体公司特定字段解析逻辑
# -----------------------------------------------------------------------------
def parse_focus_media(file_obj):
    records = []
    excel = pd.ExcelFile(file_obj)
    for sheet in excel.sheet_names:
        df = load_sheet_smart(file_obj, sheet)
        if df.empty:
            continue
        
        for _, row in df.iterrows():
            city = row.get("城市", "")
            province = row.get("省份", "") or get_province_by_city(city)
            
            records.append({
                "省份": province,
                "城市": city,
                "媒体公司": "分众公司",
                "媒体形式/名称": row.get("媒体名称", row.get("媒体类型", row.get("形式", ""))),
                "线路/站点/区域": row.get("线路", row.get("区域", row.get("位置", ""))),
                "套装/刊位编号": row.get("编号", row.get("刊位", "")),
                "客户/品牌": row.get("品牌", row.get("客户名称", row.get("客户", ""))),
                "行业": row.get("行业", ""),
                "上刊时间": row.get("上刊时间", row.get("开始时间", "")),
                "下刊时间": row.get("下刊时间", row.get("结束时间", "")),
                "面数/数量": row.get("数量", row.get("面数", 1)),
                "备注": row.get("备注", "")
            })
    return records

def parse_baima(file_obj):
    records = []
    excel = pd.ExcelFile(file_obj)
    for sheet in excel.sheet_names:
        df = load_sheet_smart(file_obj, sheet)
        if df.empty:
            continue
        
        for _, row in df.iterrows():
            city = row.get("城市", row.get("City", ""))
            province = row.get("省份", "") or get_province_by_city(city)
            
            records.append({
                "省份": province,
                "城市": city,
                "媒体公司": "白马公司",
                "媒体形式/名称": row.get("站牌名称", row.get("媒体名称", row.get("媒体类型", "候车亭广告"))),
                "线路/站点/区域": row.get("线路", row.get("站点名称", row.get("站点", row.get("位置", "")))),
                "套装/刊位编号": row.get("站牌编号", row.get("编号", row.get("套装", ""))),
                "客户/品牌": row.get("客户品牌", row.get("客户名称", row.get("品牌", row.get("客户", "")))),
                "行业": row.get("行业分类", row.get("行业", "")),
                "上刊时间": row.get("上刊日期", row.get("上刊时间", row.get("开始日期", ""))),
                "下刊时间": row.get("下刊日期", row.get("下刊时间", row.get("结束日期", ""))),
                "面数/数量": row.get("面数", row.get("数量", row.get("发布看板数", 1))),
                "备注": row.get("备注", "")
            })
    return records

def parse_sr(file_obj):
    records = []
    excel = pd.ExcelFile(file_obj)
    for sheet in excel.sheet_names:
        df = load_sheet_smart(file_obj, sheet)
        if df.empty:
            continue
        
        for _, row in df.iterrows():
            city = row.get("城市", row.get("City", ""))
            province = row.get("省份", "") or get_province_by_city(city)
            
            records.append({
                "省份": province,
                "城市": city,
                "媒体公司": "SR公司",
                "媒体形式/名称": row.get("媒体名称", row.get("媒体形式", row.get("看板类型", "SR媒体"))),
                "线路/站点/区域": row.get("线路/站点", row.get("位置", row.get("站点", row.get("区域", "")))),
                "套装/刊位编号": row.get("位号", row.get("刊位号", row.get("点位编号", row.get("编号", "")))),
                "客户/品牌": row.get("品牌", row.get("客户名称", row.get("广告主", row.get("客户", "")))),
                "行业": row.get("行业", row.get("品类", "")),
                "上刊时间": row.get("上刊时间", row.get("发布时间", row.get("开始时间", ""))),
                "下刊时间": row.get("下刊时间", row.get("撤刊时间", row.get("结束时间", ""))),
                "面数/数量": row.get("数量", row.get("频次", row.get("面数", 1))),
                "备注": row.get("备注", "")
            })
    return records

# -----------------------------------------------------------------------------
# 5. Web 主界面交互逻辑
# -----------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "📎 请选择或拖入要处理的 Excel 文件（支持同时选择多个文件）", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 开始解析并汇总", type="primary"):
        all_records = []
        
        with st.spinner("数据处理中，正在读取 Sheet 并自动映射列字段..."):
            for uploaded_file in uploaded_files:
                filename = uploaded_file.name.lower()
                
                try:
                    # 根据文件名自动选择匹配模版解析，无匹配则采用通用方式解析
                    if "分众" in filename:
                        records = parse_focus_media(uploaded_file)
                    elif "白马" in filename:
                        records = parse_baima(uploaded_file)
                    elif "sr" in filename:
                        records = parse_sr(uploaded_file)
                    else:
                        # 尝试通用解析逻辑
                        records = parse_sr(uploaded_file)
                    
                    all_records.extend(records)
                    st.toast(f"✅ 文件 '{uploaded_file.name}' 解析完成", icon="🎉")
                    
                except Exception as e:
                    st.error(f"❌ 读取文件 '{uploaded_file.name}' 时发生错误: {e}")

        if all_records:
            # 1. 组合结果
            res_df = pd.DataFrame(all_records)
            
            # 2. 补全省份信息
            res_df["省份"] = res_df.apply(
                lambda r: r["省份"] if str(r["省份"]).strip() else get_province_by_city(r["城市"]), 
                axis=1
            )
            
            # 3. 彻底清洗数据类型，解决 PyArrow 异常导致的展示卡死
            res_df_clean = clean_dataframe_for_display(res_df)
            
            st.success(f"🎉 成功完成数据汇总！共提取 {len(res_df_clean)} 条记录。")
            
            # 4. 展示预览数据（适应最新的 Streamlit 规范，使用 width='stretch'）
            st.subheader("📋 汇总结果预览（前 100 条）")
            st.dataframe(res_df_clean.head(100), width='stretch')
            
            # 5. 生成 Excel 内存流用于前端下载
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                res_df_clean.to_excel(writer, index=False, sheet_name='竞品表汇总')
            excel_data = excel_buffer.getvalue()
            
            st.download_button(
                label="📥 点击下载【最终竞品汇总表.xlsx】",
                data=excel_data,
                file_name="1、最终要填写的竞品表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ 未能从上传的文件中提取到有效数据，请检查 Excel 内容。")
else:
    st.info("👆 请在上方选择文件框上传竞品 Excel 报表。")