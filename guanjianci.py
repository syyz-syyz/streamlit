import streamlit as st
import pandas as pd
import io
import time
from functools import lru_cache

# 设置页面标题和布局
st.set_page_config(page_title="轻量化切词小工具", layout="centered")

# 主标题样式
st.markdown("""
<style>
    .stMarkdown h1 {
        color: #00008B !important;
    }
    .stMarkdown h2 {
        color: #ADD8E6 !important;
    }
    .stMarkdown p {
        color: #ADD8E6 !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("轻量化切词小工具")
st.write("### 精准匹配数据，快速生成结果")

# 操作指南（带样式）
st.markdown("""
<div style='color: #ADD8E6; padding: 10px; border-radius: 5px; background-color: rgba(0,0,0,0.05);'>
    <h4>操作指南：</h4>
    <ol>
        <li>上传 A 文件（XLSX 格式）：程序将读取第一个工作表的第一列作为源数据</li>
        <li>上传 B 文件（XLSX 格式）：程序将读取第一个工作表的第一列作为字典，后续列为标签</li>
        <li>功能实现：通过 B 文件字典匹配 A 文件数据，生成带标签的新 Excel 文件</li>
        <li>系统会自动按字典长度（从长到短）进行优先匹配</li>
    </ol>
</div>
""", unsafe_allow_html=True)

# 缓存配置
@st.cache_data(ttl=300)
def read_a_file(a_file):
    try:
        a_df = pd.read_excel(a_file, sheet_name=0, usecols=[0], header=None)
        a_df.columns = ['源数据']
        a_df = a_df[a_df['源数据'].astype(str).str.strip() != '']  # 过滤空字符串
        return a_df
    except Exception as e:
        st.error(f"读取 A 文件失败: {str(e)}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def read_b_file(b_file):
    try:
        b_df = pd.read_excel(b_file, sheet_name=0, header=None)
        if b_df.empty:
            st.error("B 文件内容为空！")
            return pd.DataFrame()
            
        b_df = b_df.rename(columns={0: '字典'})
        num_columns = b_df.shape[1]
        
        # 重命名后续列
        for i in range(1, num_columns):
            b_df.rename(columns={i: f'标签{i}'}, inplace=True)
        
        # 数据清洗
        b_df['字典'] = b_df['字典'].astype(str)
        b_df = b_df[b_df['字典'].str.strip() != '']  # 过滤空字典
        
        # 处理重复项
        duplicate_mask = b_df.duplicated(subset=['字典'], keep=False)
        if duplicate_mask.any():
            duplicates = b_df[duplicate_mask]
            st.warning(f"发现 {len(duplicates)} 条重复字典项，将保留首个出现的条目")
            b_df = b_df.drop_duplicates(subset=['字典'], keep='first')
        
        # 按字典长度排序
        b_df = b_df.sort_values(by='字典', key=lambda x: x.str.len(), ascending=False)
        return b_df
    except Exception as e:
        st.error(f"读取 B 文件失败: {str(e)}")
        return pd.DataFrame()

@st.cache_data(ttl=300)
def process_data(a_df, b_df):
    if b_df.empty:
        st.error("B 文件无有效字典数据！")
        return pd.DataFrame()
    
    # 构建字典（自动填充缺失的标签列）
    max_labels = b_df.shape[1] - 1  # 减去字典列
    b_dict = {
        row['字典']: row.drop('字典').fillna({f'标签{i+1}': None for i in range(max_labels)})
        for _, row in b_df.iterrows()
    }
    
    # 进度条初始化
    progress_bar = st.progress(0)
    progress_text = st.empty()
    result_data = []
    
    try:
        for idx, source in enumerate(a_df['源数据']):
            max_len = 0
            best_match = None
            best_labels = None
            
            # 优化匹配逻辑：优先匹配长词
            for dict_word, labels in b_dict.items():
                word_len = len(dict_word)
                if word_len <= max_len:
                    continue
                    
                start_idx = source.rfind(dict_word)
                if start_idx != -1:
                    max_len = word_len
                    best_match = dict_word
                    best_labels = labels
            
            if best_match:
                row_data = {
                    '源数据': source,
                    '匹配字典': best_match,
                    **best_labels.to_dict()
                }
                # 添加尾部匹配标记
                row_data['是否词尾'] = 'Y' if source.endswith(best_match) else 'N'
                result_data.append(row_data)
            
            # 强制更新进度条（Streamlit渲染机制需要）
            progress = (idx + 1) / len(a_df)
            progress_bar.progress(progress)
            progress_text.text(f"处理进度: {idx+1}/{len(a_df)} 条 | 耗时: {time.time()-start_time:.1f}秒")
            
        # 结果处理
        if not result_data:
            st.warning("未找到任何匹配项！")
            return pd.DataFrame()
        
        result_df = pd.DataFrame(result_data)
        return result_df
    
    except Exception as e:
        st.error(f"处理数据时发生错误: {str(e)}")
        return pd.DataFrame()

# 文件上传组件
with st.sidebar:
    st.header("文件上传")
    a_file = st.file_uploader("上传 A 文件（XLSX 格式）", type=["xlsx"], key="a")
    b_file = st.file_uploader("上传 B 文件（XLSX 格式）", type=["xlsx"], key="b")

# 主处理逻辑
if a_file and b_file:
    with st.spinner("正在加载数据..."):
        a_df = read_a_file(a_file)
        b_df = read_b_file(b_file)
        
        if a_df.empty or b_df.empty:
            st.stop()
        
        with st.spinner("正在处理数据..."):
            result_df = process_data(a_df, b_df)
        
        if not result_df.empty:
            # 显示结果
            st.success(f"✅ 找到 {len(result_df)} 条匹配结果")
            st.subheader("示例结果：")
            st.dataframe(result_df.head(5), use_container_width=True)
            
            # 导出功能
            st.subheader("下载结果文件")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name="匹配结果")
            output.seek(0)
            
            st.download_button(
                label="下载 Excel 文件",
                data=output,
                file_name="matched_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# 系统信息
with st.expander("系统信息"):
    st.write(f"当前 Streamlit 版本: {st.__version__}")
    st.write(f"当前 Pandas 版本: {pd.__version__}")
    st.write(f"服务器时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")
