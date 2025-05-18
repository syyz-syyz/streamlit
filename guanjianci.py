import streamlit as st
import pandas as pd
import io
import time
import re

# 设置页面标题和布局
st.set_page_config(page_title="轻量化切词小工具", layout="centered")


# 主标题，设置为深蓝色
st.markdown("<h1 style='color: #00008B;'>轻量化切词小工具</h1>", unsafe_allow_html=True)

# 副标题，设置为浅蓝色
st.markdown("<h2 style='color: #ADD8E6;'>精准匹配数据，快速生成结果</h2>", unsafe_allow_html=True)

# 提示词，设置为浅蓝色
st.markdown("""
<style>
   .markdown-text-container {
        color: #ADD8E6;
    }
</style>
### 操作指南
1. **读取 A 文件**：请上传一个 XLSX 文件，程序将读取该文件第一个工作表的第一列作为源数据。
2. **读取 B 文件**：请上传另一个 XLSX 文件，程序将读取该文件第一个工作表的第一列作为字典，后面的列作为标签。
3. **功能实现**：通过 B 文件中的字典数据，提取 A 文件源数据中的关键词，并将匹配结果整理成新的 Excel 文件供你下载。
4. **备注**：文档切词之前会根据单个的字典长度进行内部排列，确保洗数逻辑是从大到小，从右往左的形式。
""", unsafe_allow_html=True)

# 定义缓存函数
@st.cache_data
def read_a_file(a_file):
    a_df = pd.read_excel(a_file, sheet_name=0, usecols=[0], header=None)
    a_df.columns = ['源数据']
    a_df['源数据'] = a_df['源数据'].astype(str)
    return a_df

@st.cache_data
def read_b_file(b_file):
    b_df = pd.read_excel(b_file, sheet_name=0, header=None)
    num_columns = b_df.shape[1]
    b_df.rename(columns={0: '字典'}, inplace=True)
    for i in range(1, num_columns):
        b_df.rename(columns={i: f'标签{i}'}, inplace=True)
    b_df['字典'] = b_df['字典'].astype(str)
    
    # 按字典长度降序排序并去重
    b_df = b_df.sort_values(by='字典', key=lambda x: x.str.len(), ascending=False)
    b_df = b_df[~b_df.duplicated(subset=['字典'], keep='first')]
    
    return b_df

@st.cache_data
def process_data(a_df, b_df):
    # 构建正则表达式模式
    sorted_words = b_df['字典'].tolist()
    pattern = r'(%s)' % '|'.join(map(re.escape, sorted_words))
    
    # 创建字典映射
    b_dict = b_df.set_index('字典').to_dict('index')
    
    # 结果存储
    results = []
    total = len(a_df)
    progress_bar = st.progress(0)
    start_time = time.time()

    for idx, source in enumerate(a_df['源数据']):
        match = re.search(pattern, source)
        if match:
            word = match.group(1)
            start_idx = match.start()
            
            # 构建结果行
            result = {
                '源数据': source,
                '字典': word,
                **b_dict[word]
            }
            
            # 判断是否词尾
            result['是否词尾'] = 'Y' if source.endswith(word) else 'N'
            results.append(result)
        
        # 更新进度
        elapsed = time.time() - start_time
        progress = (idx + 1) / total
        remaining = (elapsed / progress) * (total - idx - 1) if progress > 0 else 0
        progress_bar.progress(progress)
        st.session_state['progress_text'] = f"处理进度: {idx+1}/{total} | 已用时间: {elapsed:.1f}s | 剩余时间: {remaining:.1f}s"

    # 清理进度显示
    progress_bar.empty()
    if 'progress_text' in st.session_state:
        st.write(st.session_state['progress_text'])

    return pd.DataFrame(results)

# 文件上传
with st.sidebar:
    a_file = st.file_uploader("上传 A 文件（XLSX 格式）", type=["xlsx"], key="a")
    b_file = st.file_uploader("上传 B 文件（XLSX 格式）", type=["xlsx"], key="b")

if a_file and b_file:
    try:
        a_df = read_a_file(a_file)
        b_df = read_b_file(b_file)
        
        if not b_df.empty:
            result_df = process_data(a_df, b_df)
            
            # 显示结果
            st.subheader("处理结果预览")
            st.dataframe(result_df.head(10))
            
            # 导出功能
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False)
            output.seek(0)
            
            st.download_button(
                label="下载结果文件",
                data=output,
                file_name="切词结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 统计信息
            matched = len(result_df)
            st.success(f"处理完成！共匹配 {matched} 条数据（总数据量：{len(a_df)}）")
        else:
            st.error("B 文件字典为空或格式错误")
    
    except Exception as e:
        st.error(f"处理过程中发生错误：{str(e)}")
        st.info("提示：请确保上传文件不超过1GB，且数据格式正确")
