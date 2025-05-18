import streamlit as st
import pandas as pd
import io
import time
import pyarrow as pa

# 初始化会话状态
if 'a_df' not in st.session_state:
    st.session_state.a_df = None
if 'b_df' not in st.session_state:
    st.session_state.b_df = None
if 'result_df' not in st.session_state:
    st.session_state.result_df = None

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
@st.cache_data(ttl=3600, max_entries=10)
def read_a_file(a_file):
    try:
        a_df = pd.read_excel(a_file, sheet_name=0, usecols=[0], header=None)
        a_df.columns = ['源数据']
        a_df['源数据'] = a_df['源数据'].astype(str)
        return a_df
    except Exception as e:
        st.error(f"读取 A 文件时出错: {str(e)}")
        return pd.DataFrame()

@st.cache_data(ttl=3600, max_entries=10)
def read_b_file(b_file):
    try:
        b_df = pd.read_excel(b_file, sheet_name=0, header=None)
        num_columns = b_df.shape[1]
        b_df.rename(columns={0: '字典'}, inplace=True)
        for i in range(1, num_columns):
            b_df.rename(columns={i: f'标签{i}'}, inplace=True)
        b_df['字典'] = b_df['字典'].astype(str)
        b_df = b_df.sort_values(by='字典', key=lambda x: x.str.len(), ascending=False)

        duplicate_mask = b_df['字典'].duplicated(keep=False)
        duplicate_rows = b_df[duplicate_mask]

        if not duplicate_rows.empty:
            st.warning("发现字典中有重复元素，以下是重复的行：")
            st.dataframe(duplicate_rows)
            st.info("将选用重复行中上面出现的第一条数据进行后续处理。")
            b_df = b_df[~b_df['字典'].duplicated(keep='first')]

        return b_df
    except Exception as e:
        st.error(f"读取 B 文件时出错: {str(e)}")
        return pd.DataFrame()

def process_batch(a_batch, b_df):
    """处理单个数据批次"""
    b_dict = {row['字典']: row.drop('字典') for _, row in b_df.iterrows()}
    result_data = []
    
    for _, source_data in a_batch['源数据'].items():
        matched = False
        max_length = 0
        latest_match = None
        latest_labels = None
        max_start_index = -1
        
        for dict_word, labels in b_dict.items():
            word_length = len(dict_word)
            if word_length < max_length:
                break
            start_index = source_data.rfind(dict_word)
            if start_index != -1:
                if start_index > max_start_index:
                    max_length = word_length
                    latest_match = dict_word
                    latest_labels = labels
                    max_start_index = start_index
                matched = True

        if matched:
            result_row = {
                '源数据': source_data,
                '字典': latest_match
            }
            result_row.update(latest_labels)
            result_data.append(result_row)
    
    return pd.DataFrame(result_data)

def process_data_in_batches(a_df, b_df, batch_size=1000):
    """分批处理数据，减少内存压力"""
    if a_df.empty or b_df.empty:
        st.error("无法处理空数据。请检查上传的文件是否有效。")
        return pd.DataFrame()
    
    # 统计不同长度的数量
    length_zero_count = 0
    length_one_count = 0
    other_length_count = 0

    length_zero_rows = []
    length_one_rows = []
    other_length_rows = []

    for index in range(len(b_df) - 1, -1, -1):
        row = b_df.iloc[index]
        dict_length = len(str(row['字典']))
        value = row['字典']
        if pd.isna(value):
            length_zero_count += 1
            length_zero_rows.append(row)
        elif dict_length == 1:
            length_one_count += 1
            length_one_rows.append(row)
        else:
            other_length_count = index + 1
            other_length_rows = b_df.iloc[:index + 1].to_dict(orient='records')
            break

    length_zero_df = pd.DataFrame(length_zero_rows)
    length_one_df = pd.DataFrame(length_one_rows)
    other_length_df = pd.DataFrame(other_length_rows)

    st.write(f"字典长度不为 1 和 0 的数量: {other_length_count}")
    st.write(f"字典长度为 1 的数量: {length_one_count}")
    st.write(f"字典长度为 0 的数量: {length_zero_count}")

    st.write("字典长度为 1 的行:")
    st.dataframe(length_one_df)
    st.write("字典长度为 0 的行:")
    st.dataframe(length_zero_df)

    if length_zero_count > 0:
        length_zero_df = pd.DataFrame(length_zero_rows)
        b_df = b_df[~b_df.index.isin(length_zero_df.index)]

    st.subheader("输出结果导出")
    
    all_results = []
    total_batches = len(a_df) // batch_size + (1 if len(a_df) % batch_size > 0 else 0)
    
    progress_bar = st.progress(0)
    progress_text = st.empty()
    
    for batch_num in range(total_batches):
        start_idx = batch_num * batch_size
        end_idx = min((batch_num + 1) * batch_size, len(a_df))
        batch_df = a_df.iloc[start_idx:end_idx]
        
        batch_result = process_batch(batch_df, b_df)
        all_results.append(batch_result)
        
        progress = (batch_num + 1) / total_batches
        progress_bar.progress(progress)
        progress_text.text(f"处理批次: {batch_num + 1}/{total_batches}")
    
    progress_bar.empty()
    progress_text.empty()
    
    result_df = pd.concat(all_results, ignore_index=True)
    
    # 新增功能：根据 B 文件第六列数字切分源数据并对比
    if '标签5' in b_df.columns and not result_df.empty:
        b_cut_dict = {row['字典']: row['标签5'] for _, row in b_df.iterrows()}
        result_df['是否词尾'] = ''
        for index, row in result_df.iterrows():
            source_data = row['源数据']
            dict_word = row['字典']
            cut_num = b_cut_dict.get(dict_word)
            if pd.notna(cut_num):
                cut_num = int(cut_num)
                if len(source_data) >= cut_num:
                    right_part = source_data[-cut_num:]
                    if right_part == dict_word:
                        result_df.at[index, '是否词尾'] = 'Y'
                    else:
                        result_df.at[index, '是否词尾'] = 'N'
    
    return result_df

def to_feather_bytes(df):
    """将 DataFrame 转换为 feather 格式的字节流"""
    buffer = pa.BufferOutputStream()
    df.to_feather(buffer)
    return buffer.getvalue().to_pybytes()

# 上传 A 文件
a_file = st.file_uploader("上传 A 文件（XLSX 格式）", type=["xlsx"])

# 上传 B 文件
b_file = st.file_uploader("上传 B 文件（XLSX 格式）", type=["xlsx"])

if a_file and b_file:
    try:
        if (a_file != st.session_state.get('last_a_file') or 
            b_file != st.session_state.get('last_b_file') or
            st.session_state.a_df is None or 
            st.session_state.b_df is None):
            
            st.session_state.a_df = read_a_file(a_file)
            st.session_state.b_df = read_b_file(b_file)
            st.session_state.last_a_file = a_file
            st.session_state.last_b_file = b_file
            
            with st.spinner("正在处理数据..."):
                st.session_state.result_df = process_data_in_batches(
                    st.session_state.a_df, 
                    st.session_state.b_df,
                    batch_size=1000  # 可根据数据量调整批次大小
                )
        
        if st.session_state.result_df is not None and not st.session_state.result_df.empty:
            st.subheader("处理后的前十条结果")
            st.dataframe(st.session_state.result_df.head(10))

            # 提供多种格式下载选项
            col1, col2 = st.columns(2)
            
            # Excel 格式
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                st.session_state.result_df.to_excel(writer, index=False)
            output_excel.seek(0)
            
            # Feather 格式（更高效）
            output_feather = to_feather_bytes(st.session_state.result_df)
            
            col1.download_button(
                label="下载 Excel 文件",
                data=output_excel,
                file_name='output.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
            col2.download_button(
                label="下载 Feather 文件（更快）",
                data=output_feather,
                file_name='output.feather',
                mime='application/octet-stream'
            )

            matched_count = len(st.session_state.result_df[st.session_state.result_df[st.session_state.result_df.columns[2:]].notna().any(axis=1)])
            st.info(f"共处理了 {len(st.session_state.a_df)} 条源数据，匹配到 {matched_count} 条结果。")
        else:
            st.warning("处理结果为空。请检查输入数据是否符合预期。")
            
    except Exception as e:
        st.error(f"发生错误: {str(e)}")
        import traceback
        st.text(traceback.format_exc())  # 显示详细的错误堆栈
