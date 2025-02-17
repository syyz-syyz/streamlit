import streamlit as st
import pandas as pd
import io
import time

# 设置页面标题和布局
st.set_page_config(page_title="文件处理与 Excel 生成", layout="centered")

# 主标题，设置为深蓝色
st.markdown("<h1 style='color: #00008B;'>文件处理与 Excel 生成</h1>", unsafe_allow_html=True)

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
2. **读取 B 文件**：请上传另一个 XLSX 文件，程序将读取该文件第一个工作表的前两列，分别作为字典和标签。
3. **功能实现**：通过 B 文件中的字典数据，提取 A 文件源数据中的关键词，并将匹配结果整理成新的 Excel 文件供你下载。
4. **备注**：文档切词之前会根据单个的字典长度进行内部排列，确保洗数逻辑是从大到小，从右往左的形式。
""", unsafe_allow_html=True)

# 定义缓存函数
@st.cache_data
def read_a_file(a_file):
    a_df = pd.read_excel(a_file, sheet_name=0, usecols=[0], header=None)
    a_df.columns = ['源数据']
    # 将数字转换为字符串
    a_df['源数据'] = a_df['源数据'].astype(str)
    return a_df

@st.cache_data
def read_b_file(b_file):
    b_df = pd.read_excel(b_file, sheet_name=0, usecols=[0, 1], header=None)
    b_df.columns = ['字典', '标签']
    # 将数字转换为字符串
    b_df['字典'] = b_df['字典'].astype(str)
    # 按字典长度排序，优先匹配较长的字典
    b_df = b_df.sort_values(by='字典', key=lambda x: x.str.len(), ascending=False)
    return b_df

@st.cache_data
def process_data(a_df, b_df):
    # 统计不同长度的数量
    length_zero_count = 0
    length_one_count = 0
    other_length_count = 0

    # 存储不同长度的行
    length_zero_rows = []
    length_one_rows = []
    other_length_rows = []

    # 从末尾开始遍历
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
            # 遇到长度大于 1 的元素，停止遍历
            other_length_count = index + 1
            other_length_rows = b_df.iloc[:index + 1].to_dict(orient='records')
            break

    # 将列表转换为 DataFrame
    length_zero_df = pd.DataFrame(length_zero_rows)
    length_one_df = pd.DataFrame(length_one_rows)
    other_length_df = pd.DataFrame(other_length_rows)

    # 输出统计结果
    st.write(f"字典长度不为 1 和 0 的数量: {other_length_count}")
    st.write(f"字典长度为 1 的数量: {length_one_count}")
    st.write(f"字典长度为 0 的数量: {length_zero_count}")

    # 输出对应的行
    st.write("字典长度为 1 的行:")
    st.dataframe(length_one_df)
    st.write("字典长度为 0 的行:")
    st.dataframe(length_zero_df)

    if length_zero_count > 0:
        # 将 length_zero_rows 转换为 DataFrame
        length_zero_df = pd.DataFrame(length_zero_rows)
        # 从 b_df 中删除这些行
        b_df = b_df[~b_df.index.isin(length_zero_df.index)]

    # 新增小标题
    st.subheader("输出结果导出")

    # 将 B 文件数据存储为字典
    b_dict = {row['字典']: row['标签'] for _, row in b_df.iterrows()}

    # 创建一个空的 DataFrame 来存储结果
    result_data = []

    # 初始化进度条和文字信息
    progress_bar = st.progress(0)
    progress_text = st.empty()  # 用于动态更新文字信息
    total_rows = len(a_df)

    # 记录开始时间
    start_time = time.time()

    # 遍历 a 文件的源数据，查找匹配的关键词
    for index, source_data in enumerate(a_df['源数据']):
        matched = False
        max_length = 0
        latest_match = None
        latest_label = None
        for dict_word, label in b_dict.items():
            word_length = len(dict_word)
            if word_length < max_length:
                # 如果当前字典词长度小于等于已找到的最大匹配长度，跳出内层循环
                break
            last_index = source_data.rfind(dict_word)
            if last_index != -1:
                max_length = word_length
                latest_match = dict_word
                latest_label = label
                matched = True

        if matched:
            result_data.append({
                '源数据': source_data,
                '字典': latest_match,
                '标签': latest_label
            })

        # 计算已用时间和剩余时间
        elapsed_time = time.time() - start_time
        progress = (index + 1) / total_rows
        remaining_time = (elapsed_time / (index + 1)) * (total_rows - (index + 1))

        # 更新进度条和文字信息
        progress_bar.progress(progress)
        progress_text.text(f"处理进度: {index + 1}/{total_rows} | 已用时间: {elapsed_time:.2f}秒 | 剩余时间: {remaining_time:.2f}秒")

    # 将结果转换为 DataFrame
    result_df = pd.DataFrame(result_data, columns=['源数据', '字典', '标签'])

    # 清空进度条和进度文字
    progress_bar.empty()
    progress_text.empty()

    return result_df

# 上传 A 文件
a_file = st.file_uploader("上传 A 文件（XLSX 格式）", type=["xlsx"])

# 上传 B 文件
b_file = st.file_uploader("上传 B 文件（XLSX 格式）", type=["xlsx"])

if a_file and b_file:
    # 读取文件
    a_df = read_a_file(a_file)
    b_df = read_b_file(b_file)

    # 处理数据
    result_df = process_data(a_df, b_df)

    # 显示处理后的前十条结果
    st.subheader("处理后的前十条结果")
    st.dataframe(result_df.head(10))

    # 将结果保存到内存中的 Excel 文件
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        result_df.to_excel(writer, index=False)
    output.seek(0)

    # 提供下载链接
    st.download_button(
        label="下载处理后的 Excel 文件",
        data=output,
        file_name='output.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # 显示处理结果的简单统计信息
    st.info(f"共处理了 {len(a_df)} 条源数据，匹配到 {len(result_df[result_df['标签'].notna()])} 条结果。")
