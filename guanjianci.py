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
""", unsafe_allow_html=True)

# 初始化 session_state
if "a_file_read" not in st.session_state:
    st.session_state.a_file_read = False
if "previous_a_file" not in st.session_state:
    st.session_state.previous_a_file = None

# 上传 A 文件
a_file = st.file_uploader("上传 A 文件（XLSX 格式）", type=["xlsx"])

# 检查是否重新上传了 A 文件
if a_file and a_file != st.session_state.previous_a_file:
    st.session_state.a_file_read = False
    st.session_state.previous_a_file = a_file

if a_file and not st.session_state.a_file_read:
    # 读取 a 文件的第一列
    a_progress_bar = st.progress(0)
    a_progress_text = st.empty()
    a_start_time = time.time()
    a_df = pd.read_excel(a_file, sheet_name=0, usecols=[0], header=None)
    a_total_rows = len(a_df)
    for i in range(a_total_rows):
        progress = (i + 1) / a_total_rows
        elapsed_time = time.time() - a_start_time
        remaining_time = (elapsed_time / (i + 1)) * (a_total_rows - (i + 1))
        a_progress_bar.progress(progress)
        a_progress_text.text(f"A 文件读取进度: {i + 1}/{a_total_rows} | 已用时间: {elapsed_time:.2f}秒 | 剩余时间: {remaining_time:.2f}秒")
    a_df.columns = ['源数据']
    st.session_state.a_file_read = True
elif a_file and st.session_state.a_file_read:
    a_df = pd.read_excel(a_file, sheet_name=0, usecols=[0], header=None)
    a_df.columns = ['源数据']

# 上传 B 文件
b_file = st.file_uploader("上传 B 文件（XLSX 格式）", type=["xlsx"])

# 初始化其他 session_state
if "result_df" not in st.session_state:
    st.session_state.result_df = None
if "total_rows" not in st.session_state:
    st.session_state.total_rows = 0
if "previous_b_file" not in st.session_state:
    st.session_state.previous_b_file = None

# 检查是否上传了新文件
if a_file and b_file:
    if (a_file != st.session_state.previous_a_file) or (b_file != st.session_state.previous_b_file):
        # 如果上传了新文件，清空缓存结果
        st.session_state.result_df = None
        st.session_state.total_rows = 0
        st.session_state.previous_a_file = a_file
        st.session_state.previous_b_file = b_file

    # 如果结果尚未处理，则进行处理
    if st.session_state.result_df is None:
        # 读取 b 文件的前两列
        b_progress_bar = st.progress(0)
        b_progress_text = st.empty()
        b_start_time = time.time()
        b_df = pd.read_excel(b_file, sheet_name=0, usecols=[0, 1], header=None)
        b_total_rows = len(b_df)
        for i in range(b_total_rows):
            progress = (i + 1) / b_total_rows
            elapsed_time = time.time() - b_start_time
            remaining_time = (elapsed_time / (i + 1)) * (b_total_rows - (i + 1))
            b_progress_bar.progress(progress)
            b_progress_text.text(f"B 文件读取进度: {i + 1}/{b_total_rows} | 已用时间: {elapsed_time:.2f}秒 | 剩余时间: {remaining_time:.2f}秒")
        b_df.columns = ['字典', '标签']

        # 删除包含空值的行
        b_df = b_df.dropna()

        # 统计 B 文件两列中单个词和空值的数量
        dict_non_null_count = b_df['字典'].count()
        dict_null_count = len(b_df) - dict_non_null_count
        label_non_null_count = b_df['标签'].count()
        label_null_count = len(b_df) - label_non_null_count

        # 统计字典列中单个中文单元格的数量
        single_chinese_count = b_df[b_df['字典'].str.match(r'^[\u4e00-\u9fff]$', na=False)].shape[0]
        label_chinese_count = b_df[b_df['标签'].str.match(r'^[\u4e00-\u9fff]$', na=False)].shape[0]

        st.write("B 文件数据统计：")
        st.write(f"字典列 - 非空值数量: {dict_non_null_count}, 空值数量: {dict_null_count}, 单个中文单元格数量: {single_chinese_count}")
        st.write(f"标签列 - 非空值数量: {label_non_null_count}, 空值数量: {label_null_count}, 单个中文单元格数量: {label_chinese_count}")

        # 输出空值和单个中文单元格的数据
        st.subheader("B 文件空值和单个中文单元格的数据")

        # 空值数据
        st.write("空值数据：")
        null_data = b_df[(b_df['字典'].isna()) | (b_df['标签'].isna())]
        st.dataframe(null_data)

        # 单个中文单元格数据
        st.write("单个中文单元格数据：")
        single_char_data = b_df[(b_df['字典'].str.len() == 1) & (b_df['字典'].str.match(r'^[\u4e00-\u9fff]$', na=False)) | (b_df['标签'].str.len() == 1) & (b_df['标签'].str.match(r'^[\u4e00-\u9fff]$', na=False))]
        st.dataframe(single_char_data)

        # 新增小标题
        st.subheader("输出结果导出")

        # 按字典长度排序，优先匹配较长的字典
        b_df = b_df.sort_values(by='字典', key=lambda x: x.str.len(), ascending=False)

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
            current_length = None
            same_length_dicts = []
            for _, row in b_df.iterrows():
                dict_word = row['字典']
                if current_length is None or len(dict_word) == current_length:
                    current_length = len(dict_word)
                    same_length_dicts.append(row)
                else:
                    # 处理相同长度的字典词，从右到左匹配
                    for same_length_row in reversed(same_length_dicts):
                        dict_word = same_length_row['字典']
                        if dict_word in source_data:
                            result_data.append({
                                '源数据': source_data,
                                '字典': dict_word,
                                '标签': same_length_row['标签']
                            })
                            matched = True
                            break
                    if matched:
                        break
                    same_length_dicts = [row]
                    current_length = len(dict_word)

            # 处理最后一组相同长度的字典词
            if not matched:
                for same_length_row in reversed(same_length_dicts):
                    dict_word = same_length_row['字典']
                    if dict_word in source_data:
                        result_data.append({
                            '源数据': source_data,
                            '字典': dict_word,
                            '标签': same_length_row['标签']
                        })
                        matched = True
                        break

            if not matched:
                result_data.append({
                    '源数据': source_data,
                    '字典': None,
                    '标签': None
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

        # 将结果存储到 session_state
        st.session_state.result_df = result_df
        st.session_state.total_rows = total_rows

    # 显示处理后的前十条结果
    st.subheader("处理后的前十条结果")
    st.dataframe(st.session_state.result_df.head(10))

    # 将结果保存到内存中的 Excel 文件
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.result_df.to_excel(writer, index=False)
    output.seek(0)

    # 提供下载链接
    st.download_button(
        label="下载处理后的 Excel 文件",
        data=output,
        file_name='output.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # 显示处理结果的简单统计信息
    st.info(f"共处理了 {st.session_state.total_rows} 条源数据，匹配到 {len(st.session_state.result_df[st.session_state.result_df['标签'].notna()])} 条结果。")