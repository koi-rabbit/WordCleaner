import streamlit as st
from docx import Document
import re, os
from io import BytesIO
# 下面这 3 行照抄你原来文件里的常量/函数即可
from your_original_script import (
    add_heading_numbers,
    modify_document_format,
    get_outline_level_from_xml
)

st.set_page_config(page_title="Word 自动排版", layout="centered")
st.title("📄 Word 自动排版工具")
st.markdown("上传一份 `.docx`，程序会：\n"
            "1. 根据大纲级别自动套用 Heading 1-9；\n"
            "2. 按规范重新编号；\n"
            "3. 统一字体、字号、段前段后等格式；\n"
            "4. 生成可下载的新文件。")

uploaded = st.file_uploader("请选择 Word 文件", type=["docx"])
if uploaded is None:
    st.stop()

if st.button("开始排版"):
    with st.spinner("正在处理…"):
        # ① 读进内存
        doc = Document(BytesIO(uploaded.read()))

        # ② 把 Normal 段落按大纲级别改成 Heading 1-9（你原来的逻辑）
        for para in doc.paragraphs:
            lvl = get_outline_level_from_xml(para)
            if lvl and para.style.name == "Normal":
                para.style = doc.styles[f"Heading {lvl}"]

        # ③ 编号 + 格式
        add_heading_numbers(doc)
        modify_document_format(doc)

        # ④ 写回内存
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

    st.success("处理完成！")
    st.download_button(
        label="📥 下载已排版文件",
        data=buffer,
        file_name=f"{uploaded.name.stem}_已排版.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
