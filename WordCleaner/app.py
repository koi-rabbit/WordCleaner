import streamlit as st
import re
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.shared import Inches
from io import BytesIO
import json

# 设置页面配置
st.set_page_config(
    page_title="Word文档格式化工具",
    page_icon="📝",
    layout="wide"
)

# 默认配置
DEFAULT_CONFIG = {
    "title_settings": {
        "apply_numbering": True,
        "max_levels": 9,
        "numbering_formats": {
            1: "chinese",  # 一、
            2: "chinese_bracket",  # （一）
            3: "arabic_dot",  # 1.
            4: "arabic_bracket",  # （1）
            5: "arabic_dot",  # 1.
            6: "arabic_bracket",  # （1）
            7: "arabic_dot",  # 1.
            8: "arabic_bracket",  # （1）
            9: "arabic_dot",  # 1.
        }
    },
    "body_settings": {
        "cz_font_name": "宋体",
        "font_name": "Times New Roman",
        "font_size": 12,
        "space_before": 12,
        "space_after": 12,
        "line_spacing": 1.5,
        "first_line_indent": 0.5
    },
    "table_settings": {
        "cz_font_name": "宋体",
        "font_name": "Times New Roman",
        "font_size": 10,
        "space_before": 6,
        "space_after": 6,
        "width": 6
    }
}

# 初始化session state
if 'config' not in st.session_state:
    st.session_state.config = DEFAULT_CONFIG.copy()
if 'processed' not in st.session_state:
    st.session_state.processed = False

# 样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.3rem;
        font-weight: 600;
        color: #374151;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
        border-bottom: 2px solid #E5E7EB;
        padding-bottom: 0.5rem;
    }
    .upload-box {
        border: 2px dashed #4F46E5;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        margin: 1rem 0;
    }
    .upload-box:hover {
        background: linear-gradient(135deg, #5a6fd8 0%, #6a4090 100%);
        border-color: #4338CA;
    }
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 1.5rem;
        border-radius: 8px;
        font-weight: 600;
        width: 100%;
        transition: all 0.3s ease;
    }
    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 20px rgba(0,0,0,0.2);
        background: linear-gradient(135deg, #5a6fd8 0%, #6a4090 100%);
    }
    .config-section {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #E5E7EB;
        margin-bottom: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .success-box {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        margin: 1rem 0;
    }
    .info-box {
        background: linear-gradient(135deg, #3B82F6 0%, #1D4ED8 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def get_outline_level_from_xml(p):
    """从段落的XML中提取大纲级别，并加1"""
    xml = p._p.xml
    m = re.search(r'<w:outlineLvl w:val="(\d)"/>', xml)
    level = int(m.group(1)) if m else None
    if level is not None:
        level += 1
    return level

def set_font(run, cz_font_name, font_name):
    """设置字体"""
    rPr = run.element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), cz_font_name)
    rFonts.set(qn('w:ascii'), font_name)

def number_to_chinese(number):
    """数字转中文"""
    if number < 0 or number > 100:
        raise ValueError("数字必须在0到100之间")
    
    chinese_numbers = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    chinese_units = ["", "十", "百"]
    
    if number < 10:
        return chinese_numbers[number]
    elif number < 20:
        return "十" + (chinese_numbers[number - 10] if number != 10 else "")
    elif number < 100:
        tens = number // 10
        ones = number % 10
        return chinese_numbers[tens] + "十" + (chinese_numbers[ones] if ones != 0 else "")
    else:
        return "一百"

def format_number(level, number, format_type):
    """根据格式类型格式化数字"""
    formats = {
        "chinese": lambda n: f"{number_to_chinese(n)}、",
        "chinese_bracket": lambda n: f"（{number_to_chinese(n)}）",
        "arabic_dot": lambda n: f"{n}.",
        "arabic_bracket": lambda n: f"（{n}）",
        "roman_lower": lambda n: f"{to_roman(n).lower()}.",
        "roman_upper": lambda n: f"{to_roman(n)}.",
        "alphabet_lower": lambda n: f"{chr(96 + n)}." if n <= 26 else f"{n}.",
        "alphabet_upper": lambda n: f"{chr(64 + n)}." if n <= 26 else f"{n}.",
    }
    return formats.get(format_type, lambda n: f"{n}.")(number)

def to_roman(num):
    """转换为罗马数字"""
    roman_map = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
                 (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
                 (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
    result = ""
    for value, numeral in roman_map:
        while num >= value:
            result += numeral
            num -= value
    return result

def add_heading_numbers(doc, config):
    """根据配置添加标题序号"""
    if not config["title_settings"]["apply_numbering"]:
        return
    
    max_levels = config["title_settings"]["max_levels"]
    heading_numbers = [0] * max_levels
    numbering_formats = config["title_settings"]["numbering_formats"]
    number_pattern = re.compile(r'^[\d一二三四五六七八九十（）\.、\s]+')

    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Heading'):
            try:
                level = int(paragraph.style.name.split(' ')[1]) - 1
                if level >= max_levels:
                    continue
                    
                paragraph.text = number_pattern.sub('', paragraph.text).strip()
                heading_numbers[level] += 1
                for i in range(level + 1, len(heading_numbers)):
                    heading_numbers[i] = 0
                
                format_type = numbering_formats.get(level + 1, "arabic_dot")
                number_str = format_number(level, heading_numbers[level], format_type)
                paragraph.text = number_str + paragraph.text
            except Exception:
                continue

def modify_document_format(doc, config):
    """修改文档格式"""
    body = config["body_settings"]
    table = config["table_settings"]
    
    # 处理正文
    for paragraph in doc.paragraphs:
        if not paragraph.style.name.startswith("Heading"):
            paragraph.paragraph_format.space_before = Pt(body['space_before'])
            paragraph.paragraph_format.space_after = Pt(body['space_after'])
            paragraph.paragraph_format.line_spacing = body['line_spacing']
            paragraph.paragraph_format.first_line_indent = Inches(body['first_line_indent'])
            for run in paragraph.runs:
                set_font(run, body['cz_font_name'], body['font_name'])
                run.font.size = Pt(body['font_size'])

    # 处理表格
    for table_obj in doc.tables:
        table_obj.width = Inches(table['width'])
        for row in table_obj.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        set_font(run, table['cz_font_name'], table['font_name'])
                        run.font.size = Pt(table['font_size'])
                    paragraph.paragraph_format.space_before = Pt(table['space_before'])
                    paragraph.paragraph_format.space_after = Pt(table['space_after'])

def process_document(uploaded_file, config):
    """处理上传的文档"""
    try:
        doc = Document(uploaded_file)
        
        # 转换大纲级别为标题样式
        for para in doc.paragraphs:
            outline_level = get_outline_level_from_xml(para)
            if outline_level is not None and para.style.name == 'Normal':
                if outline_level <= 9:
                    heading_style = f"Heading {outline_level}"
                    if heading_style in doc.styles:
                        para.style = doc.styles[heading_style]
        
        # 添加标题序号
        add_heading_numbers(doc, config)
        
        # 修改格式
        modify_document_format(doc, config)
        
        # 保存到内存
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"处理失败: {str(e)}")
        return None

def main():
    # 主标题
    st.markdown('<h1 class="main-header">📝 Word文档格式化工具</h1>', unsafe_allow_html=True)
    
    # 创建两列布局
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # 文件上传区域
        st.markdown('<div class="sub-header">📤 上传文档</div>', unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "",
            type=['docx'],
            help="上传需要格式化的Word文档",
            label_visibility="collapsed"
        )
        
        if uploaded_file:
            st.markdown(f'<div class="info-box">📄 已上传: {uploaded_file.name}<br>大小: {len(uploaded_file.getvalue()) / 1024:.1f} KB</div>', unsafe_allow_html=True)
    
    with col2:
        # 配置区域
        st.markdown('<div class="sub-header">⚙️ 基本设置</div>', unsafe_allow_html=True)
        
        with st.container():
            st.markdown('<div class="config-section">', unsafe_allow_html=True)
            
            # 标题设置
            st.markdown("**📝 标题设置**")
            col_a, col_b = st.columns(2)
            with col_a:
                apply_num = st.toggle("添加序号", value=st.session_state.config["title_settings"]["apply_numbering"])
                st.session_state.config["title_settings"]["apply_numbering"] = apply_num
            
            with col_b:
                if apply_num:
                    max_levels = st.select_slider("最大级别", options=list(range(1, 10)), value=st.session_state.config["title_settings"]["max_levels"])
                    st.session_state.config["title_settings"]["max_levels"] = max_levels
            
            st.divider()
            
            # 正文设置
            st.markdown("**📄 正文设置**")
            col_c, col_d = st.columns(2)
            with col_c:
                st.session_state.config["body_settings"]["font_size"] = st.number_input("字号", min_value=6, max_value=72, value=int(st.session_state.config["body_settings"]["font_size"]))
                st.session_state.config["body_settings"]["line_spacing"] = st.number_input("行距", min_value=1.0, max_value=3.0, value=float(st.session_state.config["body_settings"]["line_spacing"]), step=0.1)
            
            with col_d:
                st.session_state.config["body_settings"]["first_line_indent"] = st.number_input("缩进(英寸)", min_value=0.0, max_value=2.0, value=float(st.session_state.config["body_settings"]["first_line_indent"]), step=0.1)
            
            st.divider()
            
            # 表格设置
            st.markdown("**📊 表格设置**")
            col_e, col_f = st.columns(2)
            with col_e:
                st.session_state.config["table_settings"]["font_size"] = st.number_input("表格字号", min_value=6, max_value=72, value=int(st.session_state.config["table_settings"]["font_size"]))
            
            with col_f:
                st.session_state.config["table_settings"]["width"] = st.number_input("表格宽度", min_value=1, max_value=20, value=int(st.session_state.config["table_settings"]["width"]))
            
            st.markdown('</div>', unsafe_allow_html=True)
    
    # 高级设置（可折叠）
    with st.expander("⚙️ 高级设置", expanded=False):
        tab1, tab2, tab3 = st.tabs(["标题格式", "字体设置", "间距设置"])
        
        with tab1:
            if st.session_state.config["title_settings"]["apply_numbering"]:
                max_levels = st.session_state.config["title_settings"]["max_levels"]
                numbering_options = {
                    "chinese": "一、二、三",
                    "chinese_bracket": "（一）（二）（三）",
                    "arabic_dot": "1.2.3.",
                    "arabic_bracket": "（1）（2）（3）",
                    "roman_lower": "i.ii.iii.",
                    "roman_upper": "I.II.III.",
                    "alphabet_lower": "a.b.c.",
                    "alphabet_upper": "A.B.C."
                }
                
                cols = st.columns(min(3, max_levels))
                for level in range(1, max_levels + 1):
                    with cols[(level-1) % 3]:
                        current = st.session_state.config["title_settings"]["numbering_formats"].get(level, "arabic_dot")
                        selected = st.selectbox(
                            f"第{level}级格式",
                            options=list(numbering_options.keys()),
                            format_func=lambda x: numbering_options[x],
                            index=list(numbering_options.keys()).index(current) if current in numbering_options else 0,
                            key=f"format_{level}"
                        )
                        st.session_state.config["title_settings"]["numbering_formats"][level] = selected
        
        with tab2:
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**正文字体**")
                st.session_state.config["body_settings"]["cz_font_name"] = st.text_input("中文字体", value=st.session_state.config["body_settings"]["cz_font_name"])
                st.session_state.config["body_settings"]["font_name"] = st.text_input("英文字体", value=st.session_state.config["body_settings"]["font_name"])
            
            with col2:
                st.markdown("**表格字体**")
                st.session_state.config["table_settings"]["cz_font_name"] = st.text_input("表格中文字体", value=st.session_state.config["table_settings"]["cz_font_name"])
                st.session_state.config["table_settings"]["font_name"] = st.text_input("表格英文字体", value=st.session_state.config["table_settings"]["font_name"])
        
        with tab3:
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**正文间距**")
                st.session_state.config["body_settings"]["space_before"] = st.number_input("段前间距", min_value=0, max_value=100, value=int(st.session_state.config["body_settings"]["space_before"]))
                st.session_state.config["body_settings"]["space_after"] = st.number_input("段后间距", min_value=0, max_value=100, value=int(st.session_state.config["body_settings"]["space_after"]))
            
            with col2:
                st.markdown("**表格间距**")
                st.session_state.config["table_settings"]["space_before"] = st.number_input("表格段前间距", min_value=0, max_value=100, value=int(st.session_state.config["table_settings"]["space_before"]))
                st.session_state.config["table_settings"]["space_after"] = st.number_input("表格段后间距", min_value=0, max_value=100, value=int(st.session_state.config["table_settings"]["space_after"]))
    
    # 处理按钮
    if uploaded_file:
        st.markdown("---")
        if st.button("🚀 开始处理文档", type="primary", use_container_width=True):
            with st.spinner("正在处理文档，请稍候..."):
                processed_doc = process_document(uploaded_file, st.session_state.config)
                
                if processed_doc:
                    st.session_state.processed = True
                    st.session_state.processed_data = processed_doc
                    st.session_state.output_filename = f"已处理_{uploaded_file.name}"
                    
                    st.markdown('<div class="success-box">✅ 文档处理完成！</div>', unsafe_allow_html=True)
    
    # 下载区域
    if st.session_state.processed:
        st.markdown('<div class="sub-header">📥 下载文档</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.download_button(
                label=f"📥 下载 {st.session_state.output_filename}",
                data=st.session_state.processed_data.getvalue(),
                file_name=st.session_state.output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        
        with col2:
            if st.button("🔄 重新处理", use_container_width=True):
                st.session_state.processed = False
                st.rerun()
        
        with col3:
            if st.button("⚡ 处理新文件", use_container_width=True):
                st.session_state.processed = False
                st.rerun()
    
    # 使用说明
    with st.expander("📖 使用说明", expanded=True):
        st.markdown("""
        ### ✨ 功能介绍
        
        **自动格式化 Word 文档：**
        1. 📝 **标题处理** - 自动转换大纲级别为标题样式
        2. 🔢 **智能编号** - 为标题添加规范的序号（可选）
        3. 🎨 **格式统一** - 统一正文和表格的字体、间距
        
        ### 🚀 使用步骤
        
        1. **上传文档** - 在左侧上传需要处理的 Word 文档
        2. **配置设置** - 根据需要调整基本设置
        3. **高级设置** - 如需更多控制，展开高级设置
        4. **开始处理** - 点击蓝色按钮开始处理
        5. **下载结果** - 处理完成后下载新文档
        
        ### ⚙️ 主要设置说明
        
        - **添加序号**：是否给标题添加自动编号
        - **最大级别**：设置标题的最大层级数
        - **字号/行距**：控制正文的基本格式
        - **缩进**：正文首行缩进距离
        
        ### 💡 小贴士
        
        - 高级设置中的"标题格式"可以自定义各级标题的编号样式
        - 可以同时调整中文字体和英文字体
        - 支持 9 级标题的自动编号
        """)

if __name__ == "__main__":
    main()
