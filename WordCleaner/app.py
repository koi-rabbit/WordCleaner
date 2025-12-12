import streamlit as st
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from io import BytesIO

# 设置页面配置
st.set_page_config(
    page_title="Word文档格式化工具",
    page_icon="📝",
    layout="wide"
)

# 默认配置
DEFAULT_CONFIG = {
    "标题": {
        "应用序号": True,
        "各级标题设置": {
            1: {"应用序号": True, "格式": "chinese"},
            2: {"应用序号": True, "格式": "chinese_bracket"},
            3: {"应用序号": True, "格式": "arabic_dot"},
            4: {"应用序号": True, "格式": "arabic_bracket"},
            5: {"应用序号": True, "格式": "arabic_dot"},
            6: {"应用序号": True, "格式": "arabic_bracket"},
            7: {"应用序号": True, "格式": "arabic_dot"},
            8: {"应用序号": True, "格式": "arabic_bracket"},
            9: {"应用序号": True, "格式": "arabic_dot"},
        }
    },
    "正文": {
        "中文字体": "宋体",
        "英文字体": "Times New Roman",
        "字号": 12,
        "段前间距": 12,
        "段后间距": 12,
        "行距": 1.5,
        "首行缩进": 0.5
    },
    "表格": {
        "中文字体": "宋体",
        "英文字体": "Times New Roman",
        "字号": 10,
        "段前间距": 6,
        "段后间距": 6,
        "表格宽度": 6
    }
}

# 初始化session state
if 'config' not in st.session_state:
    st.session_state.config = DEFAULT_CONFIG.copy()
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'current_tab' not in st.session_state:
    st.session_state.current_tab = "📝 标题"

# 样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 1rem;
        padding-top: 0.5rem;
    }
    .upload-box {
        border: 2px dashed #4F46E5;
        border-radius: 10px;
        padding: 2.5rem 2rem;
        text-align: center;
        background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%);
        margin: 1rem 0;
        transition: all 0.3s ease;
    }
    .upload-box:hover {
        border-color: #4338CA;
        background: linear-gradient(135deg, #667eea25 0%, #764ba225 100%);
    }
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px rgba(102, 126, 234, 0.2);
    }
    .success-box {
        background: linear-gradient(135deg, #10B981 0%, #059669 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        margin: 1rem 0;
        animation: fadeIn 0.5s ease-in;
    }
    .file-info {
        background: linear-gradient(135deg, #3B82F6 0%, #1D4ED8 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .config-section {
        background: white;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #E5E7EB;
        margin-bottom: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    .level-setting {
        background: #F9FAFB;
        padding: 1rem;
        border-radius: 8px;
        margin-bottom: 0.5rem;
        border-left: 4px solid #4F46E5;
    }
    .help-section {
        background: linear-gradient(135deg, #F3F4F6 0%, #E5E7EB 100%);
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        border-left: 4px solid #6B7280;
    }
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(-10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 1rem;
        padding: 0.5rem 0;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 0.75rem 1.5rem;
        border-radius: 8px;
        background: #F9FAFB;
        border: 1px solid #E5E7EB;
        transition: all 0.3s ease;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-color: #4F46E5;
        box-shadow: 0 2px 8px rgba(102, 126, 234, 0.2);
    }
    .stTabs [data-baseweb="tab"]:hover {
        background: #F3F4F6;
        transform: translateY(-2px);
    }
    .stTabs [aria-selected="true"]:hover {
        background: linear-gradient(135deg, #5a6fd8 0%, #6a4090 100%) !important;
    }
</style>
""", unsafe_allow_html=True)

def get_outline_level_from_xml(p):
    """从段落的XML中提取大纲级别"""
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
        return str(number)
    
    chinese_numbers = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    
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

def format_number(number, format_type):
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

def add_heading_numbers(doc, config):
    """根据配置添加标题序号"""
    if not config["标题"]["应用序号"]:
        return
    
    heading_numbers = [0] * 9  # 最多9级标题
    heading_settings = config["标题"]["各级标题设置"]
    
    # 匹配常见序号格式
    number_pattern = re.compile(r'^[\d一二三四五六七八九十（）\.、\s]+')

    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Heading'):
            try:
                level = int(paragraph.style.name.split(' ')[1]) - 1
                
                # 检查该级别是否应用序号
                if level + 1 not in heading_settings or not heading_settings[level + 1]["应用序号"]:
                    continue
                
                # 清理原有序号
                paragraph.text = number_pattern.sub('', paragraph.text).strip()
                
                # 更新序号
                heading_numbers[level] += 1
                for i in range(level + 1, 9):
                    heading_numbers[i] = 0
                
                # 获取格式并添加序号
                format_type = heading_settings[level + 1]["格式"]
                number_str = format_number(heading_numbers[level], format_type)
                paragraph.text = number_str + paragraph.text
            except Exception:
                continue

def modify_document_format(doc, config):
    """修改文档格式"""
    # 处理正文
    body_config = config["正文"]
    for paragraph in doc.paragraphs:
        if not paragraph.style.name.startswith("Heading"):
            paragraph.paragraph_format.space_before = Pt(body_config['段前间距'])
            paragraph.paragraph_format.space_after = Pt(body_config['段后间距'])
            paragraph.paragraph_format.line_spacing = body_config['行距']
            paragraph.paragraph_format.first_line_indent = Inches(body_config['首行缩进'])
            for run in paragraph.runs:
                set_font(run, body_config['中文字体'], body_config['英文字体'])
                run.font.size = Pt(body_config['字号'])

    # 处理表格
    table_config = config["表格"]
    for table_obj in doc.tables:
        table_obj.width = Inches(table_config['表格宽度'])
        for row in table_obj.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        set_font(run, table_config['中文字体'], table_config['英文字体'])
                        run.font.size = Pt(table_config['字号'])
                    paragraph.paragraph_format.space_before = Pt(table_config['段前间距'])
                    paragraph.paragraph_format.space_after = Pt(table_config['段后间距'])

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

def config_main():
    """主配置区域"""
    # 使用tabs组织三大类设置
    tab1, tab2, tab3 = st.tabs(["📝 标题设置", "📄 正文设置", "📊 表格设置"])
    
    with tab1:
        st.markdown('<div class="config-section">', unsafe_allow_html=True)
        
        # 全局标题设置
        col1, col2 = st.columns([1, 2])
        with col1:
            st.markdown("**启用标题自动编号**")
            st.session_state.config["标题"]["应用序号"] = st.toggle(
                "",
                value=st.session_state.config["标题"]["应用序号"],
                key="global_numbering_toggle",
                help="是否给标题添加自动序号"
            )
        
        if st.session_state.config["标题"]["应用序号"]:
            st.divider()
            st.markdown("**各级标题设置**")
            
            # 序号格式选项
            format_options = {
                "chinese": "中文数字（一、二、三）",
                "chinese_bracket": "中文数字加括号（（一）（二）（三））",
                "arabic_dot": "阿拉伯数字加点（1.2.3.）",
                "arabic_bracket": "阿拉伯数字加括号（（1）（2）（3））",
                "roman_lower": "小写罗马数字（i.ii.iii.）",
                "roman_upper": "大写罗马数字（I.II.III.）",
                "alphabet_lower": "小写字母（a.b.c.）",
                "alphabet_upper": "大写字母（A.B.C.）"
            }
            
            # 使用3列布局显示1-9级标题设置
            cols = st.columns(3)
            for level in range(1, 10):
                with cols[(level-1) % 3]:
                    st.markdown(f'<div class="level-setting">', unsafe_allow_html=True)
                    st.markdown(f"**第{level}级标题**")
                    
                    # 是否应用序号
                    apply = st.checkbox(
                        "应用序号",
                        value=st.session_state.config["标题"]["各级标题设置"][level]["应用序号"],
                        key=f"apply_{level}"
                    )
                    st.session_state.config["标题"]["各级标题设置"][level]["应用序号"] = apply
                    
                    if apply:
                        # 序号格式选择
                        current_format = st.session_state.config["标题"]["各级标题设置"][level]["格式"]
                        selected = st.selectbox(
                            "序号格式",
                            options=list(format_options.keys()),
                            format_func=lambda x: format_options[x],
                            index=list(format_options.keys()).index(current_format) if current_format in format_options else 2,
                            key=f"format_{level}"
                        )
                        st.session_state.config["标题"]["各级标题设置"][level]["格式"] = selected
                    
                    st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with tab2:
        st.markdown('<div class="config-section">', unsafe_allow_html=True)
        
        # 字体设置
        st.markdown("**字体设置**")
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.config["正文"]["中文字体"] = st.text_input(
                "中文字体",
                value=st.session_state.config["正文"]["中文字体"],
                key="body_cz_font"
            )
        with col2:
            st.session_state.config["正文"]["英文字体"] = st.text_input(
                "英文字体",
                value=st.session_state.config["正文"]["英文字体"],
                key="body_en_font"
            )
        
        st.divider()
        
        # 字号与行距
        st.markdown("**字号与行距**")
        col3, col4 = st.columns(2)
        with col3:
            st.session_state.config["正文"]["字号"] = st.number_input(
                "字号 (pt)",
                min_value=6,
                max_value=72,
                value=int(st.session_state.config["正文"]["字号"]),
                key="body_font_size"
            )
        with col4:
            st.session_state.config["正文"]["行距"] = st.number_input(
                "行距倍数",
                min_value=1.0,
                max_value=3.0,
                value=float(st.session_state.config["正文"]["行距"]),
                step=0.1,
                key="body_line_spacing"
            )
        
        st.divider()
        
        # 段落格式
        st.markdown("**段落格式**")
        col5, col6, col7 = st.columns(3)
        with col5:
            st.session_state.config["正文"]["段前间距"] = st.number_input(
                "段前间距 (pt)",
                min_value=0,
                max_value=100,
                value=int(st.session_state.config["正文"]["段前间距"]),
                key="body_before"
            )
        with col6:
            st.session_state.config["正文"]["段后间距"] = st.number_input(
                "段后间距 (pt)",
                min_value=0,
                max_value=100,
                value=int(st.session_state.config["正文"]["段后间距"]),
                key="body_after"
            )
        with col7:
            st.session_state.config["正文"]["首行缩进"] = st.number_input(
                "首行缩进 (英寸)",
                min_value=0.0,
                max_value=2.0,
                value=float(st.session_state.config["正文"]["首行缩进"]),
                step=0.1,
                key="body_indent"
            )
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with tab3:
        st.markdown('<div class="config-section">', unsafe_allow_html=True)
        
        # 字体设置
        st.markdown("**字体设置**")
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.config["表格"]["中文字体"] = st.text_input(
                "中文字体",
                value=st.session_state.config["表格"]["中文字体"],
                key="table_cz_font"
            )
        with col2:
            st.session_state.config["表格"]["英文字体"] = st.text_input(
                "英文字体",
                value=st.session_state.config["表格"]["英文字体"],
                key="table_en_font"
            )
        
        st.divider()
        
        # 字号与宽度
        st.markdown("**字号与宽度**")
        col3, col4 = st.columns(2)
        with col3:
            st.session_state.config["表格"]["字号"] = st.number_input(
                "表格字号 (pt)",
                min_value=6,
                max_value=72,
                value=int(st.session_state.config["表格"]["字号"]),
                key="table_font_size"
            )
        with col4:
            st.session_state.config["表格"]["表格宽度"] = st.number_input(
                "表格宽度 (英寸)",
                min_value=1,
                max_value=20,
                value=int(st.session_state.config["表格"]["表格宽度"]),
                key="table_width"
            )
        
        st.divider()
        
        # 间距设置
        st.markdown("**间距设置**")
        col5, col6 = st.columns(2)
        with col5:
            st.session_state.config["表格"]["段前间距"] = st.number_input(
                "表格段前间距 (pt)",
                min_value=0,
                max_value=100,
                value=int(st.session_state.config["表格"]["段前间距"]),
                key="table_before"
            )
        with col6:
            st.session_state.config["表格"]["段后间距"] = st.number_input(
                "表格段后间距 (pt)",
                min_value=0,
                max_value=100,
                value=int(st.session_state.config["表格"]["段后间距"]),
                key="table_after"
            )
        
        st.markdown('</div>', unsafe_allow_html=True)

def help_sidebar():
    """侧边栏帮助信息"""
    with st.sidebar:
        st.markdown("### 📖 使用说明")
        
        st.markdown('<div class="help-section">', unsafe_allow_html=True)
        st.markdown("**📤 上传文档**")
        st.markdown("""
        1. 点击上传区域选择.docx文件
        2. 支持批量处理（可逐个上传）
        3. 文件大小建议不超过50MB
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="help-section">', unsafe_allow_html=True)
        st.markdown("**⚙️ 配置说明**")
        st.markdown("""
        - **标题设置**：控制各级标题的自动编号
        - **正文设置**：调整文档正文的格式样式
        - **表格设置**：设置表格的字体和间距
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="help-section">', unsafe_allow_html=True)
        st.markdown("**🚀 处理流程**")
        st.markdown("""
        1. 上传文档后设置参数
        2. 点击"开始处理文档"
        3. 下载处理后的文件
        4. 可重新处理或处理新文件
        """)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.divider()
        
        st.markdown("### 🔧 功能特点")
        st.markdown("""
        - ✅ **自动编号**：支持9级标题自动编号
        - ✅ **多种格式**：中文、数字、字母、罗马数字
        - ✅ **灵活控制**：各级标题可单独设置
        - ✅ **批量处理**：可连续处理多个文档
        - ✅ **格式统一**：确保文档格式一致性
        """)
        
        st.divider()
        
        # 重置按钮
        if st.button("🔄 重置所有设置", use_container_width=True):
            st.session_state.config = DEFAULT_CONFIG.copy()
            st.success("设置已重置！")
            st.rerun()

def main():
    # 主标题
    st.markdown('<h1 class="main-header">📝 Word文档格式化工具</h1>', unsafe_allow_html=True)
    
    # 创建两列布局：左侧上传/处理，右侧帮助信息
    col1, col2 = st.columns([3, 1])
    
    with col1:
        # 上传区域
        st.markdown("### 📤 上传文档")
        uploaded_file = st.file_uploader(
            "",
            type=['docx'],
            help="选择需要格式化的Word文档",
            label_visibility="collapsed"
        )
        
        if uploaded_file:
            st.markdown(f'''
            <div class="file-info">
                <div style="font-size: 1.2rem; font-weight: 600; margin-bottom: 0.5rem;">
                    📄 {uploaded_file.name}
                </div>
                <div style="font-size: 0.9rem; opacity: 0.9;">
                    文件大小: {len(uploaded_file.getvalue()) / 1024:.1f} KB
                </div>
            </div>
            ''', unsafe_allow_html=True)
        
        # 参数设置区域
        st.markdown("### ⚙️ 参数设置")
        config_main()
        
        # 处理按钮区域
        st.markdown("---")
        if uploaded_file:
            if st.button("🚀 开始处理文档", type="primary", use_container_width=True, key="process_button"):
                with st.spinner("正在处理文档，请稍候..."):
                    processed_doc = process_document(uploaded_file, st.session_state.config)
                    
                    if processed_doc:
                        st.session_state.processed = True
                        st.session_state.processed_data = processed_doc
                        st.session_state.output_filename = f"已处理_{uploaded_file.name}"
                        st.rerun()
        
        # 结果展示区域
        if st.session_state.processed:
            st.markdown("### 📥 处理结果")
            st.markdown('<div class="success-box">✅ 文档处理完成！</div>', unsafe_allow_html=True)
            
            col_a, col_b = st.columns([2, 1])
            with col_a:
                st.download_button(
                    label=f"📥 下载 {st.session_state.output_filename}",
                    data=st.session_state.processed_data.getvalue(),
                    file_name=st.session_state.output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            with col_b:
                if st.button("🔄 重新处理", use_container_width=True):
                    st.session_state.processed = False
                    st.rerun()
    
    with col2:
        # 侧边栏帮助信息
        help_sidebar()

if __name__ == "__main__":
    main()
