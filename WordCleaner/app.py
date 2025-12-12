import streamlit as st
import re
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.shared import Inches
import tempfile
from io import BytesIO
import base64
import json
from typing import Dict, List

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
    "style_rules": {
        1: {'style_name': 'Heading 1', 'font_name': 'Arial','cz_font_name': '楷体', 'font_size': 10, 'bold': True, 'space_before': 12, 'space_after': 12, 'line_spacing': 1.5, 'first_line_indent': 18},
        2: {'style_name': 'Heading 2', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 14, 'bold': True, 'space_before': 10, 'space_after': 10, 'line_spacing': 1.5, 'first_line_indent': 18},
        3: {'style_name': 'Heading 3', 'font_name': 'Arial','cz_font_name': '宋体','font_size': 12, 'bold': False, 'space_before': 8, 'space_after': 8, 'line_spacing': 1.5, 'first_line_indent': 0},
        4: {'style_name': 'Heading 4', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 11, 'bold': False, 'space_before': 6, 'space_after': 6, 'line_spacing': 1.5, 'first_line_indent': 0},
        5: {'style_name': 'Heading 5', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 10, 'bold': False, 'space_before': 4, 'space_after': 4, 'line_spacing': 1.5, 'first_line_indent': 0},
        6: {'style_name': 'Heading 6', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 9, 'bold': False, 'space_before': 2, 'space_after': 2, 'line_spacing': 1.5, 'first_line_indent': 0},
        7: {'style_name': 'Heading 7', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 8, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 18},
        8: {'style_name': 'Heading 8', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 7, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 18},
        9: {'style_name': 'Heading 9', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 6, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 18},
    },
    "body_settings": {
        "cz_font_name": "宋体",
        "font_name": "Times New Roman",
        "font_size": 12,
        "space_before": 12,
        "space_after": 12,
        "line_spacing": 1.0,
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

def get_outline_level_from_xml(p):
    """
    从段落的XML中提取大纲级别，并加1
    """
    xml = p._p.xml
    m = re.search(r'<w:outlineLvl w:val="(\d)"/>', xml)
    level = int(m.group(1)) if m else None
    if level is not None:
        level += 1  # 加1
    return level

def set_font(run, cz_font_name, font_name):
    """
    设置字体。

    :param run: 文本运行对象
    :param chinese_font_name: 中文字体名称
    :param english_font_name: 英文字体名称
    """
    # 获取或创建字体属性
    rPr = run.element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    # 设置中文字体和英文字体
    rFonts.set(qn('w:eastAsia'), cz_font_name)
    rFonts.set(qn('w:ascii'), font_name)

# 手动实现数字到中文大写数字的转换
def number_to_chinese(number):
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
    """
    根据格式类型格式化数字
    """
    if format_type == "chinese":
        return f"{number_to_chinese(number)}、"
    elif format_type == "chinese_bracket":
        return f"（{number_to_chinese(number)}）"
    elif format_type == "arabic_dot":
        return f"{number}."
    elif format_type == "arabic_bracket":
        return f"（{number}）"
    elif format_type == "roman_lower":
        roman_map = [(1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'),
                     (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'),
                     (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')]
        result = ""
        num = number
        for value, numeral in roman_map:
            while num >= value:
                result += numeral
                num -= value
        return f"{result}."
    elif format_type == "roman_upper":
        roman_map = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'),
                     (100, 'C'), (90, 'XC'), (50, 'L'), (40, 'XL'),
                     (10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
        result = ""
        num = number
        for value, numeral in roman_map:
            while num >= value:
                result += numeral
                num -= value
        return f"{result}."
    elif format_type == "alphabet_lower":
        if number <= 26:
            return f"{chr(96 + number)}."
        else:
            return f"{number}."
    elif format_type == "alphabet_upper":
        if number <= 26:
            return f"{chr(64 + number)}."
        else:
            return f"{number}."
    else:
        return f"{number}."

# 添加标题序号并清洗原有序号
def add_heading_numbers(doc, config):
    """
    根据配置添加标题序号
    """
    if not config["title_settings"]["apply_numbering"]:
        return
    
    # 初始化标题序号
    max_levels = config["title_settings"]["max_levels"]
    heading_numbers = [0] * max_levels
    
    # 获取序号格式配置
    numbering_formats = config["title_settings"]["numbering_formats"]
    
    # 定义正则表达式，匹配常见的序号格式
    number_pattern = re.compile(r'^[\d一二三四五六七八九十（）\.、\s]+')

    # 遍历文档中的所有段落
    for paragraph in doc.paragraphs:
        # 检查段落是否是标题
        if paragraph.style.name.startswith('Heading'):
            # 获取标题级别
            try:
                level = int(paragraph.style.name.split(' ')[1]) - 1
                if level >= max_levels:
                    continue
                    
                # 清洗原文档中的序号
                paragraph.text = number_pattern.sub('', paragraph.text).strip()

                # 更新序号
                heading_numbers[level] += 1
                for i in range(level + 1, len(heading_numbers)):
                    heading_numbers[i] = 0  # 重置下级标题序号

                # 获取该级别的格式类型
                format_type = numbering_formats.get(level + 1, "arabic_dot")
                
                # 构造序号字符串
                number_str = format_number(level, heading_numbers[level], format_type)

                # 添加序号到标题文本
                paragraph.text = number_str + paragraph.text
            except (ValueError, IndexError):
                continue

def modify_document_format(doc, config):
    """
    修改 Word 文档中正文和表格的格式。
    """
    style_rules = config["style_rules"]
    body_settings = config["body_settings"]
    table_settings = config["table_settings"]
    
    # 遍历文档中的每个段落
    for paragraph in doc.paragraphs:
        # 检查是否是标题（标题的 style 通常以 "Heading" 开头）
        if paragraph.style.name.startswith("Heading"):
            style_name = paragraph.style.name
            # 查找匹配的样式规则
            for level, rule in style_rules.items():
                if rule['style_name'] == style_name:
                    # 修改段前段后行距和首行缩进
                    paragraph.style.paragraph_format.space_before = Pt(rule['space_before'])
                    paragraph.style.paragraph_format.space_after = Pt(rule['space_after'])
                    paragraph.style.paragraph_format.line_spacing = rule['line_spacing']
                    paragraph.style.paragraph_format.first_line_indent = Pt(rule['first_line_indent'])
                    # 修改字体字号和粗体
                    for run in paragraph.runs:
                        set_font(run, rule['cz_font_name'], rule['font_name'])
                        run.font.size = Pt(rule['font_size'])
                        run.font.bold = rule['bold']
                    break
        else:            
            # 修改段前段后行距和首行缩进
            paragraph.paragraph_format.space_before = Pt(body_settings['space_before'])
            paragraph.paragraph_format.space_after = Pt(body_settings['space_after'])
            paragraph.paragraph_format.line_spacing = body_settings['line_spacing']
            paragraph.paragraph_format.first_line_indent = Inches(body_settings['first_line_indent'])
            # 修改字体字号
            for run in paragraph.runs:
                set_font(run, body_settings['cz_font_name'], body_settings['font_name'])
                run.font.size = Pt(body_settings['font_size'])

    # 遍历文档中的每个表格
    for table in doc.tables:
        table.width = Inches(table_settings['width'])
        # 遍历表格中的每个单元格
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    # 修改字体和字号
                    for run in paragraph.runs:
                        # 设置中文字体和英文字体
                        set_font(run, table_settings['cz_font_name'], table_settings['font_name'])
                        # 设置字号
                        run.font.size = Pt(table_settings['font_size'])

                    # 修改段前段后行距
                    paragraph.paragraph_format.space_before = Pt(table_settings['space_before'])
                    paragraph.paragraph_format.space_after = Pt(table_settings['space_after'])

def process_document(uploaded_file, config):
    """
    处理上传的Word文档
    """
    try:
        # 读取上传的文件
        doc = Document(uploaded_file)
        
        # 应用大纲级别转换
        for para in doc.paragraphs:
            outline_level = get_outline_level_from_xml(para)
            style_name = para.style.name

            # 如果获取到大纲级别且当前样式为正文，根据大纲级别设置对应的标题样式
            if outline_level is not None and style_name == 'Normal':
                # 根据大纲级别设置标题样式
                if outline_level <= 9:  # 只处理1-9级标题
                    heading_style = f"Heading {outline_level}"
                    if heading_style in doc.styles:
                        para.style = doc.styles[heading_style]
        
        # 添加标题序号并清洗原有序号
        add_heading_numbers(doc, config)
        
        # 应用样式规则
        modify_document_format(doc, config)
        
        # 将处理后的文档保存到字节流
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        
        return output
    except Exception as e:
        st.error(f"处理文档时出错: {str(e)}")
        return None

def config_sidebar():
    """
    配置侧边栏
    """
    st.sidebar.title("⚙️ 配置设置")
    
    # 标题设置
    st.sidebar.subheader("📝 标题设置")
    
    # 是否应用序号
    st.session_state.config["title_settings"]["apply_numbering"] = st.sidebar.checkbox(
        "应用标题序号", 
        value=st.session_state.config["title_settings"]["apply_numbering"]
    )
    
    # 最大标题级别
    max_levels = st.sidebar.slider(
        "最大标题级别", 
        min_value=1, 
        max_value=9, 
        value=st.session_state.config["title_settings"]["max_levels"]
    )
    st.session_state.config["title_settings"]["max_levels"] = max_levels
    
    # 序号格式配置
    st.sidebar.markdown("**序号格式配置**")
    
    # 定义序号格式选项
    numbering_options = {
        "chinese": "中文数字（一、二、三）",
        "chinese_bracket": "中文数字加括号（（一）（二）（三））",
        "arabic_dot": "阿拉伯数字加点（1.2.3.）",
        "arabic_bracket": "阿拉伯数字加括号（（1）（2）（3））",
        "roman_lower": "小写罗马数字（i.ii.iii.）",
        "roman_upper": "大写罗马数字（I.II.III.）",
        "alphabet_lower": "小写字母（a.b.c.）",
        "alphabet_upper": "大写字母（A.B.C.）"
    }
    
    # 为每个级别配置序号格式
    for level in range(1, max_levels + 1):
        current_format = st.session_state.config["title_settings"]["numbering_formats"].get(level, "arabic_dot")
        
        # 获取格式名称
        format_name = numbering_options.get(current_format, "阿拉伯数字加点")
        
        # 创建选择框
        selected_format = st.sidebar.selectbox(
            f"第{level}级标题格式",
            options=list(numbering_options.keys()),
            format_func=lambda x: numbering_options[x],
            index=list(numbering_options.keys()).index(current_format) if current_format in numbering_options else 0,
            key=f"heading_format_{level}"
        )
        
        st.session_state.config["title_settings"]["numbering_formats"][level] = selected_format
    
    # 正文设置
    st.sidebar.subheader("📄 正文设置")
    
    col1, col2 = st.sidebar.columns(2)
    
    with col1:
        st.session_state.config["body_settings"]["cz_font_name"] = st.text_input(
            "中文字体", 
            value=st.session_state.config["body_settings"]["cz_font_name"],
            key="body_cz_font"
        )
    
    with col2:
        st.session_state.config["body_settings"]["font_name"] = st.text_input(
            "英文字体", 
            value=st.session_state.config["body_settings"]["font_name"],
            key="body_en_font"
        )
    
    st.session_state.config["body_settings"]["font_size"] = st.sidebar.number_input(
        "字号 (pt)", 
        min_value=6.0, 
        max_value=72.0, 
        value=float(st.session_state.config["body_settings"]["font_size"]),
        step=0.5,
        key="body_font_size"
    )
    
    col3, col4 = st.sidebar.columns(2)
    with col3:
        st.session_state.config["body_settings"]["space_before"] = st.number_input(
            "段前间距 (pt)", 
            min_value=0.0, 
            max_value=100.0, 
            value=float(st.session_state.config["body_settings"]["space_before"]),
            step=0.5,
            key="body_space_before"
        )
    
    with col4:
        st.session_state.config["body_settings"]["space_after"] = st.number_input(
            "段后间距 (pt)", 
            min_value=0.0, 
            max_value=100.0, 
            value=float(st.session_state.config["body_settings"]["space_after"]),
            step=0.5,
            key="body_space_after"
        )
    
    st.session_state.config["body_settings"]["line_spacing"] = st.sidebar.number_input(
        "行距倍数", 
        min_value=1.0, 
        max_value=3.0, 
        value=float(st.session_state.config["body_settings"]["line_spacing"]),
        step=0.1,
        key="body_line_spacing"
    )
    
    st.session_state.config["body_settings"]["first_line_indent"] = st.sidebar.number_input(
        "首行缩进 (英寸)", 
        min_value=0.0, 
        max_value=2.0, 
        value=float(st.session_state.config["body_settings"]["first_line_indent"]),
        step=0.1,
        key="body_indent"
    )
    
    # 表格设置
    st.sidebar.subheader("📊 表格设置")
    
    col5, col6 = st.sidebar.columns(2)
    
    with col5:
        st.session_state.config["table_settings"]["cz_font_name"] = st.text_input(
            "表格中文字体", 
            value=st.session_state.config["table_settings"]["cz_font_name"],
            key="table_cz_font"
        )
    
    with col6:
        st.session_state.config["table_settings"]["font_name"] = st.text_input(
            "表格英文字体", 
            value=st.session_state.config["table_settings"]["font_name"],
            key="table_en_font"
        )
    
    st.session_state.config["table_settings"]["font_size"] = st.sidebar.number_input(
        "表格字号 (pt)", 
        min_value=6.0, 
        max_value=72.0, 
        value=float(st.session_state.config["table_settings"]["font_size"]),
        step=0.5,
        key="table_font_size"
    )
    
    col7, col8 = st.sidebar.columns(2)
    with col7:
        st.session_state.config["table_settings"]["space_before"] = st.number_input(
            "表格段前间距 (pt)", 
            min_value=0.0, 
            max_value=100.0, 
            value=float(st.session_state.config["table_settings"]["space_before"]),
            step=0.5,
            key="table_space_before"
        )
    
    with col8:
        st.session_state.config["table_settings"]["space_after"] = st.number_input(
            "表格段后间距 (pt)", 
            min_value=0.0, 
            max_value=100.0, 
            value=float(st.session_state.config["table_settings"]["space_after"]),
            step=0.5,
            key="table_space_after"
        )
    
    st.session_state.config["table_settings"]["width"] = st.sidebar.number_input(
        "表格宽度 (英寸)", 
        min_value=1.0, 
        max_value=20.0, 
        value=float(st.session_state.config["table_settings"]["width"]),
        step=0.5,
        key="table_width"
    )
    
    # 保存和重置按钮
    st.sidebar.markdown("---")
    col9, col10 = st.sidebar.columns(2)
    
    with col9:
        if st.button("💾 保存配置", use_container_width=True):
            # 将配置保存到文件
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(st.session_state.config, f, ensure_ascii=False, indent=2)
            st.sidebar.success("配置已保存！")
    
    with col10:
        if st.button("🔄 重置配置", use_container_width=True):
            st.session_state.config = DEFAULT_CONFIG.copy()
            st.sidebar.success("配置已重置！")
            st.rerun()
    
    # 加载配置
    if os.path.exists("config.json"):
        if st.sidebar.button("📂 加载配置", use_container_width=True):
            try:
                with open("config.json", "r", encoding="utf-8") as f:
                    st.session_state.config = json.load(f)
                st.sidebar.success("配置已加载！")
                st.rerun()
            except Exception as e:
                st.sidebar.error(f"加载配置失败: {str(e)}")

def main():
    st.title("📝 Word文档格式化工具")
    
    # 配置侧边栏
    config_sidebar()
    
    # 主内容区域
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.subheader("📤 上传和处理")
        
        # 文件上传区域
        uploaded_file = st.file_uploader(
            "选择.docx文件",
            type=['docx'],
            help="请上传需要格式化的Word文档",
            key="file_uploader"
        )
        
        if uploaded_file:
            st.success(f"已上传: {uploaded_file.name}")
            
            # 显示文件信息
            file_size = len(uploaded_file.getvalue()) / 1024  # KB
            st.info(f"文件大小: {file_size:.2f} KB")
            
            # 当前配置预览
            with st.expander("📋 查看当前配置", expanded=False):
                config_col1, config_col2 = st.columns(2)
                
                with config_col1:
                    st.markdown("**标题设置**")
                    st.json(st.session_state.config["title_settings"])
                
                with config_col2:
                    st.markdown("**正文和表格设置**")
                    st.json({
                        "body": st.session_state.config["body_settings"],
                        "table": st.session_state.config["table_settings"]
                    })
            
            # 处理按钮
            if st.button("🚀 开始处理文档", type="primary", use_container_width=True):
                with st.spinner("正在处理文档，请稍候..."):
                    # 处理文档
                    processed_doc = process_document(uploaded_file, st.session_state.config)
                    
                    if processed_doc:
                        st.success("✅ 文档处理完成！")
                        
                        # 显示下载链接
                        st.subheader("📥 下载处理后的文档")
                        output_filename = f"已处理_{uploaded_file.name}"
                        
                        # 创建下载按钮
                        st.download_button(
                            label=f"下载 {output_filename}",
                            data=processed_doc.getvalue(),
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            type="primary"
                        )
    
    with col2:
        st.subheader("📊 配置预览")
        
        # 显示配置概览
        config = st.session_state.config
        
        st.markdown("**标题设置**")
        if config["title_settings"]["apply_numbering"]:
            st.markdown("✅ 启用序号")
            st.markdown(f"最大级别: {config['title_settings']['max_levels']}")
        else:
            st.markdown("❌ 禁用序号")
        
        st.markdown("---")
        
        st.markdown("**正文设置**")
        st.markdown(f"""
        - 字体: {config['body_settings']['cz_font_name']} / {config['body_settings']['font_name']}
        - 字号: {config['body_settings']['font_size']}pt
        - 行距: {config['body_settings']['line_spacing']}倍
        - 缩进: {config['body_settings']['first_line_indent']}英寸
        """)
        
        st.markdown("---")
        
        st.markdown("**表格设置**")
        st.markdown(f"""
        - 字体: {config['table_settings']['cz_font_name']} / {config['table_settings']['font_name']}
        - 字号: {config['table_settings']['font_size']}pt
        - 宽度: {config['table_settings']['width']}英寸
        """)
        
        st.markdown("---")
        
        # 帮助信息
        st.markdown("### 💡 使用说明")
        st.markdown("""
        1. 在左侧配置所有设置
        2. 上传Word文档
        3. 点击"开始处理文档"
        4. 下载处理后的文档
        
        支持标题级别自动转换和多种序号格式。
        """)

if __name__ == "__main__":
    main()
