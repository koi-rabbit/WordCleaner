# app.py
import streamlit as st
from docx import Document
import re, os
from io import BytesIO
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
# ……后面直接粘你原来的常量、函数……

# 标题样式
style_rules = {
    1: {'style_name': 'Heading 1', 'font_name': 'Arial','cz_font_name': '楷体', 'font_size': 10, 'bold': True, 'space_before': 12, 'space_after': 12, 'line_spacing': 1.5, 'first_line_indent': 18},
    2: {'style_name': 'Heading 2', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 14, 'bold': True, 'space_before': 10, 'space_after': 10, 'line_spacing': 1.5, 'first_line_indent': 18},
    3: {'style_name': 'Heading 3', 'font_name': 'Arial','cz_font_name': '宋体','font_size': 12, 'bold': False, 'space_before': 8, 'space_after': 8, 'line_spacing': 1.5, 'first_line_indent': 0},
    4: {'style_name': 'Heading 4', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 11, 'bold': False, 'space_before': 6, 'space_after': 6, 'line_spacing': 1.5, 'first_line_indent': 0},
    5: {'style_name': 'Heading 5', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 10, 'bold': False, 'space_before': 4, 'space_after': 4, 'line_spacing': 1.5, 'first_line_indent': 0},
    6: {'style_name': 'Heading 6', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 9, 'bold': False, 'space_before': 2, 'space_after': 2, 'line_spacing': 1.5, 'first_line_indent': 0},
    7: {'style_name': 'Heading 7', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 8, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 18},
    8: {'style_name': 'Heading 8', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 7, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 18},
    9: {'style_name': 'Heading 9', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 6, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 18},

}

# 正文格式
bdy_cz_font_name = "宋体"  # 字体
bdy_font_name = "Times New Roman"
bdy_font_size = Pt(12)  # 字号
bdy_space_before = Pt(12)  # 段前行距
bdy_space_after = Pt(12)  # 段后行距
bdy_line_spacing = 1.0  #行距
bdy_first_line_indent = Inches(0.5)  # 首行缩进

# 表格格式
tbl_cz_font_name = "宋体"  # 中文字体
tbl_font_name = "Times New Roman"  # 英文字体
tbl_font_size = Pt(10)  # 表格字号
tbl_space_before = Pt(6)  # 表格段前行距
tbl_space_after = Pt(6)  # 表格段后行距
tbl_width = Inches(6)

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
   
# 添加标题序号并清洗原有序号
def add_heading_numbers(doc):
    # 初始化标题序号
    heading_numbers = [0, 0, 0, 0, 0, 0, 0, 0, 0]  # 假设最多有九级标题

    # 定义不同层级的序号格式
    def format_number(level, number):
        if level == 0:
            return f"{number_to_chinese(number)}、"  # 第一层级：一、二、三、
        elif level == 1:
            return f"（{number_to_chinese(number)}）"  # 第二层级：（一）（二）（三）
        elif level == 2:
            return f"{number}."  # 第三层级：1.2.3.
        elif level == 3:
            return f"（{number}）"  # 第四层级：（1）（2）（3）
        elif level == 4:
            return f"{number}."  # 第五层级：1.2.3.
        elif level == 5:
            return f"（{number}）"  # 第六层级：（1）（2）（3）
        elif level == 6:
            return f"{number}."  # 第七层级：1.2.3.
        elif level == 7:
            return f"（{number}）"  # 第八层级：（1）（2）（3）
        elif level == 8:
            return f"{number}."  # 第九层级：1.2.3.
        else:
            return f"{number}."  # 默认格式

    # 定义正则表达式，匹配常见的序号格式
    number_pattern = re.compile(r'^[\d一二三四五六七八九十（）\.、\s]+')

    # 遍历文档中的所有段落
    for paragraph in doc.paragraphs:
        # 检查段落是否是标题
        if paragraph.style.name.startswith('Heading'):
            # 获取标题级别
            level = int(paragraph.style.name.split(' ')[1]) - 1

            # 清洗原文档中的序号
            paragraph.text = number_pattern.sub('', paragraph.text).strip()

            # 更新序号
            heading_numbers[level] += 1
            for i in range(level + 1, len(heading_numbers)):
                heading_numbers[i] = 0  # 重置下级标题序号

            # 构造序号字符串
            number_str = format_number(level, heading_numbers[level])

            # 添加序号到标题文本
            paragraph.text = number_str + paragraph.text

def modify_document_format(doc):
    """
    修改 Word 文档中正文和表格的格式。

    :param file_path: 输入的 Word 文档路径
    :param output_path: 输出的 Word 文档路径，默认为 "modified.docx"
    """    
    # 遍历文档中的每个段落
    for paragraph in doc.paragraphs:
        # 检查是否是标题（标题的 style 通常以 "Heading" 开头）
        if  paragraph.style.name.startswith("Heading"):
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
        else:            
            # 修改段前段后行距和首行缩进
            paragraph.paragraph_format.space_before = bdy_space_before
            paragraph.paragraph_format.space_after = bdy_space_after
            paragraph.paragraph_format.line_spacing = bdy_line_spacing
            paragraph.paragraph_format.first_line_indent = bdy_first_line_indent
            # 修改字体字号
            for run in paragraph.runs:
                set_font(run, bdy_cz_font_name, bdy_font_name)
                run.font.size = bdy_font_size

                
    # 遍历文档中的每个表格
    for table in doc.tables:
        table.width = tbl_width 
        # 遍历表格中的每个单元格
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    # 修改字体和字号
                    for run in paragraph.runs:
                        # 设置中文字体和英文字体
                        set_font(run, tbl_cz_font_name, tbl_font_name)
                        # 设置字号
                        run.font.size = tbl_font_size

                    # 修改段前段后行距
                    paragraph.paragraph_format.space_before = tbl_space_before
                    paragraph.paragraph_format.space_after = tbl_space_after


def process_doc(uploaded_bytes):
    doc = Document(BytesIO(uploaded_bytes))
    # 下面就是你原来的 main 逻辑里“处理”部分
    for para in doc.paragraphs:
        lvl = get_outline_level_from_xml(para)
        if lvl and para.style.name == "Normal":
            para.style = doc.styles[f"Heading {lvl}"]
    add_heading_numbers(doc)
    modify_document_format(doc)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------- Streamlit 界面 ----------------
if f and st.button("开始排版"):
    with st.spinner("处理中…"):
        out = process_doc(f.read())
    base = f.name.replace(".docx", "")
    st.download_button(
        label="📥 下载已排版文件",
        data=out,
        file_name=f"{base}_已排版.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

