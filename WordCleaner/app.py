# app.py
import streamlit as st
from docx import Document
import re, os
from io import BytesIO
from docx.shared import Pt, Inches
from docx.oxml.ns import qn
from docx.shared import Cm

# -------------- 默认值 --------------
DEFAULTS = {
    # 正文
    "bdy_cz_font_name": "宋体",
    "bdy_font_name": "Times New Roman",
    "bdy_font_size": 10.5,
    "bdy_space_before": 6.0,
    "bdy_space_after": 6.0,
    "bdy_line_spacing": 1.0,
    "bdy_first_line_indent": 0.75,
    # 表格
    "tbl_cz_font_name": "宋体",
    "tbl_font_name": "Times New Roman",
    "tbl_font_size": 10.5,
    "tbl_space_before": 4.0,
    "tbl_space_after": 4.0,
    "tbl_line_spacing": 1.0,
    "tbl_width": 6.0,
}
# -------------- 初始化 / 重置 --------------
def init_state():
    for k, v in DEFAULTS.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# -------------- 侧边栏：参数面板 --------------
with st.sidebar:
    st.title("📏 格式参数")
    st.markdown("---")
    with st.expander("正文格式", expanded=True):
        st.session_state["bdy_cz_font_name"] = st.text_input("中文字体", st.session_state["bdy_cz_font_name"])
        st.session_state["bdy_font_name"] = st.text_input("英文字体", st.session_state["bdy_font_name"])
        st.session_state["bdy_font_size"] = st.number_input("字号(pt)", 5.0, 30.0, st.session_state["bdy_font_size"], 0.5)
        st.session_state["bdy_space_before"] = st.number_input("段前行距(pt)", 0.0, 50.0, st.session_state["bdy_space_before"])
        st.session_state["bdy_space_after"] = st.number_input("段后行距(pt)", 0.0, 50.0, st.session_state["bdy_space_after"])
        st.session_state["bdy_line_spacing"] = st.number_input("行距(倍)", 0.5, 3.0, st.session_state["bdy_line_spacing"], 0.1)
        st.session_state["bdy_first_line_indent"] = st.number_input("首行缩进(cm)", 0.0, 5.0, st.session_state["bdy_first_line_indent"], 0.05)

    with st.expander("表格格式", expanded=True):
        st.session_state["tbl_cz_font_name"] = st.text_input("表格中文字体", st.session_state["tbl_cz_font_name"])
        st.session_state["tbl_font_name"] = st.text_input("表格英文字体", st.session_state["tbl_font_name"])
        st.session_state["tbl_font_size"] = st.number_input("表格字号(pt)", 5.0, 30.0, st.session_state["tbl_font_size"], 0.5)
        st.session_state["tbl_space_before"] = st.number_input("表格段前行距(pt)", 0.0, 50.0, st.session_state["tbl_space_before"])
        st.session_state["tbl_space_after"] = st.number_input("表格段后行距(pt)", 0.0, 50.0, st.session_state["tbl_space_after"])
        st.session_state["tbl_line_spacing"] = st.number_input("表格行距(倍)", 0.5, 3.0, st.session_state["tbl_line_spacing"], 0.1)
        st.session_state["tbl_width"] = st.number_input("表格宽度(inches)", 1.0, 10.0, st.session_state["tbl_width"], 0.1)

    if st.button("重置全部参数"):
        for k, v in DEFAULTS.items():
            st.session_state[k] = v
        st.rerun()

# 标题样式
style_rules = {
    1: {'style_name': 'Heading 1', 'font_name': 'Arial','cz_font_name': '黑体', 'font_size': 14, 'bold': True, 'space_before': 12, 'space_after': 12, 'line_spacing': 1.5, 'first_line_indent': 0},
    2: {'style_name': 'Heading 2', 'font_name': 'Arial','cz_font_name': '黑体', 'font_size': 12, 'bold': True, 'space_before': 12, 'space_after': 12, 'line_spacing': 1.5, 'first_line_indent': 0.75},
    3: {'style_name': 'Heading 3', 'font_name': 'Times New Roman','cz_font_name': '宋体','font_size': 10.5, 'bold': False, 'space_before': 8, 'space_after': 8, 'line_spacing': 1.0, 'first_line_indent': 1.5},
    4: {'style_name': 'Heading 4', 'font_name': 'Times New Roman','cz_font_name': '宋体', 'font_size': 10.5, 'bold': False, 'space_before': 8, 'space_after': 8, 'line_spacing': 1.0, 'first_line_indent': 2.25},
    5: {'style_name': 'Heading 5', 'font_name': 'Times New Roman','cz_font_name': '宋体', 'font_size': 10.5, 'bold': False, 'space_before': 6, 'space_after': 6, 'line_spacing': 1.0, 'first_line_indent': 3.0},
    6: {'style_name': 'Heading 6', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 9, 'bold': False, 'space_before': 2, 'space_after': 2, 'line_spacing': 1.0, 'first_line_indent': 0},
    7: {'style_name': 'Heading 7', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 8, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 0},
    8: {'style_name': 'Heading 8', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 7, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 0},
    9: {'style_name': 'Heading 9', 'font_name': 'Arial','cz_font_name': '宋体', 'font_size': 6, 'bold': False, 'space_before': 0, 'space_after': 0, 'line_spacing': 1.0, 'first_line_indent': 0},

}

# 正文格式
bdy_cz_font_name = st.session_state["bdy_cz_font_name"]  # 字体
bdy_font_name = st.session_state["bdy_font_name"]
bdy_font_size = Pt(st.session_state["bdy_font_size"])  # 字号
bdy_space_before = Pt(st.session_state["bdy_space_before"])  # 段前行距
bdy_space_after = Pt(st.session_state["bdy_space_after"])  # 段后行距
bdy_line_spacing = st.session_state["bdy_line_spacing"]  #行距
bdy_first_line_indent = Cm(st.session_state["bdy_first_line_indent"])  # 首行缩进

# 表格格式
tbl_cz_font_name = st.session_state["tbl_cz_font_name"]  # 中文字体
tbl_font_name = st.session_state["tbl_font_name"]  # 英文字体
tbl_font_size = Pt(st.session_state["tbl_font_size"])  # 表格字号
tbl_space_before = Pt(st.session_state["tbl_space_before"])  # 表格段前行距
tbl_space_after = Pt(st.session_state["tbl_space_after"])  # 表格段后行距
tbl_line_spacing = st.session_state["tbl_line_spacing"]  #行距
tbl_width = Inches(st.session_state["tbl_width"])

KNOWN_STYLES = {
    "Normal",
    "List Paragraph",
    "Heading 1", "Heading 2", "Heading 3", "Heading 4",
    "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9"
}

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

def restructure_outline(doc):
    # ---------- 1. 升级：XML 大纲 → Heading ----------
    for p in doc.paragraphs:
        zero_indent(p)
        lvl = get_outline_level_from_xml(p)
        if lvl and p.style.name == "Normal":
            # Heading 1~9 才存在
            heading_style = f"Heading {lvl}"
            if heading_style in doc.styles:
                p.style = doc.styles[heading_style]

    # ---------- 2. 降级：空标题 ----------
    headings_idx: List[int] = []
    for idx, p in enumerate(doc.paragraphs):
        if p.style.name.startswith("Heading"):
            headings_idx.append(idx)
            if not p.text.strip():          # 空
                p.style = doc.styles["Normal"]
           
def zero_indent(p):
    pf = p.paragraph_format
    pf.left_indent       = Cm(0)
    pf.first_line_indent = Cm(0)
    pf.right_indent      = Cm(0)
    pf.tab_stops.clear_all()   # 清制表位
    # 再删段首空格/Tab
    if p.text:
        p.text = p.text.lstrip()

def kill_all_numbering(doc):
    """样式级 + 段落级 编号全部清零"""
    # 1. 样式级：把所有带 numId 的样式拔掉
    for st_name in ['List Paragraph', 'Heading 1', 'Heading 2', 'Heading 3',
                    'Heading 4', 'Heading 5', 'Heading 6', 'Heading 7',
                    'Heading 8', 'Heading 9']:
        try:
            style = doc.styles[st_name]
        except KeyError:
            continue
        style_el = style._element
        for num_id in style_el.xpath('.//w:numId'):
            num_id.getparent().remove(num_id) 
               
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
           
# 添加标题序号
def add_heading_numbers(doc):
    
    number_pattern = re.compile(
        r'^\s*'                                      # 前导空格
        r'[（(]?'                                     # 可选左括号（全角/半角）
        r'[\d一二三四五六七八九十零]{1,3}'            # 首位数字（阿拉伯或中文）
        r'[\.、）)]'                                  # 必须跟一个点号/顿号/右括号（把“锚”做实）
        r'(?:[（(]?\s*[\d一二三四五六七八九十零]{1,3}[\.、）)]\s*)*'  # 后面可再出现“数字+点/顿/括号”
        r'\s*$',                                     # 尾部空格
        re.UNICODE
    )
    
    # 初始化标题序号
    heading_numbers = [0, 0, 0, 0, 0, 0, 0, 0, 0]  # 假设最多有九级标题
    
    def circled_num(n: int) -> str:
        if 1 <= n <= 20:                       # 目前 Unicode 只到 ⑳
            return chr(0x245F + n)             # 0x2460 - 1 + n
        return str(n)                          # 超出 fallback
        
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
            return f"{circled_num(number)} "  # 第五层级：圈1 圈2 圈3
        elif level == 5:
            return f"{circled_num(number)} "  # 第六层级：圈1 圈2 圈3
        elif level == 6:
            return f"{circled_num(number)} "  # 第七层级：圈1 圈2 圈3
        elif level == 7:
            return f"{circled_num(number)} "  # 第八层级：圈1 圈2 圈3
        elif level == 8:
            return f"{circled_num(number)} "  # 第九层级：圈1 圈2 圈3
        else:
            return f"{number}."  # 默认格式

    # 遍历文档中的所有段落
    for paragraph in doc.paragraphs:
        # 检查段落是否是标题
        if paragraph.style.name.startswith('Heading'):
            #清洗手写序号
            if paragraph.text == "Ellipsis" or not paragraph.text.strip():
                continue            
            for p in doc.paragraphs:
                p.text = number_pattern.sub('', p.text).lstrip()
                p_pr = p._p.get_or_add_pPr()
                num_pr = p_pr.find(qn('w:numPr'))
                if num_pr is not None:
                    p_pr.remove(num_pr)
            paragraph.text = number_pattern.sub('', paragraph.text).strip()
            # 获取标题级别
            level = int(paragraph.style.name.split(' ')[1]) - 1

            # 更新序号
            heading_numbers[level] += 1
            for i in range(level + 1, len(heading_numbers)):
                heading_numbers[i] = 0  # 重置下级标题序号

            # 构造序号字符串
            number_str = format_number(level, heading_numbers[level])

            # 添加序号到标题文本
            paragraph.text = number_str + paragraph.text

def modify_document_format(doc):
    skipped = set()

    for p in doc.paragraphs:
        style_name = p.style.name

        # 跳过空段落
        if p.text == "Ellipsis" or not p.text.strip():
            continue

        if style_name not in KNOWN_STYLES:
            skipped.add(style_name)
            continue

        if style_name.startswith("Heading"):
            for rule_lvl, rule in style_rules.items():
                if rule['style_name'] == style_name:
                    p.style.paragraph_format.space_before = Pt(rule['space_before'])
                    p.style.paragraph_format.space_after = Pt(rule['space_after'])
                    p.style.paragraph_format.line_spacing = rule['line_spacing']
                    p.style.paragraph_format.first_line_indent = Cm(rule['first_line_indent'])
                    for run in p.runs:
                        set_font(run, rule['cz_font_name'], rule['font_name'])
                        run.font.size = Pt(rule['font_size'])
                        run.font.bold = rule['bold']
        else:
            p.paragraph_format.space_before = bdy_space_before
            p.paragraph_format.space_after = bdy_space_after
            p.paragraph_format.line_spacing = bdy_line_spacing
            p.paragraph_format.first_line_indent = bdy_first_line_indent
            for run in p.runs:
                set_font(run, bdy_cz_font_name, bdy_font_name)
                run.font.size = bdy_font_size

    # 表格处理不变
    for tbl in doc.tables:
        tbl.width = tbl_width
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.style.name != "Normal":
                        skipped.add(f"表格内：{p.style.name}")
                        continue
                    for run in p.runs:
                        set_font(run, tbl_cz_font_name, tbl_font_name)
                        run.font.size = tbl_font_size
                    p.paragraph_format.space_before = tbl_space_before
                    p.paragraph_format.space_after = tbl_space_after
                    p.paragraph_format.line_spacing = tbl_line_spacing

    if skipped:
        st.warning("以下样式未被处理（已跳过）：")
        st.text("\n".join(sorted(skipped)))
    else:
        st.success("所有段落/表格样式均已成功处理！")

def process_doc(uploaded_bytes):
    doc = Document(BytesIO(uploaded_bytes))
    restructure_outline(doc)
    kill_all_numbering(doc)
    add_heading_numbers(doc)
    modify_document_format(doc)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------- Streamlit 界面 ----------------
st.title("Word 自动排版")

files = st.file_uploader("上传一个或多个 docx",
                         type=["docx"],
                         accept_multiple_files=True)

if files and st.button("开始批量排版"):
    if len(files) == 0:
        st.warning("请先上传文件")
        st.stop()

    with st.spinner(f"共 {len(files)} 个文件，正在逐个处理…"):
        for f in files:
            out_buffer = process_doc(f.read())
            st.download_button(
                label=f"下载 ➤ {f.name.replace('.docx', '')}_已排版.docx",
                data=out_buffer,
                file_name=f"{f.name.replace('.docx', '')}_已排版.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )











