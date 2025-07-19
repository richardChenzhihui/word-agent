"""
word_editor/docx_utils.py

python-docx 全功能工具库，涵盖：
- 文档操作（新建、打开、保存、属性）
- 段落/标题/分页/列表
- 字体与样式
- 表格
- 图片/超链接
- 页眉页脚/节
- 高级文本替换

作者：AI助手
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from typing import Optional, List, Union, Dict, Any
from docx.table import Table
from docx.text.paragraph import Paragraph

# =====================
# 文档操作
# =====================
def create_new_docx() -> Document:
    print("[docx_utils] create_new_docx() called")
    """新建一个空白 Word 文档对象"""
    return Document()

def open_docx(path: str) -> Document:
    print(f"[docx_utils] open_docx(path={path}) called")
    """打开一个已有的 Word 文档"""
    return Document(path)

def save_docx(doc: Document, path: str):
    print(f"[docx_utils] save_docx(path={path}) called")
    """保存 Word 文档到指定路径"""
    doc.save(path)

def set_doc_property(doc: Document, prop: str, value: str):
    print(f"[docx_utils] set_doc_property(prop={prop}, value={value}) called")
    """设置文档属性（如作者、标题）"""
    setattr(doc.core_properties, prop, value)

def get_doc_property(doc: Document, prop: str) -> str:
    """获取文档属性"""
    return getattr(doc.core_properties, prop, "")

# =====================
# 段落/标题/分页/列表
# =====================
def add_paragraph(doc: Document, text: str, style: Optional[str] = None):
    print(f"[docx_utils] add_paragraph(text={text}, style={style}) called")
    """添加段落，可选样式"""
    return doc.add_paragraph(text, style=style)

def insert_paragraph_before(paragraph, text: str, style: Optional[str] = None):
    print(f"[docx_utils] insert_paragraph_before(text={text}, style={style}) called")
    """在指定段落前插入段落"""
    return paragraph.insert_paragraph_before(text, style=style)

def insert_paragraph_after(paragraph, text: str, style: Optional[str] = None):
    print(f"[docx_utils] insert_paragraph_after(text={text}, style={style}) called")
    """
    在指定段落后插入段落（兼容python-docx所有版本，底层XML实现）
    """
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    parent_elm = paragraph._element.getparent()
    idx = list(parent_elm).index(paragraph._element)
    new_p = OxmlElement('w:p')
    parent_elm.insert(idx + 1, new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    if style:
        new_para.style = style
    return new_para

def add_heading(doc: Document, text: str, level: int = 1):
    print(f"[docx_utils] add_heading(text={text}, level={level}) called")
    """添加标题，level=1~9"""
    return doc.add_heading(text, level=level)

def add_page_break(doc: Document):
    print(f"[docx_utils] add_page_break() called")
    """添加分页符"""
    doc.add_page_break()

def add_list(doc: Document, items: List[str], ordered: bool = False):
    print(f"[docx_utils] add_list(items={items}, ordered={ordered}) called")
    """添加有序/无序列表"""
    for i, item in enumerate(items, 1):
        if ordered:
            # 有序列表：手动添加数字前缀
            list_text = f"{i}. {item}"
        else:
            # 无序列表：手动添加项目符号
            list_text = f"• {item}"
        
        para = doc.add_paragraph(list_text)
        # 设置段落缩进以模拟列表效果
        para.paragraph_format.left_indent = Pt(18)  # 左缩进
        para.paragraph_format.first_line_indent = Pt(-18)  # 首行悬挂缩进

# =====================
# 字体与样式
# =====================
def normalize_font_name(font_name: str) -> str:
    """
    标准化字体名称，用于模糊匹配
    """
    return font_name.lower().replace(" ", "").replace("-", "").replace("_", "")

def find_best_font_match(requested_font: str) -> str:
    """
    查找最佳匹配的字体名称
    支持模糊匹配常见字体名称
    """
    # 常见字体映射表
    font_mappings = {
        # Times 系列
        "times": "Times New Roman",
        "timesnewroman": "Times New Roman",
        "timesroman": "Times New Roman",
        "times-roman": "Times New Roman",
        
        # Arial 系列
        "arial": "Arial",
        "arialnormal": "Arial",
        "arial-normal": "Arial",
        
        # Calibri 系列
        "calibri": "Calibri",
        "calibrinormal": "Calibri",
        
        # 宋体系列
        "simsun": "SimSun",
        "宋体": "SimSun",
        "songti": "SimSun",
        "sim-sun": "SimSun",
        
        # 黑体系列
        "simhei": "SimHei",
        "黑体": "SimHei",
        "heiti": "SimHei",
        "sim-hei": "SimHei",
        
        # 微软雅黑系列
        "microsoftyahei": "Microsoft YaHei",
        "yahei": "Microsoft YaHei",
        "雅黑": "Microsoft YaHei",
        "微软雅黑": "Microsoft YaHei",
        "msyh": "Microsoft YaHei",
        
        # 楷体系列
        "kaiti": "KaiTi",
        "楷体": "KaiTi",
        "kai-ti": "KaiTi",
        
        # Helvetica 系列
        "helvetica": "Helvetica",
        "helveticanormal": "Helvetica",
        
        # Georgia 系列
        "georgia": "Georgia",
        "georgianormal": "Georgia",
        
        # Verdana 系列
        "verdana": "Verdana",
        "verdananormal": "Verdana",
    }
    
    # 标准化输入字体名称
    normalized_requested = normalize_font_name(requested_font)
    
    # 直接匹配
    if normalized_requested in font_mappings:
        matched_font = font_mappings[normalized_requested]
        print(f"[DEBUG] Font '{requested_font}' matched to '{matched_font}'")
        return matched_font
    
    # 部分匹配
    for key, value in font_mappings.items():
        if normalized_requested in key or key in normalized_requested:
            print(f"[DEBUG] Font '{requested_font}' partially matched to '{value}'")
            return value
    
    # 如果没有匹配，返回原字体名称
    print(f"[DEBUG] Font '{requested_font}' not matched, using original name")
    return requested_font

def set_run_style(run, font_name=None, font_size=None, bold=None, italic=None, underline=None, color=None):
    print(f"[docx_utils] set_run_style(font_name={font_name}, font_size={font_size}, bold={bold}, italic={italic}, underline={underline}, color={color}) called")
    """设置 run 的字体、字号、加粗、斜体、下划线、颜色"""
    try:
        font = run.font
        if font_name:
            # 使用模糊匹配查找最佳字体
            matched_font = find_best_font_match(font_name)
            font.name = matched_font
            # 设置字体族，确保兼容性
            font._element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', matched_font)
            font._element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', matched_font)
            font._element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', matched_font)
            font._element.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cs', matched_font)
        if font_size:
            font.size = Pt(font_size)
        if bold is not None:
            font.bold = bold
        if italic is not None:
            font.italic = italic
        if underline is not None:
            font.underline = underline
        if color:
            if isinstance(color, str) and color.startswith('#'):
                color = RGBColor.from_string(color[1:])
            font.color.rgb = color
    except Exception as e:
        print(f"[ERROR] Failed to set run style: {e}")

def set_paragraph_style(paragraph, alignment=None, line_spacing=None, indent=None):
    print(f"[docx_utils] set_paragraph_style(alignment={alignment}, line_spacing={line_spacing}, indent={indent}) called")
    """设置段落对齐、行距、首行缩进"""
    if alignment:
        paragraph.alignment = alignment
    if line_spacing:
        paragraph.paragraph_format.line_spacing = line_spacing
    if indent:
        paragraph.paragraph_format.first_line_indent = Pt(indent)

def change_font_comprehensive(doc: Document, font_name: str, target_type: str = "all", target_indices: list = None):
    """
    全面的字体更改功能，支持局部和全局变更
    
    参数:
        doc: Document对象
        font_name: 字体名称（支持模糊匹配）
        target_type: 目标类型 - "all"(全部), "paragraphs"(段落), "tables"(表格), "headers"(页眉), "footers"(页脚)
        target_indices: 目标索引列表，仅在target_type不为"all"时生效
    """
    print(f"[docx_utils] change_font_comprehensive(font_name={font_name}, target_type={target_type}, target_indices={target_indices}) called")
    
    # 使用模糊匹配查找最佳字体
    matched_font = find_best_font_match(font_name)
    changed_count = 0
    
    def change_runs_font(runs):
        """更改runs列表中所有run的字体"""
        nonlocal changed_count
        for run in runs:
            try:
                set_run_style(run, font_name=matched_font)
                changed_count += 1
            except Exception as e:
                print(f"[WARNING] Failed to change font for run: {e}")
    
    def change_paragraph_font(paragraph):
        """更改段落字体"""
        change_runs_font(paragraph.runs)
    
    def change_table_font(table):
        """更改表格字体"""
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    change_runs_font(para.runs)
    
    # 根据target_type执行不同的字体更改策略
    if target_type == "all":
        # 全局更改：遍历所有可能的文本元素
        
        # 1. 主文档段落
        for para in doc.paragraphs:
            change_paragraph_font(para)
        
        # 2. 主文档表格
        for table in doc.tables:
            change_table_font(table)
        
        # 3. 页眉页脚
        try:
            for section in doc.sections:
                # 页眉
                if section.header:
                    for para in section.header.paragraphs:
                        change_paragraph_font(para)
                    for table in section.header.tables:
                        change_table_font(table)
                
                # 页脚
                if section.footer:
                    for para in section.footer.paragraphs:
                        change_paragraph_font(para)
                    for table in section.footer.tables:
                        change_table_font(table)
        except Exception as e:
            print(f"[WARNING] Could not change header/footer font: {e}")
        
        # 4. 使用更深层的XML遍历确保不遗漏
        try:
            from docx.oxml.ns import qn
            # 遍历所有文本运行元素
            for element in doc.element.iter():
                if element.tag == qn('w:r'):  # 文本运行元素
                    try:
                        # 获取或创建字体属性
                        rPr = element.find(qn('w:rPr'))
                        if rPr is None:
                            rPr = element.makeelement(qn('w:rPr'))
                            element.insert(0, rPr)
                        
                        # 设置字体
                        fonts = rPr.find(qn('w:rFonts'))
                        if fonts is None:
                            fonts = rPr.makeelement(qn('w:rFonts'))
                            rPr.append(fonts)
                        
                        fonts.set(qn('w:ascii'), matched_font)
                        fonts.set(qn('w:eastAsia'), matched_font)
                        fonts.set(qn('w:hAnsi'), matched_font)
                        fonts.set(qn('w:cs'), matched_font)
                        
                    except Exception as e:
                        print(f"[WARNING] XML-level font change failed: {e}")
        except Exception as e:
            print(f"[WARNING] Deep XML traversal failed: {e}")
    
    elif target_type == "paragraphs":
        # 段落字体更改
        paragraphs = doc.paragraphs
        if target_indices:
            for idx in target_indices:
                if 0 <= idx < len(paragraphs):
                    change_paragraph_font(paragraphs[idx])
        else:
            for para in paragraphs:
                change_paragraph_font(para)
    
    elif target_type == "tables":
        # 表格字体更改
        tables = doc.tables
        if target_indices:
            for idx in target_indices:
                if 0 <= idx < len(tables):
                    change_table_font(tables[idx])
        else:
            for table in tables:
                change_table_font(table)
    
    elif target_type == "headers":
        # 页眉字体更改
        try:
            sections = doc.sections
            if target_indices:
                for idx in target_indices:
                    if 0 <= idx < len(sections) and sections[idx].header:
                        for para in sections[idx].header.paragraphs:
                            change_paragraph_font(para)
                        for table in sections[idx].header.tables:
                            change_table_font(table)
            else:
                for section in sections:
                    if section.header:
                        for para in section.header.paragraphs:
                            change_paragraph_font(para)
                        for table in section.header.tables:
                            change_table_font(table)
        except Exception as e:
            print(f"[WARNING] Header font change failed: {e}")
    
    elif target_type == "footers":
        # 页脚字体更改
        try:
            sections = doc.sections
            if target_indices:
                for idx in target_indices:
                    if 0 <= idx < len(sections) and sections[idx].footer:
                        for para in sections[idx].footer.paragraphs:
                            change_paragraph_font(para)
                        for table in sections[idx].footer.tables:
                            change_table_font(table)
            else:
                for section in sections:
                    if section.footer:
                        for para in section.footer.paragraphs:
                            change_paragraph_font(para)
                        for table in section.footer.tables:
                            change_table_font(table)
        except Exception as e:
            print(f"[WARNING] Footer font change failed: {e}")
    
    print(f"[DEBUG] Changed font to {matched_font} for {changed_count} text runs (target: {target_type})")
    return changed_count

def change_document_font(doc: Document, font_name: str):
    """
    更改整个文档的字体（兼容性函数，调用新的comprehensive函数）
    """
    print(f"[docx_utils] change_document_font(font_name={font_name}) called")
    return change_font_comprehensive(doc, font_name, target_type="all")

def change_paragraph_font(doc: Document, font_name: str, paragraph_indices: list = None):
    """
    更改指定段落的字体
    """
    print(f"[docx_utils] change_paragraph_font(font_name={font_name}, paragraph_indices={paragraph_indices}) called")
    return change_font_comprehensive(doc, font_name, target_type="paragraphs", target_indices=paragraph_indices)

def change_table_font(doc: Document, font_name: str, table_indices: list = None):
    """
    更改指定表格的字体
    """
    print(f"[docx_utils] change_table_font(font_name={font_name}, table_indices={table_indices}) called")
    return change_font_comprehensive(doc, font_name, target_type="tables", target_indices=table_indices)

# =====================
# 表格
# =====================
def add_table(doc: Document, rows: int, cols: int, style: Optional[str] = None):
    print(f"[docx_utils] add_table(rows={rows}, cols={cols}, style={style}) called")
    """添加表格，可选样式"""
    table = doc.add_table(rows=rows, cols=cols, style=style)
    return table

def set_cell_text(cell, text: str, style: Optional[str] = None):
    print(f"[docx_utils] set_cell_text(text={text}, style={style}) called")
    """设置单元格内容，可选样式"""
    cell.text = text
    if style:
        for p in cell.paragraphs:
            p.style = style

def merge_cells(table, start_row, start_col, end_row, end_col):
    print(f"[docx_utils] merge_cells(start_row={start_row}, start_col={start_col}, end_row={end_row}, end_col={end_col}) called")
    """合并表格单元格（左上到右下）"""
    cell1 = table.cell(start_row, start_col)
    cell2 = table.cell(end_row, end_col)
    cell1.merge(cell2)

def set_table_style(table, style: str):
    print(f"[docx_utils] set_table_style(style={style}) called")
    """设置表格样式"""
    table.style = style

# =====================
# 图片/超链接
# =====================
def add_picture(doc: Document, image_path: str, width=None, height=None):
    print(f"[docx_utils] add_picture(image_path={image_path}, width={width}, height={height}) called")
    """插入图片，支持设置宽高（Inches单位）"""
    if width:
        width = Inches(width)
    if height:
        height = Inches(height)
    doc.add_picture(image_path, width=width, height=height)

# 超链接插入需用底层xml hack，简单实现如下：
def add_hyperlink(paragraph, url: str, text: str):
    print(f"[docx_utils] add_hyperlink(url={url}, text={text}) called")
    """在段落中插入超链接（简单实现）"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import RGBColor
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    # 超链接样式
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')
    rPr.append(color)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

# =====================
# 页眉页脚/节
# =====================
def get_header(doc: Document):
    print(f"[docx_utils] get_header() called")
    """获取文档第一个节的页眉对象"""
    return doc.sections[0].header

def get_footer(doc: Document):
    print(f"[docx_utils] get_footer() called")
    """获取文档第一个节的页脚对象"""
    return doc.sections[0].footer

def set_section_properties(section, orientation=None, page_width=None, page_height=None, margin=None):
    print(f"[docx_utils] set_section_properties(orientation={orientation}, page_width={page_width}, page_height={page_height}, margin={margin}) called")
    """设置节的方向、纸张大小、页边距
    orientation: WD_ORIENT.LANDSCAPE/PORTRAIT
    page_width/page_height: 毫米
    margin: (top, bottom, left, right) 毫米
    """
    if orientation:
        section.orientation = orientation
    if page_width:
        section.page_width = page_width
    if page_height:
        section.page_height = page_height
    if margin:
        section.top_margin = margin[0]
        section.bottom_margin = margin[1]
        section.left_margin = margin[2]
        section.right_margin = margin[3]

def find_and_replace_text_smart(doc: Document, search_text: str, replace_text: str, nth: Union[int, None] = 1) -> int:
    """
    智能文本查找替换，处理跨runs的文本和格式问题
    """
    print(f"[docx_utils] find_and_replace_text_smart(search_text={search_text[:50]}..., replace_text={replace_text[:50]}..., nth={nth}) called")
    
    total_replaced = 0
    
    # 处理段落中的文本
    for para in doc.paragraphs:
        if search_text in para.text:
            replaced = replace_text_preserve_format(para, search_text, replace_text, nth=None if nth is None else nth-total_replaced)
            total_replaced += replaced
            if nth is not None and total_replaced >= nth:
                break
    
    # 处理表格中的文本
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if search_text in para.text:
                        replaced = replace_text_preserve_format(para, search_text, replace_text, nth=None if nth is None else nth-total_replaced)
                        total_replaced += replaced
                        if nth is not None and total_replaced >= nth:
                            return total_replaced
    
    print(f"[DEBUG] Smart replace total: {total_replaced}")
    return total_replaced

# =====================
# 高级文本替换（保留格式，见前述实现）
# =====================
def replace_text_preserve_format(paragraph, search_text: str, replace_text: str, nth: Union[int, None] = 1, case_sensitive: bool = False) -> int:
    print(f"[docx_utils] replace_text_preserve_format(search_text={search_text[:50]}..., replace_text={replace_text[:50]}..., nth={nth}, case_sensitive={case_sensitive}) called")
    """
    在段落中查找并替换第 nth 个 search_text，并用 replace_text 替换，最大限度保留原有格式。
    参数：
        paragraph: docx.paragraph.Paragraph 对象
        search_text: 要查找的字符串
        replace_text: 替换为的字符串
        nth: 替换第几个出现（默认第1个，None表示全部替换）
        case_sensitive: 是否区分大小写（默认False）
    返回：
        替换次数
    """
    if not search_text:
        return 0
    
    # 构建完整文本和run索引
    full_text = ""
    run_indices = []
    for run in paragraph.runs:
        start = len(full_text)
        full_text += run.text
        end = len(full_text)
        run_indices.append((run, start, end))
    
    # 调试输出
    print(f"[DEBUG] Paragraph text: {full_text[:100]}...")
    print(f"[DEBUG] Search text: {search_text[:50]}...")
    print(f"[DEBUG] Paragraph has {len(paragraph.runs)} runs")
    
    # 查找所有匹配位置（支持大小写不敏感）
    match_positions = []
    search_lower = search_text.lower() if not case_sensitive else search_text
    text_to_search = full_text.lower() if not case_sensitive else full_text
    
    idx = 0
    while True:
        idx = text_to_search.find(search_lower, idx)
        if idx == -1:
            break
        match_positions.append(idx)
        idx += len(search_text) if len(search_text) > 0 else 1
    
    print(f"[DEBUG] Found {len(match_positions)} matches")
    
    if not match_positions:
        return 0
    
    # 根据nth参数选择要替换的位置
    if nth is not None:
        if nth <= 0 or nth > len(match_positions):
            return 0
        match_positions = [match_positions[nth-1]]
    
    replaced_count = 0
    # 从后往前替换，避免位置偏移
    for start_pos in reversed(match_positions):
        end_pos = start_pos + len(search_text)
        
        # 找到受影响的runs
        affected = []
        for i, (run, s, e) in enumerate(run_indices):
            if s < end_pos and e > start_pos:
                affected.append((i, run, s, e))
        
        if not affected:
            continue
        
        # 处理替换
        first_i, first_run, first_s, first_e = affected[0]
        last_i, last_run, last_s, last_e = affected[-1]
        
        # 计算前缀和后缀
        prefix = first_run.text[:start_pos - first_s] if start_pos > first_s else ""
        suffix = last_run.text[end_pos - last_s:] if end_pos < last_e else ""
        
        # 执行替换
        first_run.text = prefix + replace_text + suffix
        
        # 清空其他受影响的runs
        for i, run, s, e in affected[1:]:
            run.text = ""
        
        replaced_count += 1
        
        # 更新full_text用于下一次替换
        full_text = "".join(run.text for run, _, _ in run_indices)
    
    print(f"[DEBUG] Replaced {replaced_count} instances")
    return replaced_count

def replace_text_in_paragraphs(doc: Document, search_text: str, replace_text: str, nth: Union[int, None] = None, case_sensitive: bool = False) -> int:
    """
    在文档段落中查找并替换文本
    """
    print(f"[docx_utils] replace_text_in_paragraphs(search_text={search_text}, replace_text={replace_text}, nth={nth}, case_sensitive={case_sensitive}) called")
    total = 0
    for para in doc.paragraphs:
        total += replace_text_preserve_format(para, search_text, replace_text, nth=None if nth is None else nth-total, case_sensitive=case_sensitive)
        if nth is not None and total >= nth:
            break
    return total

def replace_text_in_tables(doc: Document, search_text: str, replace_text: str, nth: Union[int, None] = None, case_sensitive: bool = False) -> int:
    """
    在文档表格中查找并替换文本
    """
    print(f"[docx_utils] replace_text_in_tables(search_text={search_text}, replace_text={replace_text}, nth={nth}, case_sensitive={case_sensitive}) called")
    total = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    total += replace_text_preserve_format(para, search_text, replace_text, nth=None if nth is None else nth-total, case_sensitive=case_sensitive)
                    if nth is not None and total >= nth:
                        return total
    return total 

def delete_paragraph(doc: Document, index: int):
    print(f"[docx_utils] delete_paragraph(index={index}) called")
    """
    删除指定索引的段落。
    用法: delete_paragraph(doc, 2) # 删除第3个段落
    """
    p = doc.paragraphs[index]
    p._element.getparent().remove(p._element)


def move_paragraph(doc: Document, from_idx: int, to_idx: int):
    print(f"[docx_utils] move_paragraph(from_idx={from_idx}, to_idx={to_idx}) called")
    """
    将from_idx段落移动到to_idx位置。
    用法: move_paragraph(doc, 2, 0) # 第3段移到最前
    """
    paragraphs = doc.paragraphs
    p = paragraphs[from_idx]
    elm = p._element
    parent = elm.getparent()
    parent.remove(elm)
    if to_idx >= len(paragraphs):
        doc._body._element.append(elm)
    else:
        ref_elm = paragraphs[to_idx]._element
        ref_elm.addprevious(elm)


def copy_paragraph(doc: Document, from_idx: int, to_idx: int):
    print(f"[docx_utils] copy_paragraph(from_idx={from_idx}, to_idx={to_idx}) called")
    """
    复制from_idx段落到to_idx后。
    用法: copy_paragraph(doc, 1, 2)
    """
    from copy import deepcopy
    paragraphs = doc.paragraphs
    p = paragraphs[from_idx]
    new_p = deepcopy(p._element)
    if to_idx >= len(paragraphs):
        doc._body._element.append(new_p)
    else:
        ref_elm = paragraphs[to_idx]._element
        ref_elm.addnext(new_p)


def find_paragraphs(doc: Document, keyword: str) -> list:
    """
    查找包含关键字的段落索引列表。
    用法: idxs = find_paragraphs(doc, '目标')
    """
    return [i for i, p in enumerate(doc.paragraphs) if keyword in p.text]


def batch_replace(doc: Document, replace_map: dict, case_sensitive: bool = False):
    print(f"[docx_utils] batch_replace(replace_map={replace_map}, case_sensitive={case_sensitive}) called")
    """
    批量查找替换，replace_map={'old1':'new1', ...}
    
    参数:
        doc: Document对象
        replace_map: 替换映射字典 {'旧文本': '新文本', ...}
        case_sensitive: 是否区分大小写（默认False）
    
    返回:
        替换的总次数
    """
    total_replaced = 0
    for old, new in replace_map.items():
        print(f"[DEBUG] 批量替换: '{old}' -> '{new}'")
        
        # 使用统一的替换函数
        replaced_count = replace_text_everywhere(doc, old, new, nth=None, case_sensitive=case_sensitive)
        print(f"[DEBUG] 替换了 {replaced_count} 处")
        
        total_replaced += replaced_count
    
    print(f"[DEBUG] 批量替换总计: {total_replaced} 处")
    return total_replaced

# =====================
# 兼容性函数（保持向后兼容）
# =====================

def replace_text_in_docx(doc: Document, search_text: str, replace_text: str, nth: Union[int, None] = None, case_sensitive: bool = False) -> int:
    """兼容性函数：在文档段落中替换文本"""
    return replace_text_in_paragraphs(doc, search_text, replace_text, nth, case_sensitive)

def comprehensive_replace(doc: Document, search_text: str, replace_text: str, case_sensitive: bool = False) -> int:
    """兼容性函数：全面文本替换"""
    return replace_text_everywhere(doc, search_text, replace_text, nth=None, case_sensitive=case_sensitive)

def replace_text_in_headers_footers(doc: Document, search_text: str, replace_text: str, case_sensitive: bool = False) -> int:
    """
    在页眉页脚中查找并替换文本
    """
    print(f"[docx_utils] replace_text_in_headers_footers(search_text={search_text}, replace_text={replace_text}, case_sensitive={case_sensitive}) called")
    total = 0
    
    try:
        for section in doc.sections:
            # 处理页眉
            if section.header:
                for para in section.header.paragraphs:
                    total += replace_text_preserve_format(para, search_text, replace_text, nth=None, case_sensitive=case_sensitive)
                for table in section.header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                total += replace_text_preserve_format(para, search_text, replace_text, nth=None, case_sensitive=case_sensitive)
            
            # 处理页脚
            if section.footer:
                for para in section.footer.paragraphs:
                    total += replace_text_preserve_format(para, search_text, replace_text, nth=None, case_sensitive=case_sensitive)
                for table in section.footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                total += replace_text_preserve_format(para, search_text, replace_text, nth=None, case_sensitive=case_sensitive)
    except Exception as e:
        print(f"[WARNING] 页眉页脚替换失败: {e}")
    
    return total

def replace_text_everywhere(doc: Document, search_text: str, replace_text: str, nth: Union[int, None] = None, case_sensitive: bool = False) -> int:
    """
    在整个文档中查找并替换文本（段落、表格、页眉页脚）
    
    参数:
        doc: Document对象
        search_text: 要查找的文本
        replace_text: 替换的文本
        nth: 替换第几个匹配项（None表示全部替换）
        case_sensitive: 是否区分大小写（默认False）
    
    返回:
        替换的总次数
    """
    print(f"[docx_utils] replace_text_everywhere(search_text={search_text}, replace_text={replace_text}, nth={nth}, case_sensitive={case_sensitive}) called")
    
    total = 0
    
    # 段落中替换
    total += replace_text_in_paragraphs(doc, search_text, replace_text, nth=None if nth is None else nth-total, case_sensitive=case_sensitive)
    if nth is not None and total >= nth:
        return total
    
    # 表格中替换
    total += replace_text_in_tables(doc, search_text, replace_text, nth=None if nth is None else nth-total, case_sensitive=case_sensitive)
    if nth is not None and total >= nth:
        return total
    
    # 页眉页脚中替换
    total += replace_text_in_headers_footers(doc, search_text, replace_text, case_sensitive=case_sensitive)
    
    print(f"[DEBUG] 全面替换 '{search_text}' -> '{replace_text}' 总计: {total} 处")
    return total


def clear_formatting(paragraph):
    print(f"[docx_utils] clear_formatting() called")
    """
    清除段落所有格式（仅保留纯文本）。
    用法: clear_formatting(doc.paragraphs[0])
    """
    text = paragraph.text
    for run in paragraph.runs:
        run.clear()
    paragraph.clear()
    paragraph.add_run(text)


def extract_paragraphs(doc: Document, indices: list) -> Document:
    print(f"[docx_utils] extract_paragraphs(indices={indices}) called")
    """
    提取指定段落到新文档。
    用法: new_doc = extract_paragraphs(doc, [0,2])
    """
    new_doc = create_new_docx()
    for i in indices:
        new_doc.add_paragraph(doc.paragraphs[i].text)
    return new_doc


def merge_documents(doc_paths: list, output_path: str):
    print(f"[docx_utils] merge_documents(doc_paths={doc_paths}, output_path={output_path}) called")
    """
    合并多个docx文档为一个，保留格式。
    用法: merge_documents(['a.docx','b.docx'],'out.docx')
    """
    merged = create_new_docx()
    for path in doc_paths:
        d = open_docx(path)
        for p in d.paragraphs:
            merged.add_paragraph(p.text, style=p.style)
        for t in d.tables:
            tbl = merged.add_table(rows=len(t.rows), cols=len(t.columns))
            for i, row in enumerate(t.rows):
                for j, cell in enumerate(row.cells):
                    tbl.cell(i,j).text = cell.text
    save_docx(merged, output_path)


def split_document_by_heading(doc: Document, level: int) -> list:
    print(f"[docx_utils] split_document_by_heading(level={level}) called")
    """
    按指定级别标题拆分文档，返回新文档列表。
    用法: docs = split_document_by_heading(doc, 1)
    """
    docs = []
    curr_doc = None
    for p in doc.paragraphs:
        if p.style.name == f'Heading {level}':
            if curr_doc:
                docs.append(curr_doc)
            curr_doc = create_new_docx()
            curr_doc.add_paragraph(p.text, style=p.style)
        elif curr_doc:
            curr_doc.add_paragraph(p.text, style=p.style)
    if curr_doc:
        docs.append(curr_doc)
    return docs

# =====================
# 文档结构导出（结构化，供LLM/agent理解）
# =====================
def export_docx_structure(doc: Document) -> Dict[str, Any]:
    """
    导出 Word 文档结构和内容，便于 LLM/agent 解析和后续操作。
    返回结构示例：
    {
        "blocks": [
            {"type": "paragraph", "index": 0, "block_idx": 0, "text": "...", "style": "...", "heading_level": 1/None, "raw_object_type": "paragraph"},
            {"type": "table", "index": 0, "block_idx": 1, "rows": [[...], [...]], "raw_object_type": "table"},
            ...
        ]
    }
    """
    blocks = []
    para_idx = 0
    table_idx = 0
    for block in _iter_block_items(doc):
        if isinstance(block, Paragraph):
            style = block.style.name if block.style else None
            heading_level = None
            if style and style.startswith("Heading"):
                try:
                    heading_level = int(style.split()[-1])
                except Exception:
                    heading_level = None
            blocks.append({
                "type": "paragraph",
                "index": para_idx,
                "block_idx": len(blocks),
                "text": block.text,
                "style": style,
                "heading_level": heading_level,
                "raw_object_type": "paragraph"
            })
            para_idx += 1
        elif isinstance(block, Table):
            table_data = []
            for row in block.rows:
                table_data.append([cell.text for cell in row.cells])
            blocks.append({
                "type": "table",
                "index": table_idx,
                "block_idx": len(blocks),
                "rows": table_data,
                "raw_object_type": "table"
            })
            table_idx += 1
    return {"blocks": blocks}

def _iter_block_items(parent):
    """
    迭代文档中的段落和表格，保持顺序。
    """
    for child in parent.element.body.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)

# =====================
# 用法示范（可在独立脚本或交互环境中参考）
# =====================
"""
# 示例1：删除、移动、复制段落
from docx_utils import *
doc = create_new_docx()
add_paragraph(doc, 'A')
add_paragraph(doc, 'B')
add_paragraph(doc, 'C')
delete_paragraph(doc, 1)  # 删除B
move_paragraph(doc, 1, 0)  # C移到最前
copy_paragraph(doc, 1, 1)  # 再复制C到C后
save_docx(doc, 'demo.docx')

# 示例2：批量替换、清除格式
batch_replace(doc, {'A':'Alpha', 'C':'Charlie'})
clear_formatting(doc.paragraphs[0])

# 示例3：内容提取、合并
new_doc = extract_paragraphs(doc, [0,2])
merge_documents(['demo.docx','demo.docx'],'merged.docx')

# 示例4：按标题拆分
parts = split_document_by_heading(doc, 1)
for i, d in enumerate(parts):
    save_docx(d, f'part_{i}.docx')
""" 