"""
word_editor/test_docx_utils_office_agent.py

办公Agent扩展功能紧凑测试集，验证段落删除、移动、复制、查找、批量替换、内容提取、合并、清除格式、按标题拆分等。
"""
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))
from docx_utils import *

def main():
    doc = create_new_docx()
    add_heading(doc, '第一章', level=1)
    add_paragraph(doc, 'A')
    add_paragraph(doc, 'B')
    add_paragraph(doc, 'C')
    add_heading(doc, '第二章', level=1)
    add_paragraph(doc, 'D')
    add_paragraph(doc, 'E')
    save_docx(doc, 'word_editor/test_office_agent_add_para.docx')
    # 删除
    delete_paragraph(doc, 2)  # 删除B
    save_docx(doc, 'word_editor/test_office_agent_copy_para_0.docx')
    # 移动
    move_paragraph(doc, 3, 1)  # E移到A后
    save_docx(doc, 'word_editor/test_office_agent_copy_para_1.docx')
    # 复制
    copy_paragraph(doc, 1, 2)  # A复制到C后
    save_docx(doc, 'word_editor/test_office_agent_copy_para_3.docx')
    # 查找
    idxs = find_paragraphs(doc, 'A')
    print(f'包含A的段落索引: {idxs}')
    # 批量替换
    batch_replace(doc, {'A':'Alpha', 'C':'Charlie'})
    # 清除格式
    clear_formatting(doc.paragraphs[0])
    # 内容提取
    new_doc = extract_paragraphs(doc, [0,2])
    save_docx(new_doc, 'word_editor/test_office_agent_extract.docx')
    # 合并
    merge_documents(['word_editor/test_office_agent_extract.docx','word_editor/test_office_agent_extract.docx'],
                   'word_editor/test_office_agent_merged.docx')
    # 按标题拆分
    parts = split_document_by_heading(doc, 1)
    for i, d in enumerate(parts):
        save_docx(d, f'word_editor/test_office_agent_part_{i}.docx')
    # 保存主文档
    save_docx(doc, 'word_editor/test_office_agent_main.docx')
    print('办公Agent扩展功能测试完成，请检查word_editor目录下相关输出文件。')

if __name__ == '__main__':
    main() 