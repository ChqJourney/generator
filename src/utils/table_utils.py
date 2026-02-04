"""
表格处理工具函数
"""
from typing import Any


def set_cell_value(cell: Any, value: str):
    """
    设置单元格值
    
    Args:
        cell: docx 单元格对象
        value: 要设置的值
    """
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.text = value
            return  # 找到run并设置后直接返回
    # 没有run则使用第一个paragraph
    if cell.paragraphs:
        cell.paragraphs[0].add_run(value)
    else:
        cell.add_paragraph(value)
