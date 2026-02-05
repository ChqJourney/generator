from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
import sys
import shutil
import os
from typing import List, Dict, Optional, Tuple, Any
from abc import ABC, abstractmethod
from copy import deepcopy
from table_processor import TableDataTransformer
from utils.table_utils import set_cell_value

class DocxTemplateError(Exception):
    pass

class PlaceholderNotFoundError(DocxTemplateError):
    def __init__(self, placeholder: str, location: str = 'body'):
        self.placeholder = placeholder
        self.location = location
        super().__init__(f"Placeholder '{placeholder}' not found in {location}")

class InvalidLocationError(DocxTemplateError):
    def __init__(self, location: str, content_type: str):
        super().__init__(f"Invalid location '{location}' for {content_type}. Valid locations: body, header, footer")

class PlaceholderFinder:
    @staticmethod
    def _iterate_container(container):
        yield from container.paragraphs
        for table_idx, table in enumerate(container.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    for p_idx, paragraph in enumerate(cell.paragraphs):
                        yield (table_idx, row_idx, cell_idx, p_idx), paragraph

    @staticmethod
    def _search_paragraphs_in_container(container, placeholder):
        for i, paragraph in enumerate(container.paragraphs):
            if placeholder in paragraph.text:
                yield i, paragraph
        for table_idx, table in enumerate(container.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    for p_idx, paragraph in enumerate(cell.paragraphs):
                        if placeholder in paragraph.text:
                            yield (table_idx, row_idx, cell_idx, p_idx), paragraph

    @staticmethod
    def find_all_placeholders_in_location(doc: Document, placeholder: str, location: str = 'body') -> List[Tuple[Any, Any]]:
        results = []
        if location == 'body':
            results.extend(list(PlaceholderFinder._search_paragraphs_in_container(doc, placeholder)))
        elif location == 'header':
            for section in doc.sections:
                for header in [section.header, section.first_page_header, section.even_page_header]:
                    if header:
                        results.extend(list(PlaceholderFinder._search_paragraphs_in_container(header, placeholder)))
        elif location == 'footer':
            for section in doc.sections:
                for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
                    if footer:
                        results.extend(list(PlaceholderFinder._search_paragraphs_in_container(footer, placeholder)))
        return results

    @staticmethod
    def find_paragraph_with_placeholder(doc: Document, placeholder: str, location: str = 'body') -> Tuple[Any, Any]:
        results = PlaceholderFinder.find_all_placeholders_in_location(doc, placeholder, location)
        return results[0] if results else (None, None)

    @staticmethod
    def replace_paragraph_with_element(paragraph, element, location: str = 'body'):
        try:
            p_element = paragraph._element
            p_parent = p_element.getparent()
            
            if p_parent is None:
                raise DocxTemplateError("Cannot find parent element of paragraph")
            
            # 获取父容器中的索引位置
            try:
                index = list(p_parent).index(p_element)
            except ValueError:
                # 如果在父容器中找不到，可能是因为paragraph在特殊容器中
                raise DocxTemplateError(f"Paragraph not found in parent container")
            
            # 移除原段落
            p_parent.remove(p_element)
            
            # 插入新元素
            # 注意：对于表格元素，需要使用深拷贝以避免重复引用
            new_element = deepcopy(element)
            p_parent.insert(index, new_element)
            
        except (AttributeError, ValueError) as e:
            raise DocxTemplateError(f"Failed to replace paragraph with element: {str(e)}")

    @staticmethod
    def _iterate_placeholders(doc, location):
        if location == 'body':
            yield from doc.paragraphs
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        yield from cell.paragraphs
        elif location in ('header', 'footer'):
            for section in doc.sections:
                container = getattr(section, location + 's')
                yield from container.paragraphs
                for table in container.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            yield from cell.paragraphs

    @staticmethod
    def _replace_in_paragraph(paragraph, placeholder, value):
        if placeholder in paragraph.text:
            runs = paragraph.runs
            if runs:
                for run in runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, value)
                        return True
        return False

class ContentInserter(ABC):
    def __init__(self, doc: Document):
        self.doc = doc
    
    @abstractmethod
    def insert(self, *args, **kwargs):
        pass
    
    def validate_location(self, location: str, valid_locations: List[str]):
        if location not in valid_locations:
            raise InvalidLocationError(location, self.__class__.__name__)

class TextInserter(ContentInserter):
    def insert(self, placeholder: str, value: str, location: str = 'body'):
        self.validate_location(location, ['body', 'header', 'footer'])
        
        # 始终在所有位置查找占位符
        results = []
        searched_locations = []
        
        for loc in ['body', 'header', 'footer']:
            loc_results = PlaceholderFinder.find_all_placeholders_in_location(self.doc, placeholder, loc)
            if loc_results:
                searched_locations.append(loc)
                results.extend(loc_results)
                print(f"在 {loc} 中找到占位符 '{placeholder}' ({len(loc_results)} 处)")
        
        if not results:
            print(f"警告: 占位符 '{placeholder}' 在所有位置都未找到，跳过此操作")
            return
        
        # 直接在所有找到的paragraphs上尝试替换
        # _replace_in_paragraph会检查placeholder是否存在，不存在则返回False
        replaced_count = 0
        for idx, paragraph in results:
            if TextInserter._replace_in_paragraph(paragraph, placeholder, value):
                replaced_count += 1
        
        print(f"成功替换占位符 '{placeholder}' {replaced_count} 处，总计 {len(results)} 处")
    
    def _iterate_placeholders(self, location):
        if location == 'body':
            yield from self.doc.paragraphs
            for table in self.doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        yield from cell.paragraphs
        elif location == 'header':
            for section in self.doc.sections:
                for header in [section.header, section.first_page_header, section.even_page_header]:
                    if header:
                        yield from header.paragraphs
                        for table in header.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    yield from cell.paragraphs
        elif location == 'footer':
            for section in self.doc.sections:
                for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
                    if footer:
                        yield from footer.paragraphs
                        for table in footer.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    yield from cell.paragraphs

    @staticmethod
    def _replace_in_paragraph(paragraph, placeholder, value):
        full_placeholder = '{{' + placeholder + '}}'
        if full_placeholder in paragraph.text:
            runs = paragraph.runs
            if runs:
                for run in runs:
                    if full_placeholder in run.text:
                        run.text = run.text.replace(full_placeholder, value)
                        return True
                    elif placeholder in run.text:
                        run.text = run.text.replace(placeholder, value)
                        return True
        return False

class TableInserter(ContentInserter):
    def insert(self, placeholder: str, table_template_path: str, 
               raw_data: Optional[List[List[str]]] = None,
               transformations: Optional[List[Dict]] = None,
               metadata: Optional[Dict] = None,
                targets_data: Optional[Dict] = None,
               row_strategy: str = 'fixed_rows',
               skip_columns: Optional[List[int]] = None,
               header_rows: int = 1,
               location: str = 'body'):
        self.validate_location(location, ['body'])
        
        if not os.path.exists(table_template_path):
            raise DocxTemplateError(f"Table template file not found: {table_template_path}")
        
        transformer = TableDataTransformer()
        
        processed_data = raw_data
        if raw_data and transformations:
            processed_data = transformer.transform(raw_data, transformations, metadata, targets_data)
            
        
        table_template = Document(table_template_path)
        if not table_template.tables:
            raise DocxTemplateError(f"No tables found in template file: {table_template_path}")
        
        template_table = table_template.tables[0]
        
        if row_strategy == 'fixed_rows':
            self._fill_fixed_rows(template_table, processed_data, skip_columns, header_rows)
        elif row_strategy == 'dynamic_rows':
            self._fill_dynamic_rows(template_table, processed_data, skip_columns, header_rows)
        
        results = PlaceholderFinder.find_all_placeholders_in_location(self.doc, placeholder, location)
        
        if not results:
            print(f"警告: 在 {location} 中未找到占位符 '{placeholder}'，尝试在 body 中查找...")
            if location != 'body':
                results = PlaceholderFinder.find_all_placeholders_in_location(self.doc, placeholder, 'body')
                if results:
                    print(f"在 body 中找到占位符 '{placeholder}'")
                    location = 'body'
        
        if not results:
            print(f"警告: 占位符 '{placeholder}' 未找到，跳过表格插入操作")
            return
        
        for idx, paragraph in results:
            try:
                PlaceholderFinder.replace_paragraph_with_element(paragraph, template_table._element)
            except (AttributeError, ValueError, TypeError, DocxTemplateError) as e:
                try:
                    print(f"尝试在单元格内插入表格 '{placeholder}'...")
                    self._insert_table_in_cell(paragraph, template_table)
                except Exception as e2:
                    print(f"警告: 无法替换占位符 '{placeholder}' 为表格: {str(e)}")
                    print(f"      尝试在单元格内插入也失败: {str(e2)}")
                    continue
    
    def _fill_fixed_rows(self, table: Any, data: List[List[Any]], skip_columns: Optional[List[int]], header_rows: int):
        if not data:
            return
        
        data_row_idx = 0
        for row_idx, row in enumerate(table.rows):
            if row_idx < header_rows:
                continue
            
            if data_row_idx >= len(data):
                break
            
            data_row = data[data_row_idx]
            data_col_idx = 0
            
            for col_idx, cell in enumerate(row.cells):
                if skip_columns and col_idx in skip_columns:
                    continue
                
                if data_col_idx < len(data_row):
                    value = data_row[data_col_idx]
                    if value is None or value == '':
                        pass
                    else:
                        self._set_cell_value(cell, str(value))
                    data_col_idx += 1
            
            data_row_idx += 1
    
    def _fill_dynamic_rows(self, table: Any, data: List[List[Any]], skip_columns: Optional[List[int]], header_rows: int):
        if not data:
            return
        
        while len(table.rows) > header_rows:
            table._tbl.remove(table.rows[-1]._tr)
        
        num_columns = len(table.rows[0].cells) if table.rows else 0
        
        for data_row in data:
            new_row = table.add_row()
            
            if len(new_row.cells) < num_columns:
                for _ in range(num_columns - len(new_row.cells)):
                    new_row.add_cell()
            
            data_col_idx = 0
            for col_idx, cell in enumerate(new_row.cells):
                if skip_columns and col_idx in skip_columns:
                    continue
                
                if data_col_idx < len(data_row):
                    value = data_row[data_col_idx]
                    self._set_cell_value(cell, str(value) if value else '')
                    data_col_idx += 1
    
    def _set_cell_value(self, cell: Any, value: str):
        """设置单元格值"""
        set_cell_value(cell, value)
    
    def _insert_table_in_cell(self, paragraph, template_table):
        cell = self._find_parent_cell(paragraph)
        if cell is None:
            raise DocxTemplateError("Cannot find parent cell for paragraph")
        
        paragraph.clear()
        cell_element = cell._element
        new_table_element = deepcopy(template_table._element)
        p_element = paragraph._element
        p_index = list(cell_element).index(p_element)
        cell_element.remove(p_element)
        cell_element.insert(p_index, new_table_element)
        print(f"成功在单元格内插入表格")
    
    def _find_parent_cell(self, paragraph):
        try:
            p_element = paragraph._element
            current = p_element.getparent()
            
            while current is not None:
                if current.tag.endswith('tc'):
                    from docx.table import _Cell
                    return _Cell(current, None)
                current = current.getparent()
            
            return None
        except Exception:
            return None

class ImageInserter(ContentInserter):
    def insert(self, placeholder: str, image_paths: List[str], width: Optional[Any] = None, 
               height: Optional[Any] = None, alignment: Optional[str] = None, location: str = 'body'):
        self.validate_location(location, ['body', 'header', 'footer'])
        
        if not image_paths:
            raise DocxTemplateError(f"No image paths provided for placeholder '{placeholder}'")
        
        # 处理图片路径，自动添加 data_files 前缀
        processed_image_paths = []
        for img_path in image_paths:
            processed_path = self._resolve_image_path(img_path)
            if not os.path.exists(processed_path):
                raise DocxTemplateError(f"Image file not found: {processed_path} (original: {img_path})")
            processed_image_paths.append(processed_path)
        
        # 使用处理后的路径
        image_paths = processed_image_paths
        
        self._validate_image_dimensions(width, height)
        
        def create_image_paragraphs(doc, parent_element=None):
            paragraphs = []
            for idx, img_path in enumerate(image_paths):
                if parent_element is None:
                    new_p = doc.add_paragraph()
                else:
                    new_p = parent_element.add_paragraph()
                
                if alignment == 'center':
                    new_p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                elif alignment == 'right':
                    new_p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                elif alignment == 'left':
                    new_p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                
                run = new_p.add_run()
                try:
                    if width and height:
                        run.add_picture(img_path, width=width, height=height)
                    elif width:
                        run.add_picture(img_path, width=width)
                    elif height:
                        run.add_picture(img_path, height=height)
                    else:
                        run.add_picture(img_path, width=Inches(4.0))
                except (ValueError, TypeError, OSError) as e:
                    raise DocxTemplateError(f"Failed to insert image '{img_path}': {str(e)}")
                
                paragraphs.append(new_p)
            return paragraphs
        
        results = PlaceholderFinder.find_all_placeholders_in_location(self.doc, placeholder, location)
        
        # 如果在指定位置找不到，尝试在所有位置查找
        if not results:
            print(f"警告: 在 {location} 中未找到占位符 '{placeholder}'，尝试在所有位置查找...")
            for loc in ['header', 'body', 'footer']:
                if loc != location:
                    results = PlaceholderFinder.find_all_placeholders_in_location(self.doc, placeholder, loc)
                    if results:
                        print(f"在 {loc} 中找到占位符 '{placeholder}'")
                        location = loc
                        break
        
        if not results:
            print(f"警告: 占位符 '{placeholder}' 在所有位置都未找到，跳过图片插入操作")
            return
        
        for idx, paragraph in results:
            try:
                parent_element = self._get_parent_element(paragraph)
                self._replace_placeholder_with_images(paragraph, create_image_paragraphs, parent_element)
            except (AttributeError, ValueError, TypeError) as e:
                raise DocxTemplateError(f"Failed to replace placeholder '{placeholder}': {str(e)}")

    def _replace_placeholder_with_images(self, paragraph, create_fn, parent_element):
        p_element = paragraph._element
        p_parent = p_element.getparent()
        if p_parent is None:
            raise DocxTemplateError("Cannot find parent element of paragraph")
        
        index = list(p_parent).index(p_element)
        p_parent.remove(p_element)
        
        paragraphs = create_fn(self.doc, parent_element)
        for idx, new_p in enumerate(paragraphs):
            if idx == 0:
                p_parent.insert(index, new_p._element)
            else:
                if parent_element is None:
                    empty_p = self.doc.add_paragraph()
                else:
                    empty_p = parent_element.add_paragraph()
                p_parent.insert(index + idx * 2 - 1, empty_p._element)
                p_parent.insert(index + idx * 2, new_p._element)

    def _get_parent_element(self, paragraph):
        try:
            p_element = paragraph._element
            p_parent = p_element.getparent()
            if p_parent is None:
                return None
            for table in self.doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell._element == p_parent:
                            return cell
            return None
        except Exception:
            return None

    def _resolve_image_path(self, img_path: str) -> str:
        """解析图片路径，如果是相对路径则添加 data_files 前缀"""
        # 如果路径已经存在，直接返回
        if os.path.exists(img_path):
            return img_path
        
        # 如果是相对路径（以 ./ 或直接文件名开头），尝试添加 data_files 前缀
        if not os.path.isabs(img_path):
            # 移除开头的 ./
            clean_path = img_path.lstrip('./')
            clean_path = clean_path.lstrip('.\\')
            
            # 尝试在 data_files 目录中查找
            data_files_path = os.path.join('data_files', clean_path)
            if os.path.exists(data_files_path):
                return data_files_path
            
            # 尝试相对于当前工作目录的 data_files
            cwd_data_files_path = os.path.join(os.getcwd(), 'data_files', clean_path)
            if os.path.exists(cwd_data_files_path):
                return cwd_data_files_path
        
        # 如果都找不到，返回原路径（会在后续检查中报错）
        return img_path

    def _validate_image_dimensions(self, width, height):
        from docx.shared import Length
        valid_dimensions = []
        if width is not None:
            valid_dimensions.append(width)
        if height is not None:
            valid_dimensions.append(height)
        
        for dimension in valid_dimensions:
            if not isinstance(dimension, Length):
                raise DocxTemplateError(f"Invalid image dimension '{dimension}'. Must be a Length object (e.g., Inches, Mm, Cm, Pt)")

class CheckboxInserter(ContentInserter):
    """批量更新Word文档中的checkbox状态"""
    
    def insert(self, checkbox_mapping: Dict[str, bool]):
        """
        根据checkbox名称批量更新checkbox的勾选状态
        
        Args:
            checkbox_mapping: 字典，key为checkbox的name属性值，value为bool表示是否勾选
        """
        root = self.doc.part.element
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        
        checkboxes = root.findall('.//w:checkBox', namespaces=ns)
        updated = {}
        not_found = set(checkbox_mapping.keys())
        
        for checkbox in checkboxes:
            ffdata = checkbox.getparent()
            if ffdata is not None:
                name = ffdata.find('w:name', namespaces=ns)
                if name is not None:
                    field_name = name.get(w_ns + 'val')
                    
                    if field_name in checkbox_mapping:
                        should_check = checkbox_mapping[field_name]
                        not_found.discard(field_name)
                        
                        checked = checkbox.find('w:checked', namespaces=ns)
                        default = checkbox.find('w:default', namespaces=ns)
                        
                        if should_check:
                            if checked is None:
                                new_checked = parse_xml(f'<w:checked w:val="1" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                                checkbox.append(new_checked)
                            else:
                                checked.set(w_ns + 'val', '1')
                            if default is not None:
                                default.set(w_ns + 'val', '1')
                            updated[field_name] = True
                        else:
                            if checked is not None:
                                checkbox.remove(checked)
                            if default is not None:
                                default.set(w_ns + 'val', '0')
                            updated[field_name] = False
        
        # 打印更新结果
        if updated:
            print(f"成功更新 {len(updated)} 个checkbox:")
            for name, checked in updated.items():
                status = "勾选" if checked else "取消勾选"
                print(f"  {name}: {status}")
        
        # 警告未找到的checkbox
        if not_found:
            print(f"警告: 以下 {len(not_found)} 个checkbox在文档中未找到，已跳过:")
            for name in sorted(not_found):
                print(f"  - {name}")

class DocxTemplateProcessor:
    def __init__(self, template_path: str, output_path: str):
        if not os.path.exists(template_path):
            raise DocxTemplateError(f"Template file not found: {template_path}")
        
        self.template_path = template_path
        self.output_path = output_path
        self.operations = []
        # if file already exist and in use, throw exeption,don't read content 
        if DocxTemplateProcessor.is_word_file_open(output_path):
            raise DocxTemplateError(f"Output file is currently open in Word: {output_path}")
        shutil.copy(template_path, output_path)
        self.doc = Document(output_path)
    @staticmethod
    def is_word_file_open(file_path):
        directory = os.path.dirname(file_path)
        filename = os.path.basename(file_path)
    
        # Word 的锁文件规则：将文件名前两个字符替换为 ~$
        # 比如 test.docx -> ~$st.docx (注意：如果文件名很短，规则可能略有不同，但通常是 ~$ + 原名)
        # 标准做法是直接在原文件名前加 ~$ (对于由 Word 创建的临时文件)
    
        # Word 具体的临时文件名为：~$ + 文件名
        lock_filename = f"~${filename}"
        lock_file_path = os.path.join(directory, lock_filename)
    
        return os.path.exists(lock_file_path)
    def add_text(self, placeholder: str, value: str, location: str = 'body'):
        self.operations.append({
            'type': 'text',
            'placeholder': placeholder,
            'value': value,
            'location': location
        })
        return self
    
    def add_table(self, placeholder: str, table_template_path: str, 
                  raw_data: Optional[List[List[str]]] = None,
                  transformations: Optional[List[Dict]] = None,
                  metadata: Optional[Dict] = None,
                  targets_data: Optional[Dict] = None,
                  row_strategy: str = 'fixed_rows',
                  skip_columns: Optional[List[int]] = None,
                  header_rows: int = 1):
        self.operations.append({
            'type': 'table',
            'placeholder': placeholder,
            'table_template_path': table_template_path,
            'raw_data': raw_data,
            'transformations': transformations,
            'metadata': metadata,
            'targets_data': targets_data,
            'row_strategy': row_strategy,
            'skip_columns': skip_columns,
            'header_rows': header_rows
        })
        return self
    
    def add_image(self, placeholder: str, image_paths: List[str], 
                  width: Optional[any] = None, height: Optional[any] = None, 
                  alignment: Optional[str] = None, location: str = 'body'):
        self.operations.append({
            'type': 'image',
            'placeholder': placeholder,
            'image_paths': image_paths,
            'width': width,
            'height': height,
            'alignment': alignment,
            'location': location
        })
        return self
    
    def add_checkboxes(self, checkbox_mapping: Dict[str, bool]):
        """
        批量添加checkbox状态更新操作
        
        Args:
            checkbox_mapping: 字典，key为checkbox的name属性值，value为bool表示是否勾选
        """
        self.operations.append({
            'type': 'checkbox',
            'checkbox_mapping': checkbox_mapping
        })
        return self
    
    def get_all_placeholders(self) -> List[str]:
        """
        获取模板中所有的占位符名称
        
        Returns:
            List[str]: 占位符名称列表
        """
        import re
        pattern = re.compile(r'\{\{(\w+)\}\}')
        result = []
        
        for para in self.doc.paragraphs:
            matches = pattern.findall(para.text)
            result.extend(matches)
        
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        matches = pattern.findall(para.text)
                        result.extend(matches)
        
        for section in self.doc.sections:
            for header in [section.header, section.first_page_header, section.even_page_header]:
                if header:
                    for para in header.paragraphs:
                        matches = pattern.findall(para.text)
                        result.extend(matches)
                    for table in header.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for para in cell.paragraphs:
                                    matches = pattern.findall(para.text)
                                    result.extend(matches)
            
            for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
                if footer:
                    for para in footer.paragraphs:
                        matches = pattern.findall(para.text)
                        result.extend(matches)
                    for table in footer.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for para in cell.paragraphs:
                                    matches = pattern.findall(para.text)
                                    result.extend(matches)
        
        return list(dict.fromkeys(result))
    
    def process(self):
        try:
            for op in self.operations:
                if op['type'] == 'text':
                    inserter = TextInserter(self.doc)
                    inserter.insert(op['placeholder'], op['value'], op['location'])
                
                elif op['type'] == 'table':
                    inserter = TableInserter(self.doc)
                    inserter.insert(
                        op['placeholder'], 
                        op['table_template_path'],
                        op.get('raw_data'),
                        op.get('transformations'),
                        op.get('metadata'),
                        op.get('targets_data'),
                        op.get('row_strategy', 'fixed_rows'),
                        op.get('skip_columns'),
                        op.get('header_rows', 1),
                        'body'
                    )
                
                elif op['type'] == 'image':
                    inserter = ImageInserter(self.doc)
                    inserter.insert(op['placeholder'], op['image_paths'], 
                                  op['width'], op['height'], op['alignment'], op['location'])
                
                elif op['type'] == 'checkbox':
                    inserter = CheckboxInserter(self.doc)
                    inserter.insert(op['checkbox_mapping'])
            
            self.doc.save(self.output_path)
            print(f"文档已保存至: {self.output_path}")
            return self.output_path
        
        except DocxTemplateError as e:
            print(f"错误: {str(e)}")
            raise
        except Exception as e:
            print(f"处理文档时发生未知错误: {str(e)}")
            raise DocxTemplateError(f"Unknown error while processing document: {str(e)}")

