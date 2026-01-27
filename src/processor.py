from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import sys
import shutil
import os
from typing import List, Dict, Optional, Tuple, Any
from abc import ABC, abstractmethod

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
        for idx, paragraph in enumerate(container.paragraphs):
            if placeholder in paragraph.text:
                yield idx, paragraph
        for result in PlaceholderFinder._iterate_container(container):
            if isinstance(result, tuple) and len(result) == 2:
                idx, paragraph = result
            else:
                idx, paragraph = None, result
            if placeholder in paragraph.text:
                yield idx, paragraph

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
            
            index = list(p_parent).index(p_element)
            p_parent.remove(p_element)
            p_parent.insert(index, element)
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
        
        results = PlaceholderFinder.find_all_placeholders_in_location(self.doc, placeholder, location)
        if not results:
            raise PlaceholderNotFoundError(placeholder, location)
        
        replaced = False
        for idx, paragraph in results:
            if TextInserter._replace_in_paragraph(paragraph, placeholder, value):
                replaced = True
        
        if not replaced:
            raise PlaceholderNotFoundError(placeholder, location)
    
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
        if placeholder in paragraph.text:
            runs = paragraph.runs
            if runs:
                for run in runs:
                    if placeholder in run.text:
                        if '{{' + placeholder + '}}' in run.text:
                            run.text = run.text.replace('{{' + placeholder + '}}', value)
                        else:
                            run.text = run.text.replace(placeholder, value)
                        return True
        return False

class TableInserter(ContentInserter):
    def insert(self, placeholder: str, table_template_path: str, table_data: Optional[List[List[str]]] = None, offset_x: int = 0, offset_y: int = 0, location: str = 'body'):
        self.validate_location(location, ['body'])
        
        if not os.path.exists(table_template_path):
            raise DocxTemplateError(f"Table template file not found: {table_template_path}")
        
        table_template = Document(table_template_path)
        if not table_template.tables:
            raise DocxTemplateError(f"No tables found in template file: {table_template_path}")
        
        template_table = table_template.tables[0]
        
        if table_data:
            data_rows = len(table_data)
            data_cols = max(len(row) for row in table_data) if table_data else 0
            
            for row_idx, row in enumerate(template_table.rows):
                if row_idx < offset_y:
                    continue
                
                data_row_idx = row_idx - offset_y
                if data_row_idx >= data_rows:
                    break
                
                for cell_idx, cell in enumerate(row.cells):
                    if cell_idx < offset_x:
                        continue
                    
                    data_cell_idx = cell_idx - offset_x
                    if data_cell_idx >= len(table_data[data_row_idx]):
                        break
                    
                    cell_value = table_data[data_row_idx][data_cell_idx]
                    if cell_value and str(cell_value).strip():
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.text = str(cell_value)
                                break
                            if not cell.paragraphs or not any(p.runs for p in cell.paragraphs):
                                if cell.paragraphs:
                                    cell.paragraphs[0].add_run(str(cell_value))
                                else:
                                    cell.add_paragraph(str(cell_value))
                                if paragraph.runs:
                                    break
        
        results = PlaceholderFinder.find_all_placeholders_in_location(self.doc, placeholder, location)
        if not results:
            raise PlaceholderNotFoundError(placeholder, location)
        
        for idx, paragraph in results:
            try:
                if paragraph._element.getparent() is None:
                    continue
                PlaceholderFinder.replace_paragraph_with_element(paragraph, template_table._element)
            except (AttributeError, ValueError, TypeError) as e:
                raise DocxTemplateError(f"Failed to replace placeholder '{placeholder}': {str(e)}")

class ImageInserter(ContentInserter):
    def insert(self, placeholder: str, image_paths: List[str], width: Optional[Any] = None, 
               height: Optional[Any] = None, alignment: Optional[str] = None, location: str = 'body'):
        self.validate_location(location, ['body', 'header', 'footer'])
        
        if not image_paths:
            raise DocxTemplateError(f"No image paths provided for placeholder '{placeholder}'")
        
        for img_path in image_paths:
            if not os.path.exists(img_path):
                raise DocxTemplateError(f"Image file not found: {img_path}")
        
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
        if not results:
            raise PlaceholderNotFoundError(placeholder, location)
        
        for idx, paragraph in results:
            try:
                if paragraph._element.getparent() is None:
                    continue
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

class DocxTemplateProcessor:
    def __init__(self, template_path: str, output_path: str):
        if not os.path.exists(template_path):
            raise DocxTemplateError(f"Template file not found: {template_path}")
        
        self.template_path = template_path
        self.output_path = output_path
        self.operations = []
        
        shutil.copy(template_path, output_path)
        self.doc = Document(output_path)
    
    def add_text(self, placeholder: str, value: str, location: str = 'body'):
        self.operations.append({
            'type': 'text',
            'placeholder': placeholder,
            'value': value,
            'location': location
        })
        return self
    
    def add_table(self, placeholder: str, table_template_path: str, 
                  table_data: Optional[List[List[str]]] = None,
                  offset_x: int = 0, offset_y: int = 0):
        self.operations.append({
            'type': 'table',
            'placeholder': placeholder,
            'table_template_path': table_template_path,
            'table_data': table_data,
            'offset_x': offset_x,
            'offset_y': offset_y
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
                    inserter.insert(op['placeholder'], op['table_template_path'], 
                                  op['table_data'], 
                                  int(op.get('offset_x', 0)), 
                                  int(op.get('offset_y', 0)), 
                                  'body')
                
                elif op['type'] == 'image':
                    inserter = ImageInserter(self.doc)
                    inserter.insert(op['placeholder'], op['image_paths'], 
                                  op['width'], op['height'], op['alignment'], op['location'])
            
            self.doc.save(self.output_path)
            print(f"文档已保存至: {self.output_path}")
            return self.output_path
        
        except DocxTemplateError as e:
            print(f"错误: {str(e)}")
            raise
        except Exception as e:
            print(f"处理文档时发生未知错误: {str(e)}")
            raise DocxTemplateError(f"Unknown error while processing document: {str(e)}")

