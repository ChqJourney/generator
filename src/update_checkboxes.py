#!/usr/bin/env python3
import json
import sys
import argparse
from docx import Document
from docx.oxml import parse_xml

def load_checkbox_mapping(json_path):
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def update_checkboxes(docx_path, checkbox_mapping, output_path):
    doc = Document(docx_path)
    root = doc.part.element
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    
    checkboxes = root.findall('.//w:checkBox', namespaces=ns)
    updated = {}
    
    for checkbox in checkboxes:
        ffdata = checkbox.getparent()
        if ffdata is not None:
            name = ffdata.find('w:name', namespaces=ns)
            if name is not None:
                field_name = name.get(w_ns + 'val')
                
                if field_name in checkbox_mapping:
                    should_check = checkbox_mapping[field_name]
                    
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
    
    doc.save(output_path)
    return updated

def main():
    parser = argparse.ArgumentParser(description='Update Word document checkboxes based on JSON configuration')
    parser.add_argument('json_path', help='Path to JSON file containing checkbox mappings')
    parser.add_argument('docx_path', help='Path to input Word document')
    parser.add_argument('output_path', help='Path to output Word document')
    
    args = parser.parse_args()
    
    try:
        checkbox_mapping = load_checkbox_mapping(args.json_path)
        updated = update_checkboxes(args.docx_path, checkbox_mapping, args.output_path)
        
        print(f"成功更新 {len(updated)} 个checkbox:")
        for name, checked in updated.items():
            status = "勾选" if checked else "取消勾选"
            print(f"  {name}: {status}")
        print(f"已保存到: {args.output_path}")
        
    except FileNotFoundError as e:
        print(f"错误: 文件未找到 - {e}", file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"错误: JSON格式错误 - {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()