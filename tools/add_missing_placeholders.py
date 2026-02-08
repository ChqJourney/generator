"""
添加缺失的占位符到配置文件和数据文件
"""
import json
import re
import sys
sys.path.insert(0, 'src')
from template_validator import TemplateValidator

def extract_all_placeholders(template_path):
    """从模板中提取所有占位符"""
    validator = TemplateValidator(template_path)
    all_placeholders = set()
    
    # 遍历文档中的所有段落和表格
    for para in validator.doc.paragraphs:
        matches = re.findall(r'\{\{(.*?)\}\}', para.text)
        for m in matches:
            all_placeholders.add(m)
    
    # 遍历表格
    for table in validator.doc.tables:
        for row in table.rows:
            for cell in row.cells:
                matches = re.findall(r'\{\{(.*?)\}\}', cell.text)
                for m in matches:
                    all_placeholders.add(m)
    
    # 遍历页眉页脚
    for section in validator.doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for para in header.paragraphs:
                    matches = re.findall(r'\{\{(.*?)\}\}', para.text)
                    for m in matches:
                        all_placeholders.add(m)
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for para in footer.paragraphs:
                    matches = re.findall(r'\{\{(.*?)\}\}', para.text)
                    for m in matches:
                        all_placeholders.add(m)
    
    return sorted(all_placeholders)

def load_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_json(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def main():
    # 提取模板中的所有占位符
    print("[Scan] Extracting placeholders from template...")
    all_placeholders = extract_all_placeholders('report_templates/production_template.docx')
    print(f"[Info] Found {len(all_placeholders)} placeholders")
    
    # 加载现有配置
    config = load_json('config/report_config.json')
    data = load_json('config/report_data.json')
    
    # 获取已配置的占位符
    configured_placeholders = {fm['template_field'] for fm in config['field_mappings']}
    print(f"[Info] Already configured {len(configured_placeholders)} placeholders")
    
    # 找出缺失的占位符
    missing = [p for p in all_placeholders if p not in configured_placeholders]
    print(f"[Info] Need to add {len(missing)} placeholders")
    
    # 添加到配置
    for placeholder in missing:
        field_mapping = {
            "template_field": placeholder,
            "source_field": f"calculated_data.{placeholder}",
            "type": "text"
        }
        config['field_mappings'].append(field_mapping)
        
        # 添加到数据文件的 calculated_data
        if 'calculated_data' not in data:
            data['calculated_data'] = {}
        data['calculated_data'][placeholder] = "not set"
        
        print(f"  + {placeholder}")
    
    # 保存更新后的文件
    save_json('config/report_config.json', config)
    save_json('config/report_data.json', data)
    
    print(f"\n[Done] Updated:")
    print(f"  - config/report_config.json: 添加了 {len(missing)} 个 field_mappings")
    print(f"  - config/report_data.json: 在 calculated_data 中添加了 {len(missing)} 个字段")

if __name__ == '__main__':
    main()
