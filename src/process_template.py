"""
使用processor.py处理Word模板
"""
import sys
import json
from pathlib import Path
from docx.shared import Inches

from processor import DocxTemplateError

sys.path.insert(0, str(Path(__file__).parent.parent / 'scripts'))

try:
    from processor import DocxTemplateProcessor
except ImportError:
    print("Error: processor.py not found. Please ensure it is in the scripts directory.", file=sys.stderr)
    sys.exit(1)

def main():
    import argparse
    parser = argparse.ArgumentParser(description='Process Word template using processor.py')
    parser.add_argument('--template', required=True, help='Path to Word template file')
    parser.add_argument('--operations', required=True, help='Path to operations.json')
    parser.add_argument('--metadata', required=True, help='Path to metadata.json')
    parser.add_argument('--targets', required=True, help='Path to targets.json')
    parser.add_argument('--output', required=True, help='Path to output file')
    args = parser.parse_args()
    
    try:
        template_path = Path(args.template)
        operations_path = Path(args.operations)
        output_path = Path(args.output)
        metadata_path = Path(args.metadata)
        targets_path = Path(args.targets)

        if not template_path.exists():
            print(f"Error: Template file not found: {template_path}", file=sys.stderr)
            return 1
        
        if not operations_path.exists():
            print(f"Error: Operations file not found: {operations_path}", file=sys.stderr)
            return 1
        
        if not metadata_path.exists():
            print(f"Error: Metadata file not found: {metadata_path}", file=sys.stderr)
            return 1
        if not targets_path.exists():
            print(f"Error: Targets file not found: {targets_path}", file=sys.stderr)
            return 1
        
        with open(operations_path, 'r', encoding='utf-8') as f:
            operations_data = json.load(f)
        
        with open(metadata_path, 'r', encoding='utf-8') as f:
            metadata_data = json.load(f)
        with open(targets_path, 'r', encoding='utf-8') as f:
            targets_data = json.load(f)

        processor = DocxTemplateProcessor(str(template_path), str(output_path))
        
        op_count = 0
        for op in operations_data.get('operations', []):
            if op['type'] == 'text':
                processor.add_text(
                    op['placeholder'],
                    op['value'],
                    op.get('location', 'body')
                )
                op_count += 1
            
            elif op['type'] == 'image':
                width = op.get('width')
                height = op.get('height')
                
                if width is not None:
                    if isinstance(width, (int, float)):
                        width = Inches(width)
                
                if height is not None:
                    if isinstance(height, (int, float)):
                        height = Inches(height)
                
                processor.add_image(
                    op['placeholder'],
                    op['image_paths'],
                    width,
                    height,
                    op.get('alignment'),
                    op.get('location', 'body')
                )
                op_count += 1
            
            elif op['type'] == 'table':
                processor.add_table(
                    op['placeholder'],
                    op['table_template_path'],
                    op.get('table_data'),
                    op.get('transformations'),
                    metadata_data,
                    targets_data,
                    op.get('row_strategy', 'fixed_rows'),
                    op.get('skip_columns'),
                    op.get('header_rows', 1)
                )
                op_count += 1
        
        print(f"Executing {op_count} operations...", file=sys.stderr)
        result = processor.process()
        
        print(f"Report generated successfully: {result}")
        # auto open the output file in windows
        if sys.platform == "win32":
            import os
            os.startfile(output_path)

        return 0
    
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}", file=sys.stderr)
        return 1
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in operations file - {e}", file=sys.stderr)
        return 1
    except DocxTemplateError as e:
        print(f"DocxTemplateError: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1

if __name__ == '__main__':
    sys.exit(main())
