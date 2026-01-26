"""
使用processor.py处理Word模板
"""
import sys
import json
from pathlib import Path

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
    parser.add_argument('--output', required=True, help='Path to output file')
    args = parser.parse_args()
    
    try:
        template_path = Path(args.template)
        operations_path = Path(args.operations)
        output_path = Path(args.output)
        
        if not template_path.exists():
            print(f"Error: Template file not found: {template_path}", file=sys.stderr)
            return 1
        
        if not operations_path.exists():
            print(f"Error: Operations file not found: {operations_path}", file=sys.stderr)
            return 1
        
        with open(operations_path, 'r', encoding='utf-8') as f:
            operations_data = json.load(f)
        
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
                processor.add_image(
                    op['placeholder'],
                    op['image_paths'],
                    op.get('width'),
                    op.get('height'),
                    op.get('alignment'),
                    op.get('location', 'body')
                )
                op_count += 1
            
            elif op['type'] == 'table':
                processor.add_table(
                    op['placeholder'],
                    op['table_template_path'],
                    op.get('table_data')
                )
                op_count += 1
        
        print(f"Executing {op_count} operations...", file=sys.stderr)
        result = processor.process()
        
        print(f"Report generated successfully: {result}")
        return 0
    
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}", file=sys.stderr)
        return 1
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in operations file - {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1

if __name__ == '__main__':
    sys.exit(main())
