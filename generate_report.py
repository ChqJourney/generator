#!/usr/bin/env python3
"""
报告生成串联脚本
将完整的报告生成流程串联起来：
1. 验证 report_data
2. 验证 template
3. 计算 report data (calculator)
4. 生成 operations (field_mapper)
5. 生成最终报告 (process_template)

使用方法:
    # 基本用法（使用默认路径）
    python generate_report.py
    
    # 自定义路径
    python generate_report.py \
        --report config/report_data.json \
        --config config/report_config.json \
        --template report_templates/production_template.docx \
        --output output/final_report.docx
    
    # 严格模式（验证失败时停止）
    python generate_report.py --strict
    
    # 跳过验证步骤
    python generate_report.py --skip-validation
"""

import argparse
import subprocess
import sys
from pathlib import Path
from typing import List, Optional


def run_command(cmd: List[str], description: str, strict: bool = False) -> bool:
    """
    运行命令并处理输出
    
    Args:
        cmd: 命令列表
        description: 步骤描述
        strict: 严格模式（非零退出码时返回False）
    
    Returns:
        bool: 是否成功
    """
    print(f"\n{'='*70}")
    print(f"📋 {description}")
    print(f"{'='*70}")
    print(f"命令: {' '.join(cmd)}\n")
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=False,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode != 0:
            print(f"\n❌ {description} 失败 (退出码: {result.returncode})")
            if strict:
                return False
            # 非严格模式下询问是否继续
            response = input("\n是否继续? [y/N]: ").strip().lower()
            return response == 'y'
        
        print(f"\n✅ {description} 完成")
        return True
        
    except Exception as e:
        print(f"\n❌ {description} 出错: {e}")
        if strict:
            return False
        response = input("\n是否继续? [y/N]: ").strip().lower()
        return response == 'y'


def ensure_dir(path: Path) -> None:
    """确保目录存在"""
    path.parent.mkdir(parents=True, exist_ok=True)


def main():
    parser = argparse.ArgumentParser(
        description='报告生成串联脚本 - 一键完成从验证到生成的完整流程',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    # 使用默认配置
    python generate_report.py
    
    # 自定义输入输出
    python generate_report.py \
        --report config/my_report.json \
        --config config/my_config.json \
        --template report_templates/my_template.docx \
        --output output/my_report.docx
    
    # 严格模式（任何验证失败都停止）
    python generate_report.py --strict
    
    # 跳过所有验证
    python generate_report.py --skip-validation
    
    # 跳过模板验证（只验证数据）
    python generate_report.py --skip-template-validation
        """
    )
    
    # 输入文件路径
    parser.add_argument(
        '--report',
        default='config/report_data.json',
        help='输入的 report_data.json 路径 (默认: config/report_data.json)'
    )
    parser.add_argument(
        '--config',
        default='config/report_config.jsonc',
        help='report_config.json 路径 (默认: config/report_config.jsonc)'
    )
    parser.add_argument(
        '--template',
        default='report_templates/production_template.docx',
        help='Word 模板路径 (默认: report_templates/production_template.docx)'
    )
    
    # 中间文件输出路径
    parser.add_argument(
        '--calculated-report',
        default='output/calculated_report.json',
        help='计算的 report 输出路径 (默认: output/calculated_report.json)'
    )
    parser.add_argument(
        '--operations',
        default='output/operations.json',
        help='operations.json 输出路径 (默认: output/operations.json)'
    )
    
    # 最终输出
    parser.add_argument(
        '--output', '-o',
        default='output/final_report.docx',
        help='最终报告输出路径 (默认: output/final_report.docx)'
    )
    
    # 选项
    parser.add_argument(
        '--strict',
        action='store_true',
        help='严格模式：验证失败时停止流程'
    )
    parser.add_argument(
        '--strict-mode',
        action='store_true',
        help='Calculator 严格模式：字段缺失时报错'
    )
    parser.add_argument(
        '--skip-validation',
        action='store_true',
        help='跳过所有验证步骤'
    )
    parser.add_argument(
        '--skip-template-validation',
        action='store_true',
        help='跳过模板验证（只验证数据）'
    )
    parser.add_argument(
        '--functions-module',
        help='自定义计算函数模块路径',
        default='custom_calculations'
    )
    
    args = parser.parse_args()
    
    # 转换路径为 Path 对象
    report_path = Path(args.report)
    config_path = Path(args.config)
    template_path = Path(args.template)
    calculated_report_path = Path(args.calculated_report)
    operations_path = Path(args.operations)
    output_path = Path(args.output)
    
    # 检查输入文件是否存在
    if not report_path.exists():
        print(f"❌ 错误: Report 文件不存在: {report_path}")
        return 1
    
    if not config_path.exists():
        print(f"❌ 错误: Config 文件不存在: {config_path}")
        return 1
    
    if not template_path.exists():
        print(f"❌ 错误: Template 文件不存在: {template_path}")
        return 1
    
    # 确保输出目录存在
    ensure_dir(calculated_report_path)
    ensure_dir(operations_path)
    ensure_dir(output_path)
    
    print(f"\n{'#'*70}")
    print("#" + " "*68 + "#")
    print("#" + " 报告生成流程".center(64) + "#")
    print("#" + " "*68 + "#")
    print(f"{'#'*70}")
    
    print("\n📁 输入文件:")
    print(f"  • Report Data: {report_path}")
    print(f"  • Config: {config_path}")
    print(f"  • Template: {template_path}")
    
    print("\n📁 输出文件:")
    print(f"  • Calculated Report: {calculated_report_path}")
    print(f"  • Operations: {operations_path}")
    print(f"  • Final Report: {output_path}")
    
    # ========== 第1步: 验证 Report Data ==========
    if not args.skip_validation:
        cmd = [
            sys.executable, 'src/report_data_validator.py',
            '--report', str(report_path),
            '--config', str(config_path)
        ]
        if args.strict:
            cmd.append('--strict')
        
        if not run_command(cmd, "步骤 1/5: 验证 Report Data", strict=args.strict):
            print("\n" + "="*70)
            print("❌ 流程中止: Report Data 验证失败")
            print("="*70)
            return 1
    else:
        print("\n" + "="*70)
        print("⏭️  跳过步骤 1/5: Report Data 验证 (--skip-validation)")
        print("="*70)
    
    # ========== 第2步: 验证 Template ==========
    if not args.skip_validation and not args.skip_template_validation:
        cmd = [
            sys.executable, 'src/template_validator.py',
            '--template', str(template_path),
            '--config', str(config_path)
        ]
        
        if not run_command(cmd, "步骤 2/5: 验证 Template", strict=args.strict):
            print("\n" + "="*70)
            print("❌ 流程中止: Template 验证失败")
            print("   提示: 使用 --skip-template-validation 跳过此步骤")
            print("="*70)
            return 1
    else:
        print("\n" + "="*70)
        if args.skip_validation:
            print("⏭️  跳过步骤 2/5: Template 验证 (--skip-validation)")
        else:
            print("⏭️  跳过步骤 2/5: Template 验证 (--skip-template-validation)")
        print("="*70)
    
    # ========== 第3步: 计算 Report Data ==========
    cmd = [
        sys.executable, 'src/calculator.py',
        '--config', str(config_path),
        '--report', str(report_path),
        '--output', str(calculated_report_path)
    ]
    if args.strict_mode:
        cmd.append('--strict-mode')
    if args.functions_module:
        cmd.extend(['--functions-module', args.functions_module])
    
    if not run_command(cmd, "步骤 3/5: 计算 Report Data", strict=True):
        print("\n" + "="*70)
        print("❌ 流程中止: Report Data 计算失败")
        print("="*70)
        return 1
    
    # ========== 第4步: 生成 Operations ==========
    cmd = [
        sys.executable, 'src/field_mapper.py',
        '--config', str(config_path),
        '--report', str(calculated_report_path),
        '--output', str(operations_path)
    ]
    
    if not run_command(cmd, "步骤 4/5: 生成 Operations", strict=True):
        print("\n" + "="*70)
        print("❌ 流程中止: Operations 生成失败")
        print("="*70)
        return 1
    
    # ========== 第5步: 生成最终报告 ==========
    cmd = [
        sys.executable, 'src/process_template.py',
        '--template', str(template_path),
        '--operations', str(operations_path),
        '--calculated-report', str(calculated_report_path),
        '--output', str(output_path)
    ]
    
    if not run_command(cmd, "步骤 5/5: 生成最终报告", strict=True):
        print("\n" + "="*70)
        print("❌ 流程中止: 报告生成失败")
        print("="*70)
        return 1
    
    # 完成
    print("\n" + "="*70)
    print("✅ 报告生成流程完成!")
    print("="*70)
    print(f"\n📄 最终报告: {output_path}")
    
    # 显示中间文件
    print("\n📝 生成的中间文件:")
    print(f"  • {calculated_report_path}")
    print(f"  • {operations_path}")
    
    return 0


if __name__ == '__main__':
    sys.exit(main())
