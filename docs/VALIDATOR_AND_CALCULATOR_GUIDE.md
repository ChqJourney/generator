# Report Data 验证器与计算器使用指南

本文档详细介绍如何使用 `ReportDataValidator` 验证数据，以及如何配置 `Calculator` 自动计算字段。

---

## 📑 目录

1. [ReportDataValidator 使用指南](#1-reportdatavalidator-使用指南)
   - [命令行使用](#命令行使用)
   - [在代码中使用](#在代码中使用)
   - [验证结果解读](#验证结果解读)
2. [Calculator 配置指南](#2-calculator-配置指南)
   - [配置结构](#配置结构)
   - [计算函数配置](#计算函数配置)
   - [内置函数列表](#内置函数列表)
   - [完整配置示例](#完整配置示例)
3. [工作流程](#3-工作流程)
4. [常见问题](#4-常见问题)

---

## 1. ReportDataValidator 使用指南

### 命令行使用

#### 基本验证（仅检查数据格式）

```bash
python src/report_data_validator.py \
    --report config/report_data_example.json
```

#### 完整验证（包括配置一致性检查）

```bash
python src/report_data_validator.py \
    --report config/report_data_example.json \
    --config config/report_config.jsonc \
    --base-path .
```

#### 严格模式（有警告时返回错误码）

```bash
python src/report_data_validator.py \
    --report config/report_data_example.json \
    --config config/report_config.jsonc \
    --strict
```

#### 参数说明

| 参数 | 必需 | 说明 |
|------|------|------|
| `--report` | ✅ | 报告数据 JSON 文件路径 |
| `--config` | ❌ | 配置文件路径（用于检查配置一致性） |
| `--base-path` | ❌ | 基础路径（用于验证图像文件是否存在） |
| `--strict` | ❌ | 严格模式，有警告时返回非零退出码 |

### 在代码中使用

#### 基础用法

```python
from src.report_data_validator import ReportDataValidator
import json
from pathlib import Path

# 加载数据
with open('config/report_data_example.json', 'r', encoding='utf-8') as f:
    report_data = json.load(f)

# 创建验证器
validator = ReportDataValidator(
    report_data=report_data,
    base_path=Path('.')
)

# 执行验证
report = validator.validate()

# 打印报告
report.print_report()

# 检查结果
if report.is_valid:
    print("✅ 验证通过！")
else:
    print(f"❌ 验证失败，有 {len(report.errors)} 个错误")
```

#### 带配置文件的完整验证

```python
from src.report_data_validator import ReportDataValidator
import json
from pathlib import Path

# 加载数据
with open('config/report_data_example.json', 'r', encoding='utf-8') as f:
    report_data = json.load(f)

with open('config/report_config.jsonc', 'r', encoding='utf-8') as f:
    config_data = json.load(f)

# 创建验证器
validator = ReportDataValidator(
    report_data=report_data,
    config_data=config_data,  # 传入配置以检查一致性
    base_path=Path('.')
)

# 执行验证
report = validator.validate()

# 获取详细信息
print(f"错误数: {len(report.errors)}")
print(f"警告数: {len(report.warnings)}")
print(f"信息数: {len(report.infos)}")

# 获取可用字段
available = validator.get_available_fields()
print(f"metadata 字段: {available.get('metadata', [])}")
print(f"extracted_data 字段: {available.get('extracted_data', [])}")

# 获取缺少配置映射的字段
missing_mappings = validator.get_missing_config_mappings()
print(f"缺少配置映射的字段: {missing_mappings}")
```

#### 处理验证结果

```python
# 遍历错误
for error in report.errors:
    print(f"[错误] {error.path}: {error.message}")
    print(f"  建议: {error.suggestion}")

# 遍历警告
for warning in report.warnings:
    print(f"[警告] {warning.path}: {warning.message}")

# 访问摘要信息
for key, value in report.summary.items():
    print(f"{key}: {value}")
```

### 验证结果解读

#### 验证级别

| 级别 | 说明 | 影响 |
|------|------|------|
| **ERROR** | 错误，数据格式或内容有问题 | 会阻止处理，必须修复 |
| **WARNING** | 警告，建议修复但可以继续 | 不影响处理，但建议修复 |
| **INFO** | 信息，仅供参考 | 不影响处理 |

#### 常见错误类型

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| `缺少必需的顶级字段` | 缺少 `metadata`/`extracted_data`/`calculated_data` | 添加缺失的顶级字段 |
| `路径引用的字段不存在` | 配置引用的字段在数据中不存在 | 在数据中添加该字段，或修改配置 |
| `表格数据必须是列表` | 表格数据格式错误 | 确保表格数据是列表的列表 |
| `列数与表头不一致` | 表格行列数不匹配 | 确保所有数据行列数与表头一致 |
| `图像文件不存在` | 图像路径错误或文件缺失 | 确认图像路径正确 |

#### 验证报告示例

```
======================================================================
📋 Report Data 完整验证报告
======================================================================

📊 数据摘要:
  • 报告编号: RPT-2024-001
  • 申请人: ABC Lighting Co., Ltd.
  • metadata 字段数: 20
  • extracted_data 字段数: 8
  • calculated_data 字段数: 9

⚠️  [警告] 发现 3 个警告（建议修复）:
  1. [extracted_data.images[0]] 图像文件不存在
     💡 建议: 请确认文件路径正确

======================================================================
⚠️  [通过但有警告] 格式基本正确，但建议修复上述警告。
======================================================================
```

---

## 2. Calculator 配置指南

### 配置结构

#### 方式 1: 预填充值（直接从数据读取）

```json
{
  "template_field": "energy_class",
  "source_field": "calculated_data.energy_class",
  "type": "text"
}
```

**适用场景**: `calculated_data` 中已预先填充了计算好的值。

#### 方式 2: 自动计算（推荐）

```json
{
  "template_field": "energy_class",
  "source_field": "calculated_data.energy_class",
  "type": "text",
  "function": "calculate_energy_class_rating",
  "args": ["extracted_data.rated_wattage", "extracted_data.useful_luminous_flux"]
}
```

**适用场景**: 需要从 `extracted_data` 或 `metadata` 中的原始数据计算得出。

#### 配置字段说明

| 字段 | 必需 | 说明 |
|------|------|------|
| `template_field` | ✅ | Word 模板中的占位符名称 |
| `source_field` | ✅ | 数据源字段路径（支持点号路径） |
| `type` | ✅ | 字段类型: `text`, `table`, `image` |
| `function` | ❌ | 计算函数名称（需要计算时填写） |
| `args` | ❌ | 计算函数参数列表（函数必需时填写） |

### 计算函数配置

#### 基本结构

```json
{
  "template_field": "目标占位符",
  "source_field": "calculated_data.输出字段",
  "type": "text",
  "function": "函数名称",
  "args": ["输入参数1", "输入参数2", ...]
}
```

#### Args 参数格式

`args` 中的每个参数都是**点号路径**，指向数据源中的字段：

```json
{
  "args": [
    "extracted_data.rated_wattage",      // 从 extracted_data 获取
    "extracted_data.useful_luminous_flux",
    "metadata.report_no"                  // 从 metadata 获取
  ]
}
```

#### 计算执行流程

```
1. 读取 args 中每个路径的值
   ↓
2. 调用 function 指定的函数，传入参数值
   ↓
3. 将返回值写入 source_field 指定的位置
   ↓
4. 生成 calculated_report.json
```

### 内置函数列表

#### 能源相关

| 函数名 | 功能 | 参数 | 返回值示例 |
|--------|------|------|-----------|
| `calculate_energy_class_rating` | 计算能源等级 | `(wattage, flux)` | `"A++"` ~ `"E"` |
| `calculate_energy_efficacy` | 计算能源效率 | `(wattage, flux)` | `"100.00"` |

**使用示例**:
```json
{
  "template_field": "efficacy",
  "source_field": "calculated_data.efficacy",
  "type": "text",
  "function": "calculate_energy_efficacy",
  "args": ["extracted_data.rated_wattage", "extracted_data.useful_luminous_flux"]
}
```

**计算逻辑**:
```
efficacy = flux / wattage
例如: 1050 / 10.5 = 100.00 lm/W
```

#### 数学计算

| 函数名 | 功能 | 参数 | 返回值示例 |
|--------|------|------|-----------|
| `multiply` | 乘法 | `(a, b)` | `21.0` |
| `divide` | 除法 | `(a, b, default=0)` | `2.5` |
| `calculate_percentage` | 百分比 | `(value, total)` | `"50.00%"` |
| `format_number` | 格式化数字 | `(value, decimal_places)` | `"100.00"` |

**使用示例**:
```json
{
  "template_field": "percentage",
  "source_field": "calculated_data.percentage",
  "type": "text",
  "function": "calculate_percentage",
  "args": ["extracted_data.partial_value", "extracted_data.total_value"]
}
```

#### 字符串处理

| 函数名 | 功能 | 参数 | 返回值示例 |
|--------|------|------|-----------|
| `concat` | 字符串拼接 | `(*args, separator=" ")` | `"A B C"` |

**使用示例**:
```json
{
  "template_field": "ratings",
  "source_field": "calculated_data.ratings",
  "type": "text",
  "function": "concat",
  "args": ["metadata.mains_or_not", "metadata.connected_or_not"],
  "separator": ", "
}
```

### 完整配置示例

#### report_config.json（完整版）

```json
{
  "template_path": "report_templates/report_template1.docx",
  "output_dir": "output/",
  "field_mappings": [
    // ==================== metadata 字段 ====================
    {
      "template_field": "report_no",
      "source_field": "metadata.report_no",
      "type": "text"
    },
    {
      "template_field": "issue_date",
      "source_field": "metadata.issue_date",
      "type": "text"
    },
    {
      "template_field": "applicant_name",
      "source_field": "metadata.applicant_name",
      "type": "text"
    },
    {
      "template_field": "applicant_address",
      "source_field": "metadata.applicant_address",
      "type": "text"
    },
    {
      "template_field": "product_name",
      "source_field": "metadata.product_name",
      "type": "text"
    },
    {
      "template_field": "manufacturer",
      "source_field": "metadata.manufacturer",
      "type": "text"
    },
    {
      "template_field": "test_period",
      "source_field": "metadata.test_period",
      "type": "text"
    },
    
    // ==================== extracted_data 字段 ====================
    {
      "template_field": "model_identifier",
      "source_field": "extracted_data.model_identifier",
      "type": "text"
    },
    {
      "template_field": "rated_wattage",
      "source_field": "extracted_data.rated_wattage",
      "type": "text"
    },
    {
      "template_field": "useful_luminous_flux",
      "source_field": "extracted_data.useful_luminous_flux",
      "type": "text"
    },
    {
      "template_field": "cct",
      "source_field": "extracted_data.cct",
      "type": "text"
    },
    {
      "template_field": "standby_power",
      "source_field": "extracted_data.standby_power",
      "type": "text"
    },
    
    // ==================== 表格数据 ====================
    {
      "template_field": "photometric_data",
      "source_field": "extracted_data.photometric_data",
      "table_template_path": "report_templates/tables/photometric_table_template.docx",
      "type": "table",
      "row_strategy": "fixed_rows",
      "header_rows": 2,
      "skip_columns": [1],
      "transformations": [
        {
          "type": "skip_columns",
          "columns": [1]
        },
        {
          "type": "custom_transform",
          "transformer": "photometric_data_transformer",
          "calculate_columns": [5],
          "formulas": {
            "5": "E{row}/C{row}"
          },
          "average_columns": [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
          "format_rules": {
            "2": [{"condition": "x < 100", "format": "{:.2f}"}],
            "4": [{"condition": "x >= 100", "format": "{:.1f}"}, {"condition": "x < 100", "format": "{:.2f}"}],
            "5": [{"condition": "x >= 100", "format": "{:.1f}"}, {"condition": "x < 100", "format": "{:.2f}"}]
          }
        }
      ]
    },
    
    // ==================== 图像 ====================
    {
      "template_field": "images",
      "source_field": "extracted_data.images",
      "type": "image",
      "width": 4.0,
      "alignment": "center"
    },
    
    // ==================== calculated_data 字段（自动计算）====================
    {
      "template_field": "efficacy",
      "source_field": "calculated_data.efficacy",
      "type": "text",
      "function": "calculate_energy_efficacy",
      "args": ["extracted_data.rated_wattage", "extracted_data.useful_luminous_flux"]
    },
    {
      "template_field": "energy_class",
      "source_field": "calculated_data.energy_class",
      "type": "text",
      "function": "calculate_energy_class_rating",
      "args": ["extracted_data.rated_wattage", "extracted_data.useful_luminous_flux"]
    },
    {
      "template_field": "sample_size",
      "source_field": "calculated_data.sample_size",
      "type": "text"
    },
    {
      "template_field": "PonMax",
      "source_field": "calculated_data.PonMax",
      "type": "text"
    }
  ],
  "merge_strategy": "prefer_first"
}
```

#### 对应的输入数据 (report.json)

```json
{
  "metadata": {
    "report_no": "RPT-2024-001",
    "issue_date": "2024-03-15",
    "applicant_name": "ABC Lighting Co., Ltd.",
    "applicant_address": "123 Industrial Road, Shenzhen, China",
    "product_name": "LED Downlight Series A",
    "manufacturer": "ABC Lighting Co., Ltd.",
    "test_period": "2024-03-01 to 2024-03-10"
  },
  "extracted_data": {
    "model_identifier": "DL-A-10W-4000K",
    "rated_wattage": "10.5",
    "useful_luminous_flux": "1050",
    "cct": "4000K",
    "standby_power": "0.344W",
    "images": [
      "data_files/images/product_front.jpg",
      "data_files/images/product_side.jpg",
      "data_files/images/product_back.jpg"
    ],
    "photometric_data": {
      "type": "table",
      "value": [
        ["Current(A)", "Power Pon(W)", "Luminous Flux(lm)", "Efficacy(lm/W)"],
        ["0.10", "10.5", "1050", "100.0"]
      ]
    }
  },
  "calculated_data": {
    "sample_size": "5 units",
    "PonMax": "11.5W"
  }
}
```

#### Calculator 输出 (calculated_report.json)

```json
{
  "metadata": { ... },
  "extracted_data": { ... },
  "calculated_data": {
    "sample_size": "5 units",        // 预填充值
    "PonMax": "11.5W",               // 预填充值
    "efficacy": "100.00",            // 计算生成
    "energy_class": "A+"             // 计算生成
  }
}
```

---

## 3. 工作流程

### 完整处理流程

```
┌─────────────────────────────────────────────────────────────────┐
│  Step 1: 验证输入数据                                            │
│  python src/report_data_validator.py                            │
│    --report config/report.json                                  │
│    --config config/report_config.jsonc                           │
└─────────────────────────────────────────────────────────────────┘
                              ↓
                              ↓ 验证通过 ✅
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│  Step 2: 执行计算（如果配置了 function）                          │
│  python src/calculator.py                                       │
│    --config config/report_config.jsonc                           │
│    --report config/report.json                                  │
│    --output output/calculated_report.json                       │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│  Step 3: 生成操作队列                                            │
│  python src/field_mapper.py                                     │
│    --config config/report_config.jsonc                           │
│    --report output/calculated_report.json                       │
│    --output output/operations.json                              │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│  Step 4: 处理模板生成报告                                        │
│  python src/process_template.py                                 │
│    --template report_templates/report_template1.docx            │
│    --operations output/operations.json                          │
│    --output output/final_report.docx                            │
└─────────────────────────────────────────────────────────────────┘
```

### 混合使用模式

`calculated_data` 可以同时包含预填充值和计算值：

```json
// 配置中
{
  "field_mappings": [
    // 预填充值（数据中直接提供）
    {
      "template_field": "sample_size",
      "source_field": "calculated_data.sample_size",
      "type": "text"
    },
    // 自动计算值
    {
      "template_field": "efficacy",
      "source_field": "calculated_data.efficacy",
      "type": "text",
      "function": "calculate_energy_efficacy",
      "args": ["extracted_data.rated_wattage", "extracted_data.useful_luminous_flux"]
    }
  ]
}
```

```json
// 输入数据
{
  "calculated_data": {
    "sample_size": "5 units"    // 预填充
    // efficacy 将由 calculator 计算生成
  }
}
```

---

## 4. 常见问题

### Q1: 如何添加自定义计算函数？

**A:** 创建自定义计算模块：

```python
# src/custom_calculations.py
from src.calculator import CalculationRegistry

@CalculationRegistry.register("my_custom_calculation")
def my_custom_calculation(value1, value2):
    """自定义计算函数"""
    result = float(value1) * float(value2) * 1.5
    return f"{result:.2f}"
```

在配置中使用：

```json
{
  "template_field": "custom_field",
  "source_field": "calculated_data.custom_field",
  "type": "text",
  "function": "my_custom_calculation",
  "args": ["extracted_data.value1", "extracted_data.value2"]
}
```

运行 Calculator 时加载自定义模块：

```bash
python src/calculator.py \
    --config config/report_config.jsonc \
    --report config/report.json \
    --output output/calculated_report.json \
    --functions-module custom_calculations
```

### Q2: 验证器报告"calculated_data 字段可能由计算生成"是什么意思？

**A:** 这是信息提示（INFO），不是错误。表示配置引用了 `calculated_data` 中的字段，但输入数据中该字段为空。如果该字段确实由 `Calculator` 计算生成，可以忽略此提示。

### Q3: 如何处理可选字段？

**A:** 如果某个计算字段可能没有值，有两种处理方式：

**方式 1**: 在报告中显示空值（默认行为）
```json
{
  "template_field": "optional_field",
  "source_field": "calculated_data.optional_field",
  "type": "text",
  "function": "some_calculation",
  "args": ["extracted_data.input1"]
}
```

**方式 2**: 条件计算（需要自定义函数）
```python
@CalculationRegistry.register("safe_calculation")
def safe_calculation(value):
    if value is None or value == "":
        return "N/A"
    return float(value) * 2
```

### Q4: 计算函数的参数可以是常量吗？

**A:** 目前 `args` 只支持字段路径，不支持直接常量。如果需要常量，可以：

1. 将常量放在 `metadata` 中
2. 在自定义函数中硬编码常量

```python
@CalculationRegistry.register("multiply_by_factor")
def multiply_by_factor(value):
    FACTOR = 1.1  # 硬编码常量
    return float(value) * FACTOR
```

## 5. Template Validator 使用指南

`Template Validator` 用于验证 Word 模板中的占位符状态，特别是检查占位符是否被分割到多个 runs 中（这是导致 `{{}}` 残留的根本原因）。

### 命令行使用

#### 基本用法（只检查占位符是否被分割）

```bash
python src/template_validator.py \
    --template report_templates/production_template.docx
```

#### 完整用法（同时检查占位符是否在 config 中有定义）

```bash
python src/template_validator.py \
    --template report_templates/production_template.docx \
    --config config/report_config.jsonc
```

#### 参数说明

| 参数 | 必需 | 说明 |
|------|------|------|
| `--template` | ✅ | Word 模板文件路径（.docx） |
| `--config` | ❌ | 配置文件路径（用于验证占位符是否在配置中有定义） |

### 输出解读

#### 占位符统计

验证器会输出三类占位符：

1. **单个 run 中的占位符**: 格式正确，可以被正常替换 ✅
2. **被分割到多个 runs 的占位符**: ⚠️ **需要修复**，这是导致 `{{}}` 残留的原因
3. **在配置中未定义的占位符**: 需要在 config 中添加配置

#### 被分割的占位符示例

```
占位符: {{efficacy}}
位置: table 0, row 7, cell 1
Runs 详情:
  Run 0: '{{'
  Run 1: 'efficacy'
  Run 2: '}}'
```

这种情况说明 `{{efficacy}}` 被 Word 分割成了 3 个 runs，导致替换后 `{{}}` 无法被清除。

### 修复方法

#### 1. 修复被分割的占位符

在 Word 中重新输入占位符：

1. 删除原有的占位符
2. 重新输入 `{{placeholder_name}}`
3. **确保不要复制粘贴**，因为复制可能带来格式问题
4. 建议直接在 Word 中手动输入

#### 2. 添加缺失的占位符配置

在 `config/report_config.json` 中添加：

```json
{
  "field_mappings": [
    {
      "template_field": "placeholder_name",
      "source_field": "path.to.data",
      "type": "text"
    }
  ]
}
```

### 完整验证工作流

推荐的报告生成前验证流程：

```bash
# 步骤 1: 验证 report.json 数据格式和内容
python src/report_data_validator.py \
    --report config/report.json \
    --config config/report_config.jsonc

# 步骤 2: 验证 Word 模板（检查占位符是否会被正确处理）
python src/template_validator.py \
    --template report_templates/production_template.docx \
    --config config/report_config.jsonc

# 如果以上两个验证都通过，继续生成报告...

# 步骤 3: 执行计算（如果需要）
python src/calculator.py \
    --config config/report_config.jsonc \
    --report config/report.json \
    --output output/calculated_report.json

# 步骤 4: 生成操作
python src/field_mapper.py \
    --config config/report_config.jsonc \
    --report output/calculated_report.json \
    --output output/operations.json

# 步骤 5: 处理模板生成报告
python src/process_template.py \
    --template report_templates/production_template.docx \
    --operations output/operations.json \
    --output output/final_report.docx
```

---

## 6. 综合示例

### 完整的工作流程示例

#### 1. 准备 report.json

```json
{
  "metadata": {
    "report_no": "RPT-2024-001",
    "issue_date": "2024-03-15",
    "applicant_name": "ABC Lighting Co., Ltd.",
    "product_name": "LED Downlight Series A"
  },
  "extracted_data": {
    "model_identifier": "DL-A-10W-4000K",
    "rated_wattage": "10.5",
    "useful_luminous_flux": "1050",
    "photometric_data": {
      "type": "table",
      "value": [
        ["Current(A)", "Power Pon(W)", "Luminous Flux(lm)", "Efficacy(lm/W)"],
        ["0.10", "10.5", "1050", "100.0"]
      ]
    }
  },
  "calculated_data": {}
}
```

#### 2. 验证数据

```bash
python src/report_data_validator.py \
    --report config/report.json \
    --config config/report_config.jsonc
```

**预期输出**：
```
✅ [通过] 所有检查通过！格式正确。
```

#### 3. 验证模板

```bash
python src/template_validator.py \
    --template report_templates/production_template.docx \
    --config config/report_config.jsonc
```

**预期输出**：
```
✅ 模板验证通过！所有占位符都在单个 run 中，并且都已在配置中定义。
```

#### 4. 执行计算

```bash
python src/calculator.py \
    --config config/report_config.jsonc \
    --report config/report.json \
    --output output/calculated_report.json
```

#### 5. 生成报告

```bash
python src/field_mapper.py \
    --config config/report_config.jsonc \
    --report output/calculated_report.json \
    --output output/operations.json

python src/process_template.py \
    --template report_templates/production_template.docx \
    --operations output/operations.json \
    --output output/final_report.docx
```

---

## 7. 常见问题 (FAQ)

### Q1: 如何添加自定义计算函数？

**A:** 参考上文"自定义计算函数"部分。

### Q2: 验证器报告"calculated_data 字段可能由计算生成"是什么意思？

**A:** 这是信息提示（INFO），不是错误。表示配置引用了 `calculated_data` 中的字段，但输入数据中该字段为空。如果该字段确实由 `Calculator` 计算生成，可以忽略此提示。

### Q3: 如何处理可选字段？

**A:** 如果某个计算字段可能没有值，可以使用条件计算函数（见上文"条件计算函数"示例）。

### Q4: 计算函数的参数可以是常量吗？

**A:** 目前 `args` 只支持字段路径，不支持直接常量。如果需要常量，可以将常量放在 `metadata` 中，或在自定义函数中硬编码常量。

### Q5: Template Validator 报告"占位符被分割"怎么办？

**A:** 这是导致 `{{}}` 残留的根本原因。修复方法：

1. 在 Word 中定位到该占位符
2. 删除整个占位符（包括 `{{}}`）
3. **手动重新输入** `{{placeholder_name}}`（不要复制粘贴）
4. 保存模板后重新运行 Template Validator 验证

### Q6: 如何知道哪个验证器报告了错误？

**A:** 
- **Report Data Validator**: 验证 `report.json` 数据，错误信息包含字段路径（如 `extracted_data.rated_wattage`）
- **Template Validator**: 验证 Word 模板，错误信息包含占位符名称和位置（如 `table 0, row 7, cell 1`）

### Q7: 两个验证器都要运行吗？

**A:** **强烈推荐**在生成报告前都运行，因为：

1. **Report Data Validator** 检查数据问题（如缺失字段、类型错误、表格格式问题）
2. **Template Validator** 检查模板问题（如占位符被分割、配置缺失）

两个验证器检查的是不同层面的问题，互为补充。只有都通过了，生成报告时才不会出错。

---

## 附录：快速参考卡

### 命令速查表

| 命令 | 用途 |
|------|------|
| `python src/report_data_validator.py --report config/report.json` | 验证数据格式 |
| `python src/report_data_validator.py --report config/report.json --config config/report_config.jsonc` | 验证数据和配置一致性 |
| `python src/template_validator.py --template report_templates/template.docx` | 验证模板占位符 |
| `python src/template_validator.py --template report_templates/template.docx --config config/report_config.jsonc` | 验证模板占位符和配置 |
| `python src/calculator.py --config config/report_config.jsonc --report config/report.json --output output/calculated_report.json` | 执行计算 |
| `python src/field_mapper.py --config config/report_config.jsonc --report output/calculated_report.json --output output/operations.json` | 生成操作队列 |
| `python src/process_template.py --template report_templates/template.docx --operations output/operations.json --output output/final_report.docx` | 生成报告 |

### 退出码说明

| 工具 | 退出码 | 含义 |
|------|--------|------|
| `report_data_validator.py` | 0 | 验证通过，无错误无警告 |
| | 1 | 验证失败，有错误 |
| | 2 | 有警告（仅在 `--strict` 模式下） |
| `template_validator.py` | 0 | 验证通过 |
| | 1 | 有分割的占位符或未定义的占位符 |

### 推荐工作流程

```
准备阶段:
  1. 创建 report.json (metadata + extracted_data + calculated_data)
  2. 创建 Word 模板，添加占位符 {{placeholder_name}}
  3. 创建 report_config.json，定义字段映射

验证阶段（每次生成报告前执行）:
  4. 运行 report_data_validator.py 检查数据
  5. 运行 template_validator.py 检查模板
  6. 如果任一验证失败，修复问题后重新验证

生成阶段（验证通过后执行）:
  7. 运行 calculator.py 执行计算（如果需要）
  8. 运行 field_mapper.py 生成操作队列
  9. 运行 process_template.py 生成最终报告

检查阶段:
  10. 打开生成的报告，验证内容正确
  11. 如有问题，调整配置或数据，重新执行流程
```

---

**文档版本**: v1.0
**最后更新**: 2024
**维护者**: AI Assistant

如有问题，请参考项目 README 或联系开发团队。
    --report config/report.json \
    --config config/report_config.jsonc
```

常见问题检查清单：
- [ ] `source_field` 使用点号路径格式（如 `extracted_data.field`）
- [ ] `function` 名称正确（与注册名一致）
- [ ] `args` 中的路径都存在
- [ ] `type` 字段值为 `text`/`table`/`image` 之一

---

## 附录

### 文件路径说明

| 文件 | 路径 |
|------|------|
| 验证器 | `src/report_data_validator.py` |
| 计算器 | `src/calculator.py` |
| 字段映射器 | `src/field_mapper.py` |
| 模板处理器 | `src/process_template.py` |
| 完整指南 | `docs/VALIDATOR_AND_CALCULATOR_GUIDE.md` |

### 相关文档

- `README.md` - 项目概述
- `ARCHITECTURE_REFACTOR.md` - 架构迁移指南
- `TABLE_PROCESSOR_SUMMARY.md` - 表格处理器文档
- `AGENTS.md` - 开发规范

---

*文档版本: 1.0*
*最后更新: 2024-03*
