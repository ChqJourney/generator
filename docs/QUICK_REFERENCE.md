# 快速参考卡片

## 🚀 验证数据（一行命令）

```bash
python src/report_data_validator.py --report config/report.json --config config/report_config.json
```

## 🎯 验证模板（一行命令）

```bash
python src/template_validator.py --template report_templates/production_template.docx --config config/report_config.json
```

### 为什么需要两个验证器？

| 验证器 | 验证对象 | 检查内容 |
|--------|---------|---------|
| `report_data_validator.py` | report.json | 数据结构、字段类型、表格格式、图像路径、配置一致性 |
| `template_validator.py` | Word 模板 | 占位符是否被分割、占位符是否在配置中定义 |

**两个验证器都必须通过，才能确保报告生成成功！**

## 🧮 计算字段（一行命令）

```bash
python src/calculator.py --config config/report_config.json --report config/report.json --output output/calculated_report.json
```

## ⚙️ 配置 calculated_data 计算

### 方式 1: 预填充值（数据中直接提供）

```json
{
  "template_field": "energy_class",
  "source_field": "calculated_data.energy_class",
  "type": "text"
}
```

### 方式 2: 自动计算（推荐）

```json
{
  "template_field": "energy_class",
  "source_field": "calculated_data.energy_class",
  "type": "text",
  "function": "calculate_energy_class_rating",
  "args": ["extracted_data.rated_wattage", "extracted_data.useful_luminous_flux"]
}
```

## 📋 内置计算函数

| 函数 | 用途 | 示例配置 |
|------|------|----------|
| `calculate_energy_class_rating` | 能源等级 | `"function": "calculate_energy_class_rating", "args": ["extracted_data.wattage", "extracted_data.flux"]` |
| `calculate_energy_efficacy` | 能源效率 | `"function": "calculate_energy_efficacy", "args": ["extracted_data.wattage", "extracted_data.flux"]` |
| `calculate_percentage` | 百分比 | `"function": "calculate_percentage", "args": ["extracted_data.value", "extracted_data.total"]` |
| `multiply` | 乘法 | `"function": "multiply", "args": ["extracted_data.a", "extracted_data.b"]` |
| `divide` | 除法 | `"function": "divide", "args": ["extracted_data.a", "extracted_data.b"]` |
| `concat` | 字符串拼接 | `"function": "concat", "args": ["metadata.field1", "metadata.field2"], "separator": " "` |
| `format_number` | 格式化数字 | `"function": "format_number", "args": ["extracted_data.value", "2"]` |

## 🔍 常见验证错误速查

### Report Data Validator 错误

| 错误信息 | 原因 | 解决方案 |
|----------|------|----------|
| `缺少必需的顶级字段` | 缺少 metadata/extracted_data/calculated_data | 添加缺失字段 |
| `路径引用的字段不存在` | 配置引用的字段不存在 | 添加字段或修改配置路径 |
| `表格数据必须是列表` | 表格格式错误 | 确保是 `[["header"], ["data"]]` 格式 |
| `列数与表头不一致` | 表格行列不匹配 | 检查每行列数是否与表头相同 |
| `图像文件不存在` | 图像路径错误 | 确认路径正确或忽略（示例数据）|

### Template Validator 错误

| 错误信息 | 原因 | 解决方案 |
|----------|------|----------|
| `占位符被分割到多个 runs` | Word 将占位符分割成多个 runs | 在 Word 中删除并重新手动输入占位符 |
| `在配置中未定义的占位符` | 占位符在 config 中没有定义 | 在 report_config.json 中添加该占位符的映射 |

## 📁 文件模板

### 最小数据文件 (report.json)

```json
{
  "metadata": {
    "report_no": "RPT-001",
    "issue_date": "2024-03-15",
    "applicant_name": "Company Name"
  },
  "extracted_data": {
    "model_identifier": "MODEL-001",
    "rated_wattage": "10.5",
    "useful_luminous_flux": "1050"
  },
  "calculated_data": {}
}
```

### 最小配置文件 (report_config.json)

```json
{
  "template_path": "report_templates/template.docx",
  "output_dir": "output/",
  "field_mappings": [
    {
      "template_field": "report_no",
      "source_field": "metadata.report_no",
      "type": "text"
    },
    {
      "template_field": "model_identifier",
      "source_field": "extracted_data.model_identifier",
      "type": "text"
    },
    {
      "template_field": "energy_class",
      "source_field": "calculated_data.energy_class",
      "type": "text",
      "function": "calculate_energy_class_rating",
      "args": ["extracted_data.rated_wattage", "extracted_data.useful_luminous_flux"]
    }
  ]
}
```

## 🔧 Python 代码示例

### 验证数据

```python
from src.report_data_validator import ReportDataValidator
import json

with open('config/report.json') as f:
    data = json.load(f)

validator = ReportDataValidator(data)
report = validator.validate()
report.print_report()
```

### 计算字段

```python
from src.calculator import FieldCalculator
import json

with open('config/report.json') as f:
    data = json.load(f)
with open('config/report_config.json') as f:
    config = json.load(f)

calculator = FieldCalculator(data)
calculator.process_config(config)

with open('output/calculated_report.json', 'w') as f:
    json.dump(calculator.get_calculated_report(), f, indent=2)
```

---

*打印此页，随时查阅*
