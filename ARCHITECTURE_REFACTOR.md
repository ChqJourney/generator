# 重构后架构说明

## 新数据流

```
report.json (原始数据，包含metadata+extracted_data)
    ↓
calculator.py → calculated_report.json (计算后，添加calculated_data字段)
    ↓  
field_mapper.py → operations.json
    ↓
processor.py
```

## 核心变化

### 1. 数据合并
- **旧架构**: 3个独立JSON文件 (metadata.json, extracted_data.json, calculated_data.json)
- **新架构**: 1个分层JSON文件 (report.json)

### 2. 配置简化
- **旧配置**: 需要 `source` + `source_field`
- **新配置**: 仅需 `source_field` 使用点号路径

### 3. 表格数据源智能判断
支持两种表格数据源：
- **内嵌数据**: 直接是列表的列表
- **外部引用**: 包含 `type: "external"` 和 `source_id` 的字典

## 文件结构

### report.json 示例
```json
{
  "metadata": {
    "report_no": "RPT-2024-001",
    "applicant_name": "ABC Company",
    "issue_date": "2024-03-15"
  },
  "extracted_data": {
    "model_identifier": "DL-A-10W",
    "rated_wattage": "10.5",
    "useful_luminous_flux": "1050",
    "images": ["image1.jpg", "image2.jpg"],
    "photometric_data": {
      "type": "external",
      "source_id": "test_results.xlsx|Photometric Data",
      "start_row": 1,
      "mapping": {...}
    }
  },
  "calculated_data": {}
}
```

### report_config.json 示例
```json
{
  "field_mappings": [
    {
      "template_field": "report_no",
      "source_field": "metadata.report_no",
      "type": "text"
    },
    {
      "template_field": "rated_wattage",
      "source_field": "extracted_data.rated_wattage",
      "type": "text"
    },
    {
      "template_field": "energy_class_rating",
      "source_field": "calculated_data.energy_class_rating",
      "args": [
        "extracted_data.rated_wattage",
        "extracted_data.useful_luminous_flux"
      ],
      "function": "calculate_energy_class_rating",
      "type": "text"
    },
    {
      "template_field": "photometric_data",
      "source_field": "extracted_data.photometric_data",
      "table_template_path": "templates/table_template.docx",
      "type": "table"
    }
  ]
}
```

## 使用流程

### 1. 准备 report.json
```bash
# report.json 结构
{
  "metadata": {...},
  "extracted_data": {...},
  "calculated_data": {}
}
```

### 2. 执行计算
```bash
python src/calculator.py \
    --config config/report_config.json \
    --report config/report.json \
    --output output/calculated_report.json
```

**输出 calculated_report.json**:
```json
{
  "metadata": {...},
  "extracted_data": {...},
  "calculated_data": {
    "energy_class_rating": "A+",
    "energy_efficacy": "100.00"
  }
}
```

### 3. 生成操作队列
```bash
python src/field_mapper.py \
    --config config/report_config.json \
    --report output/calculated_report.json \
    --output output/operations.json
```

### 4. 处理模板（保持不变）
```python
from src.processor import DocxTemplateProcessor

processor = DocxTemplateProcessor(
    template_path="templates/report_template.docx",
    output_path="output/final_report.docx"
)

# 加载 operations.json 并执行
import json
with open('output/operations.json', 'r') as f:
    operations = json.load(f)

processor.operations = operations['operations']
processor.process()
```

## 关键改进点

### 1. 配置更简洁
```json
// 旧配置
{
  "template_field": "report_no",
  "source": "metadata",
  "source_field": "report_no",
  "type": "text"
}

// 新配置
{
  "template_field": "report_no",
  "source_field": "metadata.report_no",
  "type": "text"
}
```

### 2. 表格数据源灵活
支持两种方式：

**方式A - 外部Excel**（动态读取）：
```json
"photometric_data": {
  "type": "external",
  "source_id": "test_results.xlsx|Sheet1",
  "start_row": 1,
  "mapping": {...}
}
```

**方式B - 内嵌数据**（预提取）：
```json
"photometric_data": [
  ["Model-A", "0.15", "10.5", "1050"],
  ["Model-B", "0.16", "11.0", "1100"]
]
```

### 3. 计算字段配置
```json
{
  "template_field": "energy_class_rating",
  "source_field": "calculated_data.energy_class_rating",
  "args": [
    "extracted_data.rated_wattage",
    "extracted_data.useful_luminous_flux"
  ],
  "function": "calculate_energy_class_rating",
  "type": "text"
}
```

## 数据导航器 (DataNavigator)

内部工具类，支持点号路径访问：

```python
from src.calculator import DataNavigator

data = {
  "metadata": {"report_no": "RPT-001"},
  "extracted_data": {"rated_wattage": "10.5"}
}

# 获取值
value = DataNavigator.get_value(data, "extracted_data.rated_wattage")
# 结果: "10.5"

# 设置值
DataNavigator.set_value(data, "calculated_data.energy_class", "A+")
# data 变为:
# {
#   "metadata": {...},
#   "extracted_data": {...},
#   "calculated_data": {"energy_class": "A+"}
# }
```

## 错误处理

- **字段未找到**: 打印警告，跳过该字段
- **表格数据格式错误**: 提示无法识别格式
- **Excel文件不存在**: 捕获异常，返回空列表

## 向后兼容性

新架构**不兼容**旧架构：
- 删除了独立的 metadata.json, extracted_data.json
- 命令行参数变更（--metadata/--extracted_data → --report）
- 配置格式变更（移除source字段）

如需保留旧文件，需要手动合并到 report.json。

## 迁移指南

### 从旧架构迁移到新架构

1. **合并JSON文件**：
```python
import json

# 读取旧文件
with open('metadata.json') as f:
    metadata = json.load(f)
with open('extracted_data.json') as f:
    extracted = json.load(f)

# 创建report.json
report = {
    "metadata": metadata.get('fields', {}),  # 根据实际结构调整
    "extracted_data": extracted.get('targets', {}),
    "calculated_data": {}
}

with open('report.json', 'w') as f:
    json.dump(report, f, indent=2)
```

2. **更新配置文件**：
   - 将 `"source": "metadata"` + `"source_field": "xxx"` 改为 `"source_field": "metadata.xxx"`
   - 同样处理 `extracted_data` 和 `calculated_data`

3. **更新调用命令**：
   ```bash
   # 旧命令
   python src/field_mapper.py --config config.json --metadata m.json --extracted_data e.json --output out.json
   
   # 新命令（两步）
   python src/calculator.py --config config.json --report report.json --output calculated_report.json
   python src/field_mapper.py --config config.json --report calculated_report.json --output operations.json
   ```
