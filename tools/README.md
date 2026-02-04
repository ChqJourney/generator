# Report Config 配置工具集

这组工具帮助你快速从Word模板提取placeholder和checkbox，并生成完整的report_config.json配置。

## 工具列表

| 工具 | 用途 |
|------|------|
| `extract_template_elements.py` | 从Word模板提取所有placeholder和checkbox |
| `config_wizard.py` | 交互式配置向导，引导完成配置 |
| `excel_config_editor.py` | 导出到Excel批量编辑，再转回JSON |
| `quick_field_setup.py` | 批量设置同类字段的配置 |
| `generate_calculator_functions.py` | 自动生成calculator.py中的计算函数 |

## 快速开始

### 第一步：提取元素

从Word模板中提取所有placeholder和checkbox：

```bash
python tools/extract_template_elements.py report_templates/template.docx --generate-config
```

这会生成两个文件：
- `config/extracted_elements.json` - 提取的元素详情
- `config/extracted_elements_config.json` - 初步的配置框架

### 第二步：选择配置方式

#### 方式A：交互式向导（推荐新手）

```bash
python tools/config_wizard.py config/extracted_elements.json
```

向导会引导你：
1. 批量设置text字段的数据源
2. 逐个确认计算字段的函数和参数
3. 配置table和image的特殊属性
4. 验证配置完整性

#### 方式B：Excel批量编辑（推荐大量字段）

```bash
# 导出到Excel
python tools/excel_config_editor.py export config/extracted_elements.json --output config/fields.xlsx

# 在Excel中编辑后，导回JSON
python tools/excel_config_editor.py import config/fields.xlsx --output config/report_config.json
```

Excel表格包含：
- 颜色编码的字段类型
- 下拉式选项
- 使用说明sheet

#### 方式C：命令行批量设置（适合有规律的情况）

```bash
# 批量设置metadata字段
python tools/quick_field_setup.py config/extracted.json \
    --source metadata \
    --fields report_no,issue_date,applicant_name,product_name

# 批量设置extracted_data字段
python tools/quick_field_setup.py config/extracted.json \
    --source extracted_data \
    --fields model_identifier,rated_wattage,useful_luminous_flux

# 批量设置计算字段
python tools/quick_field_setup.py config/extracted.json --calculated \
    --function calculate_energy_class_rating \
    --args extracted_data.rated_wattage,extracted_data.useful_luminous_flux \
    --fields energy_class_rating
```

### 第三步：生成计算函数

根据配置自动生成calculator.py中的函数：

```bash
# 生成到单独文件
python tools/generate_calculator_functions.py config/report_config.json \
    --output src/custom_calculations.py

# 或者追加到calculator.py
python tools/generate_calculator_functions.py config/report_config.json \
    --append-to src/calculator.py
```

### 第四步：验证和调整

```bash
# 查看配置统计
python tools/quick_field_setup.py config/report_config.json --stats

# 查看未配置字段
python tools/quick_field_setup.py config/report_config.json --unconfigured

# 运行验证
python src/validate_report.py --report config/report.json --config config/report_config.json
```

## 典型工作流程

### 场景1：新模板配置（50+字段）

```bash
# 1. 提取
python tools/extract_template_elements.py template.docx --generate-config

# 2. 导出到Excel批量编辑
python tools/excel_config_editor.py export config/extracted_elements_config.json --output config/fields.xlsx

# 3. 在Excel中批量填写，保存

# 4. 导回JSON
python tools/excel_config_editor.py import config/fields.xlsx --output config/report_config.json

# 5. 生成计算函数
python tools/generate_calculator_functions.py config/report_config.json --append-to src/calculator.py

# 6. 手动实现特殊计算逻辑
```

### 场景2：修改现有配置（添加10个字段）

```bash
# 1. 提取新模板
python tools/extract_template_elements.py new_template.docx

# 2. 使用向导合并和配置
python tools/config_wizard.py config/extracted_elements.json

# 3. 生成新增的计算函数
python tools/generate_calculator_functions.py config/extracted_elements_final.json
```

### 场景3：批量修改字段类型

```bash
# 将所有包含image关键词的字段设为image类型
python tools/quick_field_setup.py config/extracted.json \
    --type image \
    --fields product_image,logo,signature,test_photo

# 批量设置image配置
python tools/quick_field_setup.py config/extracted.json --image-config \
    --width 4.0 --alignment center \
    --fields product_image,logo,signature,test_photo
```

## 智能推断规则

`extract_template_elements.py`会根据placeholder名称智能推断：

### 类型推断

| 关键词 | 推断类型 |
|--------|----------|
| image, img, photo, picture, pic, logo | image |
| table, data, list, items, records | table |
| cb_, chk_ | checkbox |
| 其他 | text |

### 数据源推断

| 关键词 | 推断数据源 |
|--------|------------|
| report, issue_date, applicant, product, manufacturer | metadata |
| energy, efficacy, class, rating, calculate | calculated_data |
| 其他 | extracted_data |

### 计算函数推断

| 字段名关键词 | 推断函数 |
|--------------|----------|
| energy_class, class_rating | calculate_energy_class_rating |
| efficacy | calculate_energy_efficacy |
| percentage, percent | calculate_percentage |

## 文件说明

### 中间文件

- `extracted_elements.json` - 原始提取结果，包含placeholder和checkbox列表
- `extracted_elements_config.json` - 初步生成的配置框架
- `fields.xlsx` - Excel编辑版本
- `*_final.json` - 向导生成的最终配置
- `*_calculator_functions.py` - 自动生成的计算函数模板

### 最终文件

- `report_config.json` - 最终的配置文件
- `custom_calculations.py` - 自定义计算函数

## 常见问题

### Q: 如何处理重复字段名？

A: 配置工具会自动去重。如果同一placeholder在多处出现，只需配置一次。

### Q: 如何批量修改多个字段？

A: 使用Excel方式或`quick_field_setup.py`的`--fields`参数（逗号分隔）。

### Q: 生成的计算函数不完整怎么办？

A: 工具会提示需要手动实现的函数。打开生成的文件，搜索`TODO`或`result = 0`，实现具体逻辑。

### Q: 如何验证配置正确性？

A: 运行验证脚本：
```bash
python src/validate_report.py --report config/report.json --config config/report_config.json
```

## 提示

1. **命名规范**: 在Word模板中使用有意义的placeholder名称，如`{{applicant_name}}`而不是`{{field1}}`，这样智能推断会更准确。

2. **分组处理**: 将同类字段一起处理，如先配置所有metadata字段，再配置所有extracted_data字段。

3. **复用函数**: 相似的计算逻辑使用相同的函数名，只需要实现一次。

4. **备份**: 修改配置前备份`report_config.json`。
