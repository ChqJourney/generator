# 快速配置流程（3步完成）

## 第1步：提取元素（30秒）

```bash
python tools/extract_template_elements.py 你的模板.docx --generate-config
```

输出：
- `config/extracted_elements.json` - 提取详情
- `config/extracted_elements_config.json` - 初步配置框架

## 第2步：批量配置（2-5分钟）

### 方式A：Excel编辑（推荐字段多）

```bash
# 导出
python tools/excel_config_editor.py export config/extracted_elements_config.json --output config/fields.xlsx

# 在Excel中编辑（会看到有颜色区分的表格）
# 编辑完成后，导回JSON
python tools/excel_config_editor.py import config/fields.xlsx --output config/report_config.json
```

### 方式B：命令行批量设置（适合有规律）

```bash
# 批量设置metadata字段
python tools/quick_field_setup.py config/extracted_elements_config.json \
    --source metadata \
    --fields report_no,issue_date,applicant_name,product_name,manufacturer \
    --output config/step1.json

# 批量设置extracted_data字段  
python tools/quick_field_setup.py config/step1.json \
    --source extracted_data \
    --fields model_identifier,rated_wattage,useful_luminous_flux \
    --output config/step2.json

# 设置计算字段
python tools/quick_field_setup.py config/step2.json --calculated \
    --function calculate_energy_class_rating \
    --args extracted_data.rated_wattage,extracted_data.useful_luminous_flux \
    --fields energy_class_rating \
    --output config/report_config.json
```

### 方式C：交互式向导（推荐新手）

```bash
python tools/config_wizard.py config/extracted_elements_config.json
# 按提示一步步确认即可
```

## 第3步：生成计算函数（30秒）

```bash
python tools/generate_calculator_functions.py config/report_config.json \
    --output src/custom_calculations.py
```

然后将生成的函数复制到 `src/calculator.py` 或保持为 `src/custom_calculations.py`。

## 完成！

现在你可以：
1. 检查 `config/report_config.json` 是否正确
2. 运行测试验证
3. 如有特殊计算逻辑，修改生成的函数

---

## 常用命令速查

| 任务 | 命令 |
|------|------|
| 提取placeholder | `python tools/extract_template_elements.py template.docx --generate-config` |
| 批量设置source | `python tools/quick_field_setup.py config.json --source metadata --fields a,b,c` |
| 批量设置计算字段 | `python tools/quick_field_setup.py config.json --calculated --function xxx --args a,b --fields c` |
| 导出Excel | `python tools/excel_config_editor.py export config.json --output fields.xlsx` |
| 导入Excel | `python tools/excel_config_editor.py import fields.xlsx --output config.json` |
| 生成计算函数 | `python tools/generate_calculator_functions.py config.json --output funcs.py` |
| 查看统计 | `python tools/quick_field_setup.py config.json --stats` |
| 查看未配置 | `python tools/quick_field_setup.py config.json --unconfigured` |

---

## 时间对比

| 方法 | 50个字段 | 100个字段 |
|------|---------|----------|
| 手工配置 | 2-3小时 | 4-6小时 |
| Excel方式 | 10-15分钟 | 20-30分钟 |
| 命令行批量 | 5-10分钟 | 10-15分钟 |
| 混合方式 | 5分钟 | 10分钟 |
