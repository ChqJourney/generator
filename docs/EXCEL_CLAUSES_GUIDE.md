# Excel 条款配置工具使用指南

## 概述

Excel 条款配置工具 (`tools/excel_config_editor.py`) 提供了一种高效的方式来批量管理 100+ 个条款配置。通过 Excel 表格的直观界面，业务人员可以轻松配置条款规则，然后通过脚本自动转换为 JSON 配置。

## 功能特性

- **Excel <-> JSON 双向转换**：支持从 Excel 生成 JSON 配置，也支持反向导出
- **批量编辑**：在 Excel 中批量修改多个条款
- **内置验证**：自动检查配置格式和语法错误
- **模板支持**：提供带示例的 Excel 模板，快速上手

## 安装依赖

```bash
pip install openpyxl json5
```

## 快速开始

### 1. 创建 Excel 模板

```bash
python tools/excel_config_editor.py template config/clauses_template.xlsx
```

这将创建一个包含示例条款的 Excel 模板文件。

### 2. 从现有配置反向导出（可选）

如果您已有 `report_config.jsonc` 配置，可以导出到 Excel 编辑：

```bash
python tools/excel_config_editor.py json2excel config/report_config.jsonc config/clauses.xlsx
```

### 3. 在 Excel 中编辑配置

打开 Excel 文件，按照模板格式填写条款配置：

| 列名 | 说明 | 示例 |
|------|------|------|
| `clause_id` | 条款唯一标识 | `v.1.a` |
| `template_field` | Word模板占位符（可选） | `v.1.a` |
| `param_names` | 参数名，逗号分隔 | `containing_product,light_sources` |
| `args` | 数据源路径，逗号分隔 | `metadata.containing_product.value,metadata.light_sources.value` |
| `rules` | 规则列表（格式：条件\|结果;条件\|结果） | `containing_product == 'true'\|Pass;containing_product == 'false'\|N/A` |
| `default` | 默认值 | `Fail` |
| `description` | 条款描述（可选） | `Containing product check` |

### 4. 验证 Excel 配置

```bash
python tools/excel_config_editor.py validate config/clauses.xlsx
```

### 5. 转换为 JSON

```bash
python tools/excel_config_editor.py excel2json config/clauses.xlsx output/clauses_config.json
```

如果要合并到现有的 `report_config.jsonc`：

```bash
python tools/excel_config_editor.py excel2json config/clauses.xlsx output/report_config_new.jsonc --config config/report_config.jsonc
```

## Excel 配置详解

### Rules 格式说明

Rules 列使用简单格式存储规则列表：

```
condition1|result1;condition2|result2;condition3|result3
```

**示例**：
```
containing_product == 'true' AND light_sources == 'true'|Pass;containing_product == 'false'|N/A
```

这表示：
- 如果 `containing_product == 'true' AND light_sources == 'true'`，返回 `Pass`
- 如果 `containing_product == 'false'`，返回 `N/A`

### 支持的表达式语法

**比较运算符**：
- `==` 等于
- `!=` 不等于
- `>` 大于
- `<` 小于
- `>=` 大于等于
- `<=` 小于等于

**逻辑运算符**：
- `AND` 逻辑与
- `OR` 逻辑或
- `NOT` 逻辑非

**特殊运算符**：
- `IN` / `NOT IN` 在列表中
- `CONTAINS` 包含子串
- `IS NULL` / `IS NOT NULL` 空值检查

**括号分组**：
使用 `()` 改变优先级，例如：
```
(A == 'true' AND B > 100) OR (C == 'false')
```

## 完整工作流程示例

### 场景：批量添加 100 个条款

1. **从现有配置导出**：
   ```bash
   python tools/excel_config_editor.py json2excel config/report_config.jsonc config/clauses_batch.xlsx
   ```

2. **在 Excel 中批量编辑**：
   - 复制示例行，修改 `clause_id` 和对应参数
   - 使用 Excel 的公式功能批量生成规则
   - 保存文件

3. **验证配置**：
   ```bash
   python tools/excel_config_editor.py validate config/clauses_batch.xlsx
   ```

4. **转换并合并**：
   ```bash
   python tools/excel_config_editor.py excel2json config/clauses_batch.xlsx config/report_config_new.jsonc --config config/report_config.jsonc
   ```

## 工具命令参考

```bash
# 创建模板
python tools/excel_config_editor.py template <输出.xlsx>

# Excel 转 JSON
python tools/excel_config_editor.py excel2json <输入.xlsx> <输出.json> [--config <原配置.jsonc>]

# JSON 转 Excel（反向导出）
python tools/excel_config_editor.py json2excel <输入.jsonc> <输出.xlsx>

# 验证配置
python tools/excel_config_editor.py validate <输入.xlsx>

# 查看帮助
python tools/excel_config_editor.py --help
```

## 最佳实践

1. **使用模板**：首次使用时先生成模板，了解格式要求
2. **定期验证**：编辑过程中定期运行 `validate` 命令检查错误
3. **备份原配置**：转换前备份 `report_config.jsonc`
4. **分批次处理**：条款数量较多时（100+），可以分多个 Excel 文件管理
5. **添加描述**：在 `description` 列添加条款含义说明，便于维护

## 常见问题

### Q: Excel 中的中文显示乱码？
A: 确保使用 UTF-8 编码保存 Excel 文件，或在 Windows 上设置正确的代码页。

### Q: 如何处理 100+ 个条款？
A: 可以：
- 在 Excel 中使用筛选和排序功能管理
- 将条款分组到多个 Excel 文件（如按模块分）
- 使用 Excel 的公式功能批量生成规则

### Q: 规则太长，Excel 单元格显示不全？
A: 双击列边界自动调整宽度，或启用单元格自动换行。

### Q: 转换后的 JSON 如何合并到 report_config.jsonc？
A: 使用 `--config` 参数指定原配置文件，工具会自动合并保留非条款配置。

## 与 CI/CD 集成

可以在构建流程中添加验证步骤：

```bash
# 在提交前验证
python tools/excel_config_editor.py validate config/clauses.xlsx
if [ $? -eq 0 ]; then
    python tools/excel_config_editor.py excel2json config/clauses.xlsx config/report_config.jsonc --config config/report_config.jsonc
fi
```
