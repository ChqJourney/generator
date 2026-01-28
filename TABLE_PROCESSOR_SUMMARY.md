# 表格数据处理与插入模块 - 实现总结

## 已完成的工作

### 1. 模块结构
```
src/table_processor/
├── __init__.py                    # 模块入口
├── data_transformer.py             # 数据转换器
├── table_inserter.py               # 增强版表格插入器
└── README.md                      # 使用文档
```

### 2. 核心功能实现

#### 数据转换器 (TableDataTransformer)
支持的数据转换类型：
- ✅ **skip_columns** - 跳过指定列
- ✅ **add_column** - 添加列（行索引、元数据、固定值）
- ✅ **calculate** - 计算列（average, sum, max, min, formula）
- ✅ **format_column** - 格式化列（固定小数位、函数式）
- ✅ **reorder** - 列重排序
- ✅ **filter_rows** - 行过滤

#### 增强版表格插入器 (EnhancedTableInserter)
- ✅ **fixed_rows** - 固定行数模式
- ✅ **dynamic_rows** - 动态行数模式
- ✅ **skip_columns** - 自动跳过模板中的固定列
- ✅ **header_rows** - 指定表头行数

### 3. 小数点控制 - 函数式实现

通过lambda表达式实现灵活的格式化规则：

```python
# CCT格式化：<10保留1位，>=10保留0位
"function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"

# 光效格式化：<1保留4位，>=1保留2位
"function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"

# XY坐标格式化：<0.1保留6位，>=0.1保留4位
"function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"

# 闪烁值格式化：<1保留4位，>=1保留2位
"function": "lambda x: f'{x:.4f}' if abs(x) < 1 else f'{x:.2f}'"
```

### 4. 测试结果

测试脚本 `test_table_processor.py` 包含：
- ✅ 数据转换器测试
- ✅ 复杂小数点格式化测试（CCT、光效、XY坐标）
- ✅ 增强版表格插入器测试（需要模板文件）

测试输出：
```
数据转换器测试: 通过
复杂小数点格式化测试: 通过
  - CCT格式化: 5.5 -> 5.5, 15.3 -> 15, 2500 -> 2500
  - 光光格式化: 0.85 -> 0.8500, 95.5 -> 95.50, 120.3 -> 120.30
  - XY坐标格式化: 0.04 -> 0.040000, 0.35 -> 0.3500
增强版表格插入器: 等待模板文件
```

### 5. 配置文件示例

已创建 `config/table_processor_config_example.json`，包含：
- 完整的表格配置示例
- 所有转换类型的配置示例
- 常用格式化规则示例

### 6. JSON配置文件中使用函数式小数点控制

在JSON配置文件中，lambda函数作为字符串写入：

#### 基本语法
```json
{
  "type": "format_column",
  "column": 4,
  "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
}
```

**注意**：
- `"function"` 的值必须是字符串（用双引号包裹）
- lambda函数内部用**单引号**或**转义的双引号**`\"`
- 使用**原生f-string语法**（不需要转义大括号）

#### 常见格式化规则示例

```json
{
  "transformations": [
    {
      "type": "format_column",
      "column": 0,
      "decimal": 4
    },
    
    {
      "type": "format_column",
      "column": 4,
      "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
    },
    
    {
      "type": "format_column",
      "column": 5,
      "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
    },
    
    {
      "type": "format_column",
      "column": 8,
      "function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"
    },
    
    {
      "type": "format_column",
      "column": 11,
      "function": "lambda x: f'{x:.4f}' if abs(x) < 1 else f'{x:.2f}'"
    }
  ]
}
```

#### 完整的 report_config.json 示例

```json
{
  "field_mappings": [
    {
      "template_field": "photometric_data",
      "source": "extracted_data",
      "source_field": "photometric_data",
      "type": "table",
      "table_template_path": "report_templates/tables/photometric_table_template.docx",
      
      "row_strategy": "fixed_rows",
      "skip_columns": [0, 1],
      "header_rows": 2,
      
      "transformations": [
        {
          "type": "skip_columns",
          "columns": [0, 1]
        },
        
        {
          "type": "calculate",
          "column": 4,
          "operation": "formula=Power{row}/Current{row}*1000",
          "decimal": 4
        },
        
        {
          "type": "format_column",
          "column": 0,
          "decimal": 4
        },
        
        {
          "type": "format_column",
          "column": 1,
          "decimal": 2
        },
        
        {
          "type": "format_column",
          "column": 2,
          "function": "lambda x: f'{x:.4f}'"
        },
        
        {
          "type": "format_column",
          "column": 4,
          "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
        },
        
        {
          "type": "format_column",
          "column": 5,
          "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
        },
        
        {
          "type": "format_column",
          "column": 8,
          "function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"
        },
        
        {
          "type": "format_column",
          "column": 9,
          "function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"
        },
        
        {
          "type": "format_column",
          "column": 11,
          "function": "lambda x: f'{x:.4f}' if abs(x) < 1 else f'{x:.2f}'"
        }
      ]
    }
  ]
}
```

#### 预定义格式化规则速查表

| 数据类型 | Lambda表达式 | 说明 |
|---------|-------------|------|
| 电流 | `lambda x: f'{x:.4f}'` | 固定4位 |
| 功率 | `lambda x: f'{x:.2f}'` | 固定2位 |
| 功率因数 | `lambda x: f'{x:.4f}'` | 固定4位 |
| 光效 | `lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'` | <1:4位, ≥1:2位 |
| 光通量 | `lambda x: f'{x:.1f}'` | 固定1位 |
| CCT | `lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'` | <10:1位, ≥10:0位 |
| Ra/R9 | `lambda x: f'{x:.0f}'` | 固定0位（整数） |
| XY坐标 | `lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'` | <0.1:6位, ≥0.1:4位 |
| SDCM | `lambda x: f'{x:.0f}'` | 固定0位 |
| 闪烁值 | `lambda x: f'{x:.4f}' if abs(x) < 1 else f'{x:.2f}'` | <1:4位, ≥1:2位 |
| 待机功率 | `lambda x: f'{x:.2f}'` | 固定2位 |

#### 工作原理

代码通过 `eval()` 将JSON字符串转换为可执行的lambda函数：

```python
# data_transformer.py 中的处理逻辑
func_str = config['function']  # 从JSON读取: "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
format_func = eval(func_str)  # 转换为函数对象

value = 0.85
formatted_value = format_func(value)  # 执行: '0.8500'
```

**安全性提示**：使用 `eval()` 会执行任意Python代码，请确保配置文件来源可信。

## 使用方法

### 基本使用流程

```python
from src.table_processor import TableDataTransformer, EnhancedTableInserter
from docx import Document

# 1. 准备数据
raw_data = [...]
metadata = {"model_name": "LED-001"}

# 2. 定义转换规则
transformations = [
    {"type": "skip_columns", "columns": [0, 1]},
    {"type": "format_column", "column": 0, "decimal": 4},
    {"type": "format_column", "column": 4, 
     "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"}
]

# 3. 插入到Word文档
doc = Document("template.docx")
inserter = EnhancedTableInserter(doc)

inserter.insert(
    placeholder="photometric_data",
    table_template_path="templates/table_template.docx",
    raw_data=raw_data,
    transformations=transformations,
    metadata=metadata,
    row="strategy='fixed_rows',
    skip_columns=[0, 1],
    header_rows=2
)

doc.save("output.docx")
```

### 运行测试

```bash
python test_table_processor.py
```

## 设计特点

1. **函数式小数点控制** - 使用lambda表达式实现任意复杂的格式化规则
2. **skip_columns机制** - 自动跳过模板中的固定列（序号、型号等）
3. **两种行策略** - fixed_rows（固定行数）和 dynamic_rows（动态行数）
4. **独立模块** - 不依赖processor.py，便于测试和单独使用
5. **配置驱动** - 所有转换规则通过JSON配置
6. **向后兼容** - 不提供transformations时行为与原系统一致

## 文件清单

```
generator/
├── src/
│   └── table_processor/
│       ├── __init__.py
│       ├── data_transformer.py
│       ├── table_inserter.py
│       └── README.md
├── config/
│   └── table_processor_config_example.json
└── test_table_processor.py
```

## 后续集成建议

测试完成并确认功能正常后，可以集成到现有系统：

### 1. 修改 `field_mapper.py`

在 `generate_operations()` 的table类型处理中增加：
```python
table_config['transformations'] = mapping.get('transformations')
table_config['row_strategy'] = mapping.get('row_strategy', 'fixed_rows')
table_config['skip_columns'] = mapping.get('skip_columns')
table_config['header_rows'] = mapping.get('header_rows', 1)
```

### 2. 修改 `processor.py`

在 `DocxTemplateProcessor.process()` 中：
```python
elif op['type'] == 'table':
    from table_processor import TableDataTransformer, EnhancedTableInserter
    transformer = TableDataTransformer()
    inserter = EnhancedTableInserter(self.doc)
    inserter.insert(
        placeholder=op['placeholder'],
        table_template_path=op['table_template_path'],
        raw_data=op.get('table_data'),
        transformations=op.get('transformations'),
        metadata=op.get('metadata'),
        row_strategy=op.get('row_strategy', 'fixed_rows'),
        skip_columns=op.get('skip_columns'),
        header_rows=op.get('header_rows', 1),
        location='body'
    )
```

### 3. 更新 `report_config.json`

```json
{
  "type": "table",
  "row_strategy": "fixed_rows",
  "skip_columns": [0, 1],
  "header_rows": 2,
  "transformations": [
    {"type": "skip_columns", "columns": [0, 1]},
    {"type": "format_column", "column": 4, 
     "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"}
  ]
}
```

## 总结

✅ 模块已独立实现
✅ 支持函数式小数点控制
✅ 支持skip_columns机制
✅ 支持fixed_rows和dynamic_rows两种策略
✅ 测试脚本已创建并验证通过
✅ 配置文件示例已提供
✅ 使用文档已完成

**当前状态**：模块实现完成，支持在JSON配置文件中使用函数式小数点控制。等待测试模板文件以进行完整的插入器测试，测试完成后可决定是否集成到processor.py。
