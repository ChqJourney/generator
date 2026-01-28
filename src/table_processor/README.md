# 表格数据处理与插入模块

独立的表格处理模块，支持灵活的数据转换和表格插入功能。

## 模块结构

```
src/table_processor/
├── __init__.py              # 模块入口
├── data_transformer.py       # 数据转换器
└── table_inserter.py         # 增强版表格插入器
```

## 核心功能

### 1. 数据转换器 (TableDataTransformer)

支持的数据转换类型：

#### skip_columns - 跳过指定列
```python
{
    "type": "skip_columns",
    "columns": [0, 1]  # 跳过第0列和第1列
}
```

#### add_column - 添加列
```python
# 添加行的索引
{
    "type": "add_column",
    "position": 0,
    "source": "row_index"
}

# 从元数据添加值
{
    "type": "add_column",
    "position": 0,
    "source": "metadata:model_name"
}

# 添加固定值
{
    "type": "add_column",
    "position": 0,
    "source": "value:MODEL-X"
}
```

#### calculate - 计算列
```python
# 计算平均值
{
    "type": "calculate",
    "column": 3,
    "operation": "average",
    "decimal": 1
}

# 计算总和
{
    "type": "calculate",
    "column": 2,
    "operation": "sum",
    "decimal": 2
}

# 使用公式计算（A=0列, B=1列, ...）
{
    "type": "calculate",
    "column": 4,
    "operation": "formula=Power{row}/Current{row}*1000",
    "decimal": 1
}
```

#### format_column - 格式化列（函数式）
```python
# 固定小数位
{
    "type": "format_column",
    "column": 1,
    "decimal": 2
}

# CCT格式化：<10保留1位，>=10保留0位
{
    "type": "format_column",
    "column": 5,
    "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
}

# 光效格式化：<1保留4位，>=1保留2位
{
    "type": "format_column",
    "column": 4,
    "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
}

# XY坐标格式化：<0.1保留6位，>=0.1保留4位
{
    "type": "format_column",
    "column": 8,
    "function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"
}

# 闪烁值格式化：<1保留4位，>=1保留2位
{
    "type": "format_column",
    "column": 11,
    "function": "lambda x: f'{x:.4f}' if abs(x) < 1 else f'{x:.2f}'"
}
```

#### reorder - 列重排序
```python
{
    "type": "reorder",
    "order": [0, 4, 1, 2, 3, 5, 6, 7, 8, 9]
}
```

#### filter_rows - 行过滤
```python
{
    "type": "filter_rows",
    "condition": "remove_empty"
}
```

### 2. 表格插入器 (EnhancedTableInserter)

支持两种行策略：

#### fixed_rows - 固定行数模式
- 保留模板中的所有行
- 数据填入时跳过表头和skip_columns指定的列
- 适用于行数固定的表格

#### dynamic_rows - 动态行数模式
- 保留表头行
- 删除所有数据行
- 根据数据量重新生成行
- 适用于行数不固定的表格

## 使用示例

### 基本使用

```python
from src.table_processor import TableDataTransformer, EnhancedTableInserter
from docx import Document

# 1. 准备数据
raw_data = [
    ["0.1500", "12.5", "0.95", "1500", "120.5", "2700", "95", "85"],
    ["0.2500", "25.0", "0.96", "2500", "100.0", "3000", "95", "85"],
    ["0.3500", "37.5", "0.97", "3500", "107.1", "3300", "95", "85"]
]

metadata = {"model_name": "LED-001"}

# 2. 定义转换规则
transformations = [
    {
        "type": "skip_columns",
        "columns": [0]  # 跳过Excel中的序号列
    },
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
    }
]

# 3. 使用数据转换器
transformer = TableDataTransformer()
transformed_data = transformer.transform(raw_data, transformations, metadata)

# 4. 插入到Word文档
doc = Document("template.docx")
inserter = EnhancedTableInserter(doc)

inserter.insert(
    placeholder="photometric_data",
    table_template_path="templates/table_template.docx",
    raw_data=raw_data,
    transformations=transformations,
    metadata=metadata,
    row_strategy='fixed_rows',
    skip_columns=[0, 1],  # 模板中的序号和型号列是固定的
    header_rows=2,          # 前2行是表头
    location='body'
)

doc.save("output.docx")
```

### 完整配置示例

参考 `config/table_processor_config_example.json` 查看所有支持的转换类型和配置。

## 运行测试

```bash
python test_table_processor.py
```

测试脚本将运行以下测试：
1. 数据转换器测试
2. 复杂小数点格式化测试
3. 增强版表格插入器测试

## 设计特点

1. **函数式小数点控制**：使用lambda表达式实现灵活的格式化规则
2. **两种行策略**：支持固定行数和动态行数两种模式
3. **skip_columns机制**：自动跳过模板中的固定列（序号、型号等）
4. **独立模块**：不依赖processor.py，便于测试和单独使用
5. **配置驱动**：所有转换规则通过JSON配置

## 后续集成

测试完成后，可以集成到 `src/processor.py`'s `DocxTemplateProcessor` 类中：

```python
# 在 processor.py 中
from table_processor import TableDataTransformer, EnhancedTableInserter

# 在 DocxTemplateProcessor.process() 中
elif op['type'] == 'table':
    transformer = TableDataTransformer()
    inserter = EnhancedTableInserter(self.doc, transformer)
    inserter.insert(...)
```

## 常见格式化规则

| 列类型 | 格式化规则 |
|--------|-----------|
| CCT | `lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'` |
| 光效 | `lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'` |
| XY坐标 | `lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'` |
| 功率因数 | `lambda x: f'{x:.4f}'` |
| 闪烁值 | `lambda x: f'{x:.4f}' if abs(x) < 1 else f'{x:.2f}'` |
| 百分比 | `lambda x: f'{x:.1f}%'` |
