# 字段计算器使用说明

## 概述

`calculator.py` 是一个灵活的字段计算器模块，用于根据配置文件中的规则计算字段值。它支持从 `extracted_data` 和 `metadata` 获取数据，并通过配置的函数进行计算。

## 主要特性

- **灵活配置**：通过 JSON 配置文件定义计算规则
- **内置函数**：提供常用的计算函数（能源等级、效率计算等）
- **可扩展**：支持自定义计算函数
- **类型转换**：自动将字符串转换为数值类型
- **错误处理**：完善的错误处理和日志输出
- **路径解析**：支持灵活的字段路径格式

## 使用方法

### 1. 独立使用计算器

```bash
python src/calculator.py \
    --config config/report_config.json \
    --metadata data_files/metadata.json \
    --extracted-data data_files/extracted_data.json \
    --output output/calculated_data.json
```

### 2. 集成到 field_mapper

```bash
python src/field_mapper.py \
    --config config/report_config.json \
    --metadata data_files/metadata.json \
    --extracted-data data_files/extracted_data.json \
    --calculated-data output/calculated_data.json \
    --output output/operations.json
```

### 3. 在 Python 代码中使用

```python
from src.calculator import FieldCalculator

# 创建计算器实例
calculator = FieldCalculator(
    metadata=metadata_dict,
    extracted_data=extracted_data_dict
)

# 处理整个配置
results = calculator.process_config(config)

# 或者计算单个字段
field_value = calculator.calculate_field(mapping_config)
```

## 配置格式

### 计算字段配置示例

```json
{
  "field_mappings": [
    {
      "template_field": "energy_class_rating",
      "source": "calculated_data",
      "source_field": "energy_class_rating",
      "args": [
        "extracted_data|rated_wattage",
        "extracted_data|useful_luminous_flux"
      ],
      "function": "calculate_energy_class_rating",
      "type": "text"
    },
    {
      "template_field": "energy_efficacy",
      "source": "calculated_data",
      "source_field": "energy_efficacy",
      "args": [
        "extracted_data|rated_wattage",
        "extracted_data|useful_luminous_flux"
      ],
      "function": "calculate_energy_efficacy",
      "type": "text"
    }
  ]
}
```

### 字段路径格式

- `extracted_data|field_name` - 从 extracted_data 获取字段
- `metadata|field_name` - 从 metadata 获取字段
- `field_name` - 默认从 extracted_data 获取

## 内置计算函数

### 1. calculate_energy_class_rating

计算能源等级评级（A++ 到 E）。

**参数：**
- `rated_wattage`: 额定功率 (W)
- `useful_luminous_flux`: 有用光通量 (lm)

**返回值：** 能源等级字符串

**计算公式：** η = Φ_use / P

### 2. calculate_energy_efficacy

计算能源效率值。

**参数：**
- `rated_wattage`: 额定功率 (W)
- `useful_luminous_flux`: 有用光通量 (lm)

**返回值：** 效率值字符串 (lm/W)，保留2位小数

### 3. calculate_percentage

计算百分比。

**参数：**
- `value`: 部分值
- `total`: 总值

**返回值：** 百分比字符串（如 "85.50%"）

### 4. format_number

格式化数字。

**参数：**
- `value`: 数值
- `decimal_places`: 小数位数（默认2）

**返回值：** 格式化后的数字字符串

### 5. concat

连接字符串。

**参数：**
- `*args`: 要连接的字符串
- `separator`: 分隔符（默认空格）

**返回值：** 连接后的字符串

### 6. multiply

乘法计算。

**参数：**
- `a`: 第一个数值
- `b`: 第二个数值

**返回值：** 乘积

### 7. divide

除法计算。

**参数：**
- `a`: 被除数
- `b`: 除数
- `default`: 除数为0时的默认值（默认0.0）

**返回值：** 商或默认值

## 自定义计算函数

### 方法1：直接注册

```python
from src.calculator import CalculationRegistry

@CalculationRegistry.register("my_custom_function")
def my_custom_function(arg1, arg2):
    # 你的计算逻辑
    return result
```

### 方法2：从模块加载

```python
# 创建 custom_calculations.py 文件
# 使用装饰器注册函数
# 然后在命令行指定：
python src/calculator.py \
    --functions-module custom_calculations \
    ...
```

## 输出格式

计算器输出 JSON 格式如下：

```json
{
  "calculated_fields": {
    "energy_class_rating": {
      "value": "A+",
      "source": "calculated_data",
      "field_name": "energy_class_rating"
    },
    "energy_efficacy": {
      "value": "95.50",
      "source": "calculated_data",
      "field_name": "energy_efficacy"
    }
  },
  "metadata": {
    "total_fields": 2,
    "calculator_config": {
      "strict_mode": false
    }
  }
}
```

## 错误处理

### 严格模式

使用 `--strict-mode` 标志启用严格模式，字段缺失时会报错。

### 常见错误

1. **FieldNotFoundError**: 字段未找到
2. **FunctionNotFoundError**: 计算函数未注册
3. **CalculatorError**: 一般计算错误

## 最佳实践

1. **使用有意义的函数名**：便于维护和理解
2. **添加类型转换**：确保输入参数是正确的类型
3. **处理边界情况**：除零检查、空值检查等
4. **编写文档字符串**：说明函数用途和参数
5. **单元测试**：为自定义函数编写测试

## 示例

### 示例1：计算光效

```json
{
  "template_field": "luminous_efficacy",
  "source": "calculated_data",
  "source_field": "luminous_efficacy",
  "args": [
    "extracted_data|total_luminous_flux",
    "extracted_data|rated_wattage"
  ],
  "function": "divide",
  "type": "text"
}
```

### 示例2：格式化带单位的值

```json
{
  "template_field": "power_with_unit",
  "source": "calculated_data",
  "source_field": "power_with_unit",
  "args": [
    "extracted_data|rated_wattage",
    "W"
  ],
  "function": "concat",
  "type": "text"
}
```

### 示例3：多个参数计算

```json
{
  "template_field": "test_result",
  "source": "calculated_data",
  "source_field": "test_result",
  "args": [
    "extracted_data|measured_power",
    "extracted_data|min_power_limit",
    "extracted_data|max_power_limit"
  ],
  "function": "check_pass_fail",
  "type": "text"
}
```

## 扩展阅读

- `src/custom_calculations_example.py`: 自定义函数示例
- `config/report_config.json`: 配置示例
