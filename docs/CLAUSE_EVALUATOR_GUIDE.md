# 通用条款判定模块使用指南

## 概述

通用条款判定模块 (`clause_evaluator.py`) 提供了一个可复用的条款判定解决方案，支持通过配置来实现复杂的判定逻辑，无需为每个条款编写独立的函数。

## 快速开始

### 1. 基础配置

在 `report_config.jsonc` 中添加条款判定映射：

```json
{
  "template_field": "v.1.a",
  "source_field": "calculated_data.v.1.a",
  "type": "text",
  "function": "evaluate_clause",
  "args": [
    "metadata.containing_product.value",
    "metadata.light_sources.value"
  ],
  "clause_config": {
    "clause_id": "v.1.a",
    "param_names": ["containing_product", "light_sources"],
    "rules": [
      {
        "condition": "containing_product == 'true' AND light_sources == 'true'",
        "result": "Pass"
      },
      {
        "condition": "containing_product == 'false'",
        "result": "N/A"
      }
    ],
    "default": "Fail"
  }
}
```

### 2. 运行计算器

```bash
python src/calculator.py \
    --config config/report_config.jsonc \
    --report config/report_data.json \
    --output output/calculated_report.json
```

## 配置详解

### 字段映射配置

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `template_field` | string | 是 | Word模板中的占位符名称 |
| `source_field` | string | 是 | 计算结果存储路径 |
| `type` | string | 是 | 数据类型，固定为 "text" |
| `function` | string | 是 | 固定为 "evaluate_clause" |
| `args` | array | 是 | 参数路径列表 |
| `clause_config` | object | 是 | 条款判定配置 |

### clause_config 配置

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `clause_id` | string | 是 | 条款ID，用于日志和调试 |
| `param_names` | array | 是 | 参数名列表，与 `args` 一一对应 |
| `rules` | array | 是 | 判定规则列表 |
| `default` | string | 是 | 默认返回值（无规则匹配时） |

### Rules 规则配置

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `condition` | string | 是 | 条件表达式 |
| `result` | string | 是 | 条件满足时的返回值 |

## 条件表达式语法

### 比较运算符

| 运算符 | 说明 | 示例 |
|--------|------|------|
| `==` | 等于 | `param == 'value'` |
| `!=` | 不等于 | `param != 0` |
| `>` | 大于 | `param > 100` |
| `<` | 小于 | `param < 50` |
| `>=` | 大于等于 | `param >= 0` |
| `<=` | 小于等于 | `param <= 100` |

### 逻辑运算符

| 运算符 | 说明 | 示例 |
|--------|------|------|
| `AND` | 逻辑与 | `param1 == 'true' AND param2 > 100` |
| `OR` | 逻辑或 | `param1 == 'true' OR param2 == 'true'` |
| `NOT` | 逻辑非 | `NOT (param == 'false')` |

### 特殊运算符

| 运算符 | 说明 | 示例 |
|--------|------|------|
| `IN` | 在列表中 | `param IN ['value1', 'value2']` |
| `NOT IN` | 不在列表中 | `param NOT IN ['value1', 'value2']` |
| `CONTAINS` | 包含子串 | `param CONTAINS 'substring'` |
| `IS NULL` | 为空 | `param IS NULL` |
| `IS NOT NULL` | 不为空 | `param IS NOT NULL` |

### 括号分组

使用括号 `()` 对条件进行分组：

```
(param1 == 'true' AND param2 > 100) OR (param3 == 'false' AND param4 < 50)
```

## 实际案例

### 案例 1：基础条款判定

```json
{
  "template_field": "v.1.a",
  "source_field": "calculated_data.v.1.a",
  "type": "text",
  "function": "evaluate_clause",
  "args": [
    "metadata.containing_product.value",
    "metadata.light_sources.value"
  ],
  "clause_config": {
    "clause_id": "v.1.a",
    "param_names": ["containing_product", "light_sources"],
    "rules": [
      {
        "condition": "containing_product == 'true' AND light_sources == 'true'",
        "result": "Pass"
      },
      {
        "condition": "containing_product == 'false'",
        "result": "N/A"
      }
    ],
    "default": "Fail"
  }
}
```

### 案例 2：复杂逻辑判定

```json
{
  "template_field": "v.II.1.a",
  "source_field": "calculated_data.v.II.1.a",
  "type": "text",
  "function": "evaluate_clause",
  "args": [
    "extracted_data.useful_luminous_flux",
    "extracted_data.Pon",
    "metadata.non_directional.value",
    "metadata.LED_source.value"
  ],
  "clause_config": {
    "clause_id": "v.II.1.a",
    "param_names": ["flux", "power", "non_directional", "LED_source"],
    "rules": [
      {
        "condition": "LED_source == 'false'",
        "result": "N/A"
      },
      {
        "condition": "flux IS NULL OR power IS NULL",
        "result": "Fail"
      },
      {
        "condition": "non_directional == 'true' AND (flux / power) >= 85",
        "result": "Pass"
      },
      {
        "condition": "non_directional == 'false' AND (flux / power) >= 75",
        "result": "Pass"
      }
    ],
    "default": "Fail"
  }
}
```

## 调试和常见问题

### 问题 1：条件不匹配

**排查**：
1. 检查 `args` 中的路径是否正确（checkbox 值需要使用 `.value` 后缀）
2. 检查 `param_names` 是否与 `args` 一一对应
3. 检查 `clause_config` 中的条件表达式语法是否正确

**示例**：

```json
// 错误示例 - args 路径不正确
"args": ["metadata.containing_product"]  // 缺少 .value

// 正确示例
"args": ["metadata.containing_product.value"]  // 完整的点号路径
```

### 问题 2：数值比较不生效

**原因**：checkbox 的值存储为字符串 `"true"` 或 `"false"`，而不是布尔值

**解决方案**：

```json
// 错误示例 - 直接比较布尔值
"condition": "param == true"

// 正确示例 - 与字符串 'true' 比较
"condition": "param == 'true'"
```

## 最佳实践

1. **使用参数名映射**：始终使用 `param_names` 将 `args` 映射为有意义的名称
2. **合理组织规则顺序**：规则按顺序匹配，先写特殊条件，后写通用条件
3. **使用括号明确优先级**：复杂条件使用括号明确优先级
4. **测试规则**：在正式使用前，使用简单的测试数据验证规则

## 测试

运行测试：

```bash
python test_clause_evaluator.py
```

## 总结

通用条款判定模块通过配置化方式实现了：

- **零代码新增条款**：新条款只需配置，无需编写函数
- **统一的判定逻辑**：所有条款使用相同的判定引擎
- **易于维护和审计**：规则以 JSON 格式存储，易于版本控制和审查

通过本方案，您可以从 100+ 个独立判定函数简化为 1 个通用判定函数 + 配置文件的组合，大幅提升开发和维护效率。
