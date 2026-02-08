# Word Document Report Generator - Agent Guide

## Project Overview

This is a Python-based Word document report generation system that automates the creation of professional reports by inserting extracted product data into pre-designed Word templates. It supports text replacement, table insertion, and image embedding with placeholder-based templating.

**Key Use Case**: Generating lighting product test reports (photometric data, energy ratings, etc.) from Excel test results and metadata.

---

## Technology Stack

- **Language**: Python 3.x
- **Core Dependencies**:
  - `python-docx` (>=1.2.0) - Word document manipulation
  - `openpyxl` (>=3.0.0) - Excel file reading
  - `lxml` (>=4.0.0) - XML processing
  - `typing_extensions` (>=4.0.0) - Type hints
- **Testing**: pytest with mock support
- **Build Tool**: pyproject.toml (minimal configuration)

---

## Project Structure

```
docx/
├── config/                          # Configuration files
│   ├── report_config.json          # Main field mapping configuration
│   ├── report_data.json            # Example report data structure
├── data_files/                     # Data sources
│   └── TDS.xlsx                   # Excel test data source
├── report_templates/               # Word templates
│   ├── production_template.docx      # Main report template
│   └── tables
├── src/                            # Source code
│   ├── processor.py               # Core template processing engine
│   ├── field_mapper.py            # Field mapping to operations
│   ├── calculator.py              # Field value calculation with registry
│   ├── process_template.py        # CLI wrapper for processor
│   ├── report_data_validator.py   # Comprehensive data validator
│   ├── template_validator.py      # Template placeholder validator
│   ├── utils                      # utils scripts, logging,path,and others
│   └── table_processor/           # Table processing module
│       ├── __init__.py
│       ├── data_transformer.py    # Data transformation rules
│       └── custom_transformers.py      # custom transformer


```

---

## Architecture and Data Flow

### Data Structure Hierarchy

The system uses a three-tier hierarchical data structure:

```json
{
  "metadata": {
    "report_no": "RPT-001",
    "issue_date": "2024-03-15",
    "applicant_name": "...",
    "containing_product": {
      "type": "checkbox",
      "value": "true"
    }
  },
  "extracted_data": {
    "model_identifier": "LED-100W",
    "rated_wattage": "100",
    "photometric_data": {...}
  },
  "calculated_data": {
    "energy_class_rating": "A+",
    "energy_efficacy": "100.00"
  }
}
```

**Field Access**: Use dot-notation paths like `metadata.report_no`, `extracted_data.rated_wattage`.

---

## Core Components

### 1. Calculator (`src/calculator.py`)

**Purpose**: Calculate derived field values from raw data.

**Key Classes**:
- `DataNavigator`: Dot-notation path access for hierarchical data
- `CalculationRegistry`: Decorator-based function registration system
- `FieldCalculator`: Main calculation engine

**Built-in Functions**:
- `calculate_energy_class_rating` - Energy efficiency class (A++ to E)
- `calculate_energy_efficacy` - Lumens per watt calculation
- `calculate_percentage`, `format_number`, `concat`, `multiply`, `divide`

**Custom Functions**: Use `@CalculationRegistry.register("function_name")` decorator.

**CLI Usage**:
```bash
python src/calculator.py \
    --config config/report_config.json \
    --report config/report.json \
    --output output/calculated_report.json
```

### 2. Field Mapper (`src/field_mapper.py`)

**Purpose**: Convert field mappings to operations array for processor.

**Key Functions**:
- `generate_operations(config, report_data)` - Main conversion function
- `get_value_by_path(data, path)` - Dot-notation value extraction
- `build_table_data_from_excel(value, target_headers)` - Excel data extraction
- `get_xlsx_to_list(...)` - Excel to list conversion

**Table Data Sources**:
- **Embedded**: Direct list-of-lists in JSON
- **External**: `{type: "external", source_id: "file.xlsx|SheetName", start_row: N, mapping: {...}}`

**Checkbox Data Handling**:
- Parses checkbox values from `{"type": "checkbox", "value": "true/false"}` format
- Defaults to `false` when checkbox field is missing in report data
- Generates `checkbox_mapping` operation for each checkbox field

**CLI Usage**:
```bash
python src/field_mapper.py \
    --config config/report_config.json \
    --report output/calculated_report.json \
    --output output/operations.json
```

### 3. Processor (`src/processor.py`)

**Purpose**: Core engine that applies operations to Word templates.

**Key Classes**:
- `DocxTemplateProcessor`: Main processor class
- `TextInserter`: Handles text replacement
- `TableInserter`: Handles table insertion with transformations
- `ImageInserter`: Handles image insertion
- `CheckboxInserter`: Updates form checkbox states

**Placeholder Format**: `{{placeholder_name}}` for text/table/image; checkbox names match Word form field names

**Supported Locations**: `body`, `header`, `footer` (including first_page_header, even_page_header, etc.)

**Checkbox Processing**:
- `CheckboxInserter` class handles form checkbox state updates
- Automatically sets unchecked state for template checkboxes not in operations
- Checkbox names must exactly match Word form field names (extract using `tools/extract_template_elements.py`)

### 4. Table Processor Module (`src/table_processor/`)

**Purpose**: Advanced table data transformation and insertion.

**DataTransformer Features**:
- `skip_columns` - Skip specified columns
- `add_column` - Add columns (row index, metadata lookup, fixed value)
- `calculate` - Calculations (average, sum, max, min, formula)
- `format_column` - Format with fixed decimals or lambda functions
- `reorder` - Column reordering
- `filter_rows` - Row filtering

**Row Strategies**:
- `fixed_rows` - Fill existing rows in template
- `dynamic_rows` - Add/remove rows to match data

### 5. Report Data Validator (`src/report_data_validator.py`)

**Purpose**: Comprehensive validation of report.json data format, content, and consistency with configuration.

**Validation Scope**:
- Data structure integrity (metadata, extracted_data, calculated_data)
- Field type and data type validation
- Table data format validation (headers, column consistency)
- Image path validation (file existence)
- Configuration consistency check (data fields mapped in config)

**CLI Usage**:
```bash
# Basic validation (data format only)
python src/report_data_validator.py --report config/report.json

# Full validation (including config consistency)
python src/report_data_validator.py --report config/report.json --config config/report_config.json

# Strict mode (non-zero exit code on warnings)
python src/report_data_validator.py --report config/report.json --config config/report_config.json --strict
```

### 6. Template Validator (`src/template_validator.py`)

**Purpose**: Validate Word template placeholders, specifically checking if placeholders are split across multiple runs (the root cause of `{{}}`残留).

**Validation Scope**:
- Check if placeholders are split across multiple runs
- Validate if placeholders are defined in configuration
- Provide detailed repair suggestions

**CLI Usage**:
```bash
# Check for split placeholders only (no config needed)
python src/template_validator.py --template report_templates/production_template.docx

# Also check if placeholders are defined in config
python src/template_validator.py --template report_templates/production_template.docx --config config/report_config.json
```

**Fixing Split Placeholders**:

When Template Validator reports a placeholder is split (e.g., `{{efficacy}}` split into `['{{', 'efficacy', '}}']`):

1. Open the template in Word
2. Locate the placeholder
3. Delete the entire placeholder (including `{{}}`)
4. **Manually retype** `{{placeholder_name}}` (do not copy-paste)
5. Save the template and re-run Template Validator

### 7. Validation Workflow (Recommended)

Before generating reports, it's strongly recommended to run both validators:

```bash
# Step 1: Validate report.json data format and content
python src/report_data_validator.py \
    --report config/report.json \
    --config config/report_config.json

# Step 2: Validate Word template (check if placeholders will be processed correctly)
python src/template_validator.py \
    --template report_templates/production_template.docx \
    --config config/report_config.json

# Step 3: If both validations pass, proceed with report generation
# (Run calculator.py → field_mapper.py → process_template.py)
```

**Why Both Validators?**

- **Report Data Validator** checks data issues (missing fields, type errors, table format problems, checkbox data format)
- **Template Validator** checks template issues (split placeholders, missing config definitions, checkbox name mismatches)

The two validators check different layers of issues and complement each other. Only when both pass can you be confident that report generation will succeed. Note that checkbox validation is automatic - template checkboxes without configuration will be set to `false`.

---

## Configuration Files

### Report Config (`config/report_config.json`)

```json
{
  "template_path": "report_templates/report_template1.docx",
  "output_dir": "output/",
  "field_mappings": [
    {
      "template_field": "report_no",
      "source_field": "metadata.report_no",
      "type": "text"
    },
    {
      "template_field": "energy_class_rating",
      "source_field": "calculated_data.energy_class_rating",
      "args": ["extracted_data.rated_wattage", "extracted_data.useful_luminous_flux"],
      "function": "calculate_energy_class_rating",
      "type": "text"
    },
    {
      "template_field": "photometric_data",
      "source_field": "extracted_data.photometric_data",
      "type": "table",
      "table_template_path": "report_templates/photometric_table_template.docx",
      "row_strategy": "fixed_rows",
      "header_rows": 2,
      "skip_columns": [1],
      "target_headers": ["Current", "Power", "PF", ...],
      "transformations": [
        {"type": "skip_columns", "columns": [1]},
        {"type": "calculate", "column": 1, "operation": "average", "decimal": 2}
      ]
    },
    {
      "template_field": "images",
      "source_field": "extracted_data.images",
      "type": "image",
      "width": 4.0,
      "alignment": "center"
    },
    {
      "template_field": "containing_product",
      "source_field": "metadata.containing_product",
      "type": "checkbox"
    }
  ]
}
```

### Checkbox Configuration

**Field Mapping for Checkbox:**
```json
{
  "template_field": "checkbox_name_in_word",
  "source_field": "metadata.checkbox_field",
  "type": "checkbox"
}
```

**Data Format in report.json:**
```json
{
  "metadata": {
    "checkbox_field": {
      "type": "checkbox",
      "value": "true"
    }
  }
}
```

**Behavior:**
1. Each checkbox is an independent field_mapping entry
2. `template_field` must match the checkbox name in Word template
3. Data always uses format `{"type": "checkbox", "value": "true/false"}`
4. If `report_config` has checkbox but `report_data` doesn't have the field, defaults to `false`
5. Template checkboxes not in operations are automatically set to `false`

**Checkbox Name Extraction:**
```bash
# Extract checkboxes from template
python tools/extract_template_elements.py report_templates/production_template.docx
```

---

## Build and Test Commands

### Running Tests

```bash
# Run all tests
pytest

# Run with verbose output
pytest -v

# Run specific test file
pytest tests/test_field_mapper.py
pytest tests/test_calculator.py

# Run specific test class
pytest tests/test_field_mapper.py::TestLoadJson

# Run with coverage (if pytest-cov installed)
pytest --cov=src --cov-report=html
```

### Running the Complete Workflow

```bash
# Step 1: Validate report.json
python src/validate_report.py \
    --report config/report.json \
    --config config/report_config.json

# Step 2: Calculate derived fields
python src/calculator.py \
    --config config/report_config.json \
    --report config/report.json \
    --output output/calculated_report.json

# Step 3: Generate operations
python src/field_mapper.py \
    --config config/report_config.json \
    --report output/calculated_report.json \
    --output output/operations.json

# Step 4: Process template
python src/process_template.py \
    --template report_templates/production_template.docx \
    --operations output/operations.json \
    --calculated-report config/report_data.json \
    --output output/report.docx
```

**Note:** In the new architecture, `process_template.py` uses `--calculated-report` instead of separate `--metadata` and `--targets` arguments. Checkbox operations are automatically handled during template processing.

---

## Code Style Guidelines

1. **Language**: Code comments and docstrings use both English and Chinese.
2. **Type Hints**: Use Python type hints for function signatures.
3. **Imports**: Group imports (stdlib, third-party, local) with blank lines.
4. **Error Handling**: Use custom exception classes (e.g., `DocxTemplateError`, `CalculatorError`).
5. **Docstrings**: Use triple-quoted docstrings for modules, classes, and functions.
6. **Naming**:
   - Classes: `PascalCase`
   - Functions/variables: `snake_case`
   - Constants: `UPPER_CASE`

---

## Testing Strategy

- **Unit Tests**: Each module has corresponding test file in `tests/`
- **Mock Usage**: Heavy use of `unittest.mock` for external dependencies (Excel, file system)
- **Test Data**: Use `tmp_path` fixture for temporary files
- **Coverage**: Tests cover normal cases, edge cases, and error conditions

---

## Security Considerations

1. **eval() Usage**: `data_transformer.py` uses `eval()` for lambda functions from JSON config. Ensure config files are from trusted sources.
2. **File Paths**: Always validate file paths before operations.
3. **Word File Locks**: Processor checks for Word lock files (`~$filename.docx`) before writing.
4. **Checkbox Auto-Setting**: Template checkboxes not in `checkbox_mapping` are automatically unchecked to ensure document consistency.

---

## Error Handling

### Common Errors

1. **PlaceholderNotFoundError**
   - Placeholder not found in specified location
   - Check template for correct placeholder name

2. **InvalidLocationError**
   - Invalid location specified for operation
   - Valid locations: body, header, footer

3. **DocxTemplateError**
   - Generic processor error
   - Check error message for details

4. **FileNotFoundError**
   - Template, image, or Excel file not found
   - Verify file paths in config

---

## Common Tasks for AI Agents

### Adding a New Field Mapping

1. Add entry to `config/report_config.json` field_mappings:
   ```json
   {
     "template_field": "new_field",
     "source_field": "metadata.new_field",
     "type": "text"
   }
   ```
2. Ensure placeholder `{{new_field}}` exists in Word template
3. Run validation to verify

### Adding a Checkbox Field

1. Add checkbox entry to `config/report_config.json` field_mappings:
   ```json
   {
     "template_field": "containing_product",
     "source_field": "metadata.containing_product",
     "type": "checkbox"
   }
   ```
2. Add checkbox data to `config/report.json`:
   ```json
   {
     "metadata": {
       "containing_product": {
         "type": "checkbox",
         "value": "true"
       }
     }
   }
   ```
3. Ensure the checkbox name matches the form field name in Word template
4. Template checkboxes without configuration will be automatically set to `false`

### Adding a Custom Calculation

1. Create/edit `src/custom_calculations.py`:
   ```python
   from src.calculator import CalculationRegistry
   
   @CalculationRegistry.register("my_calculation")
   def my_calculation(arg1, arg2):
       return f"{float(arg1) + float(arg2):.2f}"
   ```
2. Add field mapping with `"function": "my_calculation"`
3. Run calculator with `--functions-module custom_calculations`

### Debugging Placeholder Issues

1. Check if placeholder exists in template:
   ```python
   from src.processor import DocxTemplateProcessor
   processor = DocxTemplateProcessor("template.docx", "output.docx")
   print(processor.get_all_placeholders())
   ```
2. Verify placeholder format: `{{placeholder_name}}` (case-sensitive)
3. Check if placeholder is in correct location (body/header/footer)

### Modifying Table Transformations

1. Edit `transformations` array in report_config.json
2. Available types: `skip_columns`, `add_column`, `calculate`, `format_column`, `reorder`, `filter_rows`
3. For complex formatting, use lambda functions:
   ```json
   {"type": "format_column", "column": 4, "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"}
   ```

### Table Text Insert Configuration

**Purpose**: Insert text at specific row and column positions in tables, useful for inserting headers or labels from data fields.

**Configuration Example**:
```json
{
  "template_field": "photometric_data_table",
  "source_field": "extracted_data.photometric_data_table",
  "type": "table",
  "table_template_path": "report_templates/tables/photometric_table_template.docx",
  "row_strategy": "fixed_rows",
  "header_rows": 2,
  "text_insert": [
    {"column": 0, "row": 2, "value": "extracted_data.light_source_type"}
  ],
  "transformations": [...]
}
```

**Field Descriptions**:
- `text_insert`: Array of text insertion configurations
  - `column`: Column index (0-based) where text will be inserted
  - `row`: Row index (0-based) where text will be inserted
  - `value`: Dot-notation path to the data field (e.g., `extracted_data.light_source_type`, `metadata.report_no`)

**Data Source Path Format**:
- `metadata.field_name` - reads from `calculated_report.metadata.field_name`
- `extracted_data.field_name` - reads from `calculated_report.extracted_data.field_name`
- `calculated_data.field_name` - reads from `calculated_report.calculated_data.field_name`

**Implementation Details**:
- Text insertion occurs after the main data filling process
- The value is resolved from `calculated_report` using the dot-notation path
- If the path cannot be resolved or returns `None`, no text is inserted
- This feature works with both `fixed_rows` and `dynamic_rows` strategies

---

## Documentation References

- **`docs/VALIDATOR_AND_CALCULATOR_GUIDE.md`** - Complete guide for using ReportDataValidator and Calculator
  - Validator CLI and Python API usage
  - Calculator configuration with `function` and `args`
  - Built-in calculation functions reference
  - Complete configuration examples

- **`docs/QUICK_REFERENCE.md`** - Quick reference card for common tasks
  - One-line commands for validation and calculation
  - Common validation errors and solutions
  - Configuration templates

- **`config/report_config_example_complete.json`** - Complete configuration example with comments

---

## Important Notes

1. **Path Handling**: Image paths are automatically resolved relative to `data_files/` directory.
2. **Excel Reading**: Uses `data_only=True` to read calculated formula values.
3. **Nested Tables**: Placeholders in nested table cells may be skipped if parent cannot be determined.
4. **File Encoding**: All JSON files use UTF-8 encoding.
5. **Windows Compatibility**: Code handles Windows path separators and console encoding.
6. **Checkbox Auto-Setting**: Template checkboxes not configured in `report_config.json` are automatically set to `false` during processing.

---

## Migration from Old Architecture

If you encounter old configuration files:
- Old: Separate `metadata.json`, `extracted_data.json`, `calculated_data.json`
- New: Single `report.json` with hierarchical structure
- Old: `"source": "metadata"` + `"source_field": "field_name"`
- New: `"source_field": "metadata.field_name"`

See `ARCHITECTURE_REFACTOR.md` for detailed migration guide.
