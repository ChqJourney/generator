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
│   ├── report_example.json         # Example report data structure
│   ├── data.json                   # Sample extracted product data
│   ├── metadata.json               # Sample metadata
│   └── table_processor_config_example.json  # Table transformation examples
├── data_files/                     # Data sources
│   └── TDS.xlsx                   # Excel test data source
├── report_templates/               # Word templates
│   ├── report_template1.docx      # Main report template
│   ├── photometric_table_template.docx
│   ├── life_test_table_template.docx
│   ├── luminous_flux_table_template.docx
│   └── config.xlsx
├── src/                            # Source code
│   ├── processor.py               # Core template processing engine
│   ├── field_mapper.py            # Field mapping to operations
│   ├── calculator.py              # Field value calculation with registry
│   ├── process_template.py        # CLI wrapper for processor
│   ├── validate_report.py         # Report JSON validator
│   ├── update_checkboxes.py       # Checkbox state updater
│   ├── custom_calculations_example.py  # Custom calculation examples
│   └── table_processor/           # Table processing module
│       ├── __init__.py
│       ├── data_transformer.py    # Data transformation rules
│       └── table_inserter.py      # Enhanced table inserter
├── tests/                          # Test suite
│   ├── test_field_mapper.py       # Unit tests for field_mapper
│   └── test_calculator.py         # Unit tests for calculator
├── output/                         # Generated reports (output directory)
├── README.md                       # Human-readable documentation
├── ARCHITECTURE_REFACTOR.md        # Architecture migration guide
├── TABLE_PROCESSOR_SUMMARY.md      # Table processor documentation
├── pyproject.toml                  # Project configuration
├── pytest.ini                     # Test configuration
└── requirements.txt               # Dependencies
```

---

## Architecture and Data Flow

### New Architecture (Post-Refactor)

```
report.json (hierarchical data)
    ↓
calculator.py → calculated_report.json (adds calculated_data)
    ↓
field_mapper.py → operations.json
    ↓
process_template.py → processor.py → Output .docx
```

### Data Structure Hierarchy

The system uses a three-tier hierarchical data structure:

```json
{
  "metadata": {
    "report_no": "RPT-001",
    "issue_date": "2024-03-15",
    "applicant_name": "..."
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

**Placeholder Format**: `{{placeholder_name}}`

**Supported Locations**: `body`, `header`, `footer` (including first_page_header, even_page_header, etc.)

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

### 5. Report Validator (`src/validate_report.py`)

**Purpose**: Validate report.json format before processing.

**CLI Usage**:
```bash
python src/validate_report.py --report config/report.json --config config/report_config.json
```

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
    }
  ]
}
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
    --template report_templates/report_template1.docx \
    --operations output/operations.json \
    --metadata config/metadata.json \
    --targets config/data.json \
    --output output/final_report.docx
```

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

---

## Important Notes

1. **Path Handling**: Image paths are automatically resolved relative to `data_files/` directory.
2. **Excel Reading**: Uses `data_only=True` to read calculated formula values.
3. **Nested Tables**: Placeholders in nested table cells may be skipped if parent cannot be determined.
4. **File Encoding**: All JSON files use UTF-8 encoding.
5. **Windows Compatibility**: Code handles Windows path separators and console encoding.

---

## Migration from Old Architecture

If you encounter old configuration files:
- Old: Separate `metadata.json`, `extracted_data.json`, `calculated_data.json`
- New: Single `report.json` with hierarchical structure
- Old: `"source": "metadata"` + `"source_field": "field_name"`
- New: `"source_field": "metadata.field_name"`

See `ARCHITECTURE_REFACTOR.md` for detailed migration guide.
