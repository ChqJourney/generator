# Word Document Report Generator

## Project Overview

This project automates the generation of Word reports by inserting extracted data into pre-designed Word templates. The system supports text, table, image, and checkbox operations with placeholder-based templating.

## Architecture

```
┌─────────────────┐         ┌─────────────────┐         ┌─────────────────┐
│  report.json   │         │report_config.json│        │  Word Template   │
│  (Hierarchical │         │  (Field Mappings)│        │  .docx           │
│   Data)        │         │                  │        │                  │
└────────┬────────┘         └────────┬────────┘        └────────┬────────┘
         │                          │                          │
         │                          ▼                          │
         │               ┌──────────────────────┐              │
         │               │  field_mapper.py    │               │
         │               │  (Generate Ops)    │                │
         │               └────────┬───────────┘                │
         │                          │                          │
         ▼                          ▼                          ▼
┌──────────────────────┐  ┌──────────────────────┐  ┌──────────────────────┐
│validators(Validation)│  │  (Operation Array)  │   │  (CLI Wrapper)       │
└──────────────────────┘  └────────┬───────────┘    └──────────┬───────────┘
                                   │                           │
                                   ▼                           ▼
                          ┌──────────────────────┐
                          │  processor.py       │
                          │  (Apply Operations) │
                          └────────┬───────────┘
                                   │
                                   ▼
                          ┌──────────────────────┐
                          │  Output .docx        │
                          └──────────────────────┘
```

## Core Components

### 1. Configuration Files

#### `config/report_data.json`
Hierarchical data structure containing all report data.

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

#### `config/report_config.jsonc`
Master configuration defining how data maps to template placeholders (JSON with comments support).

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
      "template_field": "photometric_data",
      "source_field": "extracted_data.photometric_data",
      "type": "table",
      "table_template_path": "report_templates/photometric_table_template.docx",
      "row_strategy": "fixed_rows",
      "header_rows": 2
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

**Field Mapping Types:**
- **text**: Simple text replacement
- **table**: Table insertion with optional Excel data source and transformations
- **image**: Image insertion with dimensions and alignment
- **checkbox**: Form checkbox state control (checked/unchecked)

**Table-Specific Parameters:**
- `table_template_path`: Path to table template Word document
- `row_strategy`: `"fixed_rows"` (fill existing) or `"dynamic_rows"` (add/remove rows)
- `header_rows`: Number of header rows to skip
- `transformations`: Data transformation rules (skip_columns, calculate, format, etc.)

### 2. Source Scripts

#### `src/field_mapper.py`
Converts field mappings to operation array for processor.

**Key Functions:**

- `generate_operations(config, report_data) -> Dict`
  - Main function that generates operations from config and report data
  - Supports dot-notation field access (e.g., `metadata.report_no`)

- `get_value_by_path(data, path) -> Any`
  - Extracts field value using dot-notation path
  - Handles hierarchical data structure

- `build_table_data_from_excel(value, target_headers) -> List[List[str]]`
  - Builds table data from Excel source
  - Supports external data references

**Usage:**
```bash
python src/field_mapper.py \
  --config config/report_config.jsonc \
  --report config/report_data.json \
  --output operations.json
```

#### `src/processor.py`
Core engine that applies operations to Word templates.

**Classes:**

- `DocxTemplateError`: Base exception for processor errors
- `PlaceholderNotFoundError`: Raised when placeholder not found in template
- `InvalidLocationError`: Raised when invalid location specified

- `PlaceholderFinder`: Static utility class for finding placeholders
  - `find_all_placeholders_in_location(doc, placeholder, location)`: Finds all placeholder occurrences in body/header/footer
  - `replace_paragraph_with_element(paragraph, element)`: Replaces paragraph with XML element
  - Handles nested structures (tables within tables)

- `ContentInserter` (ABC): Base class for content inserters
  - `validate_location(location, valid_locations)`: Validates location parameter

- `TextInserter(ContentInserter)`: Handles text replacement
  - `insert(placeholder, value, location)`: Replaces text placeholders
  - Supports body, header, footer locations
  - Preserves formatting by replacing in runs

- `TableInserter(ContentInserter)`: Handles table insertion
  - `insert(placeholder, table_template_path, table_data, location)`: Inserts table
  - Loads table template and fills with data
  - Applies offset_x (column shift) and offset_y (row shift)
  - Skips paragraphs without parent (handles nested table cells)

- `ImageInserter(ContentInserter)`: Handles image insertion
  - `insert(placeholder, image_paths, width, height, alignment, location)`: Inserts images
  - Supports multiple images per placeholder
  - Validates dimensions (must be Length objects)
  - Handles alignment (left/center/right)

- `CheckboxInserter(ContentInserter)`: Handles form checkbox state updates
  - `insert(checkbox_mapping)`: Updates checkbox states
  - Checkbox mapping format: `{"checkbox_name": True/False}`
  - Auto-sets unchecked state for missing checkboxes

- `DocxTemplateProcessor`: Main processor class
  - `add_text(placeholder, value, location)`: Queues text text operation
  - `add_table(placeholder, table_template_path, table_data, offset_x, offset_y)`: Queues table operation
  - `add_image(placeholder, image_paths, width, height, alignment, location)`: Queues image operation
  - `add_checkboxes(checkbox_mapping)`: Queues checkbox update operation
  - `get_all_placeholders() -> List[str]`: Extracts all placeholders from template
  - `process()`: Executes all queued operations and saves document

**Key Implementation Details:**
- Placeholder format: `{{placeholder_name}}`
- Supports all header types: header, first_page_header, even_page_header
- Handles duplicate placeholders by processing all occurrences
- Skips unreplaceable placeholders (in table cells without parents)

#### `src/process_template.py`
CLI wrapper for processor.py.

Converts operations.json to processor calls with proper type conversions.

**Usage:**
```bash
python src/process_template.py \
  --template report_templates/report_template1.docx \
  --operations operations.json \
  --calculated-report config/report_data.json \
  --output output/report.docx
```

### 3. Template Files

#### `report_templates/report_template1.docx`
Main Word document template containing placeholders like `{{report_no}}`, `{{photometric_data}}`, etc.

#### `report_templates/tables/photometric_table_template.docx`
Table template with structure for photometric data.

### 4. Data Files

#### `data_files/TDS.xlsx`
Excel source file containing photometric measurement data.

## Workflow

### Quick Start

```bash
# Step 1: Validate data and template
python src/report_data_validator.py --report config/report_data.json --config config/report_config.jsonc
python src/template_validator.py --template report_templates/production_template.docx --config config/report_config.jsonc

# Step 2: Process template (checkbox handling is automatic)
python src/process_template.py \
  --template report_templates/production_template.docx \
  --operations operations.json \
  --calculated-report config/report_data.json \
  --output output/report.docx
```

**Note:** Checkbox operations are automatically handled - template checkboxes not in operations.json are set to false.

### Data Flow

1. **Data**: Single `report_data.json` with hierarchical structure (metadata, extracted_data, calculated_data)
2. **Configuration**: `report_config.json` defines field mappings and transformations
3. **Validation**: Run validators to check data and template before processing
4. **Operation Generation**: `field_mapper.py` creates operation array (including checkbox operations)
5. **Template Processing**: `processor.py` applies operations to Word template (auto-sets unchecked checkboxes to false)
6. **Output**: Generated report with filled placeholders and correct checkbox states

For detailed architecture and component documentation, see [AGENTS.md](AGENTS.md).

## Placeholders

### Format
`{{placeholder_name}}`

### Locations
- **Body**: Main document content
- **Headers**: document header, first page header, even page header
- **Footers**: document footer, first page footer, even page footer

### Supported Operations

#### Text Replacement
```json
{
  "type": "text",
  "placeholder":"report_no",
  "value": "250400343HZH-001",
  "location":"body"(default)
}
```

#### Table Insertion
```json
{
  "type": "table",
  "placeholder":"photometric_data",
  "table_template_path":"report_template/tables/photometric_data.docx",
  "table_data": [
    ["Current(A)","Power Pon(W)","Disp. Factor",...],
    ["0.4243", "95.99", "0.9858", ...],
    ["0.4256", "96.19", "0.9853", ...]
  ]
}

```

#### Image Insertion
```json
{
  "type": "image",
  "placeholder": "images",
  "image_paths": ["data_files/1.jpg", "data_files/2.jpg"],
  "width": 4.0,
  "height": null,
  "alignment": "center"
}
```

**Image Parameters:**
- `width`, `height`: Numeric values converted to `Inches` objects
- `alignment`: "left", "center", or "right"

#### Checkbox Update
```json
{
  "type": "checkbox",
  "checkbox_mapping": {
    "containing_product": true,
    "directional": false
  }
}
```

**Checkbox Behavior:**
- Template checkboxes not in `checkbox_mapping` are automatically set to `false`
- Data format in `report_data.json`: `{"type": "checkbox", "value": "true/false"}`
- Missing checkbox data defaults to `false`

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

## Dependencies

```
python-docx>=1.2.0
openpyxl>=3.0.0
lxml>=4.0.0
typing_extensions>=4.0.0
```

## Key Implementation Notes

### Placeholder Splitting
Placeholders can be split across multiple runs in Word documents. The processor uses `paragraph.text` (full paragraph text) for finding placeholders to handle this correctly.

### Nested Structures
The system handles placeholders in:
- Direct body paragraphs
- Table cells within body
- Headers and footers (all types)
- Nested table cells (skips unreplaceable ones)

### Parent Element Handling
When replacing paragraphs with tables or images, the code checks for parent elements:
- Paragraphs without parents are skipped (cannot be replaced)
- This prevents errors with nested structures

### Type Conversion
- Image dimensions: Numeric values converted to `Inches` objects
- Offsets: JSON strings converted to integers
- Table data: All values converted to strings

## Validation Tools

The system provides two validation tools to check data and templates before report generation:

```bash
# Validate report_data.json data format and content
python src/report_data_validator.py --report config/report_data.json --config config/report_config.jsonc

# Validate Word template (check placeholder splitting issues)
python src/template_validator.py --template report_templates/production_template.docx --config config/report_config.jsonc
```

For detailed validation documentation and troubleshooting, see [AGENTS.md](AGENTS.md).

## Testing

### Running Tests

```bash
# Run all tests
pytest

# Run with verbose output
pytest -v

# Run specific test file
pytest tests/test_field_mapper.py
pytest tests/test_calculator.py
```

### Test Workflow
1. Update `config/report_data.json` with test data
2. Run validators to check data and template
3. Run `field_mapper.py` to generate operations
4. Run `process_template.py` to generate output
5. Open output file and verify content

## Project Structure

```
.
├── config/
│   ├── report_data.json            # Report data (metadata, extracted_data, calculated_data)
│   └── report_config.jsonc         # Field mapping configuration (JSON with comments)
├── data_files/
│   └── TDS.xlsx              # Excel data source
├── report_templates/
│   └── *.docx                # Word templates
├── src/
│   ├── processor.py          # Core processing engine
│   ├── field_mapper.py       # Operation generator
│   ├── calculator.py         # Field calculation engine
│   ├── process_template.py   # CLI wrapper
│   ├── report_data_validator.py  # Data validator
│   ├── template_validator.py     # Template validator
│   └── table_processor/      # Table transformation module
├── tests/                    # Test suite
└── output/                   # Generated reports
```

For detailed component documentation and architecture, see [AGENTS.md](AGENTS.md).

## License

MIT License - See LICENSE file for details