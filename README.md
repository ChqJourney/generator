# Word Document Report Generator

## Project Overview

This project automates the generation of Word reports by inserting extracted data into pre-designed Word templates. The system supports text, table, and image insertion with placeholder-based templating.

## Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│  data.json     │    │  metadata.json   │    │report_config.json│   │  Word Template │
│  (Product Data)│    │  (Report Meta)  │    │  (Config)       │   │  .docx         │
└────────┬────────┘    └────────┬┬─────────┘    └────────┬────────┘    └────────┬────────┘
         │                     │  │                    │                     │
         └──────────────────────┘  └────────────────────┘                     │
                               │                                          │
                               ▼                                          ▼
                    ┌──────────────────────┐                  ┌─────────────────┐
                    │  field_mapper.py    │                  │ process_template│
                    │  (Generate Ops)    │◄─────────────────│  .py          │
                    └────────┬───────────┘                  └─────────────────┘
                               │
                               ▼
                    ┌──────────────────────┐
                    │  operations.json    │
                    │  (Op Array)       │
                    └────────┬───────────┘
                               │
                               ▼
                    ┌──────────────────────┐
                    │  processor.py       │
                    │  (Apply Ops)       │
                    └────────┬───────────┘
                               │
                               ▼
                    ┌──────────────────────┐
                    │  Output .docx      │
                    └──────────────────────┘
```

## Core Components

### 1. Configuration Files

#### `config/data.json`
Contains extracted product data from various sources (AI, OCR, etc.).

```json
{
  "targets": [
    {
      "name": "photometric_data",
      "type": "table",
      "value": {
        "source_id": "TDS.xlsx|100W-5000K",
        "start_row": 7,
        "mapping": [...]
      }
    },
    {
      "name": "model_identifier",
      "type": "str",
      "value": "230V-100W 5000K"
    }
  ]
}
```

**Key Fields:**
- `name`: Field identifier
- `type`: Data type (str, table, images)
- `value`: Field value (can be simple value, array, or complex object with `source_id`)

#### `config/metadata.json`
Report metadata like report number, issue date.

```json
{
  "fields": [
    {
      "name": "report_no",
      "description": "报告编号",
      "value": "250400343HZH-001"
    },
    {
      "name": "issue_date",
      "description": "发布日期",
      "default": "current_date"
    }
  ]
}
```

#### `config/report_config.json`
Master configuration defining how data maps to template placeholders.

```json
{
  "template_path": "report_templates/report_template1.docx",
  "output": "output/",
  "field_mappings": [
    {
      "template_field": "report_no",
      "source": "metadata",
      "source_field": "report_no",
      "type": "text"
    },
    {
      "template_field": "photometric_data",
      "source": "extracted_data",
      "source_field": "photometric_data",
      "type": "table",
      "headers": [...],
      "table_template_path": "report_templates/tables/photometric_table_template.docx",
      "offset_x": 3,
      "offset_y": 3
    },
    {
      "template_field": "images",
      "source": "extracted_data",
      "source_field": "images",
      "type": "image",
      "width": 4.0,
      "alignment": "center"
    }
  ]
}
```

**Field Mapping Types:**
- **text**: Simple text replacement
- **table**: Complex table insertion with Excel data source
- **image**: Image insertion with dimensions and alignment

**Table-Specific Parameters:**
- `headers`: Column definitions with `name`, `type` (copy/calculate), `digit` (precision)
- `table_template_path`: Path to table template Word document
- `offset_x`, `offset_y`: Row/column offset for data insertion (positive = right/down shift)
- `source_id`: Excel data source in format `"filename|sheetname"`

### 2. Source Scripts

#### `src/field_mapper.py`
Converts field mappings to operation array for processor.

**Key Functions:**

- `generate_operations(config, metadata, extracted_data) -> Dict`
  - Main function that generates operations from config and data
  - Iterates through `field_mappings` and creates operations for each

- `extract_field_value(data, field_name) -> Any`
  - Extracts field value from extracted_data
  - Handles both dict and list structures

- `build_table_data(value, headers) -> List[List[str]]`
  - Builds table data from Excel source
  - Calls `get_xlsx_to_list` for data extraction

- `get_xlsx_to_list(file_path, sheet_name, start_row, mapping, headers) -> List[List[str]]`
  - Reads Excel file and extracts specified columns
  - Applies column mapping based on header configuration
  - Handles header normalization and missing columns

- `check_type(data) -> bool`
  - Validates if data is a proper 2D string array

**Usage:**
```bash
python src/field_mapper.py \
  --config config/report_config.json \
  --metadata config/metadata.json \
  --extracted_data config/data.json \
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
  - `insert(placeholder, table_template_path, table_data, offset_x, offset_y, location)`: Inserts table
  - Loads table template and fills with data
  - Applies offset_x (column shift) and offset_y (row shift)
  - Skips paragraphs without parent (handles nested table cells)

- `ImageInserter(ContentInserter)`: Handles image insertion
  - `insert(placeholder, image_paths, width, height, alignment, location)`: Inserts images
  - Supports multiple images per placeholder
  - Validates dimensions (must be Length objects)
  - Handles alignment (left/center/right)

- `DocxTemplateProcessor`: Main processor class
  - `add_text(placeholder, value, location)`: Queues text text operation
  - `add_table(placeholder, table_template_path, table_data, offset_x, offset_y)`: Queues table operation
  - `add_image(placeholder, image_paths, width, height, alignment, location)`: Queues image operation
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

**Key Features:**
- Converts numeric width/height to `Inches` objects
- Passes offset_x, offset_y as integers
- Validates template and operations file existence

**Usage:**
```bash
python src/process_template.py \
  --template report_templates/report_template1.docx \
  --operations operations.json \
  --output output/report.docx
```

#### `src/validate_placeholders.py`
Validation tool for Word document templates.

**Validations:**
1. Checks if placeholders are in correct locations (paragraph vs table)
2. Verifies all field_mappings exist in template
3. Identifies extra placeholders not in config

**Features:**
- Searches body, headers (all 15 types), and footers
- Identifies placeholder context (paragraph/table)
- Reports missing and extra placeholders
- Supports custom template path via `--template` parameter

**Usage:**
```bash
python src/validate_placeholders.py --config config/report_config.json
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

### Complete Workflow

```bash
# Step 1: Generate operations from config and data
python src/field_mapper.py \
  --config config/report_config.json \
  --metadata config/metadata.json \
  --extracted_data config/data.json \
  --output operations.json

# Step 2: (Optional) Validate template
python src/validate_placeholders.py --config config/report_config.json

# Step 3: Process template with operations
python src/process_template.py \
  --template report_templates/report_template1.docx \
  --operations operations.json \
  --output output/report.docx
```

### Data Flow

1. **Data Extraction**: `data.json` contains extracted product data (from AI/OCR)
2. **Configuration**: `report_config.json` defines how data maps to template
3. **Operation Generation**: `field_mapper.py` creates operation array
4. **Template Processing**: `processor.py` applies operations to Word template
5. **Output**: Generated report with filled placeholders

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
  "placeholder": "report_no",
  "value": "250400343HZH-001"
}
```

#### Table Insertion
```json
{
  "type": "table",
  "placeholder": "photometric_data",
  "table_template_path": "report_templates/tables/photometric_table_template.docx",
  "table_data": [
    ["0.4243", "95.99", "0.9858", ...],
    ["0.4256", "96.19", "0.9853", ...]
  ],
  "offset_x": 3,
  "offset_y": 3
}
```

**Table Offsets:**
- `offset_x`: Columns to skip from left (positive = right shift)
- `offset_y`: Rows to skip from top (positive = down shift)
- Data starts filling at template cell `[offset_y][offset_x]`

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

### Excel Data Extraction
- Uses `openpyxl` with `read_only=True` for performance
- `data_only=True` reads calculated formula values
- Headers are normalized (trim whitespace, collapse multiple spaces)

### Type Conversion
- Image dimensions: Numeric values converted to `Inches` objects
- Offsets: JSON strings converted to integers
- Table data: All values converted to strings

## Testing

### Validation Tool
```bash
python src/validate_placeholders.py --config config/report_config.json
```

Expected output:
- Lists all placeholder locations
- Validates placeholder context (paragraph vs table)
- Checks for missing placeholders in config
- Identifies extra placeholders not in config

### Test Workflow
1. Update `config/data.json` with test data
2. Run `field_mapper.py` to generate operations
3. Run `validate_placeholders.py` to verify template
4. Run `process_template.py` to generate output
5. Open output file and verify content

## AI Agent Quick Start

### For AI Agents Processing This Project

1. **Understand Data Flow**: Review config files to understand data sources and mappings

2. **Check Template Validation**: Run validation to identify issues before processing

3. **Generate Operations**: Use `field_mapper.py` to create operation array

4. **Process Template**: Use `processor.py` (via `process_template.py`) to apply operations

5. **Handle Errors**: Check error messages for common issues:
   - Missing placeholders in template
   - Invalid file paths
   - Incorrect data types

### Common Tasks

**Add New Field Mapping:**
1. Add field to `config/report_config.json` field_mappings
2. Ensure placeholder exists in Word template
3. Run validation to verify

**Modify Table Offsets:**
1. Update `offset_x` and `offset_y` in config
2. Positive values shift data right/down
3. Test with sample data

**Debug Placeholder Issues:**
1. Run `validate_placeholders.py`
2. Check output for missing/extra placeholders
3. Verify placeholder names match exactly (case-sensitive)

## Project Structure

```
.
├── config/
│   ├── data.json              # Extracted product data
│   ├── metadata.json           # Report metadata
│   └── report_config.json     # Master configuration
├── data_files/
│   └── TDS.xlsx              # Excel data source
├── report_templates/
│   ├── report_template1.docx   # Main template
│   └── tables/
│       └── photometric_table_template.docx  # Table template
├── src/
│   ├── field_mapper.py         # Operation generator
│   ├── processor.py            # Core processor engine
│   ├── process_template.py      # CLI wrapper
│   └── validate_placeholders.py  # Validation tool
└── output/                    # Generated reports
```

## License

MIT License - See LICENSE file for details