# Excel Parser to JSON

A powerful and flexible TypeScript/JavaScript library to convert Excel files into JSON with advanced configuration options.

## Installation

```bash
npm install excel-parser-to-json
```

## Features

- Convert Excel files to JSON format
- Support for multiple sheets
- Configurable header rows
- Custom column mapping
- Range selection
- Buffer and file input support
- Empty lines handling
- Date parsing support

## Basic Usage

```typescript
import excelToJson from 'excel-parser-to-json';

// Basic usage with file path
const result = excelToJson({
    sourceFile: 'path/to/your/excel.xlsx'
});

// Using buffer or string source
const result = excelToJson({
    source: yourBufferOrString
});
```

## Configuration Options

### Main Configuration

The library accepts the following main configuration options:

```typescript
interface SheetConfig {
    header?: { rows: number };       // Number of header rows to skip
    range?: string;                  // Excel-style range (e.g., 'A1:D10')
    columnToKey?: {                  // Map column letters to custom keys
        [key: string]: string
    };
    includeEmptyLines?: boolean;     // Include empty rows in output
    sheetStubs?: boolean;           // Include empty cells as null
}
```

### Advanced Usage Examples

1. **Basic Configuration with Header Rows**
```typescript
const result = excelToJson({
    sourceFile: 'data.xlsx',
    header: {
        rows: 1    // Skip first row (header)
    }
});
```

2. **Custom Column Mapping**
```typescript
const result = excelToJson({
    sourceFile: 'data.xlsx',
    columnToKey: {
        'A': 'id',
        'B': 'name',
        'C': 'email',
        '*': 'defaultKey'  // Default key for unmapped columns
    }
});
```

3. **Specific Range Selection**
```typescript
const result = excelToJson({
    sourceFile: 'data.xlsx',
    range: 'A1:D10'    // Only process cells within this range
});
```

4. **Multiple Sheets with Different Configurations**
```typescript
const result = excelToJson({
    sourceFile: 'data.xlsx',
    sheets: [
        {
            name: 'Sheet1',
            header: { rows: 1 },
            columnToKey: {
                'A': 'id',
                'B': 'name'
            }
        },
        {
            name: 'Sheet2',
            range: 'A1:C10'
        },
        'Sheet3'  // Just the sheet name for default processing
    ]
});
```

5. **Using Buffer Input**
```typescript
const result = excelToJson({
    source: fs.readFileSync('data.xlsx'),
    includeEmptyLines: false
});
```

## Output Format

The output will be an object where each key is a sheet name and its value is an array of rows:

```typescript
{
    "Sheet1": [
        { "id": 1, "name": "John" },
        { "id": 2, "name": "Jane" }
    ],
    "Sheet2": [
        // ... sheet 2 data
    ]
}
```

## Advanced Features

### Header Row Variable Substitution

You can use special variables in column mapping:
- `{{columnHeader}}`: Uses the header row value
- `{{A1}}`, `{{B2}}`, etc.: Uses specific cell values

Example:
```typescript
const result = excelToJson({
    sourceFile: 'data.xlsx',
    header: { rows: 1 },
    columnToKey: {
        'A': '{{columnHeader}}',
        'B': 'user_{{A1}}'
    }
});
```

### Empty Lines Handling

Control empty line inclusion in the output:
```typescript
const result = excelToJson({
    sourceFile: 'data.xlsx',
    includeEmptyLines: true  // Include empty rows in output
});
```

## Error Handling

The library will throw an error if:
- Neither `sourceFile` nor `source` is provided
- The specified file cannot be read
- The Excel file format is invalid

## Type Support

The library is written in TypeScript and provides full type definitions for all configurations and outputs.

