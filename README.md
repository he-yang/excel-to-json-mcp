# Excel to JSON MCP by WTSolutions Documentation

## Introduction

The Excel to JSON MCP (Model Context Protocol) provides a standardized interface for converting Excel and CSV data into JSON format using the Model Context Protocol. This MCP implementation offers two specific tools for data conversion:

- **excel_to_json_mcp_from_data**: Converts tab-separated or comma-separated text data
- **excel_to_json_mcp_from_url**: Converts Excel file (.xlsx) from a provided URL

## Server Config

Transport: SSE
URL: https://mcp.wtsolutions.cn/excel-to-json-mcp-sse

Server Config JSON:
```json
{
  "mcpServers": {
    "excel_to_json_by_WTSolutions": {
      "args": [
        "mcp-remote",
        "https://mcp.wtsolutions.cn/excel-to-json-mcp-sse"
      ],
      "command": "npx",
      "tools": [
        "excel_to_json_mcp_from_data",
        "excel_to_json_mcp_from_url"
      ]
    }
  }
}

```



## MCP Tools

### excel_to_json_mcp_from_data

Converts tab-separated Excel data or comma-separated CSV text data into JSON format.

#### Parameters

| Parameter | Type   | Required | Description                                                                 |
|-----------|--------|----------|-----------------------------------------------------------------------------|
| data      | string | Yes      | Tab-separated or comma-separated text data with at least two rows (header row + data row) |

#### Example Request

```json
{
  "tool": "excel_to_json_mcp_from_data",
  "parameters": {
    "data": "Name\tAge\tIsStudent\nJohn Doe\t25\tfalse\nJane Smith\t30\ttrue"
  }
}
```

### excel_to_json_mcp_from_url

Converts an Excel or CSV file from a provided URL into JSON format.

#### Parameters

| Parameter | Type   | Required | Description                                      |
|-----------|--------|----------|--------------------------------------------------|
| url       | string | Yes      | URL pointing to an Excel (.xlsx)                 |

#### Example Request

```json
{
  "tool": "excel_to_json_mcp_from_url",
  "parameters": {
    "url": "https://example.com/path/to/your/file.xlsx"
  }
}
```

## Response Format

The MCP tools return a JSON object with the following structure:

| Field    | Type    | Description                                                                                                                               |
|----------|---------|-------------------------------------------------------------------------------------------------------------------------------------------|
| isError  | boolean | Indicates if there was an error processing the request                                                                                    |
| msg      | string  | 'success' or error description                                                                                                            |
| data     | string  | Converted data as array of sheet objects if using URL, string if using direct data, '' if there was an error. Each sheet object contains 'sheetName' (string) and 'data' (array of objects) if using URL |

### Example Success Response

```json
{
  "content": [{
    "type": "text",
    "text": "{\"isError\":false,\"msg\":\"success\",\"data\":\"[{\"Name\":\"John Doe\",\"Age\":25,\"IsStudent\":false},{\"Name\":\"Jane Smith\",\"Age\":30,\"IsStudent\":true}]\"}"
  }]
}
```

## Data Type Handling

The API automatically detects and converts different data types:

- **Numbers**: Converted to numeric values
- **Booleans**: Recognizes 'true'/'false' (case-insensitive) and converts to boolean values
- **Dates**: Detects various date formats and converts them appropriately
- **Strings**: Treated as string values
- **Empty values**: Represented as empty strings

## Requirements on data and url

### excel_to_json_mcp_from_data

- Input data must be tab-separated or comma-separated text with at least two rows (header row + data row).
  1. The first row will be considered as "header" row, and this API will use it as column names, subsequently JSON keys.
  2. The following rows will be considered as "data" rows, and this API will use them as JSON values.

### excel_to_json_mcp_from_url

- Each sheet of the Excel file should contain at least two rows (header row + data row).
  1. The first row will be considered as "header" row, and this API will use it as column names, subsequently JSON keys.
  2. The following rows will be considered as "data" rows, and this API will use them as JSON values.
- This Excel file should be in '.xlsx' format.
- Each sheet of the Excel file will be converted to a JSON object.
- Each JSON object will have 'sheetName' (string) and 'data' (array of objects) properties.
- Each JSON object in 'data' array will have properties corresponding to column names.
- Each JSON object in 'data' array will have values corresponding to cell values.

## Error Handling

The API returns descriptive error messages for common issues:

- `Excel Data Format Invalid`: When input data is not tab-separated or comma-separated
- `At least 2 rows are required`: When input data has fewer than 2 rows
- `Both data and url received`: When both 'data' and 'url' parameters are provided
- `Network Error when fetching file`: When there's an error downloading the file from the provided URL
- `File not found`: When the file at the provided URL cannot be found
- `Blank/Null/Empty cells in the first row not allowed`: When header row contains empty cells
- `Server Internal Error`: When an unexpected error occurs

## Pricing

Free for now.

## Donation

[https://buymeacoffee.com/wtsolutions](https://buymeacoffee.com/wtsolutions)