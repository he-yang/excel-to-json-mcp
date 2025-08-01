# Excel to JSON MCP by WTSolutions

[中文](README-zh.md)

## Introduction

The **Excel to JSON MCP** (Model Context Protocol) provides a standardized interface for converting Excel and CSV data into JSON format using the Model Context Protocol. This MCP implementation offers two specific tools for data conversion:

- **excel_to_json_mcp_from_data**: Converts tab-separated Excel data or comma-separated CSV text data into JSON format.
- **excel_to_json_mcp_from_url**: Converts Excel file (.xlsx) from a providedURL

Excel to JSON MCP is part of Excel to JSON by WTSolutions:
* [Excel to JSON Web App: Convert Excel to JSON directly in Web Browser.](https://excel-to-json.wtsolutions.cn/en/latest/WebApp.html)
* [Excel to JSON Excel Add-in: Convert Excel to JSON in Excel, works with Excel environment seamlessly.](https://excel-to-json.wtsolutions.cn/en/latest/ExcelAddIn.html)
* [Excel to JSON API: Convert Excel to JSON by HTTPS POST request.](https://excel-to-json.wtsolutions.cn/en/latest/API.html)
* <mark>Excel to JSON MCP Service: Convert Excel to JSON by AI Model MCP SSE/StreamableHTTP request.</mark> (<-- You are here.)

## Server Config

Available MCP Servers (SSE and Streamable HTTP):

### Using SSE

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
### Using Streamable HTTP

Transport: Streamable HTTP

URL: https://mcp.wtsolutions.cn/excel-to-json-mcp-shttp

Server Config JSON:

```json
{
  "mcpServers": {
    "excel_to_json_by_WTSolutions": {
      "args": [
        "mcp-remote",
        "https://mcp.wtsolutions.cn/excel-to-json-mcp-shttp"
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

> Note:
> Input data must be tab-separated (Excel) or comma-separated (CSV) text with at least two rows (header row + data row).
> - The first row will be considered as "header" row, and this MCP will use it as column names, subsequently JSON keys.
> - The following rows will be considered as "data" rows, and this MCP will treat them as JSON values.

#### Example Prompt 1:

Convert the following tab-separated data into JSON format:

```
Name	Age	IsStudent
John Doe	25	false
Jane Smith	30	true
```

#### Example Prompt 2:

Convert the following comma-separated data into JSON format:

```
Name,Age,IsStudent
John Doe,25,false
Jane Smith,30,true
```

### excel_to_json_mcp_from_url

Converts an Excel file from a provided URL into JSON format.

#### Parameters

| Parameter | Type   | Required | Description                                      |
|-----------|--------|----------|--------------------------------------------------|
| url       | string | Yes      | URL pointing to an Excel (.xlsx)                 |

> Note:
> - Each sheet of the Excel file should contain at least two rows (header row + data row).
>   1. The first row will be considered as "header" row, and this MCP will use it as column names, subsequently JSON keys.
>   2. The following rows will be considered as "data" rows, and this MCP will treat them as JSON values.
> - This Excel file should be in '.xlsx' format.
> - Each sheet of the Excel file will be converted to a JSON object.
> - Each JSON object will have 'sheetName' (string) and 'data' (array of objects) properties.
> - Each JSON object in 'data' array will have properties corresponding to column names.
> - Each JSON object in 'data' array will have values corresponding to cell values.


### Example Prompt 1

Convert Excel file to JSON, file URL: https://tools.wtsolutions.cn/example.xlsx

### Example Prompt 2
(applicable only when you do not have a URL and working with online AI LLM)

I've jsut uploaded one .xlsx file to you, please extract its URL and send it to MCP tool 'excel_to_json_mcp_from_url', for Excel to JSON conversion.


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

Above is the response from MCP tool, and in most cases your LLM should interpret the response and present you with a JSON object, for example as below. 
> Note, different LLM models may have different ways to interpret the JSON object, so please check if the JSON object is correctly interpreted by your LLM model.


```json
{
  "isError": false,
  "msg": "success",
  "data": "[{\"Name\":\"John Doe\",\"Age\":25,\"IsStudent\":false},{\"Name\":\"Jane Smith\",\"Age\":30,\"IsStudent\":true}]"
}
```

```json
{
  "isError": false,
  "msg": "success",
  "data": [
    {
      "Name": "John Doe",
      "Age": 25,
      "IsStudent": false
    },
    {
      "Name": "Jane Smith",
      "Age": 30,
      "IsStudent": true
    }
  ]
}

```

```json
[
  {
    "Name": "John Doe",
    "Age": 25,
    "IsStudent": false
  },
  {
    "Name": "Jane Smith",
    "Age": 30,
    "IsStudent": true
  }
]

```

### Example Failed Response

```json
{
  "content": [{
    "type": "text",
    "text": "{\"isError\": true, \"msg\": \"Network Error when fetching file\", \"data\": \"\"}"
  }]
}
```
Above is the response from MCP tool, and in most cases your LLM should interpret the response and present you with a JSON object, for example as below. 
> Note, different LLM models may have different ways to interpret the JSON object, so please check if the response is correctly interpreted by your LLM model.

```json
{
  "isError": true,
  "msg": "Network Error when fetching file",
  "data": ""
}
```
or it is also possbile that your LLM would say "Network Error when fetching file, try again later" to you.

## Data Type Handling

The API automatically detects and converts different data types:

- **Numbers**: Converted to numeric values
- **Booleans**: Recognizes 'true'/'false' (case-insensitive) and converts to boolean values
- **Dates**: Detects various date formats and converts them appropriately
- **Strings**: Treated as string values
- **Empty values**: Represented as empty strings

## Error Handling

The MCP returns descriptive error messages for common issues:

- `Excel Data Format Invalid`: When input data is not tab-separated or comma-separated
- `At least 2 rows are required`: When input data has fewer than 2 rows
- `Both data and url received`: When both 'data' and 'url' parameters are provided
- `Network Error when fetching file`: When there's an error downloading the file from the provided URL
- `File not found`: When the file at the provided URL cannot be found
- `Blank/Null/Empty cells in the first row not allowed`: When header row contains empty cells
- `Server Internal Error`: When an unexpected error occurs

## Service Agreement and Privacy Policy

By using Excel to JSON MCP, you agree to the [service agreement](TERMS.md), and [privacy policy](PRIVACY.md).


## Pricing

Free for now.

## Donation

[https://buymeacoffee.com/wtsolutions](https://buymeacoffee.com/wtsolutions)