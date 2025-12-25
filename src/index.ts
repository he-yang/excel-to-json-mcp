#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { processData, processURLExcel, URLs } from './functions.js'

// Create server instance
const server = new McpServer({
    name: "Excel to JSON MCP by WTSolutions",
    version: "0.3.1",
});

server.registerTool(
    "excel_to_json_mcp_from_data",
    {
        title: "Excel to JSON MCP by WTSolutions - from data",
        description: "Convert string format (1) tab separated Excel data or (2) comma separated CSV data to JSON. If you do not have a Pro Code, please pass only the data parameter, but not options parameter in the request",
        inputSchema: {
            data: z.string().nonempty().describe("Tab separated Excel data or CSV data in string format"),
            options: z.object({
                jsonMode: z.enum(['nested', 'flat']).optional().describe("Format mode for JSON output: 'nested' or 'flat'"),
                header: z.enum(['row', 'column']).optional().describe("Specifies which row/column to use as headers: 'row' (first row) or 'column' (first column)"),
                delimiter: z.enum(['.', '_', '__', '/']).optional().describe("Delimiter character for nested JSON keys when using nested jsonMode: '.', '_', '__', '/'"),
                emptyCell: z.enum(['emptyString', 'null', 'exclude']).optional().describe("Handling of empty cells: 'emptyString', 'null', or 'exclude'"),
                booleanFormat: z.enum(['trueFalse', '10', 'string']).optional().describe("Format for boolean values: 'trueFalse', '10', or 'string'"),
                jsonFormat: z.enum(['arrayOfObject', '2DArray']).optional().describe("Overall JSON output format: 'arrayOfObject' or '2DArray'"),
                singleObjectFormat: z.enum(['array', 'object']).optional().describe("Format when result has only one object: 'array' (keep as array) or 'object' (return as single object)"),
                proCode: z.string().optional().describe("Pro Code for Excel to JSON by WTSolutions, if you have one. If you do not have a Pro Code, please do not pass the options parameter in the request."),
            }).optional().describe("If you do not have a Pro Code, please do not pass the options parameter in the request."),
        }
    },
    async ({ data, options }) => {
        if (options && !options.proCode) {
            options.proCode = process.env.proCode || ''
        }
        
        let result = await processData(data, options)
        if (result.isError) {
            result['data'] = ''
            result['msg'] = result['msg'] + ' Refer to Documentation at  ' + URLs.mcp
        } else {
            result['msg'] = 'success'
        }
        // 返回JSON格式的结果
        return { content: [{ type: "text", text: JSON.stringify(result) }] };
    }
);

server.registerTool(
    "excel_to_json_mcp_from_url",
    {
        title: "Excel to JSON MCP by WTSolutions - from url",
        description: "Convert Excel (.xlsx) from publicly accessible URL(string format) to JSON. If you do not have a Pro Code, please pass only the url parameter, but not options parameter in the request.",
        inputSchema: {
            url: z.string().nonempty().describe("Publically accessible URL(string format) to Excel file(.xlsx)"),
            options: z.object({
                jsonMode: z.enum(['nested', 'flat']).optional().describe("Format mode for JSON output: 'nested' or 'flat'"),
                header: z.enum(['row', 'column']).optional().describe("Specifies which row/column to use as headers: 'row' (first row) or 'column' (first column)"),
                delimiter: z.enum(['.', '_', '__', '/']).optional().describe("Delimiter character for nested JSON keys when using nested jsonMode: '.', '_', '__', '/'"),
                emptyCell: z.enum(['emptyString', 'null', 'exclude']).optional().describe("Handling of empty cells: 'emptyString', 'null', or 'exclude'"),
                booleanFormat: z.enum(['trueFalse', '10', 'string']).optional().describe("Format for boolean values: 'trueFalse', '10', or 'string'"),
                jsonFormat: z.enum(['arrayOfObject', '2DArray']).optional().describe("Overall JSON output format: 'arrayOfObject' or '2DArray'"),
                singleObjectFormat: z.enum(['array', 'object']).optional().describe("Format when result has only one object: 'array' (keep as array) or 'object' (return as single object)"),
                proCode: z.string().optional().describe("Pro Code for Excel to JSON by WTSolutions, if you have one. If you do not have a Pro Code, please do not pass the options parameter in the request."),
            }).optional().describe("If you do not have a Pro Code, please do not pass the options parameter in the request."),
        }
    },
    async ({ url, options }) => {
        if (options && !options.proCode) {
            options.proCode = process.env.proCode || ''
        }
        let result = await processURLExcel(url, options)
        if (result.isError) {
            result['data'] = ''
            result['msg'] = result['msg'] + ' Refer to Documentation at  ' + URLs.mcp
        } else {
            result['msg'] = 'success'
        }
        return { content: [{ type: "text", text: JSON.stringify(result) }] };
    }
);

server.registerPrompt(
    "from-url",
    {
        title: "Convert Excel file(URL) to JSON",
        description: "Convert Excel file to JSON, the Excel file URL is publically avaliable. URL should point to .xlsx file and start with https.",
        argsSchema: { url: z.string() }
    },
    ({ url }) => ({
        messages: [{
            role: "user",
            content: {
                type: "text",
                text: `Convert Excel file to JSON, file URL:${url}. \n\n I do not have a pro code, so please do not pass options parameter in the request.`
            }
        }]
    })
);
server.registerPrompt(
    "from-url-upload",
    {
        title: "Convert Excel file(uploaded) to JSON",
        description: "Convert Excel file to JSON, the Excel (.xlsx) file is uploaded to AI model",
        argsSchema: {}
    },
    () => ({
        messages: [{
            role: "user",
            content: {
                type: "text",
                text: `I've just uploaded one .xlsx file to you, please extract its URL and send it to MCP tool 'excel_to_json_mcp_from_url', for Excel to JSON conversion. \n\n I do not have a pro code, so please do not pass options parameter in the request.`
            }
        }]
    })
);
server.registerPrompt(
    "from-data",
    {
        title: "Convert Excel data to JSON data",
        description: "Convert Excel data to JSON data",
        argsSchema: { data: z.string() }
    },
    ({ data }) => ({
        messages: [{
            role: "user",
            content: {
                type: "text",
                text: `Convert the following data into JSON: \n\n ${data} \n\n I do not have a pro code, so please do not pass options parameter in the request.`
            }
        }]
    })
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Excel to JSON by WTSolutions MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});




