# Excel to JSON MCP by WTSolutions

[English](README.md) 

## 介绍

**Excel to JSON MCP**（模型上下文协议）提供了一个标准化接口，用于通过模型上下文协议将Excel和CSV数据转换为JSON格式。此MCP实现提供了两个特定的数据转换工具：

- **excel_to_json_mcp_from_data**：将制表符分隔的Excel数据或逗号分隔的CSV文本数据转换为JSON格式。
- **excel_to_json_mcp_from_url**：从提供的URL转换Excel文件（.xlsx）

Excel to JSON MCP是WTSolutions的Excel to JSON系列的一部分：
* [Excel to JSON Web应用：直接在Web浏览器中转换Excel为JSON。](https://excel-to-json.wtsolutions.cn/zh-cn/latest/WebApp.html)
* [Excel to JSON Excel插件：在Excel中转换Excel为JSON，与Excel环境无缝协作。](https://excel-to-json.wtsolutions.cn/zh-cn/latest/ExcelAddIn.html)
* [Excel to JSON WPS插件：在WPS中转换Excel为JSON，与WPS环境无缝协作。](https://excel-to-json.wtsolutions.cn/zh-cn/latest/WPSAddIn.html)
* [Excel to JSON API：通过HTTPS POST请求转换Excel为JSON。](https://excel-to-json.wtsolutions.cn/zh-cn/latest/API.html)
* <mark>Excel to JSON MCP服务：通过AI模型MCP SSE/流式HTTP请求转换Excel为JSON。</mark>（<-- 您当前所在位置。）

## 服务器配置

可用的MCP服务器：

### 使用Stdio(NPX)

服务器配置 JSON:

Case 1 : 免费版

如果您使用的是免费版本：

```json
{
  "mcpServers": {
    "excel-to-json-mcp": {
      "args": [
        "excel-to-json-mcp"        
      ],
      "command": "npx"
    }
  }
}
```

Case 2 : Pro版本

如果您使用的是Pro版本（需要有有效Pro Code）：

```json
{
  "mcpServers": {
    "excel-to-json-mcp": {
      "args": [
        "excel-to-json-mcp"        
      ],
      "command": "npx",
      "env": {
        "proCode": "type in your proCode here"
      }
    }
  }
}
```

### 使用SSE
从 v0.3.0 开始，不再支持。

### 使用流式HTTP
从 v0.3.0 开始，不再支持。

## MCP工具

### excel_to_json_mcp_from_data

将制表符分隔的Excel数据或逗号分隔的CSV文本数据转换为JSON格式。

#### 参数

| 参数 | 类型 | 是否必需 | 描述 |
|-----------|--------|----------|-----------------------------------------------------------------------------|
| data | string | 是 | 制表符分隔或逗号分隔的文本数据，至少包含两行（标题行+数据行） |
| options | object | 否 | 用于自定义转换过程的可选配置对象。如果您没有Pro Code，请不要传递此参数。 |

> 注意：
> 输入数据必须是制表符分隔（Excel）或逗号分隔（CSV）的文本，至少包含两行（标题行+数据行）。
> - 第一行将被视为"标题"行，此MCP将使用它作为列名，进而作为JSON键。
> - 后续行将被视为"数据"行，此MCP将把它们视为JSON值。

#### 示例提示词1：

将以下数据转换为JSON格式：

```
Name	Age	IsStudent
John Doe	25	false
Jane Smith	30	true
```

#### 示例提示词2：

将以下数据转换为JSON格式：

```
Name,Age,IsStudent
John Doe,25,false
Jane Smith,30,true
```

### excel_to_json_mcp_from_url

将来自提供的URL的Excel文件转换为JSON格式。

#### 参数

| 参数 | 类型 | 是否必需 | 描述 |
|-----------|--------|----------|--------------------------------------------------|
| url | string | 是 | 指向Excel文件（.xlsx）的URL |
| options | object | 否 | 用于自定义转换过程的可选配置对象。如果您没有Pro Code，请仅传递`url`参数，而不要传递`options`参数。 |

> 注意：
> - Excel文件的每个工作表应至少包含两行（标题行+数据行）。
>   1. 第一行将被视为"标题"行，此MCP将使用它作为列名，进而作为JSON键。
>   2. 后续行将被视为"数据"行，此MCP将把它们视为JSON值。
> - 此Excel文件应为'.xlsx'格式。
> - Excel文件的每个工作表将被转换为一个JSON对象。
> - 每个JSON对象将具有'sheetName'（字符串）和'data'（对象数组）属性。
> - 'data'数组中的每个JSON对象将具有与列名对应的属性。
> - 'data'数组中的每个JSON对象将具有与单元格值对应的值。


### 示例提示词1

将Excel文件转换为JSON，文件URL：https://tools.wtsolutions.cn/example.xlsx

### 示例提示词2
（仅适用于没有URL且使用在线AI LLM的情况）

我刚刚向您上传了一个.xlsx文件，请提取其URL并将其发送到MCP工具'excel_to_json_mcp_from_url'，以进行Excel到JSON的转换。

## 选项对象

如果您没有Pro Code，您可以继续使用此MCP工具，只传递`url`或`data`参数，而不要传递`options`参数。但请注意，转换规则使用默认规则的限制(也就是下方表格中的默认值列)。

可选的`options`对象可以包含以下属性，用于自定义转换过程。此功能需要有效的Pro Code，查看[定价](https://excel-to-json.wtsolutions.cn/zh-cn/latest/pricing.html)。

| 属性             | 类型   | 默认值   | 描述                                                                 |
|----------------------|--------|-----------|-----------------------------------------------------------------------------|
| proCode              | string | ""        | 用于自定义转换规则的Pro Code，需要有效的Excel to JSON服务订阅。 |
| jsonMode             | string | "flat"    | JSON输出的格式模式："nested"（嵌套）或"flat"（扁平） |
| header               | string | "row"     | 指定使用哪一行/列作为标题："row"（第一行）或"column"（第一列） |
| delimiter            | string | "."       | 当使用`jsonMode`："nested"时，嵌套JSON键的分隔符，可接受的分隔符为："."、"_"、"__"、"/"。 |
| emptyCell            | string | "emptyString" | 空单元格的处理方式："emptyString"（空字符串）、"null"（空值）或"exclude"（排除） |
| booleanFormat        | string | "trueFalse" | 布尔值的格式："trueFalse"（true/false）、"10"（1/0）或"string"（字符串） |
| jsonFormat           | string | "arrayOfObject" | 整体JSON输出格式："arrayOfObject"（对象数组）或"2DArray"（二维数组） |
| singleObjectFormat   | string | "array"   | 当结果只有一个对象时的格式："array"（保持为数组）或"object"（返回单个对象） |

> 注意：
> - 只有当`jsonMode`为"nested"时，`delimiter`才会生效。
> - 只有当`jsonFormat`为"arrayOfObject"时，`singleObjectFormat`才会生效。 
> - 只有当`jsonMode`为"flat"时，`jsonFormat`为"2DArray"才会生效。
> - 如果您使用任何其他选项，则`proCode`是必需的。

## 响应格式

MCP工具返回具有以下结构的JSON对象：

| 字段 | 类型 | 描述 |
|----------|---------|-------------------------------------------------------------------------------------------------------------------------------------------|
| isError | boolean | 指示处理请求时是否出错 |
| msg | string | 'success'或错误描述 |
| data | string | 如果使用URL，则为工作表对象数组；如果使用直接数据，则为字符串；如果出错，则为空字符串。如果使用URL，每个工作表对象包含'sheetName'（字符串）和'data'（对象数组） |

### 成功响应示例

```json
{
  "content": [{
    "type": "text",
    "text": "{\"isError\":false,\"msg\":\"success\",\"data\":\"[{\"Name\":\"John Doe\",\"Age\":25,\"IsStudent\":false},{\"Name\":\"Jane Smith\",\"Age\":30,\"IsStudent\":true}]\"}"
  }]
}
```

以上是MCP工具的响应，在大多数情况下，您的LLM应该解释该响应并向您呈现一个JSON对象，例如如下所示。
> 注意，不同的LLM模型可能有不同的方式来解释JSON对象，因此请检查您的LLM模型是否正确解释了JSON对象。


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

### 失败响应示例

```json
{
  "content": [{
    "type": "text",
    "text": "{\"isError\": true, \"msg\": \"Network Error when fetching file\", \"data\": \"\"}"
  }]
}
```
以上是MCP工具的响应，在大多数情况下，您的LLM应该解释该响应并向您呈现一个JSON对象，例如如下所示。
> 注意，不同的LLM模型可能有不同的方式来解释JSON对象，因此请检查您的LLM模型是否正确解释了响应。

```json
{
  "isError": true,
  "msg": "Network Error when fetching file",
  "data": ""
}
```
或者，您的LLM也可能会向您显示"获取文件时网络错误，请稍后再试"。

## 数据类型处理

API会自动检测并转换不同的数据类型：

- **数字**：转换为数值
- **布尔值**：识别'true'/'false'（不区分大小写）并转换为布尔值
- **日期**：检测各种日期格式并适当转换
- **字符串**：视为字符串值
- **空值**：表示为空字符串

## 错误处理

MCP针对常见问题返回描述性错误消息：

- `Excel Data Format Invalid`：当输入数据不是制表符分隔或逗号分隔时
- `At least 2 rows are required`：当输入数据少于2行时
- `Both data and url received`：当同时提供'data'和'url'参数时
- `Network Error when fetching file`：从提供的URL下载文件时出错
- `File not found`：找不到提供的URL上的文件
- `Blank/Null/Empty cells in the first row not allowed`：当标题行包含空单元格时
- `Server Internal Error`：发生意外错误时

## 服务协议和隐私政策

使用Excel to JSON MCP，即表示您同意[服务协议](TERMS.md)和[隐私政策](PRIVACY.md)。


## 定价
使用默认的转换规则，免费。
使用自定义的转换规则，需要购买Pro Code，参考[定价](https://excel-to-json.wtsolutions.cn/zh-cn/latest/pricing.html)。

