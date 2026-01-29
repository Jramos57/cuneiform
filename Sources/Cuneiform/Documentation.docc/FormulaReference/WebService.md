# WebService Functions

Functions for working with external data sources, web services, and OLAP cubes.

## Overview

Cuneiform provides WebService functions that enable interaction with external data sources including web services, OLAP (Online Analytical Processing) cubes, and real-time data feeds. These functions extend spreadsheet capabilities beyond local data to include dynamic external information.

Most WebService functions require external connectivity, data sources, or specialized infrastructure that may not be available in all environments. Cuneiform implements these functions with appropriate stub implementations that return standard error values when external resources are unavailable.

## Quick Reference

### Web & Data Functions
- ``HYPERLINK`` - Create clickable hyperlinks
- ``WEBSERVICE`` - Retrieve data from web service (requires HTTP)
- ``FILTERXML`` - Extract XML data using XPath (requires XML parsing)
- ``ENCODEURL`` - URL-encode text for web requests

### Pivot Table Functions
- ``GETPIVOTDATA`` - Extract data from pivot tables

### Real-Time Data Functions
- ``RTD`` - Retrieve real-time data (requires RTD server)

### OLAP Cube Functions
- ``CUBEVALUE`` - Return aggregate value from cube
- ``CUBEMEMBER`` - Return member from cube hierarchy
- ``CUBEMEMBERPROPERTY`` - Return property of cube member

## Function Details

### CUBEVALUE

Returns an aggregated value from an OLAP cube.

**Syntax:** `CUBEVALUE(connection, [member_expression1, ...])`

**Parameters:**
- `connection`: Text string of the connection name to the cube
- `member_expression1, ...` *(optional)*: MDX expressions that specify members

**Returns:** Any - The aggregated value from the cube

**Examples:**
```swift
// Excel example with OLAP connection:
// =CUBEVALUE("Sales", "[Time].[2024]", "[Product].[Bikes]")
```

**Excel Documentation:** [CUBEVALUE function](https://support.microsoft.com/en-us/office/cubevalue-function-8733da24-26d1-4e34-9b3a-84a8f00dcbe0)

**Implementation Status:** ❌ Stub implementation (returns #N/A error)

**Notes:** Requires an OLAP cube data source connection. Returns `#N/A` error in Cuneiform as OLAP connectivity is not yet supported. OLAP cubes are multidimensional databases used for business intelligence and analytics.

---

### CUBEMEMBER

Returns a member or tuple from the cube hierarchy.

**Syntax:** `CUBEMEMBER(connection, member_expression, [caption])`

**Parameters:**
- `connection`: Text string of the connection name to the cube
- `member_expression`: MDX expression that evaluates to a unique member
- `caption` *(optional)*: Text to display instead of the member name

**Returns:** Text - The member name or specified caption

**Examples:**
```swift
// Excel example with OLAP connection:
// =CUBEMEMBER("Sales", "[Time].[2024].[Q1]", "First Quarter")
```

**Excel Documentation:** [CUBEMEMBER function](https://support.microsoft.com/en-us/office/cubemember-function-0f6a15b9-2c18-4819-ae89-e1b5c8b398ad)

**Implementation Status:** ❌ Stub implementation (returns #N/A error)

**Notes:** Requires OLAP cube connection. Returns `#N/A` error in Cuneiform. Used to validate that members exist in the cube and to return member names.

---

### CUBEMEMBERPROPERTY

Returns the value of a member property from the cube.

**Syntax:** `CUBEMEMBERPROPERTY(connection, member_expression, property)`

**Parameters:**
- `connection`: Text string of the connection name to the cube
- `member_expression`: MDX expression that evaluates to a unique member
- `property`: Name of the property to retrieve

**Returns:** Any - The property value

**Examples:**
```swift
// Excel example with OLAP connection:
// =CUBEMEMBERPROPERTY("Sales", "[Product].[Bikes]", "Color")
```

**Excel Documentation:** [CUBEMEMBERPROPERTY function](https://support.microsoft.com/en-us/office/cubememberproperty-function-001f57d6-b850-4a5e-9c46-e6b1467f2a9b)

**Implementation Status:** ❌ Stub implementation (returns #N/A error)

**Notes:** Requires OLAP cube connection. Returns `#N/A` error in Cuneiform. Member properties contain metadata about cube members such as names, keys, and custom attributes.

---

### ENCODEURL

Encodes a text string for use in a URL.

**Syntax:** `ENCODEURL(text)`

**Parameters:**
- `text`: The text to be URL-encoded

**Returns:** Text - The URL-encoded string

**Examples:**
```swift
let result1 = evaluator.evaluate("=ENCODEURL(\"hello world\")")  
// "hello%20world"

let result2 = evaluator.evaluate("=ENCODEURL(\"user@example.com\")")  
// "user%40example.com"

let result3 = evaluator.evaluate("=ENCODEURL(\"A&B=C\")")  
// "A%26B%3DC"
```

**Excel Documentation:** [ENCODEURL function](https://support.microsoft.com/en-us/office/encodeurl-function-07c7fb90-7c60-4bff-8c17-747d4338f5b8)

**Implementation Status:** ✅ Full implementation

**Notes:** Converts special characters to percent-encoded format per RFC 3986. Characters like spaces, ampersands, equals signs, and other special characters are encoded. Alphanumeric characters and `-._~` remain unchanged. Essential for constructing valid URLs in formulas.

**Use Cases:**
- Building query strings for web APIs
- Constructing URLs dynamically
- Preparing text for use with WEBSERVICE function

---

### FILTERXML

Extracts specific data from XML content using XPath expressions.

**Syntax:** `FILTERXML(xml, xpath)`

**Parameters:**
- `xml`: A text string containing valid XML
- `xpath`: XPath expression to extract data

**Returns:** Any - The extracted data (may be array for multiple matches)

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =FILTERXML("<root><item>A</item><item>B</item></root>", "//item")
// Returns: {"A"; "B"}

// =FILTERXML(A1, "//price")  // Extract all price nodes
```

**Excel Documentation:** [FILTERXML function](https://support.microsoft.com/en-us/office/filterxml-function-4df72efc-11ec-4951-86f5-c1374812f5b7)

**Implementation Status:** ❌ Not implemented (returns #N/A error)

**Notes:** Requires XML parsing and XPath evaluation capabilities. Returns `#N/A` error in Cuneiform as external XML processing is not yet supported. Useful for extracting data from web service responses, configuration files, and structured documents.

**Common Use Cases:**
- Parsing XML responses from WEBSERVICE
- Extracting data from RSS feeds
- Processing XML configuration data

---

### GETPIVOTDATA

Extracts data stored in a PivotTable report.

**Syntax:** `GETPIVOTDATA(data_field, pivot_table, [field1, item1, ...])`

**Parameters:**
- `data_field`: Name of the data field to extract
- `pivot_table`: Reference to any cell in the PivotTable
- `field1, item1, ...` *(optional)*: Pairs of field names and items to filter by

**Returns:** Any - The value from the PivotTable

**Examples:**
```swift
// Excel example with PivotTable:
// =GETPIVOTDATA("Sales", A3)
// =GETPIVOTDATA("Sales", A3, "Region", "East", "Product", "Bikes")
```

**Excel Documentation:** [GETPIVOTDATA function](https://support.microsoft.com/en-us/office/getpivotdata-function-8c083b99-a922-4ca0-af5e-3af55960761f)

**Implementation Status:** ❌ Stub implementation (returns #REF! error)

**Notes:** Requires access to PivotTable structures. Returns `#REF!` error in Cuneiform as PivotTable data extraction is not yet fully supported. More reliable than cell references as it continues to work even when the PivotTable layout changes.

**Advantages:**
- References remain valid when PivotTable is rearranged
- More precise than cell references
- Can filter by multiple fields

---

### HYPERLINK

Creates a clickable hyperlink in a cell.

**Syntax:** `HYPERLINK(link_location, [friendly_name])`

**Parameters:**
- `link_location`: Path to document or web page URL
- `friendly_name` *(optional)*: Text to display in cell (defaults to link_location)

**Returns:** Text - The friendly name (link functionality depends on spreadsheet application)

**Examples:**
```swift
let result1 = evaluator.evaluate("=HYPERLINK(\"https://example.com\", \"Visit Site\")")
// Returns: "Visit Site"

let result2 = evaluator.evaluate("=HYPERLINK(\"Sheet2!A1\", \"Go to Sheet2\")")
// Returns: "Go to Sheet2"

let result3 = evaluator.evaluate("=HYPERLINK(\"mailto:user@example.com\", \"Email Us\")")
// Returns: "Email Us"
```

**Excel Documentation:** [HYPERLINK function](https://support.microsoft.com/en-us/office/hyperlink-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f)

**Implementation Status:** ✅ Full implementation (returns friendly_name)

**Notes:** In Cuneiform's formula evaluator, the function returns the friendly name as text. The actual hyperlink behavior (clickability and navigation) is a presentation feature that depends on the spreadsheet application displaying the workbook. The link location can be:
- External URLs (http://, https://, ftp://)
- Internal cell references (Sheet1!A1)
- Email addresses (mailto:)
- File paths

**Use Cases:**
- Creating navigation menus within workbooks
- Linking to external documentation
- Building interactive dashboards
- Generating email links

---

### RTD

Retrieves real-time data from a program that supports COM automation.

**Syntax:** `RTD(progID, server, topic1, [topic2, ...])`

**Parameters:**
- `progID`: Program ID of the RTD server
- `server`: Name of the server where RTD server is running (empty string for local)
- `topic1, topic2, ...`: Topics that define what data to retrieve

**Returns:** Any - Real-time data from the server

**Examples:**
```swift
// Excel example with RTD server:
// =RTD("myrtd.server", "", "MSFT", "LastPrice")
// =RTD("Reuters.RTD", "", "EUR", "BID")
```

**Excel Documentation:** [RTD function](https://support.microsoft.com/en-us/office/rtd-function-e0cc001a-56f0-470a-9b19-9455dc0eb593)

**Implementation Status:** ❌ Stub implementation (returns #N/A error)

**Notes:** Requires RTD (RealTimeData) COM server. Returns `#N/A` error in Cuneiform as real-time data connectivity is not supported. RTD is primarily used for live financial market data, sensor readings, and other continuously updating data streams.

**Typical Use Cases:**
- Stock prices and financial market data
- Sensor readings and IoT data
- Live system monitoring metrics
- Real-time sports scores

**Advantages over DDE:**
- More efficient update mechanism
- Better performance with multiple cells
- More reliable connection management

---

### WEBSERVICE

Retrieves data from a web service on the Internet or intranet.

**Syntax:** `WEBSERVICE(url)`

**Parameters:**
- `url`: The URL of the web service as text

**Returns:** Text - The data returned by the web service

**Examples:**
```swift
// Excel example (not supported in Cuneiform):
// =WEBSERVICE("https://api.example.com/data")
// =WEBSERVICE("https://api.weather.com/forecast?city="&ENCODEURL(A1))
```

**Excel Documentation:** [WEBSERVICE function](https://support.microsoft.com/en-us/office/webservice-function-0546a35a-ecc6-4739-aed7-c0b7ce1562c4)

**Implementation Status:** ❌ Not implemented (returns #N/A error)

**Notes:** Requires HTTP client capabilities and external network access. Returns `#N/A` error in Cuneiform as external web service calls are not supported. The function makes GET requests to the specified URL and returns the response as text.

**Common Use Cases:**
- Fetching data from REST APIs
- Retrieving JSON or XML data
- Accessing public data feeds
- Integrating with web services

**Best Practices:**
- Use ENCODEURL to encode URL parameters
- Combine with FILTERXML to parse XML responses
- Cache results to avoid excessive API calls
- Handle rate limiting from external services

**Example Workflow:**
```swift
// Excel formula pattern (illustrative):
// Step 1: Encode the query parameter
// =ENCODEURL(A1)

// Step 2: Build the URL
// ="https://api.example.com/search?q="&B1

// Step 3: Call the web service
// =WEBSERVICE(C1)

// Step 4: Parse XML response
// =FILTERXML(D1, "//result/title")
```

---

## Implementation Notes

### Stub Functions

Several WebService functions return error values because they require external resources:

| Function | Error | Reason |
|----------|-------|--------|
| `WEBSERVICE` | `#N/A` | Requires HTTP client and network access |
| `FILTERXML` | `#N/A` | Requires XML parser and XPath evaluator |
| `RTD` | `#N/A` | Requires COM automation and RTD server |
| `GETPIVOTDATA` | `#REF!` | Requires full PivotTable data structures |
| `CUBEVALUE` | `#N/A` | Requires OLAP cube connection |
| `CUBEMEMBER` | `#N/A` | Requires OLAP cube connection |
| `CUBEMEMBERPROPERTY` | `#N/A` | Requires OLAP cube connection |

These stub implementations ensure formula compatibility - spreadsheets using these functions can be read and written by Cuneiform, but the functions will return appropriate error values rather than actual data.

### Supported Functions

Two WebService functions are fully implemented:

- **HYPERLINK**: Returns the friendly name text. The actual hyperlink behavior depends on the spreadsheet application.
- **ENCODEURL**: Full URL encoding implementation using RFC 3986 percent-encoding.

### OLAP Cube Functions

OLAP (Online Analytical Processing) cube functions (`CUBEVALUE`, `CUBEMEMBER`, `CUBEMEMBERPROPERTY`, and related functions) are enterprise business intelligence features that:

- Connect to multidimensional databases
- Use MDX (Multidimensional Expressions) query language
- Provide fast aggregation and analysis of large datasets
- Support drill-down, slice-and-dice operations

These functions are typically used in corporate environments with dedicated OLAP servers like Microsoft SQL Server Analysis Services (SSAS).

### Real-Time Data

The RTD function provides real-time data updates through a COM automation interface. This Windows-specific technology allows Excel to receive continuous updates from external programs without polling. Common RTD servers include:

- Financial market data providers (Reuters, Bloomberg)
- Custom data feeds
- IoT sensor networks
- Live monitoring systems

### Web Service Integration

While `WEBSERVICE` and `FILTERXML` are not currently implemented in Cuneiform, these functions enable powerful web integration scenarios in Excel:

1. **Data Retrieval**: Fetch data from REST APIs
2. **XML Parsing**: Extract specific data from XML responses
3. **Dynamic Updates**: Create formulas that update from external sources
4. **Integration**: Connect spreadsheets with external systems

## Alternative Approaches

When WebService functions are not available, consider these alternatives:

### For Web Data
- Pre-fetch data and import into workbook
- Use external scripts to populate cells
- Export data to CSV/JSON and import

### For Real-Time Data
- Use scheduled data refreshes
- Implement polling mechanisms
- Create data import workflows

### For OLAP Cubes
- Export cube data to flat files
- Use SQL queries against relational sources
- Generate reports outside the spreadsheet

### For Pivot Data
- Use direct cell references (with caution)
- Export pivot data to separate sheet
- Rebuild calculations outside pivot

## Security Considerations

Functions that access external resources pose security risks:

- **WEBSERVICE**: Can send data to external servers, potential for data exfiltration
- **RTD**: Executes external code through COM automation
- **HYPERLINK**: Can link to potentially malicious sites or files

When implementing or using these functions:

1. Validate and sanitize URLs
2. Restrict to allowed domains/protocols
3. Implement rate limiting
4. Log external access attempts
5. Require user confirmation for external requests
6. Use HTTPS for web requests

## See Also

- ``TextFunctions`` - Text manipulation including URL building
- ``LookupFunctions`` - Data lookup and references
- ``InformationFunctions`` - Type checking and cell information
- ``CompatibilityFunctions`` - Legacy function compatibility
- <doc:FormulaEvaluator> - Formula evaluation engine
