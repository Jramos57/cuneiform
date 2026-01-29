# Text Functions

Text manipulation, formatting, and string operations.

## Overview

Cuneiform provides 40 text functions for manipulating strings, searching text, converting case, formatting numbers as text, and working with Unicode characters. These functions are essential for data cleaning, text analysis, and report formatting.

All text functions are Excel-compatible and support Unicode text operations. Functions marked with "B" variants (LENB, LEFTB, RIGHTB, MIDB, FINDB, SEARCHB) are byte-oriented versions that treat multi-byte characters specially.

### Function Categories

**Text Extraction**:
- ``LEFT`` - Extract leftmost characters
- ``RIGHT`` - Extract rightmost characters  
- ``MID`` - Extract characters from middle
- ``TEXTBEFORE`` - Extract text before delimiter (Excel 365)
- ``TEXTAFTER`` - Extract text after delimiter (Excel 365)

**Text Information**:
- ``LEN`` - Count characters in text
- ``EXACT`` - Compare text with case sensitivity
- ``FIND`` - Find text position (case-sensitive)
- ``SEARCH`` - Find text position (case-insensitive)

**Case Conversion**:
- ``UPPER`` - Convert to uppercase
- ``LOWER`` - Convert to lowercase
- ``PROPER`` - Convert to title case

**Text Cleaning**:
- ``TRIM`` - Remove extra spaces
- ``CLEAN`` - Remove non-printable characters

**Text Combination**:
- ``CONCAT`` - Concatenate text values
- ``CONCATENATE`` - Join multiple text strings
- ``TEXTJOIN`` - Join text with delimiter (Excel 365)

**Text Modification**:
- ``REPLACE`` - Replace characters by position
- ``SUBSTITUTE`` - Replace text by value
- ``REPT`` - Repeat text multiple times

**Text to Number Conversion**:
- ``VALUE`` - Convert text to number
- ``NUMBERVALUE`` - Convert with locale settings

**Number to Text Formatting**:
- ``TEXT`` - Format number with custom pattern
- ``FIXED`` - Format with fixed decimals
- ``DOLLAR`` - Format as currency

**Character Codes**:
- ``CHAR`` - Get character from ASCII code
- ``CODE`` - Get ASCII code from character
- ``UNICODE`` - Get Unicode code point
- ``UNICHAR`` - Get character from Unicode code

**Advanced Text Operations** (Excel 365):
- ``TEXTSPLIT`` - Split text into array
- ``ARRAYTOTEXT`` - Convert array to text
- ``VALUETOTEXT`` - Convert value to text

**Special Functions**:
- ``T`` - Return text or empty string
- ``BAHTTEXT`` - Convert to Thai Baht text (stub)

## Detailed Function Reference

### LEN

Returns the number of characters in a text string.

**Syntax:** `LEN(text)`

**Parameters:**
- `text`: The text string to measure

**Returns:** Number of characters

**Examples:**
```swift
let evaluator = FormulaEvaluator(cellResolver: resolver)
let result = try evaluator.evaluate(.functionCall("LEN", [.string("Hello")]))
// Returns: .number(5.0)

// Using in formula
sheet.writeFormula("LEN(A1)", to: "B1")  // Count characters in A1
sheet.writeFormula("IF(LEN(A1)>10, \"Long\", \"Short\")", to: "B2")
```

**Excel Documentation:** [LEN function](https://support.microsoft.com/en-us/office/len-function-29236f94-cedc-429d-affd-b5e33d2c67cb)

**Implementation Status:** ✅ Full implementation

**See Also:** ``LENB``, ``CODE``

---

### LENB

Returns the number of bytes used to represent characters in a text string.

**Syntax:** `LENB(text)`

**Parameters:**
- `text`: The text string to measure

**Returns:** Number of bytes

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("LENB", [.string("Hello")]))
// Returns: .number(5.0)  // In single-byte encoding
```

**Excel Documentation:** [LENB function](https://support.microsoft.com/en-us/office/lenb-function-29236f94-cedc-429d-affd-b5e33d2c67cb)

**Implementation Status:** ✅ Full implementation (delegates to LEN)

**Note:** In Cuneiform, LENB behaves identically to LEN as Swift strings are Unicode-based.

**See Also:** ``LEN``

---

### UPPER

Converts all letters in a text string to uppercase.

**Syntax:** `UPPER(text)`

**Parameters:**
- `text`: The text to convert to uppercase

**Returns:** Uppercase text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("UPPER", [.string("hello world")]))
// Returns: .string("HELLO WORLD")

// Using in formula
sheet.writeFormula("UPPER(A1)", to: "B1")
sheet.writeFormula("UPPER(\"john doe\")", to: "B2")  // "JOHN DOE"
```

**Excel Documentation:** [UPPER function](https://support.microsoft.com/en-us/office/upper-function-c11f29b3-d1a3-4537-8df6-04d0049963d6)

**Implementation Status:** ✅ Full implementation

**See Also:** ``LOWER``, ``PROPER``

---

### LOWER

Converts all letters in a text string to lowercase.

**Syntax:** `LOWER(text)`

**Parameters:**
- `text`: The text to convert to lowercase

**Returns:** Lowercase text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("LOWER", [.string("HELLO WORLD")]))
// Returns: .string("hello world")

// Using in formula
sheet.writeFormula("LOWER(A1)", to: "B1")
sheet.writeFormula("LOWER(\"COMPANY NAME\")", to: "B2")  // "company name"
```

**Excel Documentation:** [LOWER function](https://support.microsoft.com/en-us/office/lower-function-3f21df02-a80c-44b2-afaf-81358f9fdeb4)

**Implementation Status:** ✅ Full implementation

**See Also:** ``UPPER``, ``PROPER``

---

### PROPER

Converts text to title case (first letter of each word capitalized).

**Syntax:** `PROPER(text)`

**Parameters:**
- `text`: The text to convert to proper case

**Returns:** Title case text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("PROPER", [.string("john doe")]))
// Returns: .string("John Doe")

// Using in formula
sheet.writeFormula("PROPER(A1)", to: "B1")
sheet.writeFormula("PROPER(\"this is a title\")", to: "B2")  // "This Is A Title"
```

**Excel Documentation:** [PROPER function](https://support.microsoft.com/en-us/office/proper-function-52a5a283-e8b2-49be-8506-b2887b889f94)

**Implementation Status:** ✅ Full implementation

**See Also:** ``UPPER``, ``LOWER``

---

### CONCAT

Concatenates a list of text strings or values.

**Syntax:** `CONCAT(text1, [text2], ...)`

**Parameters:**
- `text1`, `text2`, ...: One or more values to join

**Returns:** Combined text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("CONCAT", [
    .string("Hello"), 
    .string(" "), 
    .string("World")
]))
// Returns: .string("Hello World")

// Using in formula
sheet.writeFormula("CONCAT(A1, \" \", B1)", to: "C1")
sheet.writeFormula("CONCAT(\"Total: \", A1)", to: "C2")
```

**Excel Documentation:** [CONCAT function](https://support.microsoft.com/en-us/office/concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2)

**Implementation Status:** ✅ Full implementation

**See Also:** ``CONCATENATE``, ``TEXTJOIN``, `&` operator

---

### CONCATENATE

Joins multiple text strings into one string.

**Syntax:** `CONCATENATE(text1, [text2], ...)`

**Parameters:**
- `text1`, `text2`, ...: Text items to join (up to 255)

**Returns:** Combined text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("CONCATENATE", [
    .string("First"),
    .string(" "),
    .string("Last")
]))
// Returns: .string("First Last")

// Using in formula
sheet.writeFormula("CONCATENATE(A1, \", \", B1)", to: "C1")
sheet.writeFormula("CONCATENATE(\"Value: $\", TEXT(A1, \"0.00\"))", to: "C2")
```

**Excel Documentation:** [CONCATENATE function](https://support.microsoft.com/en-us/office/concatenate-function-8f8ae884-2ca8-4f7a-b093-75d702bea31d)

**Implementation Status:** ✅ Full implementation

**Note:** In modern Excel, CONCAT and TEXTJOIN are preferred over CONCATENATE.

**See Also:** ``CONCAT``, ``TEXTJOIN``

---

### TEXTJOIN

Joins text from multiple ranges with a delimiter (Excel 365).

**Syntax:** `TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`

**Parameters:**
- `delimiter`: Text to insert between values
- `ignore_empty`: If TRUE, ignore empty cells
- `text1`, `text2`, ...: Text values or ranges to join

**Returns:** Joined text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("TEXTJOIN", [
    .string(", "),
    .boolean(true),
    .string("Apple"),
    .string(""),
    .string("Orange")
]))
// Returns: .string("Apple, Orange")  // Empty value ignored

// Using in formula
sheet.writeFormula("TEXTJOIN(\", \", TRUE, A1:A10)", to: "B1")
sheet.writeFormula("TEXTJOIN(\" | \", FALSE, A1, B1, C1)", to: "D1")
```

**Excel Documentation:** [TEXTJOIN function](https://support.microsoft.com/en-us/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c)

**Implementation Status:** ✅ Full implementation

**See Also:** ``CONCAT``, ``CONCATENATE``

---

### LEFT

Returns the leftmost characters from a text string.

**Syntax:** `LEFT(text, [num_chars])`

**Parameters:**
- `text`: The text string
- `num_chars` *(optional)*: Number of characters to extract (default: 1)

**Returns:** Leftmost characters

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("LEFT", [
    .string("Hello World"),
    .number(5.0)
]))
// Returns: .string("Hello")

// Using in formula
sheet.writeFormula("LEFT(A1, 3)", to: "B1")
sheet.writeFormula("LEFT(A1)", to: "B2")  // First character only
```

**Excel Documentation:** [LEFT function](https://support.microsoft.com/en-us/office/left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c)

**Implementation Status:** ✅ Full implementation

**See Also:** ``RIGHT``, ``MID``, ``LEFTB``

---

### LEFTB

Returns the leftmost characters based on byte count.

**Syntax:** `LEFTB(text, [num_bytes])`

**Parameters:**
- `text`: The text string
- `num_bytes` *(optional)*: Number of bytes to extract (default: 1)

**Returns:** Leftmost characters

**Excel Documentation:** [LEFTB function](https://support.microsoft.com/en-us/office/left-leftb-functions-9203d2d2-7960-479b-84c6-1ea52b99640c)

**Implementation Status:** ✅ Full implementation (delegates to LEFT)

**Note:** In Cuneiform, LEFTB behaves identically to LEFT.

**See Also:** ``LEFT``

---

### RIGHT

Returns the rightmost characters from a text string.

**Syntax:** `RIGHT(text, [num_chars])`

**Parameters:**
- `text`: The text string
- `num_chars` *(optional)*: Number of characters to extract (default: 1)

**Returns:** Rightmost characters

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("RIGHT", [
    .string("Hello World"),
    .number(5.0)
]))
// Returns: .string("World")

// Using in formula
sheet.writeFormula("RIGHT(A1, 4)", to: "B1")
sheet.writeFormula("RIGHT(A1)", to: "B2")  // Last character only
```

**Excel Documentation:** [RIGHT function](https://support.microsoft.com/en-us/office/right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f)

**Implementation Status:** ✅ Full implementation

**See Also:** ``LEFT``, ``MID``, ``RIGHTB``

---

### RIGHTB

Returns the rightmost characters based on byte count.

**Syntax:** `RIGHTB(text, [num_bytes])`

**Parameters:**
- `text`: The text string
- `num_bytes` *(optional)*: Number of bytes to extract (default: 1)

**Returns:** Rightmost characters

**Excel Documentation:** [RIGHTB function](https://support.microsoft.com/en-us/office/right-rightb-functions-240267ee-9afa-4639-a02b-f19e1786cf2f)

**Implementation Status:** ✅ Full implementation (delegates to RIGHT)

**Note:** In Cuneiform, RIGHTB behaves identically to RIGHT.

**See Also:** ``RIGHT``

---

### MID

Returns characters from the middle of a text string.

**Syntax:** `MID(text, start_num, num_chars)`

**Parameters:**
- `text`: The text string
- `start_num`: Starting position (1-based)
- `num_chars`: Number of characters to extract

**Returns:** Substring from specified position

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("MID", [
    .string("Hello World"),
    .number(7.0),
    .number(5.0)
]))
// Returns: .string("World")

// Using in formula
sheet.writeFormula("MID(A1, 3, 5)", to: "B1")
sheet.writeFormula("MID(A1, FIND(\" \", A1)+1, 10)", to: "B2")  // After first space
```

**Excel Documentation:** [MID function](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028)

**Implementation Status:** ✅ Full implementation

**See Also:** ``LEFT``, ``RIGHT``, ``MIDB``

---

### MIDB

Returns characters from the middle of a text string based on byte count.

**Syntax:** `MIDB(text, start_num, num_bytes)`

**Parameters:**
- `text`: The text string
- `start_num`: Starting byte position (1-based)
- `num_bytes`: Number of bytes to extract

**Returns:** Substring from specified position

**Excel Documentation:** [MIDB function](https://support.microsoft.com/en-us/office/mid-midb-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028)

**Implementation Status:** ✅ Full implementation (delegates to MID)

**Note:** In Cuneiform, MIDB behaves identically to MID.

**See Also:** ``MID``

---

### TRIM

Removes leading, trailing, and excess internal spaces from text.

**Syntax:** `TRIM(text)`

**Parameters:**
- `text`: The text to clean

**Returns:** Text with normalized spaces

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("TRIM", [
    .string("  Hello   World  ")
]))
// Returns: .string("Hello World")

// Using in formula
sheet.writeFormula("TRIM(A1)", to: "B1")
sheet.writeFormula("LEN(TRIM(A1))", to: "B2")  // Count after trimming
```

**Excel Documentation:** [TRIM function](https://support.microsoft.com/en-us/office/trim-function-410388fa-c5df-49c6-b16c-9e5630b479f9)

**Implementation Status:** ✅ Full implementation

**Note:** TRIM removes all spaces except single spaces between words.

**See Also:** ``CLEAN``

---

### CLEAN

Removes non-printable characters from text.

**Syntax:** `CLEAN(text)`

**Parameters:**
- `text`: The text to clean

**Returns:** Text with printable characters only

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("CLEAN", [
    .string("Hello\u{0001}World\u{001F}")
]))
// Returns: .string("HelloWorld")

// Using in formula
sheet.writeFormula("CLEAN(A1)", to: "B1")
sheet.writeFormula("TRIM(CLEAN(A1))", to: "B2")  // Clean and trim
```

**Excel Documentation:** [CLEAN function](https://support.microsoft.com/en-us/office/clean-function-26f3d7e5-475e-4a9c-90e5-4b8ba987ba41)

**Implementation Status:** ✅ Full implementation

**Note:** Removes characters with ASCII values 0-31.

**See Also:** ``TRIM``

---

### CHAR

Returns the character for a given ASCII code.

**Syntax:** `CHAR(number)`

**Parameters:**
- `number`: ASCII code (1-255)

**Returns:** Character corresponding to the code

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("CHAR", [.number(65.0)]))
// Returns: .string("A")

// Using in formula
sheet.writeFormula("CHAR(10)", to: "B1")  // Line feed
sheet.writeFormula("\"Line 1\" & CHAR(10) & \"Line 2\"", to: "B2")
sheet.writeFormula("CHAR(169)", to: "B3")  // Copyright symbol ©
```

**Excel Documentation:** [CHAR function](https://support.microsoft.com/en-us/office/char-function-bbd249c8-b36e-4a91-8017-1c133f9b837a)

**Implementation Status:** ✅ Full implementation

**See Also:** ``CODE``, ``UNICHAR``

---

### CODE

Returns the ASCII code of the first character in text.

**Syntax:** `CODE(text)`

**Parameters:**
- `text`: Text containing the character

**Returns:** ASCII code (numeric)

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("CODE", [.string("A")]))
// Returns: .number(65.0)

// Using in formula
sheet.writeFormula("CODE(A1)", to: "B1")
sheet.writeFormula("CODE(\"Hello\")", to: "B2")  // 72 (code for 'H')
```

**Excel Documentation:** [CODE function](https://support.microsoft.com/en-us/office/code-function-c32b692b-2ed0-4a04-bdd9-75640144b928)

**Implementation Status:** ✅ Full implementation

**See Also:** ``CHAR``, ``UNICODE``

---

### EXACT

Compares two text strings with case sensitivity.

**Syntax:** `EXACT(text1, text2)`

**Parameters:**
- `text1`: First text string
- `text2`: Second text string

**Returns:** 1 if identical, 0 if different (as number)

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("EXACT", [
    .string("Hello"),
    .string("hello")
]))
// Returns: .number(0.0)  // Not exact match (case differs)

// Using in formula
sheet.writeFormula("EXACT(A1, B1)", to: "C1")
sheet.writeFormula("IF(EXACT(A1, \"Password\"), \"Match\", \"No Match\")", to: "C2")
```

**Excel Documentation:** [EXACT function](https://support.microsoft.com/en-us/office/exact-function-d3087698-fc15-4a15-9631-12575cf29926)

**Implementation Status:** ✅ Full implementation

**Note:** Returns numeric 1 or 0 rather than boolean for Excel compatibility.

**See Also:** `=` operator (case-insensitive)

---

### REPLACE

Replaces part of a text string with different text.

**Syntax:** `REPLACE(old_text, start_num, num_chars, new_text)`

**Parameters:**
- `old_text`: Text to modify
- `start_num`: Starting position (1-based)
- `num_chars`: Number of characters to replace
- `new_text`: Replacement text

**Returns:** Modified text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("REPLACE", [
    .string("Hello World"),
    .number(7.0),
    .number(5.0),
    .string("Swift")
]))
// Returns: .string("Hello Swift")

// Using in formula
sheet.writeFormula("REPLACE(A1, 1, 3, \"***\")", to: "B1")
sheet.writeFormula("REPLACE(A1, 5, 0, \"-\")", to: "B2")  // Insert at position 5
```

**Excel Documentation:** [REPLACE function](https://support.microsoft.com/en-us/office/replace-replaceb-functions-8d799074-2425-4a8a-84bc-82472868878a)

**Implementation Status:** ✅ Full implementation

**See Also:** ``SUBSTITUTE``

---

### SUBSTITUTE

Substitutes new text for old text in a string.

**Syntax:** `SUBSTITUTE(text, old_text, new_text, [instance_num])`

**Parameters:**
- `text`: Text to search
- `old_text`: Text to replace
- `new_text`: Replacement text
- `instance_num` *(optional)*: Which occurrence to replace (default: all)

**Returns:** Text with substitutions made

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("SUBSTITUTE", [
    .string("Hello World"),
    .string("o"),
    .string("0")
]))
// Returns: .string("Hell0 W0rld")  // All occurrences

// Replace only second occurrence
let result2 = try evaluator.evaluate(.functionCall("SUBSTITUTE", [
    .string("one one one"),
    .string("one"),
    .string("two"),
    .number(2.0)
]))
// Returns: .string("one two one")

// Using in formula
sheet.writeFormula("SUBSTITUTE(A1, \" \", \"\")", to: "B1")  // Remove spaces
sheet.writeFormula("SUBSTITUTE(A1, \"old\", \"new\", 1)", to: "B2")
```

**Excel Documentation:** [SUBSTITUTE function](https://support.microsoft.com/en-us/office/substitute-function-6434944e-a904-4336-a9b0-1e58df3bc332)

**Implementation Status:** ✅ Full implementation

**See Also:** ``REPLACE``

---

### REPT

Repeats text a specified number of times.

**Syntax:** `REPT(text, number_times)`

**Parameters:**
- `text`: Text to repeat
- `number_times`: Number of repetitions (0-32767)

**Returns:** Repeated text

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("REPT", [
    .string("*"),
    .number(5.0)
]))
// Returns: .string("*****")

// Using in formula
sheet.writeFormula("REPT(\"-\", 10)", to: "B1")  // ----------
sheet.writeFormula("REPT(\" \", A1)", to: "B2")  // Indent by A1 spaces
sheet.writeFormula("REPT(\"=\", 50)", to: "B3")  // Divider line
```

**Excel Documentation:** [REPT function](https://support.microsoft.com/en-us/office/rept-function-04c4d778-e712-43b4-9c15-d656582bb061)

**Implementation Status:** ✅ Full implementation

**See Also:** ``CONCAT``

---

### VALUE

Converts text to a number.

**Syntax:** `VALUE(text)`

**Parameters:**
- `text`: Text representing a number

**Returns:** Numeric value

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("VALUE", [.string("123.45")]))
// Returns: .number(123.45)

// Using in formula
sheet.writeFormula("VALUE(A1)", to: "B1")
sheet.writeFormula("VALUE(\"$1,234.56\")", to: "B2")  // 1234.56
sheet.writeFormula("VALUE(\"10%\")", to: "B3")  // 10
```

**Excel Documentation:** [VALUE function](https://support.microsoft.com/en-us/office/value-function-257d0108-07dc-437d-ae1c-bc2d3953d8c2)

**Implementation Status:** ✅ Full implementation

**Note:** Automatically removes common formatting characters ($, comma, %).

**See Also:** ``NUMBERVALUE``, ``TEXT``

---

### NUMBERVALUE

Converts text to number using locale-specific separators.

**Syntax:** `NUMBERVALUE(text, [decimal_separator], [group_separator])`

**Parameters:**
- `text`: Text to convert
- `decimal_separator` *(optional)*: Decimal point character
- `group_separator` *(optional)*: Thousands separator character

**Returns:** Numeric value

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("NUMBERVALUE", [
    .string("1.234,56"),
    .string(","),
    .string(".")
]))
// Returns: .number(1234.56)  // European format

// Using in formula
sheet.writeFormula("NUMBERVALUE(A1, \",\", \".\")", to: "B1")
sheet.writeFormula("NUMBERVALUE(\"1'234.56\", \".\", \"'\")", to: "B2")  // Swiss format
```

**Excel Documentation:** [NUMBERVALUE function](https://support.microsoft.com/en-us/office/numbervalue-function-1b05c8cf-2bfa-4437-af70-596c7ea7d879)

**Implementation Status:** ✅ Full implementation

**See Also:** ``VALUE``

---

### FIND

Finds text within text (case-sensitive).

**Syntax:** `FIND(find_text, within_text, [start_num])`

**Parameters:**
- `find_text`: Text to find
- `within_text`: Text to search
- `start_num` *(optional)*: Starting position (default: 1)

**Returns:** Position of found text (1-based), or #VALUE! error

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("FIND", [
    .string("World"),
    .string("Hello World"),
    .number(1.0)
]))
// Returns: .number(7.0)

// Using in formula
sheet.writeFormula("FIND(\" \", A1)", to: "B1")  // Find first space
sheet.writeFormula("MID(A1, FIND(\"@\", A1)+1, 50)", to: "B2")  // After @ sign
```

**Excel Documentation:** [FIND function](https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628)

**Implementation Status:** ✅ Full implementation

**See Also:** ``SEARCH``, ``FINDB``

---

### FINDB

Finds text within text based on byte count (case-sensitive).

**Syntax:** `FINDB(find_text, within_text, [start_num])`

**Parameters:**
- `find_text`: Text to find
- `within_text`: Text to search
- `start_num` *(optional)*: Starting byte position (default: 1)

**Returns:** Byte position of found text, or #VALUE! error

**Excel Documentation:** [FINDB function](https://support.microsoft.com/en-us/office/find-findb-functions-c7912941-af2a-4bdf-a553-d0d89b0a0628)

**Implementation Status:** ✅ Full implementation (delegates to FIND)

**Note:** In Cuneiform, FINDB behaves identically to FIND.

**See Also:** ``FIND``

---

### SEARCH

Finds text within text (case-insensitive, supports wildcards).

**Syntax:** `SEARCH(find_text, within_text, [start_num])`

**Parameters:**
- `find_text`: Text to find (supports * and ? wildcards)
- `within_text`: Text to search
- `start_num` *(optional)*: Starting position (default: 1)

**Returns:** Position of found text (1-based), or #VALUE! error

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("SEARCH", [
    .string("world"),
    .string("Hello World"),
    .number(1.0)
]))
// Returns: .number(7.0)  // Case-insensitive match

// Using wildcards
let result2 = try evaluator.evaluate(.functionCall("SEARCH", [
    .string("W*d"),
    .string("Hello World")
]))
// Returns: .number(7.0)  // Matches "World"

// Using in formula
sheet.writeFormula("SEARCH(\"*@*.com\", A1)", to: "B1")  // Find email pattern
sheet.writeFormula("ISNUMBER(SEARCH(\"total\", A1))", to: "B2")  // Contains "total"?
```

**Excel Documentation:** [SEARCH function](https://support.microsoft.com/en-us/office/search-searchb-functions-9ab04538-0e55-4719-a72e-b6f54513b495)

**Implementation Status:** ✅ Full implementation

**See Also:** ``FIND``, ``SEARCHB``

---

### SEARCHB

Finds text within text based on byte count (case-insensitive).

**Syntax:** `SEARCHB(find_text, within_text, [start_num])`

**Parameters:**
- `find_text`: Text to find (supports wildcards)
- `within_text`: Text to search
- `start_num` *(optional)*: Starting byte position (default: 1)

**Returns:** Byte position of found text, or #VALUE! error

**Excel Documentation:** [SEARCHB function](https://support.microsoft.com/en-us/office/search-searchb-functions-9ab04538-0e55-4719-a72e-b6f54513b495)

**Implementation Status:** ✅ Full implementation (delegates to SEARCH)

**Note:** In Cuneiform, SEARCHB behaves identically to SEARCH.

**See Also:** ``SEARCH``

---

### TEXT

Formats a number as text using a custom format.

**Syntax:** `TEXT(value, format_text)`

**Parameters:**
- `value`: Number to format
- `format_text`: Format string (Excel number format)

**Returns:** Formatted text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("TEXT", [
    .number(1234.5),
    .string("#,##0.00")
]))
// Returns: .string("1,234.50")

// Using in formula
sheet.writeFormula("TEXT(A1, \"$#,##0.00\")", to: "B1")  // Currency
sheet.writeFormula("TEXT(A1, \"0.00%\")", to: "B2")  // Percentage
sheet.writeFormula("TEXT(A1, \"yyyy-mm-dd\")", to: "B3")  // Date
sheet.writeFormula("TEXT(TODAY(), \"dddd, mmmm d, yyyy\")", to: "B4")  // Long date
```

**Excel Documentation:** [TEXT function](https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c)

**Implementation Status:** ✅ Full implementation

**Supported Formats:**
- Number: `0`, `#`, `#,##0`, `0.00`
- Percentage: `0%`, `0.00%`
- Currency: `$#,##0`, `$#,##0.00`
- Date: `yyyy`, `mm`, `dd`, `m`, `d`

**See Also:** ``FIXED``, ``DOLLAR``

---

### FIXED

Formats a number as text with fixed decimal places.

**Syntax:** `FIXED(number, [decimals], [no_commas])`

**Parameters:**
- `number`: Number to format
- `decimals` *(optional)*: Number of decimal places (default: 2)
- `no_commas` *(optional)*: If TRUE, omit thousands separators (default: FALSE)

**Returns:** Formatted text string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("FIXED", [
    .number(1234.567),
    .number(2.0),
    .number(0.0)
]))
// Returns: .string("1,234.57")

// Using in formula
sheet.writeFormula("FIXED(A1, 2)", to: "B1")  // 2 decimals with commas
sheet.writeFormula("FIXED(A1, 0)", to: "B2")  // Integer with commas
sheet.writeFormula("FIXED(A1, 2, TRUE)", to: "B3")  // No commas
```

**Excel Documentation:** [FIXED function](https://support.microsoft.com/en-us/office/fixed-function-ffd5723c-324c-45e9-8b96-e41be2a8274a)

**Implementation Status:** ✅ Full implementation

**See Also:** ``TEXT``, ``DOLLAR``

---

### DOLLAR

Formats a number as currency text with dollar sign.

**Syntax:** `DOLLAR(number, [decimals])`

**Parameters:**
- `number`: Number to format
- `decimals` *(optional)*: Number of decimal places (default: 2)

**Returns:** Formatted currency string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("DOLLAR", [
    .number(1234.567),
    .number(2.0)
]))
// Returns: .string("$1,234.57")

// Using in formula
sheet.writeFormula("DOLLAR(A1)", to: "B1")  // 2 decimals by default
sheet.writeFormula("DOLLAR(A1, 0)", to: "B2")  // No decimals
sheet.writeFormula("DOLLAR(A1, 4)", to: "B3")  // 4 decimals
```

**Excel Documentation:** [DOLLAR function](https://support.microsoft.com/en-us/office/dollar-function-a6cd05d9-9740-4ad3-a469-8109d18ff611)

**Implementation Status:** ✅ Full implementation

**See Also:** ``TEXT``, ``FIXED``

---

### TEXTBEFORE

Returns text before a delimiter (Excel 365).

**Syntax:** `TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end])`

**Parameters:**
- `text`: Text to search
- `delimiter`: Text marking the boundary
- `instance_num` *(optional)*: Which occurrence (default: 1)
- `match_mode` *(optional)*: 0=case-sensitive, 1=case-insensitive (default: 0)
- `match_end` *(optional)*: Return text before end if delimiter not found

**Returns:** Text before delimiter, or #N/A error

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("TEXTBEFORE", [
    .string("apple,orange,banana"),
    .string(",")
]))
// Returns: .string("apple")

// Using in formula
sheet.writeFormula("TEXTBEFORE(A1, \"@\")", to: "B1")  // Username from email
sheet.writeFormula("TEXTBEFORE(A1, \" - \")", to: "B2")  // Before separator
```

**Excel Documentation:** [TEXTBEFORE function](https://support.microsoft.com/en-us/office/textbefore-function-d099c28a-dba8-448e-ac6c-f086d0fa1b29)

**Implementation Status:** ⚠️ Partial implementation (basic delimiter only)

**See Also:** ``TEXTAFTER``, ``TEXTSPLIT``

---

### TEXTAFTER

Returns text after a delimiter (Excel 365).

**Syntax:** `TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end])`

**Parameters:**
- `text`: Text to search
- `delimiter`: Text marking the boundary
- `instance_num` *(optional)*: Which occurrence (default: 1)
- `match_mode` *(optional)*: 0=case-sensitive, 1=case-insensitive (default: 0)
- `match_end` *(optional)*: Return text after end if delimiter not found

**Returns:** Text after delimiter, or #N/A error

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("TEXTAFTER", [
    .string("apple,orange,banana"),
    .string(",")
]))
// Returns: .string("orange,banana")

// Using in formula
sheet.writeFormula("TEXTAFTER(A1, \"@\")", to: "B1")  // Domain from email
sheet.writeFormula("TEXTAFTER(A1, \": \")", to: "B2")  // After label
```

**Excel Documentation:** [TEXTAFTER function](https://support.microsoft.com/en-us/office/textafter-function-c8db2546-5b51-416a-9690-c7e6722e90b4)

**Implementation Status:** ⚠️ Partial implementation (basic delimiter only)

**See Also:** ``TEXTBEFORE``, ``TEXTSPLIT``

---

### TEXTSPLIT

Splits text into an array using delimiters (Excel 365).

**Syntax:** `TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty])`

**Parameters:**
- `text`: Text to split
- `col_delimiter`: Column delimiter (splits into columns)
- `row_delimiter` *(optional)*: Row delimiter (splits into rows)
- `ignore_empty` *(optional)*: If TRUE, ignore empty cells

**Returns:** Array of split values

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("TEXTSPLIT", [
    .string("apple,orange,banana"),
    .string(",")
]))
// Returns: .array([["apple", "orange", "banana"]])

// Using in formula
sheet.writeFormula("TEXTSPLIT(A1, \",\")", to: "B1")  // Split CSV
sheet.writeFormula("TEXTSPLIT(A1, \" \")", to: "B2")  // Split words
```

**Excel Documentation:** [TEXTSPLIT function](https://support.microsoft.com/en-us/office/textsplit-function-b1ca414e-4c21-4ca0-b1b7-bdecace8a6e7)

**Implementation Status:** ⚠️ Partial implementation (column delimiter only)

**See Also:** ``TEXTBEFORE``, ``TEXTAFTER``

---

### T

Returns text or an empty string.

**Syntax:** `T(value)`

**Parameters:**
- `value`: Value to check

**Returns:** Text if value is text, otherwise empty string

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("T", [.string("Hello")]))
// Returns: .string("Hello")

let result2 = try evaluator.evaluate(.functionCall("T", [.number(123.0)]))
// Returns: .string("")

// Using in formula
sheet.writeFormula("T(A1)", to: "B1")
sheet.writeFormula("CONCAT(\"Value: \", T(A1))", to: "B2")
```

**Excel Documentation:** [T function](https://support.microsoft.com/en-us/office/t-function-fb83aeec-45e7-4924-af95-53e073541228)

**Implementation Status:** ✅ Full implementation

**See Also:** ``ISTEXT``

---

### UNICODE

Returns the Unicode code point of the first character.

**Syntax:** `UNICODE(text)`

**Parameters:**
- `text`: Text containing the character

**Returns:** Unicode code point (numeric)

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("UNICODE", [.string("A")]))
// Returns: .number(65.0)

let result2 = try evaluator.evaluate(.functionCall("UNICODE", [.string("😀")]))
// Returns: .number(128512.0)

// Using in formula
sheet.writeFormula("UNICODE(A1)", to: "B1")
sheet.writeFormula("UNICODE(\"€\")", to: "B2")  // 8364
```

**Excel Documentation:** [UNICODE function](https://support.microsoft.com/en-us/office/unicode-function-adb74aaa-a2a5-4dde-aff6-966e4e81f16f)

**Implementation Status:** ✅ Full implementation

**See Also:** ``UNICHAR``, ``CODE``

---

### UNICHAR

Returns the character for a given Unicode code point.

**Syntax:** `UNICHAR(number)`

**Parameters:**
- `number`: Unicode code point

**Returns:** Character for the code point

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("UNICHAR", [.number(65.0)]))
// Returns: .string("A")

let result2 = try evaluator.evaluate(.functionCall("UNICHAR", [.number(128512.0)]))
// Returns: .string("😀")

// Using in formula
sheet.writeFormula("UNICHAR(169)", to: "B1")  // ©
sheet.writeFormula("UNICHAR(8364)", to: "B2")  // €
sheet.writeFormula("UNICHAR(128512)", to: "B3")  // 😀
```

**Excel Documentation:** [UNICHAR function](https://support.microsoft.com/en-us/office/unichar-function-ffeb64f5-f131-44c6-b332-5cd72f0659b8)

**Implementation Status:** ✅ Full implementation

**See Also:** ``UNICODE``, ``CHAR``

---

### ARRAYTOTEXT

Converts an array to text representation.

**Syntax:** `ARRAYTOTEXT(array, [format])`

**Parameters:**
- `array`: Array to convert
- `format` *(optional)*: 0=concise, 1=strict (default: 0)

**Returns:** Text representation of array

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("ARRAYTOTEXT", [
    .array([[.number(1.0), .number(2.0)], [.number(3.0), .number(4.0)]])
]))
// Returns: .string("{1, 2; 3, 4}")

// Using in formula
sheet.writeFormula("ARRAYTOTEXT(A1:B2)", to: "C1")
```

**Excel Documentation:** [ARRAYTOTEXT function](https://support.microsoft.com/en-us/office/arraytotext-function-9cdcad46-2fa5-4c6b-ac92-14e7bc862b8b)

**Implementation Status:** ✅ Full implementation

**See Also:** ``VALUETOTEXT``

---

### VALUETOTEXT

Converts any value to text.

**Syntax:** `VALUETOTEXT(value, [format])`

**Parameters:**
- `value`: Value to convert
- `format` *(optional)*: 0=concise, 1=strict (default: 0)

**Returns:** Text representation of value

**Examples:**
```swift
let result = try evaluator.evaluate(.functionCall("VALUETOTEXT", [.number(123.45)]))
// Returns: .string("123.45")

let result2 = try evaluator.evaluate(.functionCall("VALUETOTEXT", [.boolean(true)]))
// Returns: .string("TRUE")

// Using in formula
sheet.writeFormula("VALUETOTEXT(A1)", to: "B1")
```

**Excel Documentation:** [VALUETOTEXT function](https://support.microsoft.com/en-us/office/valuetotext-function-5fff61a2-301a-4ab2-9ffa-0a5242a08fea)

**Implementation Status:** ✅ Full implementation

**See Also:** ``ARRAYTOTEXT``

---

### BAHTTEXT

Converts a number to Thai Baht text format.

**Syntax:** `BAHTTEXT(number)`

**Parameters:**
- `number`: Number to convert

**Returns:** Thai Baht text representation

**Examples:**
```swift
// Using in formula
sheet.writeFormula("BAHTTEXT(A1)", to: "B1")
```

**Excel Documentation:** [BAHTTEXT function](https://support.microsoft.com/en-us/office/bahttext-function-5ba4d0b4-abd3-4325-8d22-7a92d59aab9c)

**Implementation Status:** 🔄 Stub (returns #CALC! error)

**Note:** This function requires Thai language text formatting which is not yet implemented.

---

## Common Use Cases

### Data Cleaning

Remove extra spaces and non-printable characters:
```swift
sheet.writeFormula("TRIM(CLEAN(A1))", to: "B1")
```

### Email Processing

Extract username and domain:
```swift
sheet.writeFormula("TEXTBEFORE(A1, \"@\")", to: "B1")  // Username
sheet.writeFormula("TEXTAFTER(A1, \"@\")", to: "C1")   // Domain
```

Or using traditional functions:
```swift
sheet.writeFormula("LEFT(A1, FIND(\"@\", A1)-1)", to: "B1")  // Username
sheet.writeFormula("MID(A1, FIND(\"@\", A1)+1, 100)", to: "C1")  // Domain
```

### Name Formatting

Convert names to proper case:
```swift
sheet.writeFormula("PROPER(TRIM(A1))", to: "B1")
```

Split first and last name:
```swift
sheet.writeFormula("LEFT(A1, FIND(\" \", A1)-1)", to: "B1")  // First name
sheet.writeFormula("MID(A1, FIND(\" \", A1)+1, 100)", to: "C1")  // Last name
```

### Number Formatting

Format currency:
```swift
sheet.writeFormula("DOLLAR(A1, 2)", to: "B1")
sheet.writeFormula("TEXT(A1, \"$#,##0.00\")", to: "B2")
```

Format percentages:
```swift
sheet.writeFormula("TEXT(A1, \"0.00%\")", to: "B1")
```

Format with thousands separator:
```swift
sheet.writeFormula("FIXED(A1, 0)", to: "B1")
```

### Text Search and Extract

Find and extract specific patterns:
```swift
// Find phone number pattern
sheet.writeFormula("ISNUMBER(SEARCH(\"[0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]\", A1))", to: "B1")

// Extract after colon
sheet.writeFormula("TRIM(MID(A1, FIND(\":\", A1)+1, 1000))", to: "B2")
```

### CSV Parsing

Split comma-separated values (Excel 365):
```swift
sheet.writeFormula("TEXTSPLIT(A1, \",\")", to: "B1")
```

Traditional approach:
```swift
sheet.writeFormula("LEFT(A1, FIND(\",\", A1)-1)", to: "B1")  // First value
sheet.writeFormula("MID(A1, FIND(\",\", A1)+1, FIND(\",\", A1, FIND(\",\", A1)+1)-FIND(\",\", A1)-1)", to: "C1")  // Second value
```

### Concatenation with Formatting

Build formatted strings:
```swift
sheet.writeFormula("CONCAT(\"Total: \", DOLLAR(A1))", to: "B1")
sheet.writeFormula("TEXTJOIN(\" | \", TRUE, A1:A10)", to: "B2")
sheet.writeFormula("\"Last updated: \" & TEXT(NOW(), \"yyyy-mm-dd hh:mm\")", to: "B3")
```

### Character Manipulation

Insert special characters:
```swift
sheet.writeFormula("A1 & CHAR(10) & A2", to: "B1")  // Line break
sheet.writeFormula("REPT(\"=\", 50)", to: "B2")  // Divider
sheet.writeFormula("UNICHAR(9658) & \" \" & A1", to: "B3")  // Arrow bullet
```

## Performance Considerations

Text functions in Cuneiform are highly optimized:

- **LEFT/RIGHT/MID**: O(n) where n is the number of characters extracted
- **FIND/SEARCH**: O(n*m) where n is text length and m is search pattern length
- **SUBSTITUTE**: O(n*m) where n is text length and m is pattern length
- **CONCAT/TEXTJOIN**: O(n) where n is total length of all inputs
- **TRIM/CLEAN**: O(n) where n is text length

For processing large ranges:
- Use ``TEXTJOIN`` instead of multiple ``CONCATENATE`` calls
- Use ``SUBSTITUTE`` with instance_num to replace specific occurrences
- Combine ``TRIM`` and ``CLEAN`` in a single formula for efficiency

## See Also

- <doc:FormulaReference> - Complete function reference
- <doc:Logical> - Logical functions for conditional text operations
- <doc:Information> - Type checking functions (ISTEXT, etc.)
- ``FormulaEvaluator`` - Evaluate text formulas programmatically
