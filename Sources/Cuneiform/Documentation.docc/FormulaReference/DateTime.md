# Date & Time Functions

Excel-compatible date and time functions for working with dates, times, and durations in Cuneiform spreadsheets.

## Overview

Date and time functions enable you to work with temporal data in your spreadsheets. Cuneiform implements Excel's date/time system, where dates are stored as serial numbers representing the number of days since December 30, 1899 (Excel's epoch). Times are represented as decimal fractions of a day (e.g., 0.5 = 12:00 PM).

### Excel Date Serial Number System

Excel uses a serial number system for dates:
- **Serial number 1** = January 1, 1900
- **Serial number 44562** = January 1, 2022
- **Time component** = decimal fraction (0.0 to 0.99999...)

**Important**: Excel incorrectly treats 1900 as a leap year. Cuneiform replicates this behavior for Excel compatibility.

### Time Representation

Times are stored as decimal fractions of a 24-hour day:
- **0.0** = 12:00 AM (midnight)
- **0.5** = 12:00 PM (noon)
- **0.25** = 6:00 AM
- **0.75** = 6:00 PM

## Function Categories

### Current Date & Time

| Function | Description |
|----------|-------------|
| ``TODAY`` | Returns the current date (without time) |
| ``NOW`` | Returns the current date and time |

### Date Construction

| Function | Description |
|----------|-------------|
| ``DATE`` | Creates a date from year, month, and day values |
| ``DATEVALUE`` | Converts a date text string to a serial number |

### Time Construction

| Function | Description |
|----------|-------------|
| ``TIME`` | Creates a time value from hour, minute, and second |
| ``TIMEVALUE`` | Converts a time text string to a serial number |

### Date Part Extraction

| Function | Description |
|----------|-------------|
| ``YEAR`` | Extracts the year from a date |
| ``MONTH`` | Extracts the month from a date (1-12) |
| ``DAY`` | Extracts the day of the month from a date (1-31) |
| ``WEEKDAY`` | Returns the day of the week (1-7) |
| ``WEEKNUM`` | Returns the week number of the year |
| ``ISOWEEKNUM`` | Returns the ISO week number of the year |

### Time Part Extraction

| Function | Description |
|----------|-------------|
| ``HOUR`` | Extracts the hour component (0-23) |
| ``MINUTE`` | Extracts the minute component (0-59) |
| ``SECOND`` | Extracts the second component (0-59) |

### Date Calculations

| Function | Description |
|----------|-------------|
| ``DAYS`` | Returns the number of days between two dates |
| ``DAYS360`` | Calculates days between dates using 360-day year |
| ``EDATE`` | Returns a date n months from a start date |
| ``EOMONTH`` | Returns the last day of a month n months away |
| ``DATEDIF`` | Calculates the difference between dates in various units |
| ``YEARFRAC`` | Returns the fraction of a year between two dates |

### Business Day Calculations

| Function | Description |
|----------|-------------|
| ``NETWORKDAYS`` | Counts working days between dates (excluding weekends) |
| ``WORKDAY`` | Returns a date n working days from a start date |
| ``NETWORKDAYS.INTL`` | Counts working days with custom weekend patterns |
| ``WORKDAY.INTL`` | Calculates workday with custom weekend patterns |

---

## Current Date & Time Functions

### TODAY

Returns the current date as an Excel serial number (without time component).

**Syntax:** `TODAY()`

**Parameters:** None

**Returns:** Number - The current date as a serial number

**Examples:**
```swift
let evaluator = FormulaEvaluator()
let result = evaluator.evaluate("=TODAY()")  // 45324 (for Jan 28, 2024)

// Calculate days until future date
evaluator.evaluate("=DATE(2024,12,31)-TODAY()")  // Days until end of year
```

**Excel Documentation:** [TODAY function](https://support.microsoft.com/en-us/office/today-function-5eb6078d-a82c-4736-8930-2f51a028397f)

**Implementation Status:** ✅ Full implementation

**See Also:** ``NOW``, ``DATE``

---

### NOW

Returns the current date and time as an Excel serial number.

**Syntax:** `NOW()`

**Parameters:** None

**Returns:** Number - The current date and time as a serial number with decimal fraction

**Examples:**
```swift
let evaluator = FormulaEvaluator()
let result = evaluator.evaluate("=NOW()")  // 45324.68542 (date + time)

// Extract just the time portion
evaluator.evaluate("=NOW()-TODAY()")  // 0.68542 (time as decimal)

// Calculate hours since midnight
evaluator.evaluate("=(NOW()-TODAY())*24")  // 16.45 hours
```

**Excel Documentation:** [NOW function](https://support.microsoft.com/en-us/office/now-function-3337fd29-145a-4347-b2e6-20c904739c46)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:1553

**See Also:** ``TODAY``, ``TIME``, ``HOUR``, ``MINUTE``, ``SECOND``

---

## Date Construction Functions

### DATE

Creates a date serial number from year, month, and day components.

**Syntax:** `DATE(year, month, day)`

**Parameters:**
- `year`: The year (0-99 interpreted as 1900-1999, 100-9999 as-is)
- `month`: The month (1-12, overflow/underflow handled automatically)
- `day`: The day of month (1-31, overflow/underflow handled automatically)

**Returns:** Number - Excel serial number for the specified date

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=DATE(2024,1,28)")  // 45324

// Overflow handling
evaluator.evaluate("=DATE(2023,13,1)")  // Jan 1, 2024 (month overflow)
evaluator.evaluate("=DATE(2024,1,32)")  // Feb 1, 2024 (day overflow)

// Two-digit year
evaluator.evaluate("=DATE(24,1,28)")  // January 28, 1924
```

**Excel Documentation:** [DATE function](https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:1572. Handles Excel's 1900 leap year bug and automatic overflow adjustment.

**See Also:** ``YEAR``, ``MONTH``, ``DAY``, ``DATEVALUE``

---

### DATEVALUE

Converts a date text string to an Excel serial number.

**Syntax:** `DATEVALUE(date_text)`

**Parameters:**
- `date_text`: A text string representing a date in a recognized format

**Returns:** Number - The serial number of the date

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=DATEVALUE(\"1/1/2024\")")  // Date serial number
evaluator.evaluate("=DATEVALUE(\"January 1, 2024\")")  // Same result
```

**Excel Documentation:** [DATEVALUE function](https://support.microsoft.com/en-us/office/datevalue-function-df8b07d4-7761-4a93-bc33-b7471bbff252)

**Implementation Status:** 🔄 Stub - Returns #CALC! error

**Implementation Note:** Located in FormulaEvaluator.swift:11958. Requires date string parsing implementation.

**See Also:** ``DATE``, ``TIMEVALUE``

---

## Time Construction Functions

### TIME

Creates a time serial number from hour, minute, and second components.

**Syntax:** `TIME(hour, minute, second)`

**Parameters:**
- `hour`: The hour (0-23)
- `minute`: The minute (0-59)
- `second`: The second (0-59)

**Returns:** Number - Decimal fraction representing the time (0.0 to 0.99999...)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=TIME(12,0,0)")  // 0.5 (noon)
evaluator.evaluate("=TIME(18,30,0)")  // 0.770833 (6:30 PM)
evaluator.evaluate("=TIME(0,0,0)")   // 0.0 (midnight)

// Combine with date
evaluator.evaluate("=DATE(2024,1,28)+TIME(14,30,0)")  // Jan 28, 2024 2:30 PM
```

**Excel Documentation:** [TIME function](https://support.microsoft.com/en-us/office/time-function-9a5aff99-8f7d-4611-845e-747d0b8d5457)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:5923. Converts hours, minutes, and seconds to a fraction of a 24-hour day.

**See Also:** ``HOUR``, ``MINUTE``, ``SECOND``, ``TIMEVALUE``, ``NOW``

---

### TIMEVALUE

Converts a time text string to a decimal serial number.

**Syntax:** `TIMEVALUE(time_text)`

**Parameters:**
- `time_text`: A text string in time format (e.g., "14:30:00" or "14:30")

**Returns:** Number - Decimal fraction representing the time

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=TIMEVALUE(\"12:00:00\")")  // 0.5
evaluator.evaluate("=TIMEVALUE(\"18:30\")")     // 0.770833
evaluator.evaluate("=TIMEVALUE(\"6:30 PM\")")   // Parse AM/PM format
```

**Excel Documentation:** [TIMEVALUE function](https://support.microsoft.com/en-us/office/timevalue-function-0b615c12-33d8-4431-bf3d-f3eb6d186645)

**Implementation Status:** ⚠️ Partial - Basic "HH:MM:SS" and "HH:MM" format parsing only

**Implementation Note:** Located in FormulaEvaluator.swift:5944. Currently supports simple colon-separated formats. AM/PM parsing not yet implemented.

**See Also:** ``TIME``, ``DATEVALUE``, ``HOUR``, ``MINUTE``, ``SECOND``

---

## Date Part Extraction Functions

### YEAR

Extracts the year component from a date serial number.

**Syntax:** `YEAR(serial_number)`

**Parameters:**
- `serial_number`: An Excel date serial number

**Returns:** Number - The year as an integer (1900-9999)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=YEAR(45324)")  // 2024
evaluator.evaluate("=YEAR(DATE(2024,1,28))")  // 2024
evaluator.evaluate("=YEAR(TODAY())")  // Current year
```

**Excel Documentation:** [YEAR function](https://support.microsoft.com/en-us/office/year-function-c64f017a-1354-4f75-8f07-24a64f3a9845)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:1591

**See Also:** ``MONTH``, ``DAY``, ``DATE``

---

### MONTH

Extracts the month component from a date serial number.

**Syntax:** `MONTH(serial_number)`

**Parameters:**
- `serial_number`: An Excel date serial number

**Returns:** Number - The month as an integer (1-12)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=MONTH(45324)")  // 1 (January)
evaluator.evaluate("=MONTH(DATE(2024,6,15))")  // 6 (June)
```

**Excel Documentation:** [MONTH function](https://support.microsoft.com/en-us/office/month-function-579a2881-199b-48b2-ab90-ddba0eba86e8)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:1606

**See Also:** ``YEAR``, ``DAY``, ``DATE``

---

### DAY

Extracts the day of the month from a date serial number.

**Syntax:** `DAY(serial_number)`

**Parameters:**
- `serial_number`: An Excel date serial number

**Returns:** Number - The day as an integer (1-31)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=DAY(45324)")  // 28
evaluator.evaluate("=DAY(DATE(2024,1,28))")  // 28
```

**Excel Documentation:** [DAY function](https://support.microsoft.com/en-us/office/day-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:1621

**See Also:** ``YEAR``, ``MONTH``, ``DATE``

---

### WEEKDAY

Returns the day of the week for a given date.

**Syntax:** `WEEKDAY(serial_number, [return_type])`

**Parameters:**
- `serial_number`: An Excel date serial number
- `return_type` *(optional)*: Determines the return value format
  - 1 or omitted: Numbers 1 (Sunday) through 7 (Saturday)
  - 2: Numbers 1 (Monday) through 7 (Sunday)
  - 3: Numbers 0 (Monday) through 6 (Sunday)

**Returns:** Number - The day of the week as specified by return_type

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=WEEKDAY(DATE(2024,1,28))")     // 1 (Sunday)
evaluator.evaluate("=WEEKDAY(DATE(2024,1,28), 2)")  // 7 (Sunday in Mon-Sun system)
evaluator.evaluate("=WEEKDAY(DATE(2024,1,29), 2)")  // 1 (Monday)
```

**Excel Documentation:** [WEEKDAY function](https://support.microsoft.com/en-us/office/weekday-function-60e44483-2ed1-439f-8bd0-e404c190949a)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:5675

**See Also:** ``WEEKNUM``, ``ISOWEEKNUM``

---

### WEEKNUM

Returns the week number of the year for a given date.

**Syntax:** `WEEKNUM(serial_number, [return_type])`

**Parameters:**
- `serial_number`: An Excel date serial number
- `return_type` *(optional)*: System to use (1 = week starts Sunday, 2 = week starts Monday)

**Returns:** Number - The week number (1-53)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=WEEKNUM(DATE(2024,1,28))")  // Week number of the year
evaluator.evaluate("=WEEKNUM(DATE(2024,1,1))")   // 1
```

**Excel Documentation:** [WEEKNUM function](https://support.microsoft.com/en-us/office/weeknum-function-e5c43a03-b4ab-426c-b411-b18c13c75340)

**Implementation Status:** ⚠️ Partial - Simplified calculation based on day of year

**Implementation Note:** Located in FormulaEvaluator.swift:5704. Current implementation uses simplified day-of-year calculation.

**See Also:** ``ISOWEEKNUM``, ``WEEKDAY``

---

### ISOWEEKNUM

Returns the ISO week number of the year for a given date.

**Syntax:** `ISOWEEKNUM(serial_number)`

**Parameters:**
- `serial_number`: An Excel date serial number

**Returns:** Number - The ISO week number (1-53)

**Notes:** ISO weeks start on Monday, and week 1 is the week containing January 4th.

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=ISOWEEKNUM(DATE(2024,1,28))")  // ISO week number
evaluator.evaluate("=ISOWEEKNUM(DATE(2024,1,1))")   // Could be week 52 or 1
```

**Excel Documentation:** [ISOWEEKNUM function](https://support.microsoft.com/en-us/office/isoweeknum-function-1c2d0afe-d25b-4ab1-8894-8d0520e90e0e)

**Implementation Status:** ⚠️ Partial - Simplified calculation

**Implementation Note:** Located in FormulaEvaluator.swift:5725. Simplified implementation may not handle edge cases correctly.

**See Also:** ``WEEKNUM``, ``WEEKDAY``

---

## Time Part Extraction Functions

### HOUR

Extracts the hour component from a time serial number.

**Syntax:** `HOUR(serial_number)`

**Parameters:**
- `serial_number`: An Excel date/time serial number

**Returns:** Number - The hour component (0-23)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=HOUR(0.5)")  // 12 (noon)
evaluator.evaluate("=HOUR(TIME(14,30,0))")  // 14
evaluator.evaluate("=HOUR(NOW())")  // Current hour
```

**Excel Documentation:** [HOUR function](https://support.microsoft.com/en-us/office/hour-function-a3afa879-86cb-4339-b1b5-2dd2d7310ac7)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:5972

**See Also:** ``MINUTE``, ``SECOND``, ``TIME``, ``NOW``

---

### MINUTE

Extracts the minute component from a time serial number.

**Syntax:** `MINUTE(serial_number)`

**Parameters:**
- `serial_number`: An Excel date/time serial number

**Returns:** Number - The minute component (0-59)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=MINUTE(TIME(14,30,0))")  // 30
evaluator.evaluate("=MINUTE(NOW())")  // Current minute
```

**Excel Documentation:** [MINUTE function](https://support.microsoft.com/en-us/office/minute-function-af728df0-05c4-4b07-9eed-a84801a60589)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:5990

**See Also:** ``HOUR``, ``SECOND``, ``TIME``, ``NOW``

---

### SECOND

Extracts the second component from a time serial number.

**Syntax:** `SECOND(serial_number)`

**Parameters:**
- `serial_number`: An Excel date/time serial number

**Returns:** Number - The second component (0-59)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=SECOND(TIME(14,30,45))")  // 45
evaluator.evaluate("=SECOND(NOW())")  // Current second
```

**Excel Documentation:** [SECOND function](https://support.microsoft.com/en-us/office/second-function-740d1cfc-553c-4099-b668-80eaa24e8af1)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:6008

**See Also:** ``HOUR``, ``MINUTE``, ``TIME``, ``NOW``

---

## Date Calculation Functions

### DAYS

Returns the number of days between two dates.

**Syntax:** `DAYS(end_date, start_date)`

**Parameters:**
- `end_date`: The ending date serial number
- `start_date`: The starting date serial number

**Returns:** Number - The number of days between the dates (can be negative)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=DAYS(DATE(2024,12,31), DATE(2024,1,1))")  // 365
evaluator.evaluate("=DAYS(TODAY(), DATE(2024,1,1))")  // Days since Jan 1
evaluator.evaluate("=DAYS(DATE(2024,1,1), DATE(2024,12,31))")  // -365 (negative)
```

**Excel Documentation:** [DAYS function](https://support.microsoft.com/en-us/office/days-function-57740535-d549-4395-8728-0f07bff0b9df)

**Implementation Status:** ✅ Full implementation

**Implementation Note:** Located in FormulaEvaluator.swift:9874

**See Also:** ``DATEDIF``, ``DAYS360``

---

### DAYS360

Calculates the number of days between two dates using a 360-day year (12 months of 30 days each).

**Syntax:** `DAYS360(start_date, end_date, [method])`

**Parameters:**
- `start_date`: The starting date serial number
- `end_date`: The ending date serial number
- `method` *(optional)*: FALSE (US/NASD method) or TRUE (European method)

**Returns:** Number - The number of days on a 360-day year basis

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=DAYS360(DATE(2024,1,1), DATE(2024,12,31))")  // 360
evaluator.evaluate("=DAYS360(DATE(2024,1,1), DATE(2024,7,1))")    // 180
```

**Excel Documentation:** [DAYS360 function](https://support.microsoft.com/en-us/office/days360-function-b9a509fd-49ef-407e-94df-0cbda5718c2a)

**Implementation Status:** 🔄 Stub - Returns #CALC! error

**Implementation Note:** Located in FormulaEvaluator.swift:11953. Requires 360-day year calculation implementation.

**See Also:** ``DAYS``, ``YEARFRAC``

---

### EDATE

Returns a date that is a specified number of months before or after a start date.

**Syntax:** `EDATE(start_date, months)`

**Parameters:**
- `start_date`: The starting date serial number
- `months`: Number of months to add (negative to subtract)

**Returns:** Number - The resulting date serial number

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=EDATE(DATE(2024,1,15), 3)")   // April 15, 2024
evaluator.evaluate("=EDATE(DATE(2024,6,15), -2)")  // April 15, 2024
evaluator.evaluate("=EDATE(TODAY(), 6)")  // Date 6 months from now
```

**Excel Documentation:** [EDATE function](https://support.microsoft.com/en-us/office/edate-function-3c920eb2-6e66-44e7-a1f5-753ae47ee4f5)

**Implementation Status:** ⚠️ Partial - Uses approximate 30-day month calculation

**Implementation Note:** Located in FormulaEvaluator.swift:5790. Current implementation multiplies months by 30 days (approximate).

**See Also:** ``EOMONTH``, ``DATE``

---

### EOMONTH

Returns the last day of the month that is a specified number of months before or after the start date.

**Syntax:** `EOMONTH(start_date, months)`

**Parameters:**
- `start_date`: The starting date serial number
- `months`: Number of months to add (negative to subtract)

**Returns:** Number - The serial number of the last day of the target month

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=EOMONTH(DATE(2024,1,15), 0)")   // Jan 31, 2024
evaluator.evaluate("=EOMONTH(DATE(2024,1,15), 1)")   // Feb 29, 2024 (leap year)
evaluator.evaluate("=EOMONTH(DATE(2024,2,15), 0)")   // Feb 29, 2024
```

**Excel Documentation:** [EOMONTH function](https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628)

**Implementation Status:** ⚠️ Partial - Approximate calculation with leap year handling

**Implementation Note:** Located in FormulaEvaluator.swift:5743. Uses simplified date arithmetic.

**See Also:** ``EDATE``, ``DAY``

---

### DATEDIF

Calculates the difference between two dates in various units.

**Syntax:** `DATEDIF(start_date, end_date, unit)`

**Parameters:**
- `start_date`: The starting date serial number
- `end_date`: The ending date serial number
- `unit`: The type of difference to return:
  - "D": Days
  - "M": Months (approximate)
  - "Y": Years (approximate)
  - "MD": Days ignoring months and years
  - "YM": Months ignoring years
  - "YD": Days ignoring years

**Returns:** Number - The calculated difference in the specified unit

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=DATEDIF(DATE(2020,1,1), DATE(2024,1,1), \"Y\")")   // 4 years
evaluator.evaluate("=DATEDIF(DATE(2024,1,1), DATE(2024,6,15), \"M\")")  // ~5 months
evaluator.evaluate("=DATEDIF(DATE(2024,1,1), DATE(2024,1,15), \"D\")")  // 14 days
```

**Excel Documentation:** [DATEDIF function](https://support.microsoft.com/en-us/office/datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c)

**Implementation Status:** ⚠️ Partial - Uses approximate calculations (30-day months, 365-day years)

**Implementation Note:** Located in FormulaEvaluator.swift:5852. The implementation uses simplified calculations and may not match Excel exactly for complex date differences.

**See Also:** ``DAYS``, ``YEARFRAC``

---

### YEARFRAC

Returns the fraction of a year represented by the number of whole days between two dates.

**Syntax:** `YEARFRAC(start_date, end_date, [basis])`

**Parameters:**
- `start_date`: The starting date serial number
- `end_date`: The ending date serial number
- `basis` *(optional)*: The day count basis to use:
  - 0 or omitted: US (NASD) 30/360
  - 1: Actual/actual
  - 2: Actual/360
  - 3: Actual/365
  - 4: European 30/360

**Returns:** Number - The fraction of a year (e.g., 0.5 = 6 months)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=YEARFRAC(DATE(2024,1,1), DATE(2024,7,1))")      // ~0.5
evaluator.evaluate("=YEARFRAC(DATE(2024,1,1), DATE(2025,1,1), 1)")   // ~1.0
evaluator.evaluate("=YEARFRAC(DATE(2024,1,1), DATE(2024,12,31), 3)") // 365/365
```

**Excel Documentation:** [YEARFRAC function](https://support.microsoft.com/en-us/office/yearfrac-function-3844141e-c76d-4143-82b6-208454ddc6a8)

**Implementation Status:** ✅ Full implementation with all basis options

**Implementation Note:** Located in FormulaEvaluator.swift:5888

**See Also:** ``DATEDIF``, ``DAYS``

---

## Business Day Calculation Functions

### NETWORKDAYS

Returns the number of whole working days between two dates (excludes weekends).

**Syntax:** `NETWORKDAYS(start_date, end_date, [holidays])`

**Parameters:**
- `start_date`: The starting date serial number
- `end_date`: The ending date serial number
- `holidays` *(optional)*: Range of dates to exclude as holidays (not yet implemented)

**Returns:** Number - The number of working days (Monday-Friday)

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,31))")  // ~22 workdays
evaluator.evaluate("=NETWORKDAYS(DATE(2024,1,1), DATE(2024,1,7))")   // 5 (Mon-Fri)
```

**Excel Documentation:** [NETWORKDAYS function](https://support.microsoft.com/en-us/office/networkdays-function-48e717bf-a7a3-495f-969e-5005e3eb18e7)

**Implementation Status:** ⚠️ Partial - Basic calculation without holiday exclusion

**Implementation Note:** Located in FormulaEvaluator.swift:5808. Assumes 5 working days per week (Monday-Friday). Holiday parameter not yet implemented.

**See Also:** ``WORKDAY``, ``NETWORKDAYS.INTL``

---

### WORKDAY

Returns a date that is a specified number of working days from a start date.

**Syntax:** `WORKDAY(start_date, days, [holidays])`

**Parameters:**
- `start_date`: The starting date serial number
- `days`: Number of working days to add (negative to subtract)
- `holidays` *(optional)*: Range of dates to exclude as holidays (not yet implemented)

**Returns:** Number - The resulting date serial number

**Examples:**
```swift
let evaluator = FormulaEvaluator()
evaluator.evaluate("=WORKDAY(DATE(2024,1,1), 10)")   // 10 workdays after Jan 1
evaluator.evaluate("=WORKDAY(DATE(2024,1,15), -5)")  // 5 workdays before Jan 15
```

**Excel Documentation:** [WORKDAY function](https://support.microsoft.com/en-us/office/workday-function-f764a5b7-05fc-4494-9486-60d494efbf33)

**Implementation Status:** ⚠️ Partial - Approximate calculation using 1.4 calendar days per workday

**Implementation Note:** Located in FormulaEvaluator.swift:5832. Current implementation uses a simplified ratio (5/7). Holiday parameter not yet implemented.

**See Also:** ``NETWORKDAYS``, ``WORKDAY.INTL``

---

### NETWORKDAYS.INTL

Returns the number of working days between two dates with custom weekend patterns.

**Syntax:** `NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])`

**Parameters:**
- `start_date`: The starting date serial number
- `end_date`: The ending date serial number
- `weekend` *(optional)*: Weekend day pattern (1-17 for predefined patterns, or custom string)
  - 1 or omitted: Saturday, Sunday
  - 2: Sunday, Monday
  - 3: Monday, Tuesday
  - 4: Tuesday, Wednesday
  - 5: Wednesday, Thursday
  - 6: Thursday, Friday
  - 7: Friday, Saturday
  - 11-17: Single weekend days
- `holidays` *(optional)*: Range of dates to exclude (not yet implemented)

**Returns:** Number - The number of working days

**Examples:**
```swift
let evaluator = FormulaEvaluator()
// Standard weekends (Sat-Sun)
evaluator.evaluate("=NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,31))")

// Friday-Saturday weekend (Middle East pattern)
evaluator.evaluate("=NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,31), 7)")

// Sunday-Monday weekend
evaluator.evaluate("=NETWORKDAYS.INTL(DATE(2024,1,1), DATE(2024,1,31), 2)")
```

**Excel Documentation:** [NETWORKDAYS.INTL function](https://support.microsoft.com/en-us/office/networkdays-intl-function-a9b26239-4f20-46a1-9ab8-4e925bfd5e28)

**Implementation Status:** ✅ Full implementation with predefined weekend patterns

**Implementation Note:** Located in FormulaEvaluator.swift:12112. Supports patterns 1-17. Custom string patterns not yet implemented. Holiday parameter not yet implemented.

**See Also:** ``WORKDAY.INTL``, ``NETWORKDAYS``

---

### WORKDAY.INTL

Returns a date n working days from a start date with custom weekend patterns.

**Syntax:** `WORKDAY.INTL(start_date, days, [weekend], [holidays])`

**Parameters:**
- `start_date`: The starting date serial number
- `days`: Number of working days to add (negative to subtract)
- `weekend` *(optional)*: Weekend day pattern (1-17 for predefined patterns)
  - See ``NETWORKDAYS.INTL`` for pattern details
- `holidays` *(optional)*: Range of dates to exclude (not yet implemented)

**Returns:** Number - The resulting date serial number

**Examples:**
```swift
let evaluator = FormulaEvaluator()
// 10 workdays with standard weekend
evaluator.evaluate("=WORKDAY.INTL(DATE(2024,1,1), 10)")

// 10 workdays with Friday-Saturday weekend
evaluator.evaluate("=WORKDAY.INTL(DATE(2024,1,1), 10, 7)")

// Go back 5 workdays
evaluator.evaluate("=WORKDAY.INTL(DATE(2024,1,15), -5)")
```

**Excel Documentation:** [WORKDAY.INTL function](https://support.microsoft.com/en-us/office/workday-intl-function-a378391c-9ba7-4678-8a39-39611a9bf81d)

**Implementation Status:** ✅ Full implementation with predefined weekend patterns

**Implementation Note:** Located in FormulaEvaluator.swift:12148. Supports patterns 1-17. Custom string patterns not yet implemented. Holiday parameter not yet implemented.

**See Also:** ``NETWORKDAYS.INTL``, ``WORKDAY``

---

## See Also

- ``FormulaReference`` - Complete formula function reference
- ``FormulaEvaluator`` - Core formula evaluation engine
- [Excel Date and Time Functions](https://support.microsoft.com/en-us/office/date-and-time-functions-reference-fd1b5961-c1ae-4677-be58-074152f97b81) - Microsoft Office documentation

## Topics

### Current Date & Time
- ``TODAY``
- ``NOW``

### Date Construction
- ``DATE``
- ``DATEVALUE``

### Time Construction
- ``TIME``
- ``TIMEVALUE``

### Date Part Extraction
- ``YEAR``
- ``MONTH``
- ``DAY``
- ``WEEKDAY``
- ``WEEKNUM``
- ``ISOWEEKNUM``

### Time Part Extraction
- ``HOUR``
- ``MINUTE``
- ``SECOND``

### Date Calculations
- ``DAYS``
- ``DAYS360``
- ``EDATE``
- ``EOMONTH``
- ``DATEDIF``
- ``YEARFRAC``

### Business Day Calculations
- ``NETWORKDAYS``
- ``WORKDAY``
- ``NETWORKDAYS.INTL``
- ``WORKDAY.INTL``
