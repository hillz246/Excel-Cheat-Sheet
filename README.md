# Excel-Cheat-Sheet
Excel Cheat Sheet Formulas for Data Analysis

****17 Excel cheat sheet formulas commonly used for data analysis.****

***1. SUM***
The SUM function in Excel is a fundamental tool used to add together a range of cells. It's essential for tasks involving numerical data aggregation, such as totaling sales, calculating expenses, or summarizing data sets.


**Syntax Explanation:**

    =SUM(number1, [number2], …)

**Advanced Usage:**

    Dynamic Ranges: Use OFFSET or INDIRECT with SUM to create dynamic ranges that update based on your data.

    =SUM(OFFSET(A1, 0, 0, COUNTA(A:A), 1))

    Array Formulas: Sum only certain values within an array based on conditions.

    =SUM(IF(A1:A10>5, A1:A10, 0))  // Array formula, press Ctrl+Shift+Enter

***2. AVERAGE***

The AVERAGE function calculates the mean of a group of numbers. It is essential for understanding the central tendency of your data, such as average sales, grades, or measurements.

**Syntax Explanation:**

    =AVERAGE(number1, [number2], …)
      
      number1, [number2], …: The numbers or ranges to average.

**Advanced Usage:**

    Conditional Averages: Use AVERAGEIF or AVERAGEIFS for conditional averages.


  =AVERAGEIF(A1:A10, “>5”)

    Weighted Averages: Calculate weighted averages using SUMPRODUCT.


  =SUMPRODUCT(A1:A10, B1:B10) / SUM(B1:B10)

***3. COUNT***

The COUNT function counts the number of cells that contain numbers within a range. It is useful for understanding the quantity of numeric entries in a dataset, such as counting the number of sales transactions or survey responses.
Syntax Explanation:

    =COUNT(value1, [value2], …)

    value1, [value2], …: The values or ranges to count.

**Advanced Usage:**

    Count Non-Empty Cells: Use COUNTA to count non-empty cells, which can include text and dates.

    =COUNTA(A1:A10)

    Conditional Counting: Combine with IF to count based on a condition.

    =COUNT(IF(A1:A10>5, A1:A10))  // Array formula, press Ctrl+Shift+Enter

***4. COUNTA***

COUNTA counts the number of non-empty cells in a range. It is useful for counting all data entries, including text and numbers, which is helpful in many reporting and data-cleaning tasks.
Syntax Explanation:

    =COUNTA(value1, [value2], …)

    value1, [value2], …: The values or ranges to count.

**Advanced Usage:**

    Count Specific Types of Data: Combine with ISTEXT, ISNUMBER, etc., to count specific types of data.

    =SUM(IF(ISTEXT(A1:A10), 1, 0))  // Array formula, press Ctrl+Shift+Enter

    Counting Non-Blanks in Specific Conditions: Use COUNTA with criteria.

    =COUNTA(A1:A10) – COUNTBLANK(A1:A10)

****5. IF****

The IF function performs a logical test and returns one value if true and another if false. It is essential for conditional calculations, data validation, and dynamic data analysis.
Syntax Explanation:

    =IF(logical_test, value_if_true, value_if_false)

    logical_test: The condition to test.
    value_if_true: The value to return if the condition is true.
    value_if_false: The value to return if the condition is false.

**Advanced Usage:**

    1. Nested IF Statements: Handle multiple conditions by nesting IF statements.

    =IF(A1>10, “High”, IF(A1>5, “Medium”, “Low”))

    2. Combine with AND/OR: Use with AND/OR for more complex conditions.

    =IF(AND(A1>5, B1<10), “Yes”, “No”)

****6. VLOOKUP****

VLOOKUP searches for a value in the first column of a table and returns a value in the same row from a specified column. It is vital for looking up and retrieving data, such as finding prices for products.
Syntax Explanation:

    =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])

    lookup_value: The value to search for.
    table_array: The table to search within.
    col_index_num: The column number to return a value from.
    [range_lookup]: TRUE for approximate match, FALSE for exact match.

**Advanced Usage:**

    1. Exact vs. Approximate Match: Understand the importance of the [range_lookup] parameter.

    =VLOOKUP(F1, A1:C10, 3, FALSE)

    2. Dynamic Column Index: Use MATCH to dynamically determine the column index.

    =VLOOKUP(F1, A1:E10, MATCH(“ColumnName”, A1:E1, 0), FALSE)

***7. HLOOKUP***

HLOOKUP searches for a value in the first row of a table and returns a value in the same column from a specified row. It is useful for horizontally oriented data lookups, like finding student scores across multiple tests.
Syntax Explanation:

    =HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])

    lookup_value: The value to search for.
    table_array: The table to search within.
    row_index_num: The row number to return a value from.
    [range_lookup]: TRUE for approximate match, FALSE for exact match.

**Advanced Usage:**

    Exact vs. Approximate Match: Similar to VLOOKUP, use [range_lookup] appropriately.

    =HLOOKUP(G1, A1:J3, 2, TRUE)

    Dynamic Row Index: Combine with MATCH for a dynamic row index.

    =HLOOKUP(G1, A1:J10, MATCH(“RowName”, A1:A10, 0), FALSE)

***8. INDEX***

INDEX returns the value of a cell in a specified row and column within a range. It is powerful for more flexible lookups and can be used in combination with other functions like MATCH for dynamic data retrieval.
Syntax Explanation:

    =INDEX(array, row_num, [column_num])

    array: The range to search within.
    row_num: The row number to look in.
    [column_num]: The column number to look in.

**Advanced Usage:**

    Dynamic Ranges: Use with MATCH to create dynamic lookups.

    =INDEX(A1:C10, MATCH(G1, A1:A10, 0), 2)

    Two-Dimensional Lookups: Retrieve data from a two-dimensional range.

    =INDEX(A1:C10, 4, 2)

***9. MATCH***

MATCH searches for a value in a range and returns the relative position of that value. It is often used with INDEX for advanced lookups and data retrieval.
Syntax Explanation:

    =MATCH(lookup_value, lookup_array, [match_type])

    lookup_value: The value to search for.
    lookup_array: The range to search within.
    [match_type]: 0 for exact match, 1 for less than, -1 for greater than.

**Advanced Usage:**

    Dynamic Indexing: Combine with INDEX for flexible lookups.

    =INDEX(A1:C10, MATCH(G1, A1:A10, 0), 2)

    Approximate Matching: Use [match_type] for approximate matches in sorted data.

    =MATCH(H1, A1:A10, 1)

***10. SUMIF***

SUMIF adds the cells specified by a given condition or criteria, combining summation with conditional logic. It is ideal for tasks like summing sales from a specific region or time period.
Syntax Explanation:

    =SUMIF(range, criteria, [sum_range])

    range: The range to apply the criteria.
    criteria: The condition to meet.
    [sum_range]: The range to sum.

**Advanced Usage:**

    Multiple Criteria: Use SUMIFS for multiple conditions.

    =SUMIFS(B1:B10, A1:A10, “>5”, C1:C10, “<10”)

    Using Wildcards: Use * and ? in the criteria for partial matching.

    =SUMIF(A1:A10, “Region*”, B1:B10)

***11. COUNTIF***

COUNTIF counts the number of cells that meet a specified condition, useful for simple frequency counts, such as counting occurrences of a specific value or condition in a dataset.
Syntax Explanation:

    =COUNTIF(range, criteria)

    range: The range to apply the criteria.
    criteria: The condition to meet.

Advanced Usage:

    Multiple Conditions: Use COUNTIFS for multiple criteria.

    =COUNTIFS(A1:A10, “>5”, B1:B10, “<10”)

    Using Wildcards: Use * and ? in criteria for partial matching.

    =COUNTIF(C1:C10, “A*”)

***12. AVERAGEIF***

AVERAGEIF calculates the average of cells that meet a given condition, useful for calculating conditional averages, such as the average sales of a particular product or in a specific region.
Syntax Explanation:

    =AVERAGEIF(range, criteria, [average_range])

    range: The range to apply the criteria.
    criteria: The condition to meet.
    [average_range]: The range to average.

Advanced Usage:

    Multiple Criteria: Use AVERAGEIFS for multiple conditions.

    =AVERAGEIFS(B1:B10, A1:A10, “>5”, C1:C10, “<10”)

    Using Functions in Criteria: Use functions within criteria for dynamic conditions.

    =AVERAGEIF(A1:A10, “>” & TODAY()-30, B1:B10)

***13. SUMPRODUCT***

SUMPRODUCT multiplies corresponding components in the given arrays and returns the sum of those products. It is versatile for various calculations, such as weighted averages and conditional sums.
Syntax Explanation:

    =SUMPRODUCT(array1, [array2], …)

    array1, [array2], …: The arrays to multiply and sum.

Advanced Usage:

    1. Conditional Sums: Use logical conditions within arrays.

    =SUMPRODUCT((A1:A10>5)*(B1:B10))

    2. Weighted Averages: Calculate weighted averages directly.

    =SUMPRODUCT(A1:A10, B1:B10) / SUM(B1:B10)

***14. LEFT***

The LEFT function extracts a specified number of characters from the left side of a text string. It is useful for parsing data, such as extracting area codes from phone numbers or prefixes from codes.
Syntax Explanation:

    =LEFT(text, [num_chars])

    text: The text string to extract from.
    [num_chars]: The number of characters to extract (default is 1).

Advanced Usage:

    1. Extracting Fixed-Width Data: Use LEFT to parse fixed-width text fields.

    =LEFT(A1, 5)

    2. Dynamic Character Extraction: Combine with FIND to extract variable lengths.

    =LEFT(A1, FIND(“-“, A1)-1)

***15. RIGHT***

The RIGHT function extracts a specified number of characters from the right side of a text string. It is useful for parsing data, such as extracting file extensions or the last digits of codes.
Syntax Explanation:

    =RIGHT(text, [num_chars])

    text: The text string to extract from.
    [num_chars]: The number of characters to extract (default is 1).

Advanced Usage:

    Extracting File Extensions: Use RIGHT to get file extensions from file names.

    =RIGHT(B1, 3)

    Combining with LEN and FIND: Extract text dynamically based on length and position.

    =RIGHT(B1, LEN(B1) – FIND(” “, B1))

***16. YEAR***

The YEAR function extracts the year from a date. It is useful for date analysis, such as determining the year part of sales dates or events.
Syntax Explanation:

    =YEAR(date)

    date: The date from which to extract the year.

Advanced Usage:

    Combining with TODAY: Calculate the current year.

    =YEAR(TODAY())

    Using with Other Date Functions: Combine with DATE to manipulate dates.

    =DATE(YEAR(A1), 1, 1)  // Get the first day of the year

***17. MONTH***

The MONTH function extracts the month from a date. It is useful for date analysis, such as determining the month part of sales dates or events.
Syntax Explanation:

    =MONTH(date)

    date: The date from which to extract the month.

Advanced Usage:

    Combining with TODAY: Calculate the current month.

    =MONTH(TODAY())

    Using with Other Date Functions: Combine with DATE to manipulate dates.

=DATE(YEAR(A1), MONTH(A1), 1)  // Get the first day of the month




