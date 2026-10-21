# 📊 Day 07 — Excel Lookup Functions

## 🎯 Objective

Learn how to find and retrieve information from a dataset using Excel lookup functions.

## 🛠️ Functions Covered

* `VLOOKUP`
* `HLOOKUP`
* `XLOOKUP`
* `LOOKUP`
* `INDEX` and `MATCH`

## 1. VLOOKUP

**VLOOKUP** searches for a value in the first column of a selected table and returns a value from another column in the same row.

**Syntax:**

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

**Example:** Find the department of an employee whose ID is in J2.

```excel
=VLOOKUP(J2,A2:H101,3,FALSE)
```

This searches for the Employee ID in column A and returns the Department from column C.

## 2. HLOOKUP

**HLOOKUP** searches for a value across the first row of a table and returns a value from a specified row.

**Syntax:**

```excel
=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])
```

It is useful when your data is arranged horizontally.

## 3. XLOOKUP

**XLOOKUP** searches for a value in one range and returns a matching value from another range. It can look left or right.

**Syntax:**

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
```

**Example:** Find an employee's name using their ID in J2.

```excel
=XLOOKUP(J2,A2:A101,B2:B101,"Not Found")
```

This searches column A for the employee ID and returns the corresponding name from column B. If there is no match, it displays "Not Found".

## 4. LOOKUP

**LOOKUP** searches for a value in a single row or column and returns a corresponding value from another row or column.

**Syntax:**

```excel
=LOOKUP(lookup_value, lookup_vector, [result_vector])
```

**Example:**

```excel
=LOOKUP(J2,A2:A101,B2:B101)
```

**Important:** LOOKUP assumes the lookup data is sorted in ascending order. For exact employee ID matches, XLOOKUP or VLOOKUP with `FALSE` is generally safer.

## 5. INDEX and MATCH

**INDEX** returns a value from a specified position in a range. **MATCH** finds the position of a value within a range. Together, they can perform flexible lookups.

**Example:** Find an employee's name using their ID in J2.

```excel
=INDEX(B2:B101,MATCH(J2,A2:A101,0))
```

MATCH finds the employee ID's position in column A. INDEX returns the name at that same position in column B.

## 🧪 Practical Exercise

Using the employee dataset, practise the following:

1. Find an employee's full name using their Employee ID.
2. Find their department using their Employee ID.
3. Find their monthly salary using their Employee ID.
4. Use XLOOKUP to display "Not Found" when an ID does not exist.
5. Use INDEX and MATCH to retrieve an employee's job title.
6. Explain the difference between VLOOKUP and XLOOKUP.
7. Identify why LOOKUP may return an unexpected result when the lookup data is unsorted.

## 📚 Key Lessons

* VLOOKUP searches vertically and returns values from columns to the right of the lookup column.
* HLOOKUP searches horizontally.
* XLOOKUP supports flexible lookups in either direction and can specify a result when no match is found.
* LOOKUP is most suitable for sorted lookup data.
* INDEX and MATCH work together for flexible lookups.

## 💡 Reflection

Lookup functions help analysts retrieve relevant information from large datasets quickly and accurately. Choosing the right function depends on how the data is arranged and whether exact matching is required.

**Day by day. Skill by skill. Project by project.** 🚀
