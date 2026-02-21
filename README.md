# ExcelParser

A TypeScript library for reading and parsing Excel files built on top of [ExcelJS](https://github.com/exceljs/exceljs). Provides a clean, typed API for iterating worksheet rows without worrying about the underlying Excel internals.

---

## Installation

```bash
npm install exceljs
```

---

## Core Concepts

### Key Resolution

When parsing headers, the library resolves each column's key in the following order:

1. **Normalize** — trim whitespace, collapse multiple spaces, strip control characters
2. **Length check** — if the key exceeds `maxKeyLength`, fall back to column address
3. **Identifier check** — if the value isn't a usable string, fall back to column address
4. **Collision check** — if the key already exists, fall back to column address

The column address format is the column letter followed by the row number e.g. `A1`, `B12`.

### `useFirstRowAsHeader`

When `true` (default), the first row is treated as the header row and its values become the keys for all subsequent rows. The header row itself is not included in `foreach` results.

When `false`, no row is treated as a header — all rows are returned as data and keys are resolved from cell values directly, falling back to column addresses.

---

## `ExcelParser`

The main class for parsing a single worksheet.

### `ExcelParser.readFile(config)`

Static factory method. Always use this instead of the constructor directly.

```typescript
const parser = await ExcelParser.readFile({
    excelFilePath: "./data.xlsx",
    worksheet: "Sheet1"
});
```

**Config options:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `excelFilePath` | `string` | required | Path to the `.xlsx` file |
| `worksheet` | `string` | required | Name of the worksheet to parse |
| `useFirstRowAsHeader` | `boolean` | `true` | Treat first row as header |
| `maxKeyLength` | `number` | `128` | Maximum key length before falling back to address |
| `filterOption` | `FilterOption` | none | Include or exclude specific columns |

---

### `foreach(callback, limit?)`

Iterates over each data row. Accepts an optional `limit` to stop after a set number of rows.

```typescript
parser.foreach((row, controller) => {
    console.log(row.getValueAsString("Name"));
}, 100);
```

**Callback receives:**
- `row` — a `RowEvent` instance (see below)
- `controller` — a `ForEachController` instance (see below)

---

### `ForEachController`

Passed as the second argument to the `foreach` callback.

| Method | Returns | Description |
|--------|---------|-------------|
| `abort()` | `void` | Stops iteration immediately |
| `currentRowIndex()` | `number` | The current row's position in the sheet |
| `totalRows()` | `number` | Total row count in the sheet |

```typescript
parser.foreach((row, controller) => {
    if (row.getValueAsString("Status") === "Inactive") {
        controller.abort(); // stop as soon as we hit an inactive row
    }
    console.log(`Row ${controller.currentRowIndex()} of ${controller.totalRows()}`);
});
```

---

### `RowEvent`

Wraps a single parsed row and provides typed accessors.

| Method | Returns | Description |
|--------|---------|-------------|
| `getValueAsString(key, default?)` | `string \| null` | Returns string value or default |
| `getValueAsNumber(key, default?)` | `number \| null` | Returns number value or default |
| `getValueAsDate(key, default?)` | `Date \| null` | Returns Date value or default |
| `getValueAsBoolean(key, default?)` | `boolean \| null` | Returns boolean value or default |
| `getRawValue(key)` | `any` | Returns raw untyped value |
| `hasKey(key)` | `boolean` | Checks if key exists in the row |
| `getKeys()` | `string[]` | Returns all keys in the row |
| `getAll()` | `Record<string, any>` | Returns the raw row object |
| `foreach(callback)` | `void` | Iterates key-value pairs of the row |

```typescript
parser.foreach((row) => {
    const name = row.getValueAsString("Name", "Unknown");
    const age = row.getValueAsNumber("Age", 0);
    const dob = row.getValueAsDate("Date Added");
    const active = row.getValueAsBoolean("Active");

    // iterate all key-value pairs
    row.foreach((key, value) => {
        console.log(`${key}: ${value}`);
    });
});
```

---

### `toJson()`

Returns all rows as a plain array of objects. Keys follow the same resolution logic as `foreach`.

```typescript
const rows = parser.toJson();
// [{ Name: "Person 1", Age: 25, ... }, ...]
```

---

### `getHeaders()`

Returns the internal header map. Note: headers are only populated after `foreach` or `toJson` has been called at least once.

```typescript
parser.toJson();
const headers = parser.getHeaders();
// Map<number, HeaderData>
```

---

### `getName()`

Returns the worksheet name.

```typescript
console.log(parser.getName()); // "Sheet1"
```

---

### `writeJsonFile(filepath, format?)`

Writes the parsed worksheet data to a JSON file.

```typescript
parser.writeJsonFile("./output/data.json");          // formatted by default
parser.writeJsonFile("./output/data.json", false);   // minified
```

---

### `setFilter(filters, mode?)`

Updates the active filter after construction. **Important:** calling this resets all parsed headers, so any subsequent `foreach` or `toJson` call will re-parse from scratch with the new filter applied.

```typescript
parser.setFilter([{ header: "Department" }], "includes");
```

---

## Filtering Columns

Use `filterOption` in the config to include or exclude specific columns.

```typescript
const parser = await ExcelParser.readFile({
    excelFilePath: "./data.xlsx",
    worksheet: "Sheet1",
    filterOption: {
        mode: "includes",   // or "excludes"
        filters: [
            { header: "Name" },
            { header: "Age" },
            { header: "B", useColumnLetter: true } // filter by column letter
        ]
    }
});
```

**`Filter` options:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `header` | `string` | required | Header value or column letter to match |
| `useColumnLetter` | `boolean` | `false` | Match by column letter instead of header value |

**Modes:**
- `"includes"` — only columns matching the filter are parsed
- `"excludes"` — all columns are parsed except those matching the filter

> You can only use one mode at a time. `includes` and `excludes` are mutually exclusive.

---

## `ExcelDocument`

Wraps an entire workbook and provides access to all worksheets at once.

```typescript
const doc = new ExcelDocument({
    excelFilePath: "./data.xlsx"
});
```

**Config options:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `excelFilePath` | `string` | required | Path to the `.xlsx` file |
| `worksheetOptions` | `ExcelParserOptions` | `{}` | Options passed to each `ExcelParser` |
| `throwErrors` | `boolean` | `true` | Whether to throw on file read errors |
| `filterOptions` | `WorksheetFilterOptions` | none | Include or exclude specific worksheets by name |

---

### `findWorksheet(name, options?)`

Returns an `ExcelParser` for a specific worksheet by name.

```typescript
const parser = await doc.findWorksheet("Sheet1");
```

---

### `foreach(callback)`

Iterates over each allowed worksheet as an `ExcelParser`.

```typescript
await doc.foreach((parser) => {
    console.log(parser.getName());
    console.log(parser.toJson());
});
```

---

### `parseAsJson()`

Returns all worksheets as a `WorksheetData[]` array.

```typescript
const data = await doc.parseAsJson();
// [{ name: "Sheet1", headers: [...], rows: [...] }, ...]
```

---

### `writeJsonFile(config)`

Writes the entire workbook to a JSON file.

```typescript
// Single file
await doc.writeJsonFile({
    filepath: "./output/data.json",
    format: true
});

// Separate file per worksheet
await doc.writeJsonFile({
    filepath: "./output/data.json",
    seperateWorksheetFiles: true
});
// produces: ./output/data_Sheet1.json, ./output/data_Sheet2.json, etc.
```

---

### Filtering Worksheets

Use `filterOptions` to control which worksheets are processed.

```typescript
const doc = new ExcelDocument({
    excelFilePath: "./data.xlsx",
    filterOptions: {
        mode: "includes",
        filters: ["Sheet1", "Sheet3"]  // only process these sheets
    }
});
```

---

## Full Example

```typescript
import path from "path";

// Single worksheet
const parser = await ExcelParser.readFile({
    excelFilePath: path.resolve(__dirname, "./data.xlsx"),
    worksheet: "Sheet1",
    filterOption: {
        mode: "includes",
        filters: [
            { header: "Name" },
            { header: "Department" },
            { header: "Status" }
        ]
    }
});

parser.foreach((row, controller) => {
    const name = row.getValueAsString("Name", "Unknown");
    const dept = row.getValueAsString("Department", "");
    const status = row.getValueAsString("Status", "");

    console.log({ name, dept, status });

    if (name === "Person 5") {
        controller.abort();
    }
});

// Entire workbook
const doc = new ExcelDocument({
    excelFilePath: path.resolve(__dirname, "./data.xlsx"),
    filterOptions: {
        mode: "excludes",
        filters: ["Summary"] // skip the summary sheet
    }
});

const json = await doc.parseAsJson();
console.log(json);
```
