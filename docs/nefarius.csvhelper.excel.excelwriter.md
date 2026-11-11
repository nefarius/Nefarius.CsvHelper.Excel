# ExcelWriter

Namespace: Nefarius.CsvHelper.Excel

Writes CsvHelper records into an Excel workbook when disposed.

```csharp
public sealed class ExcelWriter : System.IDisposable, System.IAsyncDisposable
```

Inheritance [Object](https://learn.microsoft.com/dotnet/api/system.object) → [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md)<br>
Implements [IDisposable](https://learn.microsoft.com/dotnet/api/system.idisposable), [IAsyncDisposable](https://learn.microsoft.com/dotnet/api/system.iasyncdisposable)

**Remarks:**

Use [ExcelWriter.Writer](./nefarius.csvhelper.excel.excelwriter.md#writer) to write records. Intermediate CSV is buffered in memory until disposal.

## Properties

### <a id="properties-writer"/>**Writer**

Gets the CsvHelper writer used to format fields and records.

```csharp
public CsvWriter Writer { get; }
```

#### Property Value

CsvWriter<br>

## Constructors

### <a id="constructors-.ctor"/>**ExcelWriter(String)**

Creates or overwrites a file with an export worksheet.

```csharp
public ExcelWriter(string path)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

### <a id="constructors-.ctor"/>**ExcelWriter(String, CultureInfo)**

Creates or overwrites a file using the specified culture.

```csharp
public ExcelWriter(string path, CultureInfo culture)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`culture` [CultureInfo](https://learn.microsoft.com/dotnet/api/system.globalization.cultureinfo)<br>

### <a id="constructors-.ctor"/>**ExcelWriter(String, String)**

Creates or overwrites a file with a named worksheet.

```csharp
public ExcelWriter(string path, string sheetName)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

### <a id="constructors-.ctor"/>**ExcelWriter(String, String, CultureInfo)**

Creates or overwrites a file with a named worksheet and culture.

```csharp
public ExcelWriter(string path, string sheetName, CultureInfo culture)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`culture` [CultureInfo](https://learn.microsoft.com/dotnet/api/system.globalization.cultureinfo)<br>

### <a id="constructors-.ctor"/>**ExcelWriter(String, String, CsvWriterOptions)**

Creates or overwrites a file using CsvHelper writer options.

```csharp
public ExcelWriter(string path, string sheetName, CsvWriterOptions options)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`options` CsvWriterOptions<br>

### <a id="constructors-.ctor"/>**ExcelWriter(Stream, CultureInfo, Boolean)**

Writes to a stream using the specified culture. Set leaveOpen to retain stream ownership.

```csharp
public ExcelWriter(Stream stream, CultureInfo culture, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://learn.microsoft.com/dotnet/api/system.io.stream)<br>

`culture` [CultureInfo](https://learn.microsoft.com/dotnet/api/system.globalization.cultureinfo)<br>

`leaveOpen` [Boolean](https://learn.microsoft.com/dotnet/api/system.boolean)<br>

### <a id="constructors-.ctor"/>**ExcelWriter(Stream, String, CultureInfo, Boolean)**

Writes a named worksheet to a stream using the specified culture.

```csharp
public ExcelWriter(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://learn.microsoft.com/dotnet/api/system.io.stream)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`culture` [CultureInfo](https://learn.microsoft.com/dotnet/api/system.globalization.cultureinfo)<br>

`leaveOpen` [Boolean](https://learn.microsoft.com/dotnet/api/system.boolean)<br>

### <a id="constructors-.ctor"/>**ExcelWriter(Stream, String, CsvWriterOptions, Boolean)**

Writes a named worksheet to a stream using CsvHelper writer options.

```csharp
public ExcelWriter(Stream stream, string sheetName, CsvWriterOptions options, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://learn.microsoft.com/dotnet/api/system.io.stream)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`options` CsvWriterOptions<br>

`leaveOpen` [Boolean](https://learn.microsoft.com/dotnet/api/system.boolean)<br>

**Remarks:**

CSV dialect options are normalized for the internal transport; formatting and mapping options apply.

## Methods

### <a id="methods-dispose"/>**Dispose()**

Finishes the records, saves the workbook and releases owned resources.

```csharp
public void Dispose()
```

### <a id="methods-disposeasync"/>**DisposeAsync()**

Flushes the records asynchronously, saves the workbook and releases owned resources.

```csharp
public ValueTask DisposeAsync()
```

#### Returns

[ValueTask](https://learn.microsoft.com/dotnet/api/system.threading.tasks.valuetask)

**Remarks:**

ClosedXML workbook serialization is synchronous.
