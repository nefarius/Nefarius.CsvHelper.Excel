# ExcelParser

Namespace: Nefarius.CsvHelper.Excel

Opens an Excel worksheet for reading with CsvHelper.

```csharp
public sealed class ExcelParser : System.IDisposable
```

Inheritance [Object](https://learn.microsoft.com/dotnet/api/system.object) → [ExcelParser](./nefarius.csvhelper.excel.excelparser.md)<br>
Implements [IDisposable](https://learn.microsoft.com/dotnet/api/system.idisposable)

**Remarks:**

Use [ExcelParser.Reader](./nefarius.csvhelper.excel.excelparser.md#reader) for record mapping and enumeration. The worksheet is buffered as CSV in memory. Dispose this owner after reading.

## Properties

### <a id="properties-reader"/>**Reader**

Gets the CsvHelper reader over the selected worksheet.

```csharp
public CsvReader Reader { get; }
```

#### Property Value

CsvReader<br>

## Constructors

### <a id="constructors-.ctor"/>**ExcelParser(String)**

Opens the first worksheet of a file.

```csharp
public ExcelParser(string path)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

### <a id="constructors-.ctor"/>**ExcelParser(String, String)**

Opens a named worksheet of a file.

```csharp
public ExcelParser(string path, string sheetName)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

### <a id="constructors-.ctor"/>**ExcelParser(String, CultureInfo)**

Opens a file using the specified culture for cell values and conversions.

```csharp
public ExcelParser(string path, CultureInfo culture)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`culture` [CultureInfo](https://learn.microsoft.com/dotnet/api/system.globalization.cultureinfo)<br>

### <a id="constructors-.ctor"/>**ExcelParser(String, String, CultureInfo)**

Opens a named worksheet using the specified culture.

```csharp
public ExcelParser(string path, string sheetName, CultureInfo culture)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`culture` [CultureInfo](https://learn.microsoft.com/dotnet/api/system.globalization.cultureinfo)<br>

### <a id="constructors-.ctor"/>**ExcelParser(String, String, CsvReaderOptions)**

Opens a file using CsvHelper reader options.

```csharp
public ExcelParser(string path, string sheetName, CsvReaderOptions options)
```

#### Parameters

`path` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`options` CsvReaderOptions<br>

### <a id="constructors-.ctor"/>**ExcelParser(Stream, CultureInfo, Boolean)**

Opens the first worksheet of a stream. Set leaveOpen to retain ownership of the stream.

```csharp
public ExcelParser(Stream stream, CultureInfo culture, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://learn.microsoft.com/dotnet/api/system.io.stream)<br>

`culture` [CultureInfo](https://learn.microsoft.com/dotnet/api/system.globalization.cultureinfo)<br>

`leaveOpen` [Boolean](https://learn.microsoft.com/dotnet/api/system.boolean)<br>

### <a id="constructors-.ctor"/>**ExcelParser(Stream, String, CultureInfo, Boolean)**

Opens a named worksheet of a stream using the specified culture.

```csharp
public ExcelParser(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://learn.microsoft.com/dotnet/api/system.io.stream)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`culture` [CultureInfo](https://learn.microsoft.com/dotnet/api/system.globalization.cultureinfo)<br>

`leaveOpen` [Boolean](https://learn.microsoft.com/dotnet/api/system.boolean)<br>

### <a id="constructors-.ctor"/>**ExcelParser(Stream, String, CsvReaderOptions, Boolean)**

Opens a worksheet of a stream using CsvHelper reader options.

```csharp
public ExcelParser(Stream stream, string sheetName, CsvReaderOptions options, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://learn.microsoft.com/dotnet/api/system.io.stream)<br>

`sheetName` [String](https://learn.microsoft.com/dotnet/api/system.string)<br>

`options` CsvReaderOptions<br>

`leaveOpen` [Boolean](https://learn.microsoft.com/dotnet/api/system.boolean)<br>

**Remarks:**

CSV dialect and encoding options are normalized for the internal UTF-8 transport.

## Methods

### <a id="methods-dispose"/>**Dispose()**

Disposes the reader and, unless leaveOpen was set, the input stream.

```csharp
public void Dispose()
```
