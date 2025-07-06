# ExcelParser

Namespace: Nefarius.CsvHelper.Excel

Parses an Excel file.

```csharp
public class ExcelParser : CsvHelper.IParser, System.IDisposable
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [ExcelParser](./nefarius.csvhelper.excel.excelparser.md)<br>
Implements IParser, [IDisposable](https://docs.microsoft.com/en-us/dotnet/api/system.idisposable)

## Properties

### <a id="properties-bytecount"/>**ByteCount**

```csharp
public long ByteCount { get; }
```

#### Property Value

[Int64](https://docs.microsoft.com/en-us/dotnet/api/system.int64)<br>

### <a id="properties-charcount"/>**CharCount**

```csharp
public long CharCount { get; }
```

#### Property Value

[Int64](https://docs.microsoft.com/en-us/dotnet/api/system.int64)<br>

### <a id="properties-configuration"/>**Configuration**

```csharp
public IParserConfiguration Configuration { get; }
```

#### Property Value

IParserConfiguration<br>

### <a id="properties-context"/>**Context**

```csharp
public CsvContext Context { get; }
```

#### Property Value

CsvContext<br>

### <a id="properties-count"/>**Count**

```csharp
public int Count { get; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### <a id="properties-delimiter"/>**Delimiter**

```csharp
public string Delimiter { get; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### <a id="properties-item"/>**Item**

```csharp
public string Item { get; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### <a id="properties-rawrecord"/>**RawRecord**

```csharp
public string RawRecord { get; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### <a id="properties-rawrow"/>**RawRow**

```csharp
public int RawRow { get; private set; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### <a id="properties-record"/>**Record**

```csharp
public String[] Record { get; private set; }
```

#### Property Value

[String[]](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### <a id="properties-row"/>**Row**

```csharp
public int Row { get; private set; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

## Constructors

### <a id="constructors-.ctor"/>**ExcelParser(String)**

Initializes a new instance of the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md) class.

```csharp
public ExcelParser(string path)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The path.

### <a id="constructors-.ctor"/>**ExcelParser(String, String)**

Initializes a new instance of the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md) class.

```csharp
public ExcelParser(string path, string sheetName)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The path.

`sheetName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The sheet name

### <a id="constructors-.ctor"/>**ExcelParser(String, CultureInfo)**

Initializes a new instance of the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md) class.

```csharp
public ExcelParser(string path, CultureInfo culture)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The path.

`culture` [CultureInfo](https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)<br>
The culture.

### <a id="constructors-.ctor"/>**ExcelParser(String, String, CultureInfo)**

Initializes a new instance of the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md) class.

```csharp
public ExcelParser(string path, string sheetName, CultureInfo culture)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The path.

`sheetName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The sheet name

`culture` [CultureInfo](https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)<br>
The culture.

### <a id="constructors-.ctor"/>**ExcelParser(Stream, CultureInfo, Boolean)**

Initializes a new instance of the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md) class.

```csharp
public ExcelParser(Stream stream, CultureInfo culture, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://docs.microsoft.com/en-us/dotnet/api/system.io.stream)<br>
The stream.

`culture` [CultureInfo](https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)<br>
The culture.

`leaveOpen` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>
`true` to leave the [TextWriter](https://docs.microsoft.com/en-us/dotnet/api/system.io.textwriter) open after the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md)
 object is disposed, otherwise `false`.

### <a id="constructors-.ctor"/>**ExcelParser(Stream, String, CultureInfo, Boolean)**

Initializes a new instance of the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md) class.

```csharp
public ExcelParser(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://docs.microsoft.com/en-us/dotnet/api/system.io.stream)<br>
The stream.

`sheetName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The sheet name

`culture` [CultureInfo](https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)<br>
The culture.

`leaveOpen` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>
`true` to leave the [TextWriter](https://docs.microsoft.com/en-us/dotnet/api/system.io.textwriter) open after the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md)
 object is disposed, otherwise `false`.

### <a id="constructors-.ctor"/>**ExcelParser(String, String, CsvConfiguration)**

Initializes a new instance of the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md) class.

```csharp
public ExcelParser(string path, string sheetName, CsvConfiguration configuration)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The stream.

`sheetName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The sheet name

`configuration` CsvConfiguration<br>
The configuration.

### <a id="constructors-.ctor"/>**ExcelParser(Stream, String, CsvConfiguration, Boolean)**

Initializes a new instance of the [ExcelParser](./nefarius.csvhelper.excel.excelparser.md) class.

```csharp
public ExcelParser(Stream stream, string sheetName, CsvConfiguration configuration, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://docs.microsoft.com/en-us/dotnet/api/system.io.stream)<br>
The stream.

`sheetName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The sheet name

`configuration` CsvConfiguration<br>
The configuration.

`leaveOpen` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>
Whether to leave the stream open when done or dispose it.

## Methods

### <a id="methods-dispose"/>**Dispose()**

```csharp
public void Dispose()
```

### <a id="methods-dispose"/>**Dispose(Boolean)**

```csharp
protected void Dispose(bool disposing)
```

#### Parameters

`disposing` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### <a id="methods-read"/>**Read()**

```csharp
public bool Read()
```

#### Returns

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)

### <a id="methods-readasync"/>**ReadAsync()**

```csharp
public Task<Boolean> ReadAsync()
```

#### Returns

[Task&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.threading.tasks.task-1)
