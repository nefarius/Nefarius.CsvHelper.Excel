# ExcelWriter

Namespace: Nefarius.CsvHelper.Excel

Used to write CSV files.

```csharp
public class ExcelWriter : CsvHelper.CsvWriter, CsvHelper.IWriter, CsvHelper.IWriterRow, System.IDisposable, System.IAsyncDisposable
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → CsvWriter → [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md)<br>
Implements IWriter, IWriterRow, [IDisposable](https://docs.microsoft.com/en-us/dotnet/api/system.idisposable), IAsyncDisposable

## Properties

### <a id="properties-configuration"/>**Configuration**

```csharp
public IWriterConfiguration Configuration { get; }
```

#### Property Value

IWriterConfiguration<br>

### <a id="properties-context"/>**Context**

```csharp
public CsvContext Context { get; }
```

#### Property Value

CsvContext<br>

### <a id="properties-headerrecord"/>**HeaderRecord**

```csharp
public String[] HeaderRecord { get; }
```

#### Property Value

[String[]](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### <a id="properties-index"/>**Index**

```csharp
public int Index { get; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### <a id="properties-row"/>**Row**

```csharp
public int Row { get; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

## Constructors

### <a id="constructors-.ctor"/>**ExcelWriter(String)**

Initializes a new instance of the [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md) class.

```csharp
public ExcelWriter(string path)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The path.

### <a id="constructors-.ctor"/>**ExcelWriter(String, CultureInfo)**

Initializes a new instance of the [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md) class.

```csharp
public ExcelWriter(string path, CultureInfo culture)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The path.

`culture` [CultureInfo](https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)<br>
The culture.

### <a id="constructors-.ctor"/>**ExcelWriter(String, String)**

Initializes a new instance of the [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md) class.

```csharp
public ExcelWriter(string path, string sheetName)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The path.

`sheetName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The sheet name

### <a id="constructors-.ctor"/>**ExcelWriter(String, String, CultureInfo)**

Initializes a new instance of the [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md) class.

```csharp
public ExcelWriter(string path, string sheetName, CultureInfo culture)
```

#### Parameters

`path` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The path.

`sheetName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The sheet name

`culture` [CultureInfo](https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)<br>
The culture.

### <a id="constructors-.ctor"/>**ExcelWriter(Stream, CultureInfo, Boolean)**

Initializes a new instance of the [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md) class.

```csharp
public ExcelWriter(Stream stream, CultureInfo culture, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://docs.microsoft.com/en-us/dotnet/api/system.io.stream)<br>
The stream.

`culture` [CultureInfo](https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)<br>
The culture.

`leaveOpen` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>
`true` to leave the [TextWriter](https://docs.microsoft.com/en-us/dotnet/api/system.io.textwriter) open after the [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md)
 object is disposed, otherwise `false`.

### <a id="constructors-.ctor"/>**ExcelWriter(Stream, String, CultureInfo, Boolean)**

Initializes a new instance of the [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md) class.

```csharp
public ExcelWriter(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen)
```

#### Parameters

`stream` [Stream](https://docs.microsoft.com/en-us/dotnet/api/system.io.stream)<br>
The stream.

`sheetName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
The sheet name

`culture` [CultureInfo](https://docs.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo)<br>
The culture.

`leaveOpen` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>
`true` to leave the [TextWriter](https://docs.microsoft.com/en-us/dotnet/api/system.io.textwriter) open after the [ExcelWriter](./nefarius.csvhelper.excel.excelwriter.md)
 object is disposed, otherwise `false`.

## Methods

### <a id="methods-dispose"/>**Dispose(Boolean)**

```csharp
protected void Dispose(bool disposing)
```

#### Parameters

`disposing` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### <a id="methods-flush"/>**Flush()**

```csharp
public void Flush()
```

### <a id="methods-flushasync"/>**FlushAsync()**

```csharp
public Task FlushAsync()
```

#### Returns

[Task](https://docs.microsoft.com/en-us/dotnet/api/system.threading.tasks.task)

### <a id="methods-nextrecord"/>**NextRecord()**

```csharp
public void NextRecord()
```

### <a id="methods-nextrecordasync"/>**NextRecordAsync()**

```csharp
public Task NextRecordAsync()
```

#### Returns

[Task](https://docs.microsoft.com/en-us/dotnet/api/system.threading.tasks.task)

### <a id="methods-writefield"/>**WriteField(String, Boolean)**

```csharp
public void WriteField(string field, bool shouldQuote)
```

#### Parameters

`field` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`shouldQuote` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>
