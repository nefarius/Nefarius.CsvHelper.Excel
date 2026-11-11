# <img src="assets/NSS-128x128.png" align="left" />Nefarius.CsvHelper.Excel

[![Build](https://github.com/nefarius/Nefarius.CsvHelper.Excel/actions/workflows/build.yml/badge.svg)](https://github.com/nefarius/Nefarius.CsvHelper.Excel/actions/workflows/build.yml)
![Requirements](https://img.shields.io/badge/Requires-.NET%20Standard%202.0%20%7C%202.1-blue.svg)
[![NuGet](https://img.shields.io/nuget/v/Nefarius.CsvHelper.Excel)](https://www.nuget.org/packages/Nefarius.CsvHelper.Excel/)
[![NuGet downloads](https://img.shields.io/nuget/dt/Nefarius.CsvHelper.Excel)](https://www.nuget.org/packages/Nefarius.CsvHelper.Excel/)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)

Read and write Excel workbooks with [CsvHelper](https://github.com/JoshClose/CsvHelper) record mapping and [ClosedXML](https://github.com/ClosedXML/ClosedXML).

[NuGet package](https://www.nuget.org/packages/Nefarius.CsvHelper.Excel/) · [Source](https://github.com/nefarius/Nefarius.CsvHelper.Excel) · [API documentation](docs/index.md)

```powershell
dotnet add package Nefarius.CsvHelper.Excel
```

The examples require version 34. That version has not been published yet, so this command currently installs version 33.

The package continues to target **.NET Standard 2.0 and 2.1**. Building the repository requires the .NET 10 SDK; running all tests also requires the .NET 8 runtime.

## Reading

Mark record types with CsvHelper 34's `[CsvRecord]` attribute and make them `partial` so its source generator can create their mappings:

```csharp
using CsvHelper;
using Nefarius.CsvHelper.Excel;

[CsvRecord]
public partial class Person
{
    public string Name { get; set; }
    public int Age { get; set; }
}
```

Use `ExcelParser.Reader` to access CsvHelper's reader. The first worksheet is selected by default; pass a worksheet name to select another one.

```csharp
using var parser = new ExcelParser("people.xlsx");
var people = parser.Reader.GetRecords<Person>().ToArray();
```

For streams and reader options:

```csharp
using var stream = File.OpenRead("people.xlsx");
using var parser = new ExcelParser(stream, "People", new CsvReaderOptions
{
    FormatProvider = CultureInfo.InvariantCulture,
    ShouldSkipRecord = row => row.ToArray().All(string.IsNullOrWhiteSpace)
}, leaveOpen: true);

await foreach (var person in parser.Reader.GetRecordsAsync<Person>())
{
    Console.WriteLine(person.Name);
}
```

Both file and stream constructors own their input by default. Pass `leaveOpen: true` to retain ownership of a stream. Dispose the `ExcelParser` after enumeration completes.

CsvHelper 34 validates generated mappings eagerly. For worksheets that may be completely empty, call `parser.Reader.ReadHeader()` and check `parser.Reader.Headers.Count` before requesting typed records.

To bound conversion work and CSV buffering for sparse sheets, the rectangle from A1 to the last used row and column must contain at most 1,000,000 fields. Larger rectangles throw `InvalidDataException` before conversion.

## Writing

Use `ExcelWriter.Writer` to access CsvHelper's writer. Disposing the owner saves the workbook; flushing or disposing the inner `Writer` alone does not save it.

```csharp
using (var excel = new ExcelWriter("people.xlsx", "People"))
{
    excel.Writer.WriteRecords(people);
}
```

The path constructors validate worksheet names and writer options before overwriting existing files. To keep an output stream open:

```csharp
using var stream = new MemoryStream();
await using (var excel = new ExcelWriter(stream, "People", new CsvWriterOptions
{
    FormatProvider = CultureInfo.InvariantCulture
}, leaveOpen: true))
{
    await excel.Writer.WriteRecordsAsync(people);
}

var bytes = stream.ToArray();
```

`DisposeAsync` flushes CsvHelper asynchronously; ClosedXML workbook serialization remains synchronous. Excel values are written as text, including strings beginning with `=`. CsvHelper formatting, generated mappings, header settings, and injection options apply.

## Migrating from CsvHelper 33

This is an API-breaking update intended for a new major package release:

- Replace `new CsvReader(parser)` with `parser.Reader`.
- Replace calls inherited from `CsvWriter`, such as `excel.WriteRecords(...)`, with `excel.Writer.WriteRecords(...)`.
- Replace `CsvConfiguration` with `CsvReaderOptions` or `CsvWriterOptions`.
- Replace `Context.AutoMap<T>()` and legacy class maps with CsvHelper 34's generated mappings or explicit mapping APIs. Record types used by `GetRecords<T>()` and `WriteRecords<T>()` require `[CsvRecord]` and `partial`.
- Dispose the adapter owner to release its stream and, for writing, save the workbook.

CsvHelper 34 seals its reader and writer and removes `IParser`. The adapter therefore transports worksheet values through an internal UTF-8 CSV buffer instead of inheriting or implementing those APIs. Both adapters buffer intermediate CSV in memory; account for this additional memory when processing large workbooks. CSV dialect and encoding settings are normalized for this transport; cell delimiters, quotes, Unicode, whitespace, and embedded line breaks remain cell values. ClosedXML normalizes CRLF cell line endings to LF.

ClosedXML 0.105.1 uses OpenXML SDK 3.x, addressing the OpenXML 2/3 dependency conflict described in [issue #1](https://github.com/nefarius/Nefarius.CsvHelper.Excel/issues/1). The .NET 10 regression suite also installs Clippit to verify coexistence with its OpenXML dependency.

## Development and releases

```powershell
dotnet build -c Release
dotnet test -c Release -p:LibraryTargetFramework=netstandard2.0
dotnet test -c Release -p:LibraryTargetFramework=netstandard2.1
dotnet pack src/CsvHelper.Excel/CsvHelper.Excel.csproj -c Release -p:GeneratePackageOnBuild=false --output artifacts
```

CI tests both library assets on .NET 8 and .NET 10, on Windows and Linux, and packs the library only after those tests pass. Dependabot checks NuGet packages, the local documentation tool, and GitHub Actions weekly.

Push a `v`-prefixed version tag to release. `publish-to-nuget.yml` runs the same validation workflow and publishes its package artifact using NuGet trusted publishing for account `Nefarius`. Its NuGet.org policy must match repository owner `nefarius`, repository `Nefarius.CsvHelper.Excel`, and workflow `publish-to-nuget.yml`, without an environment restriction. No stored `NUGET_API_KEY` is used.

Restore and run the pinned documentation tool with:

```powershell
dotnet tool restore
dotnet tool run xmldoc2md ./bin/CsvHelper.Excel/Release/netstandard2.0/Nefarius.CsvHelper.Excel.dll ./docs/
```

## Sources & 3rd party credits

This library benefits from these awesome projects ❤ (appearance in no special order):

- [CsvHelper.Excel](https://github.com/hurles/CsvHelper.Excel) by Chris Young, continued by hurles, Joost Meulenkamp, and other contributors
- [CsvHelper](https://github.com/JoshClose/CsvHelper)
- [ClosedXML](https://github.com/ClosedXML/ClosedXML)

## Credits

- This package continues the Apache-2.0-licensed [CsvHelper.Excel](https://github.com/hurles/CsvHelper.Excel) project. Its current release brings that work to CsvHelper 34 and ClosedXML 0.105.
