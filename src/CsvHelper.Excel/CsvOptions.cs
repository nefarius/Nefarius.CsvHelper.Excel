using System;

namespace Nefarius.CsvHelper.Excel;

// CsvHelper 34 embeds IsExternalInit in its .NET Standard assemblies but uses the BCL
// type in its modern assemblies. Direct init-setter calls compiled against the former
// throw MissingMethodException when NuGet selects the latter at runtime. Resolve the
// small set of transport options against the actual runtime type instead.
internal static class CsvOptions
{
    internal static T Configure<T>(T options, params (string Name, object Value)[] values)
    {
        foreach (var (name, value) in values)
        {
            typeof(T).GetProperty(name).SetValue(options, value);
        }
        return options;
    }
}
