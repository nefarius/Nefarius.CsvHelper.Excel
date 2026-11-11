using System.Globalization;
using System.Linq;
using CsvHelper;

namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingPathSpecWithBlankRow : ExcelParserSpec
    {
        public ParseUsingPathSpecWithBlankRow() : base("parse_by_path_with_blank_row", includeBlankRow: true)
        {
            var csvConfiguration = new CsvReaderOptions
            {
                FormatProvider = CultureInfo.InvariantCulture,
                ShouldSkipRecord = x => x.ToArray().All(string.IsNullOrWhiteSpace)
            };
            using var parser = new ExcelParser(Path, null, csvConfiguration);
            Run(parser);
        }
    }
}
