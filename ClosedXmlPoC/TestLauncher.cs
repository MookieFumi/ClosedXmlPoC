using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Xunit;
using Xunit.Abstractions;

namespace ClosedXmlPoC
{
    public class TestLauncher
    {
        private readonly ITestOutputHelper _output;

        public TestLauncher(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void CreateExcelFile()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), "Qux.xlsx");
            CreateExcelFile(path, 2, 2);
            using (var workbook = new XLWorkbook(path, XLEventTracking.Disabled))
            using (var worksheet = workbook.Worksheet(1))
            {
                int lastCellColumn = worksheet.LastColumnUsed().ColumnNumber();
                int lastCellRow = worksheet.LastRowUsed().RowNumber();
                _output.WriteLine($"{nameof(lastCellRow)} {lastCellRow}");
                _output.WriteLine($"{nameof(lastCellColumn)} {lastCellColumn}");
                var range = worksheet.Range(2, 1, lastCellRow, lastCellColumn);
                var stopwatch = Stopwatch.StartNew();

                //OK
                foreach (var column in range.Columns())
                {
                    //column.SetDataType(XLCellValues.Number);
                    foreach (var cell in column.CellsUsed())
                    {
                        _output.WriteLine(cell.Address.ToString());
                        cell.SetDataType(XLCellValues.Number);
                        //cell.SetDataValidation().Decimal.Between(0, 5);
                    }
                }

                //// KO
                //// https://twitter.com/panicoenlaxbox/status/884712707039121408
                //foreach (var column in worksheet.Columns())
                //{
                //    column.SetDataType(XLCellValues.Text);
                //}
                _output.WriteLine(stopwatch.Elapsed.ToString("g"));
                workbook.Save();
                _output.WriteLine(stopwatch.Elapsed.ToString("g"));
                stopwatch.Stop();
            }
        }

        [Fact]
        public void ReadExcelFile()
        {
            var path = Path.Combine(Directory.GetCurrentDirectory(), "Qux.xlsx");
            _output.WriteLine(path);
            var stopwatch = Stopwatch.StartNew();
            using (var workbook = new XLWorkbook(path, XLEventTracking.Disabled))
            {
                using (var worksheet = workbook.Worksheet(1))
                using (var range = worksheet.RangeUsed())
                using (var table = range.AsTable())
                using (var dataRange = table.DataRange)
                {
                    var rows = dataRange.Rows().Select(row => new Baz()
                    {
                        C1 = row.Field(1).GetValue<string>(),
                        C2 = row.Field(2).GetValue<string>(),
                        C3 = row.Field(3).GetValue<string>(),
                        C4 = row.Field(4).GetValue<string>(),
                        C5 = row.Field(5).GetValue<string>(),
                        C6 = row.Field(6).GetValue<string>(),
                        C7 = row.Field(7).GetValue<string>(),
                        C8 = row.Field(8).GetValue<string>(),
                        C9 = row.Field(9).GetValue<string>(),
                        C10 = row.Field(10).GetValue<string>()
                    });
                    _output.WriteLine(rows.Count().ToString());
                }
                _output.WriteLine(stopwatch.Elapsed.ToString("g"));
                stopwatch.Stop();
            }
        }

        private static void CreateExcelFile(string path, int columns, int rows)
        {
            var table = new DataTable("Worksheet1");
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add($"C{i + 1}", typeof(string));
            }

            for (int i = 0; i < rows; i++)
            {
                var values = new object[columns];
                for (int j = 0; j < columns; j++)
                {
                    values[j] = (j + 1).ToString();
                }
                table.Rows.Add(values);
            }
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(table);
                workbook.SaveAs(path);
            }
        }

    }
}