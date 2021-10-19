using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GemBox.Spreadsheet;

namespace ExcelProject1.Models
{
    class ExcelDataProvider
    {
        public async Task<List<TimedValueDBO>> GetXlsxData(string path)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile workbook = ExcelFile.Load(path);
            List <TimedValueDBO> output = new List<TimedValueDBO>();
            foreach (ExcelWorksheet worksheet in workbook.Worksheets)
            {
                foreach (ExcelRow row in worksheet.Rows)
                {
                    output.Add(new TimedValueDBO { DateTime = row.Cells[0].Value?.ToString() ?? "EMPTY", 
                        TagName = row.Cells[1].Value?.ToString() ?? "EMPTY", 
                        Type = row.Cells[2].Value?.ToString() ?? "EMPTY", 
                        Value = row.Cells[3].Value?.ToString() ?? "EMPTY"});
                }
            }
            return output;
        }
        public async Task<List<ValueDBO>> GetCsvData(string path)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            List<ValueDBO> output = new List<ValueDBO>();
            var workbook = ExcelFile.Load(path, new CsvLoadOptions(CsvType.CommaDelimited));
            var worksheet = workbook.Worksheets[0];
            foreach (ExcelRow row in worksheet.Rows)
            {
                output.Add(new ValueDBO
                {
                    TagName = row.Cells[0].Value?.ToString() ?? "EMPTY",
                    Type = row.Cells[1].Value?.ToString() ?? "EMPTY",
                    Value = row.Cells[2].Value?.ToString() ?? "EMPTY"
                });
            }
            return output;
        }
        public void SaveDataCsv(List<ValueDBO> values)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile workbook = new ExcelFile();
            var options = SaveOptions.CsvDefault;
            var worksheet = workbook.Worksheets.Add("Sheet");
            for (int index = 0; index < values.Count; index++)
            {
                worksheet.Cells[index, 0].SetValue(values[index].TagName);
                worksheet.Cells[index, 1].SetValue(values[index].Type);
                worksheet.Cells[index, 2].SetValue(values[index].Value);
            }
            workbook.Save("Output.csv", new CsvSaveOptions(CsvType.CommaDelimited));
        }
        public void SaveDataXlsx(List<TimedValueDBO> values)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("DataTable to Sheet");
            var dataTable = new DataTable();
            dataTable.Columns.Add("DateTime", typeof(string));
            dataTable.Columns.Add("TagName", typeof(int));
            dataTable.Columns.Add("Type", typeof(string));
            dataTable.Columns.Add("Value", typeof(string));
            for (int index = 0; index < values.Count; index++)
            {
                worksheet.Cells[index, 0].SetValue(values[index].DateTime);
                worksheet.Cells[index, 1].SetValue(values[index].TagName);
                worksheet.Cells[index, 2].SetValue(values[index].Type);
                worksheet.Cells[index, 3].SetValue(values[index].Value);
            }
            worksheet.InsertDataTable(dataTable,
            new InsertDataTableOptions()
            {
                ColumnHeaders = true,
                StartRow = 0
            });
            workbook.Save("Output.xlsx", SaveOptions.XlsxDefault);
        }
    }
}
