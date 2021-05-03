using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Excel2Json
{
    public static class ExcelHelper
    {

        public static MemoryStream FetchFileFromDisk(string filename)
        {
            Debug.Assert(File.Exists(filename), "File must exist");

            MemoryStream memoryStream = new MemoryStream();
            using (var fileStream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                fileStream.CopyTo(memoryStream);
            }
            return memoryStream;
        }

        public static DataTable GetDataFromExcelSheet(ExcelWorksheet worksheet)
        {
            DataTable table = new DataTable();
            foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                table.Columns.Add(firstRowCell.Text);
            }
            var startRow = 2;
            for (int rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = worksheet.Cells[rowNum, 1, rowNum, table.Columns.Count];
                DataRow row = table.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }

            return table;
        }

        public static Dictionary<string, DataTable> GetDataTablesFromExcelPackage(ExcelPackage excelPackage)
        {
            var dataTables = new Dictionary<string, DataTable>();

            foreach (var worksheet in excelPackage.Workbook.Worksheets)
            {
                var dataTable = GetDataFromExcelSheet(worksheet);

                dataTables.Add(worksheet.Name, dataTable);
            }

            return dataTables;
        }

        public static Dictionary<string, DataTable> GetExcelTabData(string filename)
        {
            return GetExcelTabData(FetchFileFromDisk(filename));
        }

        public static Dictionary<string, DataTable> GetExcelTabData(MemoryStream stream)
        {
            Dictionary<string, DataTable> dataTables = new Dictionary<string, DataTable>();

            if (stream != null && stream.Length > 0)
            {
                using (ExcelPackage xlPackage = new ExcelPackage(stream))
                {
                    dataTables = GetDataTablesFromExcelPackage(xlPackage);
                }
            }
            return dataTables;
        }
    }
}
