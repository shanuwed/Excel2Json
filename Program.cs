using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;

namespace Excel2Json
{
    /*
     * Console program to convert an Excel file to Json file
     * 
    */
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("Excel2Json [input.xlsx] [output.json]");
                return;
            }

            string infilename = args[0];
            string outfilename = args[1];

            var dataTables = ExcelHelper.GetExcelTabData(infilename);
            var sheetData = new JArray();
            foreach(var kvPair in dataTables)
            {
                dynamic sheet = new JObject();
                sheet.sheetname = kvPair.Key; 
                sheet.data = JsonConvert.DeserializeObject(JsonConvert.SerializeObject(kvPair.Value));
                sheetData.Add(sheet);
            }

            dynamic output = new JObject();
            output.filename = infilename;
            output.sheets = sheetData;

            File.WriteAllText(outfilename, output.ToString());
        }
    }
}
