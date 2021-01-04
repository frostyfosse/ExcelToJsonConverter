using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelToJsonConverter
{
    public static class Converter
    {
        static ExcelPackage GetPackage(string spreadsheetFilePath)
        {
            var fileInfo = new FileInfo(spreadsheetFilePath);
            
            return new ExcelPackage(fileInfo);
        }

        static ExcelWorksheet GetWorksheet(string spreadsheetFilePath, int workSheetIndex)
        {
            return GetPackage(spreadsheetFilePath).Workbook.Worksheets[workSheetIndex];
        }

        static ExcelWorksheet GetWorksheet(string spreadsheetFilePath, string workSheetName)
        {
            return GetPackage(spreadsheetFilePath).Workbook?
                                                  .Worksheets?
                                                  .FirstOrDefault(x => x.Name.Equals(workSheetName, StringComparison.OrdinalIgnoreCase));
        }

        static void Convert(ExcelWorksheet workSheet, string jsonFilePath)
        {
            bool headerRowCollected = false;
            int endColumnIndex = workSheet.Dimension.End.Column;
            var jArray = new JArray();
            var headers = new List<string>();

            for (int i = workSheet.Dimension.Start.Row;
                     i <= workSheet.Dimension.End.Row;
                     i++)
            {
                var jObject = new JObject();

                for (int j = workSheet.Dimension.Start.Column;
                         j <= endColumnIndex;
                         j++)
                {
                    object cellValue = workSheet.Cells[i, j].Value;

                    if (!headerRowCollected && cellValue == null)
                    {
                        endColumnIndex = j - 1;
                        break;
                    }

                    if (!headerRowCollected)
                        headers.Add(cellValue as string);
                    else
                        jObject.Add(new JProperty(headers[j - 1], cellValue));
                }

                if (!headerRowCollected)
                    headerRowCollected = true;
                else
                    jArray.Add(jObject);
            }

            if (!jsonFilePath.EndsWith(".json"))
            {
                throw new Exception($@"A valid json file name must be specified 
                                       for the output path. For example, 'C:\MyDirectory\MyFile.json' 
                                       (received '{jsonFilePath}')");
            }

            var directoryPath = Path.GetDirectoryName(jsonFilePath);

            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            File.WriteAllText(jsonFilePath, jArray.ToString());
        }

        public static void Convert(string spreadsheetFilePath, int workSheetIndex, string jsonFilePath)
        {
            var workSheet = GetWorksheet(spreadsheetFilePath, workSheetIndex);

            Convert(workSheet, jsonFilePath);
        }

        public static void Convert(string spreadsheetFilePath, string workSheetName, string jsonFilePath)
        {
            var workSheet = GetWorksheet(spreadsheetFilePath, workSheetName);

            Convert(workSheet, jsonFilePath);
        }
    }
}
