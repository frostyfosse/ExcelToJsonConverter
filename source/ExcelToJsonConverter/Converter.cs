using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelToJsonConverter
{
    public static class Converter
    {
        public static void Convert(string spreadsheetFilePath, int workSheetIndex, string jsonFilePath)
        {
            var fileInfo = new FileInfo(spreadsheetFilePath);
            var package = new ExcelPackage(fileInfo);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[workSheetIndex];
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

            File.WriteAllText(jsonFilePath, jArray.ToString());
        }
    }
}
