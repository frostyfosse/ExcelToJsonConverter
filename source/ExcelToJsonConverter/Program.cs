using System;
using System.IO;

namespace ExcelToJsonConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            var directory = @"C:\Users\jfosse\Downloads\XlsxToJsonParser";
            Converter.Convert(Path.Combine(directory, "Sample.xlsx"), 0, Path.Combine(directory, "Sample.json"));
        }
    }
}
