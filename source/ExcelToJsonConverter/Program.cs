using System;
using System.IO;

namespace ExcelToJsonConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            var parameters = new CommandLineArgs();
            var showHelp = false;
            var exitCode = 0;

            if (CommandLine.Parser.Default.ParseArguments(args, parameters))
            {
                if (parameters.Help)
                {
                    showHelp = true;
                }
                else
                {
                    try
                    {
                        var parameterText = string.Join(Environment.NewLine, $"{nameof(parameters.SpreadsheetFilePath)}: '{parameters.SpreadsheetFilePath}'",
                                                                             $"{nameof(parameters.OutputPath)}: '{parameters.OutputPath}'",
                                                                             $"{nameof(parameters.WorkSheetName)}: '{parameters.WorkSheetName}'");

                        Console.WriteLine("Parameters received");
                        Console.WriteLine(parameterText);

                        Converter.Convert(parameters.SpreadsheetFilePath, parameters.WorkSheetName, parameters.OutputPath);
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine($"Error occurred while processing data: {e}");
                        exitCode = -1;
                    }
                }
            }
            else
            {
                showHelp = true;
            }

            if (showHelp)
                CommandLineArgs.DisplayHelp();

            Environment.Exit(exitCode);
        }

        /// <summary>
        /// Proof of concept method
        /// </summary>
        public static void SampleCallToConverter()
        {
            var directory = @"C:\MyDirectory\XlsxToJsonParser";
            Converter.Convert(Path.Combine(directory, "Sample.xlsx"), "My Excel Data", Path.Combine(directory, "Sample.json"));
        }
    }
}
