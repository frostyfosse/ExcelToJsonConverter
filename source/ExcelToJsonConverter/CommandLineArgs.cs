using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelToJsonConverter
{
    public class CommandLineArgs
    {
        [Option(shortName:'s', Required = true, HelpText = "The full file path to the spreadsheet. (i.e. C:\\MyFolder\\MyFile.xlsx)")]
        public string SpreadsheetFilePath { get; set; }

        [Option(shortName: 't', Required = true, HelpText = "The tab name within the spreadsheet to output. Example, \"My Tab Name\"")]
        public string WorkSheetName { get; set; }

        [Option(shortName: 'o', Required = true, HelpText = "The full file path of the output file. (i.e. C:\\MyFolder\\MyFile.json)")]
        public string OutputPath { get; set; }

        [Option(shortName: 'h', Required = false, HelpText = "Displays the parameters that are available.")]
        public bool Help { get; set; }

        /// <summary>
        /// Writes all command line arguments with their description to the console.
        /// </summary>
        public static void DisplayHelp()
        {
            var argsType = typeof(CommandLineArgs);
            var list = new List<OptionAttribute>();
            foreach (var prop in argsType.GetType().GetProperties())
            {
                var opt = prop.GetCustomAttribute<OptionAttribute>(true);
                if (opt == null)
                    continue;

                list.Add(opt);
            }

            Console.WriteLine();
            foreach (var opt in list.OrderBy(o => o.LongName))
            {
                string str;

                //if (opt.DefaultValue == null)
                str = string.Format("/{0},/{1}:\r\n   {2}", opt.LongName, opt.ShortName, opt.HelpText);
                //else
                //    str = string.Format("{0}|{1}\r\n {2} (Default='{3}')", opt.LongName, opt.ShortName, opt.HelpText, opt.DefaultValue);

                Console.WriteLine(str);
                Console.WriteLine();
            }
        }
    }
}
