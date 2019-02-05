using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using CommandLine;
using OfficeOpenXml;

namespace Xlsx2Csv
{

    enum Quote
    {
        WhenNeeded,
        Always,
        Never
    }

    class Options
    {
        [Value(0, Required = true, HelpText = "Input file to be processed.")]
        public string InputFileName { get; set; }

        [Value(1, Required = true, HelpText = "Worksheet name to be processed.")]
        public string WorksheetName { get; set; }

        [Value(2, Required = false, HelpText = "Output file to write data to.")]
        public string OutputFileName { get; set; }

        [Option("password", Required = false, HelpText = "Password for open xlsx file.")]
        public string Password { get; set; }

        [Option("encoding", Required = false, HelpText = "CSV file encoding.", Default = "utf-8")]
        public string Encoding { get; set; }

        [Option("separator", Required = false, HelpText = "CSV file separator.")]
        public string Separator { get; set; }

        [Option("language", Required = false, HelpText = "CSV file language (culture).")]
        public string Language { get; set; }

        [Option("use-quote", Required = false, HelpText = "When use quote (values WhenNeeded, Always, Never).", Default = Quote.WhenNeeded)]
        public Quote DataQuote { get; set; }

        [Option("header-quote", Required = false, HelpText = "When use quote in header row (values WhenNeeded, Always, Never).", Default = Quote.WhenNeeded)]
        public Quote HeaderQuote { get; set; }


        [Option("silent", Required = false, HelpText = "Set output to silent (no messages).")]
        public bool Silent { get; set; }

        [Option("verbose", Required = false, HelpText = "Set output to verbose messages.")]
        public bool Verbose { get; set; }
    }

    class Program
    {

        static void Main(string[] args)
        {
            try
            {
                Parser.Default.ParseArguments<Options>(args)
                    .WithParsed<Options>(o =>
                {
                    FileInfo inputFile = new FileInfo(o.InputFileName);
                    if (!inputFile.Exists)
                    {
                        Console.WriteLine("File not exists");
                        return;
                    }

                    if (String.IsNullOrEmpty(o.OutputFileName))
                        o.OutputFileName = Path.ChangeExtension(Path.Combine(Path.GetDirectoryName(o.InputFileName), o.WorksheetName), ".csv");

                    CultureInfo cultureInfo = String.IsNullOrEmpty(o.Language) ?
                        CultureInfo.CurrentCulture :
                        CultureInfo.GetCultureInfo(o.Language);

                    if (String.IsNullOrEmpty(o.Separator))
                        o.Separator = cultureInfo.TextInfo.ListSeparator;

                    if (!o.Silent)
                    {
                        Console.WriteLine("Reading   {0}", o.InputFileName);
                        Console.WriteLine("Sheet     {0}", o.WorksheetName);
                        Console.WriteLine("Output    {0}", o.OutputFileName);
                    }
                    using (ExcelPackage package =
                        String.IsNullOrEmpty(o.Password) ?
                        new ExcelPackage(inputFile) :
                        new ExcelPackage(inputFile, o.Password))
                    {
                        var sheet = package.Workbook.Worksheets[o.WorksheetName];
                        if (sheet.Dimension == null)
                        {
                            Console.WriteLine("Error: There is no dimension in this sheet.");
                            return;
                        }
                        Encoding encoding = Encoding.GetEncoding(o.Encoding);
                        if (o.Verbose)
                        {
                            Console.WriteLine("Encoding  {0}", encoding.EncodingName);
                            Console.WriteLine("Language  {0}", cultureInfo.Name);
                            Console.WriteLine("Separator {0}", o.Separator);
                            Console.WriteLine("Dimension {0}", sheet.Dimension.Address);
                        }
                        using (StreamWriter file = new StreamWriter(o.OutputFileName, false, encoding))
                        {
                            for (int row = sheet.Dimension.Start.Row; row <= sheet.Dimension.End.Row; row++)
                            {
                                if (o.Verbose)
                                    Console.Write("Row ..... {0}\r", row);

                                for (int col = sheet.Dimension.Start.Column; col <= sheet.Dimension.End.Column; col++)
                                {
                                    ExcelRange cell = sheet.Cells[row, col];
                                    string valueString = String.Format(cultureInfo, "{0}", cell.Value);

                                    // Quote selector
                                    Quote quote = row == sheet.Dimension.Start.Row ? o.HeaderQuote : o.DataQuote;

                                    // Use quote for cell
                                    bool useQuote = (cell.Value is string) &&
                                            (quote == Quote.Always ||
                                            (quote == Quote.WhenNeeded && 
                                                (valueString.Contains("\"") || valueString.Contains("\r") || valueString.Contains("\n"))));                                    

                                    // Write value
                                    if (useQuote)
                                        file.Write("\"{0}\"", valueString.Replace("\"", "\"\""));
                                    else
                                        file.Write("{0}", valueString);

                                    // Write separator
                                    if (col < sheet.Dimension.End.Column)
                                        file.Write(o.Separator);
                                }
                                file.WriteLine();
                            }
                            if (o.Verbose)
                                Console.WriteLine();
                            if (!o.Silent)
                                Console.WriteLine("Export done");

                        }
                    }
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }
            if (System.Diagnostics.Debugger.IsAttached)
                Console.ReadKey(false);
        }
    }
}
