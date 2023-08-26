using System.Runtime.ConstrainedExecution;
using OfficeOpenXml;

namespace Svero.HelloExcel
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length < 1 || string.IsNullOrWhiteSpace(args[0])) {
                Console.Error.WriteLine("Please specify the path and name of the Excel document you want to read as argument");
                Environment.Exit(1);
            }

            string filename = args[0];
            try {
                using var package = new ExcelPackage(new FileInfo(filename));

                foreach (var sheet in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"Found worksheet \"{sheet.Name}\"");
                }
            } catch (Exception ex) {
                Console.Error.WriteLine($"An error occurred while processing {filename}: {ex.Message}");
                Environment.Exit(2);
            }
        }
    }
}