using OfficeOpenXml;

namespace Svero.HelloExcel
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length < 1 || string.IsNullOrWhiteSpace(args[0]))
            {
                Console.Error.WriteLine("Please specify the full name including the path of the Excel document and " +
                    "optionally the name of the sheet you want to process.");
                Environment.Exit(1);
            }

            var filename = args[0];
            if (!File.Exists(filename))
            {
                Console.Error.WriteLine($"The specified Excel document \"{filename}\" was not found");
                Environment.Exit(2);
            }

            try
            {
                var sheetName = string.Empty;
                if (args.Length >= 2 && !string.IsNullOrWhiteSpace(args[1]))
                {
                    sheetName = args[1];
                }

                using var package = new ExcelPackage(new FileInfo(filename));

                var sheet = TryGetWorksheet(sheetName, package);

                if (sheet == null)
                {
                    Console.WriteLine("Sheet not found");
                    Environment.Exit(3);
                }

                var row = 2;
                while (true) // Dangerous
                {
                    var transaction = ConvertRowToBankTransaction(row, sheet);
                    if (transaction == null)
                    {
                        break;
                    }

                    Console.WriteLine(transaction);

                    row++;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred while processing \"{filename}\": {ex.Message}");
                Environment.Exit(2);
            }
        }

        private static BankTransaction? ConvertRowToBankTransaction(int rowNumber, ExcelWorksheet sheet)
        {
            if (rowNumber < 1)
            {
                throw new ArgumentOutOfRangeException("rowNumber");
            }

            if (sheet == null)
            {
                throw new ArgumentNullException("sheet");
            }

            if (sheet.Cells[rowNumber, 1].Value == null)
            {
                return default;
            }

            var transaction = new BankTransaction
            {
                PlanDate = sheet.Cells[rowNumber, 1].GetValue<DateTime>(),
                ActualDate = sheet.Cells[rowNumber, 2].Value == null ?
                default(DateTime?) : sheet.Cells[rowNumber, 2].GetValue<DateTime>(),
                Description = sheet.Cells[rowNumber, 3].Text,
                Value = sheet.Cells[rowNumber, 4].GetValue<Decimal>()
            };

            return transaction;
        }

        private static ExcelWorksheet? TryGetWorksheet(string? sheetName, ExcelPackage package)
        {
            ExcelWorksheet? sheet;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                sheet = package.Workbook.Worksheets[0];
            }
            else
            {
                sheet = package.Workbook.Worksheets[sheetName];
            }

            return sheet;
        }
    }
}