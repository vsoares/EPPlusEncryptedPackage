using System;
using System.IO;
using OfficeOpenXml;

namespace EPPlusSimulation
{
    static class CreateExcelUsingMemoryStream
    {
        public static void Run()
        {
            string outputFilename = string.Format("C:\\Temp\\EPPlusSimulation\\{0}.xlsx", DateTime.Now.ToString("yyyymmddss"));

            // Create Excel package
            ExcelPackage excelPackage = new ExcelPackage();
            var worksheet = excelPackage.Workbook.Worksheets.Add("Worksheet1");

            // Encrypt package
            excelPackage.Encryption.IsEncrypted = true;

            // Save to MemoryStream
            MemoryStream outputStream = new MemoryStream();
            excelPackage.SaveAs(outputStream);
            Console.WriteLine("Output stream length: {0}", outputStream.Length);

            // Save to FileStream
            outputStream.Seek(0, SeekOrigin.Begin);
            FileStream outputFileStream = new FileStream(outputFilename, FileMode.CreateNew);
            outputStream.WriteTo(outputFileStream);
            Console.WriteLine("Output file stream length: {0}", outputFileStream.Length);

            // Excel package bytes
            Console.WriteLine("Excel package length: {0}", excelPackage.Stream.Length);
            Console.WriteLine("Excel package byte array length: {0}", excelPackage.GetAsByteArray().Length);

            // Close the streams
            outputFileStream.Close();
            outputStream.Close();

            // Open excel file using the operating system
            System.Diagnostics.Process.Start(outputFilename);

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
