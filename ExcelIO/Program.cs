// See https://aka.ms/new-console-template for more information

using ExcelDataReader;

namespace ExcelIO;

internal class Program
{
    public static void Main(string[] args)
    {
        ReadExcel(@"Data/Book.xlsx");
    }
    
    static void  ReadExcel(string filePath)
    {
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            // Auto-detect format, supports:
            //  - Binary Excel files (2.0-2003 format; *.xls)
            //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                // Choose one of either 1 or 2:

                // 1. Use the reader methods
                do
                {
                    while (reader.Read())
                    {
                         
                        //print content of col A
                        var a= reader.GetString(0);
                        
                        Console.WriteLine(a);
                    }
                } while (reader.NextResult());
                
                // 2. Use the AsDataSet extension method
                var result = reader.AsDataSet();

                // The result of each spreadsheet is in result.Tables
            }
        }
    }
}