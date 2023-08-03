using System;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace FILECONVERSION
{
    public class Program
    {
        public static Task Main(string[] args)
        {
            string txtFilePath;
            try 
            {
                Console.WriteLine("Fetching the file...");
                //string file = "C:\\Users\\ggakenge\\Downloads\\Debit_Freeze.xlsx";

                //args = new string[] { "C:\\Users\\ggakenge\\Downloads\\Debit_Freeze.xlsx" };
                Console.WriteLine("Argument size: " + args.Length);
                //var fileName = file.FullName;
                
                if (args.Length != 1)
                {
                    Console.WriteLine("Please provide excel file location");
                    return Task.FromResult(0);
                }
                string file = args[0].ToString();

                using var package = new ExcelPackage(file);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var currentSheet = package.Workbook.Worksheets;
                var workSheet = currentSheet.First();
                var noOfCol = workSheet.Dimension.End.Column;
                var noOfRow = workSheet.Dimension.End.Row;
                Console.WriteLine("Reading the file...");
                string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString().Substring(10);

                txtFilePath = $@"C:\Users\{userName}\Documents\FreezeAcounts_" + DateTime.Now.ToString("ddMMyyyyhhmmtt") + ".txt";

                for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                {

                    string BankID = workSheet.Cells[rowIterator, 1].Value?.ToString() + "      ";
                    string AccountNo = workSheet.Cells[rowIterator, 2].Value?.ToString() + "                                                      "; 
                    string FreezeType = workSheet.Cells[rowIterator, 3].Value?.ToString().PadRight(26);
                    string FreezeReason = workSheet.Cells[rowIterator, 4].Value?.ToString().PadRight(300);
                    string RequestDpt = workSheet.Cells[rowIterator, 5].Value?.ToString().PadRight(170);
                    string combineCells = BankID.Replace("\n", String.Empty).Replace("\t", String.Empty).Replace("\r", String.Empty) +
                        AccountNo.Replace("\n", String.Empty).Replace("\t", String.Empty).Replace("\r", String.Empty) +
                        FreezeType.Replace("\n", String.Empty).Replace("\t", String.Empty).Replace("\r", String.Empty)
                        + FreezeReason.Replace("\n", String.Empty).Replace("\t", String.Empty).Replace("\r", String.Empty)
                         + RequestDpt.Replace("\n", String.Empty).Replace("\t", String.Empty).Replace("\r", String.Empty);
                    using (StreamWriter writer = new StreamWriter(txtFilePath, append: true))
                    {
                        writer.WriteLine(combineCells);
       
                        
                    }

                }

                Console.WriteLine("Successfully generated text file at: " + txtFilePath);
            }

            catch (Exception e) {
                Console.WriteLine(e.Message);   
            }
            
            return Task.CompletedTask;
        }
    }
}

