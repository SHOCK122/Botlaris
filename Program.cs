using System;
using System.IO;
using OfficeOpenXml

namespace Botlaris
{
    class Program
    {
        string lastUserInput;
        static void Main(string[] args)
        {
            string templateFilePath;
            Console.WriteLine("Hello!");
            Console.WriteLine("Welcome to Botlaris!");
            Console.WriteLine("Please select the file you want to use as the template!");

            Console.WriteLine("Hello! This application lets you write application entries!");
            Console.WriteLine("Please Enter The text File: ");
            templateFilePath = Console.ReadLine();
            try
            {
                if (File.Exists(templateFilePath) && File.type)
                {
                    Console.WriteLine("File Select: " + templateFilePath);

                    var fi = new FileInfo(@"c:\workbooks\myworkbook.xlsx")

    using (var p = new ExcelPackage(fi))
                    {
                        //Get the Worksheet created in the previous codesample. 
                        var ws = p.Workbook.Worksheets["MySheet"];
                        //Set the cell value using row and column.
                        ws.Cells[2, 1].Value = "This is cell A2. It is set to bolds";
                        //The style object is used to access most cells formatting and styles.
                        ws.Cells[2, 1].Style.Font.Bold = true;
                        //Save and close the package.
                        p.Save();
                    }
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("That file does not exist!");
            }
            catch (DirectoryNotFoundException)
            {
                Console.WriteLine("Directory does not exist!");
            }
            catch (IOException)
            {
                Console.WriteLine("Nonsense!");
            }
            Console.ReadKey();
        }


        static string GetInput()
        {
            string userInput = Console.ReadLine();
            return userInput;

        }

    }
}
