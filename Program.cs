using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Botlaris
{
    class Program
    {
        private static ExcelPackage tFile;
        private static bool done = false;
        private static string pathStr = "";
        private static string source;
        private static bool errorTrip = false;
        static void Main(string[] args)
        {
            Console.WriteLine("Hello!");
            Console.WriteLine("Welcome to Botlaris!");

            SelectFile();
            
            while (!done)
            {
                Console.WriteLine("\r\nPlease select a command:");
                Console.WriteLine("  move or m: move files into folders based on selected template file.");
                Console.WriteLine("  rename or r: rename without moving the files based on the selected template file.");
                Console.WriteLine("  undo or u: undo previous Botlaris actions (requires valid moveLog.txt).");
                Console.WriteLine("  new or n: select a new template file.");
                Console.WriteLine("  exit or e: quit Botlaris.");
                HandleInput(Console.ReadLine());
            }
        }

        private static void SelectFile()
        {
            bool validFile = false;
            pathStr = "";
            while (!validFile)
            {
                if (File.Exists(pathStr) && (Path.GetExtension(pathStr) == ".xls" || Path.GetExtension(pathStr) == ".xlsx"))
                {
                    source = Path.GetDirectoryName(pathStr);
                    FileInfo fi = new FileInfo(pathStr);
                    tFile = new ExcelPackage(fi);
                    Directory.CreateDirectory(source + @"\BOTLARIS_LOGS");
                    Log("File selected: " + Path.GetFullPath(pathStr));
                    validFile = true;
                }
                else if (pathStr.ToLower() == "exit")
                {
                    Exit();
                }
                else
                {
                    Console.WriteLine("Please select the file you want to use as the template!");
                    pathStr = Console.ReadLine();
                    pathStr = pathStr.Split('"').Length > 1 ? pathStr.Split('"')[1] : pathStr;
                }
            }
        }

        public static void Log(string logMessage, int logType = 0)
        {
            string message = $"{DateTime.Now.ToLongDateString()} {DateTime.Now.ToLongTimeString()}:{logMessage}\r\n";
            switch (logType)
            {
                case 0:
                    File.AppendAllText(source + @"\BOTLARIS_LOGS\errorLog.txt", message);
                    break;
                case 1:
                    File.AppendAllText(source + @"\BOTLARIS_LOGS\moveLog.txt", message);
                    break;
                case 2:
                    File.AppendAllText(source + @"\BOTLARIS_LOGS\undoLog.txt", message);
                    break;
                default:
                    File.AppendAllText(source + @"\BOTLARIS_LOGS\errorLog.txt", message);
                    break;
            }
            Console.WriteLine(logMessage);
        }

        static void HandleInput(string userInput)
        {
            switch (userInput.ToLower())
            {
                case "move":
                case "m":
                    MoveFiles();
                    break;
                case "undo":
                case "u":
                    UndoMove();
                    break;
                case "rename":
                case "r":
                    MoveFiles(true);
                    break;
                case "exit":
                case "e":
                    Exit();
                    break;
                case "new":
                case "n":
                    SelectFile();
                    break;
                default:
                    Console.WriteLine("Unrecognized command!");
                    break;
            }
        }

        static void MoveFiles(bool isRenaming = false)
        {
            string action = isRenaming ? "rename" : "move";
            Console.WriteLine($"You chose to {action} files:");
            ExcelWorksheet ws = tFile.Workbook.Worksheets[0];
            int row = 2;
            int col = 1;
            string val;
            string srcFile;
            string target = source;
            string trgName = "";
            string trgFile;
            List<string> suffices = new List<string>();
            int maxRow = GetMaxRow(ws);
            int maxCol = GetMaxCol(ws);
            while (row <= maxRow)
            {
                while (col <= maxCol)
                {
                    //skip empty cells
                    if (ws.Cells[row, col].Value == null)
                    {
                        col += 1;
                        continue;
                    }
                    val = ws.Cells[row, col].Value.ToString();
                    //set target folder name
                    if (col == 1 && row > 2)
                    {
                        target = Path.GetDirectoryName(pathStr) + @"\" + val;
                        trgName = val;
                        if (!Directory.Exists(target) && !isRenaming)
                        {
                            Directory.CreateDirectory(target);
                            Log($"Create folder: '{target}'", 1);
                        }
                    }
                    //do selected action
                    else if (row > 2 && col >1)
                    {
                        srcFile = source + @"\" + val;
                        string filename = @"\" + trgName + GetSuff(ws, col) + Path.GetExtension(val);
                        if (isRenaming)
                        {
                            trgFile = source + filename ;
                        }
                        else
                        {
                            trgFile = target + filename;
                        }
                        Mover(srcFile, trgFile, 1);
                    }
                    col += 1;
                }
                col = 1;
                row += 1;
            }
        }

        private static string GetSuff(ExcelWorksheet ws, int col)
        {
            return ws.Cells[2, col].Value != null ? ws.Cells[2, col].Value.ToString() : "";
        }

        private static int GetMaxCol(ExcelWorksheet ws)
        {
            int maxCol = 1;
            while (ws.Cells[1, maxCol].Value != null)
            {
                maxCol += 1;
            }
            return maxCol;
        }

        private static int GetMaxRow(ExcelWorksheet ws)
        {
            int maxRow = 3;
            while (ws.Cells[maxRow, 1].Value != null)
            {
                maxRow += 1;
            }
            return maxRow;
        }

        static void Mover(string sourceFile, string targetFile, int logType = 0)
        {
            try
            {
                File.Move(sourceFile, targetFile);
                Log($"Move file: '{sourceFile}' ~to~ '{targetFile}'", logType);
            }
            catch (System.IO.FileNotFoundException e)
            {
                errorTrip = true;
                Log(e.Message);
            }
            catch (Exception e)
            {
                errorTrip = true;
                Log(e.Message);
            }
        }

        static void UndoMove()
        {
            Stack folders = new Stack();
            Regex rFilePaths = new Regex(@"Move\sfile:\s\'(?<prev>.+)\'\s~to~\s\'(?<now>.+)\'", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            Regex rFolders = new Regex(@"Create\sfolder:\s\'(?<folder>.+)\'", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            Console.WriteLine("You chose to undo moving of files");
            if (File.Exists(source + @"\BOTLARIS_LOGS\moveLog.txt"))
            {
                foreach (string line in File.ReadAllLines(source + @"\BOTLARIS_LOGS\moveLog.txt"))
                {
                    Console.WriteLine($"Undoing {line}");
                    Match matchFiles = rFilePaths.Match(line);
                    string now = matchFiles.Groups["now"].ToString();
                    string prev = matchFiles.Groups["prev"].ToString();
                    Match matchFolders = rFolders.Match(line);
                    if (matchFiles.Captures.Count > 0)
                    {
                        Mover(now, prev, 2);
                    }
                    else if (rFolders.Match(line).Groups.Count > 0)
                    {
                        folders.Push(matchFolders.Groups["folder"]);
                    }
                }
                while (folders.Count > 0)
                {
                    string folder = folders.Pop().ToString();
                    try
                    {
                        Directory.Delete(folder);
                        Log($"Delete folder: '{folder}'", 2);
                    }
                    catch (Exception e)
                    {
                        errorTrip = true;
                        Log(e.Message);
                    }
                }
            }
            else
            {
                Console.WriteLine(@"No valid moveLog.txt file found! Ensure the file is located at ..\BOTLARIS_LOGS\moveLog.txt");
            }
        }
        static void Exit()
        {
            while (errorTrip)
            {
                Console.WriteLine($@"Errors have occurred. Please check {source}\errorLog.txt");
                Console.WriteLine($"Please type 'ack' to exit.");
                if (Console.ReadLine() == "ack")
                {
                    errorTrip = false;
                }
            }
            done = true;
        }
    }
}
