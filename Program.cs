using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TestConsoleApplication.Files_Handle;
//using Twilio;
//using Twilio.Rest.Api.V2010.Account;

namespace TestConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Process p = Process.GetCurrentProcess();
            ShowWindow(p.MainWindowHandle, 3); //SW_MAXIMIZE = 3            
            while (true)
            {
                int option = Print_List_Options();
                switch (option)
                {
                    case 1:
                        Create_Folders_For_Month_Everyday_Has_Folder();
                        break;
                    case 2:
                        Read_Cell_From_All_EXcel_Sheets_And_Save_In_TXT_File();
                        break;
                    case 3:
                        Read_TXT_File_And_Check_Every_Line_Exists_In_PDF_Files_And_Save_In_TXT_File();
                        break;
                    case 4:
                        Rename_Files_From_Text_File();
                        break;

                    case -1:
                        return;
                    default:
                        break;
                }
            }

            void Create_Folders_For_Month_Everyday_Has_Folder()
            {
                Console.Write("Please Enter Parent Folder (if Empty Current Folder Is Parent Folder) : ");
                string parentFolder = Console.ReadLine();
                Console.Write("Please Enter Year : ");
                int year;
                if (int.TryParse(Console.ReadLine(), out year))
                {
                    int month;
                    Console.Write("Please Enter Month : ");
                    if (int.TryParse(Console.ReadLine(), out month))
                    {
                        int days = DateTime.DaysInMonth(year, month);
                        for (int day = 1; day <= days; day++)
                        {
                            string folderName = parentFolder;
                            if (!string.IsNullOrEmpty(parentFolder))
                            {
                                folderName += @"\";
                            }
                            folderName += new DateTime(year, month, day).ToString("yyyy_MM_dd");

                            DirectoryInfo directoryInfo = Directory.CreateDirectory(folderName);
                            Console.WriteLine("Folder " + directoryInfo.FullName + " Created.");
                            System.Threading.Thread.Sleep(1000);
                        }
                    }
                }
            }

            void Read_Cell_From_All_EXcel_Sheets_And_Save_In_TXT_File()
            {
                Console.Write("Please Enter Excel File Path : ");
                string xlsFile = Console.ReadLine();
                Console.Write("Please Enter Excel Row Number : ");
                int row = int.Parse(Console.ReadLine());
                Console.Write("Please Enter Excel Col Number : ");
                int col = int.Parse(Console.ReadLine());
                List<string> data = ExcelFileHandle.GetCellData(xlsFile, row, col);

                Console.Write("Please Enter Text File Path : ");
                string txtFile = Console.ReadLine();
                TextFileHandle.SaveListString(data, txtFile);
            }            

            void Read_TXT_File_And_Check_Every_Line_Exists_In_PDF_Files_And_Save_In_TXT_File()
            {
                Console.Write("Please Enter TXT File Path (Input): ");
                string TXT_File_Input = Console.ReadLine();
                Console.Write("Please Enter Folder Path That Has PDF Files : ");
                string PDF_Folder = Console.ReadLine();
                Console.Write("Please Enter TEXT File Path (Output) : ");
                string TXT_File_Output = Console.ReadLine();
                List<string> lines = TextFileHandle.ReadListString(TXT_File_Input);
                string[] files = Directory.GetFiles(PDF_Folder);
                List<string> result_TXT_File = new List<string>();
                                    
                foreach (string file in files)
                {
                    if (Path.GetExtension(file).ToLower() != ".pdf")
                    {
                        continue;
                    }
                    string ResultLineSave = "";
                    string PDF_File_Content = PDF_FileHandle.ReadPdfFile(file);
                    foreach (string line in lines)
                    {
                        if (PDF_File_Content.Contains(line))
                        {
                            ResultLineSave = line + ", " + file;
                            break;
                        }
                    }
                    result_TXT_File.Add(ResultLineSave);
                    TextFileHandle.SaveListString(result_TXT_File, TXT_File_Output);
                }
            }

            void Rename_Files_From_Text_File()
            {
                Console.Write("Please Enter TXT File Path (Input): ");
                string TXT_File_Input = Console.ReadLine();
                List<string> lines = TextFileHandle.ReadListString(TXT_File_Input);
                foreach (string line in lines)
                {
                    string oldPath = line.Split(',')[1];
                    string newPath = Path.GetDirectoryName(oldPath) + "\\" + line.Split(',')[0] + Path.GetExtension(oldPath);
                    File.Move(oldPath, newPath);
                }
            }
        }
        
        static int Print_List_Options()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("****************************");
            Console.WriteLine("1) Create Folders For Month (One Folder for Everyday)");
            Console.WriteLine("2) Save EXCEL CELL From All Sheets in TXT File");
            Console.WriteLine("3) Check If Every Line In TXT File Exists PDF Files Result is TXT File ");
            Console.WriteLine("4) Rename Files From TXT File Depend On Result Number (3) ex : NewName, fullPathFile");
            Console.WriteLine("****************************");
            Console.Write("Please Enter Option : ");            
            int option = 0;
            if (int.TryParse(Console.ReadLine(), out option))
            {
                Console.ForegroundColor = ConsoleColor.White;
                return option;
            }
            else
            {
                Console.WriteLine("*** Please Enter Valid Option *** ");
                return Print_List_Options();
            }            
        }

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(System.IntPtr hWnd, int cmdShow);
    }

    class TestReadOnlyAndConstant
    {
        readonly int readOnly = 10;
        const int constant = 10;

        public TestReadOnlyAndConstant()
        {
            readOnly = 100;
        }

        public void Check()
        {
            Console.WriteLine("Read Only : {0}", readOnly);
            Console.WriteLine("Constant : {0}", constant);
        }

        public void ChangeValues(int value)
        {
            
        }
    }    
}
