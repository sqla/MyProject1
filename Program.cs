using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
namespace SpecEdit
{
    class FilesDirectory
    {
        public string[] trtfiles;
        public string xls;

        public FilesDirectory()
        {
         string uriPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
         string localPath = new Uri(uriPath).LocalPath;
         string[] xlsfiles = Directory.GetFiles(localPath, "*.xls");
         string[] titfiles = Directory.GetFiles(localPath, "*.tit");
         if (xlsfiles.Length != 1) { Console.WriteLine("Ошибка: в каталоге существует {0} excel-файлов. Excel-файл должен быть один. Нажмите любую клавишу для выхода.", xlsfiles.Length);  Console.ReadKey(); Environment.Exit(0); }
         if (titfiles.Length != 0) { Console.WriteLine("Ошибка: в каталоге существует {0} tit-файлов. Удалите их перед созданием новых. Нажмите любую клавишу для выхода.",  titfiles.Length); Console.ReadKey(); Environment.Exit(0); }
         xls = xlsfiles[0];
         trtfiles = Directory.GetFiles(localPath, "*.trt");
        }
    }


    class ExcelFile
    {
        private string mFileName;
        private Excel.Workbook mWorkbook;
        private Excel.Application mExcel;
        public string[,] list;
        public int maxcolumn;
        public int maxrow;

        public ExcelFile(string xls)
        {
            this.Open(xls);
            this.GetData();
            this.CloseExcel();
        }

        public void Open(string fileName)
        {
            mFileName = fileName;

            try
            {
                if (!File.Exists(mFileName))
                {
                    Console.WriteLine("Ошибка открытия файла Excel: файл {0} не найден. Нажмите любую клавишу для выхода.", mFileName);
                    Console.ReadKey();
                    Environment.Exit(0);
                }
                    
                System.Threading.Thread.CurrentThread.CurrentCulture =
                    System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

                mExcel = new Excel.Application(); //открыть эксель
                mExcel.Visible = false;
                mExcel.DisplayAlerts = false; //чтобы не показывало "сохранить"
                object missing = Type.Missing;

                mWorkbook = mExcel.Workbooks.Open(mFileName, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing, missing,
                    missing, missing);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine(ex);
                CloseExcel();
                Console.WriteLine("Ошибка открытия файла Excel. Нажмите любую клавишу для выхода.");
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        protected Excel.Workbook Workbook
        {
            get { return mWorkbook; }
        }

        private void CloseExcel()
        {
            if (mExcel != null)
            {
                mExcel.Workbooks.Close();
                mExcel.Quit();
                mExcel = null;
            }
        }

        public void GetData()
        {
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)mWorkbook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            maxcolumn = (int)lastCell.Column;
            maxrow = (int)lastCell.Row;
            list = new string[maxcolumn, maxrow]; // массив значений с листа равен по размеру листу

            Console.Write("Считывается поле №:");
            for (int i = 0; i < maxcolumn; i++)
            {
                if (i == 1 || i == 3) { continue; }
                Console.Write(" {0}", i + 1);
                for (int j = 0; j < maxrow; j++)
                {
                    list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();
                }
                Console.Write(",");
            }
            Console.WriteLine();
            Console.WriteLine("Успех");
            Console.WriteLine();
        }
    }

    class TrtFile
    {
        private string mFileName;
        private FileStream mTrt;
        public List<string> column1;
        public List<string> column2;
        private StreamReader sr;



        public TrtFile(string trt)
        {
            this.Open(trt);
            this.GetData();
            this.CloseTrt();
        }

        public void Open(string fileName)
        {
            mFileName = fileName;
            try
            {
                if (!File.Exists(mFileName))
                {
                    Console.WriteLine("Ошибка открытия trt: файл {0} не найден. Нажмите любую клавишу для выхода.", mFileName);
                    Console.ReadKey();
                    Environment.Exit(0);
                }

                System.Threading.Thread.CurrentThread.CurrentCulture =
                    System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

                mTrt = new FileStream(mFileName, FileMode.Open, FileAccess.Read);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine(ex);
                mTrt.Close();
                mTrt = null;
                Console.WriteLine("Ошибка открытия файла trt. Нажмите любую клавишу для выхода.");
                Console.ReadKey();
                Environment.Exit(0);
            }
        }

        private void CloseTrt()
        {
            if (mTrt != null)
            {
                mTrt.Close();
                mTrt = null;
            }
            if (sr != null)
            {
                sr.Close();
                sr = null;
            }
        }

        public void GetData()
        {
            Console.WriteLine("Считывается trt-файл: {0}", mFileName);
            sr = new StreamReader(mTrt);
            column1 = new List<string>();
            column2 = new List<string>();

            for (int i=1; i < 9 && !sr.EndOfStream; i++) sr.ReadLine(); //начинаем обработку с 9-ой строки
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                var values = line.Split(';');

                column1.Add(values[0]);
                column2.Add(values[1]);
            }

            Console.WriteLine("Успех");
        }

    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("СТАРТ");
            FilesDirectory fd = new FilesDirectory();
            Console.WriteLine("Обрабатывается excel-файл: {0}", fd.xls);
            ExcelFile xlsfile = new ExcelFile(fd.xls);

            foreach (string trt in fd.trtfiles) //получаем имена файлов trt в директории утилиты
                {
                TrtFile trtfile = new TrtFile(trt); //получаем данные из trt

              string tit = trt.Replace(".trt", ".tit");
                Console.WriteLine("Создание tit-файла: {0}", tit);
                 var sw = new StreamWriter(tit, true, System.Text.Encoding.GetEncoding(1251));

                 for (int j = 0; j < xlsfile.maxrow; j++)
                 {
                     sw.WriteLine("{0};{1};{2};{3};{4};{5}", xlsfile.list[0, j], trtfile.column1[j], xlsfile.list[2, j], trtfile.column2[j], xlsfile.list[4, j], xlsfile.list[5, j]);
                 }

                sw.Close();
                sw = null;
                Console.WriteLine("Успех");
                Console.WriteLine();
            }
            Console.WriteLine("ФИНИШ. Нажмите любую клавишу для выхода.");
            Console.ReadKey();

            }
        }
}