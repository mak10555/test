using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows;

namespace TSP.Export
{
    class ResultAlgorithm
    {
        public class Result
        {
            public double Time { get; set; }
            public double Length { get; set; }

            public Result(double time, double length)
            {
                Time = time;
                Length = length;
            }
        };
        
        public string Name { get; private set; }
        public int CountCities { get; private set; }

        private List<Result> _data;

        public Result this[int i]
        {
            get { return _data[i]; }
            set { _data[i] = value; }
        }

        public int CountData
        {
            get { return _data.Count; }
        }

        public ResultAlgorithm (int countCities, string name)
        {
            CountCities = countCities;
            Name = name;
            _data = new List<Result>();
        }

        public void Add(double time, double length)
        {
            _data.Add(new Result(time, length));
        }
    }

    class ExcelReport
    {
        public static void Save(ResultAlgorithm[] data)
        {
            Excel.Application excelApp = null;
            Worksheet workSheet = null;
            string path = null;

            try
            {
                //создаём новое Excel приложение
                excelApp = new Excel.Application();

                //добавляем рабочую книгу
                excelApp.Workbooks.Add();

                //обращаемся к активному листу (по умолчанию он первый)
                workSheet = (Worksheet)excelApp.ActiveSheet;

                //добавляем строку в Excel файл
                workSheet.Cells[1, 1] = "Вершин";

                for (int j = 0; j < data[0].CountData; j ++)
                    workSheet.Cells[j + 2, 1] = data[0].CountCities;

                for (int i = 0; i < data.Length; i++)
                {
                    int c = (i + 1) * 2;
                    workSheet.Cells[1, c] = data[i].Name + " Длина";
                    workSheet.Cells[1, c + 1] = data[i].Name + " Время";
                    for (int j = 0; j < data[i].CountData; j++)
                    {
                        ResultAlgorithm.Result res = data[i][j];
                        workSheet.Cells[j + 2, c] = res.Length;
                        workSheet.Cells[j + 2, c + 1] = res.Time;
                    }
                }
                workSheet.Columns.AutoFit();

                object misValue = System.Reflection.Missing.Value;

                //Сохранение в Excel файл;
                path = AppDomain.CurrentDomain.BaseDirectory + data[0].CountCities + "_TSP.xlsx";

                workSheet.SaveAs(path, 51);

            }
            catch (COMException)
            {
                MessageBox.Show("Не сохранено. Закройте документ " + path);
            }
            finally
            {
                excelApp.Quit();
                ServiceMethods.ReleaseComObject(excelApp);
                ServiceMethods.ReleaseComObject(workSheet);
            }
        }


    }
}
