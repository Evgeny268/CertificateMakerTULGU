using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Класс для работы с документами Microsoft Excel
/// </summary>
namespace CertificateMaker.core.office
{
    class ExcelWorker
    {
        /// <summary>
        /// Чтение из Excel необходимых ячеек
        /// </summary>
        /// <param name="filepath">путь до файла на диске</param>
        /// <param name="startRow">строка, с которое следует начать чтение</param>
        /// <param name="endRow">строка, на которое следует завершить чтение</param>
        /// <param name="cellsNum">номера столбцов, которые следует прочитать</param>
        /// <returns></returns>
        public static List<string[]> ReadCells(string filepath, int startRow, int endRow, int[] cellsNum)
        {
            List<string[]> returnCellsVaule = new List<string[]>();
            Excel.Application app = null;
            try
            {
                app = new Excel.Application();
                Excel.Workbook workbook = app.Workbooks.Open(filepath, 0, true);
                Excel._Worksheet worksheet = workbook.Sheets[1];
                Excel.Range range = worksheet.UsedRange;

                for (int i = startRow; i <= endRow; i++)
                {
                    string[] row = new string[cellsNum.Length];
                    for (int j = 0; j < cellsNum.Length; j++)
                    {
                        if (range.Cells[i, cellsNum[j]] != null && range.Cells[i, cellsNum[j]].Value2 != null)
                        {
                            row[j] = range.Cells[i, cellsNum[j]].Value2.ToString();
                        }
                        else
                        {
                            row[j] = "NULL";
                        }
                    }
                    returnCellsVaule.Add(row);
                }
                workbook.Close();
            }
            finally
            {
                try
                {
                    app.Quit();
                }
                catch (Exception) { }
            }
            return returnCellsVaule;
        }
    }
}
