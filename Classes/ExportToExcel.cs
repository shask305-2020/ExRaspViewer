using System;
using System.Data;
using System.Windows.Forms;
using NSExcel = Microsoft.Office.Interop.Excel;


namespace ExRaspViewer.Classes
{
    public static class ExportToExcel
    {
        //Создание процесса Excel. Он будет висеть в памяти, пока основное приложение запущено
        public static NSExcel.Application AppExcel = new NSExcel.Application();
        private static int _currentRow = 0;
        public static string datN = "", datK = "";

        public static void Export(DataTable DataTable, string pedagog)
        {
            try
            {
                int ColumnsCount = DataTable.Columns.Count;   //Количество столбцов
                if (DataTable == null || ColumnsCount == 0)
                    throw new Exception("Экспорт в файл Excel: несуществующая или пустая входная таблица!\n");

                if (AppExcel.Workbooks.Count < 1)
                    AppExcel.Workbooks.Add();   //Добавление рабочей книги


                NSExcel._Worksheet Worksheet = AppExcel.ActiveSheet;    //Установление рабочего листа
                object[] Header = new object[ColumnsCount - 3];     //Массив для названий колонок (первые три с id не берём)
                for (int i = 0; i < ColumnsCount - 3; i++)
                    Header[i] = DataTable.Columns[i + 3].ColumnName;    //Заполнение массива названиями колонок

                if (_currentRow == 0)
                {
                    _currentRow++;
                    NSExcel.Range _excelCells = Worksheet.get_Range((NSExcel.Range)(Worksheet.Cells[_currentRow, 1]), (NSExcel.Range)(Worksheet.Cells[_currentRow, 7]));
                    _excelCells.Merge(Type.Missing);
                    Worksheet.Cells[_currentRow, 1] = $"Выполнение педагогической нагрузки за период с {datN} по {datK}";
                    _excelCells.Font.Bold = true;
                    _excelCells.HorizontalAlignment = NSExcel.XlHAlign.xlHAlignCenter;
                }

                _currentRow++;
                //Запись Ф.И.О. преподавателя в первую ячейку
                Worksheet.Cells[_currentRow, 2] = pedagog;
                NSExcel.Range teacher = Worksheet.Cells[_currentRow, 2];
                teacher.Font.Bold = true;

                _currentRow++;
                //Выделение диапазона ячеек для заполнения заголовками
                //Worksheet.Cells[номер строки, номер столбца]
                NSExcel.Range HeaderRange = Worksheet.get_Range((NSExcel.Range)(Worksheet.Cells[_currentRow, 1]), (NSExcel.Range)(Worksheet.Cells[_currentRow, ColumnsCount - 3]));
                HeaderRange.Value = Header; //Заполнение первой строки заголовками столбцов


                HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);  //Перекраска строки с заголовками
                HeaderRange.Font.Bold = true;   //Установление полужирного начертания для строки заголовков

                //Границы (все границы)
                HeaderRange.Cells.Borders[NSExcel.XlBordersIndex.xlEdgeBottom].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                HeaderRange.Cells.Borders[NSExcel.XlBordersIndex.xlEdgeLeft].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                HeaderRange.Cells.Borders[NSExcel.XlBordersIndex.xlEdgeRight].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                HeaderRange.Cells.Borders[NSExcel.XlBordersIndex.xlEdgeTop].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                HeaderRange.Cells.Borders[NSExcel.XlBordersIndex.xlInsideVertical].LineStyle = NSExcel.XlLineStyle.xlContinuous;

                int RowsCount = DataTable.Rows.Count;   //Получеие количества строк

                //Двуменрый массив для хранения данных (основные данные)
                object[,] Cells = new object[RowsCount, ColumnsCount - 3];

                //Заполнение двумерного массива данными из ячеек DataTable
                for (int j = 0; j < RowsCount; j++)
                    for (int i = 0; i < ColumnsCount - 3; i++)
                        Cells[j, i] = DataTable.Rows[j][i + 3];

                _currentRow++;
                int startRow = _currentRow;

                //Выделение диапазона ячеек для заполнения данными
                NSExcel.Range DataRange = Worksheet.get_Range((NSExcel.Range)(Worksheet.Cells[_currentRow, 1]), (NSExcel.Range)(Worksheet.Cells[_currentRow + RowsCount - 1, ColumnsCount - 3]));
                DataRange.Value = Cells;
                //Выделение границ таблицы (все границы)
                DataRange.Cells.Borders[NSExcel.XlBordersIndex.xlEdgeBottom].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                DataRange.Cells.Borders[NSExcel.XlBordersIndex.xlEdgeLeft].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                DataRange.Cells.Borders[NSExcel.XlBordersIndex.xlEdgeRight].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                DataRange.Cells.Borders[NSExcel.XlBordersIndex.xlEdgeTop].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                DataRange.Cells.Borders[NSExcel.XlBordersIndex.xlInsideVertical].LineStyle = NSExcel.XlLineStyle.xlContinuous;
                DataRange.Cells.Borders[NSExcel.XlBordersIndex.xlInsideHorizontal].LineStyle = NSExcel.XlLineStyle.xlContinuous;

                Worksheet.Columns.AutoFit();

                _currentRow += RowsCount;
                int endRow = _currentRow - 1;
                //Вычисление сумм по столбцам
                Worksheet.Cells[_currentRow, 2] = "Итого";
                
                if (startRow != endRow)
                {
                    Worksheet.Cells[_currentRow, 5].Formula = $"=SUM(E{startRow}:E{endRow})";   //Вычисление суммы "Выполнено за период"
                    Worksheet.Cells[_currentRow, 6].Formula = $"=SUM(F{startRow}:F{endRow})";   //Вычисление суммы "Выполнено с начала семестра"
                    Worksheet.Cells[_currentRow, 7].Formula = $"=SUM(G{startRow}:G{endRow})";   //Вычисление суммы "Остаток на конец периода"
                }
                else
                {
                    Worksheet.Cells[_currentRow, 5] = "0";   //Вычисление суммы "Выполнено за период"
                    Worksheet.Cells[_currentRow, 6] = "0";   //Вычисление суммы "Выполнено с начала семестра"
                    Worksheet.Cells[_currentRow, 7] = "0";   //Вычисление суммы "Остаток на конец периода"
                }

                HeaderRange = Worksheet.get_Range((NSExcel.Range)(Worksheet.Cells[_currentRow, 2]), (NSExcel.Range)(Worksheet.Cells[_currentRow, 7]));
                HeaderRange.Font.Bold = true;   //Установление полужирного начертания для строки заголовков

                _currentRow++;
            }
            catch (Exception ex)
            {
                throw new Exception("Экспорт в файл Excel:\n" + ex.Message);
            }
        }

        public static void SaveFile(string ExcelFilePath)
        {
            //Сохранение данных в файл
            if (ExcelFilePath != null && ExcelFilePath != "")
            {
                try
                {
                    NSExcel._Worksheet Worksheet = AppExcel.ActiveSheet;    //Установление рабочего листа
                    Worksheet.SaveAs(ExcelFilePath);    //Сохранение файла

                    DialogResult result = MessageBox.Show("Файл успешно сохранен!\nКоличество строк: " + _currentRow.ToString()
                        + "\n\nОткрыть файл?",
                        "Информация", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                        AppExcel.Visible = true;
                    else
                        AppExcel.Quit();

                    _currentRow = 0;
                }
                catch (Exception ex)
                {
                    throw new Exception("Экспорт в файл Excel: Не удалось сохранить файл!\nПроверьте путь к файлу.\n"
                        + ex.Message);
                }
            }
            else    // путь к файлу не указан
            {
                AppExcel.Visible = true;
            }
        }
    }
}
