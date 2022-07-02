using System;
using System.Data;
using System.Windows.Forms;
using NSExcel = Microsoft.Office.Interop.Excel;

namespace ExRaspViewer.Classes
{
    public static class Excel
    {
        public static NSExcel.Application AppExcel = new NSExcel.Application();

        public static void ExportToExcel(this DataTable DataTable, string ExcelFilePath = null)
        {
            try
            {
                int ColumnsCount = DataTable.Columns.Count;   //Количество столбцов
                if (DataTable == null || ColumnsCount == 0)
                    throw new Exception("Экспорт в файл Excel: несуществующая или пустая входная таблица!\n");

                AppExcel.Workbooks.Add();   //Добавление рабочей книги
                NSExcel._Worksheet Worksheet = AppExcel.ActiveSheet;    //Установление рабочего листа
                object[] Header = new object[ColumnsCount];     //Массив для названий колонок
                for (int i = 0; i < ColumnsCount; i++)
                    Header[i] = DataTable.Columns[i].ColumnName;    //Заполнение массива названиями колонок
                
                //Выяснение, сколько места нужно под заголовки (сколько ячеек)
                NSExcel.Range HeaderRange = Worksheet.get_Range((NSExcel.Range)(Worksheet.Cells[1, 1]), (NSExcel.Range)(Worksheet.Cells[1, ColumnsCount]));
                HeaderRange.Value = Header; //Заполнние первой строки заголовками столбцов
                
                HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);  //Перекраска строки с заголовками
                HeaderRange.Font.Bold = true;   //Установление полужирного начертания для строки заголовков
                int RowsCount = DataTable.Rows.Count;   //Получеие количства строк
                
                //Массив для хранения данных
                object[,] Cells = new object[RowsCount, ColumnsCount];

                //Заполнение двумерного массива данными из ячеек DataTable
                for (int j = 0; j < RowsCount; j++)
                    for (int i = 0; i < ColumnsCount; i++)
                        Cells[j, i] = DataTable.Rows[j][i];

                //Выяснение, сколько места нужно под строки с данными
                Worksheet.get_Range((NSExcel.Range)(Worksheet.Cells[2, 1]), (NSExcel.Range)(Worksheet.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;
                
                //Сохранение данных в файл
                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {
                        Worksheet.SaveAs(ExcelFilePath);
                        AppExcel.Quit();
                        MessageBox.Show("Файл сохранен!");
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
            catch (Exception ex)
            {
                throw new Exception("Экспорт в файл Excel: \n" + ex.Message);
            }
        }

        public static DataTable DataGridView_To_Datatable(DataGridView dg)
        {
            DataTable ExportDataTable = new DataTable();
            foreach (DataGridViewColumn col in dg.Columns)
            {
                ExportDataTable.Columns.Add(col.Name);
            }
            foreach (DataGridViewRow row in dg.Rows)
            {
                DataRow dRow = ExportDataTable.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                ExportDataTable.Rows.Add(dRow);
            }
            return ExportDataTable;
        }
    }
}
