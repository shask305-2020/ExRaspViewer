using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ExRaspViewer.Classes
{
    internal class ClassOleDB
    {
        public ClassOleDB()
        {

        }

        //Открытие файла БД расписания
        public string OpenFileDB()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "Файл БД MS Access (*.mdb)|*.mdb";
            string filePath = "";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
                return filePath;
            }
            else
            {
                return "error";
            }
            
        }
        
        //Загрузка обновлённых таблиц из БД расписания
        public DataTable UpdateSqlDBTable(string tableN, string connection)
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + connection);
            using (OleDbCommand oleDbCommand = new OleDbCommand("SELECT * FROM " + tableN + "", con))
            {
                oleDbCommand.CommandType = CommandType.Text;
                OleDbDataAdapter oleDbAdapter = new OleDbDataAdapter(oleDbCommand);
                DataTable dataTable = new DataTable();
                try
                {
                    oleDbAdapter.Fill(dataTable);
                    return dataTable;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                return dataTable;
            }
        }
    }
}
