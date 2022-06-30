using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using ExRaspViewer;

namespace ExRaspViewer.Classes
{
    internal class ClassLoadData
    {
        public string filePath = String.Empty;

        private OleDbConnection connOleDB = null;
        
        //Выбрать файл базы данных программы Экспресс-расписание
        public void OpenFileDB()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Файл БД MS Access (*.mdb)|*.mdb";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Получить путь к указанному файлу
                    filePath = openFileDialog.FileName;
                    connOleDB = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath);
                    try
                    {
                        connOleDB.Open();
                        if (connOleDB.State == ConnectionState.Open)
                        {
                            MessageBox.Show("Соединение открыто");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        connOleDB.Close();
                    }
                }
            }
        }

        public DataSet UpdateSqlDB(string tableName)
        {
            connOleDB = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath);
            using (OleDbCommand oleDbCommand = new OleDbCommand("SELECT * FROM [" + tableName + "]", connOleDB))
            {
                oleDbCommand.CommandType = CommandType.Text;
                OleDbDataAdapter oleDbAdapter = new OleDbDataAdapter(oleDbCommand);
                DataSet dataSet = new DataSet();
                oleDbAdapter.Fill(dataSet, tableName);
                return dataSet;
            }    
        }
    }
}
