using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ExRaspViewer
{
    public partial class MainForm : Form
    {
        private OleDbConnection connOleDB = null;
        private string sqlConnection;
        public MainForm()
        {
            InitializeComponent();
        }

        private void OpenFileDB()
        {
            var filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "MS Access database file (*.mdb)|*.mdb";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Получить путь к указанному файлу
                    filePath = openFileDialog.FileName;
                    this.Text = "Планировщик расписания    ExRaspisView v0.1    Файл: " + filePath;
                    connOleDB = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath);
                }

            }
        }

        private void LoadDB()
        {
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

        private void открытьБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDB();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
