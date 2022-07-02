using ExRaspViewer.Classes;
using System;
using System.Data;
using System.Windows.Forms;
using NSExcel = Microsoft.Office.Interop.Excel;

namespace ExRaspViewer
{
    public partial class Otchet : Form
    {
        private ClassSqlDB data = new ClassSqlDB();
        private Timer timer = new Timer();
        private int iteration = 0;
        private SaveFileDialog saveFileDialog = new SaveFileDialog();

        public Otchet()
        {
            InitializeComponent();
            radioButton2.Checked = true;
            timer.Tick += Timer_Tick;
            timer.Interval = 700;
        }

        private void Otchet_Load(object sender, EventArgs e)
        {
            LoadMonth();
            LoadPrepod();
            LoadDataNagrPrepod();
        }

        //Загрузка месяцев
        private void LoadMonth()
        {
            cbMonth.DisplayMember = "NAIM";
            cbMonth.ValueMember = "MN";
            cbMonth.DataSource = data.LoadMonth();
        }
        //Загрузка списка преподавателей в ListBox
        private void LoadPrepod()
        {
            listBox1.DisplayMember = "FAMIO";
            listBox1.ValueMember = "IDP";
            listBox1.DataSource = data.LoadPrepodTable();
        }

        //Загрузка нагрузки преподавателей
        private void LoadDataNagrPrepod()
        {
            int id = (int)listBox1.SelectedValue;
            dataGridView1.DataSource = data.LoadNagrPrepodOtchet(id);

            dataGridView1.Columns["IDG"].Visible = false;
            dataGridView1.Columns["IDP"].Visible = false;
            dataGridView1.Columns["IDD"].Visible = false;

            //Установление ширины столбцов
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //Группа
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //Дисциплина
            dataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; //подгруппа
            dataGridView1.Columns[5].Width = 40;
            dataGridView1.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; //Всего часов (по нагрузке)
            dataGridView1.Columns[6].Width = 50;
            dataGridView1.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; //Выполнено за период
            dataGridView1.Columns[7].Width = 50;
            dataGridView1.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.None; //Выполнено с начала семестра
            dataGridView1.Columns[8].Width = 60;
            dataGridView1.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; //Остаток на конец периода
            CurrentHoursPrepod();

            //Отключение пользовательской сортировки
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        //Отображние данных по пройденным урокам и остатку (для преподавателей)
        private void CurrentHoursPrepod()
        {
            int idg, idp, idd, num, vsego;
            string dat_N = date1.Value.ToString("d");
            string dat_K = date2.Value.ToString("d");
            int countDGV = dataGridView1.Rows.Count;
            if (dataGridView1.Rows.Count > 0)
            {
                for (int i = 0; i < countDGV; i++)
                {
                    idp = (int)dataGridView1[0, i].Value;
                    idg = (int)dataGridView1[1, i].Value;
                    idd = (int)dataGridView1[2, i].Value;
                    num = data.CountUroki(idp, idg, idd, dat_N, dat_K);    //Выполнено за период
                    vsego = Convert.ToInt32(dataGridView1[6, i].Value);     //Всего часов
                    dataGridView1[7, i].Value = num;            //Выполнено за период
                    int nachGoda = data.CountUroki(idp, idg, idd, dat_K);
                    dataGridView1[8, i].Value = nachGoda; //Выполнено с начала года
                    dataGridView1[9, i].Value = vsego - nachGoda;    //Остаток на конец периода
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CurrentHoursPrepod();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            txTeacher.Text = listBox1.Text;
            LoadDataNagrPrepod();
        }

        //Работа с интервалами
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            date1.Enabled = radioButton2.Checked;
            date2.Enabled = radioButton2.Checked;
            cbMonth.Enabled = radioButton1.Checked;
            if (radioButton1.Checked)
            {
                int m1 = Convert.ToInt32(cbMonth.SelectedValue);
                int m2 = m1 + 1;
                DateTime d1 = new DateTime(DateTime.Now.Year, m1, 1);
                DateTime d2 = new DateTime(DateTime.Now.Year, m2, 1);
                date1.Value = d1;
                date2.Value = d2;
            }
            if (radioButton2.Checked)
            {
                int m1 = DateTime.Now.Month - 1;
                int m2 = m1 + 1;
                DateTime d1 = new DateTime(DateTime.Now.Year, m1, 26);
                DateTime d2 = new DateTime(DateTime.Now.Year, m2, 26);
                date1.Value = d1;
                date2.Value = d2;
            }
        }

        //Смена интервала
        private void cbMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                int m1 = Convert.ToInt32(cbMonth.SelectedValue);
                int m2 = m1 + 1;
                DateTime d1 = new DateTime(DateTime.Now.Year, m1, 1);
                DateTime d2 = new DateTime(DateTime.Now.Year, m2, 1);
                date1.Value = d1;
                date2.Value = d2;
            }
        }

        //Экспорт в Excel
        private void button4_Click(object sender, EventArgs e)
        {            
            saveFileDialog.InitialDirectory = "c:\\";
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|Excel 2003 (*.xls)|*.xls";
            saveFileDialog.FilterIndex = 1;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            listBox1.SelectedIndex = iteration;
            string pedagog = txTeacher.Text;
            DataTable dt = ExportToExcel.DataGridView_To_Datatable(dataGridView1);
            dt.Export(pedagog);
            iteration++;
            if (iteration >= listBox1.Items.Count)
            {
                timer.Stop();
                iteration = 0;
                ExportToExcel.SaveFile(saveFileDialog.FileName);
            }
        }
    }
}
