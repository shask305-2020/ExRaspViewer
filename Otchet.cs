using ExRaspViewer.Classes;
using System;
using System.ComponentModel;
using System.Threading;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace ExRaspViewer
{
    public partial class Otchet : Form
    {
        private ClassSqlDB data = new ClassSqlDB();
        private SaveFileDialog saveFileDialog = new SaveFileDialog();
        private List<Pedagog> list;  //Создание списка преподавателей

        public Otchet()
        {
            InitializeComponent();
            radioButton2.Checked = true;
        }

        private void Otchet_Load(object sender, EventArgs e)
        {
            LoadMonth();    //Загрузка списка месяцев
            LoadPrepod();   //Загруза списка преподавателей
            LoadDataNagrPrepod();   //Загрузка нагрузки преподавателей
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
            DataTable dt = data.LoadPrepodTable();
            listBox1.DataSource = dt;

            //Создание списка преподавателей
            list = new List<Pedagog>(dt.Rows.Count);
            foreach (DataRow item in dt.Rows)
            {
                var cells = item.ItemArray;
                list.Add(new Pedagog()
                {
                    Id_Pedagog = (int)cells[0],
                    Fio_Pedagog = cells[1].ToString()
                });
            }
        }

        //Загрузка нагрузки преподавателей
        private void LoadDataNagrPrepod()
        {
            int id = 0;
            if (listBox1.Items.Count > 0)
                id = (int)listBox1.SelectedValue;

            string dat_N = date1.Value.ToString("d");
            string dat_K = date2.Value.ToString("d");
            dataGridView1.DataSource = data.LoadNagrPrepodOtchet(id, dat_N, dat_K);

            //Скрытие столбцов с id
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
            
            //Отключение пользовательской сортировки
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        //Смена преподавателя
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
                LoadDataPedagog();
        }

        //Экспорт данных по всем преподавателям
        private void LoadDataPedagog()
        {
            string dat_N = date1.Value.ToString("d");
            string dat_K = date2.Value.ToString("d");
            int idp;
            string fio_pedagog;
            ExportToExcel.datN = dat_N;
            ExportToExcel.datK = dat_K;
            DataTable dt = new DataTable();
            foreach (var item in list)
            {
                idp = item.Id_Pedagog;
                fio_pedagog = item.Fio_Pedagog;
                dt = data.LoadNagrPrepodOtchet(idp, dat_N, dat_K);
                ExportToExcel.Export(dt, fio_pedagog);
            }
            ExportToExcel.SaveFile(saveFileDialog.FileName);
        }

        //Обновление данных
        private void date1_ValueChanged(object sender, EventArgs e)
        {
            LoadDataNagrPrepod();
        }
        private void date2_ValueChanged(object sender, EventArgs e)
        {
            LoadDataNagrPrepod();
        }
    }
}
