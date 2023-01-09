using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using ExRaspViewer.Classes;


namespace ExRaspViewer
{
    public partial class MainForm : Form
    {
        private ClassSqlDB data = new ClassSqlDB();
        private ClassOleDB oleDB = new ClassOleDB();
        private Service service = new Service();
        private List<Service> servicesList = new List<Service>();

        private SqlCommandBuilder builderGroup;
        private SqlCommandBuilder builderPrep;
        private SqlDataAdapter adapterGroup;
        private SqlDataAdapter adapterPrep;
        private DataSet dsGroup;
        private DataSet dsPrep;

        private int semestr;    //Семестр
        private int tekNed = 4; //Текущая неделя
        private int state = 0;  //если 0, то данные сохранены, 1 - данные еще не сохранены
        private int _information = 0;
        private bool _listGroupClick = false;

        public MainForm()
        {
            InitializeComponent();
        }

        //Загрузка данных при загрузке формы
        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadGroup();    //После загружается LoadDataNagruzkaGrupp()
            LoadPrepod();   //После загружается LoadDataNagrPrepod()
            AddColumnPlan();
            LoadDataPlanGroup();
            LoadDataPlanPrep();
            FirstLoadData();
        }

        private void FirstLoadData()
        {
            lbDay.Text = dateTimePicker1.Value.ToString("ddd"); //День недели
            semestr = Semestr();    //Определение текущего семестра
            servicesList = service.GetListNedely(semestr);
            tekNed = Nedelya() + 3;
            labelNumNed.Text = Convert.ToString(tekNed - 3);
        }

        //Определение текущего семестра
        private int Semestr()
        {
            DateTime date = DateTime.Now;
            if (date.Month >= 1 && date.Month <= 6)
                return 2;
            else
                return 1;
        }

        //Обновить таблицу с расписанием (из файла БД программы Экспресс-расписание Колледж)
        private void открытьБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string connection = oleDB.OpenFileDB();
            if (connection == "error")
                return;
            data.DeleteDataFromDB();            //Удаление данных из рабочей БД
            data.UpdateSqlTable(connection);    //Обновление данных в таблицах рабочей БД
            data.SyncPlan();                    //Синронизация данных в таблице плана
            LoadGroup();
            LoadPrepod();
        }
        
        //Выход из приложения
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //Загрузка списка групп в ListBox
        private void LoadGroup()
        {
            listBox1.DisplayMember = "NAIM";
            listBox1.ValueMember = "IDG";
            listBox1.DataSource = SqlDB.LoadTable("ListOfGroups");
        }

        //Загрузка списка преподавателей в ListBox
        private void LoadPrepod()
        {
            listBox2.DisplayMember = "FAMIO";
            listBox2.ValueMember = "IDP";
            listBox2.DataSource = SqlDB.LoadTable("ListOfTeachers");
        }

        //Загрузка данных по нагрузке групп в DataGridView
        //Нужно добавить столбец и строку с суммами (еще не реализовано)
        private void LoadDataNagruzkaGrupp()
        {
            int id = (int)listBox1.SelectedValue;
            string workDate = dateTimePicker1.Value.ToString();
            dataGridView1.DataSource = SqlDB.GroupLoad(id, workDate);

            dataGridView1.Columns["IDG"].Visible = false;
            dataGridView1.Columns["IDP"].Visible = false;
            dataGridView1.Columns["IDD"].Visible = false;
            
            //Установление ширины столбцов
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[5].Width = 40;
            dataGridView1.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[6].Width = 50;
            dataGridView1.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[7].Width = 50;
            dataGridView1.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView1.Columns[8].Width = 50;
            dataGridView1.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //CurrentHoursGroup();

            //Отключение пользовательской сортировки
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        //Загрузка нагрузки преподавателей
        //Нужно добавить строку с суммами (не реализовано)
        private void LoadDataNagrPrepod()
        {
            int id = (int)listBox2.SelectedValue;
            string workDate = dateTimePicker1.Value.ToString();
            dataGridView2.DataSource = SqlDB.TeacherLoad(id, workDate);

            dataGridView2.Columns["IDG"].Visible = false;
            dataGridView2.Columns["IDP"].Visible = false;
            dataGridView2.Columns["IDD"].Visible = false;
            
            //Установление ширины столбцов
            dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridView2.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView2.Columns[5].Width = 40;
            dataGridView2.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView2.Columns[6].Width = 50;
            dataGridView2.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView2.Columns[7].Width = 50;
            dataGridView2.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dataGridView2.Columns[8].Width = 50;
            dataGridView2.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //CurrentHoursPrepod();

            //Отключение пользовательской сортировки
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        //Обновление таблицы с нагрузкой групп
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (state == 1)
                UpdateTablePlanGroup();
            LoadDataNagruzkaGrupp();
            LoadDataPlanGroup();
        }

        //Обновление таблицы с нагрузкой преподавателей
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (state == 1)
                UpdateTablePlanPrepod();
            LoadDataNagrPrepod();
            LoadDataPlanPrep();
        }

        
        //Отображние данных по пройденным урокам и остатку (для преподавателей)
        private void CurrentHoursPrepod()
        {
            int idg, idp, idd, num, vsego;
            string dat = dateTimePicker1.Value.ToString("d");
            int countDGV = dataGridView2.Rows.Count;
            for (int i = 0; i < countDGV; i++)
            {
                idp = (int)dataGridView2[0, i].Value;
                idg = (int)dataGridView2[1, i].Value;
                idd = (int)dataGridView2[2, i].Value;
                num = data.CountUroki(idp, idg, idd, dat);
                vsego = Convert.ToInt32(dataGridView2[6, i].Value);
                dataGridView2[7, i].Value = num;            //Выполнено
                dataGridView2[8, i].Value = vsego - num;    //Остаток
            }
        }

        //Изменение даты
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            lbDay.Text = dateTimePicker1.Value.ToString("ddd");
            tekNed = Nedelya() + 3;
            labelNumNed.Text = Convert.ToString(tekNed - 3);

            LoadDataNagruzkaGrupp();
            SummRowPlanGroup();

            LoadDataNagrPrepod();
            SummRowPlanPrep();
            //CurrentHoursGroup();
            //CurrentHoursPrepod();
        }

        //Вычисление номера недели
        private int Nedelya()
        {
            DateTime date = dateTimePicker1.Value.Date;
            int result = 0;
            foreach (var item in servicesList)
            {
                if (date >= item.DateNachNed && date <= item.DateKonNed)
                {
                    result = item.NumNed;
                    break;
                }
                else
                    result = 1;
            }
            if (result < 0)
                result = 1;
            return result;
        }

        //Добавление колонок с наименованиями недель
        private void AddColumnPlan()
        {
            List<string> list;
            list = data.LoadListNedely();
            string ned;
            for (int i = 0; i < list.Count; i++)
            {
                string header = list[i];
                if (i < 4)
                {
                    ned = list[i];
                    dataGridView3.Columns.Add(ned, header);
                    dataGridView3.Columns[i].DataPropertyName = header;
                    dataGridView4.Columns.Add(ned, header);
                    dataGridView4.Columns[i].DataPropertyName = header;
                }
                else
                {
                    ned = "Ned_" + (i - 3);
                    dataGridView3.Columns.Add(ned, header);
                    dataGridView3.Columns[i].DataPropertyName = ned;
                    dataGridView4.Columns.Add(ned, header);
                    dataGridView4.Columns[i].DataPropertyName = ned;
                }
            }
        }

        //Загрузка данных по плану в DGV (для групп)
        private void LoadDataPlanGroup()
        {
            adapterGroup = new SqlDataAdapter();
            builderGroup = new SqlCommandBuilder();
            dsGroup = new DataSet();
            int idg = Convert.ToInt32(listBox1.SelectedValue);
            dataGridView3.AutoGenerateColumns = false;
            SqlCommand cmd = new SqlCommand("SELECT * FROM [PLAN] WHERE IDG = @IDG ORDER BY [IDP], [IDD]", data.ConnSql);
            cmd.CommandType = CommandType.Text;
            cmd.Parameters.Add("@IDG", SqlDbType.Int);
            cmd.Parameters["@IDG"].Value = idg;
            adapterGroup.SelectCommand = cmd;   //Выбор даных из таблицы плана (с фильтром)
            builderGroup.DataAdapter = adapterGroup;
            adapterGroup.Fill(dsGroup);
            dataGridView3.DataSource = dsGroup.Tables[0];

            //Отключение пользовательской сортировки
            for (int i = 0; i < dataGridView3.ColumnCount; i++)
            {
                dataGridView3.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }

            //Скрытие столбцов
            if (dataGridView3.ColumnCount > 0)
            {
                dataGridView3.Columns["IDN"].Visible = false;
                dataGridView3.Columns["IDG"].Visible = false;
                dataGridView3.Columns["IDP"].Visible = false;
                dataGridView3.Columns["IDD"].Visible = false;
            }
            
            SummRowPlanGroup(); //Сумма по строкам
            SummColumnPlanLoadGroup();
            ColumnPlanWidthGroup(); //Установка ширины столбцов

            //Автовысота таблицы
            if (_listGroupClick)
            {
                _information = 20 * dataGridView3.RowCount + 50;
                splitContainer1.SplitterDistance = _information;
                _listGroupClick = false;
            }
        }

        //Загрузка данных по плану в DGV (для преподавателей)
        private void LoadDataPlanPrep()
        {
            adapterPrep = new SqlDataAdapter();
            builderPrep = new SqlCommandBuilder();
            dsPrep = new DataSet();
            int idp = Convert.ToInt32(listBox2.SelectedValue);
            dataGridView4.AutoGenerateColumns = false;
            SqlCommand cmd = new SqlCommand("SELECT * FROM [PLAN] WHERE IDP = @IDP ORDER BY [IDG], [IDD]", data.ConnSql);
            cmd.CommandType = CommandType.Text;
            cmd.Parameters.Add("@IDP", SqlDbType.Int);
            cmd.Parameters["@IDP"].Value = idp;
            adapterPrep.SelectCommand = cmd;    //Выбор даных из таблицы плана (с фильтром)
            builderPrep.DataAdapter = adapterPrep;
            adapterPrep.Fill(dsPrep);
            dataGridView4.DataSource = dsPrep.Tables[0];
            
            //Отключение пользовательской сортировки
            for (int i = 0; i < dataGridView4.ColumnCount; i++)
            {
                dataGridView4.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }

            //Скрытие столбцов
            if (dataGridView4.ColumnCount > 0)
            {
                dataGridView4.Columns["IDN"].Visible = false;
                dataGridView4.Columns["IDG"].Visible = false;
                dataGridView4.Columns["IDP"].Visible = false;
                dataGridView4.Columns["IDD"].Visible = false;
            }
            SummRowPlanPrep();  //Сумма по строкам
            SummColumnPlanLoadPrep();
            ColumnPlanWidthPrep();  //Установка ширины столбцов
        }

        //Задание ширины столбцов в таблице плана у групп
        private void ColumnPlanWidthGroup()
        {
            int colCount = dataGridView3.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                dataGridView3.Columns[i].Width = 40;
            }
        }

        //Задание ширины столбцов в таблице плана у преподавателей
        private void ColumnPlanWidthPrep()
        {
            int colCount = dataGridView4.Columns.Count;
            for (int i = 0; i < colCount; i++)
            {
                dataGridView4.Columns[i].Width = 40;
            }
        }

        //Сохранение таблицы плана в БД (для групп)
        private void UpdateTablePlanGroup()
        {
            builderGroup.GetUpdateCommand();
            adapterGroup.Update(dsGroup.Tables[0]);
        }

        //Сохранение таблицы плана в БД (для преподавателей)
        private void UpdateTablePlanPrepod()
        {
            builderPrep.GetUpdateCommand();
            adapterPrep.Update(dsPrep.Tables[0]);
        }

        //Обновление данных в БД (для групп)
        private void dataGridView3_Leave(object sender, EventArgs e)
        {
            dataGridView3.EndEdit();
            UpdateTablePlanGroup();
            LoadDataPlanPrep();
            lbStatus.Text = "Данные сохранены";
            lbStatus.ForeColor = Color.Green;
            state = 0;

        }

        //Обновление данных в БД (для преподавателей)
        private void dataGridView4_Leave(object sender, EventArgs e)
        {
            dataGridView4.EndEdit();
            UpdateTablePlanPrepod();
            LoadDataPlanGroup();
            lbStatus.Text = "Данные сохранены";
            lbStatus.ForeColor = Color.Green;
            state = 0;
        }

        //Смена строки в нагрузке групп
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            int rowIndex = dataGridView1.CurrentCell.RowIndex;
            if (dataGridView1.Focused && rowIndex != -1 && rowIndex != dataGridView3.RowCount)
                dataGridView3.Rows[rowIndex].Cells[tekNed].Selected = true;
        }

        //Смена строки в нагрузке преподавателей
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            int rowIndex = dataGridView2.CurrentCell.RowIndex;
            if (dataGridView2.Focused && rowIndex != -1 && rowIndex != dataGridView4.RowCount)
                dataGridView4.Rows[rowIndex].Cells[tekNed].Selected = true;
        }

        //Смена строки в плане групп
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            int rowIndex = dataGridView3.CurrentCell.RowIndex;
            if (dataGridView3.Focused && rowIndex != -1 && rowIndex != dataGridView3.RowCount - 1)
                if (dataGridView1.Rows.Count !=0)
                    dataGridView1.Rows[rowIndex].Selected = true;
        }

        //Смена строки в плане преподавателей
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            int rowIndex = dataGridView4.CurrentCell.RowIndex;
            if (dataGridView4.Focused && rowIndex != -1 && rowIndex != dataGridView4.RowCount - 1)
                if (dataGridView2.Rows.Count != 0)
                    dataGridView2.Rows[rowIndex].Selected = true;
        }

        //Подсчет элементов в столбце (план групп)
        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView3.BindingContext[dataGridView3.DataSource].EndCurrentEdit();    //Завершение редактирования текущей ячейки
            lbStatus.Text = "Ячейка не редактируется";
            lbStatus.ForeColor = Color.Black;
            SummPlanGroupColumn();
        }

        //Подсчет элементов в столбце (план преподавателей)
        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView4.BindingContext[dataGridView4.DataSource].EndCurrentEdit();    //Завершение редактирования текущей ячейки
            lbStatus.Text = "Ячейка не редактируется";
            lbStatus.ForeColor = Color.Black;
            SummPlanPrepColumn();
        }

        //Суммирование значений в строке (нагрузка у групп)
        private void SummRowPlanGroup()
        {
            int rowCount = dataGridView3.Rows.Count - 1;
            int columnCount = dataGridView3.Columns.Count;
            int summUr = 0;
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 4; j < columnCount; j++)
                {
                    object item = dataGridView3.Rows[i].Cells[j].Value;
                    if (item != null && item != DBNull.Value)
                    {
                        string a = item.ToString();
                        if (service.StringIsDigit(a))
                        {
                            summUr += Convert.ToInt32(a);
                        }
                    }
                }
                int ostUr = (int)dataGridView1[8, i].Value;
                dataGridView1[9, i].Value = ostUr - summUr;
                summUr = 0;
            }
        }

        //Суммирование значений в строке (нагрузка у преподавателей)
        private void SummRowPlanPrep()
        {
            int rowsCount = dataGridView4.Rows.Count - 1;
            int columnCount = dataGridView4.Columns.Count;
            int summUr = 0;
            for (int i = 0; i < rowsCount; i++)
            {
                for (int j = 4; j < columnCount; j++)
                {
                    object item = dataGridView4.Rows[i].Cells[j].Value;
                    if (item != null && item != DBNull.Value)
                    {
                        string a = item.ToString();
                        if (service.StringIsDigit(a))
                        {
                            summUr += Convert.ToInt32(a);
                        }
                    }
                }
                int ostUr = (int)dataGridView2[8, i].Value;
                dataGridView2[9, i].Value = ostUr - summUr;
                summUr = 0;
            }
        }

        //Сумма по столбцам в плане (для групп при загрузке плана)
        private void SummColumnPlanLoadGroup()
        {
            int rowsCount = dataGridView3.Rows.Count - 1;    //Количество строк (без последней строки)
            int columnCount = dataGridView3.Columns.Count;
            int summUr = 0;
            for (int i = 4; i < columnCount; i++)
            {
                for (int j = 0; j < rowsCount; j++)
                {
                    object item = dataGridView3.Rows[j].Cells[i].Value;
                    int podGruppa = Convert.ToInt32(dataGridView1.Rows[j].Cells[5].Value);
                    if (item != DBNull.Value && item != null && podGruppa != 2)
                    {
                        string a = item.ToString();
                        if (service.StringIsDigit(a))
                        { summUr += Convert.ToInt32(item); }
                        else
                        {
                            SetColorCellPlan(a, 3, j, i);   //Форматирование ячеек
                        }
                    }
                }
                dataGridView3.Rows[rowsCount].Cells[i].Value = summUr; //Отображение суммы столбца
                SetColorRowSumm(3, summUr, rowsCount, i);
                summUr = 0;
            }
        }

        //Форматирование ячеек в соответствии с правилами
        private void SetColorCellPlan(string cell, int dgv, int rowIndex, int columnIndex)
        {
            switch (dgv)
            {
                case 3:  
                    if (cell == "п" || cell == "пп")
                    {
                        dataGridView3.Rows[rowIndex].Cells[columnIndex].Style.BackColor = Color.DarkOrange;
                        dataGridView3.Rows[rowIndex].Cells[columnIndex].Style.SelectionBackColor = Color.Orange;
                    }
                    else if (cell == "у")
                    {
                        dataGridView3.Rows[rowIndex].Cells[columnIndex].Style.BackColor = Color.LimeGreen;
                        dataGridView3.Rows[rowIndex].Cells[columnIndex].Style.SelectionBackColor = Color.Lime;
                    }
                    break;

                case 4:
                    if (cell == "п" || cell == "пп")
                    {
                        dataGridView4.Rows[rowIndex].Cells[columnIndex].Style.BackColor = Color.DarkOrange;
                        dataGridView4.Rows[rowIndex].Cells[columnIndex].Style.SelectionBackColor = Color.Orange;
                    }
                    else if (cell == "у")
                    {
                        dataGridView4.Rows[rowIndex].Cells[columnIndex].Style.BackColor = Color.LimeGreen;
                        dataGridView4.Rows[rowIndex].Cells[columnIndex].Style.SelectionBackColor = Color.Lime;
                    }
                    break;
                default:
                    break;
            }
        }

        //Сумма по столбцам в плане (для преподавателей при загрузке плана)
        private void SummColumnPlanLoadPrep()
        {
            int rowsCount = dataGridView4.Rows.Count - 1;    //Количество строк (без последней строки)
            int columnCount = dataGridView4.Columns.Count;
            int summUr = 0;
            for (int i = 4; i < columnCount; i++)
            {
                for (int j = 0; j < rowsCount; j++)
                {
                    object item = dataGridView4.Rows[j].Cells[i].Value;
                    int podGruppa = Convert.ToInt32(dataGridView2.Rows[j].Cells[5].Value);
                    if (item != DBNull.Value && item != null)
                    {
                        string a = item.ToString();
                        if (service.StringIsDigit(a))
                        { summUr += Convert.ToInt32(item); }
                        else
                        {
                            SetColorCellPlan(a, 4, j, i);   //Форматирование ячеек
                        }
                    }
                }
                dataGridView4.Rows[rowsCount].Cells[i].Value = summUr; //Отображение суммы столбца
                SetColorRowSumm(4, summUr, rowsCount, i);
                summUr = 0;
            }
        }

        //Форматирование ячеек при превышении нагрузки
        private void SetColorRowSumm(int dgv, int summUr, int rowsIndex, int columnIndex)
        {
            if (summUr <= 36 && dgv == 3)
            {
                dataGridView3.Rows[rowsIndex].Cells[columnIndex].Style.BackColor = Color.White;
                dataGridView3.Rows[rowsIndex].Cells[columnIndex].Style.ForeColor = Color.Black;
            }
            else if (summUr <= 36 && dgv == 4)
            {
                dataGridView4.Rows[rowsIndex].Cells[columnIndex].Style.BackColor = Color.White;
                dataGridView4.Rows[rowsIndex].Cells[columnIndex].Style.ForeColor = Color.Black;
            }

            if (summUr > 36 && dgv == 3)
            {
                dataGridView3.Rows[rowsIndex].Cells[columnIndex].Style.BackColor = Color.Red;
                dataGridView3.Rows[rowsIndex].Cells[columnIndex].Style.ForeColor = Color.White;
            }
            else if (summUr > 36 && dgv == 4)
            {
                dataGridView4.Rows[rowsIndex].Cells[columnIndex].Style.BackColor = Color.Red;
                dataGridView4.Rows[rowsIndex].Cells[columnIndex].Style.ForeColor = Color.White;
            }
        }

        //Подсчет количества уроков в плане группы (по одному столбцу)
        private void SummPlanGroupColumn()
        {
            int columnIndex = dataGridView3.CurrentCell.ColumnIndex;  //Индекс текущего столбца
            int rowsCount = dataGridView3.Rows.Count - 1;    //Количество строк (без последней строки)
            int summUr = 0;
            for (int i = 0; i < rowsCount; i++)
            {
                object yacheyka = dataGridView3.Rows[i].Cells[columnIndex].Value;
                int podGruppa = Convert.ToInt32(dataGridView1.Rows[i].Cells[5].Value);
                if (yacheyka != DBNull.Value && yacheyka != null && podGruppa != 2)
                {
                    string a = yacheyka.ToString();
                    if (service.StringIsDigit(a))
                    { summUr += Convert.ToInt32(yacheyka); }
                    else
                    {
                        SetColorCellPlan(a, 3, i, columnIndex);   //Форматирование ячеек
                    }
                }
            }
            dataGridView3.Rows[rowsCount].Cells[columnIndex].Value = summUr; //Отображение суммы столбца
            SetColorRowSumm(3, summUr, rowsCount, columnIndex);
            SummRowPlanGroup();
        }

        //Подсчет количества уроков в плане преподавателей (по столбцам)
        private void SummPlanPrepColumn()
        {
            int columnIndex = dataGridView4.CurrentCell.ColumnIndex;  //Индекс текущего столбца
            int rowsCount = dataGridView4.Rows.Count - 1;    //Количество строк (без последней строки)
            int summUr = 0;
            for (int i = 0; i < rowsCount; i++)
            {
                object yacheyka = dataGridView4.Rows[i].Cells[columnIndex].Value;
                int podGruppa = Convert.ToInt32(dataGridView2.Rows[i].Cells[5].Value);
                if (yacheyka != DBNull.Value)
                {
                    string a = yacheyka.ToString();
                    if (service.StringIsDigit(a))
                    {
                        summUr += Convert.ToInt32(yacheyka);
                    }
                    else
                    {
                        SetColorCellPlan(a, 4, i, columnIndex);   //Форматирование ячеек
                    }
                }
            }
            dataGridView4.Rows[rowsCount].Cells[columnIndex].Value = summUr; //Отображение суммы столбца
            SetColorRowSumm(4, summUr, rowsCount, columnIndex);
            SummRowPlanPrep();
        }

        //Удаление значения текущей ячейки
        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                dataGridView3.CurrentCell.Value = null;
                dataGridView3.BindingContext[dataGridView3.DataSource].EndCurrentEdit();
                SummPlanGroupColumn();
                state = 1;
            }
        }

        //Удаление значения текущей ячейки
        private void dataGridView4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                dataGridView4.CurrentCell.Value = null;
                dataGridView4.BindingContext[dataGridView4.DataSource].EndCurrentEdit();
                SummPlanPrepColumn();
                state = 1;
            }
        }

        private void dataGridView3_Enter(object sender, EventArgs e)
        {
            lbStatus.Text = "Данные изменены";
            lbStatus.ForeColor = Color.Red;
            state = 1;
        }

        private void dataGridView4_Enter(object sender, EventArgs e)
        {
            lbStatus.Text = "Данные изменены";
            lbStatus.ForeColor = Color.Red;
            state = 1;
        }

        private void dataGridView3_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            lbStatus.Text = "Ячейка редактируется";
            lbStatus.ForeColor = Color.DarkGoldenrod;
        }

        private void dataGridView4_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            lbStatus.Text = "Ячейка редактируется";
            lbStatus.ForeColor = Color.DarkGoldenrod;
        }

        //Заполнение таблицы с наименованием недель (1 семестр)
        private void заполнитьНеделиДля1СемToolStripMenuItem_Click(object sender, EventArgs e)
        {
            data.DeleteDataFromDB("NED");
            data.LoadNed(1);
            dataGridView3.Columns.Clear();
            dataGridView4.Columns.Clear();
            AddColumnPlan();
            LoadDataPlanGroup();
            LoadDataPlanPrep();
            semestr = 1;
        }

        //Заполнение таблицы с наименованием недель (2 семестр)
        private void заполнитьНеделиДля2СемToolStripMenuItem_Click(object sender, EventArgs e)
        {
            data.DeleteDataFromDB("NED");
            data.LoadNed(2);
            dataGridView3.Columns.Clear();
            dataGridView4.Columns.Clear();
            AddColumnPlan();
            LoadDataPlanGroup();
            LoadDataPlanPrep();
            semestr = 2;
        }

        //Кнопки для выделения цветом ячеек
        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView3.CurrentCell.Style.BackColor = Color.PaleGreen;
            dataGridView3.Focus();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView3.CurrentCell.Style.BackColor = Color.White;
            dataGridView3.Focus();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView4.CurrentCell.Style.BackColor = Color.PaleGreen;
            dataGridView4.Focus();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView4.CurrentCell.Style.BackColor = Color.White;
            dataGridView4.Focus();
        }

        private void отчётЗаМесяцToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Otchet form = new Otchet();
            form.ShowDialog();
        }

        private void оПрограммеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            About form = new About();
            form.ShowDialog();
        }

        private void listBox2_Click(object sender, EventArgs e)
        {
            //Автовысота таблицы
            splitContainer1.SplitterDistance = 80;
            dataGridView4.Focus();
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            _listGroupClick = true;
            dataGridView3.Focus();
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            dataGridView3.Focus();
            if (_information == 0)
            {
                _information = 20 * dataGridView3.RowCount + 50;
            }
            splitContainer1.SplitterDistance = _information;
        }

        private void dataGridView2_Click(object sender, EventArgs e)
        {
            dataGridView4.Focus();
            splitContainer1.SplitterDistance = 80;
        }

        private void dataGridView3_Click(object sender, EventArgs e)
        {
            if (_information == 0)
            {
                _information = 20 * dataGridView3.RowCount + 50;
            }
            splitContainer1.SplitterDistance = _information;
        }

        private void dataGridView4_Click(object sender, EventArgs e)
        {
            splitContainer1.SplitterDistance = 80;
        }
    }
}
