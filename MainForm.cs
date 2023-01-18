using ExRaspViewer.Classes;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;


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

        private int _semestr;    //Семестр
        private int tekNed = 4; //Текущая неделя
        private bool is_edit = false;   //если false, то данные сохранены, true - данные еще не сохранены
        private int _tableHeight = 0;
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
            lbDay.Text = dateTimePicker1.Value.ToString("ddd"); //День недели
            _semestr = Service.Semestr();    //Определение текущего семестра
            servicesList = service.GetListNedely(_semestr);
            tekNed = Nedelya() + 3;     //+3 стоит потому, что первые три столбца для идентификаторов групп, преподавателей и предметов
            labelNumNed.Text = Convert.ToString(tekNed - 3);
        }


        #region Загрузка списков групп и преподавателей
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
        #endregion


        #region Загрузка данных по нагрузке групп и преподавателей
        //Загрузка данных по нагрузке групп в DataGridView
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

            //Отключение пользовательской сортировки
            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }

        //Загрузка нагрузки преподавателей
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

            //Отключение пользовательской сортировки
            for (int i = 0; i < dataGridView2.ColumnCount; i++)
            {
                dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.Programmatic;
            }
        }
        #endregion


        #region Обновление таблиц с нагрузкой групп и преподавателей
        //Обновление таблицы с нагрузкой групп
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (is_edit)
                UpdateTablePlanGroup();
            LoadDataNagruzkaGrupp();
            LoadDataPlanGroup();
        }

        //Обновление таблицы с нагрузкой преподавателей
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (is_edit)
                UpdateTablePlanPrepod();
            LoadDataNagrPrepod();
            LoadDataPlanPrep();
        }
        #endregion


        #region Работа с планом

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
                _tableHeight = 20 * dataGridView3.RowCount + 50;
                splitContainer1.SplitterDistance = _tableHeight;
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
                dataGridView3.Columns[i].Width = 40;
        }

        //Задание ширины столбцов в таблице плана у преподавателей
        private void ColumnPlanWidthPrep()
        {
            int colCount = dataGridView4.Columns.Count;
            for (int i = 0; i < colCount; i++)
                dataGridView4.Columns[i].Width = 40;
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

            is_edit = false;
        }

        //Обновление данных в БД (для преподавателей)
        private void dataGridView4_Leave(object sender, EventArgs e)
        {
            dataGridView4.EndEdit();
            UpdateTablePlanPrepod();
            LoadDataPlanGroup();

            is_edit = false;
        }
        #endregion


        #region Сервисные функции
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
        }

        //Вычисление номера недели (нужно перебросить в класс Service)
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

        //Заполнение таблицы с наименованиями недель для плана (1 или 2 семестр)
        private void SetWeek(int semestr)
        {
            data.DeleteDataFromDB("NED");
            data.LoadNed(semestr);
            dataGridView3.Columns.Clear();
            dataGridView4.Columns.Clear();
            AddColumnPlan();
            LoadDataPlanGroup();
            LoadDataPlanPrep();
            _semestr = semestr;
        }
        #endregion


        #region Смена строк в планах и нагрузке
        //Смена строки в нагрузке групп
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            int rowIndex = dataGridView1.CurrentCell.RowIndex;
            if (dataGridView1.Focused)
                dataGridView3.Rows[rowIndex].Cells[tekNed].Selected = true;
        }

        //Смена строки в нагрузке преподавателей
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            int rowIndex = dataGridView2.CurrentCell.RowIndex;
            if (dataGridView2.Focused)
                dataGridView4.Rows[rowIndex].Cells[tekNed].Selected = true;
        }

        //Смена строки в плане групп
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            int rowIndex = dataGridView3.CurrentCell.RowIndex;
            int dgv1RowsCount = dataGridView1.Rows.Count;
            if (dataGridView3.Focused && rowIndex < dgv1RowsCount)
                dataGridView1.Rows[rowIndex].Selected = true;
        }

        //Смена строки в плане преподавателей
        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            int rowIndex = dataGridView4.CurrentCell.RowIndex;
            int dgv2RowsCount = dataGridView2.Rows.Count;
            if (dataGridView4.Focused && rowIndex < dgv2RowsCount)
                dataGridView2.Rows[rowIndex].Selected = true;
        }
        #endregion


        #region Суммирование значений в столбце (план)
        //Подсчет элементов в столбце (план групп)
        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView3.BindingContext[dataGridView3.DataSource].EndCurrentEdit();    //Завершение редактирования текущей ячейки
            SummPlanGroupColumn();
        }

        //Подсчет элементов в столбце (план преподавателей)
        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView4.BindingContext[dataGridView4.DataSource].EndCurrentEdit();    //Завершение редактирования текущей ячейки
            SummPlanPrepColumn();
        }
        #endregion


        #region Суммирование значений в строке (нагрузка)
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
        #endregion


        #region Сумма по столбцам в плане (для групп и преподавателей при загрузке плана)
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
        #endregion


        #region Форматирование ячеек
        //Форматирование ячеек в соответствии с правилами (У, П и ПП)
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

        //Форматирование ячеек при превышении нагрузки (36 часов в неделю)
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
        #endregion


        #region Подсчет количества уроков в плане групп и преподавателей (по столбцам)
        //Подсчет количества уроков в плане группы (по столбцам)
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
        #endregion


        #region KeyDown (Delete)
        //Удаление значения текущей ячейки
        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                dataGridView3.CurrentCell.Value = null;
                dataGridView3.BindingContext[dataGridView3.DataSource].EndCurrentEdit();
                SummPlanGroupColumn();
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
            }
        }
        #endregion


        #region Enter
        private void dataGridView3_Enter(object sender, EventArgs e)
        {
            //Данные изменены
            is_edit = true;
        }

        private void dataGridView4_Enter(object sender, EventArgs e)
        {
            //Данные изменены
            is_edit = true;
        }
        #endregion


        #region Управление фокусом и размер сплит контейнера
        private void listBox1_Click(object sender, EventArgs e)
        {
            _listGroupClick = true;
            dataGridView3.Focus();
        }

        private void listBox2_Click(object sender, EventArgs e)
        {
            //Автовысота таблицы
            splitContainer1.SplitterDistance = 150;
            dataGridView4.Focus();
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            dataGridView3.Focus();
            
            if (_tableHeight == 0)
            {
                _tableHeight = 20 * dataGridView3.RowCount + 70;
            }
            splitContainer1.SplitterDistance = _tableHeight;
        }

        private void dataGridView2_Click(object sender, EventArgs e)
        {
            dataGridView4.Focus();
            splitContainer1.SplitterDistance = 150;
        }

        private void dataGridView3_Click(object sender, EventArgs e)
        {
            if (_tableHeight == 0)
            {
                _tableHeight = 20 * dataGridView3.RowCount + 70;
            }
            splitContainer1.SplitterDistance = _tableHeight;
        }

        private void dataGridView4_Click(object sender, EventArgs e)
        {
            splitContainer1.SplitterDistance = 150;
        }
        #endregion


        #region Menu
        //Обновить таблицу с расписанием (из файла БД программы Экспресс-расписание Колледж)
        private void menuOpenDB_Click(object sender, EventArgs e)
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

        //Заполнение таблицы с наименованием недель (1 и 2 семестры)
        private void menuFill_1term_Click(object sender, EventArgs e)
        {
            SetWeek(1);
        }
        private void menuFill_2term_Click(object sender, EventArgs e)
        {
            SetWeek(2);
        }

        //Меню для открытия окна отчета за месяц
        private void menuMonthlyReport_Click(object sender, EventArgs e)
        {
            Otchet form = new Otchet();
            form.ShowDialog();
        }

        //Меню информации о программе
        private void menuAbout_Click(object sender, EventArgs e)
        {
            About form = new About();
            form.ShowDialog();
        }

        //Выход из приложения
        private void menuExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        #endregion
    }
}
