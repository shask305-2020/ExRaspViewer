using System;
using System.Data.SqlClient;
using ExRaspViewer.Classes;
using System.Windows.Forms;
using System.Data;
using System.Collections.Generic;

namespace ExRaspViewer.Classes
{
    internal class ClassSqlDB
    {
        private SqlConnection connSql = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\ExRaspisDB.mdf;Integrated Security=True");
        private string[] tables = { "SPGRUP", "SPNAGR", "SPPRED", "SPPREP", "UROKI" };
        private ClassOleDB oleDB = new ClassOleDB();

        public SqlConnection ConnSql { get { return connSql; } }    //Получение строки подключения в другом классе

        //Загрузка списка групп в ListBox1
        public DataTable LoadGroupTable()
        {
            using (SqlCommand command = new SqlCommand("SELECT * FROM [dbo].[LoadGroup]", connSql))
            {
                DataTable table = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(table);
                return table;
            }
        }

        //Загрузка данных по нагрузке групп в DataGridView1
        //public DataTable LoadNagruzkaGrupp(int id)
        //{
        //    using (SqlCommand command = new SqlCommand("pr_LoadNagrGroup", connSql))
        //    {
        //        command.CommandType = CommandType.StoredProcedure;
        //        command.Parameters.AddWithValue("@IDG", id);
        //        DataTable table = new DataTable();
        //        SqlDataAdapter adapter = new SqlDataAdapter(command);
        //        adapter.Fill(table);
        //        table.Columns.Add("Вып.", typeof(int));
        //        table.Columns.Add("Ост.", typeof(int));
        //        table.Columns.Add("Ост.план", typeof(int));
        //        return table;
        //    }
        //}

        //Загрузка данных по нагрузке групп в DataGridView1
        public DataTable LoadNagruzkaGrupp(int id, string dat)
        {
            using (SqlCommand command = new SqlCommand("pr_LoadNagrGroup", connSql))
            {
                int idg, idp, idd, num, vsego;

                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@IDG", id);
                DataTable table = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(table);
                table.Columns.Add("Вып", typeof(int));
                table.Columns.Add("Ост", typeof(int));
                table.Columns.Add("Ост_план", typeof(int));
                int count = table.Rows.Count;
                if (count > 0)
                {
                    int sumVsego = 0, sumVyp = 0, sumOst = 0;
                    for (int i = 0; i < count; i++)
                    {
                        idg = (int)table.Rows[i][0];
                        idp = (int)table.Rows[i][1];
                        idd = (int)table.Rows[i][2];
                        num = CountUroki(idp, idg, idd, dat);
                        vsego = Convert.ToInt32(table.Rows[i][6]);
                        table.Rows[i][7] = num;             //Выполнено
                        table.Rows[i][8] = vsego - num;     //Остаток

                        //Вычисление суммы по столбцам
                        //Если подгруппа 2, то сумму не считаем
                        byte grup2 = 2;
                        if ((byte)table.Rows[i][5] != grup2)
                        {
                            sumVsego += vsego;
                            sumVyp += num;
                            sumOst += (int)table.Rows[i][8];
                        }
                    }
                    table.Rows.Add();
                    table.Rows[count][4] = "Сумма";
                    table.Rows[count][6] = sumVsego;
                    table.Rows[count][7] = sumVyp;
                    table.Rows[count][8] = sumOst;
                }

                return table;
            }
        }

        //Загрузка списка преподавателей в ListBox2
        public DataTable LoadPrepodTable()
        {
            using (SqlCommand command = new SqlCommand("pr_LoadPrepod", connSql))
            {
                command.CommandType = CommandType.StoredProcedure;

                DataTable table = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(table);
                return table;
            }
        }

        //Загрузка данных по нагрузке преподавателей в DataGridView2
        public DataTable LoadNagrPrepod(int id)
        {
            using (SqlCommand command = new SqlCommand("pr_LoadNagrPrepod", connSql))
            {
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@IDP", id);
                DataTable table = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(table);
                table.Columns.Add("Вып.", typeof(int));
                table.Columns.Add("Ост.", typeof(int));
                table.Columns.Add("Ост.план", typeof(int));
                return table;
            }
        }
                
        //Загрузка данных по нагрузке преподавателей в dgv в отчете
        public DataTable LoadNagrPrepodOtchet(int idPrepod, string dat_N, string dat_K)
        {
            using (SqlCommand command = new SqlCommand("pr_LoadNagrTeacher", connSql))
            {
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@IDP", idPrepod);
                DataTable table = new DataTable();
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                adapter.Fill(table);
                
                //Добавление столбцов
                table.Columns.Add("Вып. за период", typeof(int));
                table.Columns.Add("Вып. с нач. семестра", typeof(int));
                table.Columns.Add("Ост. на конец периода", typeof(int));
                
                //Заполнение добавленных столбцов данными
                int idg, idp, idd, num, vsego, countDT;
                idp = idPrepod;
                countDT = table.Rows.Count;
                if (countDT > 0)
                {
                    for (int i = 0; i < countDT; i++)
                    {
                        idg = (int)table.Rows[i][1];
                        idd = (int)table.Rows[i][2];
                        num = CountUroki(idp, idg, idd, dat_N, dat_K);    //Выполнено за период
                        vsego = Convert.ToInt32(table.Rows[i][6]);     //Всего часов
                        table.Rows[i][7] = num;            //Выполнено за период
                        int nachGoda = CountUroki(idp, idg, idd, dat_K);
                        table.Rows[i][8] = nachGoda; //Выполнено с начала года
                        table.Rows[i][9] = vsego - nachGoda;    //Остаток на конец периода
                    }
                }
                return table;
            }
        }

        //Удаление данных из рабочей таблицы перед обновлением данных
        public void DeleteDataFromDB()
        {
            using (SqlCommand command = new SqlCommand())
            {
                connSql.Open();
                try
                {
                    for (int i = 0; i < tables.Length; i++)
                    {
                        command.CommandText = "DELETE FROM " + tables[i];
                        command.Connection = connSql;
                        command.CommandType = CommandType.Text;
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    connSql.Close();
                }
            }
        }

        //Удаление данных из конкретной таблицы
        public void DeleteDataFromDB(string table)
        {
            using (SqlCommand command = new SqlCommand())
            {
                connSql.Open();
                try
                {
                    command.CommandText = "TRUNCATE TABLE " + table;
                    command.Connection = connSql;
                    command.CommandType = CommandType.Text;
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    connSql.Close();
                }
            }
        }
 
        //Обновление таблиц в рабочей базе данных
        public void UpdateSqlTable(string connection)
        {
            using (SqlBulkCopy copy = new SqlBulkCopy(connSql))
            {
                try
                {
                    connSql.Open();
                    for (int i = 0; i < tables.Length; i++)
                    {
                        copy.DestinationTableName = tables[i];
                        copy.WriteToServer(oleDB.UpdateSqlDBTable(tables[i], connection));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    connSql.Close();
                }
            }
        }

        //Подсчет количества проведенных уроков (с учетом даты)
        public int CountUroki(int idp, int idg, int idd, string dat)
        {
            using (SqlCommand command = new SqlCommand("CountUroki"))
            {
                connSql.Open();
                try
                {
                    command.Connection = connSql;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.Add("@IDP", SqlDbType.Int);
                    command.Parameters["@IDP"].Value = idp;

                    command.Parameters.Add("@IDG", SqlDbType.Int);
                    command.Parameters["@IDG"].Value = idg;

                    command.Parameters.Add("@IDD", SqlDbType.Int);
                    command.Parameters["@IDD"].Value = idd;

                    command.Parameters.Add("@DAT", SqlDbType.Date);
                    command.Parameters["@DAT"].Value = dat;

                    int value = (int)command.ExecuteScalar();
                    return value * 2;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return 0;
                }
                finally
                {
                    connSql.Close();
                }
            }
        }

        //Подсчет количества проведенных уроков (с учетом даты)
        public int CountUroki(int idp, int idg, int idd, string dat_N, string dat_K)
        {
            using (SqlCommand command = new SqlCommand("CountUroki_Otchet_interval"))
            {
                connSql.Open();
                try
                {
                    command.Connection = connSql;
                    command.CommandType = CommandType.StoredProcedure;

                    command.Parameters.Add("@IDP", SqlDbType.Int);
                    command.Parameters["@IDP"].Value = idp;

                    command.Parameters.Add("@IDG", SqlDbType.Int);
                    command.Parameters["@IDG"].Value = idg;

                    command.Parameters.Add("@IDD", SqlDbType.Int);
                    command.Parameters["@IDD"].Value = idd;

                    command.Parameters.Add("@DAT_N", SqlDbType.Date);
                    command.Parameters["@DAT_N"].Value = dat_N;

                    command.Parameters.Add("@DAT_K", SqlDbType.Date);
                    command.Parameters["@DAT_K"].Value = dat_K;

                    int value = (int)command.ExecuteScalar();
                    return value * 2;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return 0;
                }
                finally
                {
                    connSql.Close();
                }
            }
        }

        //Получение списка наименований недель, зависит от семестра
        private List<string> GetNedely(int semestr)
        {
            List<string> vs = new List<string>();
            int year = (int)DateTime.Today.Year;
            int st = 0;     //Количество недель
            int month;
            int dayNachSem;     //День начала семестра
            int kolNed;     //Количество недель в семестре
            if (semestr == 1)
            {
                month = 9;
                dayNachSem = 1;
                kolNed = 19;
            }
            else
            {
                month = 1;
                dayNachSem = 10;
                kolNed = 26;
            }
            DateTime date = new DateTime(year, month, dayNachSem);
            DateTime mon;   //понедельник
            DateTime sat;   //суббота

            string stMon;
            string stSat;
            string interval;

            while (st < kolNed)
            {
                if (date.DayOfWeek == DayOfWeek.Monday)
                {
                    mon = date;
                    stMon = mon.ToString("dd/MM");
                    sat = date.AddDays(5);
                    stSat = sat.ToString("dd/MM");
                    date = date.AddDays(7);
                    interval = String.Format("{0}-{1}", stMon, stSat);
                    vs.Add(interval);
                    st++;
                }
                else if ((date.Month == 9 && date.Day == 1 || date.Month == 1 && date.Day == 10) && date.DayOfWeek != DayOfWeek.Sunday)
                {
                    mon = date;
                    stMon = mon.ToString("dd/MM");
                    DateTime day = date;
                    while (day.DayOfWeek != DayOfWeek.Saturday)
                    { day = day.AddDays(1); }
                    stSat = day.ToString("dd/MM");
                    date = day;
                    interval = String.Format("{0}-{1}", stMon, stSat);
                    vs.Add(interval);
                }
                else
                {
                    date = date.AddDays(1);
                }
            }
            return vs;
        }

        //Запись наименований недель в таблицу с неделями
        public void LoadNed(int semestr)
        {
            List<string> ned = GetNedely(semestr);
            int length = ned.Count;
            using (SqlCommand command = new SqlCommand("INSERT INTO [NED] VALUES (@Naim)", connSql))
            {
                command.CommandType = CommandType.Text;
                connSql.Open();
                command.Parameters.Add("@Naim", SqlDbType.VarChar, 20);
                try
                {
                    for (int i = 0; i < length; i++)
                    {
                        command.Parameters["@Naim"].Value = ned[i];
                        command.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    connSql.Close();
                }
            }
        }

        //Загрузка наименований недель из БД
        public List<string> LoadListNedely()
        {
            List<string> resultList = new List<string>();
            resultList.Add("IDN");  //ИД нагрузки
            resultList.Add("IDP");  //Преподаватель
            resultList.Add("IDG");  //Группа
            resultList.Add("IDD");  //Дисциплина
            SqlDataReader reader = null;
            using (SqlCommand cmd = new SqlCommand("SELECT [NAIM] FROM [NED]", connSql))
            {
                cmd.CommandType = CommandType.Text;
                try
                {
                    connSql.Open();
                    reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        resultList.Add(reader[0].ToString());
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (reader != null && !reader.IsClosed)
                    {
                        reader.Close();
                        connSql.Close();
                    }
                }
            }
            return resultList;
        }

        //Синхронизация преподавателей, групп, предметов в таблице План
        public void SyncPlan()
        {
            SqlCommand cmd = new SqlCommand("MergePlan", connSql);
            cmd.CommandType = CommandType.StoredProcedure;
            connSql.Open();
            cmd.ExecuteNonQuery();
            connSql.Close();
        }

        //Загрузка списка месяцев
        public DataTable LoadMonth()
        {
            string sql = "SELECT * FROM [dbo].[MONTH]";
            SqlCommand cmd = new SqlCommand(sql, connSql);
            DataTable dt = new DataTable();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            adapter.Fill(dt);
            return dt;
        }
    }
}
