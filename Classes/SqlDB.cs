using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ExRaspViewer.Classes
{
    public static class SqlDB
    {
        private static SqlConnection connection = new SqlConnection(Properties.Settings.Default.connection);
        private static string[] tables = { "SPGRUP", "SPNAGR", "SPPRED", "SPPREP", "UROKI" };
        private static string sql;
        private static SqlCommand cmd;
        private static SqlDataAdapter sqlAdapter;
        private static DataTable dt;

        public static SqlConnection ConnSql { get { return connection; } }    //Получение строки подключения в другом классе

        //Загрузка данных из таблиц и представлений
        public static DataTable LoadTable(string tableName)
        {
            sql = $"SELECT * FROM {tableName}";
            cmd = new SqlCommand(sql, connection);
            dt = new DataTable();
            sqlAdapter = new SqlDataAdapter(cmd);
            sqlAdapter.Fill(dt);
            return dt;
        }


        //Загрузка данных по нагрузке групп в DataGridView1
        public static DataTable GroupLoad(int id, string data)
        {
            using (cmd = new SqlCommand("pr_GroupLoad", connection))
            {
                int num, vsego;
                int sumVsego = 0, sumVyp = 0, sumOst = 0;

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@IDG", id);

                cmd.Parameters.Add("@DAT", SqlDbType.Date);
                cmd.Parameters["@DAT"].Value = data;

                dt = new DataTable();
                sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);

                dt.Columns.Add("Ост", typeof(int));
                dt.Columns.Add("Ост_план", typeof(int));
                int count = dt.Rows.Count;
                if (count > 0)
                {
                    for (int i = 0; i < count; i++)
                    {
                        num = (int)dt.Rows[i][7];           //Выполнено
                        vsego = Convert.ToInt32(dt.Rows[i][6]);
                        dt.Rows[i][8] = vsego - num;        //Остаток

                        //Вычисление суммы по столбцам
                        //Если подгруппа 2, то сумму не считаем
                        byte grup2 = 2;
                        if ((byte)dt.Rows[i][5] != grup2)
                        {
                            sumVsego += vsego;
                            sumVyp += num;
                            sumOst += (int)dt.Rows[i][8];
                        }
                    }
                    dt.Rows.Add();
                    dt.Rows[count][4] = "Сумма";
                    dt.Rows[count][6] = sumVsego;
                    dt.Rows[count][7] = sumVyp;
                    dt.Rows[count][8] = sumOst;
                }
                return dt;
            }
        }

        //Загрузка данных по нагрузке преподавателей в DataGridView2
        public static DataTable TeacherLoad(int id, string data)
        {
            using (cmd = new SqlCommand("pr_TeacherLoad", connection))
            {
                int num, vsego;
                int sumVsego = 0, sumVyp = 0, sumOst = 0;

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@IDP", id);

                cmd.Parameters.Add("@DAT", SqlDbType.Date);
                cmd.Parameters["@DAT"].Value = data;

                dt = new DataTable();
                sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);

                dt.Columns.Add("Ост", typeof(int));
                dt.Columns.Add("Ост_план", typeof(int));
                int count = dt.Rows.Count;
                if (count > 0)
                {
                    for (int i = 0; i < count; i++)
                    {
                        num = (int)dt.Rows[i][7];           //Выполнено
                        vsego = Convert.ToInt32(dt.Rows[i][6]);
                        dt.Rows[i][8] = vsego - num;        //Остаток

                        //Вычисление суммы по столбцам
                        sumVsego += vsego;
                        sumVyp += num;
                        sumOst += (int)dt.Rows[i][8];
                    }
                    dt.Rows.Add();
                    dt.Rows[count][4] = "Сумма";
                    dt.Rows[count][6] = sumVsego;
                    dt.Rows[count][7] = sumVyp;
                    dt.Rows[count][8] = sumOst;
                }
                return dt;
            }
        }
    }
}
