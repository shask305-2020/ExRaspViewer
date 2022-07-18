using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExRaspViewer.Classes
{
    public class Pedagog
    {
        public int Id_Pedagog { get; set; }
        public string Fio_Pedagog { get; set; }
    }
    public class Service
    {
        public int NumNed { get; set; }
        public DateTime DateNachNed { get; set; }
        public DateTime DateKonNed { get; set; }


        //Создание списка для вычисления номера недели
        public List<Service> GetListNedely(int semestr)
        {
            int year = (int)DateTime.Today.Year;
            int st = 0;     //Количество недель
            int month;
            int dayNachSem;     //День начала семестра
            int kolNed;     //Количество недель в семестре

            List<Service> list = new List<Service>();

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
            DateTime sun;   //воскресенье


            while (st < kolNed)
            {
                if (date.DayOfWeek == DayOfWeek.Monday)
                {
                    mon = date;
                    sun = date.AddDays(6);
                    date = date.AddDays(7);
                    st++;
                    list.Add(new Service() { NumNed = st, DateNachNed = mon, DateKonNed = sun });
                }
                else if ((date.Month == 9 && date.Day == 1 || date.Month == 1 && date.Day == 10) && date.DayOfWeek != DayOfWeek.Sunday)
                {
                    mon = date;
                    DateTime day = date;
                    while (day.DayOfWeek != DayOfWeek.Sunday)
                    { day = day.AddDays(1); }
                    sun = day;
                    date = sun;
                    st++;
                    list.Add(new Service() { NumNed = st, DateNachNed = mon, DateKonNed = sun });
                }
                else
                {
                    date = date.AddDays(1);
                }
            }
            return list;
        }


        //Проверка, является ли строка числом
        public bool StringIsDigit(string str)
        {
            bool result = true;
            foreach (char item in str)
            {
                result = result && Char.IsDigit(item);
            }
            return result;
        }
    }
}
