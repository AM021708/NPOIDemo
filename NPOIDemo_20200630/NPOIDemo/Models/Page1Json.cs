using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;

namespace NPOIDemo.Models
{
    public class Page1Json
    {
    }

    public static class DateTimeExtensions
    {
        /// <summary>
        /// To the full taiwan date.
        /// </summary>
        /// <param name="datetime">The datetime.</param>
        /// <returns></returns>
        public static string ToFullTaiwanDate(this DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            return string.Format(" {0} 年 {1} 月 {2} 日",
                taiwanCalendar.GetYear(datetime),
                datetime.Month,
                datetime.Day);
        }

        /// <summary>
        /// To the simple taiwan date.
        /// </summary>
        /// <param name="datetime">The datetime.</param>
        /// <returns></returns>
        public static string ToSimpleTaiwanDate(this DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            return string.Format("{0}/{1}/{2}",
                taiwanCalendar.GetYear(datetime),
                datetime.Month,
                datetime.Day);
        }
    }


    public class Rootobject
    {
        public Datum[] data { get; set; }
    }

    public class Datum
    {
        public string date { get; set; }
        public int status { get; set; }
        public int next_audit { get; set; }
        public int weather { get; set; }
        public Officer officer { get; set; }
        public Transaction[] transaction { get; set; }
        public Inventory[] inventory { get; set; }
        public Others[] others { get; set; }
        public string note { get; set; }
        public string approve { get; set; }
        public Pass_In_Out[] pass_in_out { get; set; }
    }

    public class Officer
    {
        public Main main { get; set; }
        public Sub sub { get; set; }
    }

    public class Main
    {
        public string give { get; set; }
        public string recieve { get; set; }
    }

    public class Sub
    {
        public string give { get; set; }
        public string recieve { get; set; }
    }

    public class Transaction
    {
        public string item { get; set; }
        public string amount { get; set; }
        public bool give { get; set; }
        public bool recieve { get; set; }
    }

    public class Inventory
    {
        public string title { get; set; }
        public Item[] items { get; set; }
        public Liaison liaison { get; set; }
    }

    public class Liaison
    {
        public string morning { get; set; }
        public string afternoon { get; set; }
        public string evening { get; set; }
        public string midnight { get; set; }
    }

    public class Item
    {
        public string item { get; set; }
        public bool check { get; set; }
    }

    public class Others
    {
        public string title { get; set; }
        public string description { get; set; }
        public string remark { get; set; }
    }

    public class Pass_In_Out
    {
        public string unit { get; set; }
        public string place { get; set; }
        public Firm_Leader firm_leader { get; set; }
        public int amount { get; set; }
        public string work { get; set; }
        public Oversee oversee { get; set; }
        public string remark { get; set; }
    }

    public class Firm_Leader
    {
        public string name { get; set; }
        public string tel { get; set; }
    }

    public class Oversee
    {
        public string name { get; set; }
        public string tel { get; set; }
    }

}