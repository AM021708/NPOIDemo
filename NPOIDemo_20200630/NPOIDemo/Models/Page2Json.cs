using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace NPOIDemo.Models
{
    public class Page2Json
    {
    }


    public class Rootobject2
    {
        public Datum2[] data { get; set; }
    }

    public class Datum2
    {
        public string date { get; set; }
        public int status { get; set; }
        public Duty duty { get; set; }
        public string remark { get; set; }
    }

    public class Duty
    {
        public Morning morning { get; set; }
        public Evening evening { get; set; }
    }

    public class Morning
    {
        public string time { get; set; }
        public string work { get; set; }
        public bool execute { get; set; }
        public Remark[] remark { get; set; }
    }

    public class Remark
    {
        public string label { get; set; }
        public string type { get; set; }
        public string value { get; set; }
    }

    public class Evening
    {
        public string time { get; set; }
        public string work { get; set; }
        public bool execute { get; set; }
        public string remark { get; set; }
    }

}