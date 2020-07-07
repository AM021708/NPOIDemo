using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace NPOIDemo.Models
{
    public class RequestObj
    {
        public ReInventory[] inventory { get; set; }
        public ReOTHER[] OTHER { get; set; }
        public RePassinout[] Passinout { get; set; }
        public ReTran[] Trans { get; set; }
        public int id { get; set; }
        public string date { get; set; }
        public int status { get; set; }
        public int next_audit { get; set; }
        public int weather { get; set; }
        public string officer_main_give { get; set; }
        public string officer_sub_give { get; set; }
        public string note { get; set; }
    }


    public class ReInventory
    {
        public Inventoryitem[] inventoryitems { get; set; }
        public int id { get; set; }
        public string title { get; set; }
        public string liaisonmorning { get; set; }
        public string liaisonevening { get; set; }
        public int dutylog_id { get; set; }
    }

    public class Inventoryitem
    {
        public int id { get; set; }
        public string item { get; set; }
        public int checking { get; set; }
        public int inventory_id { get; set; }
    }

    public class ReOTHER
    {
        public int id { get; set; }
        public string title { get; set; }
        public string description { get; set; }
        public string remark { get; set; }
        public int dutylog_id { get; set; }
    }

    public class RePassinout
    {
        public int id { get; set; }
        public string unit { get; set; }
        public string place { get; set; }
        public string firm_leader_name { get; set; }
        public string firm_leader_tel { get; set; }
        public int amount { get; set; }
        public string works { get; set; }
        public string oversee_name { get; set; }
        public string oversee_tel { get; set; }
        public string remark { get; set; }
        public int dutylog_id { get; set; }
    }

    public class ReTran
    {
        public int id { get; set; }
        public string item { get; set; }
        public string amount { get; set; }
        public byte give { get; set; }
        public byte recieve { get; set; }
        public int dutylog_id { get; set; }
    }


}