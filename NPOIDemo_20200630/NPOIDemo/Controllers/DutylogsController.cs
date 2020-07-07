using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Http.Description;
using System.Web.ModelBinding;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using NPOIDemo.Models;

namespace NPOIDemo.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class DutylogsController : ApiController
    {
        private PAdminEntities db = new PAdminEntities();

        /*public string GetDutylog()
        {
            List<PAdmin> pAdmins = new List<PAdmin>();
            var data = new
            {
                db.Dutylog,
                db.Trans,
                db.inventory,
                db.inventoryitems,
                db.OTHER,
                db.Passinout
            };
            //dutylogs.ToList();
            
            var data2 = db.Trans.Select(s => s).ToList();

            
            string jsonData = JsonConvert.SerializeObject(data);

            return jsonData;
        }*/
        // 
        // GET: api/Dutylogs

        public IHttpActionResult GetDutylog()
        {
            System.Net.Http.Headers.HttpRequestHeaders headers = this.Request.Headers;
            int page = 0;
            int amount = 0;
            int start = 0;
            string sPage = string.Empty;
            string sAmount = string.Empty;


            if (headers.Contains("page"))
            {
                sPage = headers.GetValues("page").First();
            }
            if (headers.Contains("amount"))
            {
                sAmount = headers.GetValues("amount").First();
            }
            page = int.Parse(sPage);
            amount = int.Parse(sAmount);

            //where = "status = 0"
            //var dutylogdata2 = db.Dutylog.Where().ToList();
            if (page <= 0)
            {
                return BadRequest();
            }

            if (page == 1)
            {
                start = 0;
            }
            else
            {
                start = (page - 1) * 10;

            }
            //dutylogdata[data1, data2, data3]
            //10 * 3
            

            var dutylogdata = db.Dutylog.OrderByDescending(d => d.id).Skip(start).Take(amount).Select(d => new { d.id, d.date, d.status, d.next_audit, d.officer_main_give, d.officer_sub_give, d.note }).ToList();


            var outputData = new
            {
                total = db.Dutylog.Count(),
                dutylogdata
            };

            return Ok(outputData);
        }

        // GET: api/Dutylogs/5
        [ResponseType(typeof(Dutylog))]
        public IHttpActionResult GetDutylog(int id)
        {
            Dutylog dutylog = db.Dutylog.Find(id);
            var test = from t in db.Trans where t.dutylog_id == id select t;
            if (dutylog == null)
            {
                return NotFound();
            }
            var dutylogdata = db.Dutylog.Where(d => d.id == id);
            //var dutylogdata2 = db.Dutylog.Where(d => d.status == 0).ToList();
            /*
            var Transdata = db.Trans.Where(t => t.dutylog_id == id);
            var Inventorydata = db.inventory.Where(i => i.dutylog_id == id);
            var Inventoryitemsdata = from it in db.inventoryitems join i in db.inventory on it.inventory_id equals i.id where i.dutylog_id == id select it;
            var OTHERdata = db.OTHER.Where(o => o.dutylog_id == id);
            var Passinout = db.Passinout.Where(p => p.dutylog_id == id);
            var data = new
            {
                dutylogdata,
                Transdata,
                Inventorydata,
                Inventoryitemsdata,
                OTHERdata,
                Passinout
            };
            data.ToString();
            */

            //string jsonData = JsonConvert.SerializeObject(data);

            return Ok(dutylogdata);
        }

        struct TestObject
        {
            int id;
        }
        // PUT: api/Dutylogs/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutDutylog(int id)
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            Rootobject robj = JsonConvert.DeserializeObject<Rootobject>(System.IO.File.ReadAllText(path + "\\page1.json"));
            Dutylog dutylog = new Dutylog();
            dutylog = db.Dutylog.Find(id);
            string req_txt;
            using (StreamReader reader = new StreamReader(HttpContext.Current.Request.InputStream))
            {
                req_txt = reader.ReadToEnd();
            }

            RequestObj reobj = JsonConvert.DeserializeObject<RequestObj>(req_txt);


            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != dutylog.id)
            {
                return BadRequest();
            }
            dutylog.id = id;
            dutylog.date = reobj.date;
            dutylog.status = reobj.status;
            dutylog.next_audit = reobj.next_audit;
            dutylog.weather = reobj.weather;
            dutylog.officer_main_give = reobj.officer_main_give;
            dutylog.officer_sub_give = reobj.officer_sub_give;
            dutylog.note = reobj.note;
            db.Entry(dutylog).State = EntityState.Modified;
            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!DutylogExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            int tC = 0;

            for (int i = 0; i <= reobj.Trans.Length * id; i++)
            {
                Trans trans = db.Trans.Find(i);
                if (trans == null)
                    continue;
                if (trans.dutylog_id == id)
                {
                    trans.item = reobj.Trans[tC].item;
                    trans.amount = reobj.Trans[tC].amount;
                    trans.give = reobj.Trans[tC].give;
                    trans.recieve = reobj.Trans[tC].recieve;
                    trans.dutylog_id = id;
                    db.Entry(trans).State = EntityState.Modified;
                    try
                    {
                        db.SaveChanges();
                        tC++;
                    }
                    catch (DbUpdateConcurrencyException)
                    {
                        if (!DutylogExists(id))
                        {
                            return NotFound();
                        }
                        else
                        {
                            throw;
                        }
                    }
                }
            }
            tC = 0;

            int iC = 0;
            int itC = 0;

            for (int i = 0; i < reobj.inventory.Length * id; i++)
            {
                inventory inventory = new inventory();
                if (inventory == null)
                {
                    continue;
                }
                if (inventory.dutylog_id == id)
                {
                    for (int j = 0; j < reobj.inventory[i].inventoryitems.Length; j++)
                    {
                        inventoryitems inventoryitems = new inventoryitems();
                        if (inventoryitems == null)
                        {
                            continue;
                        }
                        if (inventoryitems.inventory_id == inventory.id)
                        {
                            inventoryitems.item = reobj.inventory[i].inventoryitems[j].item;
                            inventoryitems.checking = (byte?)reobj.inventory[i].inventoryitems[j].checking;
                            inventoryitems.inventory_id = inventory.id;
                            db.Entry(inventoryitems).State = EntityState.Modified;
                            try
                            {
                                db.SaveChanges();
                                itC++;
                            }
                            catch (DbUpdateConcurrencyException)
                            {
                                if (!DutylogExists(id))
                                {
                                    return NotFound();
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }
                    itC++;
                    inventory.title = robj.data[0].inventory[i].title;
                    inventory.liaisonevening = robj.data[0].inventory[i].liaison.evening;
                    inventory.dutylog_id = id;
                    db.Entry(inventory).State = EntityState.Modified;
                }
                try
                {
                    db.SaveChanges();
                    iC++;
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!DutylogExists(id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }

            }
            iC++;

            //修改到這邊，剩兩個TABLE


            int oC = 0;

            for (int i = 0; i < reobj.OTHER.Length * id; i++)
            {
                OTHER others = db.OTHER.Find(i);
                if (others == null)
                    continue;
                if (others.dutylog_id == id)
                {
                    others.title = reobj.OTHER[oC].title;
                    others.description = reobj.OTHER[oC].description;
                    others.remark = reobj.OTHER[oC].remark;
                    others.dutylog_id = id;
                    db.Entry(others).State = EntityState.Modified;

                    try
                    {
                        db.SaveChanges();
                        oC++;
                    }
                    catch (DbUpdateConcurrencyException)
                    {
                        if (!DutylogExists(id))
                        {
                            return NotFound();
                        }
                        else
                        {
                            throw;
                        }
                    }
                }
            }
            oC = 0;

            int pC = 0;
            for (int i = 0; i < reobj.Passinout.Length * id; i++)
            {
                Passinout passinout = db.Passinout.Find(i);
                if (passinout == null)
                    continue;
                if (passinout.dutylog_id == id)
                {
                    passinout.unit = reobj.Passinout[pC].unit;
                    passinout.place = reobj.Passinout[pC].place;
                    passinout.firm_leader_name = reobj.Passinout[pC].firm_leader_name;
                    passinout.firm_leader_tel = reobj.Passinout[pC].firm_leader_tel;
                    passinout.amount = reobj.Passinout[pC].amount;
                    passinout.works = reobj.Passinout[pC].works;
                    passinout.oversee_name = reobj.Passinout[pC].oversee_name;
                    passinout.oversee_tel = reobj.Passinout[pC].oversee_tel;
                    passinout.remark = reobj.Passinout[pC].remark;
                    passinout.dutylog_id = id;
                    db.Entry(passinout).State = EntityState.Modified;

                    try
                    {
                        db.SaveChanges();
                        pC++;
                    }
                    catch (DbUpdateConcurrencyException)
                    {
                        if (!DutylogExists(id))
                        {
                            return NotFound();
                        }
                        else
                        {
                            throw;
                        }
                    }
                }
            }
            pC = 0;



            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/Dutylogs
        [ResponseType(typeof(Dutylog))]
        public IHttpActionResult PostDutylog()
        {
            /*
            Rootobject robj = new Rootobject();
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            robj = JsonConvert.DeserializeObject<Rootobject>(System.IO.File.ReadAllText(path + "\\page1.json"));
            Dutylog dutylog = new Dutylog();

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            dutylog.date = robj.data[0].date;
            dutylog.status = robj.data[0].status;
            dutylog.next_audit = robj.data[0].next_audit;
            dutylog.weather = robj.data[0].weather;
            dutylog.officer_main_give = robj.data[0].officer.main.give;
            dutylog.officer_sub_give = robj.data[0].officer.sub.give;
            dutylog.note = "put";
            db.Dutylog.Add(dutylog);

            Trans trans = new Trans();
            for (int i = 0; i < robj.data[0].transaction.Length; i++)
            {
                trans.item = robj.data[0].transaction[i].item;
                trans.amount = robj.data[0].transaction[i].amount;
                if (robj.data[0].transaction[i].give)
                {
                    trans.give = 1;
                }
                else
                {
                    trans.give = 0;
                }
                if (robj.data[0].transaction[i].recieve)
                {
                    trans.recieve = 1;
                }
                else
                {
                    trans.recieve = 0;
                }
                trans.dutylog_id = dutylog.id;
                db.Trans.Add(trans);
                db.SaveChanges();
            }
            inventory inventory = new inventory();
            inventoryitems inventoryitems = new inventoryitems();
            for (int i = 0; i < robj.data[0].inventory.Length; i++)
            {
                inventory.title = robj.data[0].inventory[i].title;
                inventory.liaisonmorning = robj.data[0].inventory[i].liaison.morning;
                inventory.liaisonevening = robj.data[0].inventory[i].liaison.evening;
                inventory.dutylog_id = dutylog.id;
                db.inventory.Add(inventory);
                for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                {
                    inventoryitems.item = robj.data[0].inventory[i].items[j].item;
                    if (robj.data[0].inventory[i].items[j].check)
                    {
                        inventoryitems.checking = 1;
                    }
                    else
                    {
                        inventoryitems.checking = 0;
                    }
                    inventoryitems.inventory_id = inventory.id;
                    db.inventoryitems.Add(inventoryitems);
                    db.SaveChanges();
                }
            }

            OTHER others = new OTHER();
            for (int i = 0; i < robj.data[0].others.Length; i++)
            {
                others.title = robj.data[0].others[i].title;
                others.description = robj.data[0].others[i].description;
                others.remark = robj.data[0].others[i].remark;
                others.dutylog_id = dutylog.id;
                db.OTHER.Add(others);
                db.SaveChanges();
            }
            Passinout passinout = new Passinout();
            for (int i = 0; i < robj.data[0].pass_in_out.Length; i++)
            {
                passinout.unit = robj.data[0].pass_in_out[i].unit;
                passinout.place = robj.data[0].pass_in_out[i].place;
                passinout.firm_leader_name = robj.data[0].pass_in_out[i].firm_leader.name;
                passinout.firm_leader_tel = robj.data[0].pass_in_out[i].firm_leader.tel;
                passinout.amount = robj.data[0].pass_in_out[i].amount;
                passinout.works = robj.data[0].pass_in_out[i].work;
                passinout.oversee_name = robj.data[0].pass_in_out[i].oversee.name;
                passinout.oversee_tel = robj.data[0].pass_in_out[i].oversee.tel;
                passinout.remark = robj.data[0].pass_in_out[i].remark;
                passinout.dutylog_id = dutylog.id;
                db.Passinout.Add(passinout);
                db.SaveChanges();
            }

            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = dutylog.id }, dutylog);
            */


            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            Rootobject robj = JsonConvert.DeserializeObject<Rootobject>(System.IO.File.ReadAllText(path + "\\page1.json"));
            Dutylog dutylog = new Dutylog();
            string postReq = "";
            StreamReader reader = new StreamReader(HttpContext.Current.Request.InputStream);
            postReq = reader.ReadToEnd();
            RequestObj reobj = JsonConvert.DeserializeObject<RequestObj>(postReq);

            dutylog.date = reobj.date;
            dutylog.status = reobj.status;
            dutylog.next_audit = reobj.next_audit;
            dutylog.weather = reobj.weather;
            dutylog.officer_main_give = reobj.officer_main_give;
            dutylog.officer_sub_give = reobj.officer_sub_give;
            dutylog.note = reobj.note;
            db.Dutylog.Add(dutylog);
            db.SaveChanges();

            Trans trans = new Trans();
            for (int i = 0; i < reobj.Trans.Length; i++)
            {
                trans.item = reobj.Trans[i].item;
                trans.amount = reobj.Trans[i].amount;
                trans.give = reobj.Trans[i].give;
                trans.recieve = reobj.Trans[i].recieve;
                trans.dutylog_id = dutylog.id;
                db.Trans.Add(trans);
                db.SaveChanges();
            }
            inventory inventory = new inventory();
            inventoryitems inventoryitems = new inventoryitems();
            for (int i = 0; i < reobj.inventory.Length; i++)
            {
                inventory.title = reobj.inventory[i].title;
                inventory.liaisonmorning = reobj.inventory[i].liaisonmorning;
                inventory.liaisonevening = reobj.inventory[i].liaisonevening;
                inventory.dutylog_id = dutylog.id;
                db.inventory.Add(inventory);
                for (int j = 0; j < reobj.inventory[i].inventoryitems.Length; j++)
                {
                    inventoryitems.item = reobj.inventory[i].inventoryitems[j].item;
                    inventoryitems.checking = (byte?)reobj.inventory[i].inventoryitems[j].checking;
                    inventoryitems.inventory_id = inventory.id;
                    db.inventoryitems.Add(inventoryitems);
                    db.SaveChanges();
                }
            }

            OTHER others = new OTHER();
            for (int i = 0; i < reobj.OTHER.Length; i++)
            {
                others.title = reobj.OTHER[i].title;
                others.description = reobj.OTHER[i].description;
                others.remark = reobj.OTHER[i].remark;
                others.dutylog_id = dutylog.id;
                db.OTHER.Add(others);
                db.SaveChanges();
            }
            Passinout passinout = new Passinout();
            for (int i = 0; i < reobj.Passinout.Length; i++)
            {
                passinout.unit = reobj.Passinout[i].unit;
                passinout.place = reobj.Passinout[i].place;
                passinout.firm_leader_name = reobj.Passinout[i].firm_leader_name;
                passinout.firm_leader_tel = reobj.Passinout[i].firm_leader_tel;
                passinout.amount = reobj.Passinout[i].amount;
                passinout.works = reobj.Passinout[i].works;
                passinout.oversee_name = reobj.Passinout[i].oversee_name;
                passinout.oversee_tel = reobj.Passinout[i].oversee_tel;
                passinout.remark = reobj.Passinout[i].remark;
                passinout.dutylog_id = dutylog.id;
                db.Passinout.Add(passinout);
                db.SaveChanges();
            }


            return CreatedAtRoute("DefaultApi", new { id = dutylog.id }, dutylog);
        }

        // DELETE: api/Dutylogs/5
        [ResponseType(typeof(Dutylog))]
        public IHttpActionResult DeleteDutylog(int id)
        {
            Dutylog dutylog = db.Dutylog.Find(id);
            if (dutylog == null)
            {
                return NotFound();
            }

            db.Dutylog.Remove(dutylog);
            db.SaveChanges();

            return Ok(dutylog);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool DutylogExists(int id)
        {
            return db.Dutylog.Count(e => e.id == id) > 0;
        }

        private bool TransExists(int id)
        {
            return db.Trans.Count(t => t.dutylog_id == id) > 0;
        }

        private bool InventoryExists(int id)
        {
            return db.inventory.Count(i => i.dutylog_id == id) > 0;
        }

        private bool OTHERExists(int id)
        {
            return db.OTHER.Count(o => o.dutylog_id == id) > 0;
        }

        private bool PassinoutExists(int id)
        {
            return db.Passinout.Count(p => p.dutylog_id == id) > 0;
        }

    }
}