using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.IO;
using System.Linq;
using System.Web;

namespace NPOIDemo.Models
{
    public class PAdminDutylogWebApi
    {
        private PAdminEntities db = new PAdminEntities();
        
        public bool SetDutylog(int id)
        {
            Dutylog dutylog = new Dutylog();
            dutylog = db.Dutylog.Find(id);
            string req_txt;
            using (StreamReader reader = new StreamReader(HttpContext.Current.Request.InputStream))
            {
                req_txt = reader.ReadToEnd();
            }

            RequestObj reobj = JsonConvert.DeserializeObject<RequestObj>(req_txt);

            if (id != dutylog.id)
            {
                return false;
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
                    return false;
                }
                else
                {
                    throw;
                }
            }

            int tC = 0;
            for (int i = 0; i <= reobj.Trans.Length * id; i++)
            {
                Trans trans = db.Trans.Find(i + 1);
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
                            return false;
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
                inventory inventory = db.inventory.Find(id + 1);
                if (inventory == null)
                {
                    continue;
                }
                if (inventory.dutylog_id == id)
                {
                    for (int j = 0; j < reobj.inventory[i].inventoryitems.Length; j++)
                    {
                        inventoryitems inventoryitems = db.inventoryitems.Find(id + 1);
                        if (inventoryitems == null)
                        {
                            continue;
                        }
                        if (inventoryitems.inventory_id == inventory.id)
                        {
                            inventoryitems.item = reobj.inventory[iC].inventoryitems[itC].item;
                            inventoryitems.checking = (byte?)reobj.inventory[iC].inventoryitems[itC].checking;
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
                                    return false;
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }
                    inventory.title = reobj.inventory[iC].title;
                    inventory.liaisonmorning = reobj.inventory[iC].liaisonmorning;
                    inventory.liaisonevening = reobj.inventory[iC].liaisonevening;
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
                        return false;
                    }
                    else
                    {
                        throw;
                    }
                }

            }
            iC++;

            int oC = 0;
            for (int i = 0; i < reobj.OTHER.Length * id; i++)
            {
                OTHER others = db.OTHER.Find(i + 1);
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
                            return false;
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
                Passinout passinout = db.Passinout.Find(i + 1);
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
                            return false;
                        }
                        else
                        {
                            throw;
                        }
                    }
                }
            }
            pC = 0;
            return true;
        }
        public bool AddDutylog()
        {
            Dutylog dutylog = new Dutylog();
            string postReq = "";
            using (StreamReader reader = new StreamReader(HttpContext.Current.Request.InputStream))
            {
                postReq = reader.ReadToEnd();
            }
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
            return true;
        }

        public bool AddFakeDutylogData()
        {
            Rootobject robj = new Rootobject();
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            robj = JsonConvert.DeserializeObject<Rootobject>(System.IO.File.ReadAllText(path + "\\page1.json"));
            
            Dutylog dutylog = new Dutylog();
            dutylog.date = robj.data[0].date;
            dutylog.status = robj.data[0].status;
            dutylog.next_audit = robj.data[0].next_audit;
            dutylog.weather = robj.data[0].weather;
            dutylog.officer_main_give = robj.data[0].officer.main.give;
            dutylog.officer_sub_give = robj.data[0].officer.sub.give;
            dutylog.note = robj.data[0].note;
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
            return true;
        }
        private bool DutylogExists(int id)
        {
            return db.Dutylog.Count(e => e.id == id) > 0;
        }
    }
}