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
        private PAdminDutylogWebApi PAdminDutylogWebApi = new PAdminDutylogWebApi();
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

            return Ok(dutylogdata);
        }

        // PUT: api/Dutylogs/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutDutylog(int id)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            PAdminDutylogWebApi.SetDutylog(id);
            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/Dutylogs
        [ResponseType(typeof(Dutylog))]
        public IHttpActionResult PostDutylog()
        {
            bool bFakeData = false;
            if (bFakeData)
            {
                if (PAdminDutylogWebApi.AddDutylog())
                {
                    return Ok();
                }
                else
                {
                    return BadRequest();
                }  
            }
            else
            {
                if (PAdminDutylogWebApi.AddFakeDutylogData())
                {
                    return Ok();
                }
                else
                {
                    return BadRequest();
                }
            } 
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
    }
}