using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using BTLPMQL.Models;

namespace BTLPMQL.Controllers
{
    public class NhapKhausController : Controller
    {
        private RuouDbContext db = new RuouDbContext();

        // GET: NhapKhaus
        [Authorize]
        public ActionResult Index()
        {
            return View(db.NhapKhaus.ToList());
        }

        // GET: NhapKhaus/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            NhapKhau nhapKhau = db.NhapKhaus.Find(id);
            if (nhapKhau == null)
            {
                return HttpNotFound();
            }
            return View(nhapKhau);
        }

        // GET: NhapKhaus/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: NhapKhaus/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "IDRuou,TenRuou,NongDo,TinhChat,SoLuong,DonVi,TheTich,NguonGoc,NguoiXuat,NguoiNhap,Note")] NhapKhau nhapKhau)
        {
            if (ModelState.IsValid)
            {
                db.NhapKhaus.Add(nhapKhau);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(nhapKhau);
        }

        // GET: NhapKhaus/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            NhapKhau nhapKhau = db.NhapKhaus.Find(id);
            if (nhapKhau == null)
            {
                return HttpNotFound();
            }
            return View(nhapKhau);
        }

        // POST: NhapKhaus/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "IDRuou,TenRuou,NongDo,TinhChat,SoLuong,DonVi,TheTich,NguonGoc,NguoiXuat,NguoiNhap,Note")] NhapKhau nhapKhau)
        {
            if (ModelState.IsValid)
            {
                db.Entry(nhapKhau).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(nhapKhau);
        }

        // GET: NhapKhaus/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            NhapKhau nhapKhau = db.NhapKhaus.Find(id);
            if (nhapKhau == null)
            {
                return HttpNotFound();
            }
            return View(nhapKhau);
        }

        // POST: NhapKhaus/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            NhapKhau nhapKhau = db.NhapKhaus.Find(id);
            db.NhapKhaus.Remove(nhapKhau);
            db.SaveChanges();
            return RedirectToAction("Index");
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
