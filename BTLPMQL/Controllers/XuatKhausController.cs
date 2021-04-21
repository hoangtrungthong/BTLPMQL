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
    public class XuatKhausController : Controller
    {
        private RuouDbContext db = new RuouDbContext();

        // GET: XuatKhaus
        [Authorize]
        public ActionResult Index()
        {
            return View(db.XuatKhaus.ToList());
        }

        // GET: XuatKhaus/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            XuatKhau xuatKhau = db.XuatKhaus.Find(id);
            if (xuatKhau == null)
            {
                return HttpNotFound();
            }
            return View(xuatKhau);
        }

        // GET: XuatKhaus/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: XuatKhaus/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "IDRuou,TenRuou,NongDo,TinhChat,SoLuong,DonVi,TheTich,NguonGoc,NguoiXuat,NguoiNhap,Note")] XuatKhau xuatKhau)
        {
            if (ModelState.IsValid)
            {
                db.XuatKhaus.Add(xuatKhau);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(xuatKhau);
        }

        // GET: XuatKhaus/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            XuatKhau xuatKhau = db.XuatKhaus.Find(id);
            if (xuatKhau == null)
            {
                return HttpNotFound();
            }
            return View(xuatKhau);
        }

        // POST: XuatKhaus/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "IDRuou,TenRuou,NongDo,TinhChat,SoLuong,DonVi,TheTich,NguonGoc,NguoiXuat,NguoiNhap,Note")] XuatKhau xuatKhau)
        {
            if (ModelState.IsValid)
            {
                db.Entry(xuatKhau).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(xuatKhau);
        }

        // GET: XuatKhaus/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            XuatKhau xuatKhau = db.XuatKhaus.Find(id);
            if (xuatKhau == null)
            {
                return HttpNotFound();
            }
            return View(xuatKhau);
        }

        // POST: XuatKhaus/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            XuatKhau xuatKhau = db.XuatKhaus.Find(id);
            db.XuatKhaus.Remove(xuatKhau);
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
