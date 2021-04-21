﻿using System;
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
    public class KhoRuousController : Controller
    {
        private RuouDbContext db = new RuouDbContext();

        // GET: KhoRuous
        [Authorize]
        public ActionResult Index()
        {
            return View(db.KhoRuous.ToList());
        }

        // GET: KhoRuous/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            KhoRuou khoRuou = db.KhoRuous.Find(id);
            if (khoRuou == null)
            {
                return HttpNotFound();
            }
            return View(khoRuou);
        }

        // GET: KhoRuous/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: KhoRuous/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "IDRuou,TenRuou,NongDo,TinhChat,SoLuong,DonVi,TheTich,NguonGoc,DanhGia")] KhoRuou khoRuou)
        {
            if (ModelState.IsValid)
            {
                db.KhoRuous.Add(khoRuou);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(khoRuou);
        }

        // GET: KhoRuous/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            KhoRuou khoRuou = db.KhoRuous.Find(id);
            if (khoRuou == null)
            {
                return HttpNotFound();
            }
            return View(khoRuou);
        }

        // POST: KhoRuous/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "IDRuou,TenRuou,NongDo,TinhChat,SoLuong,DonVi,TheTich,NguonGoc,DanhGia")] KhoRuou khoRuou)
        {
            if (ModelState.IsValid)
            {
                db.Entry(khoRuou).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(khoRuou);
        }

        // GET: KhoRuous/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            KhoRuou khoRuou = db.KhoRuous.Find(id);
            if (khoRuou == null)
            {
                return HttpNotFound();
            }
            return View(khoRuou);
        }

        // POST: KhoRuous/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            KhoRuou khoRuou = db.KhoRuous.Find(id);
            db.KhoRuous.Remove(khoRuou);
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
