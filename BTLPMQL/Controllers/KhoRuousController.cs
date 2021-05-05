using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using BTLPMQL.Models;
using System.Data.OleDb;

namespace BTLPMQL.Controllers
{
    public class KhoRuousController : Controller
    {
        private RuouDbContext db = new RuouDbContext();
        ReadExcel excel = new ReadExcel();
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

        public ActionResult UploadFile(HttpPostedFileBase file)
        {
            try
            {
                if (file.ContentLength > 0)
                {

                    string _FileName = DateTime.Now.Year.ToString() + DateTime.Now.Date.ToString();
 
                    string _path = Path.Combine(Server.MapPath("~/Uploads/ExcelFile/"), _FileName);
 
                    file.SaveAs(_path);
  
                    CopyDataByBulk(ReadDataFromExcelFile(_path));
                    ViewBag.ThongBao = "Cập Nhật Kho Rượu Thành Công";
                }
            }
            catch (Exception ex)
            {
                ViewBag.ThongBao = "Lỗi Rồi";
            }
            return View("Index");
        }

        private void UploadExcelFile(HttpPostedFileBase file)
        {

            string _FileName = "ruoucapnhat.xlsx";

            string _path = Path.Combine(Server.MapPath("~/Uploads/ExcelFile"), _FileName);

            file.SaveAs(_path);
            DataTable dt = ReadDataFromExcelFile(_path);

            CopyDataByBulk(dt);
        }
 
        public ActionResult DownloadFile()
        {
          
            string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/ExcelFile/";
          
            byte[] fileBytes = System.IO.File.ReadAllBytes(path + "ruoucapnhat.xlsx");
           
            string fileName = "ruoucapnhat.xlsx";
       
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }

        
    private void CopyDataByBulk(DataTable dt)
        {
           
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["RuouDbContext"].ConnectionString);
            SqlBulkCopy bulkcopy = new SqlBulkCopy(con);
            bulkcopy.DestinationTableName = "KhoRuous";
            bulkcopy.ColumnMappings.Add(0, "IDRuou");
            bulkcopy.ColumnMappings.Add(1, "TenRuou");
            bulkcopy.ColumnMappings.Add(3, "NongDo");
            bulkcopy.ColumnMappings.Add(4, "TinhChat");
            bulkcopy.ColumnMappings.Add(5, "SoLuong");
            bulkcopy.ColumnMappings.Add(6, "DonVi");
            bulkcopy.ColumnMappings.Add(7, "TheTich");
            bulkcopy.ColumnMappings.Add(8, "NguonGoc");
            bulkcopy.ColumnMappings.Add(9, "DanhGia");
            con.Open();
            bulkcopy.WriteToServer(dt);
            con.Close();
        }


        public DataTable ReadDataFromExcelFile(string filepath)
        {
            string connectionString = " ";
            string fileExtention = filepath.Substring(filepath.Length - 4).ToLower();
            if (fileExtention.IndexOf("xlsx") == 0)
            {
                connectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + filepath + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO\"";
            }
            else if (fileExtention.IndexOf("xls") == 0)
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filepath + ";Extended Properties=Excel 8.0";
            }
            OleDbConnection oledbConn = new OleDbConnection(connectionString);
            DataTable data = null;
            try
            {
               
                oledbConn.Open();
       
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn);
         
                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;
      
                DataSet ds = new DataSet();

                oleda.Fill(ds);

                data = ds.Tables[0];
            }
            catch
            {
            }
            finally
            {
              oledbConn.Close();
            }
            return data;
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


