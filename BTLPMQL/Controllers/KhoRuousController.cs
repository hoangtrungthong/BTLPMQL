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
       /* ReadExcel excel = new ReadExcel();*/
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



        public ActionResult DownloadFile()
        {
            //duong dan chua file muon download
            string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/Excels/"; // tu muc chua fiel excel
            //chuyen file sang dang byte
            byte[] fileBytes = System.IO.File.ReadAllBytes(path + "fielexcel.xlsx"); //doan nay de file excel
            //ten file khi download ve
            string fileName = "Sinhviencapnhat.xlsx";
            //tra ve file
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
        }
        public ActionResult UploadFile(HttpPostedFileBase file)
        {
            //dat ten cho file
            string _FileName = "ruoumoi.xls";
            //duong dan luu file
            string _path = Path.Combine(Server.MapPath("~/Uploads/Excels"), _FileName);
            //luu file len server
            file.SaveAs(_path);

            //doc du lieu tu file excel

            DataTable dt = ReadDataFromExcelFile(_path);
            /* CopyDataByBulk(dt);*/
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                KhoRuou kr = new KhoRuou();
                kr.TenRuou = dt.Rows[i][0].ToString();
                kr.NongDo = dt.Rows[i][1].ToString();
                kr.TinhChat = dt.Rows[i][2].ToString();
                kr.SoLuong = dt.Rows[i][3].ToString();
                kr.DonVi = dt.Rows[i][4].ToString();
                kr.TheTich = dt.Rows[i][5].ToString();
                kr.NguonGoc = dt.Rows[i][6].ToString();
                kr.DanhGia = dt.Rows[i][7].ToString();
               
                db.KhoRuous.Add(kr);
                db.SaveChanges();
            }

            /* CopyDataByBulk(excel.ReadDataFromExcelFile(_path));*/
            return RedirectToAction("Index");
        }

        private void CopyDataByBulk(object v)
        {
            throw new NotImplementedException();
        }

        //upload file
        public DataTable ReadDataFromExcelFile(string filepath)
        {
            string connectionString = "";
            string fileExtention = filepath.Substring(filepath.Length - 4).ToLower();
            if (fileExtention.IndexOf("xlsx") == 0)
            {
                connectionString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source =" + filepath + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO\"";
            }
            else if (fileExtention.IndexOf(".xls") == 0)
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filepath + ";Extended Properties=Excel 8.0";
            }

            // Tạo đối tượng kết nối
            OleDbConnection oledbConn = new OleDbConnection(connectionString);
            DataTable data = null;
            try
            {
                // Mở kết nối
                oledbConn.Open();

                // Tạo đối tượng OleDBCommand và query data từ sheet có tên "Sheet1"
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn);

                // Tạo đối tượng OleDbDataAdapter để thực thi việc query lấy dữ liệu từ tập tin excel
                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                // Tạo đối tượng DataSet để hứng dữ liệu từ tập tin excel
                DataSet ds = new DataSet();

                // Đổ đữ liệu từ tập excel vào DataSet
                oleda.Fill(ds);

                data = ds.Tables[0];
            }
            catch
            {
            }
            finally
            {
                // Đóng chuỗi kết nối
                oledbConn.Close();
            }
            return data;
        }

        private void CopyDataByBulk(DataTable dt)
        {
            //lay ket noi voi database luu trong file webconfig
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["RuouDbContext"].ConnectionString);
            SqlBulkCopy bulkcopy = new SqlBulkCopy(con);
            bulkcopy.DestinationTableName = "KhoRuous";
            bulkcopy.ColumnMappings.Add(0, "Tên Rượu");
            bulkcopy.ColumnMappings.Add(1, "Nồng Độ");
            bulkcopy.ColumnMappings.Add(2, "Tính Chất");
            bulkcopy.ColumnMappings.Add(3, " Số Lượng");
            bulkcopy.ColumnMappings.Add(4, " Đơn Vị");
            bulkcopy.ColumnMappings.Add(5, "Thể Tích");
            bulkcopy.ColumnMappings.Add(6, "Nguồn Gốc");
            bulkcopy.ColumnMappings.Add(7, "Đánh Giá");
            con.Open();
            bulkcopy.WriteToServer(dt);
            con.Close();
        }
        /* public ActionResult UploadFile (HttpPostedFileBase file)
         {
             try
             {
                 if (file.ContentLength > 0)
                 {

                     string _FileName = DateTime.Now.Year.ToString() + DateTime.Now.Date.ToString();

                     string _path = Path.Combine(Server.MapPath("~/Uploads/ExcelFile/"), _FileName);

                     file.SaveAs(_path);

                     CopyDataByBulk(ReadDataFromExcelFile(_path));
                     ViewBag.ThongBao = "cap nhat thanh cong";
                 }
             }
             catch (Exception ex)
             {
                 ViewBag.ThongBao = "cap nhat thanh cong";
             }
             return RedirectToAction("Index");
         }*/

        //copy large data from datatable to sqlserver



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


