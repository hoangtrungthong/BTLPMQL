using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using BTLPMQL.Models;
namespace BTLPMQL.Controllers
{
    public class AccountController : Controller
    {
        Encryption encry = new Encryption();
        RuouDbContext db = new RuouDbContext();
        // GET: Account
        [HttpGet]
        public ActionResult Register()
        {
            return View();
        }
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult Register(Account acc)
        {
            if (ModelState.IsValid)
            {
                acc.Password = encry.PasswordEncryption(acc.Password);
                db.Accounts.Add(acc);
                db.SaveChanges();
                return RedirectToAction("Login", "Account");
            }
            return View(acc);
        }
        [HttpGet]
        public ActionResult Login()
        {
            return View();
        }
        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        public ActionResult Login(Account acc)
        {
            if (ModelState.IsValid)
            {
                string encryptionpass = encry.PasswordEncryption(acc.Password);
                var model = db.Accounts.Where(m => m.UserName == acc.UserName && m.Password == encryptionpass).ToList().Count();
                if(model == 1)
                {
                    FormsAuthentication.SetAuthCookie(acc.UserName, true);
                    return RedirectToAction("Index" , "KhoRuous");
                    return RedirectToAction("Index", "XuatKhaus");
                    return RedirectToAction("Index", "NhapKhaus");
                    return RedirectToAction("Edit", "DanhGias");
                    return RedirectToAction("Delete", "DanhGias");

                }
                else
                {
                    ModelState.AddModelError(" ", "Thông Tin Đăng Nhập Không Chính Xác");
                }    
            }
            return View(acc);
        }
        public ActionResult SignOut()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Login", "Account");
        }


    }
}