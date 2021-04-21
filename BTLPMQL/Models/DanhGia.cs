using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BTLPMQL.Models
{
    [Table("DanhGias")]
    public class DanhGia
    {
        [Key]
        public int IDKhachHang { get; set; }
        public string TenKhachHang { get; set; }
        [AllowHtml]
        public string NhanXet { get; set; }
    }
}