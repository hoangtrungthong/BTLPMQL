using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BTLPMQL.Models
{
    [Table("NhapKhaus")]
    public class NhapKhau
    {
        [Key]
        public int IDRuou { get; set; }
        public string TenRuou { get; set; }
        public int NongDo { get; set; }
        public string TinhChat { get; set; }
        public int SoLuong { get; set; }
        public string DonVi { get; set; }
        public int TheTich { get; set; }
        public string NguonGoc { get; set; }
        public string NguoiXuat { get; set; }
        public string NguoiNhap { get; set; }
        [AllowHtml]
        public string Note { get; set; }
    }
}