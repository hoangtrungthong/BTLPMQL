using System;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;
using System.Linq;

namespace BTLPMQL.Models
{
    public partial class RuouDbContext : DbContext
    {
        public RuouDbContext()
            : base("name=RuouDbContext")
        {
        }

        public virtual DbSet<KhoRuou> KhoRuous { get; set; }
        public virtual DbSet<Account> Accounts { get; set; }
      
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
        }

        public System.Data.Entity.DbSet<BTLPMQL.Models.XuatKhau> XuatKhaus { get; set; }

        public System.Data.Entity.DbSet<BTLPMQL.Models.NhapKhau> NhapKhaus { get; set; }

        public System.Data.Entity.DbSet<BTLPMQL.Models.DanhGia> DanhGias { get; set; }
    }
}
