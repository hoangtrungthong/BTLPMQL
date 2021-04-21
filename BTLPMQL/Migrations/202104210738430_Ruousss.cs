namespace BTLPMQL.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class Ruousss : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.DanhGias",
                c => new
                    {
                        IDKhachHang = c.Int(nullable: false, identity: true),
                        TenKhachHang = c.String(),
                        NhanXet = c.String(),
                    })
                .PrimaryKey(t => t.IDKhachHang);
            
            CreateTable(
                "dbo.NhapKhaus",
                c => new
                    {
                        IDRuou = c.Int(nullable: false, identity: true),
                        TenRuou = c.String(),
                        NongDo = c.Int(nullable: false),
                        TinhChat = c.String(),
                        SoLuong = c.Int(nullable: false),
                        DonVi = c.String(),
                        TheTich = c.Int(nullable: false),
                        NguonGoc = c.String(),
                        NguoiXuat = c.String(),
                        NguoiNhap = c.String(),
                        Note = c.String(),
                    })
                .PrimaryKey(t => t.IDRuou);
            
            CreateTable(
                "dbo.XuatKhaus",
                c => new
                    {
                        IDRuou = c.Int(nullable: false, identity: true),
                        TenRuou = c.String(),
                        NongDo = c.Int(nullable: false),
                        TinhChat = c.String(),
                        SoLuong = c.Int(nullable: false),
                        DonVi = c.String(),
                        TheTich = c.Int(nullable: false),
                        NguonGoc = c.String(),
                        NguoiXuat = c.String(),
                        NguoiNhap = c.String(),
                        Note = c.String(),
                    })
                .PrimaryKey(t => t.IDRuou);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.XuatKhaus");
            DropTable("dbo.NhapKhaus");
            DropTable("dbo.DanhGias");
        }
    }
}
