namespace BTLPMQL.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class KhoRuou : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.KhoRuous",
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
                        DanhGia = c.String(),
                    })
                .PrimaryKey(t => t.IDRuou);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.KhoRuous");
        }
    }
}
