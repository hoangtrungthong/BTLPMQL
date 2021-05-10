namespace BTLPMQL.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class khoruoumoi : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.KhoRuous", "NongDo", c => c.String());
            AlterColumn("dbo.KhoRuous", "SoLuong", c => c.String());
            AlterColumn("dbo.KhoRuous", "TheTich", c => c.String());
        }
        
        public override void Down()
        {
            AlterColumn("dbo.KhoRuous", "TheTich", c => c.Int(nullable: false));
            AlterColumn("dbo.KhoRuous", "SoLuong", c => c.Int(nullable: false));
            AlterColumn("dbo.KhoRuous", "NongDo", c => c.Int(nullable: false));
        }
    }
}
