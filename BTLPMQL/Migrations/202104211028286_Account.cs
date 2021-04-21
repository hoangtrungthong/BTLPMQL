namespace BTLPMQL.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class Account : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Accounts",
                c => new
                    {
                        IDTaiKhoan = c.Int(nullable: false, identity: true),
                        UserName = c.String(),
                        Password = c.String(),
                    })
                .PrimaryKey(t => t.IDTaiKhoan);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Accounts");
        }
    }
}
