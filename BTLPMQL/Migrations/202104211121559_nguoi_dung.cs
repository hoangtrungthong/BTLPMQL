namespace BTLPMQL.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class nguoi_dung : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.Accounts", "UserName", c => c.String(nullable: false));
            AlterColumn("dbo.Accounts", "Password", c => c.String(nullable: false));
        }
        
        public override void Down()
        {
            AlterColumn("dbo.Accounts", "Password", c => c.String());
            AlterColumn("dbo.Accounts", "UserName", c => c.String());
        }
    }
}
