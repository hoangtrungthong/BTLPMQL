namespace BTLPMQL.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class Danh_Gia_khach_hang : DbMigration
    {
        public override void Up()
        {
            DropPrimaryKey("dbo.Accounts");
            AlterColumn("dbo.Accounts", "UserName", c => c.String(nullable: false, maxLength: 128));
            AlterColumn("dbo.Accounts", "Password", c => c.String());
            AddPrimaryKey("dbo.Accounts", "UserName");
            DropColumn("dbo.Accounts", "IDTaiKhoan");
        }
        
        public override void Down()
        {
            AddColumn("dbo.Accounts", "IDTaiKhoan", c => c.Int(nullable: false, identity: true));
            DropPrimaryKey("dbo.Accounts");
            AlterColumn("dbo.Accounts", "Password", c => c.String(nullable: false));
            AlterColumn("dbo.Accounts", "UserName", c => c.String(nullable: false));
            AddPrimaryKey("dbo.Accounts", "IDTaiKhoan");
        }
    }
}
