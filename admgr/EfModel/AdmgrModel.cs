using admgr.Entity;
using ADmgr.Helper;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace admgr.EfModel
{
    public class AdmgrModel : DbContext
    {
        //连接字符串。
        public AdmgrModel()
            : base("name=DefaultConnection")
        {
            Database.SetInitializer<AdmgrModel>(new NewInstance());
        }

        public virtual DbSet<AdmgrSystem> AdmgrSystems { get; set; }
        public virtual DbSet<EmailData> EmailDatas { get; set; }
        public virtual DbSet<AdUser> AdUsers { get; set; }
        public virtual DbSet<Role> Roles { get; set; }
        public virtual DbSet<User> Users { get; set; }
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {

        }
    }
    public class NewInstance : CreateDatabaseIfNotExists<AdmgrModel>
    {
        protected override void Seed(AdmgrModel admgrModel)
        {
            #region 初始化角色
            Role role01 = new Role();
            role01.RoleName = "ad用户";
            role01.RoleNum = 1;
            Role role02 = new Role();
            role02.RoleName = "ad公司管理员";
            role02.RoleNum = 2;
            Role role03 = new Role();
            role03.RoleName = "exchange管理员";
            role03.RoleNum = 3;
            Role role04 = new Role();
            role04.RoleName = "lync管理员";
            role04.RoleNum = 4;
            Role role05 = new Role();
            role05.RoleName = "admgr管理员";
            role05.RoleNum = 5;
            Role role06 = new Role();
            role06.RoleName = "ad集团管理员";
            role06.RoleNum = 6;
            admgrModel.Roles.Add(role01);
            admgrModel.Roles.Add(role02);
            admgrModel.Roles.Add(role03);
            admgrModel.Roles.Add(role04);
            admgrModel.Roles.Add(role05);
            admgrModel.Roles.Add(role06);
            admgrModel.SaveChanges();
            #endregion

            #region 初始化用户
            User user01 = new User();
            user01.UserName = "admin";
            user01.UserPassword = HashUtility.GetSHA256Hash("123456");
            user01.Role = role05;
            admgrModel.Users.Add(user01);
            admgrModel.SaveChanges();
            #endregion

            #region 初始化AD系统数据
            AdmgrSystem admgrSystem = new AdmgrSystem();
            admgrSystem.AdServerIp = "10.10.2.11";
            admgrSystem.AdDomain = "test.com";
            admgrSystem.AdManagerUserName = "administrator";
            admgrSystem.AdManagerPwd = "P@ssw0rd";
            admgrSystem.SystemAccount = "admin";
            admgrSystem.SystemPwd = HashUtility.GetSHA256Hash("123456");
            admgrSystem.ExchangIp = "win-2ickeabl4tr.test.com";
            admgrSystem.ExchangeUserName = "administrator";
            admgrSystem.ExchangePwd = "P@ssw0rd";
            admgrSystem.LyncIp = "";
            admgrSystem.LyncUserName = "";
            admgrSystem.LyncPwd = "";
            admgrSystem.MsgUrl = "http://222.73.117.158:80/msg/HttpBatchSendSM";
            admgrSystem.MsgAccount = "huajian888";
            admgrSystem.MsgPwd = "Tch427321";
            admgrSystem.AccountRecycleBin = "9999-XXXXX.账号回收站";
            admgrModel.AdmgrSystems.Add(admgrSystem);
            admgrModel.SaveChanges();
            #endregion

        }
    }
}