using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace admgr.Entity
{
    [Table("AdmgrSystem")]
    public class AdmgrSystem
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int AdmgrSystemId { get; set; }
        // Ad主控服务器的IP地址
        public string AdServerIp { get; set; }
        //AD主控的域名
        public string AdDomain { get; set; }
        //AD主域控制器管理员账户
        public string AdManagerUserName { get; set; }
        //AD主域控制器管理员密码
        public string AdManagerPwd { get; set; }
        //Admgr系统的管理员用户名
        public string SystemAccount{ get; set; }
        //Admgr系统的管理员密码
        public string SystemPwd{ get; set; }
        //Exchange IP地址
        public string ExchangIp{ get; set; }
        // Exchange 账户用户名
        public string ExchangeUserName{ get; set; }
        //Exchange 管理员密码
        public string ExchangePwd{ get; set; }
        //Lync IP地址
        public string LyncIp{ get; set; }
        //Lync 管理员账户"
        public string LyncUserName{ get; set; }
        //Lync 管理员密码
        public string LyncPwd{ get; set; }
        // 企业短信通URL
        public string MsgUrl{ get; set; }
        // 企业短信通帐号
        public string MsgAccount{ get; set; }
        // 企业短信通密码
        public string MsgPwd{ get; set; }
        //帐号回收站OU的名称
        public string AccountRecycleBin { get; set; }
    }
}