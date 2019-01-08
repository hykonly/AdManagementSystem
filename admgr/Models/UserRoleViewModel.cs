using admgr.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace admgr.Models
{
    public class TitleViewModel
    {
        public string TitleId { get; set; }
        public string TitleName { get; set; }
    }
    public class UserRoleViewModel
    {
        public string Id { get; set; }
        public string AdUserId { get; set; }
        public string RoleId { get; set; }
    }
    public class AdUserViewModel
    {
        public string AdUserId { get; set; }
        public string AdUserName { get; set; }
    }
    public class AdUserPartViewModel
    {
        public string AdUserInfoId { get; set; }
        public string Sn { get; set; }
        public string Number { get; set; }
        public string AdName { get; set; }
        //public string AdPssword { get; set; }
        public string CompanyId { get; set; }
        public string DeptId { get; set; }
        public string TitleId { get; set; }
        //public string AccountExpires { get; set; }
        //public string DisplayName { get; set; }
        //public string Office { get; set; }
        //public string Officephone { get; set; }
        //public string Ext { get; set; }
        //public string Homephone { get; set; }
        public string Mobile { get; set; }
        //public string Fax { get; set; }
        public string Email { get; set; }
        //public string Country { get; set; }
        //public string Province { get; set; }
        //public string City { get; set; }
        //public string Zipcode { get; set; }
        //public string Address { get; set; }
        //public string Description { get; set; }
        //public string DefaultMailSuffix { get; set; }
        public string DisbleId { get; set; }
        //public int PasswordResetCount { get; set; }
        //public string LastResetTime { get; set; }
        //public string RoleName { get; set; }

    }
}