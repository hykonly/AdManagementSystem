using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace admgr.Entity
{
    [Table("AdUser")]
    public class AdUser
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int AdUserInfoId { get; set; }
        public string Sn { get; set; }
        public string Number { get; set; }
        public string AdName { get; set; }

        public string AdPssword { get; set; }
        //public bool Checkd { get; set; }
        public string Company { get; set; }
        public string Dept { get; set; }
        public string Title { get; set; }

        public string AccountExpires { get; set; }
        public string DisplayName { get; set; }
        public string Office { get; set; }
        public string Officephone { get; set; }
        public string Ext { get; set; }
        public string Homephone { get; set; }
        public string Mobile { get; set; }
        public string Fax { get; set; }
        public string Email { get; set; }
        public string Country { get; set; }
        public string Province { get; set; }
        public string City { get; set; }
        public string Zipcode { get; set; }
        public string Address { get; set; }
        public string Description { get; set; }
        public string DefaultMailSuffix { get; set; }
        public bool Disble { get; set; }
        public int PasswordResetCount { get; set; }
        public string LastResetTime { get; set; }

        public virtual Role Role { get; set; }
    }
}