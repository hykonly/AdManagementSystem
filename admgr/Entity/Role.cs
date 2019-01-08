using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace admgr.Entity
{
    [Table("Role")]
    public class Role
    {
        public Role()
        {
            AdUsers = new HashSet<AdUser>();
            Users = new HashSet<User>();
        }
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int RoleId { get; set; }
        public string RoleName { get; set; }
        public int RoleNum { get; set; }
        public virtual ICollection<AdUser> AdUsers { get; set; }
        public virtual ICollection<User> Users { get; set; }
    }
}