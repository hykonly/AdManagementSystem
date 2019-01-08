using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace admgr.Entity
{
    [Table("EmailData")]
    public class EmailData
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int EmailDataId { get; set; }
        public string CompanyName { get; set; }
        public string MainMailAddress { get; set; }
        public string SecondMailAddress { set; get; }
        public string EmailStoreDb { get; set; }
    }
}