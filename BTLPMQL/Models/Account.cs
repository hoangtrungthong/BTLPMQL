using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace BTLPMQL.Models
{
    public class Account
    {
        [Key]
        public string UserName { get; set; }
        public string Password { get; set; }
    }
}