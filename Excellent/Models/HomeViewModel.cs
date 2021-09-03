using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace Excellent.Models
{
    public class HomeViewModel
    {
        public string FileName { get; set; }
        public string SheetName { get; set; }
        public bool IsSuccessfull { get; set; }
        public string ErrorMessage { get; set; }
        public DataTable Data { get; set; }
    }
}