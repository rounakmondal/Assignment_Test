using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Assignment_Test.Models
{
    public class MappingViewModel
    {
        public List<string> ExcelHeaders { get; set; }
        public List<string> DatabaseFields { get; set; }
        public Dictionary<string, string> Mappings { get; set; } = new Dictionary<string, string>();
    }
}