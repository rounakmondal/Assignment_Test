using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Assignment_Test.Models
{
    public class ImportRequest
    {
        public string FileName { get; set; }
        public Dictionary<string, string> Mappings { get; set; }
    }

}