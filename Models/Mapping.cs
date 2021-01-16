using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace AutoRenewal.Models
{
    public class Mapping
    {
        public string SheetName { get; set; }
        public string ExcelCell { get; set; }
        public string WordDesignator { get; set; }
    }
}
