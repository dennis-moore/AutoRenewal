using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoRenewal.Models
{
    public class Organization
    {
        public string Name { get; set; }
        public string TemplateFileName { get; set; }
        //public IList<string> ExcelTabNames { get; set; } = new List<string>();
        public ObservableCollection<Mapping> Mappings { get; set; } = new ObservableCollection<Mapping>();
    }
}
