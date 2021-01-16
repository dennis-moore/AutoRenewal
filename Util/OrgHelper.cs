using AutoRenewal.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoRenewal.Util
{
    public class OrgHelper
    {
        private Organization organization;
        private IList<Mapping> OrderedMappings = new List<Mapping>();

        public OrgHelper(Organization organization)
        {
            this.organization = organization;
            OrderMappings();
        }

        public IList<Mapping> HasMapping(string text)
        {
            var mappings = new List<Mapping>();
            foreach(var map in OrderedMappings)
            {
                if (map.WordDesignator == text)
                    mappings.Add(map);
            }
            return mappings;
        }

        private void OrderMappings()
        {
            OrderedMappings = organization.Mappings.OrderBy(mapping =>
                Int32.Parse(mapping.WordDesignator.Replace("[", "").Replace("]", "").Split(' ')[1])).ToList();
            foreach(var map in OrderedMappings)
            {
                Debug.WriteLine(map.WordDesignator);
            }
        }

        public static int GetRowValue(string excelCell)
        {
            return Int32.Parse(excelCell.Substring(1)) - 1;
        }

        public static int GetColumnValue(string excelCell)
        {
            var letter = excelCell.Substring(0, 1)[0];
            int index = (int)letter % 32;
            return index - 1;
        }
    }
}
