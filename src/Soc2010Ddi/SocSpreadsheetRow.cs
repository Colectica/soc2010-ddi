using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Soc2010Ddi
{
    public class SocSpreadsheetRow
    {
        public string MajorGroup { get; set; }
        public string SubMajorGroup { get; set; }
        public string MinorGroup { get; set; }
        public string UnitGroup { get; set; }
        public string GroupTitle { get; set; }

        public bool IsEmpty
        {
            get
            {
                return !IsMajor && !IsSubMajor && !IsMinor && !IsUnit;
            }
        }

        public bool IsMajor
        {
            get
            {
                return !string.IsNullOrWhiteSpace(MajorGroup);
            }
        }

        public bool IsSubMajor
        {
            get
            {
                return !string.IsNullOrWhiteSpace(SubMajorGroup);
            }
        }

        public bool IsMinor
        {
            get
            {
                return !string.IsNullOrWhiteSpace(MinorGroup);
            }
        }

        public bool IsUnit
        {
            get
            {
                return !string.IsNullOrWhiteSpace(UnitGroup);
            }
        }
    }
}
