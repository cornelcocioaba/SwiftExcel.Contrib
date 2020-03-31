using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SwiftExcel.Contrib.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class IncludeInReportAttribute : Attribute
    {
        public int Order { get; set; }
    }
}
