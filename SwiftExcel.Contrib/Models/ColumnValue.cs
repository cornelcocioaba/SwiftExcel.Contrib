using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SwiftExcel.Contrib.Models
{
    public class ColumnValue
    {
        public int Order { get; set; }
        public string Path { get; set; }
        public PropertyDescriptor PropertyDescriptor { get; set; }
    }
}
