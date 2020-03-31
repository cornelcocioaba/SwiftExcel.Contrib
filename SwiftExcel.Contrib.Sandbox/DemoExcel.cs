using SwiftExcel.Contrib.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SwiftExcel.Contrib.Sandbox
{
    public class DemoExcel
    {
        public int Id { get; set; }

        [IncludeInReport(Order = 1)]
        public string Name { get; set; }

        [IncludeInReport(Order = 2)]
        public string Position { get; set; }

        [Display(Name = "Office")]
        [IncludeInReport(Order = 3)]
        public string Offices { get; set; }

        [NestedIncludeInReport]
        public DemoNestedLevelOne DemoNestedLevelOne { get; set; }
    }

    public class DemoNestedLevelOne
    {
        [IncludeInReport(Order = 4)]
        public short? Experience { get; set; }

        [DisplayName("Extn")]
        [IncludeInReport(Order = 5)]
        public int? Extension { get; set; }

        [NestedIncludeInReport]
        public DemoNestedLevelTwo DemoNestedLevelTwos { get; set; }
    }

    public class DemoNestedLevelTwo
    {
        [DisplayName("Start Date")]
        [IncludeInReport(Order = 6)]
        public DateTime? StartDates { get; set; }

        [IncludeInReport(Order = 7)]
        public long? Salary { get; set; }
    }
}
