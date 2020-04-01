using System;
using System.Collections.Generic;

namespace SwiftExcel.Contrib.Sandbox
{
    class Program
    {
        private static void Main()
        {
            var demoService = new DemoService();
            IEnumerable<DemoExcel> data = demoService.GetData(10000);

            using (var ew = new ExcelWriter("C:/temp/test.xlsx"))
            {
                ew.Write(data);
            }
        }
    }
}
