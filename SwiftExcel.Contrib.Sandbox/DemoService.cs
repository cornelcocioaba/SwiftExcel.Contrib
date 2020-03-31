using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SwiftExcel.Contrib.Sandbox
{
    public class DemoService
    {
        public IEnumerable<DemoExcel> GetData(int size)
        {
            var list = new List<DemoExcel>(size);

            for (int i = 0; i < size; i++)
            {
                list.Add(GetDemoExcel());
            }

            return list;
        }

        private DemoExcel GetDemoExcel()
        {
            var rnd = new Random();
            var next = rnd.Next(1, 100);

            var nested2 = new DemoNestedLevelTwo
            {
                Salary = next,
                StartDates = DateTime.Now
            };

            var nested1 = new DemoNestedLevelOne
            {
                DemoNestedLevelTwos = nested2,
                Experience = 10,
                Extension = next
            };

            var excel = new DemoExcel
            {
                Id = next,
                Name = "name " + next,
                Offices = "offices " + next,
                Position = "position " + next,
                DemoNestedLevelOne = nested1
            };

            return excel;
        }
    }
}
