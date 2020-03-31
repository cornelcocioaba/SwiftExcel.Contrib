using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SwiftExcel;
using SwiftExcel.Contrib.Extensions;

namespace SwiftExcel.Contrib
{
    public static class ExcelWriterExtensions
    {
        public static void Write(this ExcelWriter excelWriter, object value, int col, int row)
        {
            if (value.IsNumber())
            {
                excelWriter.Write(value.ToString(), col, row, DataType.Number);
            }
            else
            {
                excelWriter.Write(value.ToString(), col, row, DataType.Text);
            }
        }

        public static void Write<T>(this ExcelWriter excelWriter, IEnumerable<T> data)
        {
            var columns = PropertyExtensions.GetColumnsFromModel(typeof(T)).OrderBy(x => x.Value.Order);

            int col = 1;
            int row = 1;
            foreach (var column in columns)
            {
                excelWriter.Write(column.Name, col, row);
                col++;
            }

            col = 1;
            row = 2;
            foreach (var item in data)
            {
                foreach (var column in columns)
                {
                    var value = PropertyExtensions.GetPropertyValue(item, column.Value.Path) ?? DBNull.Value;
                    excelWriter.Write(value, col, row);
                    col++;
                }

                row++;
                col = 1;
            }
        }
    }
}
