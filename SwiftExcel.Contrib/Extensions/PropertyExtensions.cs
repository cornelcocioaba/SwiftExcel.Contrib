using SwiftExcel.Contrib.Attributes;
using SwiftExcel.Contrib.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SwiftExcel.Contrib.Extensions
{
    public static class PropertyExtensions
    {
        public static PropertyDescriptor GetPropertyDescriptor(this PropertyInfo propertyInfo)
        {
            return TypeDescriptor.GetProperties(propertyInfo.DeclaringType).Find(propertyInfo.Name, false);
        }

        public static string GetPropertyDisplayName(this PropertyInfo propertyInfo)
        {
            var propertyDescriptor = propertyInfo.GetPropertyDescriptor();
            var displayName = propertyInfo.IsDefined(typeof(DisplayAttribute), false) ? propertyInfo.GetCustomAttributes(typeof(DisplayAttribute),
                false).Cast<DisplayAttribute>().Single().Name : null;

            return displayName ?? propertyDescriptor.DisplayName ?? propertyDescriptor.Name;
        }

        public static object GetPropertyValue(object src, string propName)
        {
            // https://stackoverflow.com/questions/1954746/using-reflection-in-c-sharp-to-get-properties-of-a-nested-object/29823444#29823444
            if (src == null) throw new ArgumentException("Value cannot be null.", "src");
            if (propName == null) throw new ArgumentException("Value cannot be null.", "propName");

            if (propName.Contains("."))//complex type nested
            {
                var temp = propName.Split(new char[] { '.' }, 2);
                return GetPropertyValue(GetPropertyValue(src, temp[0]), temp[1]);
            }
            else
            {
                var prop = src.GetType().GetProperty(propName);
                return prop != null ? prop.GetValue(src, null) : null;
            }
        }

        public static IEnumerable<Column> GetColumnsFromModel(Type parentClass, string parentName = null)
        {
            var complexReportProperties = parentClass.GetProperties()
                       .Where(p => p.GetCustomAttributes<NestedIncludeInReportAttribute>().Any());

            var properties = parentClass.GetProperties()
                       .Where(p => p.GetCustomAttributes<IncludeInReportAttribute>().Any());

            foreach (var prop in properties.Except(complexReportProperties))
            {
                var attribute = prop.GetCustomAttribute<IncludeInReportAttribute>();

                yield return new Column
                {
                    Name = prop.GetPropertyDisplayName(),
                    Value = new ColumnValue
                    {
                        Order = attribute.Order,
                        Path = string.IsNullOrWhiteSpace(parentName) ? prop.Name : $"{parentName}.{prop.Name}",
                        PropertyDescriptor = prop.GetPropertyDescriptor()
                    }
                };
            }

            if (complexReportProperties.Any())
            {
                foreach (var parentProperty in complexReportProperties)
                {
                    var parentType = parentProperty.PropertyType;
                    var parentAttribute = parentProperty.GetCustomAttribute<NestedIncludeInReportAttribute>();

                    var complexProperties = GetColumnsFromModel(parentType, string.IsNullOrWhiteSpace(parentName) ? parentProperty.Name : $"{parentName}.{parentProperty.Name}");

                    foreach (var complexProperty in complexProperties)
                    {
                        yield return complexProperty;
                    }
                }
            }
        }

    }
}
