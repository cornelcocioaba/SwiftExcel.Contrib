# SwiftExcel.Contrib
[SwiftExcel](https://github.com/RomanPavelko/SwiftExcel) extensions


## Features

### Write`IEnumerable<T>`
```c#
Write<T>(IEnumerable<T> data)
```
### Write Object
It determines whether the value is numeric or string and set the `DataType` accordingly.
```c#
excelWriter.Write(object, col, row);
```
### Include in report
To include a property in the report add the `[IncludeInReport]` attribute. If you want to include properties from nested objects, use the `[NestedIncludeInReport]` attribute
#### Order
The order in the report can be defined by setting the `Orde`r option in the `[IncludeInReport]` attribute i.e. `[IncludeInReport(Order = 1)]`
### Column Names
Use `[Display(Name = "")]` or  `[DisplayName(“”)]` to name a column. The property name will be used otherwise.

## Nested Object example
#### Root Model

```c#
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
```

#### Nested Level One Model:

```c#
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
```

#### Nested Level Two Model:

```c#
public class DemoNestedLevelTwo
{
    [DisplayName("Start Date")]
    [IncludeInReport(Order = 6)]
    public DateTime? StartDates { get; set; }

    [IncludeInReport(Order = 7)]
    public long? Salary { get; set; }
}
```

## Usage
```c#
// Create a report that has nested objects with 10000 rows
IEnumerable<DemoExcel> data = demoService.GetData(10000);
using (var ew = new ExcelWriter("C:/temp/test.xlsx"))
{
    ew.Write(data);
}
```

Thanks to [fingers10/ExcelExport]([https://github.com/fingers10/ExcelExport](https://github.com/fingers10/ExcelExport)) for the majority of the features.