# EasyExcil

一个基于 DotnetCore.NPOI 的 快捷生成 Excel 的工具, 封装了默认样式

[TOC]

## 基本用法

```csharp
var workbook = new Workbook();

var sheet1 = workbook.AddSheet("OrderList")
    .IsFreezeTitle(false)
    .IsHasBorder(false)
    .IsStripedRow(false);

sheet1.AddColumn("Order Number")
    .Value(i => i.Id)
    .HasWidth(3);

sheet1.AddColumn("Buyer")
    .Value(i => i.BuyerName)
    .HasWidth(15);

sheet1.Fill(new List<Class>())

var workbook = workbook.Build();
```



## 具体用法



### 1. Workbook

`Workbook` 类构造函数接受一个枚举类型的参数：`ExcelType`, 默认为XSSF ， 如下：

| 值    | 对应 NPOI 类型 | 描述                                                         |
| ----- | -------------- | ------------------------------------------------------------ |
| HSSF  | HSSFWorkbook   | 适用于 Excel2003 以前的版本，后缀名为 .xls                   |
| XSSF  | XSSFWorkbook   | 适用于 Excel 2007 之后的版本，后缀名为 .xlsx                 |
| SXSSF | SXSSFWorkbook  | XSSF 的低内存占用版本，可解决其他两者数据量超出65536条后内存溢出的问题 |

通过`workbook.Build()` 方法即可返回一个构建好的 NPOI 中的 `Iworkbook`对象。



## 2. Sheet

使用 `Workbook.AddSheet<T>(string name)` 方法即可增加一个 Sheet，该方法创建一个泛型类 `Sheet<T>` 。

``` csharp
var sheet1 = workbook.AddSheet<TestClass>("OrderList")
```

### Methods

| Name                                                   | Description                                                  |
| ------------------------------------------------------ | ------------------------------------------------------------ |
| Column<T> AddColumn(string columnName)                 | 新增列，传入该列显示的名字                                   |
| Sheet<T> IsFreezeTitle(bool freeze = true)        | 是否固定表头，默认为`true`， 固定后生成的Excel 表头始终固定在第一行，不随上下滚动移动。 |
| Sheet<T> IsHasBorder(bool outBolder = true)       | 是否含有外边框，默认为true， 设定后 Excel 表格的有内容的区域会存在一层边框 |
| Sheet<T> IsStripedRow(bool has = true)            | 是否含有行间隔条纹                                           |
| Sheet<T> HasTitleStyle(Action<ICellStyle> action) | 设定表头单元格样式，一旦使用，默认样式失效！*非必须不推荐使用。* |
| Sheet<T> HasTitleRowHeight(float height = 30f)    | 设定标题行高                                                 |
| Sheet<T> HasBodyRowHeight(float height = 20f)     | 设定内容行高                                                 |

### Example

```csharp
// 仅为demo， 非特殊情况，仅需指定列名和数据
var sheet1 = workbook.AddSheet<Class>("OrderList")
    .IsFreezeTitle(false)
    .IsHasBorder(false)
    .IsStripedRow(false)
    .HasTitleStyle(style => {
        style.FillForegroundColor = Red.Index;
        style.FillPattern = FillPattern.SolidForeground;
    })
    .HasBodyRowHeight(19f)
	.Fill(new List<Class>());

sheet1.AddColumn("TestCol1");
```



## 3. Column

使用`sheet1.AddColumn("col1")` 为该 Sheet 新增一列。

### Methods

| Name                                                         | Descriptio                                                   |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| Column<T> Value(TResult value)                           | 使用固定值填充该列                                           |
| Column<T> Value<TResult>(Expression<Func<T, TResult>> func) | 使用计算值填充该列，传入一个计算值的表达式                   |
| Column<T> HasWidth(float width)                               | 指定列宽，（以字符宽度为单位），如该列大致9个字符长，则可传入 9 |
| Column<T> HasStyle(Action<ICellStyle> action)           | 指定设定该列的内容单元格样式，一旦使用，默认样式失效！*非必要不推荐使用。* |
| Column<T> HasIndex(uint index)                              | 指定在表格中该列的位置，以0起始。不指定时，则按照声明顺序排列，再将已指定的列插入到指定的位置。 |
| Column<T> HasCellType(CellType cellType)                    | 指定该列的类型，详见 [POI](https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/CellType.html) 文档， 默认为`CellType.String` |
| Column<T> HasFormat(string formatString)                   | 指定该列的格式化字符串， 若不指定则会根据数据类型进行默认格式化（如 `decimal` 则会以`$#,##0.00` 进行格式化）。*非必要不推荐使用。* |

### Example

```csharp
// 仅为举例， 非特殊情况，可仅指定列名、值、列宽和Index即可
sheet1.Column ("Test AutoStyle")
    .Value(i => i.Cost * i.Quantity)
    .HasWidth(3)
    .HasIndex(2)
    .HasStyle(style => {
         style.FillForegroundColor = Red.Index;
         style.FillPattern = FillPattern.SolidForeground;
     })
    .HasCellType(CellType.Formula)
    .HasFormat("00");
```
