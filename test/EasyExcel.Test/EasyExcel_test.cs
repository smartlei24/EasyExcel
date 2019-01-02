using System.Collections.Generic;
using Xunit;
using EasyExcel;
using NPOI.SS.UserModel;
using System.IO;
using System;

namespace EasyExcel.Test
{
    public class EasyExcel_Test
    {
        public class TestClass 
        {
            public string name;

            public int id;

            public decimal amount;
        }

        [Fact]
        public void Test_CanGenerate () 
        {
            var testData = new List<TestClass> {
                new TestClass { name = "a", id = 1, amount = 12.20M },
                new TestClass { name = "b", id = 2, amount = 2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M },
                new TestClass { name = "c", id = 5, amount = 10.2M }
            };
            var excel = new Workbook(ExcelType.XSSF);
            var book1 = excel.AddSheet<TestClass>("testSheet")
                .IsFreezeTitle (false);

            book1.AddColumn ("固定值列")
                .Value ("hhh")
                .HasCellType (CellType.String);

            book1.AddColumn("id列")
                .Value (i => i.id)
                .HasFormat("0");

            book1.AddColumn("计算列")
                .Value (i => $"name: {i.name}, amount: {i.amount}")
                .HasWidth (20);

            book1.AddColumn("金额列")
                .Value (i => i.amount);

            book1.AddColumn("名称列")
                .Value (i => i.name)
                .HasIndex (1);

            book1.Fill(testData);

            var d = excel.Build();

            using (Stream file = new FileStream (Path.Combine (Environment.CurrentDirectory, "test.xlsx"), FileMode.Create, FileAccess.Write)) {
                excel.Build().Write (file);
            }
        }
    }
}