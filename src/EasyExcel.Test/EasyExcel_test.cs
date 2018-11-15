using System.Collections.Generic;
using Xunit;
using EasyExcel;
using NPOI.SS.UserModel;
using System.IO;
using System;

namespace EasyExcel.Test {
    public class EasyExcel_Test {
        public class TestClass {
            public string name;

            public int id;

            public decimal amount;
        }

        [Fact]
        public void Test_CanGenerate () {
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
            var excel = new Excel (ExcelType.XSSF);
            var book1 = excel.AddSheetbook ("testSheet", testData)
                .HasFreezeTitle (false);

            book1.HasColumn ("固定值列")
                .WithValue ("hhh")
                .WithCellType (CellType.String);

            book1.HasColumn ("id列")
                .WithValue (i => i.id)
                .WithFormat ("0");

            book1.HasColumn ("计算列")
                .WithValue (i => $"name: {i.name}, amount: {i.amount}")
                .WithWidth (20);

            book1.HasColumn ("金额列")
                .WithValue (i => i.amount);

            book1.HasColumn ("名称列")
                .WithValue (i => i.name)
                .WithIndex (1);

            var d = excel.Build ();

            using (Stream file = new FileStream (Path.Combine (Environment.CurrentDirectory, "test.xlsx"), FileMode.Create, FileAccess.Write)) {
                excel.Build ().Write (file);
            }
        }
    }
}