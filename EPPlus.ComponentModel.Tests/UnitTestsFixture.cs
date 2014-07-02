// --------------------------------------------------------------------------------------------------------------------
// <copyright file="UnitTestsFixture.cs" company="Patrick Magee">
//   The MIT License (MIT)
//   
//   Copyright (c) 2014 Patrick Magee
//   
//   Permission is hereby granted, free of charge, to any person obtaining a copy
//   of this software and associated documentation files (the "Software"), to deal
//   in the Software without restriction, including without limitation the rights
//   to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//   copies of the Software, and to permit persons to whom the Software is
//   furnished to do so, subject to the following conditions:
//   
//   The above copyright notice and this permission notice shall be included in all
//   copies or substantial portions of the Software.
//   
//   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//   FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//   AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//   LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//   OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//   SOFTWARE.
// </copyright>
// <summary>
//   The unit test 1.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Tests
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using EPPlus.ComponentModel.Exceptions;
    using EPPlus.ComponentModel.Export;
    using EPPlus.ComponentModel.Import;
    using EPPlus.ComponentModel.Tests.Entities;

    using FizzWare.NBuilder;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using MSTestExtensions;

    using OfficeOpenXml;
    using OfficeOpenXml.DataValidation;
    using OfficeOpenXml.DataValidation.Contracts;
    using OfficeOpenXml.DataValidation.Formulas.Contracts;

    /// <summary>
    /// Exporting and Importing Tests Fixture
    /// </summary>
    [TestClass]
    public class UnitTestsFixture
    {
        #region Test resources

        /// <summary>
        /// The export service.
        /// </summary>
        private ExportService exportService;

        /// <summary>
        /// The import service.
        /// </summary>
        private ImportService importService;

        #endregion

        [TestInitialize]
        [Description("Before each test runs")]
        public void Initialize()
        {
            exportService = new ExportService();
        }

        [TestMethod]
        [Description("Exporting a sheet without a name results in an argument null exception")]
        [TestCategory("Guard")]
        public void Exporting_SheetName_Required_Test()
        {
            // Arrange
            string sheetName = null;

            // Act
            var action = new Action(() => exportService.AddSheetForExport(sheetName));

            // Assert
            ExceptionAssert.Throws<ArgumentNullException>(action);
        }

        [TestMethod]
        [Description("Exporting sheet names must be unique")]
        [TestCategory("Guard")]
        public void Exporting_Sheet_Names_Must_Be_Unique_Test()
        {
            // Arrange
            string sheetName = "Sheet";

            // Act
            exportService.AddSheetForExport(sheetName);
            var action = new Action(() => exportService.AddSheetForExport(sheetName));

            // Assert
            ExceptionAssert.Throws<SheetNameExistsException>(action);
        }

        [TestMethod]
        [Description("Exporting table names must be unique")]
        [TestCategory("Guard")]
        public void Exporting_Table_Names_Must_Be_Unique_Test()
        {
            // Arrange
            string tableName = "Table";
            string sheetName = "Sheet";
            var orders = Builder<Order>.CreateListOfSize(10).Build();
            var sheet = exportService.AddSheetForExport(sheetName);

            // Act
            var table1 = sheet.AddTableForExport(orders, tableName);
            var table2 = sheet.AddTableForExport(orders, tableName);

            Console.WriteLine("Table 1: {0}. Table 2: {1}", table1.TableName, table2.TableName);

            // Assert
            Assert.IsFalse(table1.TableName == table2.TableName);

            this.SaveToFile(exportService.Export());
        }

        [TestMethod]
        [Description("Exporting a table without a name results in the default name for that type")]
        [TestCategory("1 x Sheet"), TestCategory("1 x Type"), TestCategory("1 x Table")]
        public void Exporting_TableName_Test()
        {
            // Arrange
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            var orders = Builder<Order>.CreateListOfSize(10).Build();

            // Act
            var table1 = sheet1.AddTableForExport(orders);

            // Assert
            Assert.AreEqual(
                expected: "Sheet_One_Orders_1",
                actual: table1.TableName);
        }

        [TestMethod]
        [Description("Exporting tables without a name results in the default name for that table type")]
        [TestCategory("1 x Sheet"), TestCategory("1 x Type"), TestCategory("2 x Table")]
        public void Exporting_TableName_Unique_Test()
        {
            // Arrange
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            var orders = Builder<Order>.CreateListOfSize(10).Build();

            // Act
            var table1 = sheet1.AddTableForExport(orders);
            var table2 = sheet1.AddTableForExport(orders);

            // Assert
            Assert.AreEqual(
                expected: "Sheet_One_Orders_1",
                actual: table1.TableName);

            Assert.AreEqual(
                expected: "Sheet_One_Orders_2",
                actual: table2.TableName);
        }

        [TestMethod]
        [Description("Exporting tables without a name results in the default name for that table type")]
        [TestCategory("1 x Sheet"), TestCategory("1 x Type"), TestCategory("2 x Table")]
        public void Exporting_Multiple_Sheets_TableName_Unique_Test()
        {
            // Arrange
            var ordersToExport = Builder<Order>.CreateListOfSize(10).Build();
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            var sheet2 = exportService.AddSheetForExport("Sheet Two");

            // Act
            var sheet1Table1 = sheet1.AddTableForExport(ordersToExport);
            var sheet1Table2 = sheet1.AddTableForExport(ordersToExport);

            var sheet2Table1 = sheet2.AddTableForExport(ordersToExport);
            var sheet2Table2 = sheet2.AddTableForExport(ordersToExport);

            // Assert
            Assert.AreEqual(
                expected: "Sheet_One_Orders_1",
                actual: sheet1Table1.TableName);

            Assert.AreEqual(
                expected: "Sheet_One_Orders_2",
                actual: sheet1Table2.TableName);

            Assert.AreEqual(
                expected: "Sheet_Two_Orders_1",
                actual: sheet2Table1.TableName);

            Assert.AreEqual(
                expected: "Sheet_Two_Orders_2",
                actual: sheet2Table2.TableName);
        }

        [TestMethod]
        [Description("Exporting all orders results in importing all orders.")]
        [TestCategory("1 x Sheet"), TestCategory("1 x Table"), TestCategory("1 x Type")]
        public void Exporting_One_Table_Type_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> ordersToExport = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> allOrders = Enumerable.Empty<Order>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(ordersToExport);
            data = exportService.Export();

            importService = new ImportService(data);
            allOrders = importService.GetAll<Order>();

            // Assert
            Assert.AreEqual(
                expected: ordersToExport.Count(),
                actual: allOrders.Count());
        }

        [TestMethod]
        [Description("Exporting orders to two different tables results in importing all orders.")]
        [TestCategory("1 x Sheet"), TestCategory("2 x Table"), TestCategory("1 x Type")]
        public void Exporting_Two_Tables_Same_Type_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> firstOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> secondOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> allOrders = Enumerable.Empty<Order>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(firstOrders);
            sheet1.AddTableForExport(secondOrders);
            data = exportService.Export();

            importService = new ImportService(data);
            allOrders = importService.GetAll<Order>();

            // Assert
            var totalOrders = firstOrders.Concat(secondOrders);

            Assert.AreEqual(
                expected: totalOrders.Count(),
                actual: allOrders.Count());

        }

        [TestMethod]
        [Description("Exporting two types results in importing both of all types")]
        [TestCategory("2 x Type"), TestCategory("3 x Table"), TestCategory("1 x Sheet")]
        public void Exporting_Three_Tables_Two_Types_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> firstOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> secondOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> allOrders = Enumerable.Empty<Order>();

            IEnumerable<Person> firstPeople = Builder<Person>.CreateListOfSize(10).Build();
            IEnumerable<Person> allPeople = Enumerable.Empty<Person>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(firstOrders);
            sheet1.AddTableForExport(secondOrders);
            sheet1.AddTableForExport(firstPeople);
            data = exportService.Export();

            importService = new ImportService(data);
            allOrders = importService.GetAll<Order>();
            allPeople = importService.GetAll<Person>();

            // Assert
            var totalOrders = firstOrders.Concat(secondOrders);

            Assert.AreEqual(
                expected: totalOrders.Count(),
                actual: allOrders.Count(),
                message: "The total orders exported do not add up to the total orders imported.");

            Assert.AreEqual(
                expected: firstPeople.Count(),
                actual: allPeople.Count(),
                message: "The total people exported do not add up to the total people imported.");
        }

        [TestMethod]
        [Description("Exporting one type in three sheets results in importing all")]
        [TestCategory("1 x Type"), TestCategory("3 x Table"), TestCategory("3 x Sheet")]
        public void Exporting_Three_Sheets_One_Type_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> firstOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> secondOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> thirdOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> allOrders = Enumerable.Empty<Order>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(firstOrders);

            var sheet2 = exportService.AddSheetForExport("Sheet Two");
            sheet2.AddTableForExport(secondOrders);

            var sheet3 = exportService.AddSheetForExport("Sheet Three");
            sheet3.AddTableForExport(thirdOrders);

            data = exportService.Export();

            importService = new ImportService(data);
            allOrders = importService.GetAll<Order>();

            // Assert
            var totalOrders = firstOrders.Concat(secondOrders).Concat(thirdOrders);

            Assert.AreEqual(
                expected: totalOrders.Count(),
                actual: allOrders.Count(),
                message: "The total orders exported to three sheets do not add up to the total orders imported.");

            this.SaveToFile(data);
        }

        [TestMethod]
        [Description("Exporting one type in three sheets results in importing all")]
        [TestCategory("3 x Type"), TestCategory("3 x Table"), TestCategory("3 x Sheet")]
        public void Exporting_Three_Tables_Per_Three_Types_Per_Three_Sheets_Test()
        {
            // Arrange
            byte[] data;
            var firstOrders = Builder<Order>.CreateListOfSize(10).All().Do(order => order.Reference = "FIRST ORDER").Build();
            var secondOrders = Builder<Order>.CreateListOfSize(10).All().Do(order => order.Reference = "SECOND ORDER").Build();
            var thirdOrders = Builder<Order>.CreateListOfSize(10).All().Do(order => order.Reference = "THIRD ORDER").Build();

            var firstCars = Builder<Car>.CreateListOfSize(10).All().Do(car => car.Model = "FIRST MODEL").Build();
            var secondCars = Builder<Car>.CreateListOfSize(10).All().Do(car => car.Model = "SECOND MODEL").Build();
            var thirdCars = Builder<Car>.CreateListOfSize(10).All().Do(car => car.Model = "THIRD MODEL").Build();

            var firstPeople = Builder<Person>.CreateListOfSize(10).All().Do(person => person.MiddleName = "FIRST PERSON").Build();
            var secondPeople = Builder<Person>.CreateListOfSize(10).All().Do(person => person.MiddleName = "SECOND PERSON").Build();
            var thirdPeople = Builder<Person>.CreateListOfSize(10).All().Do(person => person.MiddleName = "THIRD PERSON").Build();

            var totalOrders = Enumerable.Empty<Order>();
            var totalCars = Enumerable.Empty<Car>();
            var totalPeople = Enumerable.Empty<Person>();

            var allOrders = Enumerable.Empty<Order>();
            var allPeople = Enumerable.Empty<Person>();
            var allCars = Enumerable.Empty<Car>();

            // Act
            for (int i = 1; i <= 3; i++)
            {
                var sheet = exportService.AddSheetForExport("Sheet " + i);

                Console.WriteLine("Added sheet: {0}", sheet.WorksheetName);

                var orderTables = new[] { firstOrders, secondOrders, thirdOrders };

                foreach (var orders in orderTables)
                {
                    totalOrders = totalOrders.Concat(orders);
                    var table = sheet.AddTableForExport(orders);

                    Console.WriteLine("added table: {0}", table.TableName);
                }

                var carTables = new[] { firstCars, secondCars, thirdCars };

                foreach (var cars in carTables)
                {
                    totalCars = totalCars.Concat(cars);
                    var table = sheet.AddTableForExport(cars);

                    Console.WriteLine("added table: {0}", table.TableName);
                }

                var personTables = new[] { firstPeople, secondPeople, thirdPeople };

                foreach (var people in personTables)
                {
                    totalPeople = totalPeople.Concat(people);
                    var table = sheet.AddTableForExport(people);

                    Console.WriteLine("added table: {0}", table.TableName);
                }
            }

            data = exportService.Export();

            importService = new ImportService(data);
            allOrders = importService.GetAll<Order>();
            allPeople = importService.GetAll<Person>();
            allCars = importService.GetAll<Car>();

            // Assert
            Assert.AreEqual(
                expected: totalOrders.Count(),
                actual: allOrders.Count(),
                message: "The total orders exported to three sheets do not add up to the total orders imported.");

            Assert.AreEqual(
                expected: totalPeople.Count(),
                actual: allPeople.Count(),
                message: "The total people exported to three sheets do not add up to the total people imported.");

            Assert.AreEqual(
                expected: totalCars.Count(),
                actual: allCars.Count(),
                message: "The total cars exported to three sheets do not add up to the total cars imported.");

            this.SaveToFile(data);
        }

        [TestMethod]
        [Description("Exporting one type in two sheets results in importing all types")]
        [TestCategory("1 x Type"), TestCategory("2 x Table"), TestCategory("2 x Sheet")]
        public void Exporting_Two_Sheets_One_Type_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> firstOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> secondOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> allOrders = Enumerable.Empty<Order>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(firstOrders);

            var sheet2 = exportService.AddSheetForExport("Sheet Two");
            sheet2.AddTableForExport(firstOrders);

            data = exportService.Export();
            importService = new ImportService(data);
            allOrders = importService.GetAll<Order>();

            // Assert
            var totalOrders = firstOrders.Concat(secondOrders);

            Assert.AreEqual(
                expected: totalOrders.Count(),
                actual: allOrders.Count(),
                message: "The total orders exported to two sheets do not add up to the total orders imported.");
        }

        [TestMethod]
        [Description("Exporting one type with list validation on a property results in export having list validation on that property")]
        [TestCategory("1 x Type"), TestCategory("2 x Table"), TestCategory("2 x Sheet"), TestCategory("1 x List Validation")]
        public void Exporting_One_Type_List_Validation_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> orders = Builder<Order>.CreateListOfSize(10).Build();
            var options = Enumerable.Range(1, 10).Select(i => string.Format("Option {0}", i));
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            var sheet1Table1 = sheet1.AddTableForExport(orders);

            // Act
            sheet1Table1.Options.AddListValidation(o => o.Reference, list =>
                    {
                        foreach (var option in options)
                        {
                            list.Formula.Values.Add(option);
                        }
                    });

            data = exportService.Export(); // Now export our sheet with validation

            // Assert
            using (var package = new ExcelPackage(new MemoryStream(data)))
            {
                var workSheet = package.Workbook.Worksheets["Sheet One"];
                var table = workSheet.Tables[sheet1Table1.TableName];
                var column = table.Columns["Reference"].Position + 1;
                var range = workSheet.Cells[table.Address.Start.Row + 1, column, table.Address.End.Row, column];
                var validation = workSheet.DataValidations.Find(v => v.Address.Address == range.Address) as IExcelDataValidationList;
                
                Assert.IsNotNull(
                    value: validation);

                Assert.AreEqual(
                    expected: typeof(ExcelDataValidationList),
                    actual: validation.GetType());

                Assert.AreEqual(
                    expected: options.Count(),
                    actual: validation.Formula.Values.Count);

                Assert.AreEqual(
                    expected: 1, 
                    actual: workSheet.DataValidations.Count);
            }

            this.SaveToFile(data);
        }

        [TestCleanup]
        [Description("Run after each unit test and clean up any left over resources")]
        public void CleanUp()
        {
            if (importService != null)
            {
                try
                {
                    importService.Dispose();
                }
                finally
                {
                    importService = null;
                }
            }

            if (exportService != null)
            {
                try
                {
                    exportService.Dispose();
                }
                finally
                {
                    exportService = null;
                }
            }
        }

        /// <summary>
        /// Save the data to file.
        /// </summary>
        /// <param name="data">The exported byte data from the exporter.</param>
        /// <param name="path">The path to save the file.</param>
        private void SaveToFile(byte[] data, string path = @"C:\\Temp\\Test.xlsx")
        {
            var isBuildServer = Environment.GetEnvironmentVariable("APPVEYOR") != null;

            if (!isBuildServer)
            {
                using (var stream = new MemoryStream(data))
                {
                    using (var package = new ExcelPackage(stream))
                    {
                        package.SaveAs(new FileInfo(path));
                    }
                }
            }
        }

    }
}