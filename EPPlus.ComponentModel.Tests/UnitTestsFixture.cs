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
    using System.Linq;

    using EPPlus.ComponentModel.Export;
    using EPPlus.ComponentModel.Import;
    using EPPlus.ComponentModel.Tests.Entities;

    using FizzWare.NBuilder;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using MSTestExtensions;

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
        public void TestInitialize()
        {
            exportService = new ExportService();
        }

        [TestMethod]
        [Description("Exporting a sheet without a name results in an argument null exception")]
        [TestCategory("Guard")]
        public void Exporting_SheetName_Required_Test()
        {
            // Arrange

            // Act

            // Assert
            ExceptionAssert.Throws<ArgumentNullException>(() => exportService.AddSheetForExport(null));
        }

        [TestMethod]
        [Description("Exporting a table without a name results in the default name for that type")]
        [TestCategory("1 x Sheet"), TestCategory("1 x Type"), TestCategory("1 x Table")]
        public void Exporting_TableName_Test()
        {
            // Arrange
            var sheet1 = exportService.AddSheetForExport("Sheet One");

            // Act
            var orders = Builder<Order>.CreateListOfSize(10).Build();
            var table1 = sheet1.AddTableForExport<Order>(orders);

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

            // Act
            var orders = Builder<Order>.CreateListOfSize(10).Build();
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
            var orders = Builder<Order>.CreateListOfSize(10).Build();
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            var sheet2 = exportService.AddSheetForExport("Sheet Two");

            // Act
            var sheet1Table1 = sheet1.AddTableForExport(orders);
            var sheet1Table2 = sheet1.AddTableForExport(orders);

            var sheet2Table1 = sheet2.AddTableForExport(orders);
            var sheet2Table2 = sheet2.AddTableForExport(orders);

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
            IEnumerable<Order> ordersToInsert = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> ordersToLoad = Enumerable.Empty<Order>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(ordersToInsert);
            data = exportService.Export();

            importService = new ImportService(data);
            ordersToLoad = importService.GetAll<Order>();

            // Assert
            Assert.AreEqual(
                expected: ordersToInsert.Count(),
                actual: ordersToLoad.Count());
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
            IEnumerable<Order> ordersToLoad = Enumerable.Empty<Order>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(firstOrders);
            sheet1.AddTableForExport(secondOrders);
            data = exportService.Export();

            importService = new ImportService(data);
            ordersToLoad = importService.GetAll<Order>();

            // Assert
            var totalOrders = firstOrders.Concat(secondOrders);

            Assert.AreEqual(
                expected: totalOrders.Count(),
                actual: ordersToLoad.Count());

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
            IEnumerable<Order> ordersToLoad = Enumerable.Empty<Order>();

            IEnumerable<Person> firstPeople = Builder<Person>.CreateListOfSize(10).Build();
            IEnumerable<Person> peopleToLoad = Enumerable.Empty<Person>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(firstOrders);
            sheet1.AddTableForExport(secondOrders);
            sheet1.AddTableForExport(firstPeople);
            data = exportService.Export();

            importService = new ImportService(data);
            ordersToLoad = importService.GetAll<Order>();
            peopleToLoad = importService.GetAll<Person>();

            // Assert
            var totalOrders = firstOrders.Concat(secondOrders);

            Assert.AreEqual(
                expected: totalOrders.Count(),
                actual: ordersToLoad.Count(),
                message: "The total orders exported do not add up to the total orders imported.");

            Assert.AreEqual(
                expected: firstPeople.Count(),
                actual: peopleToLoad.Count(),
                message: "The total people exported do not add up to the total people imported.");
        }

        [TestMethod]
        [Description("Exporting two types results in importing both of all types")]
        [TestCategory("1 x Type"), TestCategory("2 x Table"), TestCategory("2 x Sheet")]
        public void Exporting_Two_Sheets_One_Type_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> firstOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> secondOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> ordersToLoad = Enumerable.Empty<Order>();

            // Act
            var sheet1 = exportService.AddSheetForExport("Sheet One");
            sheet1.AddTableForExport(firstOrders);

            var sheet2 = exportService.AddSheetForExport("Sheet Two");
            sheet2.AddTableForExport(firstOrders);

            data = exportService.Export();
            importService = new ImportService(data);
            ordersToLoad = importService.GetAll<Order>();

            // Assert
            var totalOrders = firstOrders.Concat(secondOrders);

            Assert.AreEqual(
                expected: totalOrders.Count(),
                actual: ordersToLoad.Count(),
                message: "The total orders exported to two sheets do not add up to the total orders imported.");
        }

        [TestCleanup]
        [Description("Run after each unit test and clean up any left over resources")]
        public void TestCleanUp()
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

    }
}