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
    using System.Collections.Generic;
    using System.Linq;

    using EPPlus.ComponentModel.Export;
    using EPPlus.ComponentModel.Import;
    using EPPlus.ComponentModel.Tests.Entities;

    using FizzWare.NBuilder;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The unit test 1.
    /// </summary>
    [TestClass]
    public class UnitTestsFixture
    {
        [TestMethod]
        [Description("Tests that exporting 10 orders results in importing 10 orders.")]
        public void Exporting_One_Table_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> ordersToInsert = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> ordersToLoad = Enumerable.Empty<Order>();

            // Act
            using (var exporter = new ExportService())
            {
                var sheet1 = exporter.AddSheetForExport("Sheet One");
                sheet1.AddTableForExport(ordersToInsert);
                data = exporter.Export();
            }

            using (var importer = new ImportService(data))
            {
                ordersToLoad = importer.GetAll<Order>();
            }

            // Assert
            Assert.AreEqual(expected: ordersToInsert.Count(), actual: ordersToLoad.Count());
        }

        [TestMethod]
        [Description("Tests that exporting orders to two different tables results in importing all orders.")]
        public void Exporting_Two_Tables_Test()
        {
            // Arrange
            byte[] data;
            IEnumerable<Order> firstOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> secondOrders = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> ordersToLoad = Enumerable.Empty<Order>();

            // Act
            using (var exporter = new ExportService())
            {
                var sheet1 = exporter.AddSheetForExport("Sheet One");
                sheet1.AddTableForExport(firstOrders);
                sheet1.AddTableForExport(secondOrders);
                data = exporter.Export();
            }

            using (var importer = new ImportService(data))
            {
                ordersToLoad = importer.GetAll<Order>();
            }

            // Assert
            var totalOrders = firstOrders.Concat(secondOrders);

            Assert.AreEqual(expected: totalOrders.Count(), actual: ordersToLoad.Count());

        }
    }
}