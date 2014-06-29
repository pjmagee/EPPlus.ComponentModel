// --------------------------------------------------------------------------------------------------------------------
// <copyright file="UnitTest1.cs" company="">
//   
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
        public void Orders_Test()
        {
            // Arrange
            byte[] data = null;
            IEnumerable<Order> ordersToInsert = Builder<Order>.CreateListOfSize(10).Build();
            IEnumerable<Order> ordersToLoad = Enumerable.Empty<Order>();

            // Act
            using (var exporter = new Exporter())
            {
                var sheet1 = exporter.AddSheetForExport("Sheet One");
                sheet1.AddTableForExport(ordersToInsert);
                data = exporter.Export();
            }

            using (var importer = new Importer(data))
            {
                ordersToLoad = importer.GetList<Order>();
            }

            // Assert
            Assert.AreEqual(expected: ordersToInsert.Count(), actual: ordersToLoad.Count());
        }
    }
}