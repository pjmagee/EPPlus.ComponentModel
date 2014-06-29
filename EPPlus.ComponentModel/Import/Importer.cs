// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Importer.cs" company="Patrick Magee">
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
//   The importer.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Import
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Data.Entity.Design.PluralizationServices;
    using System.Globalization;
    using System.IO;
    using System.Linq;

    using EPPlus.ComponentModel.Common;

    using OfficeOpenXml;

    /// <summary>
    /// The importer.
    /// </summary>
    public class Importer : IImporter, IDisposable
    {
        #region Fields

        /// <summary>
        /// The package.
        /// </summary>
        private readonly ExcelPackage package;

        /// <summary>
        /// The pluraliser for tables
        /// </summary>
        private readonly PluralizationService pluralizationService;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Importer"/> class.
        /// </summary>
        /// <param name="data">
        /// The data.
        /// </param>
        public Importer(byte[] data)
            : this()
        {
            this.package = new ExcelPackage(new MemoryStream(data));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Importer"/> class.
        /// </summary>
        /// <param name="path">
        /// The path.
        /// </param>
        public Importer(string path)
            : this()
        {
            var data = File.ReadAllBytes(path);
            this.package = new ExcelPackage(new MemoryStream(data));
        }

        /// <summary>
        /// Prevents a default instance of the <see cref="Importer"/> class from being created.
        /// </summary>
        private Importer()
        {
            this.pluralizationService = PluralizationService.CreateService(CultureInfo.CurrentCulture);
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The dispose.
        /// </summary>
        public void Dispose()
        {
            try
            {
                this.package.Dispose();
            }
            catch
            {
            }
        }

        /// <summary>
        /// The get list.
        /// </summary>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="IEnumerable"/>.
        /// </returns>
        public IEnumerable<T> GetList<T>()
        {
            var name = typeof(T).Name;
            var pluralName = this.pluralizationService.Pluralize(name);
            var properties = typeof(T).GetProperties();

            var tables = from worksheet in this.package.Workbook.Worksheets
                         from table in worksheet.Tables
                         where table.Name.EndsWith(pluralName)
                         select table;

            foreach (var excelTable in tables)
            {
                var dataTable = excelTable.ToDataTable();

                foreach (var row in dataTable.Rows.Cast<DataRow>())
                {
                    var instance = Activator.CreateInstance<T>();

                    foreach (var column in dataTable.Columns.Cast<DataColumn>())
                    {
                        var property = properties.Single(p => p.Name == column.ColumnName);
                        var propertyType = property.PropertyType;
                        var propertyValue = row[column];
                        var typeConverter = TypeDescriptor.GetConverter(propertyType);
                        object value = typeConverter.ConvertFrom(propertyValue);
                        property.SetValue(instance, value);
                    }

                    yield return instance;
                }
            }
        }

        #endregion
    }
}