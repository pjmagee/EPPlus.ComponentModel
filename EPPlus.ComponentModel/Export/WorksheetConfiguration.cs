// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WorksheetConfiguration.cs" company="Patrick Magee">
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
//   The worksheet configuration.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Export
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.Entity.Design.PluralizationServices;
    using System.Globalization;
    using System.Linq;
    using System.Reflection;

    using EPPlus.ComponentModel.Common;

    using OfficeOpenXml;
    using OfficeOpenXml.Table;

    /// <summary>
    /// The worksheet configuration.
    /// </summary>
    public class WorksheetConfiguration : IWorksheetConfiguration
    {
        #region Fields

        /// <summary>
        /// The ExportService.
        /// </summary>
        private readonly IExportService exportService;

        /// <summary>
        /// Tables must have a unique table name. 
        /// This stores the number of tables for each type, so that the next table name
        /// can store the index of that type. 
        /// </summary>
        /// <example>
        /// table_type_1, table_type_2
        /// </example>
        private readonly IDictionary<Type, int> typeCount;

        /// <summary>
        /// The pluraliser for tables
        /// </summary>
        private readonly PluralizationService pluralizationService;

        /// <summary>
        /// The table configurations.
        /// </summary>
        private readonly List<ITableConfiguration> tableConfigurations;

        /// <summary>
        /// The worksheet.
        /// </summary>
        private readonly ExcelWorksheet worksheet;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetConfiguration"/> class.
        /// </summary>
        /// <param name="exportService">
        /// The ExportService.
        /// </param>
        /// <param name="worksheet">
        /// The worksheet.
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// </exception>
        public WorksheetConfiguration(IExportService exportService, ExcelWorksheet worksheet) : this()
        {
            if (exportService == null)
            {
                throw new ArgumentNullException("exportService");
            }

            if (worksheet == null)
            {
                throw new ArgumentNullException("worksheet");
            }

            this.exportService = exportService;
            this.worksheet = worksheet;
        }

        /// <summary>
        /// Prevents a default instance of the <see cref="WorksheetConfiguration"/> class from being created.
        /// </summary>
        private WorksheetConfiguration()
        {
            this.pluralizationService = PluralizationService.CreateService(CultureInfo.CurrentCulture);
            this.tableConfigurations = new List<ITableConfiguration>();
            this.typeCount = new Dictionary<Type, int>();
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets the ExportService.
        /// </summary>
        public IExportService ExportService
        {
            get
            {
                return this.exportService;
            }
        }

        /// <summary>
        /// Gets the table configurations.
        /// </summary>
        public IEnumerable<ITableConfiguration> TableConfigurations
        {
            get
            {
                return this.tableConfigurations;
            }
        }

        /// <summary>
        /// Gets the worksheet name.
        /// </summary>
        public string WorksheetName
        {
            get
            {
                return this.worksheet.Name;
            }
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The add table for export.
        /// </summary>
        /// <param name="collection">
        /// The collection.
        /// </param>
        /// <param name="tableName">
        /// The table name.
        /// </param>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="ITableConfiguration"/>.
        /// </returns>
        public ITableConfiguration<T> AddTableForExport<T>(IEnumerable<T> collection, string tableName = null)
        {
            var key = this.GetKey<T>(tableName ?? string.Empty);
            var rangeToFill = this.GetRangeToFill<T>();
            var dataTable = collection.ToDataTable(key);

            // Load from dataTable will set the table name from the table name set on the dataTable
            var filledRange = rangeToFill.LoadFromDataTable(dataTable, PrintHeaders: true, TableStyle: TableStyles.Dark1);
            FormatDates(filledRange, dataTable);
            var table = worksheet.Tables[key];
            return this.CreateTableConfiguration<T>(table);
        }

        #endregion

        #region Methods

        private void FormatDates(ExcelRangeBase range, DataTable table)
        {
            var columns = from DataColumn d in table.Columns where d.DataType == typeof(DateTime) || d.ColumnName.Contains("Date") select d.Ordinal + 1;

            foreach (var column in columns)
            {
                worksheet.Cells[range.Start.Row + 1, column, range.End.Row, column].Style.Numberformat.Format = "dd/mm/yyyy";
            }
        }

        /// <summary>
        /// The create table configuration.
        /// </summary>
        /// <param name="table">
        /// The table.
        /// </param>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="ITableConfiguration"/>.
        /// </returns>
        private ITableConfiguration<T> CreateTableConfiguration<T>(ExcelTable table)
        {
            ITableConfiguration<T> configuration = new TableConfiguration<T>(this, table);
            this.tableConfigurations.Add(configuration);
            return configuration;
        }

        /// <summary>
        /// It gets the correct ExcelRange to fill from by inserting new rows at the end of the current range.
        /// This ensures there is an empty row between new tables being inserted into the spreadsheet. 
        /// </summary>
        /// <typeparam name="T">The type</typeparam>
        /// <returns>
        /// The <see cref="ExcelRange"/> that is to be filled from a collection or datatable.
        /// </returns>
        private ExcelRange GetRangeToFill<T>()
        {
            var isEmpty = this.worksheet.Dimension == null;
            this.worksheet.InsertRow(isEmpty ? 1 : this.worksheet.Dimension.End.Row, 2);
            return this.worksheet.Cells[isEmpty ? 1 : this.worksheet.Dimension.End.Row, 1];
        }

        /// <summary>
        /// The get key.
        /// </summary>
        /// <param name="tableName">
        /// The table name.
        /// </param>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException">
        /// </exception>
        public string GetKey<T>(string tableName)
        {
            if (tableName == null)
            {
                throw new ArgumentNullException("tableName");
            }

            var type = typeof(T);
            var count = 1;

            if (!typeCount.ContainsKey(type))
            {
                typeCount[type] = 1;
            }
            else
            {
                count = typeCount[type] + 1;
                typeCount[type] = count;
            }

            tableName = tableName.Replace(" ", "_");
            var sheetName = WorksheetName.Replace(" ", "_");
            var plural = pluralizationService.Pluralize(typeof(T).Name);

            return string.Format("{0}_{1}_{2}_{3}", sheetName, tableName, plural, count).Replace("__", "_");
        }

        #endregion
    }
}