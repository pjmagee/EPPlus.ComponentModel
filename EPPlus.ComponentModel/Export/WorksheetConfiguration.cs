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
        private readonly IDictionary<string, int> typeCount;

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
            this.typeCount = new Dictionary<string, int>();
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
            var rangeToFill = this.GetRangeToFill<T>();
            var dataTable = this.CreateDataTable(collection);
            var filledRange = rangeToFill.LoadFromDataTable(dataTable, PrintHeaders: true);
            var table = this.CreateTable<T>(tableName, filledRange);
            this.worksheet.Cells[table.Address.Address].AutoFitColumns();
            return this.CreateTableConfiguration<T>(table);
        }

        #endregion

        #region Methods

        /// <summary>
        /// The create data table.
        /// </summary>
        /// <param name="collection">
        /// The collection.
        /// </param>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="DataTable"/>.
        /// </returns>
        private DataTable CreateDataTable<T>(IEnumerable<T> collection)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            var properties = this.GetProperties<T>();

            foreach (var property in properties)
            {
                dataTable.Columns.Add(property.Name);
            }

            foreach (var item in collection)
            {
                var dataRow = dataTable.NewRow();

                foreach (var property in properties)
                {
                    dataRow[property.Name] = property.GetValue(item).ToString();
                }

                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        /// <summary>
        /// The create table.
        /// </summary>
        /// <param name="tableName">
        /// The table name.
        /// </param>
        /// <param name="filledRange">
        /// The filled range.
        /// </param>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="ExcelTable"/>.
        /// </returns>
        private ExcelTable CreateTable<T>(string tableName, ExcelRangeBase filledRange)
        {
            var pluralName = this.pluralizationService.Pluralize(typeof(T).Name);
           
            var tableCount = 1;

            if (typeCount.ContainsKey(pluralName))
            {
                typeCount[pluralName] += 1;
                tableCount = typeCount[pluralName];
            }
            else
            {
                typeCount[pluralName] = tableCount;
            }

            var identifer = string.Format(TableConfiguration<T>.TableKeyIdentifer, pluralName, tableCount);
            tableName = tableName == null ? identifer : tableName + identifer;

            var excelTable = this.worksheet.Tables.Add(filledRange, tableName);
            excelTable.TableStyle = TableStyles.Dark1;
            excelTable.ShowHeader = true;
            excelTable.ShowFilter = true;

            return excelTable;
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
        /// The get columns.
        /// </summary>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        private int GetColumns<T>()
        {
            return typeof(T).GetProperties().Count();
        }

        /// <summary>
        /// The get end range.
        /// </summary>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="ExcelRange"/>.
        /// </returns>
        private ExcelRange GetEndRange<T>()
        {
            var columns = this.GetColumns<T>();

            return this.worksheet.Dimension == null ?
                this.worksheet.Cells[FromRow: 1, FromCol: 1, ToRow: 1, ToCol: columns] : 
                this.worksheet.Cells[this.worksheet.Dimension.End.Row - 1, 1, this.worksheet.Dimension.End.Row, columns];
        }

        /// <summary>
        /// The get properties.
        /// </summary>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="PropertyInfo[]"/>.
        /// </returns>
        private PropertyInfo[] GetProperties<T>()
        {
            return typeof(T).GetProperties().OrderBy(p => p.Name.Length).ToArray();
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
            var currentRange = this.GetEndRange<T>();
            this.worksheet.InsertRow(currentRange.End.Row, 2);
            var rangeToFill = this.GetEndRange<T>();
            return rangeToFill;
        }

        #endregion
    }
}