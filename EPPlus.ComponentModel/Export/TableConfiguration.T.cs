// --------------------------------------------------------------------------------------------------------------------
// <copyright file="TableConfiguration.T.cs" company="Patrick Magee">
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
//   The table configuration.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Export
{
    using System;

    using OfficeOpenXml.Table;

    /// <summary>
    /// The table configuration.
    /// </summary>
    /// <typeparam name="T">
    /// </typeparam>
    public class TableConfiguration<T> : ITableConfiguration<T>, ITableConfiguration
    {
        #region Fields

        /// <summary>
        /// The options.
        /// </summary>
        private readonly IObjectConfiguration<T> options;

        /// <summary>
        /// The table.
        /// </summary>
        private readonly ExcelTable table;

        /// <summary>
        /// The worksheet configuration.
        /// </summary>
        private readonly IWorksheetConfiguration worksheetConfiguration;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="TableConfiguration{T}"/> class.
        /// </summary>
        /// <param name="worksheetConfiguration">
        /// The worksheet configuration.
        /// </param>
        /// <param name="table">
        /// The table.
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// </exception>
        public TableConfiguration(IWorksheetConfiguration worksheetConfiguration, ExcelTable table) : this()
        {
            if (worksheetConfiguration == null)
            {
                throw new ArgumentNullException("worksheetConfiguration");
            }

            if (table == null)
            {
                throw new ArgumentNullException("table");
            }

            this.table = table;
            this.worksheetConfiguration = worksheetConfiguration;
            this.options = new ObjectConfiguration<T>(table, this);
        }

        /// <summary>
        /// Prevents a default instance of the <see cref="TableConfiguration{T}"/> class from being created.
        /// </summary>
        private TableConfiguration()
        {
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
                return this.worksheetConfiguration.ExportService;
            }
        }

        /// <summary>
        /// Gets the options.
        /// </summary>
        public IObjectConfiguration Options
        {
            get
            {
                return this.options;
            }
        }

        /// <summary>
        /// Gets the table name.
        /// </summary>
        public string TableName
        {
            get
            {
                return this.table.Name;
            }
        }

        /// <summary>
        /// Gets the worksheet configuration.
        /// </summary>
        public IWorksheetConfiguration WorksheetConfiguration
        {
            get
            {
                return this.worksheetConfiguration;
            }
        }

        #endregion

        #region Explicit Interface Properties

        /// <summary>
        /// Gets the options.
        /// </summary>
        IObjectConfiguration<T> ITableConfiguration<T>.Options
        {
            get
            {
                return this.options;
            }
        }

        #endregion
    }
}