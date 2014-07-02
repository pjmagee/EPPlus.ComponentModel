// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IWorksheetConfiguration.cs" company="Patrick Magee">
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
//   The WorksheetConfiguration interface.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Export
{
    using System.Collections.Generic;

    /// <summary>
    /// The WorksheetConfiguration interface.
    /// </summary>
    public interface IWorksheetConfiguration
    {
        #region Public Properties

        /// <summary>
        /// Gets the ExportService.
        /// </summary>
        /// <remarks>
        /// The export service that is associated with this worksheet configuration.
        /// </remarks>
        IExportService ExportService { get; }

        /// <summary>
        /// Gets the table configurations.
        /// </summary>
        IEnumerable<ITableConfiguration> TableConfigurations { get; }

        /// <summary>
        /// Gets the worksheet name.
        /// </summary>
        string WorksheetName { get; }

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
        ITableConfiguration<T> AddTableForExport<T>(IEnumerable<T> collection, string tableName = null);

        #endregion
    }
}