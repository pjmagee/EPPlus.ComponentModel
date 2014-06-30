// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IImportService.cs" company="Patrick Magee">
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
//   The ImportService interface.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Import
{
    using System.Collections.Generic;

    /// <summary>
    /// The ImportService interface.
    /// </summary>
    public interface IImportService
    {
        #region Public Methods and Operators

        /// <summary>
        /// Gets all the types found in any tables in the entire worksheet.
        /// </summary>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="IEnumerable{T}"/> found in the entire workbook.
        /// </returns>
        IEnumerable<T> GetAll<T>();

        /// <summary>
        /// Gets all the types in tables specified in the given sheet name.
        /// </summary>
        /// <typeparam name="T">The type of object.</typeparam>
        /// <param name="sheetName">The name of the sheet in the excel workbook.</param>
        /// <returns>
        /// The <see cref="IEnumerable{T}"/> found in all tabels of the given type in the given sheet.
        /// </returns>
        IEnumerable<T> GetListFromSheet<T>(string sheetName);

        /// <summary>e
        /// Gets all the types in the specified table name.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tableName"></param>
        /// <returns>
        ///  The <see cref="IEnumerable{T}"/> found in the given table name.
        /// </returns>
        IEnumerable<T> GetListFromTable<T>(string tableName);

        #endregion
    }
}