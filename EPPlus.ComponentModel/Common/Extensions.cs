// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Extensions.cs" company="Patrick Magee">
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
//   The extensions.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Common
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.Linq;

    using OfficeOpenXml.Table;

    /// <summary>
    /// The extensions.
    /// </summary>
    public static class Extensions
    {
        #region Public Methods and Operators

        /// <summary>
        /// The to data table.
        /// </summary>
        /// <param name="table">
        /// The table.
        /// </param>
        /// <returns>
        /// The <see cref="DataTable"/>.
        /// </returns>
        public static DataTable ToDataTable(this ExcelTable table)
        {
            DataTable dataTable = new DataTable();

            var tableStartRow = table.Address.Start.Row;

            var headerRow = table.WorkSheet.Cells[tableStartRow, table.Address.Start.Column, tableStartRow, table.Address.End.Column];

            foreach (var header in headerRow)
            {
                dataTable.Columns.Add(header.Text);
            }

            for (var rowIndex = table.Address.Start.Row + 1; rowIndex <= table.Address.End.Row; rowIndex++)
            {
                var tableRow = table.WorkSheet.Cells[table.Address.Start.Row + 1, table.Address.Start.Column, rowIndex, table.Address.End.Column];
                var dataRow = dataTable.NewRow();

                foreach (var cell in tableRow)
                {
                    dataRow[cell.Start.Column - 1] = cell.Text;
                }

                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        public static DataTable ToDataTable<T>(this IEnumerable<T> collection, string tableName = null)
        {
            Type type = typeof(T);
            DataTable dataTable = new DataTable(tableName ?? type.Name);
            var properties = type.GetProperties();

            foreach (var property in properties)
            {
                // Adding metadata here can change the way EPPlus will fill in the cells.
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

            if (dataTable.Rows.Count != collection.Count())
            {
                throw new Exception("DataTable conversion invalid.");
            }

            return dataTable;
        }

        #endregion
    }
}