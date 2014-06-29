// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IExporter.cs" company="Patrick Magee">
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
//   The Exporter interface.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Export
{
    using System.Collections.Generic;

    /// <summary>
    /// The Exporter interface.
    /// </summary>
    public interface IExporter
    {
        #region Public Properties

        /// <summary>
        /// Gets the worksheet configurations.
        /// </summary>
        IEnumerable<IWorksheetConfiguration> WorksheetConfigurations { get; }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The add sheet for export.
        /// </summary>
        /// <param name="sheetName">
        /// The sheet name.
        /// </param>
        /// <returns>
        /// The <see cref="IWorksheetConfiguration"/>.
        /// </returns>
        IWorksheetConfiguration AddSheetForExport(string sheetName);

        /// <summary>
        /// The export.
        /// </summary>
        /// <param name="path">
        /// The path.
        /// </param>
        void Export(string path);

        /// <summary>
        /// The export.
        /// </summary>
        /// <returns>
        /// The <see cref="byte[]"/>.
        /// </returns>
        byte[] Export();

        #endregion
    }
}