// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IObjectConfiguration.cs" company="Patrick Magee">
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
//   The ObjectConfiguration interface.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Export
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;

    using OfficeOpenXml.DataValidation.Contracts;

    /// <summary>
    /// The ObjectConfiguration interface.
    /// </summary>
    public interface IObjectConfiguration : ITableNavigatable
    {
        #region Public Properties

        /// <summary>
        /// Gets the date validations.
        /// </summary>
        IDictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationDateTime>>, Action<IExcelDataValidationDateTime>>> DateValidations { get; }

        /// <summary>
        /// Gets the decimal validators.
        /// </summary>
        IDictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationDecimal>>, Action<IExcelDataValidationDecimal>>> DecimalValidators { get; }

        /// <summary>
        /// Gets the ignored properties.
        /// </summary>
        IEnumerable<PropertyInfo> IgnoredProperties { get; }

        /// <summary>
        /// Gets the integer validators.
        /// </summary>
        IDictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationInt>>, Action<IExcelDataValidationInt>>> IntegerValidators { get; }

        /// <summary>
        /// Gets the list validations.
        /// </summary>
        IDictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationList>>, Action<IExcelDataValidationList>>> ListValidations { get; }

        /// <summary>
        /// Gets a value indicating whether print headers.
        /// </summary>
        bool PrintHeaders { get; }

        /// <summary>
        /// Gets the property substitutes.
        /// </summary>
        IDictionary<PropertyInfo, object> PropertySubstitutes { get; }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The add header row.
        /// </summary>
        /// <param name="printHeaders">
        /// The print headers.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        IObjectConfiguration AddHeaderRow(bool printHeaders = true);

        #endregion
    }
}