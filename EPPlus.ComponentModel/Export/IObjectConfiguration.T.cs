// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IObjectConfiguration.T.cs" company="Patrick Magee">
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
    using System.Linq.Expressions;

    using OfficeOpenXml.DataValidation.Contracts;

    /// <summary>
    /// The ObjectConfiguration interface.
    /// </summary>
    /// <typeparam name="T">
    /// </typeparam>
    public interface IObjectConfiguration<T> : ITableNavigatable<T>, IObjectConfiguration
    {
        #region Public Methods and Operators

        /// <summary>
        /// The add date validation.
        /// </summary>
        /// <param name="selector">
        /// The selector.
        /// </param>
        /// <param name="action">
        /// The action.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        IObjectConfiguration<T> AddDateValidation(
            Expression<Func<T, object>> selector, 
            Action<IExcelDataValidationDateTime> action);

        /// <summary>
        /// The add decimal validation.
        /// </summary>
        /// <param name="selector">
        /// The selector.
        /// </param>
        /// <param name="action">
        /// The action.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        IObjectConfiguration<T> AddDecimalValidation(
            Expression<Func<T, object>> selector, 
            Action<IExcelDataValidationDecimal> action);

        /// <summary>
        /// The add integer validation.
        /// </summary>
        /// <param name="selector">
        /// The selector.
        /// </param>
        /// <param name="action">
        /// The action.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        IObjectConfiguration<T> AddIntegerValidation(
            Expression<Func<T, object>> selector, 
            Action<IExcelDataValidationInt> action);

        /// <summary>
        /// The add list validation.
        /// </summary>
        /// <param name="selector">
        /// The selector.
        /// </param>
        /// <param name="action">
        /// The action.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        IObjectConfiguration<T> AddListValidation(
            Expression<Func<T, object>> selector, 
            Action<IExcelDataValidationList> action);

        /// <summary>
        /// The add property substitue.
        /// </summary>
        /// <param name="selector">
        /// The selector.
        /// </param>
        /// <param name="substitute">
        /// The substitute.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        IObjectConfiguration<T> AddPropertySubstitute(Expression<Func<T, object>> selector, object substitute);

        /// <summary>
        /// The add property to ignore.
        /// </summary>
        /// <param name="selector">
        /// The selector.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        IObjectConfiguration<T> AddPropertyToIgnore(Expression<Func<T, object>> selector);

        #endregion
    }
}