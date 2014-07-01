// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ObjectConfiguration.T.cs" company="Patrick Magee">
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
//   The object configuration.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Export
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;

    using OfficeOpenXml.DataValidation.Contracts;
    using OfficeOpenXml.Table;

    /// <summary>
    /// The object configuration.
    /// </summary>
    /// <typeparam name="T">
    /// </typeparam>
    public class ObjectConfiguration<T> : IObjectConfiguration<T>
    {
        #region Fields

        /// <summary>
        /// The date validations.
        /// </summary>
        private readonly Dictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationDateTime>>, Action<IExcelDataValidationDateTime>>> dateValidations;

        /// <summary>
        /// The decimal validators.
        /// </summary>
        private readonly Dictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationDecimal>>, Action<IExcelDataValidationDecimal>>> decimalValidators;

        /// <summary>
        /// The ignoredproperties.
        /// </summary>
        private readonly IList<PropertyInfo> ignoredproperties;

        /// <summary>
        /// The integer validations.
        /// </summary>
        private readonly Dictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationInt>>, Action<IExcelDataValidationInt>>> integerValidations;

        /// <summary>
        /// The list validations.
        /// </summary>
        private readonly
            Dictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationList>>, Action<IExcelDataValidationList>>> listValidations;

        /// <summary>
        /// The property substitutes.
        /// </summary>
        private readonly IDictionary<PropertyInfo, object> propertySubstitutes;

        /// <summary>
        /// The table.
        /// </summary>
        private readonly ExcelTable table;

        /// <summary>
        /// The table configuration.
        /// </summary>
        private readonly ITableConfiguration<T> tableConfiguration;

        /// <summary>
        /// The type properties.
        /// </summary>
        private readonly PropertyInfo[] typeProperties;

        /// <summary>
        /// The print headers.
        /// </summary>
        private bool printHeaders;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ObjectConfiguration{T}"/> class.
        /// </summary>
        /// <param name="table">
        /// The table.
        /// </param>
        /// <param name="tableConfiguration">
        /// The table configuration.
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// </exception>
        public ObjectConfiguration(ExcelTable table, ITableConfiguration<T> tableConfiguration) : this()
        {
            if (table == null)
            {
                throw new ArgumentNullException("table");
            }

            if (tableConfiguration == null)
            {
                throw new ArgumentNullException("tableConfiguration");
            }

            this.table = table;
            this.tableConfiguration = tableConfiguration;
        }

        /// <summary>
        /// Prevents a default instance of the <see cref="ObjectConfiguration{T}"/> class from being created.
        /// </summary>
        private ObjectConfiguration()
        {
            this.typeProperties = typeof(T).GetProperties().ToArray();
            this.printHeaders = false;
            this.ignoredproperties = new List<PropertyInfo>();
            this.propertySubstitutes = new Dictionary<PropertyInfo, object>();
            this.listValidations = new Dictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationList>>, Action<IExcelDataValidationList>>>();
            this.dateValidations = new Dictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationDateTime>>, Action<IExcelDataValidationDateTime>>>();
            this.integerValidations = new Dictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationInt>>, Action<IExcelDataValidationInt>>>();
            this.decimalValidators = new Dictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationDecimal>>, Action<IExcelDataValidationDecimal>>>();
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets the date validations.
        /// </summary>
        public IDictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationDateTime>>, Action<IExcelDataValidationDateTime>>> DateValidations
        {
            get
            {
                return this.dateValidations;
            }
        }

        /// <summary>
        /// Gets the decimal validators.
        /// </summary>
        public IDictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationDecimal>>, Action<IExcelDataValidationDecimal>>> DecimalValidators
        {
            get
            {
                return this.decimalValidators;
            }
        }

        /// <summary>
        /// Gets the ExportService.
        /// </summary>
        public IExportService ExportService
        {
            get
            {
                return this.WorksheetConfiguration.ExportService;
            }
        }

        /// <summary>
        /// Gets the ignored properties.
        /// </summary>
        public IEnumerable<PropertyInfo> IgnoredProperties
        {
            get
            {
                return this.ignoredproperties;
            }
        }

        /// <summary>
        /// Gets the integer validators.
        /// </summary>
        public IDictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationInt>>, Action<IExcelDataValidationInt>>> IntegerValidators
        {
            get
            {
                return this.integerValidations;
            }
        }

        /// <summary>
        /// Gets the list validations.
        /// </summary>
        public IDictionary<PropertyInfo, Tuple<Action<Action<IExcelDataValidationList>>, Action<IExcelDataValidationList>>> ListValidations
        {
            get
            {
                return this.listValidations;
            }
        }

        /// <summary>
        /// Gets a value indicating whether print headers.
        /// </summary>
        public bool PrintHeaders
        {
            get
            {
                return this.printHeaders;
            }
        }

        /// <summary>
        /// Gets the property substitutes.
        /// </summary>
        public IDictionary<PropertyInfo, object> PropertySubstitutes
        {
            get
            {
                return this.propertySubstitutes;
            }
        }

        /// <summary>
        /// Gets the table configuration.
        /// </summary>
        public ITableConfiguration TableConfiguration
        {
            get
            {
                return this.tableConfiguration;
            }
        }

        /// <summary>
        /// Gets the worksheet configuration.
        /// </summary>
        public IWorksheetConfiguration WorksheetConfiguration
        {
            get
            {
                return this.tableConfiguration.WorksheetConfiguration;
            }
        }

        #endregion

        #region Explicit Interface Properties

        /// <summary>
        /// Gets the table configuration.
        /// </summary>
        ITableConfiguration<T> ITableNavigatable<T>.TableConfiguration
        {
            get
            {
                return this.tableConfiguration;
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the end row.
        /// </summary>
        private int EndRow
        {
            get
            {
                return this.table.Address.End.Row;
            }
        }

        /// <summary>
        /// Gets the start row.
        /// </summary>
        private int StartRow
        {
            get
            {
                return this.table.Address.Start.Row + 1;
            }
        }

        /// <summary>
        /// Gets the type properties.
        /// </summary>
        private List<PropertyInfo> TypeProperties
        {
            get
            {
                return this.typeProperties.Except(this.ignoredproperties).ToList();
            }
        }

        #endregion

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
        public IObjectConfiguration<T> AddDateValidation(Expression<Func<T, object>> selector, Action<IExcelDataValidationDateTime> action) 
        {
            var property = this.GetProperty(selector);
            var index = this.GetPropertyIndex(property);

            var dateTimeValidation = new Action<Action<IExcelDataValidationDateTime>>(
                callback =>
                    {
                        var range = this.table.WorkSheet.Cells[this.StartRow, index, this.EndRow, index];
                        var validation = range.DataValidation.AddDateTimeDataValidation();
                        callback(validation);
                    });

            this.dateValidations[property] = Tuple.Create(dateTimeValidation, action);
            return this;
        }

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
        public IObjectConfiguration<T> AddDecimalValidation(Expression<Func<T, object>> selector, Action<IExcelDataValidationDecimal> action)
        {
            var property = this.GetProperty(selector);
            var index = this.GetPropertyIndex(property);

            var decimalValidation = new Action<Action<IExcelDataValidationDecimal>>(
                process =>
                    {
                        var range = this.table.WorkSheet.Cells[this.StartRow + 1, index, this.EndRow, index];
                        var validation = range.DataValidation.AddDecimalDataValidation();
                        process(validation);
                    });

            this.decimalValidators[property] = Tuple.Create(decimalValidation, action);
            return this;
        }

        /// <summary>
        /// The add header row.
        /// </summary>
        /// <param name="enabled">
        /// The enabled.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        public IObjectConfiguration<T> AddHeaderRow(bool enabled = true)
        {
            this.printHeaders = enabled;
            return this;
        }

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
        public IObjectConfiguration<T> AddIntegerValidation(Expression<Func<T, object>> selector, Action<IExcelDataValidationInt> action)
        {
            var property = this.GetProperty(selector);
            var index = this.GetPropertyIndex(property);

            var integerValidation = new Action<Action<IExcelDataValidationInt>>(
                process =>
                    {
                        var range = this.table.WorkSheet.Cells[this.StartRow + 1, index, this.EndRow, index];
                        var validation = range.DataValidation.AddIntegerDataValidation();
                        process(validation);
                    });

            this.integerValidations[property] = Tuple.Create(integerValidation, action);
            return this;
        }

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
        public IObjectConfiguration<T> AddListValidation(Expression<Func<T, object>> selector, Action<IExcelDataValidationList> action)
        {
            var property = this.GetProperty(selector);
            var index = this.GetPropertyIndex(property);

            var listValidation = new Action<Action<IExcelDataValidationList>>(
                process =>
                    {
                        var range = this.table.WorkSheet.Cells[this.StartRow + 1, index, this.EndRow, index];
                        var validation = range.DataValidation.AddListDataValidation();
                        process(validation);
                    });

            this.listValidations[property] = Tuple.Create(listValidation, action);

            return this;
        }

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
        public IObjectConfiguration<T> AddPropertySubstitue(Expression<Func<T, object>> selector, object substitute)
        {
            var property = this.GetProperty(selector);
            this.propertySubstitutes[property] = substitute;
            return this;
        }

        /// <summary>
        /// The add property to ignore.
        /// </summary>
        /// <param name="selector">
        /// The selector.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        public IObjectConfiguration<T> AddPropertyToIgnore(Expression<Func<T, object>> selector)
        {
            var property = this.GetProperty(selector);
            this.ignoredproperties.Add(property);
            return this;
        }

        #endregion

        #region Explicit Interface Methods

        /// <summary>
        /// The add header row.
        /// </summary>
        /// <param name="printHeaders">
        /// The print headers.
        /// </param>
        /// <returns>
        /// The <see cref="IObjectConfiguration"/>.
        /// </returns>
        IObjectConfiguration IObjectConfiguration.AddHeaderRow(bool printHeaders)
        {
            return this.AddHeaderRow(printHeaders);
        }

        #endregion

        #region Methods

        /// <summary>
        /// The get property.
        /// </summary>
        /// <param name="selector">
        /// The selector.
        /// </param>
        /// <returns>
        /// The <see cref="PropertyInfo"/>.
        /// </returns>
        /// <exception cref="ArgumentException">
        /// </exception>
        private PropertyInfo GetProperty(Expression<Func<T, object>> selector)
        {
            MemberExpression expression;

            var unaryExpression = selector.Body as UnaryExpression;

            if (unaryExpression != null)
            {
                var operand = unaryExpression.Operand as MemberExpression;

                if (operand != null)
                {
                    expression = operand;
                }
                else
                {
                    throw new ArgumentException();
                }
            }
            else
            {
                var memberExpression = selector.Body as MemberExpression;

                if (memberExpression != null)
                {
                    expression = memberExpression;
                }
                else
                {
                    throw new ArgumentException();
                }
            }

            return (PropertyInfo)expression.Member;
        }

        /// <summary>
        /// The get property index.
        /// </summary>
        /// <param name="info">
        /// The info.
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        private int GetPropertyIndex(PropertyInfo info)
        {
            return this.TypeProperties.IndexOf(info) + 1;
        }

        #endregion
    }
}