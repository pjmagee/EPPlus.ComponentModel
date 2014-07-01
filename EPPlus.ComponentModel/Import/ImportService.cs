// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ImportService.cs" company="Patrick Magee">
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
//   The importer.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Import
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Data.Entity.Design.PluralizationServices;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    using EPPlus.ComponentModel.Common;
    using EPPlus.ComponentModel.Export;

    using OfficeOpenXml;
    using OfficeOpenXml.Table;

    /// <summary>
    /// The importer.
    /// </summary>
    public class ImportService : IImportService, IDisposable
    {
        #region Static Fields

        /// <summary>
        /// The pluraliser for tables
        /// </summary>
        private static readonly PluralizationService PluralizationService;

        #endregion

        #region Fields

        /// <summary>
        /// The cached types to use when multiple types are requested. Because reflection is expensive.
        /// </summary>
        private readonly IDictionary<Type, CachedTypeInformation> cachedTypes;

        /// <summary>
        /// The package.
        /// </summary>
        private readonly ExcelPackage package;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes static members of the <see cref="ImportService"/> class.
        /// </summary>
        static ImportService()
        {
            PluralizationService = PluralizationService.CreateService(CultureInfo.CurrentCulture);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ImportService"/> class.
        /// </summary>
        /// <param name="data">
        /// The data.
        /// </param>
        public ImportService(byte[] data)
            : this()
        {
            this.package = new ExcelPackage(new MemoryStream(data));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ImportService"/> class.
        /// </summary>
        /// <param name="path">
        /// The path.
        /// </param>
        public ImportService(string path)
            : this()
        {
            var data = File.ReadAllBytes(path);
            this.package = new ExcelPackage(new MemoryStream(data));
        }

        /// <summary>
        /// Prevents a default instance of the <see cref="ImportService"/> class from being created.
        /// </summary>
        private ImportService()
        {
            this.cachedTypes = new Dictionary<Type, CachedTypeInformation>();
        }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The dispose.
        /// </summary>
        public void Dispose()
        {
            try
            {
                this.package.Dispose();
            }
            catch
            {
            }
        }

        /// <summary>
        /// Gets all the types found in any tables in the entire worksheet.
        /// </summary>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="IEnumerable"/>.
        /// </returns>
        public IEnumerable<T> GetAll<T>()
        {
            var info = this.GetOrCreate<T>();

            var tables = (from worksheet in this.package.Workbook.Worksheets
                          from table in worksheet.Tables
                          where table.Name.Contains(info.PluralTypeName)
                          select table).ToList();

            return this.GetTypesFromTables<T>(tables, info).ToList();
        }

        /// <summary>
        /// Gets all the types in tables specified in the given sheet name.
        /// </summary>
        /// <typeparam name="T">
        /// The type of object.
        /// </typeparam>
        /// <param name="sheetName">
        /// The name of the sheet in the excel workbook.
        /// </param>
        /// <returns>
        /// The <see cref="IEnumerable{T}"/> found in all tabels of the given type in the given sheet.
        /// </returns>
        /// <exception cref="ArgumentNullException">
        /// The sheetName is null.
        /// </exception>
        public IEnumerable<T> GetListFromSheet<T>(string sheetName)
        {
            if (sheetName == null)
            {
                throw new ArgumentNullException("sheetName");
            }

            sheetName = sheetName.Replace(" ", "_");

            var info = this.GetOrCreate<T>();

            var tables = from sheet in this.package.Workbook.Worksheets
                         where sheet.Name == sheetName
                         from table in sheet.Tables
                         select table;

            return this.GetTypesFromTables<T>(tables, info);
        }

        /// <summary>
        /// e
        /// Gets all the types in the specified table name.
        /// </summary>
        /// <typeparam name="T">
        /// </typeparam>
        /// <param name="tableName">
        /// </param>
        /// <returns>
        /// The <see cref="IEnumerable{T}"/> found in the given table name.
        /// </returns>
        public IEnumerable<T> GetListFromTable<T>(string tableName)
        {
            if (tableName == null)
            {
                throw new ArgumentNullException("tableName");
            }

            tableName = tableName.Replace(" ", "_");

            var info = this.GetOrCreate<T>();
            var tableKey = string.Format(TableConfiguration<T>.TableKey, info.PluralTypeName);
            tableName = tableKey.Contains(tableName) ? tableKey : tableName + "_" + tableKey;

            var tables = from sheet in this.package.Workbook.Worksheets
                         from table in sheet.Tables
                         where table.Name.Contains(tableName)
                         select table;

            return this.GetTypesFromTables<T>(tables, info);
        }

        #endregion

        #region Methods

        /// <summary>
        /// The get or create.
        /// </summary>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="CachedTypeInformation"/>.
        /// </returns>
        private CachedTypeInformation GetOrCreate<T>()
        {
            Type type = typeof(T);
            CachedTypeInformation typeInformation;

            if (!this.cachedTypes.TryGetValue(type, out typeInformation))
            {
                typeInformation = CachedTypeInformation.From(type);
                this.cachedTypes.Add(type, typeInformation);
            }

            return typeInformation;
        }

        /// <summary>
        /// The get types from tables.
        /// </summary>
        /// <param name="tables">
        /// The tables.
        /// </param>
        /// <param name="info">
        /// The info.
        /// </param>
        /// <typeparam name="T">
        /// </typeparam>
        /// <returns>
        /// The <see cref="IEnumerable"/>.
        /// </returns>
        private IEnumerable<T> GetTypesFromTables<T>(IEnumerable<ExcelTable> tables, CachedTypeInformation info)
        {
            foreach (var excelTable in tables)
            {
                var dataTable = excelTable.ToDataTable();

                foreach (var row in dataTable.Rows.Cast<DataRow>())
                {
                    var instance = Activator.CreateInstance<T>();

                    foreach (var column in dataTable.Columns.Cast<DataColumn>())
                    {
                        var property = info.Properties.Single(p => p.Name == column.ColumnName);
                        var propertyType = property.PropertyType;
                        var propertyValue = row[column];
                        var typeConverter = TypeDescriptor.GetConverter(propertyType);
                        object value = typeConverter.ConvertFrom(propertyValue);
                        property.SetValue(instance, value);
                    }

                    yield return instance;
                }
            }
        }

        #endregion

        /// <summary>
        /// The cached type information.
        /// </summary>
        private class CachedTypeInformation
        {
            /// <summary>
            /// The cached type informations.
            /// </summary>
            private static Dictionary<Type, CachedTypeInformation> CachedTypeInformations;

            #region Constructors and Destructors

            /// <summary>
            /// Initializes a new instance of the <see cref="CachedTypeInformation"/> class.
            /// </summary>
            /// <param name="pluralTypeName">
            /// The plural type name.
            /// </param>
            /// <param name="properties">
            /// The properties.
            /// </param>
            /// <param name="typeName">
            /// The type name.
            /// </param>
            private CachedTypeInformation(string pluralTypeName, PropertyInfo[] properties, string typeName)
            {
                this.PluralTypeName = pluralTypeName;
                this.Properties = properties;
                this.TypeName = typeName;
            }

            /// <summary>
            /// Initializes static members of the <see cref="CachedTypeInformation"/> class.
            /// </summary>
            static CachedTypeInformation()
            {
                CachedTypeInformations = new Dictionary<Type, CachedTypeInformation>();
            }

            #endregion

            #region Public Properties

            /// <summary>
            /// Gets or sets the plural type name.
            /// </summary>
            internal string PluralTypeName { get; set; }

            /// <summary>
            /// Gets or sets the properties.
            /// </summary>
            internal PropertyInfo[] Properties { get; set; }

            /// <summary>
            /// Gets or sets the type name.
            /// </summary>
            internal string TypeName { get; set; }

            #endregion

            #region Public Methods and Operators

            /// <summary>
            /// Gets the cached type information object from the given type.
            /// </summary>
            /// <param name="type"></param>
            /// <returns></returns>
            public static CachedTypeInformation From(Type type)
            {
                CachedTypeInformation typeInformation;

                if (!CachedTypeInformations.TryGetValue(type, out typeInformation))
                {
                    typeInformation = new CachedTypeInformation(PluralizationService.Pluralize(type.Name), type.GetProperties(), type.Name);
                    CachedTypeInformations[type] = typeInformation;
                }

                return typeInformation;
            }

            #endregion
        }
    }
}