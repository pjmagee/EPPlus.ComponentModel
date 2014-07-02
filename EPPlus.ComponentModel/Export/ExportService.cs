// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ExportService.cs" company="Patrick Magee">
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
//   The ExportService.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Export
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using EPPlus.ComponentModel.Exceptions;

    using OfficeOpenXml;

    /// <summary>
    /// The ExportService.
    /// </summary>
    public class ExportService : IExportService, IDisposable
    {
        #region Fields

        /// <summary>
        /// The package.
        /// </summary>
        private readonly ExcelPackage package;

        /// <summary>
        /// The worksheet configurations.
        /// </summary>
        private readonly List<IWorksheetConfiguration> worksheetConfigurations;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ExportService"/> class.
        /// </summary>
        public ExportService()
        {
            this.package = new ExcelPackage();
            this.worksheetConfigurations = new List<IWorksheetConfiguration>();
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets the worksheet configurations.
        /// </summary>
        public IEnumerable<IWorksheetConfiguration> WorksheetConfigurations
        {
            get
            {
                return this.worksheetConfigurations;
            }
        }

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
        /// <exception cref="Exception">
        /// </exception>
        public IWorksheetConfiguration AddSheetForExport(string sheetName)
        {
            if (this.package.Workbook.Worksheets.Any(ws => ws.Name == sheetName))
            {
                throw new SheetNameExistsException(string.Format("Sheet {0} already exists", sheetName));
            }

            var worksheet = this.package.Workbook.Worksheets.Add(sheetName);
            var worksheetConfiguration = new WorksheetConfiguration(this, worksheet);
            this.worksheetConfigurations.Add(worksheetConfiguration);
            return worksheetConfiguration;
        }

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
        /// The export.
        /// </summary>
        /// <returns>
        /// The <see cref="byte[]"/>.
        /// </returns>
        public byte[] Export()
        {
            this.AddValidations();
            this.AutoFormatColumns();
            return this.package.GetAsByteArray();
        }

        /// <summary>
        /// The export.
        /// </summary>
        /// <param name="path">
        /// The path.
        /// </param>
        public void Export(string path)
        {
            this.AddValidations();
            this.AutoFormatColumns();
            this.package.SaveAs(new FileInfo(path));
        }

        #endregion

        #region Methods

        /// <summary>
        /// Auto format the worksheets.
        /// </summary>
        private void AutoFormatColumns()
        {
            foreach (var workSheet in this.package.Workbook.Worksheets)
            {
                if (workSheet.Dimension != null)
                {
                    workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();    
                }
            }
        }

        /// <summary>
        /// The add validations.
        /// </summary>
        private void AddValidations()
        {
            foreach (var worksheetConfiguration in this.worksheetConfigurations)
            {
                foreach (var tableConfiguration in worksheetConfiguration.TableConfigurations)
                {
                    var objectConfiguration = tableConfiguration.Options;

                    foreach (var configuration in objectConfiguration.DateValidations)
                    {
                        var process = configuration.Value.Item1;
                        var validator = configuration.Value.Item2;
                        process(validator);
                    }

                    foreach (var configuration in objectConfiguration.IntegerValidators)
                    {
                        var process = configuration.Value.Item1;
                        var validator = configuration.Value.Item2;
                        process(validator);
                    }

                    foreach (var configuration in objectConfiguration.ListValidations)
                    {
                        var process = configuration.Value.Item1;
                        var validator = configuration.Value.Item2;
                        process(validator);
                    }

                    foreach (var configuration in objectConfiguration.DecimalValidators)
                    {
                        var process = configuration.Value.Item1;
                        var validator = configuration.Value.Item2;
                        process(validator);
                    }
                }
            }
        }

        #endregion
    }
}