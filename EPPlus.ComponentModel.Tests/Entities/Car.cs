namespace EPPlus.ComponentModel.Tests.Entities
{
    using System;

    /// <summary>
    /// The car.
    /// </summary>
    public class Car
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the id.
        /// </summary>
        public Guid Id { get; set; }

        /// <summary>
        /// Gets or sets the make.
        /// </summary>
        public string Make { get; set; }

        /// <summary>
        /// Gets or sets the model.
        /// </summary>
        public string Model { get; set; }

        #endregion
    }
}