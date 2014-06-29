// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Order.cs" company="">
//   
// </copyright>
// <summary>
//   The order.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace EPPlus.ComponentModel.Tests.Entities
{
    using System;

    /// <summary>
    /// The order.
    /// </summary>
    public class Order
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the cost.
        /// </summary>
        public decimal Cost { get; set; }

        /// <summary>
        /// Gets or sets the date created.
        /// </summary>
        public DateTime DateCreated { get; set; }

        /// <summary>
        /// Gets or sets the id.
        /// </summary>
        public Guid Id { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether is processed.
        /// </summary>
        public bool IsProcessed { get; set; }

        /// <summary>
        /// Gets or sets the items.
        /// </summary>
        public int Items { get; set; }

        /// <summary>
        /// Gets or sets the reference.
        /// </summary>
        public string Reference { get; set; }

        #endregion
    }
}