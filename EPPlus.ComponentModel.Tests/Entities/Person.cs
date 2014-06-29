namespace EPPlus.ComponentModel.Tests.Entities
{
    using System;

    /// <summary>
    /// The person.
    /// </summary>
    public class Person
    {
        #region Public Properties

        /// <summary>
        /// Gets or sets the dob.
        /// </summary>
        public DateTime Dob { get; set; }

        /// <summary>
        /// Gets or sets the first name.
        /// </summary>
        public string FirstName { get; set; }

        /// <summary>
        /// Gets or sets the id.
        /// </summary>
        public Guid Id { get; set; }

        /// <summary>
        /// Gets or sets the last name.
        /// </summary>
        public string LastName { get; set; }

        /// <summary>
        /// Gets or sets the middle name.
        /// </summary>
        public string MiddleName { get; set; }

        #endregion
    }
}