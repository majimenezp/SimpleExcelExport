// -----------------------------------------------------------------------
// <copyright file="Person.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace Tests
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using SimpleExcelExport;
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class Person
    {
        /// <summary>
        /// No needed,but in case you need to set a column name and columns order
        /// </summary>
        [ExcelExport("Name",order=1)]
        public string Name { get; set; }

        [ExcelExport("Last Name", order = 2)]
        public string LastName { get; set; }

        [ExcelExport("day of birth", order = 3)]
        public DateTime BirthDay { get; set; }

        [ExcelExport("Country", order = 4)]
        public string Country { get; set; }

        [ExcelExport("Genre", order = 5)]
        public Sex Sex { get; set; }

        [ExcelExport("Number of children", order = 7)]
        public int NumberOfChildren { get; set; }

        [ExcelExport("Person's height", order = 6)]
        public decimal Height { get; set; }
    }
}
