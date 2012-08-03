// -----------------------------------------------------------------------
// <copyright file="Column.cs" company="Microsoft">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace SimpleExcelExport
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    internal class Column
    {
        public string ColumnName { get; set; }
        public string PropName { get; set; }
        public Type PropType { get; set; }
        public int ColumnOrder { get; set; }
        public string CellColor { get; set; }
        public bool Ignore { get; set; }
        public Column()
        {
            CellColor = string.Empty;
            Ignore = false;
        }
    }
    
}
