// -----------------------------------------------------------------------
// <copyright file="ColumnName.cs" company="Microsoft">
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
    /// Attribute to set the column name in excel and the columns order
    /// </summary>
    /// 
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class ExcelExport:System.Attribute
    {
        private string name;
        public int order;
        public ExcelExport(string name)
        {
            this.name = name;
            this.order = 0;
        }
        public string GetName()
        {
            return name;
        }
    }
}
