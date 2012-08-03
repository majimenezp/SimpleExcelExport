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
        public string backgroundColor;
        public int order;
        public bool ignore;
        public ExcelExport(string name)
        {
            this.name = name;
            this.order = 0;
            this.backgroundColor = string.Empty;
            this.ignore = false;
        }
        public string GetName()
        {
            return name;
        }
        public string GetBackgroundColor()
        {
            return backgroundColor;
        }
        public bool GetIgnore()
        {
            return ignore;
        }
    }
}
