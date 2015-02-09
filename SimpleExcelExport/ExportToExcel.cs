using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;

namespace SimpleExcelExport
{
    public class ExportToExcel
    {
        private ExcelFileCreator globalExcelCreator;
        private int lastRowNumber=0;
        public static Regex regexFunctionColor = new Regex(
        "(?:{)(.*)(?:})",
        RegexOptions.IgnoreCase
        | RegexOptions.CultureInvariant
        | RegexOptions.IgnorePatternWhitespace
        | RegexOptions.Compiled);
        public static Regex regexRgbColor = new Regex("(\\d{1,3}),(\\d{1,3}),(\\d{1,3})", RegexOptions.IgnoreCase| RegexOptions.CultureInvariant| RegexOptions.IgnorePatternWhitespace| RegexOptions.Compiled);
        private Column initColumn;

        public int LastRowNumber { get { return lastRowNumber; } }

        public ExportToExcel()
        {
        }

        public ExportToExcel(bool headerBold, string headerFontColor="", string headerbackgroundColor="")
        {
            initColumn = new Column();
            initColumn.HFontBold = headerBold;
            initColumn.HFontColor = headerFontColor;
            initColumn.HBackColor = headerbackgroundColor;
        }

        public byte[] ListToExcel<T>(List<T> list)
        {
            ExcelFileCreator excelCreator;
            MemoryStream output = new MemoryStream();
            var columns = GetTypeDefinition(typeof(T));
            try
            {
                if (initColumn == null)
                {
                    excelCreator = new ExcelFileCreator(columns);
                }
                else
                {
                    foreach (var item in columns)
                    {
                        item.HFontBold = initColumn.HFontBold;
                        item.HFontColor = initColumn.HFontColor;
                        item.HBackColor = initColumn.HBackColor;
                    }
                    excelCreator = new ExcelFileCreator(columns);
                }
            }
            catch
            {
                excelCreator = new ExcelFileCreator(columns);
            }
            
            ProcessRows<T>(list, columns, excelCreator);
            lastRowNumber = excelCreator.LastRownNumber;
            output = (MemoryStream)excelCreator.SaveDocument();
            return output.ToArray();
        }

        public HSSFWorkbook ProcessListToExcel<T>(List<T> list)
        {
            var columns = GetTypeDefinition(typeof(T));
            try
            {
                globalExcelCreator = new ExcelFileCreator(columns);
            }
            catch
            {
                globalExcelCreator = new ExcelFileCreator(columns);
            }
            ProcessRows<T>(list, columns, globalExcelCreator);
            lastRowNumber = globalExcelCreator.LastRownNumber;
            return globalExcelCreator.GetDocument();
        }


        private void ProcessRows<T>(List<T> list, List<Column> columns, ExcelFileCreator excel)
        {
            var orderedColumns = columns.OrderBy(x => x.ColumnOrder);
            Type type = typeof(T);
            int columnNumber = 0;
            foreach (var element in list)
            {
                excel.CreateRow();
                columnNumber = 0;
                foreach (var column in orderedColumns)
                {
                    var value = type.GetProperty(column.PropName).GetValue(element, null);
                    System.Drawing.Color backgroundColor = GetColor(element, column.CellColor, type);

                    excel.CreateCellWithValue(columnNumber, value, column.PropType, backgroundColor);
                    ++columnNumber;
                }

            }
        }

        internal System.Drawing.Color GetColor(object element, string color, Type type)
        {
            System.Drawing.Color resultColor = System.Drawing.Color.Empty;
            if (!string.IsNullOrEmpty(color))
            {
                Match functionColor = regexFunctionColor.Match(color);

                if (regexFunctionColor.IsMatch(color)) // have a reference to a function,execute function
                {
                    var functionName = regexFunctionColor.Split(color)[1];
                    var methodInfo=type.GetMethod(functionName);
                    var value = (string)methodInfo.Invoke(element, null);
                    if (regexRgbColor.IsMatch(value))
                    {
                        resultColor = ProcessColorByRGB(value);
                    }
                    else if (string.IsNullOrEmpty(value))
                    {
                        resultColor = System.Drawing.Color.Empty;
                    }
                    else
                    {
                        resultColor = System.Drawing.Color.FromName(value);
                    }
                }
                else if (regexRgbColor.IsMatch(color))
                {
                    resultColor = ProcessColorByRGB(color);
                }
                else
                {
                    resultColor = System.Drawing.Color.FromName(color);
                }
                return resultColor;
            }
            else
            {
                return System.Drawing.Color.Empty;
            }
        }

        private System.Drawing.Color ProcessColorByRGB(string value)
        {
            System.Drawing.Color resultColor = System.Drawing.Color.Empty;
            int[] vals = new int[3];
            int i = 0;
            foreach (var elem in regexRgbColor.Matches(value))
            {
                vals[i] = Convert.ToInt32(elem);
                ++i;
            }
            resultColor = System.Drawing.Color.FromArgb(vals[0], vals[1], vals[2]);
            return resultColor;
        }

        private List<Column> GetTypeDefinition(Type type)
        {
            List<Column> columns = new List<Column>();
            foreach (var prop in type.GetProperties())
            {
                var tmp = new Column();
                var attrs = System.Attribute.GetCustomAttributes(prop);
                tmp.PropName = prop.Name;
                tmp.PropType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                tmp.ColumnName = prop.Name;
                tmp.ColumnOrder = 0;
                foreach (var attr in attrs)
                {
                    if (attr is ExcelExport)
                    {
                        ExcelExport attribute = (ExcelExport)attr;
                        tmp.ColumnName = attribute.GetName();
                        tmp.ColumnOrder = attribute.order;
                        tmp.CellColor = attribute.GetBackgroundColor();
                        tmp.HFontBold = attribute.GetHeaderBold();
                        tmp.HBackColor = attribute.GetHeaderBackgroundColor();
                        tmp.HFontColor = attribute.GetHeaderFontColor();
                        tmp.Ignore=attribute.ignore;
                    }
                }
                if(!tmp.Ignore)
                {
                    columns.Add(tmp);
                }
            }
            return columns;
        }

    }
}
