using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace SimpleExcelExport
{
    public class ExportToExcel
    {
        public static byte[] ListToExcel<T>(List<T> list)
        {
            MemoryStream output=new MemoryStream();
            var columns=GetTypeDefinition(typeof(T));
            var excelCreator = new ExcelFileCreator(columns);
            ProcessRows<T>(list,columns,excelCreator);
            output = (MemoryStream)excelCreator.SaveDocument();
            return output.ToArray();
        }

        private static void ProcessRows<T>(List<T> list,List<Column> columns,ExcelFileCreator excel)
        {
            var orderedColumns = columns.OrderBy(x => x.ColumnOrder);
            Type type = typeof(T);
            int columnNumber = 0;
            foreach(var element in list)
            {
                excel.CreateRow();
                columnNumber = 0;
                foreach(var column in orderedColumns)
                {
                    var value = type.GetProperty(column.PropName).GetValue(element, null);
                    excel.CreateCellWithValue(columnNumber, value, column.PropType);
                    ++columnNumber;
                }
                
            }
        }

        private static List<Column> GetTypeDefinition(Type type)
        {
            List<Column> columns = new List<Column>();
            foreach (var prop in type.GetProperties())
            {
                var tmp=new Column();
                var attrs = System.Attribute.GetCustomAttributes(prop);
                tmp.PropName = prop.Name;
                tmp.PropType = prop.PropertyType;
                tmp.ColumnName = prop.Name;
                tmp.ColumnOrder = 0;
                foreach (var attr in attrs)
                {
                    if (attr is ExcelExport)
                    {
                        ExcelExport attribute = (ExcelExport)attr;
                        tmp.ColumnName = attribute.GetName();
                        tmp.ColumnOrder = attribute.order;
                    }
                }
                columns.Add(tmp);
            }
            return columns;
        }
    
    }
}
