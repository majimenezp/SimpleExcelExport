using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using System.IO;
namespace SimpleExcelExport
{
    public class ExcelFileCreator
    {
        private List<Column> columns;
        private HSSFWorkbook document;
        private HSSFSheet currentSheet;
        private int currentRowNumber=0;
        private NPOI.SS.UserModel.IRow currentRow;
        public ExcelFileCreator()
        {
            this.columns = new List<Column>();
            CreateDocument();
        }

        public ExcelFileCreator(List<Column> columns)
        {
            CreateDocument();
            this.columns = columns;
            CreateHeader();
        }

        private void CreateDocument()
        {
            document = new HSSFWorkbook();
            currentSheet = (HSSFSheet)document.CreateSheet();
        }
        public Stream SaveDocument()
        {
            MemoryStream memory = new MemoryStream();
            document.Write(memory);
            document.Dispose();
            return memory;
        }

        private void CreateHeader()
        {
            int columnNumber=0;
            var orderedColumns = columns.OrderBy(x => x.ColumnOrder);
            var row = currentSheet.CreateRow(currentRowNumber);
            foreach (var column in orderedColumns)
            {
                row.CreateCell(columnNumber, NPOI.SS.UserModel.CellType.STRING );
                
                var cellt = GetColumnCellType(column.PropType);
                row.Cells[columnNumber].SetCellValue(column.ColumnName);
                currentSheet.SetColumnWidth(columnNumber, (int)((column.ColumnName.Length*1.5) * 256));
                ++columnNumber;
            }
        }

        private NPOI.SS.UserModel.CellType GetColumnCellType(Type type)
        {
            NPOI.SS.UserModel.CellType cellType= NPOI.SS.UserModel.CellType.STRING;
            switch (type.Name.ToLowerInvariant())
            {
                case "string":
                    cellType=NPOI.SS.UserModel.CellType.STRING;
                    break;
                case "datetime":
                    cellType = NPOI.SS.UserModel.CellType.STRING;
                    break;
                case "int":
                case "int32":
                case "int64":
                case "decimal":
                case "long":
                case "double":
                    cellType = NPOI.SS.UserModel.CellType.NUMERIC;
                    break;
                case "boolean":
                case "bool":
                    cellType = NPOI.SS.UserModel.CellType.BOOLEAN;
                    break;
            }
            return cellType;
        }


        internal void CreateRow()
        {
            ++currentRowNumber;
            currentRow=currentSheet.CreateRow(currentRowNumber);
            
        }

        internal void CreateCellWithValue(int i, object value,Type valueType)
        {
            var cell=currentRow.CreateCell(i, GetColumnCellType(valueType));
            switch (valueType.Name.ToLowerInvariant())
            {
                case "string":
                    cell.SetCellValue(value.ToString());
                    break;
                case "datetime":
                    cell.SetCellValue((DateTime)value);
                    break;
                case "int":
                case "int32":
                case "int64":
                case "decimal":
                case "long":
                case "double":
                    cell.SetCellValue(Convert.ToDouble(value));
                    break;
                case "boolean":
                case "bool":
                    cell.SetCellValue((bool)value);
                    break;
                default:
                    cell.SetCellValue(value.ToString());
                    break;
            }
            
        }
    }
}
