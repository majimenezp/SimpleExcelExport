using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using System.IO;
namespace SimpleExcelExport
{
    internal class ExcelFileCreator
    {
        private List<Column> columns;
        private HSSFWorkbook document;
        private HSSFSheet currentSheet;
        private int currentRowNumber=0;
        private HSSFDataFormat cellsFormat;
        private NPOI.SS.UserModel.IRow currentRow;
        private short DateFormat;
        private short DecimalFormat;
        private Dictionary<string, HSSFCellStyle> cellStyles=new Dictionary<string,HSSFCellStyle>();
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
            cellsFormat = (HSSFDataFormat)document.CreateDataFormat();
            DateFormat = 14;
            currentSheet = (HSSFSheet)document.CreateSheet();
        }
        public Stream SaveDocument()
        {
            MemoryStream memory = new MemoryStream();
            document.Write(memory);
            return memory;
        }

        public HSSFWorkbook GetDocument()
        {
            return document;
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
                    cellType = NPOI.SS.UserModel.CellType.NUMERIC;
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

        public int LastRownNumber { get { return currentRowNumber; } }

        internal void CreateCellWithValue(int i, object value,Type valueType,System.Drawing.Color backgroundColor)
        {
            HSSFCellStyle style;
            string valueTypeName=valueType.Name.ToLowerInvariant();
            var cell=currentRow.CreateCell(i, GetColumnCellType(valueType));
            string styleId=valueTypeName + (backgroundColor.IsEmpty?string.Empty:backgroundColor.ToArgb().ToString());
            if(cellStyles.ContainsKey(styleId))
            {
                style=cellStyles[styleId];
            }
            else{
                style= (HSSFCellStyle)document.CreateCellStyle();
                if (!backgroundColor.IsEmpty)
                {
                    style.FillForegroundColor = GetXLColour(backgroundColor);
                    style.FillPattern = NPOI.SS.UserModel.FillPatternType.SOLID_FOREGROUND;
                }
                switch (valueTypeName)
                {
                    case "string":
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("General");
                        break;
                    case "datetime":
                        style.DataFormat = 14;
                        break;
                    case "int":
                    case "int32":
                    case "int64":
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0");
                        break;
                    case "decimal":
                    case "long":
                    case "double":
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                        break;
                    default:
                        style.DataFormat = HSSFDataFormat.GetBuiltinFormat("General");
                        break;
                }
                cellStyles.Add(styleId, style);
            }
            
            
            switch (valueTypeName)
            {
                case "string":
                    cell.SetCellValue(value.ToString());
                    break;
                case "datetime":
                    if (((DateTime)value) == DateTime.MinValue)
                    {
                        cell.SetCellValue(string.Empty);
                    }
                    else
                    {
                        cell.SetCellValue((DateTime)value);
                    }
                    break;
                case "int":
                case "int32":
                case "int64":
                    cell.SetCellValue(Convert.ToDouble(value));
                    break;
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
            cell.CellStyle = style;
        }

        private short GetXLColour(System.Drawing.Color SystemColour)
        {
            HSSFPalette XlPalette = document.GetCustomPalette();
            NPOI.HSSF.Util.HSSFColor XlColour = XlPalette.FindColor(SystemColour.R, SystemColour.G, SystemColour.B);
            if (XlColour == null)
            {
                if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 255)
                {
                    if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 64)
                    {
                        NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE = 64;
                    }
                    NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE += 1;
                    XlColour = XlPalette.AddColor(SystemColour.R, SystemColour.G, SystemColour.B);
                }
                else
                {
                    XlColour = XlPalette.FindSimilarColor(SystemColour.R, SystemColour.G, SystemColour.B);
                }
                return XlColour.GetIndex();
            }
            else
            {
                
                return XlColour.GetIndex();
            }
        }

    }
}
