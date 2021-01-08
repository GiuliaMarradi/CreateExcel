using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
namespace CreateExcel
{
    public class NumberCell : Cell
    {
        public NumberCell(string header, long value, int index)
        {
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.CellValue = new CellValue(value.ToString());
        }

        public NumberCell(string header, int value, int index)
        {
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.CellValue = new CellValue(value);
        }

        public NumberCell(string header, double value, int index)
        {
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.CellValue = new CellValue(value);
        }

        public NumberCell(string header, decimal value, int index)
        {
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.CellValue = new CellValue(value);
        }
    }
}
