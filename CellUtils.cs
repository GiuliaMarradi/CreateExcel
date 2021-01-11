using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace CreateExcel
{
    public static partial class CellUtils
    {
        public static Cell CreateDateCell(string header, DateTime dateTime, int index)
            => new Cell
            {
                DataType = CellValues.Date,
                CellReference = header + index,
                StyleIndex = 1,
                CellValue = new CellValue(dateTime)
            };

        public static Cell CreateFomulaCell(string header, string text, int index)
            => new Cell
            {
                CellFormula = new CellFormula { CalculateCell = true, Text = text },
                DataType = CellValues.Number,
                CellReference = header + index,
                StyleIndex = 2
            };

        public static Cell CreateTextCell(string header, string text, int index)
            => new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = header + index,
                //Add text to the text cell.
                InlineString = new InlineString { Text = new Text { Text = text } }
            };

        public static Cell CreateFormattedNumberCell(string header, double value, int index)
        {
            var cell = CreateNumberCell(header, value, index);
            cell.StyleIndex = 2;
            return cell;
        }

        public static Cell CreateFormattedNumberCell(string header, decimal value, int index)
        {
            var cell = CreateNumberCell(header, value, index);
            cell.StyleIndex = 2;
            return cell;
        }

        public static Cell CreateNumberCell(string header, int value, int index)
            => new Cell
            {
                DataType = CellValues.Number,
                CellReference = header + index,
                CellValue = new CellValue(value),
            };

        public static Cell CreateNumberCell(string header, double value, int index)
            => new Cell
            {
                DataType = CellValues.Number,
                CellReference = header + index,
                CellValue = new CellValue(value),
            };

        public static Cell CreateNumberCell(string header, decimal value, int index)
            => new Cell
            {
                DataType = CellValues.Number,
                CellReference = header + index,
                CellValue = new CellValue(value),
            };
    }
}
