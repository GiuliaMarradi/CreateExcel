using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CreateExcel
{
    public static partial class CellUtils
    {
        public static Cell CreateHeaderCell(string header, string text, int index, Stylesheet styles,
            System.Drawing.Color fillColour, double? fontSize, bool isBold)
        {
            var cell = CreateTextCell(header, text, index);

            UInt32Value fontId = CreateFont(styles, "", fontSize, isBold, System.Drawing.Color.Black);
            UInt32Value fillId = CreateFill(styles, fillColour);
            UInt32Value formatId = CreateCellFormat(styles, fontId, fillId, 0);

            cell.StyleIndex = formatId;

            return cell;
        }


        private static UInt32Value CreateCellFormat(
            Stylesheet styleSheet,
            UInt32Value fontIndex,
            UInt32Value fillIndex,
            UInt32Value numberFormatId)
        {
            CellFormat cellFormat = new CellFormat();

            if (fontIndex != null)
                cellFormat.FontId = fontIndex;

            if (fillIndex != null)
                cellFormat.FillId = fillIndex;

            if (numberFormatId != null)
            {
                cellFormat.NumberFormatId = numberFormatId;
                cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            }

            styleSheet.CellFormats.Append(cellFormat);

            UInt32Value result = styleSheet.CellFormats.Count;
            styleSheet.CellFormats.Count++;
            return result;
        }

        private static UInt32Value CreateFill(
            Stylesheet styleSheet,
            System.Drawing.Color fillColor)
        {
            var color = System.Drawing.Color.FromArgb(
                                   fillColor.A,
                                   fillColor.R,
                                   fillColor.G,
                                   fillColor.B);

            PatternFill patternFill =
                new PatternFill(
                    new ForegroundColor()
                    {
                        Rgb = new HexBinaryValue()
                        {
                            Value = color.Name
                        }
                    });

            patternFill.PatternType = fillColor ==
                        System.Drawing.Color.White ? PatternValues.None : PatternValues.LightDown;

            Fill fill = new Fill(patternFill);

            styleSheet.Fills.Append(fill);

            UInt32Value result = styleSheet.Fills.Count;
            styleSheet.Fills.Count++;
            return result;
        }

        private static UInt32Value CreateFont(
            Stylesheet styleSheet,
            string fontName,
            double? fontSize,
            bool isBold,
            System.Drawing.Color foreColor)
        {

            Font font = new Font();

            if (!string.IsNullOrEmpty(fontName))
            {
                FontName name = new FontName()
                {
                    Val = fontName
                };
                font.Append(name);
            }

            if (fontSize.HasValue)
            {
                FontSize size = new FontSize()
                {
                    Val = fontSize.Value
                };
                font.Append(size);
            }

            if (isBold == true)
            {
                Bold bold = new Bold();
                font.Append(bold);
            }

            var colorName = System.Drawing.Color.FromArgb(
                                foreColor.A,
                                foreColor.R,
                                foreColor.G,
                                foreColor.B).Name;

            Color color = new Color()
            {
                Rgb = new HexBinaryValue()
                {
                    Value = colorName
                }
            };
            font.Append(color);

            styleSheet.Fonts.Append(font);
            UInt32Value result = styleSheet.Fonts.Count;
            styleSheet.Fonts.Count++;
            return result;
        }

    }
}
