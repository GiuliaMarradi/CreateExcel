using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CreateExcel
{
    public class FormatedNumberCell : NumberCell
    {
        public FormatedNumberCell(string header, double value, int index)
            : base(header, value, index)
        {
            this.StyleIndex = 2;
        }

        public FormatedNumberCell(string header, decimal value, int index)
          : base(header, value, index)
        {
            this.StyleIndex = 2;
        }
    }
}
