using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
namespace LiXcel
{
    public partial class LxlRibbon
    {
        private Excel.Application Application { get {return Globals.ThisAddIn.Application; }  }

        private void LxlRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range input = Application.ActiveCell;
            if (input == null)
            {
                var inputBox = Application.InputBox("seleziona casella input", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (!(inputBox is Excel.Range))
                    return;
                input = (Excel.Range) inputBox;
            }

            Excel.Range output = null;
            if (output == null)
            {
                var inputBox = Application.InputBox("seleziona range output (almeno 2 colonne)", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (!(
                    inputBox is Excel.Range
                    && ((inputBox as Excel.Range).Columns.Count > 1)
                    ))
                    return;
                output = (Excel.Range)inputBox;
            }
            double min = double.NaN;
            var minstr = Application.InputBox("Minimum value");
            if (minstr is bool  || "".Equals(minstr.ToString())|| ! double.TryParse(minstr.ToString(), out min)) min = double.NaN;
            double max = double.NaN;
            var maxstr = Application.InputBox("Maxumum value");
            if (maxstr is bool  || "".Equals(maxstr.ToString())|| ! double.TryParse(maxstr.ToString(), out max)) max = double.NaN;
            int iterazioni = 1000000;// Application.InputBox("Numero di iterazioni", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            Globals.api.Simulate(input, output, iterazioni,min,max);
        }
    }
}
