using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

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
                input = inputBox;
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
                output = inputBox;
            }
            int iterazioni = 1000000;// Application.InputBox("Numero di iterazioni", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            Globals.api.Simulate(input, output, iterazioni);
        }
    }
}
