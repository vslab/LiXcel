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
            Excel.Range selection =
                Application.Selection != null &&
                Application.Selection is Excel.Range &&
                ((Excel.Range) Application.Selection).Columns.Count > 1?
                Application.Selection : null;
            var inputBox = Application.InputBox("seleziona casella input",Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,8);
            
            if (!(inputBox is Excel.Range))
                return;

            Excel.Range input = inputBox;
            Excel.Range output = selection;
            if (selection == null)
            {
                inputBox = Application.InputBox("seleziona range output (almeno 2 colonne)", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (!(
                    inputBox is Excel.Range
                    && ((inputBox as Excel.Range).Columns.Count > 1)
                    ))
                    return;
            }
            int iterazioni = 50000;// Application.InputBox("Numero di iterazioni", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            Globals.api.Simulate(input, output, iterazioni);
        }
    }
}
