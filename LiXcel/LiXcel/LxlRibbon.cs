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
                Application.Selection is Excel.Range ?
                Application.Selection : null;
            
            Excel.Range input = Application.InputBox("seleziona casella input",Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,8);
            Excel.Range output = selection;// Application.InputBox("seleziona casella output",Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,8);
            int iterazioni = 50000;// Application.InputBox("Numero di iterazioni", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            Globals.api.Simulate(input, output, iterazioni);
        }
    }
}
