using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace LiXcel
{
    public partial class TaskPaneView : UserControl
    {
        private Excel.Application Application { get { return Globals.ThisAddIn.Application; } }

        public TaskPaneView()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (keepOnSimulating)
            {
                // stop simulating
                keepOnSimulating = false;
            }
            else
            {
                if (stillSimulating)
                {
                    // finishing last simulation
                    SetText("completing simulation, please wait", StatusLabel);
                }
                else
                {
                    // start simulating
                    iterations = (int)iterationsNumericUpDown.Value;
                    Excel.Range input = Application.ActiveCell;
                    if (input == null)
                    {
                        var inputBox = Application.InputBox("seleziona casella input", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                        if (!(inputBox is Excel.Range))
                            return;
                        input = (Excel.Range)inputBox;
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
                    var minstr = minTextBox.Text; //Application.InputBox("Minimum value");
                    if (minstr is bool || "".Equals(minstr.ToString()) || !double.TryParse(minstr.ToString(), out min)) min = double.NaN;
                    double max = double.NaN;
                    var maxstr = maxTextBox.Text; //Application.InputBox("Maxumum value");
                    if (maxstr is bool || "".Equals(maxstr.ToString()) || !double.TryParse(maxstr.ToString(), out max)) max = double.NaN;
                    //int iterazioni = 1000000;// Application.InputBox("Numero di iterazioni", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                    load = Globals.api.SimulateThreaded(input, output, iterations, min, max);
                    System.Threading.Thread t = new System.Threading.Thread(RunSimulation);
                    t.Start();
                }
            }
        }

        private System.Threading.Thread currentThread;
        private Microsoft.FSharp.Core.FSharpFunc<Microsoft.FSharp.Core.Unit, Microsoft.FSharp.Core.Unit> load;
        private bool keepOnSimulating;
        private bool stillSimulating;
        private int iterations;

        private void RunSimulation()
        {
            keepOnSimulating = true;
            stillSimulating = true;
            SetText("Stop Simulation", StartButton);
            int totiterations = 0;
            while (keepOnSimulating)
            {
                totiterations += iterations;
                SetText("simulated " + totiterations +  " samples", StatusLabel);
                load.Invoke(null);
            }
            SetText("Start Simulation", StartButton);
            SetText("simulated " + totiterations + " samples", StatusLabel);
            stillSimulating = false;


        }

        delegate void SetTextCallback(string text, System.Windows.Forms.Control c);
        private void SetText(string text, System.Windows.Forms.Control c)
        {
            if (c.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text, c });
            }
            else
            {
                c.Text = text;
            }
        }

    }
}
