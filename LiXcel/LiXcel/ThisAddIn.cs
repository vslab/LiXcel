using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace LiXcel
{
    internal sealed partial class Globals {
        static public LiXcelCore.Api api { get { return Globals.ThisAddIn.api; }  }
    }
    public partial class ThisAddIn
    {
        private LiXcelCore.Api _api;
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;

        #region properties
        public Microsoft.Office.Tools.CustomTaskPane taskPane
        {
            get
            {
                return _taskPane;
            }
        }
        public LiXcelCore.Api api
        {
            get
            {
                return _api;
            }
        }
        #endregion

        public void ShowTaskPane()
        {
            _taskPane.Visible = true;
        }

        protected override object RequestComAddInAutomationService()
        {
            lock (this)
            {
                if (api == null)
                    _api = new LiXcelCore.Api();
            }

            return api;
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var taskPaneView = new TaskPaneView();
            _taskPane = this.CustomTaskPanes.Add(taskPaneView, "LiXcel");
            _taskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Codice generato da VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
