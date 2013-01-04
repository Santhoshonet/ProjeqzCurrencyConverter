using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;

namespace ProjeqzCurrencyConverter
{
    public partial class ThisAddIn
    {
        private static void ThisAddInStartup(object sender, EventArgs e)
        {

        }

        private static void ThisAddInShutdown(object sender, EventArgs e)
        {
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {

            return base.CreateRibbonExtensibilityObject();
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }
        
        #endregion
    }
}
