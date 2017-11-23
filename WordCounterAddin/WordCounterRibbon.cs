using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;

namespace WordCounterAddin
{
    public partial class WordCounterRibbon
    {
        private void WordCounterRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            StatisticsForm form = new StatisticsForm(Globals.ThisAddIn.Application.ActiveDocument);
            form.Show();
        }
    }
}
