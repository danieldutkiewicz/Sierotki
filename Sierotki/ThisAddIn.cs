using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;

namespace Sierotki
{
    public partial class ThisAddIn
    {

        Word.Range range;
        readonly string[] characters = { " w ", " i ", " u ", " o ", " a ", " z " };
        readonly string[] charactersReplacement = { " w^s", " i^s", " u^s", " o^s", " a^s", " z^s" };

        public void SearchReplace()
        {

            Word.Document document = Application.ActiveDocument;

            object start = document.Content.Start;
            object end = document.Content.End;

            range = document.Range(start, end);

            int i = 0;
            foreach (var character in characters)
            {
                range.Find.ClearFormatting();
                object text = range.Find.Text = character;
                range.Find.Replacement.ClearFormatting();
                object replaceWith = range.Find.Replacement.Text = charactersReplacement[i++];

                Replace(text, replaceWith);
            }

            ShowEndWindow();
        }

        private void Replace(object text, object replaceWith)
        {
            object replaceAll = Word.WdReplace.wdReplaceAll;
            range.Find.Execute(ref text, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref replaceWith,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
        }

        private void ShowEndWindow()
        {
            MessageBox.Show("Program zakończony sukcesem.");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Kod wygenerowany przez program VSTO

        /// <summary>
        /// Metoda wymagana do obsługi projektanta — nie należy modyfikować
        /// jej zawartości w edytorze kodu.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
