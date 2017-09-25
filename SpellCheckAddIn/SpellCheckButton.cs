using ESRI.ArcGIS.ArcMapUI;
using ArcMapSpellCheck;
using System.Diagnostics;
using System.Windows.Forms;

namespace SpellCheckAddIn
{
    public class SpellCheckButton : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public SpellCheckButton()
        {
        }

        protected override void OnClick()
        {
            // Check to see if Word is already opened.
            var processes = Process.GetProcessesByName("WINWORD");

            if (processes.Length > 0)
            {
                string message = "The spell check operation needs to use Microsoft Word, which is already opened. Please close Word and try again.";
                MessageBox.Show(message, "Word is already open", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, 0);
            }
            else
            {
                IMxDocument doc = ArcMap.Application.Document as IMxDocument;
                using (var spellchecker = new Spellchecker())
                {
                    spellchecker.CheckDocument(doc);
                }
            }
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
