using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ESRI.ArcGIS.ArcMapUI;
using ArcMapSpellCheck;

namespace SpellCheckAddIn
{
    public class SpellCheckButton : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public SpellCheckButton()
        {
        }

        protected override void OnClick()
        {
            //
            //  TODO: Sample code showing how to access button host
            //
            ////ArcMap.Application.CurrentTool = null;
            IMxDocument doc = ArcMap.Application.Document as IMxDocument;
            Spellchecker.CheckDocument(doc);
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
