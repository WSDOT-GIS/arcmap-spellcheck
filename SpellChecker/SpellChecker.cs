using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.ArcMapUI;
//using Word = Microsoft.Office.Interop.Word;

namespace ArcMapSpellCheck
{
    /// <summary>
    /// Checks the spelling of the text in an ArcMap document.
    /// </summary>
    /// <remarks>
    ///		<para>Based on a <see href="http://edn.esri.com/index.cfm?fa=codeExch.sampleDetail&amp;pg=/arcobjects/9.1/Samples/ArcMap/SpellcheckTextElementsinArcMap.htm">.
    ///		VB code sample from the ESRI EDN site</see>.  Converted to C# and enhanced by Jeff Jacobson.</para>
    ///		<para>This version differs from the original sample on which it was based in the following ways:
    ///		<list type="bullet">
    ///			<item><description>Written in C# instead of VBA.</description></item>
    ///			<item><description>Is implemented as an ArcGIS command instead of a VBA macro.</description></item>
    ///			<item><description>This version of the tool checks the spelling of all text elements as well as the names
    ///			of all maps and layers, in the current ArcMap document.
    ///			The original version would only check a single selected text element, and would crash
    ///			if no text elements were selected.</description></item>
    ///		</list>
    ///		</para>
    ///		<note>
    ///			<para>Because tool utilizes Microsoft Word's spell check tool, this tool will not work if Word is not installed.</para>
    ///			<para>This tool was tested with Office 2000 (i.e., Word 9).</para>
    ///		</note>
    /// </remarks>
    [ComVisible(true)]
    public sealed class Spellchecker : IDisposable
    {
        //// This Regex pattern matches formatting tags that are used in text elements (e.g., <tag></tag> or <tag/>).
        //private const string xmlTagPattern = @"\<(?<slash>/)?[a-zA-Z]+(?!\k<slash>)/?\>";

        /// <summary>Creates a new instance of <see cref="Spellchecker"/>.</summary>
        public Spellchecker()
        {
            // Activate Word
            //_WordApp = new Word.ApplicationClass();
            //_WordDoc = _WordApp.Documents.Add(ref _Missing, ref _Missing, ref _Missing, ref _Missing);

            //// On Windows 7, starting Word causes the ArcMap window to be sent behind other open windows.
            //// Then, when the SpellCheck dialog was shown, it would also be hidden behind the other windows.
            //ActivateArcMap();
        }

        // The following directive disables the compiler's warning: "Ambiguity between method 'method' and non-method 'non-method'. Using method group."
#pragma warning disable 0467
        /// <summary>
        /// Checks the spelling of the given <see cref="IMxDocument"/>.
        /// </summary>
        /// <param name="document">An ArcMap document.</param>
        [CLSCompliant(false), ComVisible(false)]
        public void CheckDocument(IMxDocument document)
        {
            ThrowIfDisposed();

            if (document == null)
            {
                throw new ArgumentNullException("document");
            }
            try
            {
                _CancelSpellChecking = false;
                int spellCheckedTextCount = 0;
                int spellCheckedTocCount = 0;

                // Spellcheck the MXD's text elements.
                spellCheckedTextCount = CheckSpellingOfTextElements(document);

                // Spellcheck the names of the maps and layers in the MXD.
                spellCheckedTocCount = CheckSpellingOfTocItemNames(document.Maps);

                // Refresh the contents.
                document.CurrentContentsView.Refresh(null);

                // Refresh text on map.
                document.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);


                // Display the number of spell-checked text elements and TOC items.
                StringBuilder message = new StringBuilder();
                message.AppendFormat("Text elements: {0}{1}", spellCheckedTextCount, Environment.NewLine);
                message.AppendFormat("Table of contents elements: {0}", spellCheckedTocCount);
                ShowOKMessageBox(message.ToString(), "Information", MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowOKMessageBox(ex.ToString(), "Error", MessageBoxIcon.Error);
            }
        }


        #region DISPOSE/FINALIZE MEMBERS

        #region IDisposable Members

        /// <summary>
        /// Releases all resources used by the object.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion

        private bool _IsDisposed;
        /// <summary>
        /// Gets a value that indicates whether the object is disposed.
        /// </summary>
        public bool IsDisposed { get { return _IsDisposed; } }

        private bool _IsDisposing;
        /// <summary>
        /// Gets a value that indicates whether the object is disposing.
        /// </summary>
        public bool IsDisposing { get { return _IsDisposing; } }

        /// <summary>
        /// This method should be used to perform the actual clean up of all appropriate resources.
        /// </summary>
        private void Dispose(bool disposing)
        {
            if (!IsDisposed)
            {
                _IsDisposing = true;
                OnDisposing(EventArgs.Empty);

                // Release unmanaged resources:  ReleaseBuffer(unmanagedBuffer);
                if (_WordDoc != null)
                {
                    // Close the current document
                    object doNotSaveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    _WordDoc.Close(ref doNotSaveChanges, ref _Missing, ref _Missing);
                    _WordDoc = null;
                }

                if (_WordApp != null)
                {
                    object doNotSaveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    object skipped = Missing.Value;
                    _WordApp.Quit(ref doNotSaveChanges, ref skipped, ref skipped);
                    _WordApp = null;
                }

                if (disposing)
                {
                    // TODO: Release managed resources:  if(managedObject != null) managedObject.Dispose();

                }

                _IsDisposing = false;
                _IsDisposed = true;
                OnDisposed(EventArgs.Empty);
            }
        }

        private readonly object EventLock = new Object();

        /// <summary>
        /// Raised when this instance completes disposal.
        /// </summary>
        public event EventHandler Disposed;
        private void OnDisposed(EventArgs e)
        {
            EventHandler handler;
            lock (EventLock)
            {
                handler = Disposed;
            }
            if (handler != null)
                handler(this, e);
        }

        /// <summary>
        /// Raised when this instance begins disposal.
        /// </summary>
        public event EventHandler Disposing;
        private void OnDisposing(EventArgs e)
        {
            EventHandler handler;
            lock (EventLock)
            {
                handler = Disposing;
            }
            if (handler != null)
                handler(this, e);
        }


        /// <summary>
        /// Call this method at the beginning of each property or method that
        /// cannot be used when this object has been disposed.
        /// </summary>
        private void ThrowIfDisposed()
        {
            if (!IsDisposed)
                return;

            string thisName = this.GetType().Name;
            System.Diagnostics.StackFrame frame = new System.Diagnostics.StackFrame(1);
            string memberName = frame.GetMethod().Name;
            string memberType = "method";

            if ((memberName.Length >= 4) && (memberName.Substring(1, 3) == "et_"))
            {
                memberName = memberName.Substring(4);
                memberType = "property";
            }

            throw new ObjectDisposedException(thisName, String.Format(System.Globalization.CultureInfo.CurrentCulture, "Cannot access the {0} {1} when the IsDisposed property is true.", memberName, memberType));
        }

        /// <summary>
        /// This finalizer is called by the garbage collector automatically.  If the user of this class
        /// fails to call the Dispose method explicitly, this method will call it to free unmanaged resources.
        /// </summary>
        ~Spellchecker()
        {
            Dispose(false);
        }

        #endregion


        private Word.Application _WordApp;
        private Word.Document _WordDoc;
        private object _Missing = Missing.Value;
        private bool _CancelSpellChecking;


        /// <summary>
        /// Checks the spelling of the names of all maps and layers in the table of contents.
        /// </summary>
        /// <param name="maps">An <see cref="IMaps"/> collection of maps.</param>
        /// <returns>
        ///		The number of items that were spellchecked.  Zero will be returned under the following contiditions:
        ///		<list type="bullet">
        ///			<item><description>No mispelled items were detected.</description></item>
        ///			<item><description>Either <paramref name="maps"/>.</description></item>
        ///		</list>
        ///	</returns>
        private int CheckSpellingOfTocItemNames(IMaps maps)
        {
            int spellcheckedItems = 0;

            if (maps == null || maps.Count == 0 || _CancelSpellChecking)
                return spellcheckedItems;

            TableOfContentsItem[] tocItems = TableOfContentsItem.GetAllTableOfContentsItems(maps);
            if (tocItems == null || tocItems.Length == 0)
                return spellcheckedItems;

            foreach (TableOfContentsItem tocItem in tocItems)
            {
                if (_CancelSpellChecking)
                    break;

                tocItem.Name = CheckText(tocItem.Name);

                spellcheckedItems++;
            }

            return spellcheckedItems;
        }

        /// <summary>
        /// Checks the spelling of the text elements in an ArcMap document.
        /// </summary>
        /// <param name="mxDoc">An <see cref="IMxDocument">ArcMap docmuent</see>.</param>
        /// <returns>
        ///		The number of text elements that were spellchecked.  Zero will be returned under the following contiditions:
        ///		<list type="bullet">
        ///			<item><description>No mispelled items were detected.</description></item>
        ///		</list>
        ///	</returns>
        private int CheckSpellingOfTextElements(IMxDocument mxDoc)
        {
            int checkedElementCount = 0;

            if (_CancelSpellChecking)
                return checkedElementCount;

            // Cast both the PageLayout and FocusMap (Layout view and Data view) to IGraphicsContainers
            // and put them into an array.
            IGraphicsContainer[] gContainers = {
                                                   mxDoc.PageLayout as IGraphicsContainer,
                                                   mxDoc.FocusMap as IGraphicsContainer
                                               };

            IElement element;
            ITextElement textElement;

            // Run the MS Word spell check on the text elements in both the Layout and Data views.
            foreach (IGraphicsContainer gContainer in gContainers)
            {
                if (_CancelSpellChecking)
                    break;

                if (gContainer == null)
                    continue;

                gContainer.Reset();
                element = gContainer.Next();

                // Loop through ALL elements and (if they are text elements) check their spelling.
                while (element != null)
                {
                    if (_CancelSpellChecking)
                        break;

                    // Cast the current element as an ITextElement.
                    textElement = element as ITextElement;

                    // If the current element is not a text element, go to the next element.
                    if (textElement == null)
                    {
                        element = gContainer.Next();
                        continue;
                    }

                    // Replace carriage return not followed by a new line with Environment.NewLine.
                    textElement.Text = Regex.Replace(CheckText(textElement.Text), @"\r(?<!\n)", Environment.NewLine);

                    checkedElementCount++;

                    element = gContainer.Next();
                }
            }

            return checkedElementCount;
        }

        private string CheckText(string text)
        {
            _WordDoc.SelectAllEditableRanges(ref _Missing);
            _WordApp.Selection.Text = text;

            // TODO: Replace MS Word code with custom UI.

            //Word.Dialog spellDialog = _WordApp.Dialogs[Word.WdWordDialog.wdDialogToolsSpellingAndGrammar];

            //// Keep the dialog in front of the current process (ArcMap)
            //var hwndSpellDialog = FindWindow("OpusApp", null);
            //SetParent(hwndSpellDialog, System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);


            //int dialogReturn = spellDialog.Show(ref _Missing);
            //_CancelSpellChecking = (dialogReturn == 0 || dialogReturn == -2);


            //_WordDoc.SelectAllEditableRanges(ref _Missing);
            //return _WordApp.Selection.Text;
        }
#pragma warning restore 0467


        /// <summary>
        /// A method used for simplifying the process of showing a message box that only has an OK button.
        /// </summary>
        private static DialogResult ShowOKMessageBox(string message, string title, MessageBoxIcon icon)
        {
            return MessageBox.Show(message, title, MessageBoxButtons.OK, icon, MessageBoxDefaultButton.Button1, 0);
        }

        private static void ActivateArcMap()
        {
            BringWindowToTop(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr FindWindow(string ClassName, string WindowText);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);


    }
}
