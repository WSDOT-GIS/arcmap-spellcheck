using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.ArcMapUI;
using Word;

namespace ArcMapSpellCheck {
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
	public sealed class Spellchecker {
		//// This Regex pattern matches formatting tags that are used in text elements (e.g., <tag></tag> or <tag/>).
		//private const string xmlTagPattern = @"\<(?<slash>/)?[a-zA-Z]+(?!\k<slash>)/?\>";
		
		/// <summary>Creates a new instance of <see cref="Spellchecker"/>.</summary>
		private Spellchecker() { }

        // The following directive disables the compiler's warning: "Ambiguity between method 'method' and non-method 'non-method'. Using method group." 
#pragma warning disable 0467
        /// <summary>
		/// Checks the spelling of the given <see cref="IMxDocument"/>.
		/// </summary>
		/// <param name="document">An ArcMap document.</param>
		[CLSCompliant(false), ComVisible(false)]
		public static void CheckDocument(IMxDocument document) {
			if (document == null) {
				throw new ArgumentNullException("document");
			}
			try {
				int spellCheckedTextCount = 0;
				int spellCheckedTocCount = 0;
				bool wordHasBeenClosed = false;

				object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;
				object skipped = Missing.Value;

				// Activate Word
				Word.Application wordApp = new ApplicationClass();
				// Spellcheck the MXD's text elements.
				spellCheckedTextCount = CheckSpellingOfTextElements(document, wordApp, ref wordHasBeenClosed);
				// Spellcheck the names of the maps and layers in the MXD.
				spellCheckedTocCount = CheckSpellingOfTocItemNames(document.Maps, wordApp, ref wordHasBeenClosed);
				
				// Deactivate Word.
				if (wordApp != null) {
					wordApp.Quit(ref doNotSaveChanges, ref skipped, ref skipped);
					wordApp = null;
					wordHasBeenClosed = true;
				}

                // Refresh the contents.
                document.CurrentContentsView.Refresh(null);
                // Refresh text on map.
                //document.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Type.Missing, Type.Missing);
                document.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, null, null);
                // Display the number of spell-checked text elements and TOC items.
				StringBuilder message = new StringBuilder();
				message.AppendFormat("Text elements: {0}{1}", spellCheckedTextCount, Environment.NewLine);
				message.AppendFormat("Table of contents elements: {0}", spellCheckedTocCount);
				ShowOKMessageBox(message.ToString(), "Information", MessageBoxIcon.Information);
			} catch (Exception ex) {
				ShowOKMessageBox(ex.ToString(), "Error", MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Checks the spelling of the names of all maps and layers in the table of contents.
		/// </summary>
		/// <param name="maps">An <see cref="IMaps"/> collection of maps.</param>
		/// <param name="wordApp">A <see cref="Word.Application"/> object.</param>
		/// <param name="wordHasBeenClosed">A <see cref="bool"/> indicating if <paramref name="wordApp"/> has been closed.</param>
		/// <returns>
		///		The number of items that were spellchecked.  Zero will be returned under the following contiditions:
		///		<list type="bullet">
		///			<item><description>No mispelled items were detected.</description></item>
		///			<item><description>Either <paramref name="maps"/> or <paramref name="wordApp"/> were <see langword="null"/>.</description></item>
		///			<item><description><paramref name="wordHasBeenClosed"/> is <see langword="true"/>.</description></item>
		///		</list>
		///	</returns>
		private static int CheckSpellingOfTocItemNames(IMaps maps, Word.Application wordApp, ref bool wordHasBeenClosed) {
			int spellcheckedItems = 0;
			// Exit if there are no maps or if maps is null.
			if (maps == null || maps.Count == 0 || wordApp == null || wordHasBeenClosed) 
				return spellcheckedItems;

			// These values indicate what button was pressed on Word's spell check dialog.
			const int /*close = -2, ok = -1,*/ cancel = 0;


			// Get all items from the table of contents.
			TableOfContentsItem[] tocItems = TableOfContentsItem.GetAllTableOfContentsItems(maps);

			// Exit if tocItems contains 0 elements.
			if (tocItems == null || tocItems.Length == 0) return spellcheckedItems;

			//Word.Application wordApp = new Word.ApplicationClass();
			object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;
			// Since C# does not permit optional parameters an instance of Missing must be used in instances where these 
			// optional parameters were skipped in the VB version of the code.
			object skipped = Missing.Value;
				
			string currentText;

			Document wdDoc = null;
			foreach (TableOfContentsItem tocItem in tocItems) {
				currentText = tocItem.Name;
				
				wdDoc = wordApp.Documents.Add(ref skipped, ref skipped, ref skipped, ref skipped);
				wordApp.Selection.Text = currentText;
				int dialogReturn = wordApp.Dialogs.Item(WdWordDialog.wdDialogToolsSpellingAndGrammar).Show(ref skipped);
				// If the cancel button was clicked on the spell check dialog...
				if (dialogReturn == cancel) {
					// Quit word if the spell check was canceled.
					wordApp.Quit(ref doNotSaveChanges, ref skipped, ref skipped);
					wordApp = null;
					wordHasBeenClosed = true;
					break;
				}
				else {
					tocItem.Name = wordApp.Selection.Text;
					spellcheckedItems++;
				}

				// Close the current document
				wdDoc.Close(ref doNotSaveChanges, ref skipped, ref skipped);
			}

			//// Quit MS Word.
			//if (wordApp != null) {
			//	wordApp.Quit(ref doNotSaveChanges, ref skipped, ref skipped);
			//	wordApp = null;
			//}
			return spellcheckedItems;
		}

        /// <summary>
        /// Checks the spelling of the text elements in an ArcMap document.
        /// </summary>
        /// <param name="mxDoc">An <see cref="IMxDocument">ArcMap docmuent</see>.</param>
        /// <param name="wordApp">A <see cref="Word.Application"/> object.</param>
        /// <param name="wordHasBeenClosed">A <see cref="bool"/> indicating if <paramref name="wordApp"/> has been closed.</param>
        /// <returns>
        ///		The number of text elements that were spellchecked.  Zero will be returned under the following contiditions:
        ///		<list type="bullet">
        ///			<item><description>No mispelled items were detected.</description></item>
        ///			<item><description>Either <paramref name="maps"/> or <paramref name="wordApp"/> were <see langword="null"/>.</description></item>
        ///			<item><description><paramref name="wordHasBeenClosed"/> is <see langword="true"/>.</description></item>
        ///		</list>
        ///	</returns>
        private static int CheckSpellingOfTextElements(IMxDocument mxDoc, Word.Application wordApp, ref bool wordHasBeenClosed) {
            int checkedElementCount = 0;
            if (wordApp == null || wordHasBeenClosed) return checkedElementCount;

            // These values indicate what button was pressed on Word's spell check dialog.
            const int /*close = -2, ok = -1,*/ cancel = 0;

            // Cast both the PageLayout and FocusMap (Layout view and Data view) to IGraphicsContainers 
            // and put them into an array.
            IGraphicsContainer[] gContainers = { 
												   mxDoc.PageLayout as IGraphicsContainer, 
												   mxDoc.FocusMap as IGraphicsContainer 
											   };

            IElement element;
            ITextElement textElement;
            //Word.Application wordApp = new Word.ApplicationClass();
            object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;
            // Since C# does not permit optional parameters an instance of Missing must be used in instances where these 
            // optional parameters were skipped in the VB version of the code.
            object skipped = Missing.Value;

            string currentText;

            // Run the MS Word spell check on the text elements in both the Layout and Data views.
            foreach (IGraphicsContainer gContainer in gContainers) {
                // If the current IGraphicsContainer is null, skip to the next.
                if (gContainer == null) {
                    continue;
                }

                gContainer.Reset();
                element = gContainer.Next();

                Document wdDoc;

                // Loop through ALL elements and (if they are text elements) check their spelling.
                while (element != null) {
                    // If the application (Word) has been closed, break out of the loop.
                    if (wordApp == null) {
                        break;
                    }
                    // Cast the current element as an ITextElement.
                    textElement = element as ITextElement;
                    // If the current element is not a text element, go to the next element.
                    if (textElement == null) {
                        element = gContainer.Next();
                        continue;
                    }
                    checkedElementCount++;

                    currentText = textElement.Text;

                    // Add a new document and add the currently selected text from ArcMap into the Word document.
                    wdDoc = wordApp.Documents.Add(ref skipped, ref skipped, ref skipped, ref skipped);
                    wordApp.Selection.Text = currentText;

                    // Create and show the spelling and grammar check dialog.
                    Dialog spellDialog = wordApp.Dialogs.Item(WdWordDialog.wdDialogToolsSpellingAndGrammar);
                    int dialogReturn = spellDialog.Show(ref skipped);


                    // If the cancel button was clicked on the spell check dialog...
                    if (dialogReturn == cancel) {
                        // Quit word if the spell check was canceled.
                        wordApp.Quit(ref doNotSaveChanges, ref skipped, ref skipped);
                        wordApp = null;
                        wordHasBeenClosed = true;
                        break;
                    } else {
                        // Replace carriage return not followed by a new line with Environment.NewLine.
                        textElement.Text = Regex.Replace(wordApp.Selection.Text, @"\r(?<!\n)", Environment.NewLine);
                    }

                    // Close the current document
                    wdDoc.Close(ref doNotSaveChanges, ref skipped, ref skipped);
                    element = gContainer.Next();
                }
            }
            //// Quit MS Word.
            //if (wordApp != null) {
            //	wordApp.Quit(ref doNotSaveChanges, ref skipped, ref skipped);
            //	wordApp = null;
            //}
            //mxDoc.ActiveView.Refresh();

            return checkedElementCount;
        }
#pragma warning restore 0467
        

		/// <summary>
		/// A method used for simplifying the process of showing a message box that only has an OK button.
		/// </summary>
		private static DialogResult ShowOKMessageBox(string message, string title, MessageBoxIcon icon) {
			return MessageBox.Show(message, title, MessageBoxButtons.OK, icon, MessageBoxDefaultButton.Button1, 0);
		}
	}
}
