using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.ControlCommands;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Utility.BaseClasses;
using ESRI.ArcGIS.Utility.CATIDs;

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
	[Guid("c50838dc-0292-4b18-bc1a-2dc0407e234e"), ComVisible(true)]
	public sealed class SpellCheckCommand: BaseCommand {
		//// This Regex pattern matches formatting tags that are used in text elements (e.g., <tag></tag> or <tag/>).
		//private const string xmlTagPattern = @"\<(?<slash>/)?[a-zA-Z]+(?!\k<slash>)/?\>";
		private IApplication _app;
		
		/// <summary>
		/// Creates a new instance of <see cref="SpellCheckCommand"/>.
		/// </summary>
		public SpellCheckCommand() {
			// Get the button image from the assembly.
			Assembly asm = GetType().Assembly;
			const string imageName = "ArcMapSpellCheck.SpellCheck.png";  // The name of the spellcheck image resource.
			
			// If the assembly contains the named image resource, set the bitmap property of this class.
			ArrayList resourceNames = new ArrayList(asm.GetManifestResourceNames());
			if (resourceNames.Contains(imageName)) {
				// Get the image resource stream from the assembly.
				Stream imageStream = asm.GetManifestResourceStream(imageName);
				// Create a bitmap with the resource stream.
				Bitmap spellCheckBitmap = new Bitmap(imageStream);
				// Set m_bitmap to the spellcheck bitmap.
				base.m_bitmap = spellCheckBitmap;
			}
//#if(DEBUG)
//			else {
//				MessageBox.Show("\"" + imageName + "\" resource not found.");
//			}
//#endif
			resourceNames = null;

			base.m_caption = "Spell Check"; // The string shown when used as a menu item.
			base.m_category = "WSDOT Tools"; // The Category in which the Command appears in the Customize dialog.
			base.m_message = "Check the spelling of the text items in the current document."; // The message string that appears in the statusbar on mouseover.
			base.m_name = "WSDOTTools_Spellcheck"; // The internal name of this command.  By convention, the category and caption of the command.
			base.m_toolTip = base.m_caption; // The string that appears in the screen tip.


		}

		#region "Component Category Registration"
		[ComRegisterFunction()]
		static void Reg(string regKey) {
			MxCommands.Register(regKey);
		}

		[ComUnregisterFunction()]
		static void Unreg(string regKey) {
			MxCommands.Unregister(regKey);
		}
		#endregion


		/// <summary>
		/// Populates the hook helper variable.
		/// </summary>
		/// <param name="hook">The hook to the application.</param>
		public override void OnCreate(object hook) {
			if (hook != null) {
				//_hookHelper.Hook = hook;
				_app = (IApplication)hook;
			}
		}

		/// <summary>
		/// The currently loaded ArcMap document.
		/// </summary>
		private IMxDocument MxDocument {
			get {
				IMxDocument doc = (IMxDocument)_app.Document;
				return doc;
			}
		}

		/// <summary>
		/// Checks the spelling of text in an ArcMap document.
		/// </summary>
		public override void OnClick() {
			Spellchecker.CheckDocument(this.MxDocument);
		}
	}
}
