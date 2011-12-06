using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ArcMapSpellCheck {
	/// <summary>
	/// Summary description for InstallerConfig.
	/// </summary>
	[RunInstaller(true)]
	public class InstallerConfig : Installer {
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private Container components;

		/// <summary>
		/// Creates a new instance of <see cref="InstallerConfig"/>.
		/// </summary>
		public InstallerConfig() {
			// This call is required by the Designer.
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing ) {
			try {
				if( disposing ) {
					if(components != null) {
						components.Dispose();
					}
				}
			} finally {
				base.Dispose( disposing );
			}
		}


		#region Component Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			components = new Container();
		}
		#endregion

		/// <summary>
		/// Performs the COM registration for this tool.
		/// </summary>
		public override void Install(IDictionary stateSaver) {
			RegistrationServices regSrv = new RegistrationServices();

			try {
				base.Install (stateSaver);

				if (!regSrv.RegisterAssembly(base.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase))
					throw new InstallException("COM registration failed.  Some or all of the application classes are not properly registered in the ESRI component categories.");
			} catch (Exception ex) {
				ShowErrorMessageBox(ex, "Install Error");
			}
		}

		/// <summary>
		/// Performs the COM unregistration for this tool.
		/// </summary>
		public override void Uninstall(IDictionary savedState) {
			RegistrationServices regSrv = new RegistrationServices();

			try {
				base.Uninstall (savedState);

				if (!regSrv.UnregisterAssembly(base.GetType().Assembly))
					throw new InstallException("COM unregistration failed.  Some or all of the application classes are not properly removed from the ESRI component categories.");
			} catch (Exception ex) {
				ShowErrorMessageBox(ex, "Uninstall Error");
			}
		}

		private static DialogResult ShowErrorMessageBox(Exception ex, string caption) {
			return MessageBox.Show(ex.Message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, 0);
		}

	}
}
