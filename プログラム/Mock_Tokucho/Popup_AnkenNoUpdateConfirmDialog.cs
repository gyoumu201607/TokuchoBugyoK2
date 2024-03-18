using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
	public partial class Popup_AnkenNoUpdateConfirmDialog : Form
	{

		public string FolderPath_Before;
		public string FolderPath_After;

		public Popup_AnkenNoUpdateConfirmDialog(string folderPath_Before, string folderPath_After)
		{
			InitializeComponent();
			FolderPath_Before = folderPath_Before;
			FolderPath_After = folderPath_After;

			lbl_FolderPath_BeforeUpdate.Text = FolderPath_Before;
			lbl_FolderPath_AfterUpdate.Text = FolderPath_After;
		}

		private void BtnCancel_Click(object sender, EventArgs e)
		{
			return;
		}

		private void btnOK_Click(object sender, EventArgs e)
		{
			return;
		}
	}
}
