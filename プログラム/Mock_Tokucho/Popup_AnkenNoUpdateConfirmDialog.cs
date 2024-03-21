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
		const string HENKO_NASHI = "(変更なし)";

		public Popup_AnkenNoUpdateConfirmDialog(string folderPath_Before, string folderPath_After)
		{
			InitializeComponent();
			FolderPath_Before = folderPath_Before;
			FolderPath_After = folderPath_After;

			//変更後フォルダが空もしくは変更前と同じ場合、変更なしで表記
			if(FolderPath_After.Length == 0 || FolderPath_Before == FolderPath_After)
			{
				FolderPath_After = HENKO_NASHI;
			}

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
