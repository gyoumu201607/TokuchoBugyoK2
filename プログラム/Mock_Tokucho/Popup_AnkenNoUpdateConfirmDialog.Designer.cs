
namespace TokuchoBugyoK2
{
	partial class Popup_AnkenNoUpdateConfirmDialog
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.lbl1_title = new System.Windows.Forms.Label();
			this.lbl1_body = new System.Windows.Forms.Label();
			this.lbl2_title = new System.Windows.Forms.Label();
			this.lbl2_body = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lbl_FolderPath_BeforeUpdate = new System.Windows.Forms.Label();
			this.lbl_FolderPath_AfterUpdate = new System.Windows.Forms.Label();
			this.lbl3_title = new System.Windows.Forms.Label();
			this.lbl3_body = new System.Windows.Forms.Label();
			this.btnOK = new System.Windows.Forms.Button();
			this.BtnCancel = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// lbl1_title
			// 
			this.lbl1_title.AutoSize = true;
			this.lbl1_title.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.lbl1_title.Location = new System.Drawing.Point(12, 21);
			this.lbl1_title.Name = "lbl1_title";
			this.lbl1_title.Size = new System.Drawing.Size(339, 12);
			this.lbl1_title.TabIndex = 0;
			this.lbl1_title.Text = "①「工期開始年度」が変更され「案件番号」が変更となりました。";
			// 
			// lbl1_body
			// 
			this.lbl1_body.Location = new System.Drawing.Point(34, 42);
			this.lbl1_body.Name = "lbl1_body";
			this.lbl1_body.Size = new System.Drawing.Size(995, 33);
			this.lbl1_body.TabIndex = 0;
			this.lbl1_body.Text = "案件番号の変更を行わない場合、「キャンセル」を押下し「工期自」を変更して「工期開始年度」を修正してください。変更を行う場合、次の②を確認願います。";
			// 
			// lbl2_title
			// 
			this.lbl2_title.AutoSize = true;
			this.lbl2_title.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.lbl2_title.Location = new System.Drawing.Point(12, 90);
			this.lbl2_title.Name = "lbl2_title";
			this.lbl2_title.Size = new System.Drawing.Size(336, 12);
			this.lbl2_title.TabIndex = 0;
			this.lbl2_title.Text = "②案件番号が変更となった場合、案件フォルダを確認願います。";
			// 
			// lbl2_body
			// 
			this.lbl2_body.Location = new System.Drawing.Point(34, 112);
			this.lbl2_body.Name = "lbl2_body";
			this.lbl2_body.Size = new System.Drawing.Size(304, 20);
			this.lbl2_body.TabIndex = 0;
			this.lbl2_body.Text = "「変更後の案件フォルダ」のフォルダで変更します。";
			// 
			// label5
			// 
			this.label5.BackColor = System.Drawing.SystemColors.ControlLight;
			this.label5.Location = new System.Drawing.Point(34, 167);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(127, 20);
			this.label5.TabIndex = 0;
			this.label5.Text = "現在の案件フォルダ";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.BackColor = System.Drawing.Color.Khaki;
			this.label6.Location = new System.Drawing.Point(34, 189);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(127, 20);
			this.label6.TabIndex = 0;
			this.label6.Text = "変更後の案件フォルダ";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lbl_FolderPath_BeforeUpdate
			// 
			this.lbl_FolderPath_BeforeUpdate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lbl_FolderPath_BeforeUpdate.Location = new System.Drawing.Point(167, 166);
			this.lbl_FolderPath_BeforeUpdate.Name = "lbl_FolderPath_BeforeUpdate";
			this.lbl_FolderPath_BeforeUpdate.Size = new System.Drawing.Size(862, 20);
			this.lbl_FolderPath_BeforeUpdate.TabIndex = 0;
			this.lbl_FolderPath_BeforeUpdate.Text = "変更前";
			// 
			// lbl_FolderPath_AfterUpdate
			// 
			this.lbl_FolderPath_AfterUpdate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lbl_FolderPath_AfterUpdate.Location = new System.Drawing.Point(167, 192);
			this.lbl_FolderPath_AfterUpdate.Name = "lbl_FolderPath_AfterUpdate";
			this.lbl_FolderPath_AfterUpdate.Size = new System.Drawing.Size(862, 20);
			this.lbl_FolderPath_AfterUpdate.TabIndex = 0;
			this.lbl_FolderPath_AfterUpdate.Text = "変更後";
			// 
			// lbl3_title
			// 
			this.lbl3_title.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.lbl3_title.ForeColor = System.Drawing.Color.Red;
			this.lbl3_title.Location = new System.Drawing.Point(12, 239);
			this.lbl3_title.Name = "lbl3_title";
			this.lbl3_title.Size = new System.Drawing.Size(79, 20);
			this.lbl3_title.TabIndex = 0;
			this.lbl3_title.Text = "【注意事項】";
			// 
			// lbl3_body
			// 
			this.lbl3_body.ForeColor = System.Drawing.Color.Red;
			this.lbl3_body.Location = new System.Drawing.Point(34, 259);
			this.lbl3_body.Name = "lbl3_body";
			this.lbl3_body.Size = new System.Drawing.Size(541, 33);
			this.lbl3_body.TabIndex = 0;
			this.lbl3_body.Text = "「フォルダ変更」を行う場合、現在の「案件フォルダ」をエクスプローラで開いている場合、閉じてから実行してください。開いている場合は、新規フォルダが作成されますが、元" +
    "の案件フォルダが残ったままとなりますので、ご確認ください。";
			// 
			// btnOK
			// 
			this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(844, 295);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(75, 23);
			this.btnOK.TabIndex = 1;
			this.btnOK.Text = "OK";
			this.btnOK.UseVisualStyleBackColor = true;
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// BtnCancel
			// 
			this.BtnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.BtnCancel.Location = new System.Drawing.Point(954, 295);
			this.BtnCancel.Name = "BtnCancel";
			this.BtnCancel.Size = new System.Drawing.Size(75, 23);
			this.BtnCancel.TabIndex = 1;
			this.BtnCancel.Text = "キャンセル";
			this.BtnCancel.UseVisualStyleBackColor = true;
			this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
			// 
			// Popup_AnkenNoUpdateConfirmDialog
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(1065, 339);
			this.Controls.Add(this.BtnCancel);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lbl_FolderPath_AfterUpdate);
			this.Controls.Add(this.lbl_FolderPath_BeforeUpdate);
			this.Controls.Add(this.lbl3_body);
			this.Controls.Add(this.lbl3_title);
			this.Controls.Add(this.lbl2_body);
			this.Controls.Add(this.lbl1_body);
			this.Controls.Add(this.lbl2_title);
			this.Controls.Add(this.lbl1_title);
			this.Name = "Popup_AnkenNoUpdateConfirmDialog";
			this.Text = "案件番号変更　確認ダイアログ";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lbl1_title;
		private System.Windows.Forms.Label lbl1_body;
		private System.Windows.Forms.Label lbl2_title;
		private System.Windows.Forms.Label lbl2_body;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lbl_FolderPath_BeforeUpdate;
		private System.Windows.Forms.Label lbl_FolderPath_AfterUpdate;
		private System.Windows.Forms.Label lbl3_title;
		private System.Windows.Forms.Label lbl3_body;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button BtnCancel;
	}
}