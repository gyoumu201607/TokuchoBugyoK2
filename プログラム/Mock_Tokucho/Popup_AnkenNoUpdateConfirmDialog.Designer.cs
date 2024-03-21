
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
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lbl_FolderPath_BeforeUpdate = new System.Windows.Forms.Label();
			this.lbl_FolderPath_AfterUpdate = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.btnOK = new System.Windows.Forms.Button();
			this.BtnCancel = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label1.Location = new System.Drawing.Point(12, 21);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(339, 12);
			this.label1.TabIndex = 0;
			this.label1.Text = "①「工期開始年度」が変更され「案件番号」が変更となりました。";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(34, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(551, 33);
			this.label2.TabIndex = 0;
			this.label2.Text = "案件番号の変更を行わない場合、「キャンセル」を押下し「工期自」を変更して「工期開始年度」を修正してください。変更を行う場合、次の②を確認願います。";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
			this.label3.Location = new System.Drawing.Point(12, 90);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(336, 12);
			this.label3.TabIndex = 0;
			this.label3.Text = "②案件番号が変更となった場合、案件フォルダを確認願います。";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(34, 112);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(304, 20);
			this.label4.TabIndex = 0;
			this.label4.Text = "「変更後の案件フォルダ」のフォルダで変更します。";
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
			// label7
			// 
			this.label7.ForeColor = System.Drawing.Color.Red;
			this.label7.Location = new System.Drawing.Point(12, 239);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(79, 20);
			this.label7.TabIndex = 0;
			this.label7.Text = "【注意事項】";
			// 
			// label8
			// 
			this.label8.ForeColor = System.Drawing.Color.Red;
			this.label8.Location = new System.Drawing.Point(34, 259);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(541, 33);
			this.label8.TabIndex = 0;
			this.label8.Text = "「フォルダ変更」を行う場合、現在の「案件フォルダ」をエクスプローラで開いている場合、閉じてから実行してください。開いている場合は、新規フォルダが作成されますが、元" +
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
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Name = "Popup_AnkenNoUpdateConfirmDialog";
			this.Text = "案件番号変更　確認ダイアログ";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lbl_FolderPath_BeforeUpdate;
		private System.Windows.Forms.Label lbl_FolderPath_AfterUpdate;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button BtnCancel;
	}
}