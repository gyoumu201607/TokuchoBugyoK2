namespace TokuchoBugyoK2
{
    partial class Popup_GyouTsuika
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Popup_GyouTsuika));
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox_ChousaBusho = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox_ChousaTantousha = new System.Windows.Forms.ComboBox();
            this.textBox_TankaTekiyouChiiki = new System.Windows.Forms.TextBox();
            this.textBox_TuikaGyousuu = new System.Windows.Forms.TextBox();
            this.textBox_ZentaiJunKaishiNo = new System.Windows.Forms.TextBox();
            this.ErrorBox = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.ErrorMessage = new System.Windows.Forms.Label();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.btn_LineAdd = new System.Windows.Forms.Button();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel2.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.comboBox_ChousaBusho, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label6, 0, 5);
            this.tableLayoutPanel2.Controls.Add(this.label3, 0, 2);
            this.tableLayoutPanel2.Controls.Add(this.label4, 0, 3);
            this.tableLayoutPanel2.Controls.Add(this.label5, 0, 4);
            this.tableLayoutPanel2.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.comboBox_ChousaTantousha, 1, 1);
            this.tableLayoutPanel2.Controls.Add(this.textBox_TankaTekiyouChiiki, 1, 2);
            this.tableLayoutPanel2.Controls.Add(this.textBox_TuikaGyousuu, 1, 4);
            this.tableLayoutPanel2.Controls.Add(this.textBox_ZentaiJunKaishiNo, 1, 5);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 22);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 6;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.Size = new System.Drawing.Size(360, 160);
            this.tableLayoutPanel2.TabIndex = 31;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.White;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(3, 3);
            this.label1.Margin = new System.Windows.Forms.Padding(3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(140, 17);
            this.label1.TabIndex = 16;
            this.label1.Text = "調査担当部所";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // comboBox_ChousaBusho
            // 
            this.comboBox_ChousaBusho.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBox_ChousaBusho.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_ChousaBusho.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F);
            this.comboBox_ChousaBusho.FormattingEnabled = true;
            this.comboBox_ChousaBusho.Location = new System.Drawing.Point(149, 3);
            this.comboBox_ChousaBusho.Name = "comboBox_ChousaBusho";
            this.comboBox_ChousaBusho.Size = new System.Drawing.Size(194, 21);
            this.comboBox_ChousaBusho.TabIndex = 1;
            this.comboBox_ChousaBusho.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ComboBox_DrawItem);
            this.comboBox_ChousaBusho.SelectedIndexChanged += new System.EventHandler(this.comboBox_ChousaBushoSelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.White;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(3, 134);
            this.label6.Margin = new System.Windows.Forms.Padding(3);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(140, 17);
            this.label6.TabIndex = 16;
            this.label6.Text = "全体順開始番号";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.White;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(3, 59);
            this.label3.Margin = new System.Windows.Forms.Padding(3);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(140, 17);
            this.label3.TabIndex = 16;
            this.label3.Text = "単価適用地域";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.White;
            this.tableLayoutPanel2.SetColumnSpan(this.label4, 2);
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(3, 85);
            this.label4.Margin = new System.Windows.Forms.Padding(3);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(320, 17);
            this.label4.TabIndex = 16;
            this.label4.Text = "※追加行数を指定後、追加ボタンをクリックして下さい。";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.White;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(3, 108);
            this.label5.Margin = new System.Windows.Forms.Padding(3);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(140, 17);
            this.label5.TabIndex = 16;
            this.label5.Text = "追加行数";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.White;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(3, 30);
            this.label2.Margin = new System.Windows.Forms.Padding(3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 17);
            this.label2.TabIndex = 16;
            this.label2.Text = "調査担当者";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // comboBox_ChousaTantousha
            // 
            this.comboBox_ChousaTantousha.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBox_ChousaTantousha.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_ChousaTantousha.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F);
            this.comboBox_ChousaTantousha.FormattingEnabled = true;
            this.comboBox_ChousaTantousha.Location = new System.Drawing.Point(149, 30);
            this.comboBox_ChousaTantousha.Name = "comboBox_ChousaTantousha";
            this.comboBox_ChousaTantousha.Size = new System.Drawing.Size(194, 21);
            this.comboBox_ChousaTantousha.TabIndex = 2;
            this.comboBox_ChousaTantousha.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ComboBox_DrawItem);
            this.comboBox_ChousaTantousha.SelectedIndexChanged += new System.EventHandler(this.comboBox_ChousaTantoushaSelectedIndexChanged);
            // 
            // textBox_TankaTekiyouChiiki
            // 
            this.textBox_TankaTekiyouChiiki.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox_TankaTekiyouChiiki.Location = new System.Drawing.Point(149, 59);
            this.textBox_TankaTekiyouChiiki.Name = "textBox_TankaTekiyouChiiki";
            this.textBox_TankaTekiyouChiiki.Size = new System.Drawing.Size(100, 20);
            this.textBox_TankaTekiyouChiiki.TabIndex = 3;
            // 
            // textBox_TuikaGyousuu
            // 
            this.textBox_TuikaGyousuu.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox_TuikaGyousuu.Location = new System.Drawing.Point(149, 108);
            this.textBox_TuikaGyousuu.MaxLength = 5;
            this.textBox_TuikaGyousuu.Name = "textBox_TuikaGyousuu";
            this.textBox_TuikaGyousuu.Size = new System.Drawing.Size(100, 20);
            this.textBox_TuikaGyousuu.TabIndex = 4;
            this.textBox_TuikaGyousuu.Text = "1";
            this.textBox_TuikaGyousuu.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox_TuikaGyousuu.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textbox_KeyPress);
            this.textBox_TuikaGyousuu.Validated += new System.EventHandler(this.textBox_ValidatedNumeric);
            // 
            // textBox_ZentaiJunKaishiNo
            // 
            this.textBox_ZentaiJunKaishiNo.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox_ZentaiJunKaishiNo.Location = new System.Drawing.Point(149, 134);
            this.textBox_ZentaiJunKaishiNo.MaxLength = 5;
            this.textBox_ZentaiJunKaishiNo.Name = "textBox_ZentaiJunKaishiNo";
            this.textBox_ZentaiJunKaishiNo.Size = new System.Drawing.Size(100, 20);
            this.textBox_ZentaiJunKaishiNo.TabIndex = 5;
            this.textBox_ZentaiJunKaishiNo.Text = "0.00";
            this.textBox_ZentaiJunKaishiNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.textBox_ZentaiJunKaishiNo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textbox_KeyPress);
            this.textBox_ZentaiJunKaishiNo.Validated += new System.EventHandler(this.textBox_ValidatedDecimal);
            // 
            // ErrorBox
            // 
            this.ErrorBox.AutoScroll = true;
            this.ErrorBox.AutoSize = true;
            this.ErrorBox.ColumnCount = 1;
            this.ErrorBox.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.ErrorBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ErrorBox.Location = new System.Drawing.Point(3, 3);
            this.ErrorBox.MaximumSize = new System.Drawing.Size(0, 100);
            this.ErrorBox.Name = "ErrorBox";
            this.ErrorBox.RowCount = 1;
            this.ErrorBox.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.ErrorBox.Size = new System.Drawing.Size(379, 1);
            this.ErrorBox.TabIndex = 32;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.White;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.ErrorMessage, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel4, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.ErrorBox, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Font = new System.Drawing.Font("MS UI Gothic", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(385, 243);
            this.tableLayoutPanel1.TabIndex = 11;
            // 
            // ErrorMessage
            // 
            this.ErrorMessage.AutoSize = true;
            this.ErrorMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ErrorMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.ErrorMessage.ForeColor = System.Drawing.Color.Red;
            this.ErrorMessage.Location = new System.Drawing.Point(2, 6);
            this.ErrorMessage.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.ErrorMessage.Name = "ErrorMessage";
            this.ErrorMessage.Size = new System.Drawing.Size(381, 13);
            this.ErrorMessage.TabIndex = 33;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 2;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel4.Controls.Add(this.btn_LineAdd, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.button_Cancel, 1, 0);
            this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 188);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 39F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(310, 39);
            this.tableLayoutPanel4.TabIndex = 19;
            // 
            // btn_LineAdd
            // 
            this.btn_LineAdd.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btn_LineAdd.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(42)))), ((int)(((byte)(78)))), ((int)(((byte)(122)))));
            this.btn_LineAdd.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btn_LineAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.btn_LineAdd.ForeColor = System.Drawing.Color.White;
            this.btn_LineAdd.Location = new System.Drawing.Point(5, 5);
            this.btn_LineAdd.Margin = new System.Windows.Forms.Padding(5);
            this.btn_LineAdd.Name = "btn_LineAdd";
            this.btn_LineAdd.Size = new System.Drawing.Size(140, 29);
            this.btn_LineAdd.TabIndex = 6;
            this.btn_LineAdd.Text = "追加";
            this.btn_LineAdd.UseVisualStyleBackColor = false;
            this.btn_LineAdd.Click += new System.EventHandler(this.btn_LineAdd_Click);
            // 
            // button_Cancel
            // 
            this.button_Cancel.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.button_Cancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(42)))), ((int)(((byte)(78)))), ((int)(((byte)(122)))));
            this.button_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.button_Cancel.ForeColor = System.Drawing.Color.White;
            this.button_Cancel.Location = new System.Drawing.Point(155, 5);
            this.button_Cancel.Margin = new System.Windows.Forms.Padding(5);
            this.button_Cancel.Name = "button_Cancel";
            this.button_Cancel.Size = new System.Drawing.Size(150, 29);
            this.button_Cancel.TabIndex = 7;
            this.button_Cancel.Text = "キャンセル";
            this.button_Cancel.UseVisualStyleBackColor = false;
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);
            // 
            // Popup_GyouTsuika
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(385, 243);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Popup_GyouTsuika";
            this.Text = "調査品目追加画面";
            this.Load += new System.EventHandler(this.Popup_GyouTsuika_Load);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel4.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBox_ChousaBusho;
        private System.Windows.Forms.TableLayoutPanel ErrorBox;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label ErrorMessage;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.Button btn_LineAdd;
        private System.Windows.Forms.Button button_Cancel;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBox_ChousaTantousha;
        private System.Windows.Forms.TextBox textBox_TankaTekiyouChiiki;
        private System.Windows.Forms.TextBox textBox_TuikaGyousuu;
        private System.Windows.Forms.TextBox textBox_ZentaiJunKaishiNo;
    }
}