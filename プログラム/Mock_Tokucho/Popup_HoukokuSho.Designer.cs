namespace TokuchoBugyoK2
{
    partial class Popup_HoukokuSho
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Popup_HoukokuSho));
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.btnFileExport = new System.Windows.Forms.Button();
            this.button_end = new System.Windows.Forms.Button();
            this.radioButton_Save = new System.Windows.Forms.RadioButton();
            this.radioButton_DL = new System.Windows.Forms.RadioButton();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox_Nendo = new System.Windows.Forms.ComboBox();
            this.dateTime_KikanStart = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox_Month = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBox_Quarter = new System.Windows.Forms.ComboBox();
            this.dateTime_KikanEnd = new System.Windows.Forms.DateTimePicker();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.item1_HoukokuFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.item1_Folder_icon = new System.Windows.Forms.PictureBox();
            this.comboBox_Chohyo = new System.Windows.Forms.ComboBox();
            this.ErrorBox = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.ErrorMessage = new System.Windows.Forms.Label();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.item1_Folder_icon)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel4.ColumnCount = 4;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel4.Controls.Add(this.btnFileExport, 2, 0);
            this.tableLayoutPanel4.Controls.Add(this.button_end, 3, 0);
            this.tableLayoutPanel4.Controls.Add(this.radioButton_Save, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.radioButton_DL, 1, 0);
            this.tableLayoutPanel4.Location = new System.Drawing.Point(311, 164);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(310, 39);
            this.tableLayoutPanel4.TabIndex = 19;
            // 
            // btnFileExport
            // 
            this.btnFileExport.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnFileExport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(42)))), ((int)(((byte)(78)))), ((int)(((byte)(122)))));
            this.btnFileExport.Enabled = false;
            this.btnFileExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFileExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.btnFileExport.ForeColor = System.Drawing.Color.White;
            this.btnFileExport.Location = new System.Drawing.Point(117, 5);
            this.btnFileExport.Margin = new System.Windows.Forms.Padding(5);
            this.btnFileExport.Name = "btnFileExport";
            this.btnFileExport.Size = new System.Drawing.Size(120, 29);
            this.btnFileExport.TabIndex = 9;
            this.btnFileExport.Text = "ファイル出力";
            this.btnFileExport.UseVisualStyleBackColor = false;
            this.btnFileExport.Click += new System.EventHandler(this.btnFileExport_Click);
            // 
            // button_end
            // 
            this.button_end.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.button_end.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(42)))), ((int)(((byte)(78)))), ((int)(((byte)(122)))));
            this.button_end.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_end.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.button_end.ForeColor = System.Drawing.Color.White;
            this.button_end.Location = new System.Drawing.Point(247, 5);
            this.button_end.Margin = new System.Windows.Forms.Padding(5);
            this.button_end.Name = "button_end";
            this.button_end.Size = new System.Drawing.Size(65, 29);
            this.button_end.TabIndex = 10;
            this.button_end.Text = "終了";
            this.button_end.UseVisualStyleBackColor = false;
            this.button_end.Click += new System.EventHandler(this.button_end_Click);
            // 
            // radioButton_Save
            // 
            this.radioButton_Save.AutoSize = true;
            this.radioButton_Save.Checked = true;
            this.radioButton_Save.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.radioButton_Save.Location = new System.Drawing.Point(3, 3);
            this.radioButton_Save.Name = "radioButton_Save";
            this.radioButton_Save.Padding = new System.Windows.Forms.Padding(0, 8, 0, 0);
            this.radioButton_Save.Size = new System.Drawing.Size(54, 29);
            this.radioButton_Save.TabIndex = 7;
            this.radioButton_Save.TabStop = true;
            this.radioButton_Save.Text = "保存";
            this.radioButton_Save.UseVisualStyleBackColor = true;
            // 
            // radioButton_DL
            // 
            this.radioButton_DL.AutoSize = true;
            this.radioButton_DL.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.radioButton_DL.Location = new System.Drawing.Point(63, 3);
            this.radioButton_DL.Name = "radioButton_DL";
            this.radioButton_DL.Padding = new System.Windows.Forms.Padding(0, 8, 0, 0);
            this.radioButton_DL.Size = new System.Drawing.Size(46, 29);
            this.radioButton_DL.TabIndex = 8;
            this.radioButton_DL.Text = "ＤＬ";
            this.radioButton_DL.UseVisualStyleBackColor = true;
            this.radioButton_DL.CheckedChanged += new System.EventHandler(this.radioButton_DL_CheckedChanged);
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 6;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 70F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel3.Controls.Add(this.label5, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.label3, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.comboBox_Nendo, 1, 0);
            this.tableLayoutPanel3.Controls.Add(this.dateTime_KikanStart, 1, 1);
            this.tableLayoutPanel3.Controls.Add(this.label4, 2, 0);
            this.tableLayoutPanel3.Controls.Add(this.comboBox_Month, 3, 0);
            this.tableLayoutPanel3.Controls.Add(this.label6, 2, 1);
            this.tableLayoutPanel3.Controls.Add(this.comboBox_Quarter, 4, 0);
            this.tableLayoutPanel3.Controls.Add(this.dateTime_KikanEnd, 3, 1);
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 98);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel3.Size = new System.Drawing.Size(428, 60);
            this.tableLayoutPanel3.TabIndex = 29;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.Color.White;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(3, 30);
            this.label5.Margin = new System.Windows.Forms.Padding(3);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(44, 17);
            this.label5.TabIndex = 20;
            this.label5.Text = "期間";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.White;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(3, 3);
            this.label3.Margin = new System.Windows.Forms.Padding(3);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 17);
            this.label3.TabIndex = 16;
            this.label3.Text = "年度";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // comboBox_Nendo
            // 
            this.comboBox_Nendo.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBox_Nendo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Nendo.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox_Nendo.FormattingEnabled = true;
            this.comboBox_Nendo.Location = new System.Drawing.Point(53, 3);
            this.comboBox_Nendo.Name = "comboBox_Nendo";
            this.comboBox_Nendo.Size = new System.Drawing.Size(114, 21);
            this.comboBox_Nendo.TabIndex = 2;
            this.comboBox_Nendo.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ComboBox_DrawItem);
            this.comboBox_Nendo.SelectedIndexChanged += new System.EventHandler(this.comboBox_NendoSelectedIndexChanged);
            // 
            // dateTime_KikanStart
            // 
            this.dateTime_KikanStart.CustomFormat = " ";
            this.dateTime_KikanStart.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTime_KikanStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTime_KikanStart.Location = new System.Drawing.Point(53, 30);
            this.dateTime_KikanStart.Name = "dateTime_KikanStart";
            this.dateTime_KikanStart.Size = new System.Drawing.Size(114, 20);
            this.dateTime_KikanStart.TabIndex = 5;
            this.dateTime_KikanStart.ValueChanged += new System.EventHandler(this.dateTimePicker_ValueChanged);
            this.dateTime_KikanStart.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dateTimePicker_KeyDown);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.White;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(173, 3);
            this.label4.Margin = new System.Windows.Forms.Padding(3);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(24, 17);
            this.label4.TabIndex = 19;
            this.label4.Text = "月";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // comboBox_Month
            // 
            this.comboBox_Month.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBox_Month.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Month.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox_Month.FormattingEnabled = true;
            this.comboBox_Month.Location = new System.Drawing.Point(203, 3);
            this.comboBox_Month.Name = "comboBox_Month";
            this.comboBox_Month.Size = new System.Drawing.Size(44, 21);
            this.comboBox_Month.TabIndex = 3;
            this.comboBox_Month.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ComboBox_DrawItem);
            this.comboBox_Month.SelectedIndexChanged += new System.EventHandler(this.comboBox_MonthSelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.White;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(173, 30);
            this.label6.Margin = new System.Windows.Forms.Padding(3);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(24, 17);
            this.label6.TabIndex = 19;
            this.label6.Text = "～";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // comboBox_Quarter
            // 
            this.comboBox_Quarter.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBox_Quarter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Quarter.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox_Quarter.FormattingEnabled = true;
            this.comboBox_Quarter.Location = new System.Drawing.Point(253, 3);
            this.comboBox_Quarter.Name = "comboBox_Quarter";
            this.comboBox_Quarter.Size = new System.Drawing.Size(64, 21);
            this.comboBox_Quarter.TabIndex = 4;
            this.comboBox_Quarter.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ComboBox_DrawItem);
            this.comboBox_Quarter.SelectedIndexChanged += new System.EventHandler(this.comboBox_QuarterSelectedIndexChanged);
            // 
            // dateTime_KikanEnd
            // 
            this.tableLayoutPanel3.SetColumnSpan(this.dateTime_KikanEnd, 3);
            this.dateTime_KikanEnd.CustomFormat = " ";
            this.dateTime_KikanEnd.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTime_KikanEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTime_KikanEnd.Location = new System.Drawing.Point(203, 30);
            this.dateTime_KikanEnd.Name = "dateTime_KikanEnd";
            this.dateTime_KikanEnd.Size = new System.Drawing.Size(114, 20);
            this.dateTime_KikanEnd.TabIndex = 6;
            this.dateTime_KikanEnd.ValueChanged += new System.EventHandler(this.dateTimePicker_ValueChanged);
            this.dateTime_KikanEnd.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dateTimePicker_KeyDown);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel2.Controls.Add(this.item1_HoukokuFolder, 0, 2);
            this.tableLayoutPanel2.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.item1_Folder_icon, 1, 1);
            this.tableLayoutPanel2.Controls.Add(this.comboBox_Chohyo, 1, 0);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 22);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 3;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(420, 70);
            this.tableLayoutPanel2.TabIndex = 31;
            // 
            // item1_HoukokuFolder
            // 
            this.tableLayoutPanel2.SetColumnSpan(this.item1_HoukokuFolder, 2);
            this.item1_HoukokuFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.item1_HoukokuFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.item1_HoukokuFolder.Location = new System.Drawing.Point(3, 57);
            this.item1_HoukokuFolder.MaxLength = 512;
            this.item1_HoukokuFolder.Name = "item1_HoukokuFolder";
            this.item1_HoukokuFolder.Size = new System.Drawing.Size(414, 23);
            this.item1_HoukokuFolder.TabIndex = 22;
            this.item1_HoukokuFolder.Visible = false;
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
            this.label1.Text = "帳票選択";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.CadetBlue;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(3, 30);
            this.label2.Margin = new System.Windows.Forms.Padding(3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 21);
            this.label2.TabIndex = 17;
            this.label2.Text = "報告書フォルダを開く";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // item1_Folder_icon
            // 
            this.item1_Folder_icon.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.item1_Folder_icon.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.item1_Folder_icon.Image = ((System.Drawing.Image)(resources.GetObject("item1_Folder_icon.Image")));
            this.item1_Folder_icon.Location = new System.Drawing.Point(148, 32);
            this.item1_Folder_icon.Margin = new System.Windows.Forms.Padding(2);
            this.item1_Folder_icon.Name = "item1_Folder_icon";
            this.item1_Folder_icon.Size = new System.Drawing.Size(22, 16);
            this.item1_Folder_icon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.item1_Folder_icon.TabIndex = 20;
            this.item1_Folder_icon.TabStop = false;
            this.item1_Folder_icon.Click += new System.EventHandler(this.folderHoukokushoIcon_Click);
            // 
            // comboBox_Chohyo
            // 
            this.comboBox_Chohyo.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBox_Chohyo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Chohyo.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox_Chohyo.FormattingEnabled = true;
            this.comboBox_Chohyo.Location = new System.Drawing.Point(149, 3);
            this.comboBox_Chohyo.Name = "comboBox_Chohyo";
            this.comboBox_Chohyo.Size = new System.Drawing.Size(255, 21);
            this.comboBox_Chohyo.TabIndex = 1;
            this.comboBox_Chohyo.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ComboBox_DrawItem);
            this.comboBox_Chohyo.SelectedIndexChanged += new System.EventHandler(this.comboBox_ChohyoSelectedIndexChanged);
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
            this.ErrorBox.Size = new System.Drawing.Size(618, 1);
            this.ErrorBox.TabIndex = 32;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.White;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.ErrorMessage, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel4, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.ErrorBox, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 0, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Font = new System.Drawing.Font("Symbol", 8.25F);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(624, 321);
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
            this.ErrorMessage.Size = new System.Drawing.Size(620, 13);
            this.ErrorMessage.TabIndex = 33;
            // 
            // Popup_HoukokuSho
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(624, 321);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Popup_HoukokuSho";
            this.Text = "選択リスト 報告書";
            this.Load += new System.EventHandler(this.Popup_HoukokuSho_Load);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel4.PerformLayout();
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.item1_Folder_icon)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.Button btnFileExport;
        private System.Windows.Forms.Button button_end;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBox_Nendo;
        private System.Windows.Forms.DateTimePicker dateTime_KikanStart;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBox_Month;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBox_Quarter;
        private System.Windows.Forms.DateTimePicker dateTime_KikanEnd;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TextBox item1_HoukokuFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox item1_Folder_icon;
        private System.Windows.Forms.ComboBox comboBox_Chohyo;
        private System.Windows.Forms.TableLayoutPanel ErrorBox;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label ErrorMessage;
        private System.Windows.Forms.RadioButton radioButton_Save;
        private System.Windows.Forms.RadioButton radioButton_DL;
    }
}