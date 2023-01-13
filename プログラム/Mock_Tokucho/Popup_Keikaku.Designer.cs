namespace TokuchoBugyoK2
{
    partial class Popup_Keikaku
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Popup_Keikaku));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.label76 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.src_3 = new System.Windows.Forms.TextBox();
            this.src_1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.src_4 = new System.Windows.Forms.TextBox();
            this.src_5 = new System.Windows.Forms.TextBox();
            this.src_6 = new System.Windows.Forms.TextBox();
            this.src_7 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.src_2 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.item_KeikakuKashoShibu = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.c1FlexGrid1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.Top_Page = new System.Windows.Forms.PictureBox();
            this.Previous_Page = new System.Windows.Forms.PictureBox();
            this.After_Page = new System.Windows.Forms.PictureBox();
            this.End_Page = new System.Windows.Forms.PictureBox();
            this.tableLayoutPanel11 = new System.Windows.Forms.TableLayoutPanel();
            this.Paging_now = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.Paging_all = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid1)).BeginInit();
            this.tableLayoutPanel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Top_Page)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Previous_Page)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.After_Page)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.End_Page)).BeginInit();
            this.tableLayoutPanel11.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.White;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.button2, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.groupBox1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.groupBox2, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1163, 580);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // button2
            // 
            this.button2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(42)))), ((int)(((byte)(78)))), ((int)(((byte)(122)))));
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.button2.ForeColor = System.Drawing.Color.White;
            this.button2.Location = new System.Drawing.Point(1093, 545);
            this.button2.Margin = new System.Windows.Forms.Padding(5);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(65, 29);
            this.button2.TabIndex = 8;
            this.button2.Text = "終了";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.AutoSize = true;
            this.groupBox1.Controls.Add(this.tableLayoutPanel2);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1157, 260);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filters";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.AutoSize = true;
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.68034F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 82.31966F));
            this.tableLayoutPanel2.Controls.Add(this.label76, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.label1, 0, 2);
            this.tableLayoutPanel2.Controls.Add(this.src_3, 1, 2);
            this.tableLayoutPanel2.Controls.Add(this.src_1, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label2, 0, 3);
            this.tableLayoutPanel2.Controls.Add(this.label3, 0, 4);
            this.tableLayoutPanel2.Controls.Add(this.label4, 0, 6);
            this.tableLayoutPanel2.Controls.Add(this.label5, 0, 7);
            this.tableLayoutPanel2.Controls.Add(this.src_4, 1, 3);
            this.tableLayoutPanel2.Controls.Add(this.src_5, 1, 4);
            this.tableLayoutPanel2.Controls.Add(this.src_6, 1, 6);
            this.tableLayoutPanel2.Controls.Add(this.src_7, 1, 7);
            this.tableLayoutPanel2.Controls.Add(this.label6, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.src_2, 1, 1);
            this.tableLayoutPanel2.Controls.Add(this.label7, 0, 5);
            this.tableLayoutPanel2.Controls.Add(this.item_KeikakuKashoShibu, 1, 5);
            this.tableLayoutPanel2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(9, 18);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 8;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(707, 224);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // label76
            // 
            this.label76.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label76.AutoSize = true;
            this.label76.BackColor = System.Drawing.Color.White;
            this.label76.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label76.Location = new System.Drawing.Point(3, 5);
            this.label76.Margin = new System.Windows.Forms.Padding(3);
            this.label76.Name = "label76";
            this.label76.Size = new System.Drawing.Size(92, 17);
            this.label76.TabIndex = 12;
            this.label76.Text = "計画売上年度";
            this.label76.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.White;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label1.Location = new System.Drawing.Point(3, 61);
            this.label1.Margin = new System.Windows.Forms.Padding(3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 17);
            this.label1.TabIndex = 12;
            this.label1.Text = "計画発注者名";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // src_3
            // 
            this.src_3.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.src_3.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F);
            this.src_3.Location = new System.Drawing.Point(128, 59);
            this.src_3.Name = "src_3";
            this.src_3.Size = new System.Drawing.Size(490, 21);
            this.src_3.TabIndex = 3;
            this.src_3.TextChanged += new System.EventHandler(this.src_1_TextChanged);
            // 
            // src_1
            // 
            this.src_1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.src_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.src_1.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F);
            this.src_1.FormattingEnabled = true;
            this.src_1.Location = new System.Drawing.Point(128, 3);
            this.src_1.Name = "src_1";
            this.src_1.Size = new System.Drawing.Size(102, 22);
            this.src_1.TabIndex = 1;
            this.src_1.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.ComboBox_DrawItem);
            this.src_1.SelectedIndexChanged += new System.EventHandler(this.src_1_SelectedIndexChanged);
            this.src_1.TextChanged += new System.EventHandler(this.src_1_TextChanged);
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.White;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label2.Location = new System.Drawing.Point(3, 89);
            this.label2.Margin = new System.Windows.Forms.Padding(3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 17);
            this.label2.TabIndex = 12;
            this.label2.Text = "計画案件名";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.White;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label3.Location = new System.Drawing.Point(3, 117);
            this.label3.Margin = new System.Windows.Forms.Padding(3);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 17);
            this.label3.TabIndex = 12;
            this.label3.Text = "計画部所支部";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.White;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label4.Location = new System.Drawing.Point(3, 173);
            this.label4.Margin = new System.Windows.Forms.Padding(3);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(106, 17);
            this.label4.TabIndex = 12;
            this.label4.Text = "前年度案件番号";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.White;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label5.Location = new System.Drawing.Point(3, 201);
            this.label5.Margin = new System.Windows.Forms.Padding(3);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(106, 17);
            this.label5.TabIndex = 12;
            this.label5.Text = "前年度受託番号";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // src_4
            // 
            this.src_4.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.src_4.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F);
            this.src_4.Location = new System.Drawing.Point(128, 87);
            this.src_4.Name = "src_4";
            this.src_4.Size = new System.Drawing.Size(576, 21);
            this.src_4.TabIndex = 4;
            this.src_4.TextChanged += new System.EventHandler(this.src_1_TextChanged);
            // 
            // src_5
            // 
            this.src_5.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.src_5.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F);
            this.src_5.Location = new System.Drawing.Point(128, 115);
            this.src_5.Name = "src_5";
            this.src_5.Size = new System.Drawing.Size(490, 21);
            this.src_5.TabIndex = 5;
            this.src_5.TextChanged += new System.EventHandler(this.src_1_TextChanged);
            // 
            // src_6
            // 
            this.src_6.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.src_6.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F);
            this.src_6.Location = new System.Drawing.Point(128, 171);
            this.src_6.Name = "src_6";
            this.src_6.Size = new System.Drawing.Size(102, 21);
            this.src_6.TabIndex = 7;
            this.src_6.TextChanged += new System.EventHandler(this.src_1_TextChanged);
            // 
            // src_7
            // 
            this.src_7.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.src_7.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F);
            this.src_7.Location = new System.Drawing.Point(128, 199);
            this.src_7.Name = "src_7";
            this.src_7.Size = new System.Drawing.Size(102, 21);
            this.src_7.TabIndex = 8;
            this.src_7.TextChanged += new System.EventHandler(this.src_1_TextChanged);
            // 
            // label6
            // 
            this.label6.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.White;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label6.Location = new System.Drawing.Point(3, 33);
            this.label6.Margin = new System.Windows.Forms.Padding(3);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(64, 17);
            this.label6.TabIndex = 12;
            this.label6.Text = "計画番号";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // src_2
            // 
            this.src_2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.src_2.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F);
            this.src_2.Location = new System.Drawing.Point(128, 31);
            this.src_2.Name = "src_2";
            this.src_2.Size = new System.Drawing.Size(102, 21);
            this.src_2.TabIndex = 2;
            this.src_2.TextChanged += new System.EventHandler(this.src_1_TextChanged);
            // 
            // label7
            // 
            this.label7.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.White;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label7.Location = new System.Drawing.Point(3, 145);
            this.label7.Margin = new System.Windows.Forms.Padding(3);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(92, 17);
            this.label7.TabIndex = 12;
            this.label7.Text = "計画課所支部";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // item_KeikakuKashoShibu
            // 
            this.item_KeikakuKashoShibu.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.item_KeikakuKashoShibu.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 10F);
            this.item_KeikakuKashoShibu.Location = new System.Drawing.Point(128, 143);
            this.item_KeikakuKashoShibu.Name = "item_KeikakuKashoShibu";
            this.item_KeikakuKashoShibu.Size = new System.Drawing.Size(490, 21);
            this.item_KeikakuKashoShibu.TabIndex = 6;
            this.item_KeikakuKashoShibu.TextChanged += new System.EventHandler(this.src_1_TextChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.tableLayoutPanel3);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(3, 269);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(1157, 268);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "選択リスト";
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 1;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Controls.Add(this.c1FlexGrid1, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.tableLayoutPanel4, 0, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.tableLayoutPanel3.Location = new System.Drawing.Point(3, 15);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(1151, 250);
            this.tableLayoutPanel3.TabIndex = 0;
            // 
            // c1FlexGrid1
            // 
            this.c1FlexGrid1.AutoResize = true;
            this.c1FlexGrid1.CellButtonImage = ((System.Drawing.Image)(resources.GetObject("c1FlexGrid1.CellButtonImage")));
            this.c1FlexGrid1.ColumnInfo = resources.GetString("c1FlexGrid1.ColumnInfo");
            this.c1FlexGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1FlexGrid1.ForeColor = System.Drawing.Color.Black;
            this.c1FlexGrid1.Location = new System.Drawing.Point(3, 35);
            this.c1FlexGrid1.Name = "c1FlexGrid1";
            this.c1FlexGrid1.Rows.Count = 1;
            this.c1FlexGrid1.ShowButtons = C1.Win.C1FlexGrid.ShowButtonsEnum.Always;
            this.c1FlexGrid1.Size = new System.Drawing.Size(1145, 244);
            this.c1FlexGrid1.StyleInfo = resources.GetString("c1FlexGrid1.StyleInfo");
            this.c1FlexGrid1.TabIndex = 11;
            this.c1FlexGrid1.BeforeMouseDown += new C1.Win.C1FlexGrid.BeforeMouseDownEventHandler(this.c1FlexGrid1_BeforeMouseDown);
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.tableLayoutPanel4.AutoSize = true;
            this.tableLayoutPanel4.ColumnCount = 5;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel4.Controls.Add(this.Top_Page, 0, 0);
            this.tableLayoutPanel4.Controls.Add(this.Previous_Page, 1, 0);
            this.tableLayoutPanel4.Controls.Add(this.After_Page, 3, 0);
            this.tableLayoutPanel4.Controls.Add(this.End_Page, 4, 0);
            this.tableLayoutPanel4.Controls.Add(this.tableLayoutPanel11, 2, 0);
            this.tableLayoutPanel4.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(188, 26);
            this.tableLayoutPanel4.TabIndex = 13;
            // 
            // Top_Page
            // 
            this.Top_Page.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Top_Page.Image = ((System.Drawing.Image)(resources.GetObject("Top_Page.Image")));
            this.Top_Page.Location = new System.Drawing.Point(0, 0);
            this.Top_Page.Margin = new System.Windows.Forms.Padding(0);
            this.Top_Page.Name = "Top_Page";
            this.Top_Page.Size = new System.Drawing.Size(30, 26);
            this.Top_Page.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.Top_Page.TabIndex = 0;
            this.Top_Page.TabStop = false;
            this.Top_Page.Click += new System.EventHandler(this.Top_Page_Click);
            // 
            // Previous_Page
            // 
            this.Previous_Page.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Previous_Page.Image = ((System.Drawing.Image)(resources.GetObject("Previous_Page.Image")));
            this.Previous_Page.Location = new System.Drawing.Point(30, 0);
            this.Previous_Page.Margin = new System.Windows.Forms.Padding(0);
            this.Previous_Page.Name = "Previous_Page";
            this.Previous_Page.Size = new System.Drawing.Size(30, 26);
            this.Previous_Page.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.Previous_Page.TabIndex = 0;
            this.Previous_Page.TabStop = false;
            this.Previous_Page.Click += new System.EventHandler(this.Previous_Page_Click);
            // 
            // After_Page
            // 
            this.After_Page.Dock = System.Windows.Forms.DockStyle.Fill;
            this.After_Page.Image = ((System.Drawing.Image)(resources.GetObject("After_Page.Image")));
            this.After_Page.Location = new System.Drawing.Point(128, 0);
            this.After_Page.Margin = new System.Windows.Forms.Padding(0);
            this.After_Page.Name = "After_Page";
            this.After_Page.Size = new System.Drawing.Size(30, 26);
            this.After_Page.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.After_Page.TabIndex = 0;
            this.After_Page.TabStop = false;
            this.After_Page.Click += new System.EventHandler(this.After_Page_Click);
            // 
            // End_Page
            // 
            this.End_Page.Dock = System.Windows.Forms.DockStyle.Fill;
            this.End_Page.Image = ((System.Drawing.Image)(resources.GetObject("End_Page.Image")));
            this.End_Page.Location = new System.Drawing.Point(158, 0);
            this.End_Page.Margin = new System.Windows.Forms.Padding(0);
            this.End_Page.Name = "End_Page";
            this.End_Page.Size = new System.Drawing.Size(30, 26);
            this.End_Page.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.End_Page.TabIndex = 0;
            this.End_Page.TabStop = false;
            this.End_Page.Click += new System.EventHandler(this.End_Page_Click);
            // 
            // tableLayoutPanel11
            // 
            this.tableLayoutPanel11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel11.AutoSize = true;
            this.tableLayoutPanel11.ColumnCount = 3;
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel11.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel11.Controls.Add(this.Paging_now, 0, 1);
            this.tableLayoutPanel11.Controls.Add(this.label30, 1, 1);
            this.tableLayoutPanel11.Controls.Add(this.Paging_all, 2, 1);
            this.tableLayoutPanel11.Location = new System.Drawing.Point(63, 3);
            this.tableLayoutPanel11.Name = "tableLayoutPanel11";
            this.tableLayoutPanel11.RowCount = 2;
            this.tableLayoutPanel11.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel11.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel11.Size = new System.Drawing.Size(62, 20);
            this.tableLayoutPanel11.TabIndex = 1;
            // 
            // Paging_now
            // 
            this.Paging_now.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.Paging_now.AutoSize = true;
            this.Paging_now.Location = new System.Drawing.Point(3, 1);
            this.Paging_now.Name = "Paging_now";
            this.Paging_now.Size = new System.Drawing.Size(16, 17);
            this.Paging_now.TabIndex = 1;
            this.Paging_now.Text = "1";
            // 
            // label30
            // 
            this.label30.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label30.AutoSize = true;
            this.label30.Location = new System.Drawing.Point(25, 1);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(12, 17);
            this.label30.TabIndex = 1;
            this.label30.Text = "/";
            // 
            // Paging_all
            // 
            this.Paging_all.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.Paging_all.AutoSize = true;
            this.Paging_all.Location = new System.Drawing.Point(43, 1);
            this.Paging_all.Name = "Paging_all";
            this.Paging_all.Size = new System.Drawing.Size(16, 17);
            this.Paging_all.TabIndex = 1;
            this.Paging_all.Text = "0";
            // 
            // Popup_Keikaku
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1163, 580);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Popup_Keikaku";
            this.Text = "選択リスト 計画番号";
            this.Load += new System.EventHandler(this.Popup_Keikaku_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid1)).EndInit();
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Top_Page)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Previous_Page)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.After_Page)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.End_Page)).EndInit();
            this.tableLayoutPanel11.ResumeLayout(false);
            this.tableLayoutPanel11.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label label76;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox src_3;
        private System.Windows.Forms.ComboBox src_1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox src_4;
        private System.Windows.Forms.TextBox src_5;
        private System.Windows.Forms.TextBox src_6;
        private System.Windows.Forms.TextBox src_7;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private C1.Win.C1FlexGrid.C1FlexGrid c1FlexGrid1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox src_2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.PictureBox Top_Page;
        private System.Windows.Forms.PictureBox Previous_Page;
        private System.Windows.Forms.PictureBox After_Page;
        private System.Windows.Forms.PictureBox End_Page;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel11;
        private System.Windows.Forms.Label Paging_now;
        private System.Windows.Forms.Label label30;
        private System.Windows.Forms.Label Paging_all;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox item_KeikakuKashoShibu;
    }
}