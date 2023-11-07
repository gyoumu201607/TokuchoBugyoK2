using C1.Win.C1FlexGrid;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Windows.Forms;

namespace TokuchoBugyoK2
{
    public partial class Popup_GroupMei : Form
    {
        #region 変数など -----------------------------------------------------

        GlobalMethod GlobalMethod = new GlobalMethod();
        private TableLayoutPanel tableLayoutPanel7;
        private Button button_GyouTsuika;
        private C1.Win.C1FlexGrid.C1FlexGrid GroupMeiGrid;
        private TableLayoutPanel tableLayoutPanel1;
        private Label label72;
        private Label Label_HacchushaKamei;
        private Label label95;
        private Label Label_MadoguchiID;
        private Button button_Shuryou;
        private Button button_Update;

        public string[] UserInfos;//ログインユーザー情報
        private Label label2;
        private Label Label_GyoumuMeishou;
        public string MadoguchiID;//窓口ID
        private string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();//データベース接続先
        DataTable MadoguchiData = new DataTable();//案件情報取得用データテーブル
        DataTable MadoguchiGroupMeiData = new DataTable();//グループ名取得用データテーブル
        int intGroupMeiSuu = 15;//グループ名初期表示行数

        #endregion
        public Popup_GroupMei()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Popup_GroupMei));
            this.button_GyouTsuika = new System.Windows.Forms.Button();
            this.tableLayoutPanel7 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.Label_GyoumuMeishou = new System.Windows.Forms.Label();
            this.Label_HacchushaKamei = new System.Windows.Forms.Label();
            this.label95 = new System.Windows.Forms.Label();
            this.Label_MadoguchiID = new System.Windows.Forms.Label();
            this.label72 = new System.Windows.Forms.Label();
            this.GroupMeiGrid = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.button_Shuryou = new System.Windows.Forms.Button();
            this.button_Update = new System.Windows.Forms.Button();
            this.tableLayoutPanel7.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GroupMeiGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // button_GyouTsuika
            // 
            this.button_GyouTsuika.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.button_GyouTsuika.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(42)))), ((int)(((byte)(78)))), ((int)(((byte)(122)))));
            this.button_GyouTsuika.FlatAppearance.BorderSize = 0;
            this.button_GyouTsuika.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_GyouTsuika.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.button_GyouTsuika.ForeColor = System.Drawing.Color.White;
            this.button_GyouTsuika.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_GyouTsuika.Location = new System.Drawing.Point(8, 123);
            this.button_GyouTsuika.Margin = new System.Windows.Forms.Padding(8);
            this.button_GyouTsuika.Name = "button_GyouTsuika";
            this.button_GyouTsuika.Size = new System.Drawing.Size(154, 30);
            this.button_GyouTsuika.TabIndex = 1;
            this.button_GyouTsuika.Text = "行追加";
            this.button_GyouTsuika.UseVisualStyleBackColor = false;
            this.button_GyouTsuika.Click += new System.EventHandler(this.button_GyouTsuika_Click);
            // 
            // tableLayoutPanel7
            // 
            this.tableLayoutPanel7.AutoSize = true;
            this.tableLayoutPanel7.BackColor = System.Drawing.SystemColors.Control;
            this.tableLayoutPanel7.ColumnCount = 3;
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70.9282F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.88616F));
            this.tableLayoutPanel7.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.18564F));
            this.tableLayoutPanel7.Controls.Add(this.tableLayoutPanel1, 0, 0);
            this.tableLayoutPanel7.Controls.Add(this.GroupMeiGrid, 0, 2);
            this.tableLayoutPanel7.Controls.Add(this.button_GyouTsuika, 0, 1);
            this.tableLayoutPanel7.Controls.Add(this.button_Shuryou, 2, 3);
            this.tableLayoutPanel7.Controls.Add(this.button_Update, 1, 3);
            this.tableLayoutPanel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel7.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel7.Margin = new System.Windows.Forms.Padding(10);
            this.tableLayoutPanel7.Name = "tableLayoutPanel7";
            this.tableLayoutPanel7.RowCount = 4;
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 114F));
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 49F));
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 380F));
            this.tableLayoutPanel7.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel7.Size = new System.Drawing.Size(589, 612);
            this.tableLayoutPanel7.TabIndex = 5;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel7.SetColumnSpan(this.tableLayoutPanel1, 3);
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 160F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.Label_GyoumuMeishou, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.Label_HacchushaKamei, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label95, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.Label_MadoguchiID, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label72, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(10, 10);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(10);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(569, 94);
            this.tableLayoutPanel1.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Teal;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(2, 44);
            this.label2.Margin = new System.Windows.Forms.Padding(1);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(158, 48);
            this.label2.TabIndex = 73;
            this.label2.Text = "業務名称";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Label_GyoumuMeishou
            // 
            this.Label_GyoumuMeishou.AutoSize = true;
            this.Label_GyoumuMeishou.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Label_GyoumuMeishou.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label_GyoumuMeishou.Location = new System.Drawing.Point(165, 43);
            this.Label_GyoumuMeishou.Name = "Label_GyoumuMeishou";
            this.Label_GyoumuMeishou.Size = new System.Drawing.Size(400, 50);
            this.Label_GyoumuMeishou.TabIndex = 72;
            this.Label_GyoumuMeishou.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Label_HacchushaKamei
            // 
            this.Label_HacchushaKamei.AutoSize = true;
            this.Label_HacchushaKamei.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Label_HacchushaKamei.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label_HacchushaKamei.Location = new System.Drawing.Point(165, 22);
            this.Label_HacchushaKamei.Name = "Label_HacchushaKamei";
            this.Label_HacchushaKamei.Size = new System.Drawing.Size(400, 20);
            this.Label_HacchushaKamei.TabIndex = 71;
            this.Label_HacchushaKamei.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label95
            // 
            this.label95.AutoSize = true;
            this.label95.BackColor = System.Drawing.Color.Teal;
            this.label95.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label95.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label95.ForeColor = System.Drawing.Color.White;
            this.label95.Location = new System.Drawing.Point(2, 23);
            this.label95.Margin = new System.Windows.Forms.Padding(1);
            this.label95.Name = "label95";
            this.label95.Size = new System.Drawing.Size(158, 18);
            this.label95.TabIndex = 70;
            this.label95.Text = "発注者名・課名";
            this.label95.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Label_MadoguchiID
            // 
            this.Label_MadoguchiID.AutoSize = true;
            this.Label_MadoguchiID.BackColor = System.Drawing.SystemColors.Control;
            this.Label_MadoguchiID.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Label_MadoguchiID.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label_MadoguchiID.Location = new System.Drawing.Point(165, 1);
            this.Label_MadoguchiID.Name = "Label_MadoguchiID";
            this.Label_MadoguchiID.Size = new System.Drawing.Size(400, 20);
            this.Label_MadoguchiID.TabIndex = 69;
            this.Label_MadoguchiID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label72
            // 
            this.label72.AutoSize = true;
            this.label72.BackColor = System.Drawing.Color.Teal;
            this.label72.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label72.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label72.ForeColor = System.Drawing.Color.White;
            this.label72.Location = new System.Drawing.Point(2, 2);
            this.label72.Margin = new System.Windows.Forms.Padding(1);
            this.label72.Name = "label72";
            this.label72.Size = new System.Drawing.Size(158, 18);
            this.label72.TabIndex = 68;
            this.label72.Text = "特調番号";
            this.label72.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // GroupMeiGrid
            // 
            this.GroupMeiGrid.AllowMerging = C1.Win.C1FlexGrid.AllowMergingEnum.FixedOnly;
            this.GroupMeiGrid.AllowResizing = C1.Win.C1FlexGrid.AllowResizingEnum.Both;
            this.GroupMeiGrid.AutoClipboard = true;
            this.GroupMeiGrid.AutoResize = true;
            this.GroupMeiGrid.BackColor = System.Drawing.Color.White;
            this.GroupMeiGrid.ColumnInfo = resources.GetString("GroupMeiGrid.ColumnInfo");
            this.tableLayoutPanel7.SetColumnSpan(this.GroupMeiGrid, 3);
            this.GroupMeiGrid.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.GroupMeiGrid.HighLight = C1.Win.C1FlexGrid.HighLightEnum.WithFocus;
            this.GroupMeiGrid.KeyActionTab = C1.Win.C1FlexGrid.KeyActionEnum.MoveDown;
            this.GroupMeiGrid.Location = new System.Drawing.Point(8, 171);
            this.GroupMeiGrid.Margin = new System.Windows.Forms.Padding(8);
            this.GroupMeiGrid.Name = "GroupMeiGrid";
            this.GroupMeiGrid.PreserveEditMode = true;
            this.GroupMeiGrid.Rows.Count = 1;
            this.GroupMeiGrid.Size = new System.Drawing.Size(573, 356);
            this.GroupMeiGrid.StyleInfo = resources.GetString("GroupMeiGrid.StyleInfo");
            this.GroupMeiGrid.TabIndex = 3;
            this.GroupMeiGrid.BeforeMouseDown += new C1.Win.C1FlexGrid.BeforeMouseDownEventHandler(this.GroupMeiGrid_BeforeMouseDown);
            // 
            // button_Shuryou
            // 
            this.button_Shuryou.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.button_Shuryou.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(42)))), ((int)(((byte)(78)))), ((int)(((byte)(122)))));
            this.button_Shuryou.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Shuryou.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.button_Shuryou.ForeColor = System.Drawing.Color.White;
            this.button_Shuryou.Location = new System.Drawing.Point(519, 563);
            this.button_Shuryou.Margin = new System.Windows.Forms.Padding(5);
            this.button_Shuryou.Name = "button_Shuryou";
            this.button_Shuryou.Size = new System.Drawing.Size(65, 29);
            this.button_Shuryou.TabIndex = 8;
            this.button_Shuryou.Text = "終了";
            this.button_Shuryou.UseVisualStyleBackColor = false;
            this.button_Shuryou.Click += new System.EventHandler(this.button_Shuryou_Click);
            // 
            // button_Update
            // 
            this.button_Update.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.button_Update.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(42)))), ((int)(((byte)(78)))), ((int)(((byte)(122)))));
            this.button_Update.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_Update.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            this.button_Update.ForeColor = System.Drawing.Color.White;
            this.button_Update.Location = new System.Drawing.Point(434, 563);
            this.button_Update.Margin = new System.Windows.Forms.Padding(5);
            this.button_Update.Name = "button_Update";
            this.button_Update.Size = new System.Drawing.Size(65, 29);
            this.button_Update.TabIndex = 9;
            this.button_Update.Text = "更新";
            this.button_Update.UseVisualStyleBackColor = false;
            this.button_Update.Click += new System.EventHandler(this.button_Update_Click);
            // 
            // Popup_GroupMei
            // 
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(589, 612);
            this.Controls.Add(this.tableLayoutPanel7);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Popup_GroupMei";
            this.Text = "グループ名の設定";
            this.Load += new System.EventHandler(this.Popup_GroupMei_Load);
            this.tableLayoutPanel7.ResumeLayout(false);
            this.tableLayoutPanel7.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GroupMeiGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        private void Popup_GroupMei_Load(object sender, EventArgs e)
        {
            set_data();
        }

        /// <summary>
        /// 案件情報、グループ名をセット
        /// </summary>
        private void set_data()
        {
            //案件情報のセット
            MadoguchiData = get_data(MadoguchiID);
            Label_MadoguchiID.Text = MadoguchiData.Rows[0]["MadoguchiID"].ToString();
            Label_HacchushaKamei.Text = MadoguchiData.Rows[0]["MadoguchiHachuuKikanmei"].ToString();
            Label_GyoumuMeishou.Text = MadoguchiData.Rows[0]["MadoguchiGyoumuMeishou"].ToString();

            //グループ名をGridにセット
            MadoguchiGroupMeiData = get_MadoguchiGroupMei(MadoguchiID);
            GroupMeiGrid.Rows.Count = 1;
            for (int i = 0; i < MadoguchiGroupMeiData.Rows.Count; i++)
            {
                GroupMeiGrid.Rows.Add();
                for (int k = 0; k < MadoguchiGroupMeiData.Columns.Count - 1; k++)
                {
                    GroupMeiGrid.Rows[i + 1][k + 1] = MadoguchiGroupMeiData.Rows[i][k].ToString();
                }
            }

            //初期表示に足りない行数を空白の行で追加
            if (GroupMeiGrid.Rows.Count - 1 < intGroupMeiSuu)
            {
                int num = 0;
                if (GroupMeiGrid.Rows.Count != 1)
                {
                    num = GroupMeiGrid.Rows.Count - 1;
                }
                for (int i = num; i < intGroupMeiSuu; i++)
                {
                    GroupMeiGrid.Rows.Add();
                }
            }


        }

        /// <summary>
        /// グループ名が調査品目明細に登録されているかどうかを判定
        /// true : 登録されていない
        /// false : 登録されている
        /// </summary>
        /// <param name="htirow"></param>
        /// <returns></returns>
        private Boolean Chk_data(int htirow)
        {

            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                DataTable dt = new DataTable();
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                //調査品目明細にChousaMadoguchiGroupMasterIDが何件登録されているか。
                sSql.Append("SELECT");
                sSql.Append("      MadoguchiID ");
                sSql.Append("      ,ChousaMadoguchiGroupMasterID ");
                sSql.Append(" FROM");
                sSql.Append("    ChousaHinmoku");
                sSql.Append(" WHERE");
                sSql.Append("    MadoguchiID = ").Append(MadoguchiID);
                sSql.Append("    AND ChousaMadoguchiGroupMasterID = ").Append(GroupMeiGrid.Rows[htirow][1]);
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                //調査品目明細に登録されていなければtrue,登録されていればfalse。
                if (dt.Rows.Count == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

        }

        /// <summary>
        /// 案件情報の取得
        /// </summary>
        /// <param name="MadoguchiID"></param>
        /// <returns></returns>
        public DataTable get_data(string MadoguchiID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT TOP 1");
                sSql.Append("      MadoguchiID");
                sSql.Append("    , MadoguchiGyoumuMeishou ");
                sSql.Append("    , MadoguchiHachuuKikanmei ");
                sSql.Append(" FROM");
                sSql.Append("    MadoguchiJouhou");
                sSql.Append(" WHERE");
                sSql.Append("    MadoguchiID = ").Append(MadoguchiID);
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 窓口IDに紐づいたグループ名の取得
        /// </summary>
        /// <param name="MadoguchiID"></param>
        /// <returns></returns>
        public DataTable get_MadoguchiGroupMei(string MadoguchiID)
        {
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                sSql.Append("SELECT");
                sSql.Append("      MadoguchiGroupMasterID ");
                sSql.Append("    , MadoguchiGroupMei ");
                sSql.Append("    , Mst_Busho.ShibuMei ");
                sSql.Append("    , Mst_Chousain.ChousainMEI ");
                sSql.Append("    , MadoguchiGroupeChangeDate ");
                sSql.Append("    , MadoguchiID ");
                sSql.Append(" FROM");
                sSql.Append("    MadoguchiGroupMaster");
                sSql.Append(" JOIN");
                sSql.Append("    Mst_Busho");
                sSql.Append(" ON");
                sSql.Append("    MadoguchiGroupeChengeBushoCD = Mst_Busho.GyoumuBushoCD");
                sSql.Append(" JOIN");
                sSql.Append("    Mst_Chousain");
                sSql.Append(" ON");
                sSql.Append("    MadoguchiGropeChengeUserID = Mst_Chousain.KojinCD");
                sSql.Append(" WHERE");
                sSql.Append("    MadoguchiID = ").Append(MadoguchiID);
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
            }
            return dt;
        }

        /// <summary>
        /// 削除ボタン押下時処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GroupMeiGrid_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = ((C1FlexGrid)sender).HitTest(new System.Drawing.Point(e.X, e.Y));
            if (hti.Row > 0 && hti.Column == 0)
            {
                //確認ダイアログ
                //I10003：「グループ名を削除しますがよろしいですか？」
                if (GlobalMethod.outputMessage("I10003", "", 1) == DialogResult.OK)
                {
                    if (GroupMeiGrid.Rows[hti.Row][1] != null)
                    {
                        //調査品目明細に登録されていなければ処理を実行、登録されていればエラーダイアログ。。
                        if (Chk_data(hti.Row))
                        {
                            using (var conn = new SqlConnection(connStr))
                            {
                                //削除処理
                                conn.Open();
                                var cmd = conn.CreateCommand();
                                StringBuilder sSql = new StringBuilder();
                                sSql = new StringBuilder();
                                sSql.Append("DELETE");
                                sSql.Append(" FROM");
                                sSql.Append("    MadoguchiGroupMaster");
                                sSql.Append(" WHERE");
                                sSql.Append("    MadoguchiGroupMasterID = ").Append(GroupMeiGrid.Rows[hti.Row][1]);
                                cmd.CommandText = sSql.ToString();
                                cmd.ExecuteNonQuery();
                                GroupMeiGrid.Rows.Remove(hti.Row);
                            }
                        }
                        else
                        {
                            //E70086：「このグループ名はすでに調査品目に登録されているため削除できません。」
                            MessageBox.Show(GlobalMethod.GetMessage("E70086", ""), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    else
                    {
                        GroupMeiGrid.Rows.Remove(hti.Row);
                    }
                }
            }
        }

        /// <summary>
        /// 更新ボタン押下時処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Update_Click(object sender, EventArgs e)
        {
            //グループ名マスタ　更新処理　開始
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                DataTable dt = new DataTable();
                var cmd = conn.CreateCommand();
                StringBuilder sSql = new StringBuilder();
                //MadoguchiGroupMasterIDの最大値を取得
                sSql.Append("SELECT MAX(MadoguchiGroupMasterID) FROM MadoguchiGroupMaster");
                cmd.CommandText = sSql.ToString();
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                int MadoguchiGroupMasterIDCount = 1;
                if (dt.Rows.Count > 1)
                {
                    MadoguchiGroupMasterIDCount = int.Parse(dt.Rows[0][0].ToString()) + 1;
                }

                //グループ名登録・更新処理
                for (int i = 1; i < GroupMeiGrid.Rows.Count; i++)
                {
                    sSql = new StringBuilder();

                    //新規登録処理
                    if (GroupMeiGrid.Rows[i][1] == null && GroupMeiGrid.Rows[i][2] != null)
                    {
                        sSql.Append("INSERT INTO MadoguchiGroupMaster (");
                        sSql.Append("      MadoguchiGroupMasterID ");
                        sSql.Append("      ,MadoguchiGroupMei ");
                        sSql.Append("      ,MadoguchiGroupCreateDate ");
                        sSql.Append("      ,MadoguchiGropeCreateUserID ");
                        sSql.Append("      ,MadoguchiGroupeCreateBushoCD ");
                        sSql.Append("      ,MadoguchiGroupeChangeDate ");
                        sSql.Append("      ,MadoguchiGropeChengeUserID ");
                        sSql.Append("      ,MadoguchiGroupeChengeBushoCD ");
                        sSql.Append("      ,MadoguchiID ");
                        sSql.Append(")");
                        sSql.Append(" VALUES (");
                        sSql.Append("      \'").Append(MadoguchiGroupMasterIDCount).Append("\'");
                        sSql.Append("      ,\'").Append(GroupMeiGrid.Rows[i][2]).Append("\'");
                        sSql.Append("      ,").Append("GETDATE()");
                        sSql.Append("      ,\'").Append(UserInfos[0]).Append("\'");
                        sSql.Append("      ,\'").Append(UserInfos[2]).Append("\'");
                        sSql.Append("      ,").Append("GETDATE()");
                        sSql.Append("      ,\'").Append(UserInfos[0]).Append("\'");
                        sSql.Append("      ,\'").Append(UserInfos[2]).Append("\'");
                        sSql.Append("      ,\'").Append(MadoguchiID).Append("\'");
                        sSql.Append(" )");
                        cmd.CommandText = sSql.ToString();
                        cmd.ExecuteNonQuery();
                        MadoguchiGroupMasterIDCount++;
                    }
                    //更新処理
                    else if (GroupMeiGrid.Rows[i][1] != null)
                    {
                        sSql.Append("UPDATE MadoguchiGroupMaster");
                        sSql.Append(" SET");
                        sSql.Append("    MadoguchiGroupMei = \'").Append(GroupMeiGrid.Rows[i][2]).Append("\'");
                        sSql.Append("    ,MadoguchiGroupeChangeDate = ").Append("GETDATE()");
                        sSql.Append("    ,MadoguchiGropeChengeUserID = ").Append(UserInfos[0]);
                        sSql.Append("    ,MadoguchiGroupeChengeBushoCD = ").Append(UserInfos[2]);
                        sSql.Append(" WHERE");
                        sSql.Append("    MadoguchiID = ").Append(MadoguchiID);
                        sSql.Append("    AND MadoguchiGroupMasterID = ").Append(GroupMeiGrid.Rows[i][1]);
                        cmd.CommandText = sSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                }
            }
            //グループ名マスタ　更新処理　終了
            this.Close();
        }

        /// <summary>
        /// 終了ボタン押下時処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Shuryou_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 行追加ボタン押下自処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_GyouTsuika_Click(object sender, EventArgs e)
        {
            GroupMeiGrid.Rows.Add();
        }

    }
}
