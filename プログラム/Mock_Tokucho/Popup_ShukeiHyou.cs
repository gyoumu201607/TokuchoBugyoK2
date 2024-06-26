﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Security.Cryptography.X509Certificates;
using System.IO;

namespace TokuchoBugyoK2
{
    public partial class Popup_ShukeiHyou : Form
    {
        public string[] UserInfos;
        public string[] ReturnValue = new string[10];
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string nendo;
        private int FromNendo;
        private int ToNendo;
        public string Busho = null;
        public string Tantousha = "";
        public string MadoguchiID = "";
        public string TokuhoBangou = "";
        public string TokuhoBangouEda = "";
        public string KanriBangou = "";
        public string PrintGamen = "";

        private Boolean Busho_Ikkatu = false;
        private Boolean Hinmoku_All = false;

        // VIPS　20220316　課題管理表No1263(957)　ADD　自分大臣の時だけファイルDLタイプの選択を追加
        Popup_Download download_form = null;
        // VIPS　20220316　課題管理表No1263(957)　DEL  DLのフォルダ存在チェック行わない
        //private Boolean existFolder = false;

        public Popup_ShukeiHyou()
        {
            InitializeComponent();

            // コンボボックスにマウスホイールイベントを付与
            this.src_Busho.MouseWheel += item_MouseWheel;
            this.comboBox_Taisho.MouseWheel += item_MouseWheel;
            this.comboBox_Chohyo.MouseWheel += item_MouseWheel;

        }

        private void Popup_ShukeiHyou_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            // 年度
            if (int.TryParse(nendo, out FromNendo))
            {
                ToNendo = FromNendo + 1;
            }
            else
            {
                FromNendo = DateTime.Today.Year;
                ToNendo = FromNendo + 1;
            }
            set_combo(FromNendo.ToString());

            // 調査担当部所
            if (Busho != null && Busho != "")
            {
                src_Busho.SelectedValue = Busho;
            }
            // 調査担当者
            if (Tantousha != null && Tantousha != "")
            {
                src_Tantousha.Text = Tantousha;
            }

            // リンク先を設定するはデフォルトチェック
            item_LinkCheckBox.Checked = true;

            get_data();
            FolderPathCheck();

            //getFileName();

            if (src_Tantousha.Text != "")
            {
                DataTable dt = new DataTable();

                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

                DateTime today = DateTime.Today.Date;
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT " +
                        " BushokanriboKamei " +
                        ",KojinCD " +
                        ",ChousainMei " +
                        "FROM Mst_Chousain LEFT JOIN Mst_Busho ON Mst_Chousain.GyoumuBushoCD = Mst_Busho.GyoumuBushoCD " +
                        "WHERE (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + today + "' ) " +
                        "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + today + "' ) ";

                    if (src_Busho.Text != "")
                    {
                        cmd.CommandText += "AND Mst_Chousain.GyoumuBushoCD COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(src_Busho.SelectedValue.ToString(), 1) + "' ";
                    }
                    if (src_Tantousha.Text != "")
                    {
                        cmd.CommandText += "AND Mst_Chousain.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(src_Tantousha.Text, 1) + "' ";
                    }
                    var sda = new SqlDataAdapter(cmd);
                    dt.Clear();
                    sda.Fill(dt);

                    conn.Close();
                }

                if (dt != null && dt.Rows.Count > 0)
                {
                    label_SentakuBusho.Text = dt.Rows[0][0].ToString();
                    item1_KojinCD.Text = dt.Rows[0][1].ToString();
                    label_SentakuTantousha.Text = dt.Rows[0][2].ToString();

                    getFileName();
                }

            }

            // VIPS　20220316　課題管理表No1263(957)　ADD　自分大臣の時だけファイルDLタイプの選択を追加
            // 自分大臣以外の時、保存を初期選択として変更不可にする
            if (PrintGamen != "Jibun")
            {
                radioButton_DL.Enabled = false;
                radioButton_Save.Checked = true;
            }
            //else //  VIPS　20220322　課題管理表No1263(957)　ADD 自分大臣の時、初期選択はDLにする
            else //  VIPS　20220330　課題管理表No1298(983)　CHANGE 自分大臣の時、初期選択は保存にする
            {
                radioButton_Save.Checked = true;
            }

        }

        //コンボボックス設定
        private void set_combo(string nendo)
        {
            //受託課所支部
            //SQL変数
            //string discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            //string value = "Mst_Busho.GyoumuBushoCD ";
            //string table = "Mst_Busho";
            //string where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
            //        //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' ";
            //        "AND NOT GyoumuBushoCD LIKE '121%' AND BushoMadoguchiHyoujiFlg = 1 ";
            string discript = "Mst_Busho.BushokanriboKamei ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            string where = "KashoShibuCD != '' AND GyoumuBushoCD != '999990' AND GyoumuBushoCD != '127900' AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 ";

            int FromNendo;
            if (int.TryParse(nendo, out FromNendo))
            {
                int ToNendo = int.Parse(nendo) + 1;
                //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";
            }
            where += " ORDER BY BushoMadoguchiNarabijun";

            Console.WriteLine(where);
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            if (combodt != null)
            {
                DataRow dr = combodt.NewRow();
                dr[0] = "127000";
                dr[1] = "本部 調査部門";
                combodt.Rows.Add(dr);
            }
            src_Busho.DisplayMember = "Discript";
            src_Busho.ValueMember = "Value";
            src_Busho.DataSource = combodt;

            if (Busho != null)
            {
                src_Busho.SelectedValue = Busho;
            }

            //対象　コンボ
            DataTable tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "主+副");
            tmpdt.Rows.Add(1, "主");
            tmpdt.Rows.Add(2, "副");
            comboBox_Taisho.DisplayMember = "Discript";
            comboBox_Taisho.ValueMember = "Value";
            comboBox_Taisho.DataSource = tmpdt;

            //帳票情報
            //SQL変数
            combodt = new System.Data.DataTable();
            discript = "PrintName ";
            value = "PrintListID ";
            table = "Mst_PrintList";
            where = "MENU_ID = 203 AND PrintBunruiCD = 3 AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun ";
            combodt = GlobalMethod.getData(discript, value, table, where);
            comboBox_Chohyo.DisplayMember = "Discript";
            comboBox_Chohyo.ValueMember = "Value";
            comboBox_Chohyo.DataSource = combodt;

        }


        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            DateTime today = DateTime.Today.Date;
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT " +
                    "Mst_Chousain.GyoumuBushoCD " +
                    //",ShibuMei " +
                    ",BushokanriboKamei " +
                    ",KojinCD " +
                    ",ChousainMei " +
                    //",ChousaShozoku " +
                    //",ShozokuRyaku " +
                    //",BushoShibuCD " + //支部コード
                    //",KashoShibuCD " + //課コード
                    "FROM Mst_Chousain LEFT JOIN Mst_Busho ON Mst_Chousain.GyoumuBushoCD = Mst_Busho.GyoumuBushoCD " +
                    //"WHERE (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                    //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                    "WHERE (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + today + "' ) " +
                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + today + "' ) ";

                if (src_Busho.Text != "")
                {
                    cmd.CommandText += "AND Mst_Chousain.GyoumuBushoCD COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + GlobalMethod.ChangeSqlText(src_Busho.SelectedValue.ToString().TrimEnd('0'), 1) + "%' ESCAPE '\\' ";
                }
                if (src_Tantousha.Text != "")
                {
                    cmd.CommandText += "AND Mst_Chousain.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_Tantousha.Text, 1) + "%' ESCAPE '\\' ";
                }

                //cmd.CommandText += "ORDER BY ChousainMei ";
                //cmd.CommandText += "ORDER BY KojinCD ";
                cmd.CommandText += "ORDER BY BushoMadoguchiNarabijun, KojinCD";
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);

                // 集計表フォルダ取得
                var dtCommon = new DataTable();
                cmd.CommandText = "SELECT MadoguchiShukeiHyoFolder " +
                    "FROM MadoguchiJouhou " +
                    "WHERE MadoguchiID = '" + MadoguchiID + "' ";
                //データ取得
                var sdaC = new SqlDataAdapter(cmd);
                sdaC.Fill(dtCommon);

                if (dtCommon.Rows.Count > 0)
                {
                    item1_ShukeiFolder.Text = dtCommon.Rows[0][0].ToString();
                }
                conn.Close();
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
            Paging_now.Text = (1).ToString();
            set_data(1);
            //Resize_Grid("c1FlexGrid1");
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid1.HitTest(new Point(e.X, e.Y));

            //部所一括チェックがtrueじゃない
            if (!Busho_Ikkatu)
            {
                //if (hti.Column == 0 & hti.Row != 0)
                if (hti.Column == 0 & hti.Row > 0)
                {
                    var _row = hti.Row;
                    var _col = hti.Column;

                    //選択したデータを表示　部所名　調査員名
                    label_SentakuBusho.Text = c1FlexGrid1.Rows[_row][_col + 2].ToString();
                    item1_KojinCD.Text = c1FlexGrid1.Rows[_row][_col + 3].ToString();
                    label_SentakuTantousha.Text = c1FlexGrid1.Rows[_row][_col + 4].ToString();

                    //配列に格納
                    //ReturnValue[0] = c1FlexGrid1.Rows[_row][_col + 2].ToString();//部所名
                    //ReturnValue[1] = c1FlexGrid1.Rows[_row][_col + 4].ToString();//調査員名
                    //this.Close();
                    getFileName();

                    //  VIPS　20220322　課題管理表No1263(957)　ADD  保存にチェックがついていて、かつ、フォルダがみつからない場合にエラー
                    // フォルダチェック
                    if (!Directory.Exists(item1_ShukeiFolder.Text) && radioButton_Save.Checked)
                    {
                        // 集計表フォルダがみつかりません。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20331", ""));
                        // ファイル出力ボタンを非活性化
                        btnFileExport.Enabled = false;
                    }
                    else
                    {
                        //  VIPS　20220322　課題管理表No1263(957)　ADD  保存にチェックがついていて、ファイルが存在する場合にエラー
                        // フォルダ + ファイル名存在チェック
                        if (File.Exists(item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text) && radioButton_Save.Checked)
                        {
                            // E20332:集計表ファイルが既に存在します。
                            set_error("", 0);
                            set_error(GlobalMethod.GetMessage("E20332", ""));
                            // ファイル出力ボタンを非活性化
                            btnFileExport.Enabled = false;
                        }
                        else
                        {
                            set_error("", 0);
                            // ファイル出力ボタンを活性化
                            btnFileExport.Enabled = true;
                        }
                    }

                }
            }
        }

        //// スクロールバーが表示された場合に表示領域を調整するメソッド
        //public void Resize_Grid(string name)
        //{
        //    // 縦が伸びるとスクロールが出るので、
        //    // スクロールバーが出る分横幅を増やす

        //    Control[] cs;
        //    cs = this.Controls.Find(name, true);
        //    if (cs.Length > 0)
        //    {
        //        var fx = (C1.Win.C1FlexGrid.C1FlexGrid)cs[0];
        //        // 行の高さを足し合わせた値
        //        int h = 0;
        //        for (int i = 0; i < fx.Rows.Count; i++)
        //        {
        //            // 全行の高さを算出
        //            if (fx.Rows[i].Height == -1)
        //            {
        //                h += 22;
        //            }
        //            else
        //            {
        //                h += fx.Rows[i].Height;
        //            }
        //        }
        //        // 今回はここが不要の為、コメントアウト
        //        // 4は、上下の枠（2px + 2px）を表している + 全行の高さを足す
        //        //fx.Height = 4 + h;

        //        int w = 0;
        //        for (int i = 0; i < fx.Cols.Count; i++)
        //        {
        //            // 全列の幅を算出
        //            if (fx.Cols[i].Width == -1)
        //            {
        //                w += 100;
        //            }
        //            else
        //            {
        //                w += fx.Cols[i].Width;
        //            }
        //        }
        //        // 4は、上下の枠（2px + 2px）を表している + 全列の幅を足す
        //        if (fx.Height < 4 + h)
        //        {
        //            fx.Width += 18;
        //        }
        //    }
        //}

        private void button2_Click(object sender, EventArgs e)
        {
            // 調査品目データを取り直しさせるためにパラメータをセット
            ReturnValue[0] = "1";
            this.Close();
        }

        private void Top_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (1).ToString();
            set_data(int.Parse(Paging_now.Text));
        }

        private void Previous_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_now.Text) - 1).ToString();
            set_data(int.Parse(Paging_now.Text));
        }

        private void After_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_now.Text) + 1).ToString();
            set_data(int.Parse(Paging_now.Text));
        }

        private void End_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            set_data(int.Parse(Paging_now.Text));
        }
        private void set_data(int pagenum)
        {
            c1FlexGrid1.Rows.Count = 1;
            c1FlexGrid1.AllowAddNew = true;
            int viewnum = 20;
            int startrow = (pagenum - 1) * viewnum;
            int addnum = ListData.Rows.Count - startrow;
            if (addnum > viewnum)
            {
                addnum = viewnum;
            }
            for (int r = 0; r < addnum; r++)
            {
                c1FlexGrid1.Rows.Add();
                for (int i = 0; i < c1FlexGrid1.Cols.Count - 1; i++)
                {
                    c1FlexGrid1[r + 1, i + 1] = ListData.Rows[startrow + r][i];
                }

            }
            c1FlexGrid1.AllowAddNew = false;
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
        }

        private void set_page_enabled(int now, int last)
        {
            if (now <= 1)
            {
                Top_Page.Enabled = false;
                Previous_Page.Enabled = false;
            }
            else
            {
                Top_Page.Enabled = true;
                Previous_Page.Enabled = true;
            }
            if (now >= last)
            {
                End_Page.Enabled = false;
                After_Page.Enabled = false;
            }
            else
            {
                End_Page.Enabled = true;
                After_Page.Enabled = true;
            }
        }

        private void src_2_KeyDown(object sender, KeyEventArgs e)
        {
            get_data();
        }

        private void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }
            e.DrawBackground();

            bool selected = DrawItemState.Selected == (e.State & DrawItemState.Selected);
            var brush = (selected) ? Brushes.White : Brushes.Black;

            //e.Graphics.DrawString(((ComboBox)sender).Items[e.Index].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            DataRowView r = ((ComboBox)sender).Items[e.Index] as DataRowView;
            if (r != null)
            {
                e.Graphics.DrawString(r.Row["Discript"].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            }
            else
            {
                e.Graphics.DrawString(((ComboBox)sender).Items[e.Index].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            }
            e.DrawFocusRectangle();
        }

        // コンボボックスの選択後のイベントTextChangedで拾う
        private void src_1_TextChanged(object sender, EventArgs e)
        {
            get_data();
        }

        // 職員名
        private void src_2_TextChanged(object sender, EventArgs e)
        {
            get_data();
        }

        private void folderHoukokushoIcon_Click(object sender, EventArgs e)
        {
            if (item1_ShukeiFolder.Text == "")
            {
                System.Diagnostics.Process.Start("EXPLORER.EXE", "");
            }
            else
            {
                // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
                if (System.Text.RegularExpressions.Regex.IsMatch(item1_ShukeiFolder.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // 指定されたフォルダパスが存在するなら開く
                    if (item1_ShukeiFolder.Text != "" && item1_ShukeiFolder.Text != null && Directory.Exists(item1_ShukeiFolder.Text))
                    {
                        System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item1_ShukeiFolder.Text));
                    }
                    else
                    {
                        System.Diagnostics.Process.Start("EXPLORER.EXE", "");
                    }
                }
                else
                {
                    System.Diagnostics.Process.Start("EXPLORER.EXE", "");
                }
            }
        }
        private void folderText_Leave(object sender, EventArgs e)
        {
            FolderPathCheck();
        }
        private void FolderPathCheck()
        {
            // 集計表フォルダ
            if (Directory.Exists(item1_ShukeiFolder.Text))
            {
                item1_ShukeiFolder_icon.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
                set_error("", 0);
                // VIPS　20220316　課題管理表No1263(957)　ADD　自分大臣の時だけファイルDLタイプの選択を追加
                // VIPS　20220322　課題管理表No1263(957)　DEL　DELフォルダチェック行わない
                //existFolder = true;
            }
            else
            {
                item1_ShukeiFolder_icon.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
                set_error("", 0);
                set_error(GlobalMethod.GetMessage("E20331", ""));
                // VIPS　20220316　課題管理表No1263(957)　ADD　自分大臣の時だけファイルDLタイプの選択を追加
                // VIPS　20220322　課題管理表No1263(957)　DEL　DELフォルダチェック行わない
                //existFolder = false;
            }
        }

        private void checkBox_BushoIkkatu_CheckedChanged(object sender, EventArgs e)
        {
            //部所一括集計出力

            //チェックついたら
            if (checkBox_BushoIkkatu.Checked)
            {
                //Gridの1列目の画像を変更　押せなくする
                //c1FlexGrid1.Cols[0].Style.BackgroundImage = Image.FromFile("Resource/Image/folder_gray_s.png");
                c1FlexGrid1.Cols[0].Style.BackgroundImage = Image.FromFile("Resource/Image/ActionDeleteDisabled.png");
                Busho_Ikkatu = true;

                //出力対象の表示を変える
                label_SentakuBusho.Text = src_Busho.Text;
                label_SentakuTantousha.Text = "";

                item1_PritFileName.Enabled = false;

                //  VIPS　20220322　課題管理表No1263(957)　ADD  保存にチェックがついていて、かつ、フォルダがみつからない場合にエラー
                // フォルダチェック
                if (!Directory.Exists(item1_ShukeiFolder.Text) && radioButton_Save.Checked)
                {
                    // 集計表フォルダがみつかりません。
                    set_error("", 0);
                    set_error(GlobalMethod.GetMessage("E20331", ""));
                    // ファイル出力ボタンを非活性化
                    btnFileExport.Enabled = false;
                }
                else
                {
                    btnFileExport.Enabled = true;
                }
            }
            //チェックが外れたら
            else
            {
                //Gridの1列目の画像を変更
                c1FlexGrid1.Cols[0].Style.BackgroundImage = Image.FromFile("Resource/Image/selectRow.png");
                Busho_Ikkatu = false;

                item1_PritFileName.Enabled = true;
            }
        }

        private void checkBox_Zenhinmoku_CheckedChanged(object sender, EventArgs e)
        {
            //全品目一括集計出力
            //チェックがついた場合
            if (checkBox_Zenhinmoku.Checked)
            {
                Hinmoku_All = true;
                //gridと対象非表示
                groupBox2.Visible = false;
                groupBox3.Visible = false;
                // filterを隠す
                groupBox1.Visible = false;

                // 部所一括集計表出力のチェックを外す
                checkBox_BushoIkkatu.Checked = false;

                getFileName();

                //  VIPS　20220322　課題管理表No1263(957)　ADD  保存にチェックがついていて、かつ、フォルダがみつからない場合にエラー
                // フォルダチェック
                if (!Directory.Exists(item1_ShukeiFolder.Text) && radioButton_Save.Checked)
                {
                    // 集計表フォルダがみつかりません。
                    set_error("", 0);
                    set_error(GlobalMethod.GetMessage("E20331", ""));
                    // ファイル出力ボタンを非活性化
                    btnFileExport.Enabled = false;
                }
                else
                {
                    // VIPS 20220322 課題管理表No1263(957) ADD  保存にチェックがついていて、ファイルが存在する場合にエラー
                    if (File.Exists(item1_ShukeiFolder.Text + @"\" + item1_PritFileName) && radioButton_Save.Checked)
                    {
                        // ファイルが存在する
                        // E20332:集計表ファイルが既に存在します。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20332", ""));

                        // ファイル出力ボタンを非活性化
                        btnFileExport.Enabled = false;
                    }
                    else
                    {
                        // ファイル出力ボタンを活性化
                        btnFileExport.Enabled = true;
                    }
                }
            }
            else
            {
                Hinmoku_All = false;
                //gridと対象表示
                groupBox2.Visible = true;
                groupBox3.Visible = true;
                groupBox1.Visible = true;

                // ファイル名を空に
                item1_PritFileName.Text = "";
            }
        }

        // ファイル名を取得
        private void getFileName()
        {
            // 調査員名 + 特調番号 + 管理番号 + 拡張子・・・
            // 調査員名 + 特調番号 + 拡張子
            // 拡張子
            String extensions = ".xlsm";
            String printFileName = "";

            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                if (comboBox_Chohyo.SelectedValue != null)
                {
                    var cmd = conn.CreateCommand();
                    // printFileName取得
                    var dtCommon = new DataTable();
                    cmd.CommandText = "SELECT PrintFileName " +
                        "FROM Mst_PrintList " +
                        "WHERE PrintListID = '" + comboBox_Chohyo.SelectedValue + "' ";
                    //データ取得
                    var sdaC = new SqlDataAdapter(cmd);
                    sdaC.Fill(dtCommon);

                    if (dtCommon.Rows.Count > 0)
                    {
                        printFileName = dtCommon.Rows[0][0].ToString();
                    }
                    conn.Close();

                    if (printFileName != "" && printFileName.Length > 5)
                    {
                        // 後ろから5文字「.xlsm」を取る
                        extensions = printFileName.Substring(printFileName.Length - 5, 5);
                    }
                }
            }
            if (checkBox_Zenhinmoku.Checked)
            {
                item1_PritFileName.Text = "一括集計表" + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
            }
            else
            {
                if (label_SentakuTantousha.Text != "")
                {
                    item1_PritFileName.Text = label_SentakuTantousha.Text + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                }
                else
                {
                    item1_PritFileName.Text = "未登録" + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                }
            }
        }

        // 帳票選択
        private void Chohyo_TextChanged(object sender, EventArgs e)
        {
            getFileName();
        }

        // ファイル出力
        private void btnFileExport_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            // 部所一括集計表出力以外の場合
            if (!checkBox_BushoIkkatu.Checked)
            {
                // VIPS 20220322 課題管理表No1263(957) ADD  保存にチェックがついていて、かつ、ファイルが存在する場合にエラー
                // ファイル存在チェック
                if (File.Exists(item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text) && radioButton_Save.Checked)
                {
                    // 既にファイルが存在する
                    set_error("", 0);
                    set_error(GlobalMethod.GetMessage("E20332", "") + ":" + item1_PritFileName.Text);
                    return;
                }

                // 1:MadoguchiID     窓口ID
                // 2:ZeninSyukeihyo  全品目一括集計表 1:チェック 0:未チェック
                // 3:ShibuMei        支部名
                // 4:KojinCD         個人CD
                // 5:ShuFuku         主+副  0:主+副 1:主のみ 2:副
                // 6:FileName        ファイル名
                // 7:PrintGamen      呼び出し元画面 0:窓口ミハル 1:特命課長  2:自分大臣
                // 7個分先に用意
                string[] report_data = new string[7] { "", "", "", "", "", "", "" };

                // 窓口ID
                report_data[0] = MadoguchiID;
                // 全品目一括集計表 1:チェック 0:未チェック
                if (checkBox_Zenhinmoku.Checked)
                {
                    report_data[1] = "1";
                }
                else
                {
                    report_data[1] = "0";
                }
                // 支部名
                report_data[2] = label_SentakuBusho.Text;
                // 個人CD
                if (checkBox_Zenhinmoku.Checked)
                {
                    report_data[3] = "0";
                }
                else
                {
                    report_data[3] = item1_KojinCD.Text;
                }
                // 主+副  0:主+副 1:主のみ 2:副
                report_data[4] = comboBox_Taisho.SelectedValue.ToString();
                // ファイル名
                report_data[5] = item1_PritFileName.Text;

                report_data[6] = "0";
                switch (PrintGamen)
                {
                    case "Madoguchi":
                        report_data[6] = "0";
                        break;
                    case "Tokumei":
                        report_data[6] = "1";
                        break;
                    case "Jibun":
                        report_data[6] = "2";
                        break;
                    default:
                        break;
                }

                int listID = int.Parse(comboBox_Chohyo.SelectedValue.ToString());

                string[] result = GlobalMethod.InsertMadoguchiReportWork(listID, UserInfos[0], report_data, "Shukeihyo");
                // result
                // 成否判定 0:正常 1：エラー
                // メッセージ（主にエラー用）
                // ファイル物理パス（C:\Work\xxxx\0000000111_xxx.xlsx）
                // ダウンロード時のファイル名（xxx.xlsx）
                if (result != null && result.Length >= 4)
                {
                    if (result[0].Trim() == "1")
                    {
                        //set_error(result[1]);
                        set_error("", 0);
                        set_error(result[1]);
                    }
                    else
                    {
                        set_error("", 0);

                        // VIPS　20220316　課題管理表No1263(957)　ADD　保存、DL選択の分岐を追加	
                        // 直接フォルダに保存するかDLダイアログを表示するか
                        if (radioButton_Save.Checked)
                        {
                            // 成功時は、ファイルをフォルダにコピーする
                            try
                            {
                                System.IO.File.Copy(result[2], item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text, true);

                                set_error("集計表ファイルを出力しました。:" + item1_PritFileName.Text);

                                // リンク先を設定するチェックボックスチェック時
                                if (item_LinkCheckBox.Checked)
                                {
                                    // 対象を取得する
                                    string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                                    DataTable dt0 = new DataTable();
                                    using (var conn = new SqlConnection(connStr))
                                    {
                                        //No1381 1165 リンクについてエクセルのリンクを追加
                                        conn.Open();
                                        var cmd = conn.CreateCommand();
                                        SqlTransaction transaction = conn.BeginTransaction();
                                        cmd.Transaction = transaction;

                                        try
                                        {
                                            string linkpath = item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text;
                                            // 全品目一括集計表
                                            cmd.CommandText = "UPDATE ChousaHinmoku SET ChousaLinkSakli = N'" + linkpath + "' " +
                                                "WHERE " +
                                                "MadoguchiID = '" + MadoguchiID + "' ";
                                            // 全品目一括集計表ではない AND 個人CD が0出ない場合は、個人のみ更新
                                            if (!checkBox_Zenhinmoku.Checked && report_data[3] != "0")
                                            {
                                                cmd.CommandText += "AND (HinmokuChousainCD = '" + report_data[3] + "' " +
                                                    "OR HinmokuFukuChousainCD1 = '" + report_data[3] + "' " +
                                                    "OR HinmokuFukuChousainCD2 = '" + report_data[3] + "' )";
                                            }
                                            cmd.ExecuteNonQuery();


                                            // 全品目一括集計表
                                            cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET MadoguchiL1ShukeihyoLink = N'" + linkpath + "' " +
                                                ", MadoguchiL1AsteriaKoushinFlag = 1 " +
                                                "WHERE " +
                                                "MadoguchiID = '" + MadoguchiID + "' ";
                                            // 全品目一括集計表ではない AND 個人CD が0出ない場合は、個人のみ更新
                                            if (!checkBox_Zenhinmoku.Checked && report_data[3] != "0")
                                            {
                                                cmd.CommandText += "AND MadoguchiL1ChousaTantoushaCD = '" + report_data[3] + "' ";
                                            }
                                            cmd.ExecuteNonQuery();

                                            transaction.Commit();
                                        }
                                        catch (Exception)
                                        {
                                            transaction.Rollback();
                                            // エラーが発生しました
                                            set_error("", 0);
                                            set_error(GlobalMethod.GetMessage("E00091", ""));
                                        }
                                        conn.Close();
                                        //try
                                        //{
                                        //    conn.Open();
                                        //    var cmd = conn.CreateCommand();

                                        //    // 全品目一括集計表ではない
                                        //    if (!checkBox_Zenhinmoku.Checked)
                                        //    {
                                        //        cmd.CommandText = "UPDATE ChousaHinmoku SET ChousaLinkSakli = N'" + item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text + "' " +
                                        //            "WHERE " +
                                        //            "MadoguchiID = '" + MadoguchiID + "' ";

                                        //        // 個人CD が0出ない場合は、個人のみ更新
                                        //        if (report_data[3] != "0")
                                        //        {
                                        //            cmd.CommandText += "AND (HinmokuChousainCD = '" + report_data[3] + "' " +
                                        //                "OR HinmokuFukuChousainCD1 = '" + report_data[3] + "' " +
                                        //                "OR HinmokuFukuChousainCD2 = '" + report_data[3] + "' )";
                                        //        }
                                        //    }
                                        //    else
                                        //    {
                                        //        // 全品目一括集計表
                                        //        cmd.CommandText = "UPDATE ChousaHinmoku SET ChousaLinkSakli = N'" + item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text + "' " +
                                        //            "WHERE " +
                                        //            "MadoguchiID = '" + MadoguchiID + "' ";
                                        //    }
                                        //    cmd.ExecuteNonQuery();
                                        //    conn.Close();
                                        //    // 調査品目データを取り直しさせるためにパラメータをセット
                                        //    ReturnValue[0] = "1";
                                        //}
                                        //catch (Exception)
                                        //{
                                        //    // エラーが発生しました
                                        //    set_error("", 0);
                                        //    set_error(GlobalMethod.GetMessage("E00091", ""));
                                        //}
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                // ファイルコピー失敗
                                set_error(GlobalMethod.GetMessage("E20332", ""));
                            }

                        }
                        else // VIPS　20220316　課題管理表No1263(957)　ADD　DL処理の追加
                        {
                            // DLダイアログを表示する。
                            if (download_form != null)
                            {
                                download_form.Close();
                            }
                            // DLダイアログを表示する。
                            download_form = new Popup_Download();
                            download_form.TopLevel = false;
                            this.Controls.Add(download_form);

                            String fileName = Path.GetFileName(item1_PritFileName.Text);
                            download_form.ExcelName = fileName;
                            download_form.TotalFilePath = result[2];
                            download_form.Dock = DockStyle.Bottom;
                            download_form.Show();
                            download_form.BringToFront();
                        }
                    }
                }
                else
                {
                    // エラーが発生しました
                    set_error("", 0);
                    set_error(GlobalMethod.GetMessage("E00091", ""));
                }
            }
            // 部所一括集計表出力
            else
            {
                // 対象の担当者リスト
                List<string> kojinList = new List<string>();
                List<string> ChousainMeiList = new List<string>();

                // 対象を取得する
                string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

                //分類
                using (var conn = new SqlConnection(connStr))
                {
                    try
                    {
                        var cmd = conn.CreateCommand();
                        //cmd.CommandText = "SELECT " +
                        //        "MadoguchiL1ChousaTantoushaCD " +
                        //        ",MadoguchiL1ChousaTantousha " +
                        //        ",MadoguchiL1ChousaBushoCD " +
                        //        "FROM MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiID = '" + MadoguchiID + "' AND MadoguchiL1ChousaTantoushaCD > 0 " +
                        //        "AND MadoguchiL1ChousaBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                        //        // MadoguchiL1ChousaShinchoku = 1 //調査中
                        //        // 1:調査中　　⇒　40：集計中
                        //        //"AND MadoguchiL1ChousaShinchoku = 40";
                        //        // 旧進捗状況の　1:調査中　は 20:調査開始、30:見積中、40：集計中に該当する
                        //        //"AND MadoguchiL1ChousaShinchoku IN (20,30,40)";
                        //        "AND MadoguchiL1ChousaShinchoku != 80";

                        // 主のデータを取得
                        // 0:主+副 1:主 2:副
                        if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "0" || comboBox_Taisho.SelectedValue.ToString() == "1"))
                        {
                            cmd.CommandText = "SELECT distinct " +
                                    "HinmokuChousainCD " +
                                    ",mc.ChousainMei " +
                                    ",HinmokuRyakuBushoCD " +
                                    "FROM ChousaHinmoku ch " +
                                    "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ch.HinmokuChousainCD = mjmc.MadoguchiL1ChousaTantoushaCD " +
                                    "LEFT JOIN Mst_Chousain mc ON ch.HinmokuChousainCD = mc.KojinCD " +
                                    "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND ch.HinmokuChousainCD > 0 " +
                                    "AND HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                                    "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                            var sda = new SqlDataAdapter(cmd);
                            DataTable dt0 = new DataTable();
                            sda.Fill(dt0);

                            if (dt0 != null && dt0.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt0.Rows.Count; i++)
                                {
                                    // 重複除外
                                    if (!kojinList.Contains(dt0.Rows[i][0].ToString()))
                                    {
                                        kojinList.Add(dt0.Rows[i][0].ToString());
                                        ChousainMeiList.Add(dt0.Rows[i][1].ToString());
                                    }
                                }
                            }
                        }
                        // 副1のデータを取得
                        if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "0" || comboBox_Taisho.SelectedValue.ToString() == "2"))
                        {
                            cmd.CommandText = "SELECT distinct " +
                                    "HinmokuFukuChousainCD1 " +
                                    ",mc.ChousainMei " +
                                    ",HinmokuRyakuBushoFuku1CD " +
                                    "FROM ChousaHinmoku ch " +
                                    "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ch.HinmokuFukuChousainCD1 = mjmc.MadoguchiL1ChousaTantoushaCD " +
                                    "LEFT JOIN Mst_Chousain mc ON ch.HinmokuFukuChousainCD1 = mc.KojinCD " +
                                    "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND ch.HinmokuFukuChousainCD1 > 0 " +
                                    "AND HinmokuRyakuBushoFuku1CD = '" + src_Busho.SelectedValue.ToString() + "' " +
                                    "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                            var sda = new SqlDataAdapter(cmd);
                            DataTable dt0 = new DataTable();
                            sda.Fill(dt0);

                            if (dt0 != null && dt0.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt0.Rows.Count; i++)
                                {
                                    // 重複除外
                                    if (!kojinList.Contains(dt0.Rows[i][0].ToString()))
                                    {
                                        kojinList.Add(dt0.Rows[i][0].ToString());
                                        ChousainMeiList.Add(dt0.Rows[i][1].ToString());
                                    }
                                }
                            }
                        }
                        // 副2のデータを取得
                        if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "0" || comboBox_Taisho.SelectedValue.ToString() == "2"))
                        {
                            cmd.CommandText = "SELECT distinct " +
                                    "HinmokuFukuChousainCD2 " +
                                    ",mc.ChousainMei " +
                                    ",HinmokuRyakuBushoFuku2CD " +
                                    "FROM ChousaHinmoku ch " +
                                    "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ch.HinmokuFukuChousainCD2 = mjmc.MadoguchiL1ChousaTantoushaCD " +
                                    "LEFT JOIN Mst_Chousain mc ON ch.HinmokuFukuChousainCD2 = mc.KojinCD " +
                                    "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND ch.HinmokuFukuChousainCD2 > 0 " +
                                    "AND HinmokuRyakuBushoFuku2CD = '" + src_Busho.SelectedValue.ToString() + "' " +
                                    "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                            var sda = new SqlDataAdapter(cmd);
                            DataTable dt0 = new DataTable();
                            sda.Fill(dt0);

                            if (dt0 != null && dt0.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt0.Rows.Count; i++)
                                {
                                    // 重複除外
                                    if (!kojinList.Contains(dt0.Rows[i][0].ToString()))
                                    {
                                        kojinList.Add(dt0.Rows[i][0].ToString());
                                        ChousainMeiList.Add(dt0.Rows[i][1].ToString());
                                    }
                                }
                            }
                        }
                        conn.Close();
                    }
                    catch (Exception)
                    {
                        //    // エラーが発生しました
                        //    label3.Text = GlobalMethod.GetMessage("E00091", "");
                        //    label3.Visible = true;
                    }
                }
                // 対象者がいる場合
                //if(dt0.Rows.Count > 0)
                if (kojinList.Count > 0)
                {
                    // VIPS　20220322　課題管理表No1263(957)　ADD保存にチェックがついていて、かつ、フォルダがみつからない場合にエラー
                    // フォルダチェック
                    if (!Directory.Exists(item1_ShukeiFolder.Text) && radioButton_Save.Checked)
                    {
                        // 集計表フォルダがみつかりません。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20331", ""));
                        return;
                    }
                    String extensions = ".xlsm";
                    string fileName = "";
                    string errorMsg = "";

                    //set_error("", 0);
                    //for (int i = 0; dt0.Rows.Count > i; i++)
                    for (int i = 0; kojinList.Count > i; i++)
                    {
                        // ファイル名を作成
                        //if (KanriBangou == "")
                        //{
                        //    fileName = dt0.Rows[i][1].ToString() + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                        //}
                        //else
                        //{
                        //    fileName = dt0.Rows[i][1].ToString() + "-" + TokuhoBangou + "-" + TokuhoBangouEda + "-" + KanriBangou + extensions;
                        //}

                        //fileName = dt0.Rows[i][1].ToString() + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                        fileName = ChousainMeiList[i].ToString() + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                        // VIPS　20220322　課題管理表No1263(957)　ADD保存にチェックがついていて、かつ、ファイルが存在する場合にエラー
                        // 存在チェック
                        if (File.Exists(item1_ShukeiFolder.Text + @"\" + fileName) && radioButton_Save.Checked)
                        {
                            // E20332:集計表ファイルが既に存在します。
                            set_error(GlobalMethod.GetMessage("E20332", "") + ":" + fileName);
                        }
                        else
                        {
                            // 作れるファイルが1つでもあれば作る

                            // 1:MadoguchiID     窓口ID
                            // 2:ZeninSyukeihyo  全品目一括集計表 1:チェック 0:未チェック
                            // 3:ShibuMei        支部名
                            // 4:KojinCD         個人CD
                            // 5:ShuFuku         主+副  0:主+副 1:主のみ 2:副
                            // 6:FileName        ファイル名
                            // 7:PrintGamen      呼び出し元画面 0:窓口ミハル 1:特命課長  2:自分大臣
                            // 6個分先に用意
                            string[] report_data = new string[7] { "", "", "", "", "", "", "" };

                            // 窓口ID
                            report_data[0] = MadoguchiID;
                            // 全品目一括集計表 1:チェック 0:未チェック
                            report_data[1] = "0";
                            // 支部名
                            report_data[2] = src_Busho.Text;
                            // 個人CD
                            //report_data[3] = dt0.Rows[i][1].ToString();
                            //report_data[3] = dt0.Rows[i][0].ToString();
                            report_data[3] = kojinList[i].ToString();
                            // 主+副  0:主+副 1:主のみ 2:副
                            report_data[4] = comboBox_Taisho.SelectedValue.ToString();
                            // ファイル名
                            //report_data[5] = item1_PritFileName.Text;
                            report_data[5] = fileName;

                            report_data[6] = "0";
                            switch (PrintGamen)
                            {
                                case "Madoguchi":
                                    report_data[6] = "0";
                                    break;
                                case "Tokumei":
                                    report_data[6] = "1";
                                    break;
                                case "Jibun":
                                    report_data[6] = "2";
                                    break;
                                default:
                                    break;
                            }

                            int listID = int.Parse(comboBox_Chohyo.SelectedValue.ToString());

                            string[] result = GlobalMethod.InsertMadoguchiReportWork(listID, UserInfos[0], report_data, "Shukeihyo");
                            // result
                            // 成否判定 0:正常 1：エラー
                            // メッセージ（主にエラー用）
                            // ファイル物理パス（C:\Work\xxxx\0000000111_xxx.xlsx）
                            // ダウンロード時のファイル名（xxx.xlsx）
                            if (result != null && result.Length >= 4)
                            {
                                if (result[0].Trim() == "1")
                                {
                                    set_error(result[1]);
                                }
                                else
                                {
                                    // VIPS　20220316　課題管理表No1263(957)　ADD　保存、DL選択の分岐を追加
                                    // 直接フォルダに保存するかDLダイアログを表示するか選択させる
                                    if (radioButton_Save.Checked)
                                    {
                                        // 成功時は、ファイルをフォルダにコピーする
                                        try
                                        {
                                            System.IO.File.Copy(result[2], item1_ShukeiFolder.Text + @"\" + fileName, true);
                                            set_error("集計表ファイルを出力しました。:" + fileName);

                                            // リンク先を設定するチェックボックスチェック時
                                            if (item_LinkCheckBox.Checked)
                                            {
                                                // 対象を取得する
                                                using (var conn = new SqlConnection(connStr))
                                                {
                                                    //No1381 1165 リンクについてエクセルのリンクを追加
                                                    conn.Open();
                                                    var cmd = conn.CreateCommand();
                                                    SqlTransaction transaction = conn.BeginTransaction();
                                                    cmd.Transaction = transaction;

                                                    try
                                                    {
                                                        string linkpath = item1_ShukeiFolder.Text + @"\" + fileName;
                                                        // 全品目一括集計表
                                                        cmd.CommandText = "UPDATE ChousaHinmoku SET ChousaLinkSakli = N'" + linkpath + "' " +
                                                            "WHERE " +
                                                            "MadoguchiID = '" + MadoguchiID + "' ";
                                                        // 個人CD が0出ない場合は、個人のみ更新
                                                        if (report_data[3] != "0")
                                                        {
                                                            cmd.CommandText += "AND (HinmokuChousainCD = '" + report_data[3] + "' " +
                                                                "OR HinmokuFukuChousainCD1 = '" + report_data[3] + "' " +
                                                                "OR HinmokuFukuChousainCD2 = '" + report_data[3] + "' )";
                                                        }
                                                        cmd.ExecuteNonQuery();


                                                        // 全品目一括集計表
                                                        cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET MadoguchiL1ShukeihyoLink = N'" + linkpath + "' " +
                                                            ", MadoguchiL1AsteriaKoushinFlag = 1 " +
                                                            "WHERE " +
                                                            "MadoguchiID = '" + MadoguchiID + "' ";
                                                        // 個人CD が0出ない場合は、個人のみ更新
                                                        if (report_data[3] != "0")
                                                        {
                                                            cmd.CommandText += "AND MadoguchiL1ChousaTantoushaCD = '" + report_data[3] + "' ";
                                                        }
                                                        cmd.ExecuteNonQuery();

                                                        transaction.Commit();
                                                    }
                                                    catch (Exception)
                                                    {
                                                        transaction.Rollback();
                                                        // エラーが発生しました
                                                        set_error("", 0);
                                                        set_error(GlobalMethod.GetMessage("E00091", ""));
                                                    }
                                                    conn.Close();

                                                    //try
                                                    //{
                                                    //    conn.Open();
                                                    //    var cmd = conn.CreateCommand();
                                                    //    cmd.CommandText = "UPDATE ChousaHinmoku SET ChousaLinkSakli = N'" + item1_ShukeiFolder.Text + @"\" + fileName + "' " +
                                                    //        "WHERE " +
                                                    //        "MadoguchiID = '" + MadoguchiID + "' ";

                                                    //    // 個人CD が0出ない場合は、個人のみ更新
                                                    //    if (report_data[3] != "0")
                                                    //    {
                                                    //        cmd.CommandText += "AND (HinmokuChousainCD = '" + report_data[3] + "' " +
                                                    //            "OR HinmokuFukuChousainCD1 = '" + report_data[3] + "' " +
                                                    //            "OR HinmokuFukuChousainCD2 = '" + report_data[3] + "' )";
                                                    //    }

                                                    //    cmd.ExecuteNonQuery();
                                                    //    conn.Close();

                                                    //    // 調査品目データを取り直しさせるためにパラメータをセット
                                                    //    ReturnValue[0] = "1";
                                                    //}
                                                    //catch (Exception)
                                                    //{
                                                    //    // エラーが発生しました
                                                    //    set_error("", 0);
                                                    //    set_error(GlobalMethod.GetMessage("E00091", ""));
                                                    //}
                                                }
                                            }

                                        }
                                        catch (Exception)
                                        {
                                            // ファイルコピー失敗
                                            set_error("ファイルコピー失敗:" + fileName);
                                        }
                                    }
                                    else // VIPS　20220316　課題管理表No1263(957)　ADD　DL処理の追加
                                    {
                                        // DLダイアログを表示する。
                                        if (download_form != null)
                                        {
                                            download_form.Close();
                                        }
                                        // DLダイアログを表示する。
                                        download_form = new Popup_Download();
                                        download_form.TopLevel = false;
                                        this.Controls.Add(download_form);

                                        fileName = Path.GetFileName(item1_PritFileName.Text);
                                        download_form.ExcelName = fileName;
                                        download_form.TotalFilePath = result[2];
                                        download_form.Dock = DockStyle.Bottom;
                                        download_form.Show();
                                        download_form.BringToFront();

                                    }
                                }
                            }
                            else
                            {
                                // エラーが発生しました
                                set_error("", 0);
                                set_error(GlobalMethod.GetMessage("E00091", ""));
                            }
                        }
                    }
                }
                // 対象者がいない場合
                else
                {
                    set_error("", 0);
                    // E20350:選択された調査担当部所には、調査員が割り当てられていません。
                    set_error(GlobalMethod.GetMessage("E20350", ""));
                }
            }
        }







        private void set_error(string mes, int flg = 1)
        {
            if (flg == 0)
            {
                ErrorMessage.Text = "";
                ErrorBox.Visible = false;
            }
            else
            {
                if (ErrorMessage.Text != "")
                {
                    ErrorMessage.Text += Environment.NewLine;
                }
                ErrorMessage.Text += mes;
                ErrorBox.Visible = true;
            }
        }

        private void src_Busho_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
        }

        private void comboBox_Chohyo_SelectedIndexChanged(object sender, EventArgs e)
        {
            getFileName();
        }

        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        // ファイル名
        private void item1_PritFileName_TextChanged(object sender, EventArgs e)
        {
            if (item1_PritFileName.Text != "")
            {
                //VIPS 20220322 課題管理表No1263(957) ADD 保存にチェックがついていて、かつ、フォルダがみつからない場合にエラー
                // フォルダチェック
                if (!Directory.Exists(item1_ShukeiFolder.Text) && radioButton_Save.Checked)
                {
                    // 集計表フォルダがみつかりません。
                    set_error("", 0);
                    set_error(GlobalMethod.GetMessage("E20331", ""));
                    // ファイル出力ボタンを非活性化
                    btnFileExport.Enabled = false;
                }
                else
                {
                    //VIPS 20220322 課題管理表No1263(957) ADD 保存にチェックがついていて、ファイルが存在する場合にエラー
                    // フォルダ + ファイル名存在チェック
                    if (File.Exists(item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text) && radioButton_Save.Checked)
                    {
                        // E20332:集計表ファイルが既に存在します。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20332", ""));
                        // ファイル出力ボタンを非活性化
                        btnFileExport.Enabled = false;
                    }
                    else
                    {
                        set_error("", 0);
                        // ファイル出力ボタンを活性化
                        btnFileExport.Enabled = true;
                    }
                }
            }
            else
            {
                set_error("", 0);
                // ファイル出力ボタンを非活性化
                btnFileExport.Enabled = false;
            }
        }

        // VIPS　20220316　課題管理表No1263(957)　ADD　自分大臣の時だけファイルDLタイプの選択を追加
        private void radioButton_DL_CheckedChanged(object sender, EventArgs e)
        {
            //VIPS 20220322 課題管理表No1263(957) ADD ファイルが存在するなどのエラーを消す
            set_error("", 0);

            //VIPS 20220322 課題管理表No1263(957) DEL ファイル存在チェックは行わない
            //if (PrintGamen == "Jibun")
            //{
            //    // DLが選択されている時はフォルダ有無に関係なく出力できるようにする
            //    if (radioButton_DL.Checked || existFolder)
            //    {
            //        // ファイル出力ボタンを活性化
            //        btnFileExport.Enabled = true;
            //        btnFileExport.BackColor = Color.FromArgb(42, 78, 122);
            //    }
            //    else
            //    {
            //        // ファイル出力ボタンを非活性化
            //        btnFileExport.Enabled = false;
            //        btnFileExport.BackColor = Color.DimGray;
            //    }
            //}
        }
    }
}
