using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace TokuchoBugyoK2
{
    public partial class Tokumei : Form
    {
        public string[] UserInfos;
        private DataTable ListData = new DataTable();
        private DataTable CousainList = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public Boolean ReSearch = false;
        private Boolean warifuriFlg = false;

        private Color errorBackColor = Color.FromArgb(255, 204, 255);

        public Tokumei()
        {
            InitializeComponent();

            // コンボボックスにマウスホイールイベントを付与
            this.item_Nendo.MouseWheel += item_MouseWheel;
            this.item_ChousaBusho.MouseWheel += item_MouseWheel;
            this.item_MadoguchiBusho.MouseWheel += item_MouseWheel;
            this.item_FromTo.MouseWheel += item_MouseWheel;
            this.item_ShimekiriSentaku.MouseWheel += item_MouseWheel;
            this.item_ChousaKind.MouseWheel += item_MouseWheel;
            this.item_Shintyokujyoukyo.MouseWheel += item_MouseWheel;
            this.item_TantoushaJoukyo.MouseWheel += item_MouseWheel;
        }

        private void Tokumei_Load(object sender, EventArgs e)
        {
            //不具合No1355（1123）
            lblVersion.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                lblBootMode.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }

            //ユーザ名を設定
            label32.Text = UserInfos[3] + "：" + UserInfos[1];

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            Hashtable imgMap = new Hashtable();

            gridSizeChange();

            //ソート項目にアイコンを設定
            C1.Win.C1FlexGrid.CellRange cr;
            Bitmap bmp1 = new Bitmap("Resource/Image/SortIconDefalt.png");
            Bitmap bmpSort = new Bitmap(bmp1, bmp1.Width / 6, bmp1.Height / 6);
            cr = c1FlexGrid1.GetCellRange(0, 1);
            cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            cr.Image = bmpSort;

            Bitmap bmp2 = new Bitmap("Resource/Image/header_blank.png");

            for (int i = 3; i < c1FlexGrid1.Cols.Count; i++)
            {
                cr = c1FlexGrid1.GetCellRange(0, i);
                cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
                cr.Image = bmpSort;
            }

            // 進捗アイコン
            imgMap = new Hashtable();
            imgMap.Add("8", Image.FromFile("Resource/Image/shin_ao.png"));     // 報告済み
            imgMap.Add("5", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（二次検証済み）
            //imgMap.Add("6", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（中止）
            imgMap.Add("6", Image.FromFile("Resource/Image/shin_ao.png"));     // 中止
            imgMap.Add("7", Image.FromFile("Resource/Image/shin_midori.png")); // 担当者済み
            imgMap.Add("1", Image.FromFile("Resource/Image/shin_dokuro.png")); // 締切日経過
            imgMap.Add("2", Image.FromFile("Resource/Image/shin_aka.png"));    // 締切日が3日以内、かつ2次検証が完了していない
            imgMap.Add("3", Image.FromFile("Resource/Image/shin_kiiro.png"));  // 締切日が1週間以内、かつ2次検証が完了していない
            imgMap.Add("4", Image.FromFile("Resource/Image/blank2.png"));      // それ以外
            //imgMap.Add("4", Image.FromFile("Resource/Image/blank2.png"));
            c1FlexGrid1.Cols[1].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[1].ImageMap = imgMap;
            c1FlexGrid1.Cols[1].ImageAndText = false;

            //編集の画像切り替え
            imgMap = new Hashtable();
            imgMap.Add("0", Image.FromFile("Resource/Image/file_presentation1_g.png"));
            imgMap.Add("1", Image.FromFile("Resource/Image/file_presentation1.png"));
            c1FlexGrid1.Cols[2].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[2].ImageMap = imgMap;
            c1FlexGrid1.Cols[2].ImageAndText = false;

            set_combo();
            ClearForm();
            get_data();
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {

            var hti = this.c1FlexGrid1.HitTest(new Point(e.X, e.Y));
            if (hti.Row > 0 && hti.Column == 2)
            {
                this.ReSearch = true;
                Tokumei_Input form = new Tokumei_Input();
                form.MadoguchiID = c1FlexGrid1[hti.Row, 3].ToString();
                form.UserInfos = UserInfos;
                form.Show(this);
                this.Hide();
            }
        }


        private void set_combo()
        {

            //コンボボックスの内容を設定
            GlobalMethod GlobalMethod = new GlobalMethod();
            DataTable combodt;
            System.Data.DataTable tmpdt;
            DataRow dr;
            SortedList sl;

            //売上年度
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";

            // 売上年度は年度マスタのデータを全て表示
            String where = "";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            item_Nendo.DisplayMember = "Discript";
            item_Nendo.ValueMember = "Value";
            item_Nendo.DataSource = combodt;

            // 調査担当所
            discript = "Mst_Busho.BushokanriboKamei ";
            value = "Mst_Busho.GyoumuBushoCD ";
            table = "Mst_Busho";
            where = "KashoShibuCD != '' AND GyoumuBushoCD != '999990' AND GyoumuBushoCD != '127900' AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
                    " ORDER BY BushoMadoguchiNarabijun";
            combodt = GlobalMethod.getData(discript, value, table, where);
            dr = combodt.NewRow();
            combodt.Rows.InsertAt(dr, 0);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[20].DataMap = sl;
            item_ChousaBusho.DisplayMember = "Discript";
            item_ChousaBusho.ValueMember = "Value";
            item_ChousaBusho.DataSource = combodt;


            DataTable combodt3 = GlobalMethod.getData(discript, value, table, where);
            dr = combodt3.NewRow();
            combodt3.Rows.InsertAt(dr, 0);

            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt3);
            c1FlexGrid1.Cols[7].DataMap = sl; // 窓口部所




            // 検索期間の指定
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "以前");
            tmpdt.Rows.Add(2, "当日");
            tmpdt.Rows.Add(3, "一週間");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_FromTo.DisplayMember = "Discript";
            item_FromTo.ValueMember = "Value";
            item_FromTo.DataSource = tmpdt;

            // 締め日の選択
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(4, "前日の締めは？");
            tmpdt.Rows.Add(1, "本日の締めは？");
            tmpdt.Rows.Add(5, "翌日の締めは？");
            tmpdt.Rows.Add(2, "今週の締めは？");
            tmpdt.Rows.Add(3, "来週の締めは？");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_ShimekiriSentaku.DisplayMember = "Discript";
            item_ShimekiriSentaku.ValueMember = "Value";
            item_ShimekiriSentaku.DataSource = tmpdt;

            // 表示件数
            // 画面上で定義

            // 調査種別
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, " ");
            tmpdt.Rows.Add(1, "単品");
            tmpdt.Rows.Add(2, "一般");
            tmpdt.Rows.Add(3, "単契");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //if (tmpdt != null)
            //{
            //    dr = tmpdt.NewRow();
            //    tmpdt.Rows.InsertAt(dr, 0);
            //}
            item_ChousaKind.DisplayMember = "Discript";
            item_ChousaKind.ValueMember = "Value";
            item_ChousaKind.DataSource = tmpdt;
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[9].DataMap = sl; // 調査種別



            // 実施区分
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, " ");
            tmpdt.Rows.Add(1, "実施");
            tmpdt.Rows.Add(2, "打診中");
            tmpdt.Rows.Add(3, "中止");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //c1FlexGrid1.Cols[10].DataMap = sl; // 実施区分

            // 進捗状況
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(6, "超過");
            tmpdt.Rows.Add(5, "締め切りまで3日以内");
            tmpdt.Rows.Add(4, "締め切りまで1週間以内");
            tmpdt.Rows.Add(3, "締め切りまで1週間以上 または 中止");
            tmpdt.Rows.Add(1, "完了");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            if (tmpdt != null)
            {
                dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            item_Shintyokujyoukyo.DisplayMember = "Discript";
            item_Shintyokujyoukyo.ValueMember = "Value";
            item_Shintyokujyoukyo.DataSource = tmpdt;

            // VIPS　20220330　課題管理表No1294(982) ADD 担当者状況の追加
            // 担当者状況
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(10, "依頼");
            tmpdt.Rows.Add(20, "調査開始");
            tmpdt.Rows.Add(30, "見積中");
            tmpdt.Rows.Add(40, "集計中");
            tmpdt.Rows.Add(50, "担当者済");
            tmpdt.Rows.Add(60, "一次検済");
            tmpdt.Rows.Add(70, "二次検済");
            tmpdt.Rows.Add(200, "調査中");
            tmpdt.Rows.Add(220, "依頼・調査中");
            tmpdt.Rows.Add(230, "依頼・調査中・担当者済");
            tmpdt.Rows.Add(300, "検証中");
            tmpdt.Rows.Add(310, "検証中・二次検済");
            tmpdt.Rows.Add(80, "中止");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //if (tmpdt != null)
            //{
            //    dr = tmpdt.NewRow();
            //    tmpdt.Rows.InsertAt(dr, 0);
            //}
            item_TantoushaJoukyo.DisplayMember = "Discript";
            item_TantoushaJoukyo.ValueMember = "Value";
            item_TantoushaJoukyo.DataSource = tmpdt;

            // 管理帳票印刷コンボボックス
            discript = "PrintName";
            value = "PrintListID";
            table = "Mst_PrintList";
            //where = "";
            where = "MENU_ID = 300 AND PrintBunruiCD = 1 AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            //dr = combodt.NewRow();
            //combodt.Rows.InsertAt(dr, 0);
            item_tyouhyouInsatu.DataSource = combodt;
            item_tyouhyouInsatu.DisplayMember = "Discript";
            item_tyouhyouInsatu.ValueMember = "Value";

            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);

            for (int i = 0; i < 20; i++)
            {
                c1FlexGrid1.Cols[20 + i].DataMap = sl;
            }
        }

        private void set_combo_shibu(string nendo)
        {
            DataTable combodt;
            DataTable combodt2;
            System.Data.DataTable tmpdt;
            DataRow dr;
            SortedList sl;

            // 調査担当部所、窓口部所、受託部所の選択値を保持しておき、コンボ再作成後にセットする
            string ChousaBusho_SelectedValue = "";
            string MadoguchiBusho_SelectedValue = "";
            // 調査担当部所
            if (item_ChousaBusho.Text != "")
            {
                ChousaBusho_SelectedValue = item_ChousaBusho.SelectedValue.ToString();
            }
            // 窓口部所
            if (item_MadoguchiBusho.Text != "")
            {
                MadoguchiBusho_SelectedValue = item_MadoguchiBusho.SelectedValue.ToString();
            }

            // 調査担当部所
            string discript = "Mst_Busho.BushokanriboKamei ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            string where = "KashoShibuCD != '' AND GyoumuBushoCD != '999990' AND GyoumuBushoCD != '127900' AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 ";
            int FromNendo;
            int ToNendo;

            if (!int.TryParse(nendo, out FromNendo))
            {
                FromNendo = DateTime.Today.Year;
            }
            ToNendo = FromNendo + 1;

            if (int.TryParse(nendo, out FromNendo))
            {
                ToNendo = int.Parse(nendo) + 1;
                if (item_NendoOption3Nen.Checked)
                {
                    ToNendo -= 2;
                }
                //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";
            }

            where += " ORDER BY BushoMadoguchiNarabijun";

            combodt = GlobalMethod.getData(discript, value, table, where);

            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[20].DataMap = sl;
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }
            item_ChousaBusho.DisplayMember = "Discript";
            item_ChousaBusho.ValueMember = "Value";
            item_ChousaBusho.DataSource = combodt;

            combodt2 = GlobalMethod.getData(discript, value, table, where);

            if (combodt2 != null)
            {
                dr = combodt2.NewRow();
                combodt2.Rows.InsertAt(dr, 0);
            }

            // 窓口部所
            item_MadoguchiBusho.DisplayMember = "Discript";
            item_MadoguchiBusho.ValueMember = "Value";
            item_MadoguchiBusho.DataSource = combodt2;

            // 受託部所
            discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            value = "Mst_Busho.GyoumuBushoCD ";
            table = "Mst_Busho";
            where = "KashoShibuCD != '' AND GyoumuBushoCD != '999990' AND GyoumuBushoCD != '127900' AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 ";
            //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
            //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
            where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";
            where += " ORDER BY BushoMadoguchiNarabijun";

            combodt = GlobalMethod.getData(discript, value, table, where);
            sl = new SortedList();
            //行の数だけの数だけSortedListにIDとValueをadd
            sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[20].DataMap = sl;
            if (combodt != null)
            {
                dr = combodt.NewRow();
                combodt.Rows.InsertAt(dr, 0);
            }

            // 値を戻す
            // 調査担当部所
            if (ChousaBusho_SelectedValue != "")
            {
                item_ChousaBusho.SelectedValue = ChousaBusho_SelectedValue;
            }
            // 窓口部所
            if (MadoguchiBusho_SelectedValue != "")
            {
                item_MadoguchiBusho.SelectedValue = MadoguchiBusho_SelectedValue;
            }
        }

        // 検索条件クリア
        private void ClearForm()
        {
            //検索条件初期化
            //売上年度　受託課所支部
            /*
            String discript = "NendoSeireki ";
            String value = "NendoID ";
            String table = "Mst_Nendo ";
            String where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
            //コンボボックスデータ取得
            DataTable dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null)
            {
                item_Nendo.SelectedValue = dt.Rows[0][0].ToString();
            }
            else
            {
                item_Nendo.SelectedValue = System.DateTime.Now.Year;
            }
            */
            item_Nendo.SelectedValue = GlobalMethod.GetTodayNendo();
            set_combo_shibu(item_Nendo.Text.ToString());

            item_Hyoujikensuu.SelectedIndex = 1;

            item_NendoOptionTounen.Checked = true;
            item_NendoOption3Nen.Checked = false;
            item_ChousaBusho.SelectedValue = UserInfos[2];
            item_HachushaKamei.Text = "";
            item_TokuchoBangou.Text = "";
            item_DateFrom.Text = "";
            item_DateFrom.CustomFormat = " ";
            item_DateTo.Text = "";
            item_DateTo.CustomFormat = " ";
            item_MadoguchiBusho.SelectedIndex = -1;
            item_ChousaKbnJibusho.Checked = false;
            item_ChousaKbnShibuShibu.Checked = false;
            item_ChousaKbnHonbuShibu.Checked = false;
            item_ChousaKbnShibuHonbu.Checked = false;
            item_Gyoumumei.Text = "";
            item_KanriBangou.Text = "";
            item_FromTo.SelectedIndex = -1;
            item_MadoguchiTantousha.Text = "";
            item_Koujikenmei.Text = "";
            item_ShimekiriSentaku.SelectedIndex = -1;
            item_Hyoujikensuu.SelectedValue = 100;
            item_ChousaKind.SelectedIndex = -1;
            item_ChousaHinmoku.Text = "";
            item_Shintyokujyoukyo.SelectedIndex = -1;

            item_TantoushaJoukyo.SelectedIndex = -1;


            //グリッドコントロールを初期化
            c1FlexGrid1.Styles.Normal.WordWrap = true;
            c1FlexGrid1.Rows[0].AllowMerging = true;
            c1FlexGrid1.AllowAddNew = false;

            if (c1FlexGrid1.Rows.Count > 1)
            {
                //グリッドクリア ヘッダー以外削除
                c1FlexGrid1.Rows.Count = 1;
            }

            warifuriFlg = false; 
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid1.BeginUpdate();

            errorCheck_initialize();

            // false：正常 true：エラー
            Boolean errorFlg = false;
            // 
            //if (item_DateFrom.CustomFormat == "" && item_DateTo.CustomFormat == "")
            //{
            //    // FromがToが大きい場合、エラー
            //    if (item_DateFrom.Value > item_DateTo.Value)
            //    {
            //        errorFlg = true;
            //        set_error("", 0);
            //        // E20002 対象項目の入力に誤りがあります。
            //        set_error(GlobalMethod.GetMessage("E20002", ""));
            //        item_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        item_DateTo.BackColor = Color.FromArgb(255, 204, 255);
            //    }
            //    else
            //    {
            //        item_DateFrom.BackColor = Color.FromArgb(255, 255, 255);
            //        item_DateTo.BackColor = Color.FromArgb(255, 255, 255);
            //    }
            //}

            // 検索期間の指定に値が入っていた場合、入力チェック
            if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() != "")
            {
                // 値の再設定
                changeShimekiribi();
                errorFlg = errorCheck_Shimekiribi();
            }

            if (errorFlg == false)
            {
                get_data();
            }
            //描画再開
            c1FlexGrid1.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void Tokumei_Activated(object sender, EventArgs e)
        {
            if (ReSearch)
            {
                get_data();
                ReSearch = false;
            }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            errorCheck_initialize();

            ClearForm();
        }

        private void get_data()
        {
            if (item_ChousaBusho.Text == "")
            {
                item_ChousaBusho.BackColor = Color.FromArgb(255, 204, 255);
                set_error("", 0);
                set_error(GlobalMethod.GetMessage("E10010", ""));
                return;
            }
            else
            {
                item_ChousaBusho.BackColor = Color.White;
            }



            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            //データ取得処理
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();


                //①旧部所年度commonValue1取得
                var comboDt = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                  "commonValue1 " +
                  "FROM " + "M_CommonMaster " +
                  "WHERE CommonMasterKye = 'OLDBUSHO_NENDO'";

                String nendo1 = item_Nendo.Text.Substring(0, 4);
                String nendo2 = nendo1;
                if (item_NendoOption3Nen.Checked)
                {
                    nendo2 = (int.Parse(nendo1) - 2).ToString();
                }

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(comboDt);

                //②commonValue1がnullなら2015とする
                int commonValue1 = 2015;
                if (comboDt.Rows.Count > 0)
                {
                    DataRow nr = comboDt.Rows[0];
                    commonValue1 = int.Parse(nr["commonValue1"].ToString());
                }

                // 年度
                int FromNendo = 0;
                int ToNendo = 0;

                if (int.TryParse(item_Nendo.SelectedValue.ToString(), out FromNendo))
                {
                    ToNendo = FromNendo + 1;
                }
                else
                {
                    FromNendo = DateTime.Today.Year;
                }
                // 3年以内にチェックしていた場合
                if (item_NendoOption3Nen.Checked)
                {
                    FromNendo = FromNendo - 2;
                    ToNendo = FromNendo + 1;
                }
                else
                {
                    ToNendo = FromNendo + 1;
                }


                // 調査員1～25取得処理
                CousainList = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT distinct " +
                    "KojinCD " +
                    ",ChousainMei " +
                    "FROM Mst_Chousain " +
                    "WHERE " +
                    "TokuchoFLG = 1 AND RetireFLG = 0 " +
                    "AND GyoumuBushoCD = '" + item_ChousaBusho.SelectedValue + "' ";

                //今日日付の年度データを検索する際は、今日有効な調査員を表示
                if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(FromNendo, 4, 1))
                {
                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today + "' ) ";
                }
                else
                {
                    //cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                    //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + FromNendo + "/4/1' ) ";
                }

                cmd.CommandText += "ORDER BY KojinCD ";

                //データ取得
                sda = new SqlDataAdapter(cmd);
                sda.Fill(CousainList);

                //③窓口情報取得
                //SQL生成
                cmd.CommandText = "SELECT distinct " +
                //" T2.MadoguchiShinchokuJoukyou " +               //0:進捗状況
                " T2.MadoguchiShinchokuJoukyou " +               //0:進捗状況
                ",T2.MadoguchiID " +                             //1:窓口ID
                ",T2.MadoguchiKanriBangou " +                    //2:管理番号
                ",CASE WHEN T2.MadoguchiUketsukeBangouEdaban is null OR T2.MadoguchiUketsukeBangouEdaban = '' then T2.MadoguchiUketsukeBangou " + // 3:特調番号 + 枝番（存在すれば）
                " ELSE T2.MadoguchiUketsukeBangou + '-' + T2.MadoguchiUketsukeBangouEdaban END AS tokuchoNo " +
                ",T2.MadoguchiHachuuKikanmei " +                 //4:発注者詳細名
                ",T2.MadoguchiTantoushaBushoCD AS MadoguchiTantoushaBushoCD " + //5:窓口部所1
                ",T7.ChousainMei AS MadoguchiTantousha " +       //6:窓口担当者1
                //",T2.MadoguchiChousaShubetsu " +                 //7:調査種別
                //",T2.MadoguchiJiishiKubun " +                    //8:実施区分
                ",CASE T2.MadoguchiChousaShubetsu WHEN 1 THEN '単品' WHEN 2 THEN '一般' WHEN 3 THEN '単契' ELSE ' ' END AS MadoguchiChousaShubetsu" +                      //7:調査種別
                ",CASE T2.MadoguchiJiishiKubun WHEN 1 THEN '実施' WHEN 2 THEN '打診中' WHEN 3 THEN '中止' ELSE ' ' END AS MadoguchiJiishiKubun" +       //8:実施区分
                ",T2.MadoguchiGyoumuMeishou " +                  //9:業務名称
                ",T2.MadoguchiKoujiKenmei " +                    //10:工事件名
                ",T2.MadoguchiChousaHinmoku " +                  //11:調査品目
                //",T2.MadoguchiShimekiribi " +                    //12:締切日
                ",T2.MadoguchiShimekiribi " +                    //12:締切日
                ",T2.MadoguchiTourokubi " +                      //13:登録日
                ",T2.MadoguchiOuenUketsukebi " +                 //14:応援受付日
                ",T12.ShibuBikouChousaBusho " +                  //15:＜部所＞
                ",T12.ShibuBikouKanriNo " +                      //16:＜部所＞_管理
                ",T12.ShibuBikou ";                              //17:＜部所＞_備考

                DataRow shibudr;
                // 調査員1～調査員25 担当者状況 18~42
                for (int i = 0; i < 25; i++)
                {
                    // 調査部所に紐づく調査員のデータがあるかないか
                    if (CousainList.Rows.Count > i)
                    {
                        shibudr = CousainList.Rows[i];
                        //cmd.CommandText += ",(select TOP 1 MadoguchiL1ChousaShinchoku from MadoguchiJouhouMadoguchiL1Chou where MadoguchiL1ChousaBushoCD = T2.MadoguchiTantoushaBushoCD AND MadoguchiL1ChousaTantoushaCD = " + shibudr["KojinCD"] + " and MadoguchiID = T2.MadoguchiID) ";
                        //cmd.CommandText += ",(select TOP 1 MadoguchiL1ChousaShinchoku from MadoguchiJouhouMadoguchiL1Chou where MadoguchiL1ChousaBushoCD = '" + item_ChousaBusho.SelectedValue.ToString() + "' AND MadoguchiL1ChousaTantoushaCD = " + shibudr["KojinCD"] + " and MadoguchiID = T2.MadoguchiID) ";
                        cmd.CommandText += ",(select TOP 1 CASE MadoguchiL1ChousaShinchoku WHEN 10 THEN '依頼' WHEN 20 THEN '調査開始' WHEN 30 THEN '見積中' WHEN 40 THEN '集計中' WHEN 50 THEN '担当者済' WHEN 60 THEN '一次検済' WHEN 70 THEN '二次検済' WHEN 80 THEN '中止' ELSE '' END from MadoguchiJouhouMadoguchiL1Chou where MadoguchiL1ChousaBushoCD = '" + item_ChousaBusho.SelectedValue.ToString() + "' AND MadoguchiL1ChousaTantoushaCD = " + shibudr["KojinCD"] + " and MadoguchiID = T2.MadoguchiID) ";
                    }
                    else
                    {
                        cmd.CommandText += ",'' ";
                    }
                }
                // 調査員1～調査員25 担当者締切日 43~67
                for (int i = 0; i < 25; i++)
                {
                    // 調査部所に紐づく調査員のデータがあるかないか
                    if (CousainList.Rows.Count > i)
                    {
                        shibudr = CousainList.Rows[i];
                        //cmd.CommandText += ",(select TOP 1 MadoguchiL1ChousaShimekiribi from MadoguchiJouhouMadoguchiL1Chou where MadoguchiL1ChousaBushoCD = T2.MadoguchiTantoushaBushoCD AND MadoguchiL1ChousaTantoushaCD = " + shibudr["KojinCD"] + " and MadoguchiID = T2.MadoguchiID) ";
                        cmd.CommandText += ",(select TOP 1 MadoguchiL1ChousaShimekiribi from MadoguchiJouhouMadoguchiL1Chou where MadoguchiL1ChousaBushoCD = '" + item_ChousaBusho.SelectedValue.ToString() + "' AND MadoguchiL1ChousaTantoushaCD = " + shibudr["KojinCD"] + " and MadoguchiID = T2.MadoguchiID) ";
                    }
                    else
                    {
                        cmd.CommandText += ",'' ";
                    }
                }

                cmd.CommandText +=
                // 68:進捗アイコンの判定用
                ", " +
                "CASE " +
                "WHEN T2.MadoguchiHoukokuzumi = 1 THEN '8' " +
                "WHEN T2.MadoguchiHoukokuzumi != 1 THEN " +
                "     CASE " +
                "         WHEN T2.MadoguchiShinchokuJoukyou = 80 THEN '6' " + //中止
                "         WHEN T2.MadoguchiShinchokuJoukyou = 70 THEN '5' " + // 二次検済
                "         WHEN T2.MadoguchiShinchokuJoukyou = 50 THEN '7' " + //担当者済
                "         WHEN T2.MadoguchiShinchokuJoukyou = 60 THEN '7' " + // 一次検済
                "     ELSE " +
                "         CASE " +
                "              WHEN T2.MadoguchiShimekiribi < '" + DateTime.Today + "' THEN '1' " +
                "              WHEN T2.MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN '2' " +
                "              WHEN T2.MadoguchiShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN '3' " +
                "         ELSE '4' " +
                "         END " +
                "     END " +
                "END " +
                "FROM MadoguchiJouhou T2 " +
                //" INNER JOIN MadoguchiJouhou T2 ON T2.MadoguchiID = T0.MadoguchiID " +
                //不具合No1339（No1096）2022/08/03 調査担当部所で部署備考を取得する必要あり
                " LEFT JOIN ShibuBikou T12 ON T12.MadoguchiID = T2.MadoguchiID AND ShibuBikouBushoKanriboBushoCD = " + item_ChousaBusho.SelectedValue.ToString() +
                //" LEFT JOIN ShibuBikou T12 ON T12.MadoguchiID = T2.MadoguchiID AND ShibuBikouBushoKanriboBushoCD = T2.MadoguchiTantoushaBushoCD " +
                
                //" LEFT JOIN Mst_Busho T3 ON T3.GyoumuBushoCD = T2.MadoguchiTantoushaBushoCD  " +
                //" LEFT JOIN Mst_Busho T4 ON T4.GyoumuBushoCD = T2.MadoguchiJutakuBushoCD  " +
                //" LEFT JOIN Mst_Chousain T5 ON T5.KojinCD = T2.MadoguchiJutakuTantoushaID  " +
                " LEFT JOIN Mst_Busho T6 ON T6.GyoumuBushoCD = T2.JutakuBushoShozokuCD  " +
                " LEFT JOIN Mst_Chousain T7 ON T7.KojinCD = T2.MadoguchiTantoushaCD  ";


                String w_jyokyou = "";
                String workdayFrom = "0";
                String workdayTo = "0";
                DateTime w_Simekiribi6 = DateTime.Today;

                // 締切日From,Toは後で計算値を設定
                DateTime w_SimekiribiFrom = DateTime.Today;
                DateTime w_SimekiribiTo = DateTime.Today;
                DateTime dateTime;

                w_Simekiribi6 = w_Simekiribi6.AddDays(6);

                // 調査担当部所 が選択されている場合、MadoguchiJouhouMadoguchiL1ChouをJOINする
                if (item_ChousaBusho.Text != "" && item_ChousaBusho.SelectedValue != null)
                {
                    cmd.CommandText +=
                   //" LEFT JOIN MadoguchiJouhouMadoguchiL1Chou T11 ON T2.MadoguchiID = T11.MadoguchiID ";

                   // 同じ部所で担当者が複数入れた場合、同一列が複数出てしまわないようにする為、以下のカタチとする
                   "  LEFT JOIN " +
                   "(select distinct MadoguchiID, MadoguchiL1ChousaBushoCD ";

                    // VIPS 20220510 課題管理表No.1317(1041) ↓1311で対応した際、「MadoguchiL1ChousaShimekiribi」のカラムが重複してしまったため修正
                    // VIPS 20220427 課題管理表No.1311(1034) ADD 担当部所情報（MadoguchiJouhouMadoguchiL1Chou）の担当者締切日（MadoguchiL1ChousaShimekiribi）を参照
                    //if (item_DateFrom.CustomFormat == "" || item_DateTo.CustomFormat == "")
                    //{
                    //    cmd.CommandText += ", MadoguchiL1ChousaShimekiribi ";
                    //}
                    if ((item_TantoushaJoukyo.Text != "" && item_TantoushaJoukyo.SelectedValue != null) || (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() != "") || item_DateFrom.CustomFormat == "" || item_DateTo.CustomFormat == "")
                    {
                        cmd.CommandText += ",MadoguchiL1ChousaShinchoku, MadoguchiL1ChousaShimekiribi ";
                    }
                    //if ((item_TantoushaJoukyo.Text != "" && item_TantoushaJoukyo.SelectedValue != null) || (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() != "")) { 
                    //    cmd.CommandText += ",MadoguchiL1ChousaShinchoku, MadoguchiL1ChousaShimekiribi ";
                    //}

                    cmd.CommandText += " from MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiL1ChousaBushoCD = " + item_ChousaBusho.SelectedValue.ToString();
                    // 担当者進捗 が選択されている場合
                    if (item_TantoushaJoukyo.Text != "" && item_TantoushaJoukyo.SelectedValue != null)
                    {
                        // VIPS　20220330　課題管理表No1294(982) ADD, CHANGE 担当者状況の追加・変更
                        // 担当者状況が複数条件の場合はテーブルで直接探さない
                        if (int.Parse(item_TantoushaJoukyo.SelectedValue.ToString()) <= 80)
                        {
                            cmd.CommandText += " AND MadoguchiL1ChousaShinchoku = " + item_TantoushaJoukyo.SelectedValue.ToString() + " ";
                        }
                    }
                    // 未割振り
                    if (warifuriFlg)
                    {
                        //cmd.CommandText += " AND ISNULL(MadoguchiL1ChousaShinchoku, 0) = 0 ";
                        cmd.CommandText += " AND ISNULL(MadoguchiL1ChousaShinchoku, 0) = 10 ";
                    }
                    // 進捗状況
                    if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() != "")
                    {
                        switch (item_Shintyokujyoukyo.SelectedValue.ToString())
                        {
                            // 2次検証済み
                            case "1":
                                w_jyokyou = "3";
                                break;
                            // 担当者済
                            case "2":
                                w_jyokyou = "2";
                                break;
                            // 締切日が1週間をこえる もしくは中止
                            case "3":
                                break;
                            // 締切日が1週間以内
                            case "4":
                                workdayFrom = "-1";
                                workdayTo = "7";
                                break;
                            // 締切日が3日以内
                            case "5":
                                workdayFrom = "-1";
                                workdayTo = "3";
                                break;
                            // 締切日が超過
                            case "6":
                                workdayTo = "-1";
                                break;
                            default:
                                break;
                        }

                        // 締切日付計算
                        w_SimekiribiFrom = DateTime.Today.AddDays(int.Parse(workdayFrom));
                        w_SimekiribiTo = DateTime.Today.AddDays(int.Parse(workdayTo));

                        // 進捗状況のコンボボックスの条件　二次検証済み　か　完了かどうか
                        if (w_jyokyou != "" && (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "1"))
                        {
                            cmd.CommandText += "and MadoguchiL1ChousaShinchoku = 70 ";
                        }
                        // 締め切りまで1週間以上、または中止
                        if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "3")
                        {
                            //cmd.CommandText += "and ((     (ISNULL(MadoguchiL1ChousaShinchoku, 0) = 0 or MadoguchiL1ChousaShinchoku = 10) ";
                            cmd.CommandText += "and ((     (ISNULL(MadoguchiL1ChousaShinchoku, 0) = 0 or MadoguchiL1ChousaShinchoku < 50) ";
                            cmd.CommandText += "      and (MadoguchiL1ChousaShimekiribi >= '" + w_Simekiribi6 + "')) ";
                            cmd.CommandText += "or　(MadoguchiL1ChousaShinchoku = 80)) ";

                        }
                        // 2次検証済ではない　のと　完了 中止ではないのが条件
                        if (item_Shintyokujyoukyo.SelectedValue != null && (item_Shintyokujyoukyo.SelectedValue.ToString() == "4" || item_Shintyokujyoukyo.SelectedValue.ToString() == "5"))
                        {
                            //cmd.CommandText += "and ((ISNULL(MadoguchiL1ChousaShinchoku, 0) = 0 or MadoguchiL1ChousaShinchoku = 10) ";
                            cmd.CommandText += "and ((ISNULL(MadoguchiL1ChousaShinchoku, 0) = 0 or MadoguchiL1ChousaShinchoku < 50) ";
                            cmd.CommandText += "and (MadoguchiL1ChousaShimekiribi >= '" + w_SimekiribiFrom + "' and MadoguchiL1ChousaShimekiribi <= '" + w_SimekiribiTo + "')) ";
                        }
                        // 超過
                        if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "6")
                        {
                            //cmd.CommandText += "and ((ISNULL(MadoguchiL1ChousaShinchoku, 0) = 0 or MadoguchiL1ChousaShinchoku = 10) ";
                            cmd.CommandText += "and ((ISNULL(MadoguchiL1ChousaShinchoku, 0) = 0 or MadoguchiL1ChousaShinchoku < 50) ";
                            cmd.CommandText += "and (MadoguchiL1ChousaShimekiribi <= '" + w_SimekiribiTo + "')) ";
                        }
                    }

                    cmd.CommandText += 
                    ") T11 " +
                   "    ON T2.MadoguchiID = T11.MadoguchiID ";
                }

                cmd.CommandText +=
               "WHERE MadoguchiTourokuNendo <= '" + nendo1 + "' and MadoguchiTourokuNendo >= '" + nendo2 + "' " +
                "  AND MadoguchiDeleteFlag != 1 " +
                " AND T2.MadoguchiSystemRenban > 0 ";

                //String w_jyokyou = "";
                //String workdayFrom = "0";
                //String workdayTo = "0";
                //DateTime w_Simekiribi6 = DateTime.Today;

                //// 締切日From,Toは後で計算値を設定
                //DateTime w_SimekiribiFrom = DateTime.Today;
                //DateTime w_SimekiribiTo = DateTime.Today;
                //DateTime dateTime;

                //w_Simekiribi6 = w_Simekiribi6.AddDays(6);

                String Where = "";

                //// 進捗状況
                //if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() != "")
                //{
                //    switch (item_Shintyokujyoukyo.SelectedValue.ToString())
                //    {
                //        // 2次検証済み
                //        case "1":
                //            w_jyokyou = "3";
                //            break;
                //        // 担当者済
                //        case "2":
                //            w_jyokyou = "2";
                //            break;
                //        // 締切日が1週間をこえる もしくは中止
                //        case "3":
                //            break;
                //        // 締切日が1週間以内
                //        case "4":
                //            workdayFrom = "-1";
                //            workdayTo = "7";
                //            break;
                //        // 締切日が3日以内
                //        case "5":
                //            workdayFrom = "-1";
                //            workdayTo = "3";
                //            break;
                //        // 締切日が超過
                //        case "6":
                //            workdayTo = "-1";
                //            break;
                //        default:
                //            break;
                //    }

                //    // 締切日付計算
                //    w_SimekiribiFrom = DateTime.Today.AddDays(int.Parse(workdayFrom));
                //    w_SimekiribiTo = DateTime.Today.AddDays(int.Parse(workdayTo));

                //    // 進捗状況のコンボボックスの条件　二次検証済み　か　完了かどうか
                //    if (w_jyokyou != "" && (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "1"))
                //    {
                //        Where += "and T11.MadoguchiL1ChousaShinchoku = 70 ";
                //    }
                //    // 締め切りまで1週間以上、または中止
                //    if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "3")
                //    {
                //        Where += "and ((ISNULL(T11.MadoguchiL1ChousaShinchoku, 0) = 0 or T11.MadoguchiL1ChousaShinchoku = 10) ";
                //        Where += "and (T11.MadoguchiL1ChousaShimekiribi >= '" + w_Simekiribi6 + "')) ";
                //        Where += "or　(T11.MadoguchiL1ChousaShinchoku = 80) ";

                //    }
                //    // 2次検証済ではない　のと　完了 中止ではないのが条件
                //    if (item_Shintyokujyoukyo.SelectedValue != null && (item_Shintyokujyoukyo.SelectedValue.ToString() == "4" || item_Shintyokujyoukyo.SelectedValue.ToString() == "5"))
                //    {
                //        Where += "and ((ISNULL(T11.MadoguchiL1ChousaShinchoku, 0) = 0 or T11.MadoguchiL1ChousaShinchoku = 10) ";
                //        Where += "and (T11.MadoguchiL1ChousaShimekiribi >= '" + w_SimekiribiFrom + "' and T11.MadoguchiL1ChousaShimekiribi <= '" + w_SimekiribiTo + "')) ";
                //    }
                //    // 超過
                //    if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "6")
                //    {
                //        Where += "and ((ISNULL(T11.MadoguchiL1ChousaShinchoku, 0) = 0 or T11.MadoguchiL1ChousaShinchoku = 10) ";
                //        Where += "and (T11.MadoguchiL1ChousaShimekiribi <= '" + w_SimekiribiTo + "')) ";
                //    }
                //}

                // 特調番号
                if (item_TokuchoBangou.Text != "")
                {
                    Where += "and (CASE WHEN T2.MadoguchiUketsukeBangouEdaban is null OR T2.MadoguchiUketsukeBangouEdaban = '' then T2.MadoguchiUketsukeBangou " +
                             " ELSE T2.MadoguchiUketsukeBangou + '-' + T2.MadoguchiUketsukeBangouEdaban END) COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_TokuchoBangou.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 業務名称
                if (item_Gyoumumei.Text != "")
                {
                    Where += "and MadoguchiGyoumuMeishou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_Gyoumumei.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 工事件名
                if (item_Koujikenmei.Text != "")
                {
                    Where += "and MadoguchiKoujiKenmei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_Koujikenmei.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 管理番号
                if (item_KanriBangou.Text != "")
                {
                    Where += "and MadoguchiKanriBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_KanriBangou.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 調査品目
                if (item_ChousaHinmoku.Text != "")
                {
                    Where += "and MadoguchiChousaHinmoku COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_ChousaHinmoku.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                //// 調査区分 自部所
                //if (item_ChousaKbnJibusho.Checked)
                //{
                //    Where += "and MadoguchiChousaKubunJibusho = 1 ";
                //}
                //// 調査区分 支→支
                //if (item_ChousaKbnShibuShibu.Checked)
                //{
                //    Where += "and MadoguchiChousaKubunShibuShibu = 1 ";
                //}
                //// 調査区分 本→支
                //if (item_ChousaKbnHonbuShibu.Checked)
                //{
                //    Where += "and MadoguchiChousaKubunHonbuShibu = 1 ";
                //}
                //// 調査区分 支→本
                //if (item_ChousaKbnShibuHonbu.Checked)
                //{
                //    Where += "and MadoguchiChousaKubunShibuHonbu = 1 ";
                //}
                if (item_ChousaKbnJibusho.Checked || item_ChousaKbnShibuShibu.Checked || item_ChousaKbnHonbuShibu.Checked || item_ChousaKbnShibuHonbu.Checked)
                {
                    // OR追加フラグ true:OR追加
                    //Boolean OrAddFlg = false;

                    //cmd.CommandText += "AND (";
                    ////調査区分　自部所
                    //if (item_ChousaKbnJibusho.Checked)
                    //{
                    //    cmd.CommandText += " MadoguchiChousaKubunJibusho = 1 ";
                    //    OrAddFlg = true;
                    //}

                    ////調査区分　支部→支部
                    //if (item_ChousaKbnShibuShibu.Checked)
                    //{
                    //    if (OrAddFlg)
                    //    {
                    //        cmd.CommandText += "OR ";
                    //    }
                    //    cmd.CommandText += " MadoguchiChousaKubunShibuShibu = 1 ";
                    //    OrAddFlg = true;
                    //}

                    ////調査区分　本部→支部
                    //if (item_ChousaKbnHonbuShibu.Checked)
                    //{
                    //    if (OrAddFlg)
                    //    {
                    //        cmd.CommandText += "OR ";
                    //    }
                    //    cmd.CommandText += " MadoguchiChousaKubunHonbuShibu = 1 ";
                    //    OrAddFlg = true;
                    //}

                    ////調査区分　支部→本部
                    //if (item_ChousaKbnShibuHonbu.Checked)
                    //{
                    //    if (OrAddFlg)
                    //    {
                    //        cmd.CommandText += "OR ";
                    //    }
                    //    cmd.CommandText += " MadoguchiChousaKubunShibuHonbu = 1 ";
                    //    OrAddFlg = true;
                    //}
                    //cmd.CommandText += ")";

                    cmd.CommandText += "AND (";
                    //調査区分　自部所
                    if (item_ChousaKbnJibusho.Checked)
                    {
                        cmd.CommandText += " MadoguchiChousaKubunJibusho = 1 ";
                    }
                    else
                    {
                        cmd.CommandText += " MadoguchiChousaKubunJibusho = 0 ";
                    }

                    cmd.CommandText += "AND ";
                    //調査区分　支部→支部
                    if (item_ChousaKbnShibuShibu.Checked)
                    {
                        cmd.CommandText += " MadoguchiChousaKubunShibuShibu = 1 ";
                    }
                    else
                    {
                        cmd.CommandText += " MadoguchiChousaKubunShibuShibu = 0 ";
                    }
                    cmd.CommandText += "AND ";
                    //調査区分　本部→支部
                    if (item_ChousaKbnHonbuShibu.Checked)
                    {
                        cmd.CommandText += " MadoguchiChousaKubunHonbuShibu = 1 ";
                    }
                    else
                    {
                        cmd.CommandText += " MadoguchiChousaKubunHonbuShibu = 0 ";
                    }

                    cmd.CommandText += "AND ";
                    //調査区分　支部→本部
                    if (item_ChousaKbnShibuHonbu.Checked)
                    {
                        cmd.CommandText += " MadoguchiChousaKubunShibuHonbu = 1 ";
                    }
                    else
                    {
                        cmd.CommandText += " MadoguchiChousaKubunShibuHonbu = 0 ";
                    }
                    cmd.CommandText += ")";
                }

                // 調査担当者への締切日
                if (item_DateFrom.CustomFormat == "")
                {
                    //VIPS 20220427 課題管理表No.1311(1034) CHANGE 担当部所情報（MadoguchiJouhouMadoguchiL1Chou）の担当者締切日（MadoguchiL1ChousaShimekiribi）を参照
                    Where += "and T11.MadoguchiL1ChousaShimekiribi >= '" + item_DateFrom.Text + "' ";
                    //Where += "and T2.MadoguchiShimekiribi >= '" + item_DateFrom.Text + "' ";
                }
                if (item_DateTo.CustomFormat == "")
                {
                    //VIPS 20220427 課題管理表No.1311(1034) CHANGE 担当部所情報（MadoguchiJouhouMadoguchiL1Chou）の担当者締切日（MadoguchiL1ChousaShimekiribi）を参照
                    Where += "and T11.MadoguchiL1ChousaShimekiribi <= '" + item_DateTo.Text + "' ";
                    //Where += "and T2.MadoguchiShimekiribi <= '" + item_DateTo.Text + "' ";
                }



                // 発注者名・課名
                if (item_HachushaKamei.Text != "")
                {
                    Where += "and T2.MadoguchiHachuuKikanmei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_HachushaKamei.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 調査種別
                if (item_ChousaKind.SelectedValue != null && item_ChousaKind.Text != " ")
                {
                    Where += "and MadoguchiChousaShubetsu = " + item_ChousaKind.SelectedValue.ToString() + " ";
                }

                // 窓口部所
                if (item_MadoguchiBusho.SelectedValue != null && item_MadoguchiBusho.Text != "")
                {
                    if (item_MadoguchiBusho.SelectedIndex != 0)
                    {
                        Where += "and MadoguchiTantoushaBushoCD = " + item_MadoguchiBusho.SelectedValue.ToString() + " ";
                    }
                }
                // 窓口担当者
                if (item_MadoguchiTantousha.Text != "")
                {
                    Where += "and T7.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(item_MadoguchiTantousha.Text, 1, 0) + "%' ESCAPE '\\' ";
                }
                // 調査担当部所
                if (item_ChousaBusho.Text != "" && item_ChousaBusho.SelectedValue != null)
                {
                    Where += "and T11.MadoguchiL1ChousaBushoCD = " + item_ChousaBusho.SelectedValue.ToString() + " ";
                }

                //担当者状況
                if (item_TantoushaJoukyo.Text != "" && item_TantoushaJoukyo.SelectedValue != null)
                {
                    // VIPS　20220330　課題管理表No1294(982) ADD 担当者状況の分岐を追加
                    int tantoushaJoukyo = int.Parse(item_TantoushaJoukyo.SelectedValue.ToString());

                    //調査中
                    if (tantoushaJoukyo == 200)
                    {
                        Where += "and (T11.MadoguchiL1ChousaShinchoku = 20 " +
                        " OR T11.MadoguchiL1ChousaShinchoku = 30)";
                    }
                    //依頼・調査中
                    else if (tantoushaJoukyo == 220)
                    {
                        Where += "and (T11.MadoguchiL1ChousaShinchoku = 10 " +
                        " OR T11.MadoguchiL1ChousaShinchoku = 20" +
                        " OR T11.MadoguchiL1ChousaShinchoku = 30)";
                    }
                    //依頼・調査中・担当者済
                    else if (tantoushaJoukyo == 230)
                    {
                        Where += "and (T11.MadoguchiL1ChousaShinchoku = 10 " +
                        " OR T11.MadoguchiL1ChousaShinchoku = 20" +
                        " OR T11.MadoguchiL1ChousaShinchoku = 30" +
                        " OR T11.MadoguchiL1ChousaShinchoku = 40" +
                        " OR T11.MadoguchiL1ChousaShinchoku = 50)";
                    }
                    //検証中
                    else if (tantoushaJoukyo == 300)
                    {
                        Where += "and (T11.MadoguchiL1ChousaShinchoku = 50 " +
                        " OR T11.MadoguchiL1ChousaShinchoku = 60)";
                    }
                    //検証中・二次検済
                    else if (tantoushaJoukyo == 310)
                    {
                        Where += "and (T11.MadoguchiL1ChousaShinchoku = 50 " +
                        " OR T11.MadoguchiL1ChousaShinchoku = 60" +
                        " OR T11.MadoguchiL1ChousaShinchoku = 70)";
                    }
                    //それ以外　依頼, 調査開始, 見積中, 集計中, 担当者済, 一次検済, 二次検済
                    else
                    {
                        Where += "and T11.MadoguchiL1ChousaShinchoku = " + item_TantoushaJoukyo.SelectedValue.ToString() + " ";
                        Console.WriteLine(item_TantoushaJoukyo.SelectedValue);
                    }

                }
                // 未割振り
                if (warifuriFlg)
                {
                    //cmd.CommandText += " AND ISNULL(MadoguchiL1ChousaShinchoku, 0) = 0 ";
                    cmd.CommandText += " AND ISNULL(MadoguchiL1ChousaShinchoku, 0) = 10 ";
                }

                // 1206 超過の場合は、報告済みは除外する
                if (item_Shintyokujyoukyo.SelectedValue != null && item_Shintyokujyoukyo.SelectedValue.ToString() == "6")
                {
                    cmd.CommandText += "and MadoguchiHoukokuzumi = 0 ";
                }


                if (Where != "")
                {
                    cmd.CommandText += Where;
                }
                //特命課長の並び順　窓口の締切日（降順）、第２キー特調番号(枝番含む)（昇順）、第３キー管理番号（降順）
                cmd.CommandText += "ORDER BY ";
                cmd.CommandText += " T2.MadoguchiShimekiribi DESC ,CASE WHEN T2.MadoguchiUketsukeBangouEdaban is null OR T2.MadoguchiUketsukeBangouEdaban = '' then T2.MadoguchiUketsukeBangou ELSE T2.MadoguchiUketsukeBangou + '-' + T2.MadoguchiUketsukeBangouEdaban END ,MadoguchiKanriBangou DESC";

                Console.WriteLine(cmd.CommandText);
                GlobalMethod.outputLogger("Search_Tokumei", "開始", "GetMadoguchiJouhou", UserInfos[1]);
                //データ取得
                sda = new SqlDataAdapter(cmd);

                set_error("", 0);
                ListData = new DataTable();
                ListData.Clear();
                sda.Fill(ListData);

                // ヘッダー部だけに設定
                c1FlexGrid1.Rows.Count = 1;

                //行数決め 基本は50
                int rowsCount = 50;
                if (ListData.Rows.Count <= 50)
                {
                    rowsCount = ListData.Rows.Count;
                }
                // 0件の場合
                if (ListData.Rows.Count == 0)
                {
                    set_error("", 0);
                    // I20001:該当データは0件です。
                    set_error(GlobalMethod.GetMessage("I20001", ""));
                }

                Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / int.Parse(item_Hyoujikensuu.Text.Replace("件", "")))).ToString();
                Paging_now.Text = (1).ToString();
                set_data(int.Parse(Paging_now.Text));
                set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

                // 取得件数を表示
                if (ListData != null)
                {
                    Grid_Num.Text = "(" + ListData.Rows.Count + ")";
                }
                else
                {
                    // 念の為、ListDataがNullの時は0を表示
                    Grid_Num.Text = "(0)";
                }
            }
            //未割振りフラグを初期化
            warifuriFlg = false;
        }


        private void set_data(int pagenum)
        {
            // まずはヘッダーだけに
            c1FlexGrid1.Rows.Count = 1;
            // 表示するデータ数 + header
            int viewnum = int.Parse(item_Hyoujikensuu.Text.Replace("件", ""));
            // 表示するページ数 
            int startrow = (pagenum - 1) * viewnum;
            // 表示する行
            int addnum = ListData.Rows.Count - startrow;
            if (addnum > viewnum)
            {
                addnum = viewnum;
            }
            // 表示したい行分ループ
            for (int r = 0; r < addnum; r++)
            {
                c1FlexGrid1.Rows.Add();
                // 進捗状況
                c1FlexGrid1[r + 1, 1] = ListData.Rows[startrow + r][68];
                DateTime Simekiribi = DateTime.Today;
                DateTime TantouSimekiribi;
                DateTime.TryParse(ListData.Rows[startrow + r][12].ToString(), out Simekiribi);

                for (int i = 1; i < c1FlexGrid1.Cols.Count - 2; i++)
                {
                    c1FlexGrid1[r + 1, i + 2] = ListData.Rows[startrow + r][i];
                    if (i + 2 >= 20)
                    {
                        //switch (ListData.Rows[startrow + r][i].ToString())
                        //{
                        //    //case "10":
                        //    case "依頼":
                        //        // ピンク背景
                        //        c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                        //        break;
                        //    //case "40":
                        //    case "集計中":
                        //        // 薄緑背景
                        //        c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.BackColor = Color.FromArgb(180, 255, 180);
                        //        break;
                        //    //case "50":
                        //    case "担当者済":
                        //        // 黄色背景
                        //        c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.BackColor = Color.FromArgb(255, 255, 100);
                        //        break;
                        //    default:
                        //        c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.BackColor = Color.White;
                        //        break;
                        //}

                        // 691対応
                        // ①通常調査開始～見積中は背景色　黄色
                        // ②担当者済みで緑
                        // ③文字色は、担当部所画面で締切日が変更、取り込まれた締切日で赤
                        switch (ListData.Rows[startrow + r][i].ToString())
                        {
                            case "依頼":
                            case "調査開始":
                            case "見積中":
                            case "集計中":
                                // 黄色背景
                                c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.BackColor = Color.FromArgb(255, 255, 100);
                                break;
                            case "担当者済":
                                // 薄緑背景
                                c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.BackColor = Color.FromArgb(180, 255, 180);
                                break;
                            default:
                                c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.BackColor = Color.White;
                                break;
                        }

                        // 担当者の締切日表示判定 窓口情報の締切日（調査担当者への締切日）よりも担当部所の締切日が未来の場合に表示する
                        if (DateTime.TryParse(ListData.Rows[startrow + r][i + 25].ToString(), out TantouSimekiribi) 
                            //&& TantouSimekiribi < Simekiribi)
                            && TantouSimekiribi > Simekiribi)
                        {
                            c1FlexGrid1[r + 1, i + 2] = c1FlexGrid1.GetDataDisplay(r + 1, i + 2) + Environment.NewLine + TantouSimekiribi.ToString();
                            c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.ForeColor = Color.Red;
                        }
                        else
                        {
                            c1FlexGrid1.GetCellRange(r + 1, i + 2).StyleNew.ForeColor = Color.Black;
                        }
                    }
                }
                // 編集アイコンを表示する為に、1をセット
                c1FlexGrid1[r + 1, 2] = "1";

                // 調査員
                if (r == 0)
                {
                    // クエリ：17:調査員1～41:調査員25
                    // Grid：20:調査員1～44:調査員25
                    for (int k = 0; CousainList.Rows.Count > k; k++)
                    {
                        DataRow dr2 = CousainList.Rows[k];
                        // BushokanriboKameiRaku の頭2文字をGridのヘッダーにセットする
                        c1FlexGrid1[0, k + 20] = dr2[1];
                    }

                    // ソートマーク
                    C1.Win.C1FlexGrid.CellRange cr;
                    Bitmap bmp1 = new Bitmap("Resource/Image/SortIconDefalt.png");
                    Bitmap bmpSort = new Bitmap(bmp1, bmp1.Width / 6, bmp1.Height / 6);
                    cr = c1FlexGrid1.GetCellRange(0, 1);
                    cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
                    cr.Image = bmpSort;
                    Bitmap bmp2 = new Bitmap("Resource/Image/header_blank.png");
                    for (int i = 0; i < 25; i++)
                    {
                        cr = c1FlexGrid1.GetCellRange(0, i + 20);
                        cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
                        // 調査部所に紐づく調査員のデータがあるかないか
                        if (CousainList.Rows.Count > i)
                        {
                            cr.Image = bmpSort;
                        }
                        else
                        {
                            cr.Image = bmp2;
                        }
                    }

                    DataTable BushoKaMeidt;
                    string BushoKaMei = "";
                    // 調査担当者が空でない場合（必須なので基本空はあり得ない）
                    if (item_ChousaBusho.Text != "")
                    {
                        BushoKaMeidt = GlobalMethod.getData("BushokanriboKameiRaku", "BushokanriboKameiRaku", "Mst_Busho", "GyoumuBushoCD = '" + item_ChousaBusho.SelectedValue.ToString() + "' ");

                        if(BushoKaMeidt != null && BushoKaMeidt.Rows.Count > 0) 
                        { 
                            BushoKaMei = BushoKaMeidt.Rows[0][0].ToString();
                        }
                    }

                    c1FlexGrid1[0, 17] = BushoKaMei;
                    c1FlexGrid1[0, 18] = "管理_" + BushoKaMei;
                    c1FlexGrid1[0, 19] = "部所_" + BushoKaMei;

                }
            }
            c1FlexGrid1.AllowAddNew = false;
        }


        // エラーメッセージ表示
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

        // 先頭ページへ
        private void Top_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
        }

        // 一つ前ページへ
        private void Previous_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_now.Text) - 1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
        }

        // 次のページへ
        private void After_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_now.Text) + 1).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
        }

        // ラストページへ
        private void End_Page_Click(object sender, EventArgs e)
        {
            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            set_data(int.Parse(Paging_now.Text));
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));
        }
        // ページングアイコンのON/OFF切替
        private void set_page_enabled(int now, int last)
        {
            GlobalMethod.outputLogger("Paging_Tokumei", "ページ:" + now, "GridAll", UserInfos[1]);
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

        // コンボボックスに値を表示するイベント、コンボボックスのDrawItemイベントに設定する
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

        private void item_Nendo_Validated(object sender, EventArgs e)
        {
            set_combo_shibu(item_Nendo.SelectedValue.ToString());
        }

        private void item_MadoguchiTantousha_icon_Click(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //選択されている年度を条件に調査員プロンプトを表示
            if (item_Nendo.Text != "")
            {
                form.nendo = item_Nendo.SelectedValue.ToString();
            }
            form.program = "madoguchi";
            form.Busho = null;
            if (item_MadoguchiBusho.Text != "")
            {
                form.Busho = item_MadoguchiBusho.SelectedValue.ToString();
            }
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item_MadoguchiTantousha.Text = form.ReturnValue[1];
                item_MadoguchiBusho.SelectedValue = form.ReturnValue[2];
            }
        }

        private void button_MiWarihuri_Click(object sender, EventArgs e)
        {
            Boolean errorFlg = false;
            errorCheck_initialize();
            // 検索期間の指定に値が入っていた場合
            if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() != "")
            {
                changeShimekiribi();
                errorFlg = errorCheck_Shimekiribi();
            }
            item_TantoushaJoukyo.SelectedValue = 10;
            if (!errorFlg)
            {
                warifuriFlg = true;
                get_data();
            }
        }

        // 日付系の値変更（CustomFormatを空文字にすると表示される）
        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";
        }

        // 日付系のDeleteボタン押された場合に非表示にする
        private void dateTimePicker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Delete)
            {
                ((DateTimePicker)sender).Text = "";
                ((DateTimePicker)sender).CustomFormat = " ";
            }
        }


        // 検索期間の指定変更
        private void Kikansitei_TextChanged(object sender, EventArgs e)
        {
            //// 日付のFrom
            //label_DateFrom.BackColor = Color.CadetBlue;

            //if (item_FromTo.SelectedValue != null)
            //{
            //    // 締め日の選択を空に
            //    item_ShimekiriSentaku.SelectedValue = -1;
            //}
            //// 1:以前
            //if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() == "1")
            //{
            //    // 日付Fromが選択されているか
            //    if (item_DateFrom.CustomFormat == "")
            //    {
            //        // FromをToにコピー
            //        item_DateTo.Value = item_DateFrom.Value;
            //        item_DateTo.CustomFormat = "";
            //    }
            //    else
            //    {
            //        // Toが空だった場合
            //        if (item_DateTo.CustomFormat != "")
            //        {
            //            label_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            label_DateFrom.BackColor = Color.CadetBlue;
            //        }
            //    }
            //    // Fromを消す
            //    item_DateFrom.CustomFormat = " ";

            //}
            //// 2:当日の場合
            //else if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() == "2")
            //{
            //    // 日付Fromが選択されているか
            //    if (item_DateFrom.CustomFormat == "")
            //    {
            //        // FromをToにコピー
            //        item_DateTo.Value = item_DateFrom.Value;
            //        item_DateTo.CustomFormat = "";
            //    }
            //    else
            //    {
            //        // Toが空だった場合
            //        if (item_DateTo.CustomFormat != "")
            //        {
            //            label_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            label_DateFrom.BackColor = Color.CadetBlue;
            //            item_DateFrom.Value = item_DateTo.Value;
            //        }
            //    }
            //}
            //// 3:一週間
            //else if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() == "3")
            //{
            //    // 日付Fromが選択されているか
            //    if (item_DateFrom.CustomFormat == "")
            //    {
            //        // Fromに6日を足した日をToにセット
            //        DateTime dateTime = item_DateFrom.Value;
            //        dateTime = dateTime.AddDays(6);
            //        // FromをToにコピー
            //        item_DateTo.Value = dateTime;
            //        item_DateTo.CustomFormat = "";
            //    }
            //    else
            //    {
            //        // Toが空だった場合
            //        if (item_DateTo.CustomFormat != "")
            //        {
            //            label_DateFrom.BackColor = Color.FromArgb(255, 204, 255);
            //        }
            //        else
            //        {
            //            DateTime dateTime = item_DateTo.Value;
            //            // Toに6日を引いた日をFromにセット
            //            dateTime = dateTime.AddDays(-6);
            //            label_DateFrom.BackColor = Color.CadetBlue;
            //            item_DateFrom.Value = dateTime;
            //        }
            //    }
            //}

            errorCheck_initialize();

            // 検索期間の指定に値が入っていた場合
            if (item_FromTo.SelectedValue != null && item_FromTo.SelectedValue.ToString() != "")
            {
                changeShimekiribi();
            }
        }

        // 締め日の選択
        private void item_ShimekiriSentaku_TextChanged(object sender, EventArgs e)
        {
            //// 未選択の場合
            //if (item_ShimekiriSentaku.Text == "" )
            //{
            //    // 検索期間の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
            //    label_DateFrom.BackColor = Color.CadetBlue;
            //}

            //// 1:本日の締めは？
            //if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "1")
            //{
            //    // 本日日付を設定
            //    dateTime = DateTime.Today;
            //    item_DateFrom.Value = dateTime;
            //    item_DateFrom.CustomFormat = "";
            //    item_DateTo.Value = dateTime;
            //    item_DateFrom.CustomFormat = "";
            //    // 検索期間の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
            //}
            //// 2:今週の締めは？
            //else if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "2")
            //{
            //    // 本日日付を設定
            //    dateTime = DateTime.Today;
            //    // 曜日 0:月曜 ～ 6:日曜
            //    DayOfWeek dayOfWeek = dateTime.DayOfWeek;
            //    int dow = getDayOfWeek(dateTime);
            //    dow = -1 * dow;
            //    item_DateFrom.Value = dateTime.AddDays(dow);
            //    item_DateFrom.CustomFormat = "";
            //    item_DateTo.Value = dateTime.AddDays(dow + 6);
            //    item_DateFrom.CustomFormat = "";
            //    // 検索期間の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
            //}
            //// 3:来週の締めは？
            //else if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "3" )
            //{
            //    // 本日日付を設定
            //    dateTime = DateTime.Today;
            //    // 曜日 0:月曜 ～ 6:日曜
            //    DayOfWeek dayOfWeek = dateTime.DayOfWeek;
            //    int dow = getDayOfWeek(dateTime);
            //    dow = -1 * dow;
            //    item_DateFrom.Value = dateTime.AddDays(7 + dow);
            //    item_DateFrom.CustomFormat = "";
            //    item_DateTo.Value = dateTime.AddDays(7 + dow + 6);
            //    item_DateFrom.CustomFormat = "";
            //    // 検索期間の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
            //}
            //// 4:前日の締めは？
            //else if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "4")
            //{
            //    // 本日日付を設定
            //    dateTime = DateTime.Today.AddDays(-1);
            //    item_DateFrom.Value = dateTime;
            //    item_DateFrom.CustomFormat = "";
            //    item_DateTo.Value = dateTime;
            //    item_DateFrom.CustomFormat = "";
            //    // 検索期間の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
            //}
            //// 5:翌日の締めは？
            //else if (item_ShimekiriSentaku.SelectedValue != null && item_ShimekiriSentaku.SelectedValue.ToString() == "5")
            //{
            //    // 本日日付を設定
            //    dateTime = DateTime.Today.AddDays(1);
            //    item_DateFrom.Value = dateTime;
            //    item_DateFrom.CustomFormat = "";
            //    item_DateTo.Value = dateTime;
            //    item_DateFrom.CustomFormat = "";
            //    // 検索期間の指定を空に
            //    item_FromTo.SelectedValue = -1;
            //    item_FromTo.BackColor = Color.FromArgb(255, 255, 255);
            //}

            // 締切日の選択に値が入っていた場合
            if (item_ShimekiriSentaku.SelectedValue != null)
            {
                DateTime dateTimeFrom = DateTime.MinValue;
                DateTime dateTimeTo = DateTime.MinValue;
                int dow = 0;
                switch (item_ShimekiriSentaku.SelectedValue.ToString())
                {
                    case "1":   // 1:本日の締めは？
                        // 締切日のfrom、toに本日の日付をセットする
                        dateTimeFrom = DateTime.Today;
                        dateTimeTo = dateTimeFrom;
                        break;
                    case "2":   // 2:今週の締めは？
                        // 締切日のfromに本日を含む週の月曜日をセットする
                        // 締切日のtoに本日を含む週の日曜日をセットする
                        dow = -1 * getDayOfWeek(DateTime.Today);
                        dateTimeFrom = DateTime.Today.AddDays(dow);
                        dateTimeTo = dateTimeFrom.AddDays(6);
                        break;
                    case "3":   // 3:来週の締めは？
                        // 締切日のfromに本日+7を含む週の月曜日をセットする
                        // 締切日のtoに本日+7を含む週の日曜日をセットする
                        dow = -1 * getDayOfWeek(DateTime.Today.AddDays(7));
                        dateTimeFrom = DateTime.Today.AddDays(7 + dow);
                        dateTimeTo = dateTimeFrom.AddDays(6);
                        break;
                    case "4":   // 4:前日の締めは？
                        // 締切日のfrom、toに本日-1の日付をセットする
                        dateTimeFrom = DateTime.Today.AddDays(-1);
                        dateTimeTo = dateTimeFrom;
                        break;
                    case "5":   // 5:翌日の締めは？
                        // 締切日のfrom、toに本日+1の日付をセットする
                        dateTimeFrom = DateTime.Today.AddDays(1);
                        dateTimeTo = dateTimeFrom;
                        break;
                    default:
                        break;
                }

                // 締切日がセットされていた場合、編集する
                if (dateTimeFrom != DateTime.MinValue && dateTimeTo != DateTime.MinValue)
                {
                    // 締切日from
                    item_DateFrom.Value = dateTimeFrom.Date;
                    item_DateFrom.CustomFormat = "";
                    // 締切日to
                    item_DateTo.Value = dateTimeTo.Date;
                    item_DateFrom.CustomFormat = "";
                }

                // 検索期間の指定を空にする
                item_FromTo.SelectedValue = -1;     // Kikansitei_TextChanged が動く

            }
        }

        // 曜日の数値を返却する
        private int getDayOfWeek(DateTime dateTime)
        {
            DayOfWeek dow = dateTime.DayOfWeek;
            int num = 0;

            switch (dow)
            {
                case DayOfWeek.Monday:
                    // 月曜日
                    num = 0;
                    break;
                case DayOfWeek.Tuesday:
                    // 火曜日
                    num = 1;
                    break;
                case DayOfWeek.Wednesday:
                    // 水曜日
                    num = 2;
                    break;
                case DayOfWeek.Thursday:
                    // 木曜日
                    num = 3;
                    break;
                case DayOfWeek.Friday:
                    // 金曜日
                    num = 4;
                    break;
                case DayOfWeek.Saturday:
                    // 土曜日
                    num = 5;
                    break;
                case DayOfWeek.Sunday:
                    // 日曜日
                    num = 6;
                    break;
            }
            return num;
        }
        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void c1FlexGrid1_OwnerDrawCell(object sender, C1.Win.C1FlexGrid.OwnerDrawCellEventArgs e)
        {
            //if (e.Row > 1 && e.Col >= 20)
            //{
            //    switch(c1FlexGrid1[e.Row, e.Col])
            //    {
            //        case 10:
            //            c1FlexGrid1.GetCellRange(e.Row, e.Col).StyleNew.BackColor = Color.Yellow;
            //            break;

            //    }
            //    c1FlexGrid1.GetCellRange(e.Row, e.Col).StyleNew.BackColor = Color.Blue;
            //}
        }

        //ヘッダー特命課長ボタン
        private void btnTokumei_Click(object sender, EventArgs e)
        {
            ////Tokumei form = new Tokumei();
            ////form.UserInfos = this.UserInfos;
            ////form.Show();
            ////this.Close();
            //Form f = null;
            //Boolean openFlg = false;
            //for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            //{
            //    f = System.Windows.Forms.Application.OpenForms[i];
            //    if (f.Text.IndexOf("特命課長") >= 0 && f.Text.IndexOf("編集") <= -1)
            //    {
            //        f.Show();
            //        openFlg = true;
            //        break;
            //    }
            //}
            //if (!openFlg)
            //{
            //    Tokumei form = new Tokumei();
            //    form.UserInfos = this.UserInfos;
            //    form.Show();
            //    //this.Close();
            //}
            //this.Hide();
        }
        //ヘッダー窓口ボタン
        private void btnMadoguchi_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
            //Madoguchi form = new Madoguchi();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();
            Form f = null;
            Boolean openFlg = false;
            for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            {
                f = System.Windows.Forms.Application.OpenForms[i];
                if (f.Text.IndexOf("窓口ミハル") >= 0 && f.Text.IndexOf("編集") <= -1)
                {
                    f.Show();
                    openFlg = true;
                    break;
                }
            }
            if (!openFlg)
            {
                Madoguchi form = new Madoguchi();
                form.UserInfos = this.UserInfos;
                form.Show();
                //this.Close();
            }
            this.Hide();
        }
        //ヘッダー特調野郎TOPボタン
        private void btnTokuchoyaro_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
            //Tokuchoyaro form = new Tokuchoyaro();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();
            Form f = null;
            Boolean openFlg = false;
            for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            {
                f = System.Windows.Forms.Application.OpenForms[i];
                if (f.Text.IndexOf("特調野郎") >= 0 && f.Text.IndexOf("編集") <= -1)
                {
                    f.Show();
                    openFlg = true;
                    break;
                }
            }
            if (!openFlg)
            {
                Tokuchoyaro form = new Tokuchoyaro();
                form.UserInfos = this.UserInfos;
                form.Show();
                //this.Close();
            }
            this.Hide();
        }
        //ヘッダー自分大臣ボタン
        private void btnJibun_Click(object sender, EventArgs e)
        {
            this.ReSearch = true;
            //Jibun form = new Jibun();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();

            Form f = null;
            Boolean openFlg = false;
            for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            {
                f = System.Windows.Forms.Application.OpenForms[i];
                if (f.Text.IndexOf("自分大臣") >= 0 && f.Text.IndexOf("編集") <= -1)
                {
                    f.Show();
                    openFlg = true;
                    break;
                }
            }
            if (!openFlg)
            {
                Jibun form = new Jibun();
                form.UserInfos = this.UserInfos;
                form.Show();
                //this.Close();
            }
            this.Hide();
        }

        private void button_PrintKouteihyo_Click(object sender, EventArgs e)
        {
            errorCheck_initialize();

            if (item_tyouhyouInsatu.Text == "")
            {
                set_error("", 0);
                set_error("帳票を選択してください。");
            }
            else
            {
                string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    var Dt = new System.Data.DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT PrintDataPattern FROM Mst_PrintList"
                                    + " WHERE PrintListID = '" + item_tyouhyouInsatu.SelectedValue + "'"
                                    ;
                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(Dt);
                    Boolean errorFLG = false;
                    if (Dt.Rows.Count > 0)
                    {
                        set_error("", 0);
                        // 6:受託調査業務工程表
                        if (Dt.Rows[0][0].ToString() == "6")
                        {
                            // 締切日の入力チェック
                            errorFLG = errorCheck_Shimekiribi();

                            if (errorFLG == false)
                            {
                                // string[]
                                // 26個分先に用意
                                string[] report_data = new string[28] { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

                                // 0.呼び出し元画面（1:特命課長、0：自分大臣）
                                report_data[0] = "1";
                                // 1.部所名
                                report_data[1] = UserInfos[3];
                                // 2.職員名
                                report_data[2] = UserInfos[1];
                                // 3.登録年度
                                report_data[3] = item_Nendo.SelectedValue.ToString();
                                // 4.登録年度オプション
                                if (item_NendoOptionTounen.Checked)
                                {
                                    report_data[4] = "1";   // 当年度
                                }
                                else
                                {
                                    report_data[4] = "2";   // 3年以内
                                }
                                // 5.調査担当部所CD
                                if (item_ChousaBusho.Text != null && item_ChousaBusho.Text != "")
                                {
                                    report_data[5] = item_ChousaBusho.SelectedValue.ToString();
                                }
                                // 6.発注者名・課名
                                report_data[6] = item_HachushaKamei.Text;
                                // 7.特調番号
                                report_data[7] = item_TokuchoBangou.Text;
                                // 8.締切日from
                                report_data[8] = "null";
                                if (item_DateFrom.CustomFormat == "")
                                {
                                    report_data[8] = "'" + item_DateFrom.Text + "'";
                                }
                                // 9.締切日to
                                report_data[9] = "null";
                                if (item_DateTo.CustomFormat == "")
                                {
                                    report_data[9] = "'" + item_DateTo.Text + "'";
                                }
                                // 10.担当者状況
                                report_data[10] = "0";
                                if (item_TantoushaJoukyo.Text != null && item_TantoushaJoukyo.Text != "")
                                {
                                    report_data[10] = item_TantoushaJoukyo.SelectedValue.ToString();
                                }
                                // 11.窓口部所
                                if (item_MadoguchiBusho.Text != null && item_MadoguchiBusho.Text != "")
                                {
                                    report_data[11] = item_MadoguchiBusho.SelectedValue.ToString();
                                }
                                // 12.業務名称
                                report_data[12] = item_Gyoumumei.Text;
                                // 13.管理番号
                                report_data[13] = item_KanriBangou.Text;
                                // 14.検索期間の指定
                                report_data[14] = "0";
                                if (item_FromTo.Text != null && item_FromTo.Text != "")
                                {
                                    report_data[14] = item_FromTo.SelectedValue.ToString();
                                }
                                // 15.調査種別
                                report_data[15] = "0";
                                if (item_ChousaKind.Text != null && item_ChousaKind.Text != "")
                                {
                                    report_data[15] = item_ChousaKind.SelectedValue.ToString();
                                }
                                // 16.窓口担当者
                                report_data[16] = item_MadoguchiTantousha.Text;
                                // 17.工事件名
                                report_data[17] = item_Koujikenmei.Text;
                                // 18.締切日の選択
                                report_data[18] = "0";
                                if (item_ShimekiriSentaku.Text != null && item_ShimekiriSentaku.Text != "")
                                {
                                    report_data[18] = item_ShimekiriSentaku.SelectedValue.ToString();
                                }
                                // 19.調査区分（自部所）
                                report_data[19] = "0";
                                if (item_ChousaKbnJibusho.Checked)
                                {
                                    report_data[19] = "1";
                                }
                                // 20.調査区分（支→支）
                                report_data[20] = "0";
                                if (item_ChousaKbnShibuShibu.Checked)
                                {
                                    report_data[20] = "1";
                                }
                                // 21.調査区分（本→支）
                                report_data[21] = "0";
                                if (item_ChousaKbnHonbuShibu.Checked)
                                {
                                    report_data[21] = "1";
                                }
                                // 22.調査区分（支→本）
                                report_data[22] = "0";
                                if (item_ChousaKbnShibuHonbu.Checked)
                                {
                                    report_data[22] = "1";
                                }
                                // 23.調査品目
                                report_data[23] = item_ChousaHinmoku.Text;
                                // 24.進捗状況
                                report_data[24] = "0";
                                if (item_Shintyokujyoukyo.Text != null && item_Shintyokujyoukyo.Text != "")
                                {
                                    report_data[24] = item_Shintyokujyoukyo.SelectedValue.ToString();
                                }
                                // 25.本部単品       特命課長では未設定
                                report_data[25] = "0";
                                // 26.調査担当者名   特命課長では未設定
                                // 27.メモ          特命課長では未設定

                                string[] result = GlobalMethod.InsertMadoguchiReportWork(7, UserInfos[0], report_data,"KouteiHyo");

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
                                        Popup_Download form = new Popup_Download();
                                        form.TopLevel = false;
                                        this.Controls.Add(form);

                                        String fileName = Path.GetFileName(result[3]);
                                        form.ExcelName = fileName;
                                        form.TotalFilePath = result[2];
                                        form.Dock = DockStyle.Bottom;
                                        form.Show();
                                        form.BringToFront();
                                    }
                                }
                                else
                                {
                                    // エラーが発生しました
                                    set_error(GlobalMethod.GetMessage("E00091", ""));
                                }
                            }
                        }
                        //if (errorFLG == true)
                        //{
                        //    set_error(errorMsg);
                        //}
                    }
                    conn.Close();
                }
            }
        }

        private void errorCheck_initialize()
        {
            // エラーメッセージのクリア
            set_error("", 0);

            // 画面背景色の初期化
            //label23.BackColor = Color.CadetBlue;
            label_DateFrom.BackColor = Color.CadetBlue;
            label_DateTo.BackColor = Color.CadetBlue;
        }

        // 締切日のチェック（true:エラー、false:正常）
        private Boolean errorCheck_Shimekiribi()
        {

            // 締切日の未入力チェック
            if (item_DateFrom.CustomFormat == " " && item_DateTo.CustomFormat == " ")
            {
                // E40001 締切日を入力してください
                set_error(GlobalMethod.GetMessage("E40001", ""));
                //label23.BackColor = errorBackColor;
                label_DateFrom.BackColor = errorBackColor;

                return true;
            }

            // 締切日のFromとToの大小関係のチェック
            if (item_DateFrom.CustomFormat == "" && item_DateTo.CustomFormat == "")
            {
                if (item_DateFrom.Value > item_DateTo.Value)
                {
                    // E20002 対象項目の入力に誤りがあります。
                    set_error(GlobalMethod.GetMessage("E20002", ""));
                    //label23.BackColor = errorBackColor;
                    label_DateFrom.BackColor = errorBackColor;
                    label_DateTo.BackColor = errorBackColor;

                    return true;
                }
            }

            return false;
        }
        private void changeShimekiribi()
        {
            switch (item_FromTo.SelectedValue.ToString())
            {
                case "1":   // 1:以前
                    // 締切日のfromが設定されていた場合、fromをtoにセットして、fromは空にする
                    if (item_DateFrom.CustomFormat == "")
                    {
                        item_DateTo.Value = item_DateFrom.Value;
                        item_DateTo.CustomFormat = "";
                        item_DateFrom.Text = "";
                        item_DateFrom.CustomFormat = " ";
                    }
                    break;
                case "2":   // 2:当日の場合
                    // 締切日のfromが設定されていた場合、fromをtoにセット
                    // 締切日のtoだけ設定されていた場合、toをfromにセット
                    if (item_DateFrom.CustomFormat == "")
                    {
                        item_DateTo.Value = item_DateFrom.Value;
                        item_DateTo.CustomFormat = "";
                    }
                    else
                    {
                        if (item_DateTo.CustomFormat == "")
                        {
                            item_DateFrom.Value = item_DateTo.Value;
                            item_DateFrom.CustomFormat = "";
                        }
                    }
                    break;
                case "3":   // 3:一週間
                    // 締切日のfromが設定されていた場合、from+6をtoにセット
                    // 締切日のtoだけ設定されていた場合、to-6をfromにセット
                    if (item_DateFrom.CustomFormat == "")
                    {
                        item_DateTo.Value = item_DateFrom.Value.AddDays(6);
                        item_DateTo.CustomFormat = "";
                    }
                    else
                    {
                        if (item_DateTo.CustomFormat == "")
                        {
                            item_DateFrom.Value = item_DateTo.Value.AddDays(-6);
                            item_DateFrom.CustomFormat = "";
                        }
                    }
                    break;
                default:
                    break;
            }
            // 締め日の選択を空にする
            item_ShimekiriSentaku.SelectedValue = -1;   // item_ShimekiriSentaku_TextChanged が動く

            // 締切日が未入力だった場合、背景色の変更
            if (item_DateFrom.CustomFormat == " " && item_DateTo.CustomFormat == " ")
            {
                label_DateFrom.BackColor = errorBackColor;
            }
        }

        private void c1FlexGrid1_AfterSort(object sender, C1.Win.C1FlexGrid.SortColEventArgs e)
        {
            if (c1FlexGrid1.Rows.Count > 2)
            {
                c1FlexGrid1.Select(2, 2, true);
            }
        }

        private void btnGridSize_Click(object sender, EventArgs e)
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:667 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 667;
            //    c1FlexGrid1.Width = 1864;
            //}
            gridSizeChange();
        }
        private void gridSizeChange()
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:667 → 1086・・・調査品目明細と合わせる
            //    // width:1864 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid1.Height = 1086;
            //    c1FlexGrid1.Width = 3752;
            //}
            //else
            //{
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid1.Height = 667;
            //    c1FlexGrid1.Width = 1864;
            //}
            string num = "";
            int bigHeight = 0;
            int bigWidth = 0;
            int smallHeight = 0;
            int smallWidth = 0;

            if (btnGridSize.Text == "一覧拡大")
            {
                num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_GRID_BIG_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out bigHeight);
                    if (bigHeight == 0)
                    {
                        bigHeight = 1086;
                    }
                }
                num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_GRID_BIG_WIDTH");
                if (num != null)
                {
                    Int32.TryParse(num, out bigWidth);
                    if (bigWidth == 0)
                    {
                        bigWidth = 3752;
                    }
                }

                // height:628 → 1086・・・調査品目明細と合わせる
                // width:1864 → 3752
                btnGridSize.Text = "一覧縮小";
                c1FlexGrid1.Height = bigHeight;
                c1FlexGrid1.Width = bigWidth;

            }
            else
            {
                num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_GRID_SMALL_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out smallHeight);
                    if (smallHeight == 0)
                    {
                        smallHeight = 667;
                    }
                }
                num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_GRID_SMALL_WIDTH");
                if (num != null)
                {
                    Int32.TryParse(num, out smallWidth);
                    if (smallWidth == 0)
                    {
                        smallWidth = 1864;
                    }
                }

                btnGridSize.Text = "一覧拡大";
                c1FlexGrid1.Height = smallHeight;
                c1FlexGrid1.Width = smallWidth;
            }
        }
    }
}
