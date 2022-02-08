using System;
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
using System.Collections;

namespace TokuchoBugyoK2
{
    public partial class Popup_Anken : Form
    {
        public string[] ReturnValue = new string[17];
        private DataTable ListData = new DataTable();
        GlobalMethod GlobalMethod = new GlobalMethod();
        public string mode = "";
        public string keikakubangou = "";
        public string nendo = "";
        public string jutakuBangou = "";
        public string gyoumuBushoCD = "";
        public string hachuushaKaMei = "";
        public string gyoumuMei = "";

        public Popup_Anken()
        {
            InitializeComponent();

            //マウスホイールの制御を追加
            this.src_1.MouseWheel += item_MouseWheel;   //売上年度
            this.src_4.MouseWheel += item_MouseWheel;   //受託部所
        }

        private void Popup_Anken_Load(object sender, EventArgs e)
        {

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            get_data();
            if (mode == "kurikoshi")
            {
                groupBox1.Visible = false;
                groupBox2.Text = "繰越案件一覧";
                c1FlexGrid1.Cols.Remove(1);
                c1FlexGrid1.Cols.Remove(9);
                c1FlexGrid1.Cols[9].Visible = true;
                c1FlexGrid1.Cols.Move(2, 0);
                this.Text = "繰越案件一覧";
            }
            if (mode == "keikaku")
            {
                groupBox1.Visible = false;
                groupBox2.Text = "案件一覧";
                //c1FlexGrid1.Cols.Remove(0); // 選択
                //c1FlexGrid1.Cols.Remove(8); // 入札状況 （選択が消えたことで9 ⇒ 8）
                //c1FlexGrid1.Cols[8].Visible = true; // 落札額・・・？元から表示されている
                //c1FlexGrid1.Cols.Move(2, 0); // 受託番号を先頭に移動

                // 選択列非表示
                c1FlexGrid1.Cols[1].Visible = false;
                // 入札状況を非表示
                c1FlexGrid1.Cols[10].Visible = false;

                // タイトルと幅を変える
                c1FlexGrid1.Cols[3].Caption = "受託番号";
                c1FlexGrid1.Cols[3].Width = 100;
                c1FlexGrid1.Cols[4].Caption = "案件番号";

                this.Text = "案件一覧";
            }
            if (mode == "kakoirai")
            {
                groupBox1.Visible = false;
                groupBox3.Visible = true;
                groupBox2.Text = "受託番号一覧";
                c1FlexGrid1.Width = 160;
                c1FlexGrid1.Cols[2].Width = 130;
                c1FlexGrid1.Cols[3][0] = "特調番号";
                c1FlexGrid1.Cols[3].Width = 180;
                c1FlexGrid1.Cols[4].Visible = false;
                c1FlexGrid1.Cols[5].Visible = false;
                c1FlexGrid1.Cols[6].Visible = false;
                c1FlexGrid1.Cols[7].Visible = false;
                c1FlexGrid1.Cols[8].Visible = false;
                c1FlexGrid1.Cols[9].Visible = false;
                c1FlexGrid1.Cols[10].Visible = false;
                c1FlexGrid1.Cols[11].Visible = false;
                c1FlexGrid1.Cols[12].Visible = false;
                this.Text = "選択リスト";
                this.Width = 260;
                if(jutakuBangou != "")
                {
                    if(jutakuBangou.Length > 9) { 
                        textBox5.Text = jutakuBangou.Substring(0,9);
                    }
                    else
                    {
                        textBox5.Text = jutakuBangou;
                    }
                }
            }
            set_combo();

            if (nendo != "")
            {
                src_1.SelectedValue = nendo;
            }

            if (hachuushaKaMei != "")
            {
                src_2.Text = hachuushaKaMei;
            }

            if (gyoumuMei != "")
            {
                src_3.Text = gyoumuMei;
            }

            if (gyoumuBushoCD != "")
            {
                src_4.SelectedValue = gyoumuBushoCD;
            }
        }

        private void set_combo()
        {
            GlobalMethod GlobalMethod = new GlobalMethod();
            //売上年度
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";

            // 売上年度は年度マスタのデータを全て表示
            //String where = "NendoID <= YEAR(GETDATE()) AND NendoID >= YEAR(GETDATE()) - 3";
            String where = "";
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            src_1.DataSource = combodt;
            src_1.DisplayMember = "Discript";
            src_1.ValueMember = "Value";

            //入札状況
            //SQL変数
            discript = "RakusatsuShaMei";
            value = "RakusatsuShaID";
            table = "Mst_RakusatsuSha";
            where = "RakusatsuShaNarabijun > 0";
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            SortedList sl = GlobalMethod.Get_SortedList(combodt);
            //該当グリッドのセルにセット
            //c1FlexGrid1.Cols[9].DataMap = sl;
            c1FlexGrid1.Cols[10].DataMap = sl;

            //受託課所支部
            //SQL変数
            discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            value = "Mst_Busho.GyoumuBushoCD ";
            table = "Mst_Busho";
            where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND BushoEntryHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
                    //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' ";
                    "AND NOT GyoumuBushoCD LIKE '121%' ";
            int FromNendo;
            if (int.TryParse(nendo, out FromNendo))
            {
                int ToNendo = int.Parse(nendo) + 1;
                where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";
            }
            //コンボボックスデータ取得
            combodt = GlobalMethod.getData(discript, value, table, where);
            src_4.DataSource = combodt;
            src_4.DisplayMember = "Discript";
            src_4.ValueMember = "Value";

        }

        private void set_combo_shibu(string nendo)
        {
            //受託課所支部
            string SelectedValue = "";
            if (src_4.Text != "")
            {
                SelectedValue = src_4.SelectedValue.ToString();
            }
            //SQL変数
            string discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            string where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND BushoEntryHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
                    "AND NOT GyoumuBushoCD LIKE '121%' ";
            int FromNendo;
            if (int.TryParse(nendo, out FromNendo))
            {
                int ToNendo = int.Parse(nendo) + 1;
                //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' )";
                where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' )";
            }
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            DataRow dr = combodt.NewRow();
            combodt.Rows.InsertAt(dr, 0);
            src_4.DataSource = combodt;
            src_4.DisplayMember = "Discript";
            src_4.ValueMember = "Value";

            if (SelectedValue != "")
            {
                src_4.SelectedValue = SelectedValue;
            }
        }

        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            var dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                if (mode != "kakoirai")
                {
                    string toukai = GlobalMethod.GetCommonValue2("ENTORY_TOUKAI");

                    if(mode != "keikaku") { 
                         cmd.CommandText = "SELECT " +
                        "AnkenJouhou.AnkenJouhouID " + // 0:契約情報ID
                        ",AnkenAnkenBangou " +
                        ",AnkenJutakuBangou ";
                    }
                    else
                    {
                        // 計画詳細の場合
                        cmd.CommandText = "SELECT " +
                            "AnkenJouhou.AnkenJouhouID " + // 0:契約情報ID
                            ",AnkenJutakuBangou " +
                            ",AnkenAnkenBangou ";
                    }

                    cmd.CommandText += ",AnkenJutakuBangouEda " +
                        ",AnkenUriageNendo " +
                        ",AnkenHachuushaKaMei " + // 5:発注者名課名
                        ",AnkenJutakushibu " +
                        ",AnkenGyoumuMei " +
                        ",NyuusatsuRakusatsushaID " +
                        ",NyuusatsuRakusatsusha " +
                        ",ISNULL(NyuusatsuRakusatugaku,0) " + // 10:落札額
                        //",ISNULL(KeiyakuZeikomiKingaku,0) " + // 11:前回契約金額（税抜）
                        ",ISNULL(KeiyakuKeiyakuKingaku,0) " + // 11:前回契約金額（税抜）
                        //",ISNULL(NyuusatsuOusatugaku,0) " + // 12:前回応札額（税抜）
                        ",ISNULL(NyuusatsuOusatsuKingaku,0) " + 
                        ",ISNULL(NyuusatsuMitsumorigaku,0) " +
                        //",ISNULL(Keiyakukeiyakukingakukei,0) " + // 14
                        // 受託金額（税込） - 消費税額
                        //",ISNULL(FLOOR(Keiyakukeiyakukingakukei / (1 + (KeiyakuShouhizeiritsu / 100.00))),0) " + // 14:前回受託金額（税抜）
                        // 495 落札者が「建設物価調査会」ではない場合、前回受託金額（税抜）を0にするよう修正
                        ",CASE NyuusatsuRakusatsusha " +
                        "  WHEN '" + toukai + "' THEN ISNULL(FLOOR(Keiyakukeiyakukingakukei / (1 + (KeiyakuShouhizeiritsu / 100.00))),0) " +
                        "  ELSE 0 " +
                        "END " + // 14:前回受託金額（税抜）

                        //",NyuusatsuKyougouTashaID " +
                        ",KyougouTashaID " +
                        ",KyougouKigyouCD " +
                        //",ISNULL(TankaKeiyakuRank.TankaKeiyakuID,0) " +
                        // 1224 重複する受託番号を除外する対応
                        //",ISNULL(TankaKeiyaku.TankaKeiyakuID,0) " + 
                        ",ISNULL(TK.TankaKeiyakuID,0) " +
                        "FROM AnkenJouhou LEFT JOIN Mst_Busho ON AnkenJutakubushoCD = GyoumuBushoCD " +
                        "INNER JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                        "LEFT JOIN KeiyakuJouhouEntory ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID " +
                        "LEFT JOIN Mst_RakusatsuSha ON NyuusatsuRakusatsushaID = RakusatsuShaID " +
                        //"LEFT JOIN Mst_KyougouTasha ON KyougouTashaID = NyuusatsuKyougouTashaID " +
                        //"LEFT JOIN Mst_KyougouTasha ON KyougouMeishou = NyuusatsuRakusatsusha  " +
                        "LEFT JOIN (SELECT KyougouMeishou, KyougouKigyouCD, Max(KyougouTashaID) AS KyougouTashaID FROM Mst_KyougouTasha GROUP BY KyougouMeishou, KyougouKigyouCD ) KT " +
                                "ON KT.KyougouMeishou = NyuusatsuRakusatsusha " +
                        //"LEFT JOIN TankaKeiyaku ON TankaKeiyaku.AnkenJouhouID = AnkenJouhou.AnkenJouhouID " +
                        //"LEFT JOIN TankaKeiyaku ON TankaKeiyaku.TankakeiyakuJutakuBangou = AnkenJouhou.AnkenJutakuBangou " +
                        // 1224 重複する受託番号を除外する対応
                        //"LEFT JOIN TankaKeiyaku ON TankaKeiyaku.TankakeiyakuJutakuBangou = AnkenJouhou.AnkenJutakuBangou AND AnkenJouhou.AnkenJutakuBangou <> '' " +
                        "LEFT JOIN (SELECT TankaKeiyaku.TankakeiyakuJutakuBangou, Max(TankaKeiyaku.TankaKeiyakuID) as TankaKeiyakuID FROM TankaKeiyaku GROUP BY TankaKeiyaku.TankakeiyakuJutakuBangou) TK " +
                                "ON TK.TankakeiyakuJutakuBangou = AnkenJouhou.AnkenJutakuBangou AND AnkenJouhou.AnkenJutakuBangou <> '' " +
                        //"LEFT JOIN TankaKeiyakuRank ON TankaKeiyaku.TankaKeiyakuID = TankaKeiyakuRank.TankaKeiyakuID AND TankaRankID = 1 " +
                        "LEFT JOIN (SELECT TankaKeiyakuRank.TankaKeiyakuID, Min(TankaKeiyakuRank.TankaRankID) AS TankaRankID FROM TankaKeiyakuRank GROUP BY TankaKeiyakuRank.TankaKeiyakuID ) TR " +
                                // 1224 重複する受託番号を除外する対応
                                //"ON TankaKeiyaku.TankaKeiyakuID = TR.TankaKeiyakuID " +
                                "ON TK.TankaKeiyakuID = TR.TankaKeiyakuID " +
                        "LEFT JOIN (SELECT Mst_Koujijimusho.TankaKeiyakuID , COUNT(Mst_Koujijimusho.TankaKeiyakuID) AS 'CNT' FROM Mst_Koujijimusho GROUP BY Mst_Koujijimusho.TankaKeiyakuID ) TMP " +
                                // 1224 重複する受託番号を除外する対応
                                //"ON TankaKeiyaku.TankaKeiyakuID = TMP.TankaKeiyakuID " +
                                "ON TK.TankaKeiyakuID = TMP.TankaKeiyakuID " +
                        "LEFT JOIN (select NyuusatsuJouhouID,min(NyuusatsuOusatsuKingaku) AS NyuusatsuOusatsuKingaku FROM NyuusatsuJouhouOusatsusha where NyuusatsuOusatsusha = '" + toukai + "' group by NyuusatsuJouhouID) T1 " +
                        "  ON T1.NyuusatsuJouhouID = AnkenJouhou.AnkenJouhouID " +
                        //"WHERE AnkenSaishinFlg = 1 AND AnkenDeleteFlag = 0 AND AnkenSakuseiKubun <> '02' AND NOT AnkenJutakuBangou LIKE '%999' ";
                        // 業務日報のデータは60000～70000の間で登録している為、除外
                        "WHERE AnkenSaishinFlg = 1 AND AnkenDeleteFlag = 0 AND AnkenSakuseiKubun <> '02' AND (AnkenJouhou.AnkenJouhouID < 60000 or AnkenJouhou.AnkenJouhouID > 70000) ";

                    if (src_1.Text != "" && mode != "keikaku")
                    {
                        cmd.CommandText += " AND AnkenUriageNendo COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(src_1.SelectedValue.ToString(), 0) + "'  ";
                    }
                    if (src_2.Text != "")
                    {
                        cmd.CommandText += " AND AnkenHachuushaKaMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_2.Text, 1) + "%' ESCAPE'\\' ";
                    }
                    if (src_3.Text != "")
                    {
                        cmd.CommandText += " AND AnkenGyoumuMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_3.Text, 1) + "%' ESCAPE'\\' ";
                    }
                    //受託部所が入っているとき
                    if (src_4.SelectedValue != null && mode != "keikaku")
                    {
                        String jutakubusho = src_4.SelectedValue.ToString();

                        ////受託部所が127900
                        //if ("127900".Equals(jutakubusho))
                        //{
                        //    cmd.CommandText += "AND (AnkenJutakubushoCD LIKE '1279%' " +
                        //        "OR AnkenJutakubushoCD = '128416') ";
                        //}
                        ////受託部所が127100
                        //else if ("127100".Equals(jutakubusho))
                        //{
                        //    cmd.CommandText += "AND (AnkenJutakubushoCD LIKE '1271%' " +
                        //       "OR AnkenJutakubushoCD = '127220') ";
                        //}
                        ////受託部所が127900、127100でない場合
                        //else
                        //{
                        //    cmd.CommandText += "AND AnkenJutakubushoCD LIKE '" + jutakubusho.TrimEnd('0') + "%" + "' ";
                        //}

                        if (jutakubusho != "")
                        {
                            if ("127100".Equals(jutakubusho))
                            {
                                cmd.CommandText += "  and AnkenJutakubushoCD like '1271%' ";
                            }
                            else
                            {
                                // 128400 情報システム部関連で、2021年度よりも前の場合
                                if (jutakubusho.Substring(0, 4) == "1284" && src_1.Text != null && src_1.Text != "" && int.Parse(src_1.SelectedValue.ToString()) < 2021)
                                {
                                    // 127900 情報システム部【旧】を見る
                                    cmd.CommandText += "  and AnkenJutakubushoCD LIKE '127900'";
                                }
                                else if (jutakubusho.Substring(0, 4) == "1292")
                                {
                                    cmd.CommandText += "  and AnkenJutakubushoCD LIKE '129230'";
                                }
                                else
                                {
                                    cmd.CommandText += "  and AnkenJutakubushoCD LIKE '" + jutakubusho.TrimEnd('0') + "%'";
                                }

                                if ("127000".Equals(jutakubusho))
                                {
                                    // 2021年度より前は、1279は除外する
                                    //if (int.Parse(src_1.SelectedValue.ToString()) < 2021) {
                                    if (src_1.Text != null && src_1.Text != "" && int.Parse(src_1.SelectedValue.ToString()) < 2021)
                                    {
                                        // 127000 本部 調査部門の場合、 1279 情報システム部関連は除外
                                        cmd.CommandText += "  and NOT AnkenJutakubushoCD LIKE '1279%'";
                                    }
                                }
                            }

                        }
                    }
                    if (src_5.Text != "")
                    {
                        cmd.CommandText += " AND AnkenAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_5.Text, 1) + "%' ESCAPE'\\' ";
                    }
                    if (src_6.Text != "")
                    {
                        cmd.CommandText += " AND AnkenJutakuBangou + '-' + AnkenJutakuBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_6.Text, 1) + "%' ESCAPE'\\' ";
                    }
                    if (mode == "keikaku")
                    {
                        cmd.CommandText += " AND AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(keikakubangou, 0) + "'  ";
                    }
                    if (mode == "tanka")
                    {
                        //cmd.CommandText += " AND AnkenKianZumi = 1 ";
                        cmd.CommandText += " AND AnkenJutakuBangou <> '' ";
                    }
                    if (mode == "kakotanka")
                    {
                        //cmd.CommandText += " AND (TankaRankID IS NOT NULL OR TMP.CNT > 0 ) ";
                        cmd.CommandText += " AND (TR.TankaRankID IS NOT NULL OR TMP.CNT > 0 ) ";
                    }

                    // 499 業務日報用の案件を除外する
                    //cmd.CommandText += " AND AnkenAnkenBangou not like '%999' ";
                    // 業務日報のデータは60000～70000の間で登録している為、除外
                    cmd.CommandText += " AND (AnkenJouhou.AnkenJouhouID < 60000 or AnkenJouhou.AnkenJouhouID > 70000) ";

                    cmd.CommandText += " ORDER BY AnkenJutakuBangou";
                }
                else
                {
                    cmd.CommandText = "SELECT " +
                        "MadoguchiID " +
                        ",MadoguchiUketsukeBangou + '-' + MadoguchiUketsukeBangouEdaban " +
                        "FROM MadoguchiJouhou " +
                        "WHERE MadoguchiDeleteFlag <> 1 ";

                    if (textBox5.Text != "")
                    {
                        cmd.CommandText += " AND SUBSTRING(MadoguchiJutakuBangou,1,9) COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(textBox5.Text, 1) + "%'  ";
                    }
                    cmd.CommandText += " ORDER BY MadoguchiUketsukeBangou, MadoguchiUketsukeBangouEdaban ";
                }
                Console.WriteLine(cmd.CommandText);
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
            Paging_now.Text = (1).ToString();
            set_data(1);
        }

        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid1.HitTest(new System.Drawing.Point(e.X, e.Y));

            //if (mode != "kurikoshi" && mode != "kakoirai" && mode != "keikaku" && hti.Column == 0 && hti.Row != 0)
            if (mode != "kurikoshi" && mode != "kakoirai" && mode != "keikaku" && hti.Column == 1 && hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                ReturnValue[0] = c1FlexGrid1.Rows[_row][_col + 1].ToString();     // AnkenJouhou.AnkenJouhouID
                ReturnValue[1] = c1FlexGrid1.Rows[_row][_col + 2].ToString();     // AnkenAnkenBangou
                ReturnValue[2] = c1FlexGrid1.Rows[_row][_col + 3].ToString();     // AnkenJutakuBangouALL
                ReturnValue[3] = c1FlexGrid1.Rows[_row][_col + 4].ToString();     // AnkenJutakuBangouEda
                ReturnValue[4] = c1FlexGrid1.Rows[_row][_col + 8].ToString();     // AnkenGyoumuMei
                ReturnValue[5] = c1FlexGrid1.Rows[_row][_col + 10].ToString();    // NyuusatsuRakusatsusha
                ReturnValue[6] = c1FlexGrid1.Rows[_row][_col + 9].ToString();     // NyuusatsuRakusatsushaID
                ReturnValue[7] = c1FlexGrid1.Rows[_row][_col + 11].ToString();    // NyuusatsuRakusatugaku
                ReturnValue[8] = c1FlexGrid1.Rows[_row][_col + 13].ToString();    // NyuusatsuOusatugaku
                ReturnValue[9] = c1FlexGrid1.Rows[_row][_col + 14].ToString();    // NyuusatsuMitsumorigaku
                ReturnValue[10] = c1FlexGrid1.Rows[_row][_col + 12].ToString();   // KeiyakuZeikomiKingaku
                ReturnValue[11] = c1FlexGrid1.Rows[_row][_col + 15].ToString();   // Keiyakukeiyakukingakukei　
                ReturnValue[12] = c1FlexGrid1.Rows[_row][_col + 16].ToString();   // NyuusatsuKyougouTashaID
                ReturnValue[13] = c1FlexGrid1.Rows[_row][_col + 17].ToString();   // KyougouKigyouCD
                ReturnValue[14] = c1FlexGrid1.Rows[_row][_col + 6].ToString();    // AnkenHachuushaKaMei
                ReturnValue[15] = c1FlexGrid1.Rows[_row][_col + 7].ToString();    // AnkenJutakushibu
                if (mode == "tanka" || mode == "kakotanka")
                {
                    ReturnValue[16] = c1FlexGrid1.Rows[_row][_col + 18].ToString();    // TankaKeiyakuRank.TankaKeiyakuID
                }
                this.Close();
            }

            if (mode == "kakoirai" && hti.Column == 1 && hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                ReturnValue[0] = c1FlexGrid1.Rows[_row][_col + 1].ToString();     // MadoguchiID
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
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
                for (int i = 0; i < c1FlexGrid1.Cols.Count - 2; i++)
                {
                    if (ListData.Columns.Count > i)
                    {
                        c1FlexGrid1[r + 1, i + 2] = ListData.Rows[startrow + r][i];
                    }
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


        private void src_1_TextChanged(object sender, EventArgs e)
        {
            set_combo_shibu(src_1.SelectedValue.ToString());
            get_data();
        }

        private void src_4_TextChanged(object sender, EventArgs e)
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

        // 
        private void src2_TextChanged(object sender, EventArgs e)
        {
            get_data();
        }

        // マウスホイールイベントでコンボ値が変わらないようにする
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void src_1_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
        }

        private void src_4_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
        }
    }
}





