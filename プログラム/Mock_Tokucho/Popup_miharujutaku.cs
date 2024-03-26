using System;
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

namespace TokuchoBugyoK2
{
    public partial class Popup_miharujutaku : Form
    {
        public string jutakuno;
        public string gyoumumei;
        public string jimusyo;
        public string busyo;
        public string tantousya;

        GlobalMethod GlobalMethod = new GlobalMethod();
        private DataTable ListData = new DataTable();
        private DataTable JimushoData = new DataTable();
        public string Busho;
        public string Nendo;
        public string[] ReturnValue = new string[30];
        private int FromNendo;
        private int ToNendo;

        public Popup_miharujutaku()
        {
            InitializeComponent();
        }

        private void Popup_miharujutaku_Load(object sender, EventArgs e)
        {
            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid2.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid2.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            // ホイール制御
            this.comboBox1.MouseWheel += item_MouseWheel; // 年度
            this.comboBox2.MouseWheel += item_MouseWheel; // 起案状態
            this.cmb_JutakuBusho.MouseWheel += item_MouseWheel; // 受託部所
            
            set_combo(Nendo);
            get_data();

            if (Busho != null && Busho != "")
            {
                cmb_JutakuBusho.SelectedValue = Busho;
            }
        }
        private void set_nendo()
        {

            if (int.TryParse(Nendo, out FromNendo))
            {
                ToNendo = FromNendo + 1;
            }
            else
            {
                FromNendo = DateTime.Today.Year;
                ToNendo = FromNendo + 1;
            }
        }

        private void set_combo(String Nendo)
        {

            //年度
            string discript = "NendoSeireki";
            string value = "NendoID";
            string table = "Mst_Nendo";
            string where = "";
            //コンボボックスデータ取得
            DataTable tmpdt4 = GlobalMethod.getData(discript, value, table, where);
            comboBox1.DisplayMember = "Discript";
            comboBox1.ValueMember = "Value";
            comboBox1.DataSource = tmpdt4;

            //年度の値を持っていたら受け渡す
            if (Nendo != null)
            {
                comboBox1.SelectedValue = Nendo;
            }
            else
            {
                Nendo = comboBox1.SelectedValue.ToString();
            }

            set_nendo();

            //起案
            DataTable tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "指定なし");
            tmpdt.Rows.Add(2, "済");
            tmpdt.Rows.Add(3, "未");
            comboBox2.DataSource = tmpdt;
            comboBox2.DisplayMember = "Discript";
            comboBox2.ValueMember = "Value";

            comboBox2.SelectedValue = "1";

            //受託課所支部
            set_combo_shibu();
        }

        /// <summary>
        /// 受託課所支部
        /// </summary>
        private void set_combo_shibu()
		{

            string discript = "";
            string value = "";
            string table = "";
            string where = "";

            string SelectedValue = "";
            //選択した値を退避
            if (cmb_JutakuBusho.Text != "")
            {
                SelectedValue = cmb_JutakuBusho.SelectedValue.ToString();
            }

            //SQL変数
            discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            value = "Mst_Busho.GyoumuBushoCD ";
            table = "Mst_Busho";
            where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
                    //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' ";
                    "AND NOT GyoumuBushoCD LIKE '121%' ";
            int FromNendo;
            if (int.TryParse(comboBox1.SelectedValue.ToString(), out FromNendo))
            {
                int ToNendo = int.Parse(Nendo) + 1;

                //No1702
                //From年度は翌年の3/31以前の開始日
                //To年月は当年の4/1以降の終了日
                FromNendo += 1;
                ToNendo -= 1;
                where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/4/1' )";
            }

            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            cmb_JutakuBusho.DataSource = combodt;
            cmb_JutakuBusho.DisplayMember = "Discript";
            cmb_JutakuBusho.ValueMember = "Value";


            if (SelectedValue != "")
            {
                cmb_JutakuBusho.SelectedValue = SelectedValue;
            }
        }

        private void get_data()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT " +
                    "aj.AnkenKianZumi " +
                    ",aj.AnkenAnkenBangou " +
                    ",aj.AnkenJutakuBangou " +
                    ",aj.AnkenJutakuBangouEda " +//非表示
                    ",aj.AnkenJutakushibu " +
                    ",aj.AnkenHachuushaKaMei " +
                    ",aj.AnkenGyoumuMei " +
                    ",kje.KeiyakuZeikomiKingaku " +
                    ",aj.AnkenTantoushaMei " +
                    ",aj.AnkenTantoushaCD " +//非表示
                    ",kje.KeiyakuUriageHaibunCho " +
                    ",aj.GyoumuKanrishaMei  " +　//非表示　業務管理者
                    ",aj.GyoumuKanrishaCD  " +//非表示　業務管理者CD
                    ",aj.AnkenJouhouID " +//非表示
                    ",gj.KanriGijutsushaNM " +//非表示
                    ",gj.KanriGijutsushaCD " +//非表示
                    ",aj.AnkenGyoumuKubun " +//非表示
                    ",aj.AnkenJutakubushoCD " +//非表示　
                    ",mb.BushoShozokuChou " +//受託所属長
                    ",aj.AnkenKeiyakusho " + // 非表示 案件（受託）フォルダ
                    ",mb.JigyoubuHeadCD " + // 非表示 事業部HeadCD
                    ",gjm.GyoumuJouhouMadoKojinCD " + // 非表示　窓口担当者CD
                    ",gjm.GyoumuJouhouMadoChousainMei " + // 非表示　窓口担当者名
                    //",gjm.GyoumuJouhouMadoGyoumuBushoCD " +// 非表示　窓口部所CD
                    //",gjm.GyoumuJouhouMadoShibuMei " +// 非表示　窓口部所名
                    ",mc.GyoumuBushoCD " +
                    ",mb2.ShibuMei " +

                    "FROM AnkenJouhou aj " +
                    "LEFT JOIN KeiyakuJouhouEntory kje ON " +
                    "aj.AnkenJouhouID = kje.AnkenJouhouID " +
                    "LEFT JOIN Mst_Busho mb ON " +
                    "aj.AnkenJutakubushoCD = mb.GyoumuBushoCD " +
                    "LEFT JOIN GyoumuJouhou gj ON " +
                    "aj.AnkenJouhouID = gj.AnkenJouhouID " +
                    "LEFT JOIN GyoumuJouhouMadoguchi gjm ON " +
                    "gj.GyoumuJouhouID = gjm.GyoumuJouhouID " +

                    // 1225 窓口部所は、窓口担当者のコードから調査員マスタを引いて、業務部所コードを取得して設定
                    "LEFT JOIN Mst_Chousain mc ON " +
                    "gjm.GyoumuJouhouMadoKojinCD = mc.KojinCD " +
                    "LEFT JOIN Mst_Busho mb2 ON " +
                    "mc.GyoumuBushoCD = mb2.GyoumuBushoCD " +

                    "WHERE AnkenUriageNendo = '" + Nendo + "' " +
                    //"AND aj.AnkenJutakuBangou NOT LIKE '%9999' " +
                    "AND (aj.AnkenJouhouID < 60000 or aj.AnkenJouhouID > 70000) " +
                    "AND aj.AnkenJouhouID > 0 " +
                    "AND aj.AnkenSaishinFlg = 1 " +
                    "AND aj.AnkenDeleteFlag != 1 ";

                //受託部所が入っているとき
                if (cmb_JutakuBusho.SelectedValue != null)
                {
                    String jutakubusho = cmb_JutakuBusho.SelectedValue.ToString();

                    ////受託部所が127900
                    //if ("127900".Equals(jutakubusho))
                    //{
                    //    cmd.CommandText += "AND (aj.AnkenJutakubushoCD LIKE '1279%' " +
                    //        "OR aj.AnkenJutakubushoCD = '128416') ";
                    //}
                    ////受託部所が127100
                    //else if ("127100".Equals(jutakubusho))
                    //{
                    //    cmd.CommandText += "AND (aj.AnkenJutakubushoCD LIKE '1271%' " +
                    //       "OR aj.AnkenJutakubushoCD = '127220') ";
                    //}
                    ////受託部所が127900、127100でない場合
                    //else
                    //{
                    //    cmd.CommandText += "AND aj.AnkenJutakubushoCD LIKE '" + jutakubusho.TrimEnd('0') + "%" + "' ";
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
                            if (jutakubusho.Substring(0, 4) == "1284" && comboBox1.Text != null && comboBox1.Text != "" && int.Parse(comboBox1.SelectedValue.ToString()) < 2021)
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
                                if (comboBox1.Text != null && comboBox1.Text != "" && int.Parse(comboBox1.SelectedValue.ToString()) < 2021)
                                {
                                    // 127000 本部 調査部門の場合、 1279 情報システム部関連は除外
                                    cmd.CommandText += "  and NOT AnkenJutakubushoCD LIKE '1279%'";
                                }
                            }
                        }
                    }
                }

                //起案　済のとき
                if (comboBox2.SelectedValue != null)
                {
                    if ("2".Equals(comboBox2.SelectedValue.ToString()))
                    {
                        cmd.CommandText += "AND aj.AnkenKianZumi = 1 ";
                    }
                    //起案 未のとき
                    else if ("3".Equals(comboBox2.SelectedValue.ToString()))
                    {
                        cmd.CommandText += "AND (aj.AnkenKianZumi != 1 OR aj.AnkenKianZumi is null) ";
                    }
                }
                //案件番号
                if (textBox2.Text != "")
                {
                   cmd.CommandText += "AND aj.AnkenAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(textBox2.Text, 1) + "%' ESCAPE '\\' ";
                }

                //受託番号
                if (textBox3.Text != "")
                {
                    cmd.CommandText += "AND aj.AnkenJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(textBox3.Text, 1) + "%' ESCAPE '\\' ";
                }

                //発注者課名
                if (textBox5.Text != "")
                {
                    cmd.CommandText += "AND aj.AnkenHachuushaKaMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(textBox5.Text, 1) + "%' ESCAPE '\\' ";
                }

                //業務名称
                if (textBox1.Text != "")
                {
                    cmd.CommandText += "AND aj.AnkenGyoumuMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(textBox1.Text, 1) + "%' ESCAPE '\\' ";
                }
                //cmd.CommandText += "AND gjm.GyoumuJouhouMadoguchiID = (select min(GyoumuJouhouMadoguchiID) from GyoumuJouhouMadoguchi gm2 where gm2.GyoumuJouhouID = gjm.GyoumuJouhouID group by GyoumuJouhouID) ";
                cmd.CommandText += "AND (gjm.GyoumuJouhouMadoguchiID = (select min(GyoumuJouhouMadoguchiID) from GyoumuJouhouMadoguchi gm2 where gm2.GyoumuJouhouID = gjm.GyoumuJouhouID group by GyoumuJouhouID) or gjm.GyoumuJouhouMadoguchiID is null) ";
                cmd.CommandText += "ORDER BY aj.AnkenAnkenBangou,aj.AnkenJutakuBangou ";
                var sda = new SqlDataAdapter(cmd);
                ListData.Clear();
                sda.Fill(ListData);
            }
            Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
            Paging_now.Text = (1).ToString();
            set_data(1);
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


        private void set_jimushoData(int pagenum)
        {
            //Grid準備
            c1FlexGrid2.Rows.Count = 1;
            c1FlexGrid2.AllowAddNew = true;
            //表示件数設定
            int viewnum = 20;
            int startrow = (pagenum - 1) * viewnum;
            int addnum = JimushoData.Rows.Count - startrow;
            if (addnum > viewnum)
            {
                addnum = viewnum;
            }
            //事務所データセット
            for (int r = 0; r < addnum; r++)
            {
                c1FlexGrid2.Rows.Add();
                for (int i = 0; i < c1FlexGrid2.Cols.Count - 1; i++)
                {
                    c1FlexGrid2[r + 1, i + 1] = JimushoData.Rows[startrow + r][i];
                }

            }
            c1FlexGrid2.AllowAddNew = false;
            set_jimu_page(int.Parse(Paging_now2.Text), int.Parse(Paging_all2.Text));
        }


        private void set_jimu_page(int now, int last)
        {
            if (now <= 1)
            {
                Top_Page2.Enabled = false;
                Previous_Page2.Enabled = false;
            }
            else
            {
                Top_Page2.Enabled = true;
                Previous_Page2.Enabled = true;
            }
            if (now >= last)
            {
                End_Page2.Enabled = false;
                After_Page2.Enabled = false;
            }
            else
            {
                End_Page2.Enabled = true;
                After_Page2.Enabled = true;
            }
        }


        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid1.HitTest(new Point(e.X, e.Y));

            //if (hti.Column == 0 & hti.Row != 0)
            if (hti.Column == 0 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                //窓口画面へ渡すものをセット
                //案件番号　受託番号　受託枝番号
                ReturnValue[0] = c1FlexGrid1.Rows[_row][_col + 2].ToString();
                ReturnValue[1] = c1FlexGrid1.Rows[_row][_col + 3].ToString();
                ReturnValue[2] = c1FlexGrid1.Rows[_row][_col + 4].ToString();
                //契約担当者　契約担当者CD
                ReturnValue[3] = c1FlexGrid1.Rows[_row][_col + 9].ToString();
                ReturnValue[4] = c1FlexGrid1.Rows[_row][_col + 10].ToString();
                //業務管理者　業務管理者CD
                ReturnValue[5] = c1FlexGrid1.Rows[_row][_col + 12].ToString();
                ReturnValue[6] = c1FlexGrid1.Rows[_row][_col + 13].ToString();
                //案件情報ID　管理技術者　管理技術者CD
                ReturnValue[7] = c1FlexGrid1.Rows[_row][_col + 14].ToString();
                ReturnValue[8] = c1FlexGrid1.Rows[_row][_col + 15].ToString();
                ReturnValue[9] = c1FlexGrid1.Rows[_row][_col + 16].ToString();
                //契約区分　受託部所 部所所属長
                ReturnValue[10] = c1FlexGrid1.Rows[_row][_col + 17].ToString();
                ReturnValue[11] = c1FlexGrid1.Rows[_row][_col + 18].ToString();
                ReturnValue[12] = c1FlexGrid1.Rows[_row][_col + 19].ToString();
                //発注者課名　業務名称
                ReturnValue[13] = c1FlexGrid1.Rows[_row][_col + 6].ToString();
                ReturnValue[14] = c1FlexGrid1.Rows[_row][_col + 7].ToString();
                // 事業部HeadCDが調査部の場合、案件受託フォルダを親画面に返す
                if (c1FlexGrid1.Rows[_row][_col + 21] != null && c1FlexGrid1.Rows[_row][_col + 21].ToString() == "T")
                {
                    // 案件（受託）フォルダ
                    ReturnValue[15] = c1FlexGrid1.Rows[_row][_col + 20].ToString();
                }
                //窓口担当者ＣＤ　名
                ReturnValue[16] = c1FlexGrid1.Rows[_row][_col + 22].ToString();
                ReturnValue[17] = c1FlexGrid1.Rows[_row][_col + 23].ToString();
                //窓口部所支部ＣＤ　名
                ReturnValue[18] = c1FlexGrid1.Rows[_row][_col + 24].ToString();
                ReturnValue[19] = c1FlexGrid1.Rows[_row][_col + 25].ToString();

                //単価テーブルを見る
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    string jutakuNo = c1FlexGrid1.Rows[_row][_col + 3].ToString();

                    JimushoData.Clear();

                    if(jutakuNo != "")
                    {
                        var cmd = conn.CreateCommand();
                        cmd.CommandText = "SELECT tk.TankakeiyakuJutakuBangou" + //受託番号
                            ",mk.KoujijimushoMei " +　//事務所名
                            ",mkt.KoujiTantoushaMei " + //担当者名
                            ",tk.TankaKeiyakuID " +//業務CD
                            ",mk.KoujijimushoUketsukeNo " +//発注機関受付番号
                                                           //◆単品タブ用
                            ",mkt.KoujiTantoushaBusho " + //部署
                            ",mkt.KoujiTantoushaYakushoku " +//役職
                            ",mkt.KoujiTantoushaTEL " + //TEL
                            ",mkt.KoujiTantoushaFAX " + //FAX
                            ",mkt.KoujiTantoushaMail " + //MAIL
                            "FROM TankaKeiyaku tk " +
                            "INNER JOIN Mst_Koujijimusho mk " +
                            //"LEFT JOIN Mst_Koujijimusho mk " +
                            "ON tk.TankaKeiyakuID = mk.TankaKeiyakuID " +
                            "AND TankakeiyakuDeleteFlag = 0 " +
                            //"INNER JOIN Mst_KoujijimushoTantousha mkt " +
                            "LEFT JOIN Mst_KoujijimushoTantousha mkt " +
                            "ON mk.KoujijimushoID = mkt.KoujijimushoID " +
                            "AND mk.KoujijimushoDeleteFlag = 0 " +
                            "AND mkt.KoujiTantoushaDeleteFlag = 0 " +
                            "WHERE tk.TankakeiyakuJutakuBangou = '" + jutakuNo + "' " +
                            " ORDER BY KoujijimushoMei,KoujiTantoushaBusho,KoujiTantoushaMei";
                        var sda = new SqlDataAdapter(cmd);
                        //JimushoData.Clear();
                        sda.Fill(JimushoData);
                    }

                    //単価、事務所データを取得できた場合
                    if (JimushoData.Rows.Count > 0)
                    {
                        //事務所用Gridを可視化
                        groupBox3.Visible = true;
                        groupBox4.Visible = true;
                        textBox6.Text = jutakuNo;
                        //Gridセット
                        set_jimushoData(1);

                    }
                    //できなかった場合
                    else
                    {
                        //単品タブの項目をクリアにして渡す
                        ReturnValue[20] = "";
                        ReturnValue[21] = "";
                        ReturnValue[22] = "";
                        //TEL　FAX　MEIL
                        ReturnValue[23] = "";
                        ReturnValue[24] = "";
                        ReturnValue[25] = "";
                        ReturnValue[26] = "";
                        //プロンプトを閉じて画面へ戻る
                        this.Close();
                    }
                }//SqlConnection
                Paging_all2.Text = (Math.Ceiling((double)JimushoData.Rows.Count / 20)).ToString();
                Paging_now2.Text = (1).ToString();
                set_jimushoData(1);
            }
        }

        private void c1FlexGrid2_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid2.HitTest(new Point(e.X, e.Y));

            //if (hti.Column == 0 & hti.Row != 0)
            if (hti.Column == 0 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                //事務所、単価情報を窓口画面へ渡す
                //部署 役職　担当者名　
                ReturnValue[20] = c1FlexGrid2.Rows[_row][_col + 6].ToString(); // 6:部署
                ReturnValue[21] = c1FlexGrid2.Rows[_row][_col + 7].ToString(); // 7:役職
                ReturnValue[22] = c1FlexGrid2.Rows[_row][_col + 3].ToString(); // 3:担当者名
                //TEL　FAX　MEIL 
                ReturnValue[23] = c1FlexGrid2.Rows[_row][_col + 8].ToString(); // 8:TEL
                ReturnValue[24] = c1FlexGrid2.Rows[_row][_col + 9].ToString(); // 9:FAX
                ReturnValue[25] = c1FlexGrid2.Rows[_row][_col + 10].ToString(); // 10:MAIL
                //発注課名に事務所名を加える
                //ReturnValue[13] = ReturnValue[13] + " " + c1FlexGrid2.Rows[_row][_col + 2].ToString(); //2:事務所名
                // 工事事務所名だけを反映
                ReturnValue[13] = c1FlexGrid2.Rows[_row][_col + 2].ToString(); //2:事務所名
                ReturnValue[26] = c1FlexGrid2.Rows[_row][_col + 5].ToString(); // 5:発注機関・受付番号

                //プロンプトを閉じて画面へ戻る
                this.Close();
            }
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

        private void textChanged(object sender, EventArgs e)
        {
            get_data();
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
			if (comboBox1.SelectedIndex != 0)
            {
                Nendo = comboBox1.SelectedValue.ToString();
                set_combo_shibu();
                get_data();
            }
            set_nendo();
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

        private void Top_Page_Click2(object sender, EventArgs e)
        {
            Paging_now2.Text = (1).ToString();
            set_jimushoData(int.Parse(Paging_now2.Text));
        }

        private void Previous_Page_Click2(object sender, EventArgs e)
        {
            Paging_now2.Text = (int.Parse(Paging_now2.Text) - 1).ToString();
            set_jimushoData(int.Parse(Paging_now2.Text));
        }
        private void After_Page_Click2(object sender, EventArgs e)
        {
            Paging_now2.Text = (int.Parse(Paging_now2.Text) + 1).ToString();
            set_jimushoData(int.Parse(Paging_now2.Text));
        }

        private void End_Page_Click2(object sender, EventArgs e)
        {
            Paging_now2.Text = (int.Parse(Paging_all2.Text)).ToString();
            set_jimushoData(int.Parse(Paging_now2.Text));
        }
        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            set_combo_shibu();
            get_data();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
        }

		private void comboBox1_TextChanged_1(object sender, EventArgs e)
		{

		}
	}
}
