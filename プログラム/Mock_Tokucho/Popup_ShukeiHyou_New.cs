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
using System.Security.Cryptography.X509Certificates;
using System.IO;
using C1.Win.C1FlexGrid;

namespace TokuchoBugyoK2
{
    public partial class Popup_ShukeiHyou_New : Form
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

        // 奉行エクセル移管対応
        //対象のグループ一覧用リスト
        private List<string> BushoList = new List<string>();
        private List<string> TokuchoList = new List<string>();
        private List<string> KojincdList = new List<string>();
        private List<string> ChousainMeiList = new List<string>();
        private List<string> GroupMeiList = new List<string>();
        private List<string> BunkatsuList = new List<string>();
        private List<string> ShukeiVerList = new List<string>();

        private List<string> GroupIDList = new List<string>();
        private List<string> kojinList = new List<string>();

        //選択帳票の集計表Ver
        private int ShukeiVer;
        private string chousainShukeiFolder;

        // VIPS　20220316　課題管理表No1263(957)　ADD　自分大臣の時だけファイルDLタイプの選択を追加
        Popup_Download download_form = null;
        // VIPS　20220316　課題管理表No1263(957)　DEL  DLのフォルダ存在チェック行わない
        //private Boolean existFolder = false;

        public Popup_ShukeiHyou_New()
        {
            InitializeComponent();

            // コンボボックスにマウスホイールイベントを付与
            this.src_Busho.MouseWheel += item_MouseWheel;
            this.comboBox_Taisho.MouseWheel += item_MouseWheel;
            this.comboBox_Chohyo.MouseWheel += item_MouseWheel;

        }

        private void Popup_ShukeiHyou_New_Load(object sender, EventArgs e)
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
            // 奉行エクセル移管対応 20231004 対象のグループ一覧リスト初期化
            BushoList.Clear();
            TokuchoList.Clear();
            ChousainMeiList.Clear();
            GroupMeiList.Clear();
            BunkatsuList.Clear();
            kojinList.Clear();

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

                // 奉行エクセル移管対応 20231004
                if (comboBox_Chohyo.SelectedValue != null)
                {
                    // 選択帳票の集計表Verを取得
                    cmd.CommandText = "SELECT " +
                            "PrintShukeiVer " +
                            "FROM Mst_PrintList " +
                            "WHERE PrintListID = '" + comboBox_Chohyo.SelectedValue.ToString() + "' ";
                    var sde = new SqlDataAdapter(cmd);
                    DataTable dtVer = new DataTable();
                    sde.Fill(dtVer);
                    if (dtVer != null && dtVer.Rows.Count > 0)
                    {
                        int.TryParse(dtVer.Rows[0][0].ToString(), out ShukeiVer);
                    }
                }

                // グループ一覧用対象調査品目データの取得
                if (checkBox_BushoIkkatu.Checked)
                {
                    // 調査員リストのみ取得
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
                                "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                                "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                                "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                                "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                        var sdc = new SqlDataAdapter(cmd);
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
                                "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                                "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                                "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                                "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                        var sdc = new SqlDataAdapter(cmd);
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
                                "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                                "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                                "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                                "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                        var sdc = new SqlDataAdapter(cmd);
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
                                }
                            }
                        }
                    }
                    hinmokuListSelect();
                }
                else
                {
                    if (item1_KojinCD.Text != null && item1_KojinCD.Text != "")
                    {
                        kojinList.Add(item1_KojinCD.Text);
                        hinmokuListSelect();
                    }
                    // 以下は動確OKであれば消すこと！
                    ////  1:主のデータを取得
                    //if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "1"))
                    //{
                    //    cmd.CommandText = "SELECT distinct " +
                    //            "mc.ChousainMei " +
                    //            ",mg.MadoguchiGroupMei " +
                    //            ",ch.ChousaBunkatsuHouhou " +
                    //            "FROM ChousaHinmoku ch " +
                    //            "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ch.HinmokuChousainCD = mjmc.MadoguchiL1ChousaTantoushaCD " +
                    //            "LEFT JOIN Mst_Chousain mc ON ch.HinmokuChousainCD = mc.KojinCD " +
                    //            "LEFT JOIN MadoguchiGroupMaster mg ON ch.ChousaMadoguchiGroupMasterID = mg.MadoguchiGroupMasterID " +
                    //            "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND ch.HinmokuChousainCD = '" + item1_KojinCD.Text + "' " +
                    //            "AND ch.HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                    //            "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                    //            "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                    //            "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                    //            "AND mjmc.MadoguchiL1ChousaShinchoku != 80 " +
                    //            "ORDER BY mg.MadoguchiGroupMei, ch.ChousaBunkatsuHouhou";
                    //    var sdb = new SqlDataAdapter(cmd);
                    //    DataTable dt0 = new DataTable();
                    //    sdb.Fill(dt0);

                    //    if (dt0 != null && dt0.Rows.Count > 0)
                    //    {
                    //        for (int i = 0; i < dt0.Rows.Count; i++)
                    //        {
                    //            BushoList.Add(src_Busho.Text);
                    //            ChousainMeiList.Add(dt0.Rows[i][0].ToString());
                    //            TokuchoList.Add(TokuhoBangou.ToString() + "-" + TokuhoBangouEda.ToString());
                    //            GroupMeiList.Add(dt0.Rows[i][1].ToString());
                    //            BunkatsuList.Add(dt0.Rows[i][2].ToString());
                    //        }
                    //    }
                    //}
                    //// 2:副1と副2のデータを取得
                    //if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "2"))
                    //{
                    //    cmd.CommandText = "SELECT distinct " +
                    //            "mc.ChousainMei " +
                    //            ",mg.MadoguchiGroupMei " +
                    //            ",ch.ChousaBunkatsuHouhou " +
                    //            "FROM ChousaHinmoku ch " +
                    //            "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ((ch.HinmokuFukuChousainCD1 = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                    //            "OR (ch.HinmokuFukuChousainCD2 = mjmc.MadoguchiL1ChousaTantoushaCD)) " +
                    //            "LEFT JOIN Mst_Chousain mc ON (ch.HinmokuFukuChousainCD1 = mc.KojinCD) OR (ch.HinmokuFukuChousainCD2 = mc.KojinCD)" +
                    //            "LEFT JOIN MadoguchiGroupMaster mg ON ch.ChousaMadoguchiGroupMasterID = mg.MadoguchiGroupMasterID " +
                    //            "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND (( ch.HinmokuFukuChousainCD1 = '" + item1_KojinCD.Text + "' ) " +
                    //            "OR (ch.HinmokuFukuChousainCD2 = '" + item1_KojinCD.Text + "' )) " +
                    //            "AND ch.HinmokuRyakuBushoFuku1CD = '" + src_Busho.SelectedValue.ToString() + "' " +
                    //            "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                    //            "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                    //            "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                    //            "AND mjmc.MadoguchiL1ChousaShinchoku != 80 " +
                    //            "ORDER BY mg.MadoguchiGroupMei, ch.ChousaBunkatsuHouhou";
                    //    var sdb = new SqlDataAdapter(cmd);
                    //    DataTable dt0 = new DataTable();
                    //    sdb.Fill(dt0);

                    //    if (dt0 != null && dt0.Rows.Count > 0)
                    //    {
                    //        for (int i = 0; i < dt0.Rows.Count; i++)
                    //        {
                    //            BushoList.Add(src_Busho.Text);
                    //            ChousainMeiList.Add(dt0.Rows[i][0].ToString());
                    //            TokuchoList.Add(dt0.Rows[i][1].ToString() + "-" + dt0.Rows[i][2].ToString());
                    //            GroupMeiList.Add(dt0.Rows[i][3].ToString());
                    //            BunkatsuList.Add(dt0.Rows[i][4].ToString());
                    //        }
                    //    }
                    //}
                    //// 0:主+副
                    //if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "0"))
                    //{
                    //    cmd.CommandText = "SELECT distinct " +
                    //            "mc.ChousainMei " +
                    //            ",mg.MadoguchiGroupMei " +
                    //            ",ch.ChousaBunkatsuHouhou " +
                    //            "FROM ChousaHinmoku ch " +
                    //            "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ((ch.HinmokuChousainCD = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                    //            "OR (ch.HinmokuFukuChousainCD1 = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                    //            "OR (ch.HinmokuFukuChousainCD2 = mjmc.MadoguchiL1ChousaTantoushaCD)) " +
                    //            "LEFT JOIN Mst_Chousain mc ON ((ch.HinmokuChousainCD = mc.KojinCD) OR (ch.HinmokuFukuChousainCD1 = mc.KojinCD) OR (ch.HinmokuFukuChousainCD2 = mc.KojinCD) " +
                    //            "LEFT JOIN MadoguchiGroupMaster mg ON ch.ChousaMadoguchiGroupMasterID = mg.MadoguchiGroupMasterID " +
                    //            "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND (( ch.HinmokuChousainCD = '" + item1_KojinCD.Text + "' ) " +
                    //            "OR (ch.HinmokuFukuChousainCD1 = '" + item1_KojinCD.Text + "' ) " +
                    //            "OR (ch.HinmokuFukuChousainCD2 = '" + item1_KojinCD.Text + "' )) " +
                    //            "AND ch.HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                    //            "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                    //            "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                    //            "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                    //            "AND mjmc.MadoguchiL1ChousaShinchoku != 80 " +
                    //            "ORDER BY mg.MadoguchiGroupMei, ch.ChousaBunkatsuHouhou";
                    //    var sdb = new SqlDataAdapter(cmd);
                    //    DataTable dt0 = new DataTable();
                    //    sdb.Fill(dt0);

                    //    if (dt0 != null && dt0.Rows.Count > 0)
                    //    {
                    //        for (int i = 0; i < dt0.Rows.Count; i++)
                    //        {
                    //            BushoList.Add(src_Busho.Text);
                    //            ChousainMeiList.Add(dt0.Rows[i][0].ToString());
                    //            TokuchoList.Add(dt0.Rows[i][1].ToString() + "-" + dt0.Rows[i][2].ToString());
                    //            GroupMeiList.Add(dt0.Rows[i][3].ToString());
                    //            BunkatsuList.Add(dt0.Rows[i][4].ToString());
                    //        }
                    //    }
                    //}
                }

                if (ShukeiVer == 2 && !checkBox_Zenhinmoku.Checked)
                {
                    c1FlexGrid2.Rows.Count = 1;
                    c1FlexGrid2.AllowAddNew = true;
                    int addrow = BushoList.Count;

                    for (int r = 0; r < addrow; r++)
                    {
                        c1FlexGrid2.Rows.Add();
                        {
                            c1FlexGrid2[r + 1, 1] = BushoList[r].ToString();
                            c1FlexGrid2[r + 1, 2] = ChousainMeiList[r].ToString();
                            c1FlexGrid2[r + 1, 3] = GroupMeiList[r].ToString();
                            if (BunkatsuList[r] == "1")
                            {
                                c1FlexGrid2[r + 1, 4] = "シート分割";
                            }
                            else
                            {
                                c1FlexGrid2[r + 1, 4] = "ファイル分割";
                            }
                        }

                    }
                    c1FlexGrid2.AllowAddNew = false;
                }
                else
                {
                    c1FlexGrid2.Rows.Count = 1;
                }

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
            // 奉行エクセル移管対応 20231004
            if (!checkBox_BushoIkkatu.Checked)
            {
                Paging_all.Text = (Math.Ceiling((double)ListData.Rows.Count / 20)).ToString();
                Paging_now.Text = (1).ToString();
                set_data(1);
                //Resize_Grid("c1FlexGrid1");
            }
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

                    // 奉行エクセル移管対応 20231004
                    get_data();
                    getFileName();

                    // 奉行エクセル移管対応 20231004
                    // エラー背景色　クリア
                    Color clearColor = Color.FromArgb(255, 255, 255);
                    int filerow = c1FlexGrid3.Rows.Count;
                    for (int r = 0; r < filerow; r++)
                    {
                        c1FlexGrid3.GetCellRange(r, 0).StyleNew.BackColor = clearColor;
                    }

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
                        // 奉行エクセル移管対応 20231004
                        ////  VIPS　20220322　課題管理表No1263(957)　ADD  保存にチェックがついていて、ファイルが存在する場合にエラー
                        //// フォルダ + ファイル名存在チェック
                        //if (File.Exists(item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text) && radioButton_Save.Checked)
                        //{
                        //    // E20332:集計表ファイルが既に存在します。
                        //    set_error("", 0);
                        //    set_error(GlobalMethod.GetMessage("E20332", ""));
                        //    // ファイル出力ボタンを非活性化
                        //    btnFileExport.Enabled = false;
                        //}
                        //else
                        //{
                        //    set_error("", 0);
                        //    // ファイル出力ボタンを活性化
                        //    btnFileExport.Enabled = true;
                        //}
                        int prntflg = 0;
                        filerow = c1FlexGrid3.Rows.Count;
                        // エラー背景色
                        Color errorColor = Color.FromArgb(255, 204, 255);
                        for (int r = 0; r < filerow; r++)
                        {
                            if (File.Exists(item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[r, 0]) && radioButton_Save.Checked)
                            {
                                // E20332:集計表ファイルが既に存在します。
                                set_error("", 0);
                                set_error(GlobalMethod.GetMessage("E20332", ""));
                                // ファイル出力ボタンを非活性化
                                btnFileExport.Enabled = false;
                                c1FlexGrid3.GetCellRange(r, 0).StyleNew.BackColor = errorColor;
                            }
                            else
                            {
                                // ファイル出力ボタンを活性化
                                prntflg = 1;
                            }
                        }
                        // 出力可能なファイルがあればファイル出力ボタンを活性化
                        if (prntflg == 1)
                        {
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

                // 奉行エクセル移管対応 20231004
                //item1_PritFileName.Enabled = false;
                get_data();
                getFileName();
                // エラー背景色　クリア
                Color clearColor = Color.FromArgb(255, 255, 255);
                int filerow = c1FlexGrid3.Rows.Count;
                for (int r = 0; r < filerow; r++)
                {
                    c1FlexGrid3.GetCellRange(r, 0).StyleNew.BackColor = clearColor;
                }

                // 奉行エクセル移管対応 20231004
                ////  VIPS　20220322　課題管理表No1263(957)　ADD  保存にチェックがついていて、かつ、フォルダがみつからない場合にエラー
                //// フォルダチェック
                //if (!Directory.Exists(item1_ShukeiFolder.Text) && radioButton_Save.Checked)
                //{
                //    // 集計表フォルダがみつかりません。
                //    set_error("", 0);
                //    set_error(GlobalMethod.GetMessage("E20331", ""));
                //    // ファイル出力ボタンを非活性化
                //    btnFileExport.Enabled = false;
                //}
                //else
                //{
                //    btnFileExport.Enabled = true;
                //}
                int prntflg = 0;
                filerow = c1FlexGrid3.Rows.Count;
                // エラー背景色
                Color errorColor = Color.FromArgb(255, 204, 255);
                for (int r = 0; r < filerow; r++)
                {
                    if (File.Exists(item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[r, 0]) && radioButton_Save.Checked)
                    {
                        // E20332:集計表ファイルが既に存在します。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20332", ""));
                        // ファイル出力ボタンを非活性化
                        btnFileExport.Enabled = false;
                        c1FlexGrid3.GetCellRange(r, 0).StyleNew.BackColor = errorColor;
                    }
                    else
                    {
                        // ファイル出力ボタンを活性化
                        prntflg = 1;
                    }
                }
                // 出力可能なファイルがあればファイル出力ボタンを活性化
                if (prntflg == 1)
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

                // 奉行エクセル移管対応 20231004
                //item1_PritFileName.Enabled = true;
                c1FlexGrid2.Rows.Count = 1;
                c1FlexGrid3.Rows.Count = 0;
                // ファイル出力ボタンを非活性化
                btnFileExport.Enabled = false;
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

                // 奉行エクセル移管対応
                get_data();
                if (ShukeiVer == 1)
                {
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
                        //// VIPS 20220322 課題管理表No1263(957) ADD  保存にチェックがついていて、ファイルが存在する場合にエラー
                        //if (File.Exists(item1_ShukeiFolder.Text + @"\" + item1_PritFileName) && radioButton_Save.Checked)
                        //{
                        //    // ファイルが存在する
                        //    // E20332:集計表ファイルが既に存在します。
                        //    set_error("", 0);
                        //    set_error(GlobalMethod.GetMessage("E20332", ""));

                        //    // ファイル出力ボタンを非活性化
                        //    btnFileExport.Enabled = false;
                        //}
                        //else
                        //{
                        //    // ファイル出力ボタンを活性化
                        //    btnFileExport.Enabled = true;
                        //}
                        int prntflg = 0;
                        int filerow = c1FlexGrid3.Rows.Count;
                        // エラー背景色
                        Color errorColor = Color.FromArgb(255, 204, 255);
                        for (int r = 0; r < filerow; r++)
                        {
                            if (File.Exists(item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[r, 0]) && radioButton_Save.Checked)
                            {
                                // E20332:集計表ファイルが既に存在します。
                                set_error("", 0);
                                set_error(GlobalMethod.GetMessage("E20332", ""));
                                // ファイル出力ボタンを非活性化
                                btnFileExport.Enabled = false;
                                c1FlexGrid3.GetCellRange(r, 0).StyleNew.BackColor = errorColor;
                            }
                            else
                            {
                                // ファイル出力ボタンを活性化
                                prntflg = 1;
                            }
                        }
                        // 出力可能なファイルがあればファイル出力ボタンを活性化
                        if (prntflg == 1)
                        {
                            btnFileExport.Enabled = true;
                        }
                    }
                }
                else
                {
                    // E20XXX:全品目一括集計表で集計表Ver2は選択できません。
                    set_error("", 0);
                    set_error(GlobalMethod.GetMessage("E20XXX", ""));
                    // ファイル出力ボタンを非活性化
                    btnFileExport.Enabled = false;
                }
            }
            else
            {
                Hinmoku_All = false;
                //gridと対象表示
                groupBox2.Visible = true;
                groupBox3.Visible = true;
                groupBox1.Visible = true;

                // 奉行エクセル移管対応 20231004
                //// ファイル名を空に
                //item1_PritFileName.Text = "";
                c1FlexGrid2.Rows.Count = 1;
                c1FlexGrid3.Clear(ClearFlags.Content);
                c1FlexGrid3.Rows.Count = 1;
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
            // ファイル名を初期化
            c1FlexGrid3.Clear(ClearFlags.Content);
            c1FlexGrid3.Rows.Count = 1;

            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                if (comboBox_Chohyo.SelectedValue != null)
                {
                    // 奉行エクセル移管対応 20231004
                    if (ShukeiVer == 2)
                    {
                        btnFileExport.Text = "フォルダ作成と出力";
                    }
                    else
                    {
                        btnFileExport.Text = "ファイル出力";
                    }
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
                // 奉行エクセル移管対応 20231004
                //item1_PritFileName.Text = "一括集計表" + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                c1FlexGrid3.Rows.Count = 1;
                c1FlexGrid3.AllowAddNew = true;
                c1FlexGrid3.Rows.Add();
                c1FlexGrid3[0, 0] = "一括集計表" + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                c1FlexGrid3.AllowAddNew = false;
            }
            else
            {
                // 奉行エクセル移管対応 20231004
                //if (label_SentakuTantousha.Text != "")
                if (label_SentakuTantousha.Text != "" || checkBox_BushoIkkatu.Checked)
                {
                    // 奉行エクセル移管対応 20231004
                    //item1_PritFileName.Text = label_SentakuTantousha.Text + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                    int addrow = ChousainMeiList.Count;
                    c1FlexGrid3.Rows.Count = 1;
                    c1FlexGrid3.AllowAddNew = true;
                    if (ShukeiVer == 2)
                    {
                        for (int r = 0; r < addrow; r++)
                        {
                            c1FlexGrid3.Rows.Add();
                            {
                                if (c1FlexGrid2.Rows[r + 1][4].ToString() == "シート分割")
                                {
                                    c1FlexGrid3[r, 0] = ChousainMeiList[r].ToString() + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                                }
                                else
                                {
                                    c1FlexGrid3[r, 0] = ChousainMeiList[r].ToString() + "-" + TokuhoBangou + "-" + TokuhoBangouEda + "【" + GroupMeiList[r].ToString() + "】" + extensions;
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int r = 0; r < addrow; r++)
                        {
                            c1FlexGrid3.Rows.Add();
                            {
                                c1FlexGrid3[r, 0] = ChousainMeiList[r].ToString() + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                            }
                        }
                    }
                    c1FlexGrid3.AllowAddNew = false;

                }
                else
                {
                    // 奉行エクセル移管対応 20231004
                    //item1_PritFileName.Text = "未登録" + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                    c1FlexGrid2.Rows.Count = 1;
                    c1FlexGrid3.Clear(ClearFlags.Content);
                    c1FlexGrid3.Rows.Count = 1;
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
            // エラー背景色
            Color errorColor = Color.FromArgb(255, 204, 255);
            int prntflg = 1;

            // 部所一括集計表出力以外の場合
            if (!checkBox_BushoIkkatu.Checked)
            {
                // 奉行エクセル移管対応 20231004
                //// VIPS 20220322 課題管理表No1263(957) ADD  保存にチェックがついていて、かつ、ファイルが存在する場合にエラー
                //// ファイル存在チェック
                //if (File.Exists(item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text) && radioButton_Save.Checked)
                //{
                //    // 既にファイルが存在する
                //    set_error("", 0);
                //    set_error(GlobalMethod.GetMessage("E20332", "") + ":" + item1_PritFileName.Text);
                //    return;
                //}

                // 集計表Ver1、Ver2混在チェック
                if (!fileErrorCheck(item1_KojinCD.Text))
                {
                    prntflg = 0;
                }

                // 集計表VerがVer2の場合、調査員単位で分割方法がファイル分割・シート分割混在でメッセージ出力
                if (ShukeiVer == 2)
                {
                    if (BunkatsuList[0] == "1" && BunkatsuList.Contains("2"))
                    {
                        // E20XXX:分割方法が混在しています。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20XXX", ""));
                        // 混在系エラーは問題あるファイルを特定できないので出力を取りやめる？
                        return;
                    }
                    if (BunkatsuList[0] == "2" && BunkatsuList.Contains("1"))
                    {
                        // E20XXX:分割方法が混在しています。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20XXX", ""));
                        // 混在系エラーは問題あるファイルを特定できないので出力を取りやめる？
                        return;
                    }
                }

                int filerow = c1FlexGrid3.Rows.Count;
                for (int r = 0; r < filerow; r++)
                {
                    // 集計表Ver2でグループ名が選択されていない品目明細があった場合、エラーとする。
                    if ((ShukeiVer == 2) && (GroupMeiList[r].ToString() == "" || GroupMeiList[r].ToString() is null))
                    {
                        c1FlexGrid3.GetCellRange(r, 0).StyleNew.BackColor = errorColor;
                        // E20XXX:グループ名が選択されていない明細が存在します。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20XXX", ""));
                        prntflg = 0;
                    }

                    if (prntflg == 1)
                    {
                        // 集計表フォルダ・作業フォルダ作成
                        chousainShukeiFolder = "";
                        if (!createFolder(r))
                        {
                            prntflg = 0;
                        }

                        if (!checkBox_Zenhinmoku.Checked && ShukeiVer == 2 && BunkatsuList[r] == "2")
                        {
                            if (File.Exists(chousainShukeiFolder + @"\" + c1FlexGrid3[r, 0]) && radioButton_Save.Checked)
                            {
                                // E20332:集計表ファイルが既に存在します。
                                set_error("", 0);
                                set_error(GlobalMethod.GetMessage("E20332", ""));
                                c1FlexGrid3.GetCellRange(r, 0).StyleNew.BackColor = errorColor;
                                prntflg = 0;
                            }
                        }
                        else
                        {
                            if (File.Exists(item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[r, 0]) && radioButton_Save.Checked)
                            {
                                // E20332:集計表ファイルが既に存在します。
                                set_error("", 0);
                                set_error(GlobalMethod.GetMessage("E20332", ""));
                                c1FlexGrid3.GetCellRange(r, 0).StyleNew.BackColor = errorColor;
                                prntflg = 0;
                            }
                        }
                    }

                    // 1:MadoguchiID     窓口ID
                    // 2:ZeninSyukeihyo  全品目一括集計表 1:チェック 0:未チェック
                    // 3:ShibuMei        支部名
                    // 4:KojinCD         個人CD
                    // 5:ShuFuku         主+副  0:主+副 1:主のみ 2:副
                    // 6:FileName        ファイル名
                    // 7:PrintGamen      呼び出し元画面 0:窓口ミハル 1:特命課長  2:自分大臣
                    // 7個分先に用意
                    if (prntflg == 1)
                    {
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
                        report_data[5] = c1FlexGrid3[r, 0].ToString();

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
                                        if (!checkBox_Zenhinmoku.Checked && ShukeiVer == 2 && BunkatsuList[r] == "2")
                                        {
                                            System.IO.File.Copy(result[2], chousainShukeiFolder + @"\" + c1FlexGrid3[r, 0].ToString(), true);
                                        }
                                        else
                                        {
                                            System.IO.File.Copy(result[2], item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[r, 0].ToString(), true);
                                        }
                                        set_error("集計表ファイルを出力しました。:" + c1FlexGrid3[r, 0].ToString());

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
                                                    string linkpath;
                                                    if (!checkBox_Zenhinmoku.Checked && ShukeiVer == 2 && BunkatsuList[r] == "2")
                                                    {
                                                        linkpath = chousainShukeiFolder + @"\" + c1FlexGrid3[r, 0].ToString();
                                                    }
                                                    else
                                                    {
                                                        linkpath = item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[r, 0].ToString();
                                                    }
                                                    // 全品目一括集計表
                                                    cmd.CommandText = "UPDATE ChousaHinmoku SET ChousaLinkSakli = N'" + linkpath + "' " +
                                                        "WHERE " +
                                                        "MadoguchiID = '" + MadoguchiID + "' ";
                                                    // 全品目一括集計表ではない AND 個人CD が0でない場合は、個人のみ更新
                                                    if (!checkBox_Zenhinmoku.Checked && report_data[3] != "0")
                                                    {
                                                        cmd.CommandText += "AND (HinmokuChousainCD = '" + report_data[3] + "' " +
                                                            "OR HinmokuFukuChousainCD1 = '" + report_data[3] + "' " +
                                                            "OR HinmokuFukuChousainCD2 = '" + report_data[3] + "' )" +
                                                            "AND ChousaShuukeihyouVer = " + ShukeiVer + " ";
                                                    }
                                                    // 集計表Verが2の場合、グループ単位の更新
                                                    if (ShukeiVer == 2)
                                                    {
                                                        cmd.CommandText += "AND ChousaMadoguchiGroupMasterID = " + int.Parse(GroupIDList[r]) + " ";
                                                    }
                                                    cmd.ExecuteNonQuery();

                                                    // 担当部所テーブル更新
                                                    string shukeipath;
                                                    if (!checkBox_Zenhinmoku.Checked && ShukeiVer == 2 && BunkatsuList[r] == "2")
                                                    {
                                                        shukeipath = chousainShukeiFolder;
                                                    }
                                                    else
                                                    {
                                                        shukeipath = item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[r, 0].ToString();
                                                    }
                                                    // 全品目一括集計表
                                                    cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET MadoguchiL1ShukeihyoLink = N'" + shukeipath + "' " +
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

                                    String fileName = Path.GetFileName(c1FlexGrid3[r, 0].ToString());
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
            // 部所一括集計表出力
            else
            {
                //奉行エクセル移管対応 get_data内で対象担当者のリストは取得・保持しているためそれを使用する
                //// 対象の担当者リスト
                //List<string> kojinList = new List<string>();
                //List<string> ChousainMeiList = new List<string>();

                // 対象を取得する
                string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

                ////分類
                //using (var conn = new SqlConnection(connStr))
                //{
                //    try
                //    {
                //        var cmd = conn.CreateCommand();
                //        //cmd.CommandText = "SELECT " +
                //        //        "MadoguchiL1ChousaTantoushaCD " +
                //        //        ",MadoguchiL1ChousaTantousha " +
                //        //        ",MadoguchiL1ChousaBushoCD " +
                //        //        "FROM MadoguchiJouhouMadoguchiL1Chou WHERE MadoguchiID = '" + MadoguchiID + "' AND MadoguchiL1ChousaTantoushaCD > 0 " +
                //        //        "AND MadoguchiL1ChousaBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                //        //        // MadoguchiL1ChousaShinchoku = 1 //調査中
                //        //        // 1:調査中　　⇒　40：集計中
                //        //        //"AND MadoguchiL1ChousaShinchoku = 40";
                //        //        // 旧進捗状況の　1:調査中　は 20:調査開始、30:見積中、40：集計中に該当する
                //        //        //"AND MadoguchiL1ChousaShinchoku IN (20,30,40)";
                //        //        "AND MadoguchiL1ChousaShinchoku != 80";

                //        // 主のデータを取得
                //        // 0:主+副 1:主 2:副
                //        if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "0" || comboBox_Taisho.SelectedValue.ToString() == "1"))
                //        {
                //            cmd.CommandText = "SELECT distinct " +
                //                    "HinmokuChousainCD " +
                //                    ",mc.ChousainMei " +
                //                    ",HinmokuRyakuBushoCD " +
                //                    "FROM ChousaHinmoku ch " +
                //                    "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ch.HinmokuChousainCD = mjmc.MadoguchiL1ChousaTantoushaCD " +
                //                    "LEFT JOIN Mst_Chousain mc ON ch.HinmokuChousainCD = mc.KojinCD " +
                //                    "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND ch.HinmokuChousainCD > 0 " +
                //                    "AND HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                //                    "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                //            var sda = new SqlDataAdapter(cmd);
                //            DataTable dt0 = new DataTable();
                //            sda.Fill(dt0);

                //            if (dt0 != null && dt0.Rows.Count > 0)
                //            {
                //                for (int i = 0; i < dt0.Rows.Count; i++)
                //                {
                //                    // 重複除外
                //                    if (!kojinList.Contains(dt0.Rows[i][0].ToString()))
                //                    {
                //                        kojinList.Add(dt0.Rows[i][0].ToString());
                //                        ChousainMeiList.Add(dt0.Rows[i][1].ToString());
                //                    }
                //                }
                //            }
                //        }
                //        // 副1のデータを取得
                //        if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "0" || comboBox_Taisho.SelectedValue.ToString() == "2"))
                //        {
                //            cmd.CommandText = "SELECT distinct " +
                //                    "HinmokuFukuChousainCD1 " +
                //                    ",mc.ChousainMei " +
                //                    ",HinmokuRyakuBushoFuku1CD " +
                //                    "FROM ChousaHinmoku ch " +
                //                    "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ch.HinmokuFukuChousainCD1 = mjmc.MadoguchiL1ChousaTantoushaCD " +
                //                    "LEFT JOIN Mst_Chousain mc ON ch.HinmokuFukuChousainCD1 = mc.KojinCD " +
                //                    "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND ch.HinmokuFukuChousainCD1 > 0 " +
                //                    "AND HinmokuRyakuBushoFuku1CD = '" + src_Busho.SelectedValue.ToString() + "' " +
                //                    "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                //            var sda = new SqlDataAdapter(cmd);
                //            DataTable dt0 = new DataTable();
                //            sda.Fill(dt0);

                //            if (dt0 != null && dt0.Rows.Count > 0)
                //            {
                //                for (int i = 0; i < dt0.Rows.Count; i++)
                //                {
                //                    // 重複除外
                //                    if (!kojinList.Contains(dt0.Rows[i][0].ToString()))
                //                    {
                //                        kojinList.Add(dt0.Rows[i][0].ToString());
                //                        ChousainMeiList.Add(dt0.Rows[i][1].ToString());
                //                    }
                //                }
                //            }
                //        }
                //        // 副2のデータを取得
                //        if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "0" || comboBox_Taisho.SelectedValue.ToString() == "2"))
                //        {
                //            cmd.CommandText = "SELECT distinct " +
                //                    "HinmokuFukuChousainCD2 " +
                //                    ",mc.ChousainMei " +
                //                    ",HinmokuRyakuBushoFuku2CD " +
                //                    "FROM ChousaHinmoku ch " +
                //                    "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ch.HinmokuFukuChousainCD2 = mjmc.MadoguchiL1ChousaTantoushaCD " +
                //                    "LEFT JOIN Mst_Chousain mc ON ch.HinmokuFukuChousainCD2 = mc.KojinCD " +
                //                    "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND ch.HinmokuFukuChousainCD2 > 0 " +
                //                    "AND HinmokuRyakuBushoFuku2CD = '" + src_Busho.SelectedValue.ToString() + "' " +
                //                    "AND mjmc.MadoguchiL1ChousaShinchoku != 80";
                //            var sda = new SqlDataAdapter(cmd);
                //            DataTable dt0 = new DataTable();
                //            sda.Fill(dt0);

                //            if (dt0 != null && dt0.Rows.Count > 0)
                //            {
                //                for (int i = 0; i < dt0.Rows.Count; i++)
                //                {
                //                    // 重複除外
                //                    if (!kojinList.Contains(dt0.Rows[i][0].ToString()))
                //                    {
                //                        kojinList.Add(dt0.Rows[i][0].ToString());
                //                        ChousainMeiList.Add(dt0.Rows[i][1].ToString());
                //                    }
                //                }
                //            }
                //        }
                //        conn.Close();
                //    }
                //    catch (Exception)
                //    {
                //        //    // エラーが発生しました
                //        //    label3.Text = GlobalMethod.GetMessage("E00091", "");
                //        //    label3.Visible = true;
                //    }
                //}
                // 対象者がいる場合
                //if(dt0.Rows.Count > 0)
                //if (kojinList.Count > 0)
                if (KojincdList.Count > 0)
                {
                    //// VIPS　20220322　課題管理表No1263(957)　ADD保存にチェックがついていて、かつ、フォルダがみつからない場合にエラー
                    //// フォルダチェック
                    //if (!Directory.Exists(item1_ShukeiFolder.Text) && radioButton_Save.Checked)
                    //{
                    //    // 集計表フォルダがみつかりません。
                    //    set_error("", 0);
                    //    set_error(GlobalMethod.GetMessage("E20331", ""));
                    //    return;
                    //}
                    String extensions = ".xlsm";
                    string fileName = "";
                    string errorMsg = "";

                    //set_error("", 0);
                    //for (int i = 0; dt0.Rows.Count > i; i++)
                    for (int i = 0; KojincdList.Count > i; i++)
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
                        // 奉行エクセル移管対応 20231004
                        //fileName = ChousainMeiList[i].ToString() + "-" + TokuhoBangou + "-" + TokuhoBangouEda + extensions;
                        //// VIPS　20220322　課題管理表No1263(957)　ADD保存にチェックがついていて、かつ、ファイルが存在する場合にエラー
                        //// 存在チェック
                        //if (File.Exists(item1_ShukeiFolder.Text + @"\" + fileName) && radioButton_Save.Checked)
                        //{
                        //    // E20332:集計表ファイルが既に存在します。
                        //    set_error(GlobalMethod.GetMessage("E20332", "") + ":" + fileName);
                        //}

                        // 集計表Ver1、Ver2混在チェック
                        if (!fileErrorCheck(KojincdList[i].ToString()))
                        {
                            prntflg = 0;
                        }

                        // 集計表VerがVer2の場合、調査員単位で分割方法混在チェック
                        //int BunkatsuType;
                        int.TryParse(BunkatsuList[i].ToString(), out int BunkatsuType);
                        if (ShukeiVer == 2)
                        {
                            if (!bunkatsuCheck(KojincdList[i].ToString(), BunkatsuType))
                            {
                                prntflg = 0;
                            }
                        }

                        // 集計表Ver2でグループ名が選択されていない品目明細があった場合、エラーとする。
                        if ((ShukeiVer == 2) && (GroupMeiList[i].ToString() == "" || GroupMeiList[i].ToString() is null))
                        {
                            c1FlexGrid3.GetCellRange(i, 0).StyleNew.BackColor = errorColor;
                            // E20XXX:グループ名が選択されていない明細が存在します。
                            set_error("", 0);
                            set_error(GlobalMethod.GetMessage("E20XXX", ""));
                            prntflg = 0;
                        }

                        if (prntflg == 1)
                        {
                            // 集計表フォルダ・作業フォルダ作成
                            chousainShukeiFolder = "";
                            if (!createFolder(i))
                            {
                                prntflg = 0;
                            }

                            if (ShukeiVer == 2 && BunkatsuList[i] == "2")
                            {
                                if (File.Exists(chousainShukeiFolder + @"\" + c1FlexGrid3[i, 0]) && radioButton_Save.Checked)
                                {
                                    // E20332:集計表ファイルが既に存在します。
                                    set_error("", 0);
                                    set_error(GlobalMethod.GetMessage("E20332", ""));
                                    c1FlexGrid3.GetCellRange(i, 0).StyleNew.BackColor = errorColor;
                                    prntflg = 0;
                                }
                            }
                            else
                            {
                                if (File.Exists(item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[i, 0]) && radioButton_Save.Checked)
                                {
                                    // E20332:集計表ファイルが既に存在します。
                                    set_error("", 0);
                                    set_error(GlobalMethod.GetMessage("E20332", ""));
                                    c1FlexGrid3.GetCellRange(i, 0).StyleNew.BackColor = errorColor;
                                    prntflg = 0;
                                }
                            }

                        }

                        if (prntflg == 1)
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
                            report_data[3] = ChousainMeiList[i].ToString();
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
                                            if (ShukeiVer == 2 && BunkatsuList[i] == "2")
                                            {
                                                System.IO.File.Copy(result[2], chousainShukeiFolder + @"\" + c1FlexGrid3[i, 0].ToString(), true);
                                            }
                                            else
                                            {
                                                System.IO.File.Copy(result[2], item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[i, 0].ToString(), true);
                                            }
                                            set_error("集計表ファイルを出力しました。:" + c1FlexGrid3[i, 0].ToString());

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
                                                        string linkpath;
                                                        if (ShukeiVer == 2 && BunkatsuList[i] == "2")
                                                        {
                                                            linkpath = chousainShukeiFolder + @"\" + c1FlexGrid3[i, 0].ToString();
                                                        }
                                                        else
                                                        {
                                                            linkpath = item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[i, 0].ToString();
                                                        }
                                                        // 全品目一括集計表
                                                        cmd.CommandText = "UPDATE ChousaHinmoku SET ChousaLinkSakli = N'" + linkpath + "' " +
                                                            "WHERE " +
                                                            "MadoguchiID = '" + MadoguchiID + "' ";
                                                        // 個人CD が0出ない場合は、個人のみ更新
                                                        if (report_data[3] != "0")
                                                        {
                                                            cmd.CommandText += "AND (HinmokuChousainCD = '" + report_data[3] + "' " +
                                                                "OR HinmokuFukuChousainCD1 = '" + report_data[3] + "' " +
                                                                "OR HinmokuFukuChousainCD2 = '" + report_data[3] + "' )" +
                                                                "AND ChousaShuukeihyouVer = " + ShukeiVer + " ";
                                                        }
                                                        // 集計表Verが2の場合、グループ単位の更新
                                                        if (ShukeiVer == 2)
                                                        {
                                                            cmd.CommandText += "AND ChousaMadoguchiGroupMasterID = " + int.Parse(GroupIDList[i]) + " ";
                                                        }
                                                        cmd.ExecuteNonQuery();

                                                        // 担当部所テーブル更新
                                                        string shukeipath;
                                                        if (ShukeiVer == 2 && BunkatsuList[i] == "2")
                                                        {
                                                            shukeipath = chousainShukeiFolder;
                                                        }
                                                        else
                                                        {
                                                            shukeipath = item1_ShukeiFolder.Text + @"\" + c1FlexGrid3[i, 0].ToString();
                                                        }
                                                        // 全品目一括集計表
                                                        cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET MadoguchiL1ShukeihyoLink = N'" + shukeipath + "' " +
                                                            ", MadoguchiL1AsteriaKoushinFlag = 1 " +
                                                            "WHERE " +
                                                            "MadoguchiID = '" + MadoguchiID + "' ";

                                                        // ※ファイル出力のループ1回で１ファイル対象だが、１担当で複数グループある場合、調査品目は問題ないが
                                                        // 　MadoguchiJouhouMadoguchiL1Chouは窓口ID＋調査担当者CDで一意になっている？のでグループ分作成できない・・・

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

                                        fileName = Path.GetFileName(c1FlexGrid3[i, 0].ToString());
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

        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        // 奉行エクセル移管対応 20231004　（削除されたコントロールのイベント）
        // ファイル名
        //private void item1_PritFileName_TextChanged(object sender, EventArgs e)
        //{
        //if (item1_PritFileName.Text != "")
        //{
        //    //VIPS 20220322 課題管理表No1263(957) ADD 保存にチェックがついていて、かつ、フォルダがみつからない場合にエラー
        //    // フォルダチェック
        //    if (!Directory.Exists(item1_ShukeiFolder.Text) && radioButton_Save.Checked)
        //    {
        //        // 集計表フォルダがみつかりません。
        //        set_error("", 0);
        //        set_error(GlobalMethod.GetMessage("E20331", ""));
        //        // ファイル出力ボタンを非活性化
        //        btnFileExport.Enabled = false;
        //    }
        //    else
        //    {
        //        //VIPS 20220322 課題管理表No1263(957) ADD 保存にチェックがついていて、ファイルが存在する場合にエラー
        //        // フォルダ + ファイル名存在チェック
        //        if (File.Exists(item1_ShukeiFolder.Text + @"\" + item1_PritFileName.Text) && radioButton_Save.Checked)
        //        {
        //            // E20332:集計表ファイルが既に存在します。
        //            set_error("", 0);
        //            set_error(GlobalMethod.GetMessage("E20332", ""));
        //            // ファイル出力ボタンを非活性化
        //            btnFileExport.Enabled = false;
        //        }
        //        else
        //        {
        //            set_error("", 0);
        //            // ファイル出力ボタンを活性化
        //            btnFileExport.Enabled = true;
        //        }
        //    }
        //}
        //else
        //{
        //    set_error("", 0);
        //    // ファイル出力ボタンを非活性化
        //    btnFileExport.Enabled = false;
        //}
        //}

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

        private void comboBox_Taisho_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
            getFileName();
        }

        private void comboBox_Chohyo_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
            set_error("", 0);
            if (ShukeiVer == 2)
            {
                if (checkBox_Zenhinmoku.Checked)
                {
                    // E20XXX:全品目一括集計表で集計表Ver2は選択できません。
                    set_error("", 0);
                    set_error(GlobalMethod.GetMessage("E20XXX", ""));
                    // ファイル出力ボタンを非活性化
                    btnFileExport.Enabled = false;
                }
                else
                {
                    getFileName();
                }
            }
            else
            {
                getFileName();
            }
        }

        private void src_Busho_SelectedIndexChanged(object sender, EventArgs e)
        {
            get_data();
        }

        private void hinmokuListSelect()
        {
            string convStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(convStr))
            {
                try
                {
                    var cmd = conn.CreateCommand();
                    //  1:主のデータを取得
                    if (comboBox_Taisho.Text != null && comboBox_Taisho.SelectedValue.ToString() == "1")
                    {
                        for (int r = 0; r < kojinList.Count; r++)
                        {
                            cmd.CommandText = "SELECT distinct " +
                                    "mc.ChousainMei " +
                                    ",mg.MadoguchiGroupMei " +
                                    ",ch.ChousaBunkatsuHouhou " +
                                    ",ch.ChousaMadoguchiGroupMasterID " +
                                    "FROM ChousaHinmoku ch " +
                                    "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ch.HinmokuChousainCD = mjmc.MadoguchiL1ChousaTantoushaCD " +
                                    "LEFT JOIN Mst_Chousain mc ON ch.HinmokuChousainCD = mc.KojinCD " +
                                    "LEFT JOIN MadoguchiGroupMaster mg ON ch.ChousaMadoguchiGroupMasterID = mg.MadoguchiGroupMasterID " +
                                    "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND ch.HinmokuChousainCD = '" + kojinList[r] + "' " +
                                    "AND ch.HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                                    "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                                    "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                                    "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                                    "AND mjmc.MadoguchiL1ChousaShinchoku != 80 " +
                                    "ORDER BY mg.MadoguchiGroupMei, ch.ChousaBunkatsuHouhou";
                            var sdb = new SqlDataAdapter(cmd);
                            DataTable dt0 = new DataTable();
                            sdb.Fill(dt0);

                            if (dt0 != null && dt0.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt0.Rows.Count; i++)
                                {
                                    BushoList.Add(src_Busho.Text);
                                    ChousainMeiList.Add(dt0.Rows[i][0].ToString());
                                    TokuchoList.Add(TokuhoBangou.ToString() + "-" + TokuhoBangouEda.ToString());
                                    KojincdList.Add(kojinList[r].ToString());
                                    GroupMeiList.Add(dt0.Rows[i][1].ToString());
                                    BunkatsuList.Add(dt0.Rows[i][2].ToString());
                                    GroupIDList.Add(dt0.Rows[i][3].ToString());
                                }
                            }
                        }
                    }
                    // 2:副1と副2のデータを取得
                    if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "2"))
                    {
                        for (int r = 0; r < kojinList.Count; r++)
                        {
                            cmd.CommandText = "SELECT distinct " +
                                "mc.ChousainMei " +
                                ",mg.MadoguchiGroupMei " +
                                ",ch.ChousaBunkatsuHouhou " +
                                ",ch.ChousaMadoguchiGroupMasterID " +
                                "FROM ChousaHinmoku ch " +
                                "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ((ch.HinmokuFukuChousainCD1 = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                                "OR (ch.HinmokuFukuChousainCD2 = mjmc.MadoguchiL1ChousaTantoushaCD)) " +
                                "LEFT JOIN Mst_Chousain mc ON (ch.HinmokuFukuChousainCD1 = mc.KojinCD) OR (ch.HinmokuFukuChousainCD2 = mc.KojinCD)" +
                                "LEFT JOIN MadoguchiGroupMaster mg ON ch.ChousaMadoguchiGroupMasterID = mg.MadoguchiGroupMasterID " +
                                "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND (( ch.HinmokuFukuChousainCD1 = '" + kojinList[r] + "' ) " +
                                "OR (ch.HinmokuFukuChousainCD2 = '" + kojinList[r] + "' )) " +
                                "AND ch.HinmokuRyakuBushoFuku1CD = '" + src_Busho.SelectedValue.ToString() + "' " +
                                "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                                "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                                "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                                "AND mjmc.MadoguchiL1ChousaShinchoku != 80 " +
                                "ORDER BY mg.MadoguchiGroupMei, ch.ChousaBunkatsuHouhou";
                            var sdb = new SqlDataAdapter(cmd);
                            DataTable dt0 = new DataTable();
                            sdb.Fill(dt0);

                            if (dt0 != null && dt0.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt0.Rows.Count; i++)
                                {
                                    BushoList.Add(src_Busho.Text);
                                    ChousainMeiList.Add(dt0.Rows[i][0].ToString());
                                    TokuchoList.Add(TokuhoBangou.ToString() + "-" + TokuhoBangouEda.ToString());
                                    KojincdList.Add(kojinList[r].ToString());
                                    GroupMeiList.Add(dt0.Rows[i][1].ToString());
                                    BunkatsuList.Add(dt0.Rows[i][2].ToString());
                                    GroupIDList.Add(dt0.Rows[i][3].ToString());
                                }
                            }
                        }
                    }
                    // 0:主+副
                    if (comboBox_Taisho.Text != null && (comboBox_Taisho.SelectedValue.ToString() == "0"))
                    {
                        for (int r = 0; r < kojinList.Count; r++)
                        {
                            cmd.CommandText = "SELECT distinct " +
                                "mc.ChousainMei " +
                                ",mg.MadoguchiGroupMei " +
                                ",ch.ChousaBunkatsuHouhou " +
                                ",ch.ChousaMadoguchiGroupMasterID " +
                                "FROM ChousaHinmoku ch " +
                                "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ((ch.HinmokuChousainCD = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                                "OR (ch.HinmokuFukuChousainCD1 = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                                "OR (ch.HinmokuFukuChousainCD2 = mjmc.MadoguchiL1ChousaTantoushaCD)) " +
                                "LEFT JOIN Mst_Chousain mc ON ((ch.HinmokuChousainCD = mc.KojinCD) OR (ch.HinmokuFukuChousainCD1 = mc.KojinCD) OR (ch.HinmokuFukuChousainCD2 = mc.KojinCD)) " +
                                "LEFT JOIN MadoguchiGroupMaster mg ON ch.ChousaMadoguchiGroupMasterID = mg.MadoguchiGroupMasterID " +
                                "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND (( ch.HinmokuChousainCD = '" + kojinList[r] + "' ) " +
                                "OR (ch.HinmokuFukuChousainCD1 = '" + kojinList[r] + "' ) " +
                                "OR (ch.HinmokuFukuChousainCD2 = '" + kojinList[r] + "' )) " +
                                "AND ch.HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                                "AND ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                                "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                                "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                                "AND mjmc.MadoguchiL1ChousaShinchoku != 80 " +
                                "ORDER BY mg.MadoguchiGroupMei, ch.ChousaBunkatsuHouhou";
                            var sdb = new SqlDataAdapter(cmd);
                            DataTable dt0 = new DataTable();
                            sdb.Fill(dt0);

                            if (dt0 != null && dt0.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt0.Rows.Count; i++)
                                {
                                    BushoList.Add(src_Busho.Text);
                                    ChousainMeiList.Add(dt0.Rows[i][0].ToString());
                                    TokuchoList.Add(TokuhoBangou.ToString() + "-" + TokuhoBangouEda.ToString());
                                    KojincdList.Add(kojinList[r].ToString());
                                    GroupMeiList.Add(dt0.Rows[i][1].ToString());
                                    BunkatsuList.Add(dt0.Rows[i][2].ToString());
                                    GroupIDList.Add(dt0.Rows[i][3].ToString());
                                }
                            }
                        }
                    }
                    conn.Close();
                }
                catch (Exception)
                {
                    // エラーが発生しました
                }
            }
        }

            private bool fileErrorCheck(string chkChousain)
        {
            // 同一担当者のうちでVer1、Ver2混在をチェック→メッセージ出力
            // 対象の担当者集計表Verリスト
            ShukeiVerList.Clear();
            int checkflg = 1;
            string convStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(convStr))
            {
                try
                {
                    var cmd = conn.CreateCommand();
                    // 選択した集計表Ver以外を検索
                    cmd.CommandText = "SELECT " +
                            "ch.ChousaShuukeihyouVer " +
                            "FROM ChousaHinmoku ch " +
                            "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ((ch.HinmokuChousainCD = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                            "OR (ch.HinmokuFukuChousainCD1 = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                            "OR (ch.HinmokuFukuChousainCD2 = mjmc.MadoguchiL1ChousaTantoushaCD)) " +
                            "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND (( ch.HinmokuChousainCD = '" + chkChousain + "' ) " +
                            "OR (ch.HinmokuFukuChousainCD1 = '" + chkChousain + "' ) " +
                            "OR (ch.HinmokuFukuChousainCD2 = '" + chkChousain + "' )) " +
                            "AND ch.HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                            "AND NOT ch.ChousaShuukeihyouVer = " + ShukeiVer + " " +
                            "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                            "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                            "AND mjmc.MadoguchiL1ChousaShinchoku != 80 ";

                    var sda = new SqlDataAdapter(cmd);
                    DataTable dt0 = new DataTable();
                    sda.Fill(dt0);
                    if (dt0 != null && dt0.Rows.Count > 0)
                    {
                        // Verの混在チェック
                        // E20XXX:集計表Ver1、Ver2が混在しています。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20XXX", ""));
                        checkflg = 0;
                    }
                    else
                    {
                        checkflg = 1;
                    }
                    conn.Close();
                }
                catch (Exception)
                {
                    // エラーが発生しました
                    checkflg = 0;
                }
            }
            if (checkflg == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool bunkatsuCheck(string chkChousain, int bnkt)
        {
            // 集計表VerがVer2の場合、調査員単位で分割方法がファイル分割・シート分割混在でメッセージ出力
            int checkflg = 1;
            string convStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(convStr))
            {
                try
                {
                    var cmd = conn.CreateCommand();
                    // 対象ファイルの分割方法以外を検索
                    cmd.CommandText = "SELECT " +
                            "ch.ChousaBunkatsuHouhou " +
                            "FROM ChousaHinmoku ch " +
                            "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou mjmc ON ch.MadoguchiID = mjmc.MadoguchiID AND ((ch.HinmokuChousainCD = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                            "OR (ch.HinmokuFukuChousainCD1 = mjmc.MadoguchiL1ChousaTantoushaCD) " +
                            "OR (ch.HinmokuFukuChousainCD2 = mjmc.MadoguchiL1ChousaTantoushaCD)) " +
                            "WHERE ch.MadoguchiID = '" + MadoguchiID + "' AND (( ch.HinmokuChousainCD = '" + chkChousain + "' ) " +
                            "OR (ch.HinmokuFukuChousainCD1 = '" + chkChousain + "' ) " +
                            "OR (ch.HinmokuFukuChousainCD2 = '" + chkChousain + "' )) " +
                            "AND ch.HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue.ToString() + "' " +
                            "AND NOT ch.ChousaBunkatsuHouhou = " + bnkt + " " +
                            "AND mjmc.MadoguchiL1UketsukeBangou = '" + TokuhoBangou.ToString() + "' " +
                            "AND mjmc.MadoguchiL1UketsukeBangouEdaban = '" + TokuhoBangouEda.ToString() + "' " +
                            "AND mjmc.MadoguchiL1ChousaShinchoku != 80 ";

                    var sda = new SqlDataAdapter(cmd);
                    DataTable dt0 = new DataTable();
                    sda.Fill(dt0);
                    if (dt0 != null && dt0.Rows.Count > 0)
                    {
                        // Verの混在チェック
                        // E20XXX:分割方法が混在しています。
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E20XXX", ""));
                        checkflg = 0;
                    }
                    else
                    {
                        checkflg = 1;
                    }
                    conn.Close();
                }
                catch (Exception)
                {
                    // エラーが発生しました
                    checkflg = 0;
                }
            }
            if (checkflg == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool createFolder(int c)
        {
            // 集計表フォルダ作成
            DirectoryInfo di = new DirectoryInfo(item1_ShukeiFolder.Text + @"\" + ChousainMeiList[c].ToString() + "-" + TokuhoBangou.ToString());
            chousainShukeiFolder = item1_ShukeiFolder.Text + @"\" + ChousainMeiList[c].ToString() + "-" + TokuhoBangou.ToString();
            if (!Directory.Exists(chousainShukeiFolder))
            {
                try
                {
                    di.Create();
                }
                catch (Exception)
                {
                    // フォルダを作成する権限がありません。
                    set_error(GlobalMethod.GetMessage("E70046", "(基本情報等一覧)"));
                    return false;
                }
            }

            // 作業フォルダ作成&DB登録
            string basePath;
            DataTable BaseList = GlobalMethod.getData("CommonMasterID", "CommonValue1", "M_CommonMaster", "CommonMasterKye = 'ENTORY_SAGYOU_HOLDERBASE' ");
            if (BaseList != null && BaseList.Rows.Count > 0)
            {
                basePath = BaseList.Rows[0][0].ToString();
                basePath = basePath.Replace(@"$NENDO$", FromNendo.ToString()).Replace(@"$BUSHO$", BushoList[c].ToString()).Replace(@"$TANTOUSHA$", ChousainMeiList[c].ToString()).Replace(@"$TOKUCHOBANGOU$", TokuhoBangou.ToString());
                //作業フォルダ作成
                DirectoryInfo ds = new DirectoryInfo(basePath);
                if (!Directory.Exists(basePath))
                {
                    try
                    {
                        ds.Create();
                    }
                    catch (Exception)
                    {
                        // フォルダを作成する権限がありません。
                        set_error(GlobalMethod.GetMessage("E70046", "(基本情報等一覧)"));
                        return false;
                    }
                }
                basePath = basePath.Replace("/", @"\");
                // 作業フォルダDB登録
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {
                        cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                        "MadoguchiL1SagyouHolder = '" + basePath + "' " +
                        " WHERE MadoguchiL1ChousaShinchoku != 80 " +
                        " AND MadoguchiID = '" + MadoguchiID + "' " +
                        " AND MadoguchiL1ChousaBushoCD = '" + Busho + "' " +
                        " AND MadoguchiL1ChousaTantoushaCD = '" + KojincdList[c].ToString() + "' " +
                        " AND MadoguchiL1TokuchoBangou = '" + TokuchoList[c].ToString() + "' ";
                        cmd.ExecuteNonQuery();

                        transaction.Commit();
                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        throw;
                    }
                    conn.Close();
                }
                DataTable CreateList = GlobalMethod.getData("CommonMasterID", "CommonValue1", "M_CommonMaster", "CommonMasterKye = 'ENTORY_SAGYOU_HOLDER' ORDER BY CommonMasterID ");
                if (CreateList != null && CreateList.Rows.Count > 0)
                {
                    for (int i = 0; i < CreateList.Rows.Count; i++)
                    {
                        DirectoryInfo dm = new DirectoryInfo(basePath + @"\" + CreateList.Rows[i][0].ToString());
                        if (!Directory.Exists(basePath + @"\" + CreateList.Rows[i][0].ToString()))
                        {
                            try
                            {
                                dm.Create();
                            }
                            catch (Exception)
                            {
                                // フォルダを作成する権限がありません。
                                set_error(GlobalMethod.GetMessage("E70046", "(基本情報等一覧)"));
                                return false;
                            }
                        }
                    }
                    return true;
                }
                else
                {
                    // ENTORY_SAGYOU_HOLDERが取得できない
                    return false;
                }
            }
            else
            {
                // ENTORY_SAGYOU_HOLDERBASEが取得できない
                return false;
            }
        }
    }
}
