using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TokuchoBugyoK2.TokuchoBugyoKDataSetTableAdapters;
using System.Text.RegularExpressions;
using TokuchoBugyoK2;
using System.Collections.Specialized;

namespace TokuchoBugyoK2
{
    public partial class Tokumei_Input : Form
    {
        private string pgmName = "Tokumei_Input";

        // 右クリックメニュー
        ContextMenuStrip contextMenuStrip1 = new ContextMenuStrip();
        ToolStripMenuItem item0 = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuBusho = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuBushoClear = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuTantousha = new ToolStripMenuItem();       // 調査担当者の右メニュー
        ToolStripMenuItem contextMenuTantoushaBusho = new ToolStripMenuItem();  // 調査担当者の右メニュー

        ToolStripMenuItem contextMenuTantoushaBushoClear = new ToolStripMenuItem();      // 部所クリア
        ToolStripMenuItem contextMenuTantoushaTantoushaClear = new ToolStripMenuItem();  // 担当者クリア
        ToolStripMenuItem contextMenuHoukoku = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuIrai = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuCopy = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuPaste = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuHoukokuClear = new ToolStripMenuItem();      // 報告ランククリア
        ToolStripMenuItem contextMenuIraiClear = new ToolStripMenuItem();  // 依頼ランククリア
        GlobalMethod GlobalMethod = new GlobalMethod();

        public string[] UserInfos;
        private string beforeJutaku = "";
        private string befireTokuchoEda = "";

        private string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
        private DataTable MadoguchiData = new DataTable();
        private DataTable DT_MadoguchiL1Chou = new DataTable();
        private DataTable DT_KyouryokuIraisho = new DataTable();
        private DataTable DT_Ouenuketsuke = new DataTable();
        private DataTable DT_Tanpin = new DataTable();
        private DataTable DT_TanpinRank = new DataTable();
        private DataTable DT_Sekou = new DataTable();
        private DataTable DT_GaroonTsuikaAtesaki = new DataTable();
        private DataTable DT_ChousaHinmoku = new DataTable();
        private string Message = "";
        private int Nendo = 0;
        private String sekouMode = "0";
        private string SekouJoukenID = "";
        private int ChousaHinmokuMode = 0;//調査品目編集モード 0:表示 1:編集
        private Boolean HinmokuInitFlag = true;//調査品目タブ初回表示フラグ
        private Image Img_AddRow;
        private Image Img_AddRowNonactive;
        private Image Img_DeleteRow;
        private Image Img_DeleteRowNonactive;
        private Image Img_Sort;
        private int errorCnt = 0;
        private string ShukeiHyoFolder = "";
        private string chousaLinkFlg = "0"; // 調査品目明細のフォルダリンク先表示フラグ　1=非表示、0=表示
        //奉行エクセル
        private string sagyoForuda = "0"; //調査品目明細の作業フォルダ表示フラグ　1=非表示、0=表示
        // 調査品目明細の削除Key
        private String deleteChousaHinmokuIDs = "";
        private String MadoguchiHoukokuzumi = "";
        private String chousaHinmokuSearchWhere = "";
        // 11桁の0
        private String zeroStr = "00000000000";
        ComboBox comboBox1 = new ComboBox();
        private string folderIcon = "0"; // 集計表フォルダアイコン 0:グレー 1:イエロー
        int HinmokuRow = 0;    // 選択範囲の上端行番号
        int HinmokuRowSel = 0; // 選択範囲の下端行番号
        int HinmokuCol = 0;    // 選択範囲の上端列番号
        int HinmokuColSel = 0; // 選択範囲の下端列番号
        int BushoTantouRow = 0; // 部所、担当者の選択範囲
        int BushoTantouColumn = 0; // 部所、担当者の選択範囲
        List<List<string>> copyData = null;
        private string tabChousahinmokuFlg = "0"; // 0:調査品目明細を開いたことが無い 1:調査品目明細を開いたことがある
        private string tabChangeFlg = "0"; // 0:タブ移動してない 1:タブ移動した
        // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
        private string globalErrorFlg = "0";//0:正常、1:エラー

        public string mode = "";
        public string MadoguchiID = "";
        public string KakoIraiID = "";

        //不具合No1207
        //共通マスタからグリッド行高の設定を取得する
        private string AutoSizeGridRowMode;
        private const string GRID_ROW_AUTO_SIZE = "行高自動調整";
        private const string GRID_ROW_FIX_SIZE = "行高自動調整解除";

        // 奉行エクセル移管対応
        private string IsPopup_ShukeiHyou_New = "0";
        public Tokumei_Input()
        {
            InitializeComponent();

            // コンボボックスにマウスホイールイベントを付与
            this.src_Busho.MouseWheel += item_MouseWheel;
            this.src_ShuFuku.MouseWheel += item_MouseWheel;
            this.src_Zaikou.MouseWheel += item_MouseWheel;
            this.src_TantoushaKuuhaku.MouseWheel += item_MouseWheel;
            this.item_Hyoujikensuu.MouseWheel += item_MouseWheel;

            this.c1FlexGrid4.MouseWheel += c1FlexGrid4_MouseWheel; // 調査品目明細のGrid

            //エントリ君修正STEP2
            this.ErrorMessage.Font = new System.Drawing.Font(this.ErrorMessage.Font.Name, float.Parse(GlobalMethod.GetCommonValue1("DSP_ERROR_FONTSIZE")));
        }

        private void Tokumei_Input_Load(object sender, EventArgs e)
        {
            //不具合No1017（751）
            //タブの文字装飾変更対応
            //文字表示を大きくする場合は、デザイナでTabのItemSize.widthを変更する。窓口、特命課長、自分大臣は、125で設定すると、14ポイントぐらいのサイズでいける
            tab.DrawMode = TabDrawMode.OwnerDrawFixed;



            //不具合No1207
            //共通マスタからグリッド行高の設定を取得する
            AutoSizeGridRowMode = GlobalMethod.GetCommonValue1("CHOUSA_GYOU_FLG");

            //不具合No1355（1123）
            lblVersion.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                lblBootMode.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            //ユーザ名を設定
            label7.Text = UserInfos[3] + "：" + UserInfos[1];
            //MadoguchiID = "300027";
            //mode = "update";
            c1FlexGrid1.Rows[0].Height = 40;
            c1FlexGrid1.Height = 44 + (c1FlexGrid1.Rows.Count - 1) * 22;

            // 一覧を隠す
            this.Owner.Hide();
            //gridSizeChange();

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid4.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid4.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid5.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid5.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            BikoGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            BikoGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            item3_TargetPage.ImeMode = ImeMode.Disable;

            // c1FlexGirdを隠す
            // 担当部所
            c1FlexGrid1.Visible = false;
            c1FlexGrid5.Visible = false;
            // 調査品目明細
            c1FlexGrid4.Visible = false;
            // 備考
            BikoGrid.Visible = false;

            //担当部所のボタン　非活性
            button2_Update.Enabled = false;
            button2_Update.BackColor = Color.DarkGray;

            //調査品目一覧の設定
            c1FlexGrid4.Rows[0].AllowMerging = true;
            //ヘッダー2行表示(同一データはセルがマージされる)
            for (int i = 0; i < c1FlexGrid4.Cols.Count; i++)
            {
                c1FlexGrid4.Rows[1][i] = c1FlexGrid4.Rows[0][i];
            }
            //ヘッダー行高さを調整
            c1FlexGrid4.Rows[0].Height = 22;
            c1FlexGrid4.Rows[1].Height = 22;
            //ヘッダー2段目の表示内容
            //c1FlexGrid4.Rows[1][42] = "部所";
            //c1FlexGrid4.Rows[1][43] = "担当者";
            //c1FlexGrid4.Rows[1][44] = "部所";
            //c1FlexGrid4.Rows[1][45] = "担当者";
            //c1FlexGrid4.Rows[1][46] = "部所";
            //c1FlexGrid4.Rows[1][47] = "担当者";
            //c1FlexGrid4.Rows[1][48] = "本数";
            //c1FlexGrid4.Rows[1][49] = "ランク";
            //c1FlexGrid4.Rows[1][50] = "本数";
            //c1FlexGrid4.Rows[1][51] = "ランク";
            c1FlexGrid4.Rows[0]["RowChange"] = " ";
            c1FlexGrid4.Rows[1]["RowChange"] = " ";
            c1FlexGrid4.Rows[1]["HinmokuRyakuBushoCD"] = "部所";
            c1FlexGrid4.Rows[1]["HinmokuChousainCD"] = "担当者";
            c1FlexGrid4.Rows[1]["HinmokuRyakuBushoFuku1CD"] = "部所";
            c1FlexGrid4.Rows[1]["HinmokuFukuChousainCD1"] = "担当者";
            c1FlexGrid4.Rows[1]["HinmokuRyakuBushoFuku2CD"] = "部所";
            c1FlexGrid4.Rows[1]["HinmokuFukuChousainCD2"] = "担当者";
            c1FlexGrid4.Rows[1]["ChousaHoukokuHonsuu"] = "本数";
            c1FlexGrid4.Rows[1]["ChousaHoukokuRank"] = "ランク";
            c1FlexGrid4.Rows[1]["ChousaIraiHonsuu"] = "本数";
            c1FlexGrid4.Rows[1]["ChousaIraiRank"] = "ランク";
            //セル表示画像の取得
            Img_AddRow = Image.FromFile("Resource/Image/new.png");
            Img_AddRowNonactive = Image.FromFile("Resource/Image/UndeleteRow.gif");
            Img_DeleteRow = Image.FromFile("Resource/Image/ActionDelete.png");
            Img_DeleteRowNonactive = Image.FromFile("Resource/Image/DeleteRow.gif");
            Img_Sort = Image.FromFile("Resource/Image/SortIconDefalt.png");
            Img_Sort = new Bitmap(Img_Sort, Img_Sort.Width / 6, Img_Sort.Height / 6);

            // リンク先パスの幅を調整
            //c1FlexGrid4.Cols[41].Width = 0;
            c1FlexGrid4.Cols["ChousaLinkSakliFolder"].Width = 0;

            if (Message != "")
            {
                // 画面呼びなおし時にメッセージを表示
                set_error(Message);
                // メッセージをクリア
                Message = "";
            }

            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            DataTable dt0 = new DataTable();
            //分類
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT " +
                        "Mst_Busho.GyoumuBushoCD AS Value" +
                        ",Mst_Busho.ShibuMei + Mst_Busho.KaMei AS Discript " +
                        "FROM Mst_Busho WHERE JutakubuBushoHyoujiFlg = 1";
                var sda = new SqlDataAdapter(cmd);
                dt0.Clear();
                sda.Fill(dt0);
                conn.Close();
            }
            item0.Text = "分類";
            Set_ContextMenu(item0, dt0);

            DataTable dt = new DataTable();
            //受託課所支部
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();

                //データ取得時に年度がいない場合、当年度とする
                int Nendo;
                int ToNendo;
                if (item1_TourokuNendo.Text == "")
                {
                    Nendo = DateTime.Today.Year;
                    ToNendo = DateTime.Today.AddYears(1).Year;
                }
                else
                {
                    int.TryParse(item1_TourokuNendo.Text.ToString(), out Nendo);
                    ToNendo = Nendo + 1;
                }
                //cmd.CommandText = "SELECT " +
                //"GyoumuBushoCD  " +
                //",BushokanriboKameiRaku  " +
                //"FROM Mst_Busho  " +
                //"WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                //"ORDER BY BushoMadoguchiNarabijun";

                cmd.CommandText = "SELECT " +
                "GyoumuBushoCD  " +
                ",BushokanriboKameiRaku  " +
                "FROM Mst_Busho  " +
                "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  ";
                //// 今日日付の年度データを検索する際は、今日有効な部所を表示
                //if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                //{
                //    cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today + "' ) ";
                //}
                //else
                //{
                //    cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                //}
                //cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                ////"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                cmd.CommandText += "ORDER BY BushoMadoguchiNarabijun";

                var sda = new SqlDataAdapter(cmd);
                dt.Clear();
                sda.Fill(dt);
                conn.Close();
            }
            contextMenuBusho.Text = "部所";
            Set_ContextMenu(contextMenuBusho, dt);

            contextMenuBushoClear.Text = "部所クリア";
            contextMenuBushoClear.Click += ContextMenuBushoClearEvent;

            contextMenuTantoushaBushoClear.Text = "部所クリア";
            contextMenuTantoushaBushoClear.Click += ContextMenuTantoushaBushoClearEvent;

            contextMenuTantoushaTantoushaClear.Text = "担当者クリア";
            contextMenuTantoushaTantoushaClear.Click += ContextMenuTantoushaClearEvent;

            contextMenuCopy.Text = "コピー";
            contextMenuCopy.Click += ContextMenuCopyEvent;

            contextMenuPaste.Text = "貼り付け";
            contextMenuPaste.Click += ContextMenuPasteEvent;

            contextMenuHoukokuClear.Text = "報告ランククリア";
            contextMenuHoukokuClear.Click += ContextMenuHoukokuClearEvent;

            contextMenuIraiClear.Text = "依頼ランククリア";
            contextMenuIraiClear.Click += ContextMenuIraiClearEvent;

            //if (mode != "insert")
            //{
            //    // 単品入力から業務CDの取得
            //    int TankaKeiyakuID = 0;
            //    DataTable Tanpin_Dt = new DataTable();
            //    using (var conn = new SqlConnection(connStr))
            //    {
            //        var cmd = conn.CreateCommand();
            //        cmd.CommandText = "SELECT TanpinGyoumuCD FROM TanpinNyuuryoku"
            //                    + " WHERE MadoguchiID = " + MadoguchiID
            //                    ;

            //        var sda = new SqlDataAdapter(cmd);
            //        sda.Fill(Tanpin_Dt);
            //        if (Tanpin_Dt.Rows.Count > 0 && Tanpin_Dt.Rows[0][0] != null)
            //        {
            //            TankaKeiyakuID = int.Parse(Tanpin_Dt.Rows[0][0].ToString());
            //        }
            //    }

            //    // 報告ランク
            //    DataTable houkokuDt = new DataTable();
            //    using (var conn = new SqlConnection(connStr))
            //    {
            //        var cmd = conn.CreateCommand();
            //        cmd.CommandText = "SELECT"
            //                        + " TankaRankHinmoku AS Value"
            //                        + ", TankaRankHinmoku AS Descript"
            //                        + " FROM TankaKeiyakuRank"
            //                        + " WHERE TankaRankDeleteFlag != 1"
            //                        + " AND TankaKeiyakuID = " + TankaKeiyakuID
            //                        + " ORDER BY TankaKeiyakuID, TankaRankID"
            //                        ;

            //        var sda = new SqlDataAdapter(cmd);
            //        houkokuDt.Clear();
            //        sda.Fill(houkokuDt);
            //        conn.Close();
            //    }

            //    contextMenuHoukoku.Text = "報告ランク";
            //    //contextMenuHoukoku.DropDownItems.Add("", null, ContextMenuEvent);
            //    //contextMenuHoukoku.DropDownItems.Add("A-①", null, ContextMenuEvent);
            //    //contextMenuHoukoku.DropDownItems.Add("A-②", null, ContextMenuEvent);

            //    contextMenuHoukoku = Set_ContextMenu(contextMenuHoukoku, houkokuDt);

            //    // 依頼ランク
            //    DataTable iraiDt = new DataTable();
            //    using (var conn = new SqlConnection(connStr))
            //    {
            //        var cmd = conn.CreateCommand();
            //        cmd.CommandText = "SELECT"
            //                        + " TankaRankHinmoku AS Value"
            //                        + ", TankaRankHinmoku AS Descript"
            //                        + " FROM TankaKeiyakuRank"
            //                        + " WHERE TankaRankDeleteFlag != 1"
            //                        + " AND TankaKeiyakuID = " + TankaKeiyakuID
            //                        + " ORDER BY TankaKeiyakuID, TankaRankID"
            //                        ;

            //        var sda = new SqlDataAdapter(cmd);
            //        iraiDt.Clear();
            //        sda.Fill(iraiDt);
            //        conn.Close();
            //    }

            //    contextMenuIrai.Text = "依頼ランク";
            //    contextMenuIrai = Set_ContextMenu(contextMenuIrai, iraiDt);
            //}

            //コンボボックス取得
            get_combo();

            //新規登録以外だったらデータ取得
            if (MadoguchiID != "")
            {
                get_data(1);
                get_combo_byNendo();
                //get_data(2);
                //get_data(3); //調査品目　検索条件のコンボが正常に認識できないため、別タイミングに移動
                //get_data(4);
                //get_data(5); // 応援受付
                //get_data(6);
                //get_data(7);

                // ヘッダーのボタン下のGaroon連携設定日時の文言取得（変更出来るように）
                string garoonUpdateDisp = GlobalMethod.GetCommonValue1("GAROON_UPDATETIME_DISP");
                if (garoonUpdateDisp != null && garoonUpdateDisp == "")
                {
                    item1_GaroonUpdateDispTitle.Text = garoonUpdateDisp;
                }

                // GaroonTsuikaAtesakiから更新日時を取得する
                DataTable dt3 = new DataTable();
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT " +
                        "MadoguchiGaroonRenkeiJikouDate " +
                        "FROM MadoguchiJouhou " +
                        "WHERE MadoguchiID = '" + MadoguchiID + "' ";
                    var sda = new SqlDataAdapter(cmd);
                    dt3.Clear();
                    sda.Fill(dt3);
                    conn.Close();

                    if (dt3 != null && dt3.Rows.Count > 0)
                    {
                        item1_GaroonUpdateDisp.Text = dt3.Rows[0][0].ToString();
                    }
                }
            }
            FolderPathCheck();

            // ヘッダー表示
            // 特調番号
            Header1.Text = item1_MadoguchiUketsukeBangou.Text + "-" + item1_MadoguchiUketsukeBangouEdaban.Text;
            // 発注者名・課名
            Header3.Text = item1_MadoguchiHachuuKikanmei.Text;
            // 業務名称
            Header4.Text = item1_MadoguchiGyoumuMeishou.Text;
        }

        //public ToolStripMenuItem Set_ContextMenu(ToolStripMenuItem item, DataTable dt)
        //{
        //    for (int i = 0; i < dt.Rows.Count; i++)
        //    {
        //        if (dt.Rows[i][1].ToString() != "")
        //        {
        //            item.DropDownItems.Add(dt.Rows[i][1].ToString(), null, ContextMenuEvent);
        //        }
        //    }
        //    return item;
        //}
        //No.1443
        public ToolStripMenuItem Set_ContextMenu(ToolStripMenuItem item, DataTable dt, bool isEscape = false)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][1].ToString() != "")
                {
                    // No.1443 対応
                    //item.DropDownItems.Add(dt.Rows[i][1].ToString(), null, ContextMenuEvent);
                    if (isEscape)
                    {
                        if (dt.Rows[i][1].ToString() == "-")
                        {
                            item.DropDownItems.Add("半角ハイフン ", null, ContextMenuEvent);
                            //item.DropDownItems.Add("[" + dt.Rows[i][1].ToString() + "]", null, ContextMenuEvent);
                        }
                        else
                        {
                            item.DropDownItems.Add(dt.Rows[i][1].ToString(), null, ContextMenuEvent);
                        }
                    }
                    else
                    {
                        item.DropDownItems.Add(dt.Rows[i][1].ToString(), null, ContextMenuEvent);
                    }
                }
            }
            return item;
        }

        public ToolStripMenuItem Set_ContextBushoMenu(ToolStripMenuItem item, DataTable dt)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][1].ToString() != "")
                {
                    item.DropDownItems.Add(dt.Rows[i][1].ToString(), null, ContextMenuBushoEvent);
                }
            }
            return item;
        }

        public void ContextMenuEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(BushoTantouRow, BushoTantouColumn);
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col] = sender.ToString();
        }

        // 担当者右クリックで部所
        public void ContextMenuBushoEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(BushoTantouRow, BushoTantouColumn);
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col - 1] = sender.ToString();
        }

        // 部所 部所クリア
        public void ContextMenuBushoClearEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(BushoTantouRow, BushoTantouColumn);
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col] = "";
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col + 1] = "";
        }
        // 担当者 部所クリア
        public void ContextMenuTantoushaBushoClearEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(BushoTantouRow, BushoTantouColumn);
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col - 1] = "";
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col] = "";
        }
        // 担当者 担当者クリア
        public void ContextMenuTantoushaClearEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(BushoTantouRow, BushoTantouColumn);
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col] = "";
        }

        // コピー
        public void ContextMenuCopyEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(HinmokuRow, HinmokuCol, HinmokuRowSel, HinmokuColSel);
            SendKeys.SendWait("^(c)");
        }
        // 貼り付け
        public void ContextMenuPasteEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(HinmokuRow, HinmokuCol, HinmokuRowSel, HinmokuColSel);
            SendKeys.SendWait("^(v)");
        }

        // 報告ランククリア
        public void ContextMenuHoukokuClearEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(BushoTantouRow, BushoTantouColumn);
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col] = "";
        }
        // 依頼ランククリア
        public void ContextMenuIraiClearEvent(object sender, EventArgs e)
        {
            c1FlexGrid4.Select(BushoTantouRow, BushoTantouColumn);
            c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col] = "";
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
        //タブ遷移時
        private void tab_SelectedIndexChanged(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();

            // c1FlexGirdを隠す
            // 担当部所
            c1FlexGrid1.Visible = false;
            c1FlexGrid5.Visible = false;
            // 調査品目明細
            c1FlexGrid4.Visible = false;
            // 備考
            BikoGrid.Visible = false;

            if (((TabControl)sender).SelectedTab.Text == "担当部所")
            {
                //描画停止
                c1FlexGrid1.BeginUpdate();
                get_data(2);
                c1FlexGrid1.Visible = true;
                c1FlexGrid5.Visible = true;
                //描画再開
                c1FlexGrid1.EndUpdate();
            }
            if (((TabControl)sender).SelectedTab.Text == "調査品目明細")
            {
                //// 1201 最大化する
                //this.WindowState = FormWindowState.Maximized;
                // 表示モードに切り替える
                ChousaHinmokuMode = 0;
                ChousaHinmokuGrid_InputMode();
                btnGridSize.Text = "一覧拡大";
                tabChangeFlg = "1"; // 0:タブ移動してない 1:タブ移動した

                // 単品入力から業務CDの取得
                int TankaKeiyakuID = 0;
                DataTable Tanpin_Dt = new DataTable();
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT TanpinGyoumuCD FROM TanpinNyuuryoku"
                                + " WHERE MadoguchiID = " + MadoguchiID
                                ;

                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(Tanpin_Dt);
                    if (Tanpin_Dt.Rows.Count > 0 && Tanpin_Dt.Rows[0][0] != null)
                    {
                        TankaKeiyakuID = int.Parse(Tanpin_Dt.Rows[0][0].ToString());
                    }
                }

                // 報告ランク
                DataTable houkokuDt = new DataTable();
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT"
                                    + " TankaRankHinmoku AS Value"
                                    + ", TankaRankHinmoku AS Descript"
                                    + " FROM TankaKeiyakuRank"
                                    + " WHERE TankaRankDeleteFlag != 1"
                                    + " AND TankaKeiyakuID = " + TankaKeiyakuID
                                    + " ORDER BY TankaKeiyakuID, TankaRankID"
                                    ;

                    var sda = new SqlDataAdapter(cmd);
                    houkokuDt.Clear();
                    sda.Fill(houkokuDt);
                    conn.Close();
                }
                contextMenuHoukoku = new ToolStripMenuItem();
                contextMenuHoukoku.Text = "報告ランク";
                //No.1443
                contextMenuHoukoku = Set_ContextMenu(contextMenuHoukoku, houkokuDt, true);

                // 依頼ランク
                DataTable iraiDt = new DataTable();
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT"
                                    + " TankaRankHinmoku AS Value"
                                    + ", TankaRankHinmoku AS Descript"
                                    + " FROM TankaKeiyakuRank"
                                    + " WHERE TankaRankDeleteFlag != 1"
                                    + " AND TankaKeiyakuID = " + TankaKeiyakuID
                                    + " ORDER BY TankaKeiyakuID, TankaRankID"
                                    ;

                    var sda = new SqlDataAdapter(cmd);
                    iraiDt.Clear();
                    sda.Fill(iraiDt);
                    conn.Close();
                }
                contextMenuIrai = new ToolStripMenuItem();
                contextMenuIrai.Text = "依頼ランク";
                //No.1443
                contextMenuIrai = Set_ContextMenu(contextMenuIrai, iraiDt, true);
                gridSizeChange();
                tabChousahinmokuFlg = "1"; // 0:調査品目明細を開いたことが無い 1:調査品目明細を開いたことがある
                tabChangeFlg = "0"; // 0:タブ移動してない 1:タブ移動した

                // .00が残る対応
                c1FlexGrid4.EditOptions -= C1.Win.C1FlexGrid.EditFlags.UseNumericEditor;

                //描画停止
                c1FlexGrid4.BeginUpdate();
                if (HinmokuInitFlag)
                {
                    //調査品目タブの初期化
                    ClearHinmoku();
                    HinmokuInitFlag = false;
                }
                //調査品目タブ検索
                get_data(3);

                // 1:報告完了の場合、ボタン制御
                if (MadoguchiHoukokuzumi == "1")
                {
                    // 入力開始・文字変換を非活性
                    // 入力開始
                    button3_InputStatus.BackColor = Color.DimGray;
                    button3_InputStatus.Enabled = false;
                }
                else
                {
                    // 入力開始
                    button3_InputStatus.BackColor = Color.FromArgb(42, 78, 122);
                    button3_InputStatus.Enabled = true;
                }
                c1FlexGrid4.Visible = true;

                //描画再開
                c1FlexGrid4.EndUpdate();
            }

            // 備考選択時 
            TabControl tabcon = (TabControl)sender;
            if (tabcon.SelectedTab.Text.Equals("備考"))
            {
                // ShibuBikou 初期化
                // データを 'tokuchoBugyoK2DataSet.ShibuBikou' テーブルに読み込みます。
                ShibuBikoManager sbm = new ShibuBikoManager();
                sbm.ShibuBikoInit(this, Decimal.Parse(MadoguchiID));
                BikoGrid.Visible = true;
            }

            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void get_combo()
        {

            //受託部所
            //SQL変数
            String discript = "Mst_Busho.ShibuMei + ' ' + ISNULL(Mst_Busho.KaMei,'')";
            String value = "Mst_Busho.GyoumuBushoCD ";
            String table = "Mst_Busho";
            String where = "JutakubuBushoHyoujiFlg = 1 AND GyoumuBushoCD != '999990' AND BushoNewOld <= 1 AND BushoDeleteFlag != 1 " +
                "ORDER BY Seiretsu";
            //コンボボックスデータ取得
            DataTable tmpdt = GlobalMethod.getData(discript, value, table, where);

            SortedList sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //c1FlexGrid1.Cols[2].DataMap = sl;
            //調査品目Gridの担当部所は年度に紐づいて選択肢を作成
            //c1FlexGrid4.Cols[42].DataMap = sl;//調査品目タブGrid担当部所
            //c1FlexGrid4.Cols[44].DataMap = sl;//調査品目タブGrid担当部所
            //c1FlexGrid4.Cols[46].DataMap = sl;//調査品目タブGrid担当部所

            //窓口部所
            //SQL変数
            discript = "BushokanriboKamei";
            value = "GyoumuBushoCD ";
            table = "Mst_Busho ";
            where = "KashoShibuCD != '' AND BushoNewOld <= 1 AND GyoumuBushoCD != '999990' AND BushoDeleteFlag != 1 " +
                "AND BushoMadoguchiHyoujiFlg = 1 AND GyoumuBushoCD != '127900' ORDER BY BushoMadoguchiNarabijun ";
            //コンボボックスデータ取得
            DataTable tmpdt2 = GlobalMethod.getData(discript, value, table, where);
            if (tmpdt2 != null)
            {
                //空白行追加
                DataRow dr = tmpdt2.NewRow();
                tmpdt2.Rows.InsertAt(dr, 0);
            }
            item1_MadoguchiTantoushaBushoCD.DisplayMember = "Discript";
            item1_MadoguchiTantoushaBushoCD.ValueMember = "Value";
            item1_MadoguchiTantoushaBushoCD.DataSource = tmpdt2;
            //ユーザの部所セット
            item1_MadoguchiTantoushaBushoCD.SelectedValue = UserInfos[2];



            //契約区分
            //SQL変数
            discript = "GyoumuKubunHyouji";
            value = "GyoumuNarabijunCD";
            table = "Mst_GyoumuKubun";
            where = "GyoumuNarabijunCD < 100 ";
            //コンボボックスデータ取得
            DataTable tmpdt5 = GlobalMethod.getData(discript, value, table, where);
            if (tmpdt5 != null)
            {
                //空白行追加
                DataRow dr = tmpdt5.NewRow();
                tmpdt5.Rows.InsertAt(dr, 0);
            }
            item1_AnkenGyoumuKubun.DisplayMember = "Discript";
            item1_AnkenGyoumuKubun.ValueMember = "Value";
            item1_AnkenGyoumuKubun.DataSource = tmpdt5;


            //調査種別
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "単品");
            tmpdt.Rows.Add(2, "一般");
            tmpdt.Rows.Add(3, "単契");
            item1_MadoguchiChousaShubetsu.DisplayMember = "Discript";
            item1_MadoguchiChousaShubetsu.ValueMember = "Value";
            item1_MadoguchiChousaShubetsu.DataSource = tmpdt;


            //担当部所タブ 協力部所　進捗状況
            Hashtable imgMap = new Hashtable();
            imgMap.Add("8", Image.FromFile("Resource/Image/shin_ao.png"));     // 報告済み
            imgMap.Add("5", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（二次検証済み）
            //imgMap.Add("6", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（中止）
            imgMap.Add("6", Image.FromFile("Resource/Image/shin_ao.png"));     // 中止
            imgMap.Add("7", Image.FromFile("Resource/Image/shin_midori.png")); // 担当者済み
            imgMap.Add("1", Image.FromFile("Resource/Image/shin_dokuro.png")); // 締切日経過
            imgMap.Add("2", Image.FromFile("Resource/Image/shin_aka.png"));    // 締切日が3日以内、かつ2次検証が完了していない
            imgMap.Add("3", Image.FromFile("Resource/Image/shin_kiiro.png"));  // 締切日が1週間以内、かつ2次検証が完了していない
                                                                               //imgMap.Add("4", Image.FromFile("Resource/Image/blank2.png"));      // 上記のいずれにも該当しない

            c1FlexGrid1.Cols[0].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid1.Cols[0].ImageMap = imgMap;
            c1FlexGrid1.Cols[0].ImageAndText = false;

            //担当部所タブ 協力部所　担当者状況
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(10, "依頼");
            tmpdt.Rows.Add(20, "調査開始");
            tmpdt.Rows.Add(30, "見積中");
            tmpdt.Rows.Add(40, "集計中");
            tmpdt.Rows.Add(50, "担当者済");
            tmpdt.Rows.Add(60, "一次検済");
            tmpdt.Rows.Add(70, "二次検済");
            tmpdt.Rows.Add(80, "中止");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            c1FlexGrid1.Cols[5].DataMap = sl;

            //担当部所　調査員
            discript = "ChousainMei ";
            value = "KojinCD ";
            table = "Mst_Chousain ";
            //where = "RetireFLG = 0 AND TokuchoFLG = 1";
            where = "";
            //コンボボックスデータ取得
            DataTable tmpdt10 = GlobalMethod.getData(discript, value, table, where);
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt10);
            //該当グリッドのセルにセット
            c1FlexGrid1.Cols[3].DataMap = sl;
            c1FlexGrid5.Cols[3].DataMap = sl;//担当部所タブGaroon追加宛先Grid担当者

            //担当部所　調査員
            discript = "ChousainMei ";
            value = "KojinCD ";
            table = "Mst_Chousain ";
            where = "RetireFLG = 0 AND TokuchoFLG = 1";
            //コンボボックスデータ取得
            DataTable tmpdt11 = GlobalMethod.getData(discript, value, table, where);
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt11);
            //c1FlexGrid4.Cols[43].DataMap = sl;//調査品目タブGrid担当者
            //c1FlexGrid4.Cols[45].DataMap = sl;//調査品目タブGrid担当者
            //c1FlexGrid4.Cols[47].DataMap = sl;//調査品目タブGrid担当者
            c1FlexGrid4.Cols["HinmokuChousainCD"].DataMap = sl;//調査品目タブGrid担当者
            c1FlexGrid4.Cols["HinmokuFukuChousainCD1"].DataMap = sl;//調査品目タブGrid担当者
            c1FlexGrid4.Cols["HinmokuFukuChousainCD2"].DataMap = sl;//調査品目タブGrid担当者


            //調査品目　調査主副コンボ
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(0, "主+副");
            tmpdt.Rows.Add(1, "主");
            tmpdt.Rows.Add(2, "副");
            src_ShuFuku.DisplayMember = "Discript";
            src_ShuFuku.ValueMember = "Value";
            src_ShuFuku.DataSource = tmpdt;

            //調査品目　材工
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(0, "全て");
            tmpdt.Rows.Add(1, "材のみ");
            tmpdt.Rows.Add(2, "工のみ");
            tmpdt.Rows.Add(3, "材+D工");
            tmpdt.Rows.Add(4, "E工のみ");
            tmpdt.Rows.Add(5, "他");
            src_Zaikou.DisplayMember = "Discript";
            src_Zaikou.ValueMember = "Value";
            src_Zaikou.DataSource = tmpdt;

            //奉行エクセル
            //グループ名
            discript = "MadoguchiGroupMei ";
            value = "MadoguchiID ";
            table = "MadoguchiGroupMaster ";
            where = "MadoguchiGroupMasterID = " + MadoguchiID; //MadoguchiIDが一致するもの
            //コンボボックスデータ取得
            DataTable tmpdt22 = GlobalMethod.getData(discript, value, table, where);
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt22);
            c1FlexGrid4.Cols["GroupMei"].DataMap = sl;

            //調査品目　担当者空白リスト
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(0, "全て");
            tmpdt.Rows.Add(1, "担当者が空白のリスト");
            tmpdt.Rows.Add(2, "担当者が設定済のリスト");
            src_TantoushaKuuhaku.DisplayMember = "Discript";
            src_TantoushaKuuhaku.ValueMember = "Value";
            src_TantoushaKuuhaku.DataSource = tmpdt;

            //調査品目 進捗アイコン
            imgMap = new Hashtable();
            imgMap.Add("8", Image.FromFile("Resource/Image/shin_ao.png"));
            imgMap.Add("5", Image.FromFile("Resource/Image/greenT1.png"));
            //imgMap.Add("6", Image.FromFile("Resource/Image/greenT1.png"));
            imgMap.Add("6", Image.FromFile("Resource/Image/shin_ao.png"));     // 中止
            imgMap.Add("1", Image.FromFile("Resource/Image/shin_dokuro.png"));
            imgMap.Add("7", Image.FromFile("Resource/Image/shin_midori.png"));
            imgMap.Add("2", Image.FromFile("Resource/Image/shin_aka.png"));
            imgMap.Add("3", Image.FromFile("Resource/Image/shin_kiiro.png"));
            imgMap.Add("4", Image.FromFile("Resource/Image/blank2.png"));
            //c1FlexGrid4.Cols[5].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            //c1FlexGrid4.Cols[5].ImageMap = imgMap;
            //c1FlexGrid4.Cols[5].ImageAndText = false;
            c1FlexGrid4.Cols["ShinchokuIcon"].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid4.Cols["ShinchokuIcon"].ImageMap = imgMap;
            c1FlexGrid4.Cols["ShinchokuIcon"].ImageAndText = false;

            //調査品目　Grid材工
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            //tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "材");
            tmpdt.Rows.Add(2, "D工");
            tmpdt.Rows.Add(3, "E工");
            tmpdt.Rows.Add(4, "他");

            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            //c1FlexGrid4.Cols[8].DataMap = sl;
            c1FlexGrid4.Cols["ChousaZaiKou"].DataMap = sl;

            //奉行エクセル
            //調査品目　Grid集計表Ver
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            //tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "-");
            tmpdt.Rows.Add(2, "集計表Ver2");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            c1FlexGrid4.Cols["ShukeihyoVer"].DataMap = sl;

            //調査品目　Grid分割方法Ver
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(0, "-");
            tmpdt.Rows.Add(1, "シート分割");
            tmpdt.Rows.Add(2, "ファイル分割");
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            c1FlexGrid4.Cols["Bunkatsuhouhou"].DataMap = sl;


            //調査品目　Gridリンク先画像
            imgMap = new Hashtable();
            imgMap.Add(0, Image.FromFile("Resource/Image/folder_gray_s.png"));
            imgMap.Add(1, Image.FromFile("Resource/Image/folder_yellow_s.png"));
            imgMap.Add(2, Image.FromFile("Resource/Image/excel_icon_s.png"));
            //c1FlexGrid4.Cols[40].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            //c1FlexGrid4.Cols[40].ImageMap = imgMap;
            //c1FlexGrid4.Cols[40].ImageAndText = false;
            c1FlexGrid4.Cols["ChousaLinkSakli"].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid4.Cols["ChousaLinkSakli"].ImageMap = imgMap;
            c1FlexGrid4.Cols["ChousaLinkSakli"].ImageAndText = false;

            //調査品目　Grid属性
            discript = "ZokuseiMeishou ";
            value = "ZokuseiMeishou ";
            table = "Mst_Zokusei ";
            where = "ZokuseiID IS NOT NULL ORDER BY ZokuseiNarabijun ";
            tmpdt = new DataTable();
            tmpdt = GlobalMethod.getData(discript, value, table, where);
            if (tmpdt != null)
            {
                //空白行追加
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            sl = new SortedList();
            sl = GlobalMethod.Get_SortedList(tmpdt);
            //該当グリッドのセルにセット
            //c1FlexGrid4.Cols[24].DataMap = sl;
            c1FlexGrid4.Cols["ChousaObiMei"].DataMap = sl;

        }

        private void get_combo_byNendo()
        {
            //データ取得時に年度がいない場合、当年度とする
            int ToNendo;
            if (Nendo == 0)
            {
                Nendo = (int)DateTime.Today.Year;
                ToNendo = DateTime.Today.AddYears(1).Year;
            }
            else
            {
                ToNendo = Nendo + 1;
            }

            //協力部所タブ　部所一覧の更新
            string discript = "BushokanriboKamei ";
            string value = "GyoumuBushoCD ";
            string table = "Mst_Busho ";
            string where = "BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != '' " +
                                //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                                ////" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                                //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                                " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                                " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                                " ORDER BY BushoMadoguchiNarabijun ";
            //コンボボックスデータ取得
            DataTable tmpdt = GlobalMethod.getData(discript, value, table, where);
            for (int i = 0; i < tmpdt.Rows.Count; i++)
            {
                //tableLayoutPanel17.Controls["KyoroykuBusho" + (i + 1)].Text = tmpdt.Rows[i][1].ToString();
                //tableLayoutPanel17.Controls["KyoroykuBusho" + (i + 1)].Visible = true;
                if (i == 0) { KyoroykuBusho1.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho1.Visible = true; }
                if (i == 1) { KyoroykuBusho2.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho2.Visible = true; }
                if (i == 2) { KyoroykuBusho3.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho3.Visible = true; }
                if (i == 3) { KyoroykuBusho4.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho4.Visible = true; }
                if (i == 4) { KyoroykuBusho5.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho5.Visible = true; }
                if (i == 5) { KyoroykuBusho6.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho6.Visible = true; }
                if (i == 6) { KyoroykuBusho7.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho7.Visible = true; }
                if (i == 7) { KyoroykuBusho8.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho8.Visible = true; }
                if (i == 8) { KyoroykuBusho9.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho9.Visible = true; }
                if (i == 9) { KyoroykuBusho10.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho10.Visible = true; }
                if (i == 10) { KyoroykuBusho11.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho11.Visible = true; }
                if (i == 11) { KyoroykuBusho12.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho12.Visible = true; }
                if (i == 12) { KyoroykuBusho13.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho13.Visible = true; }
                if (i == 13) { KyoroykuBusho14.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho14.Visible = true; }
                if (i == 14) { KyoroykuBusho15.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho15.Visible = true; }
                if (i == 15) { KyoroykuBusho16.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho16.Visible = true; }
                if (i == 16) { KyoroykuBusho17.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho17.Visible = true; }
                if (i == 17) { KyoroykuBusho18.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho18.Visible = true; }
                if (i == 18) { KyoroykuBusho19.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho19.Visible = true; }
                if (i == 19) { KyoroykuBusho20.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho20.Visible = true; }
                if (i == 20) { KyoroykuBusho21.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho21.Visible = true; }
                if (i == 21) { KyoroykuBusho22.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho22.Visible = true; }
                if (i == 22) { KyoroykuBusho23.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho23.Visible = true; }
                if (i == 23) { KyoroykuBusho24.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho24.Visible = true; }
                if (i == 24) { KyoroykuBusho25.Text = tmpdt.Rows[i][1].ToString(); KyoroykuBusho25.Visible = true; }
            }
            ////部所一覧のチェックボックスを再セット
            //set_data(2);
            SortedList sl = GlobalMethod.Get_SortedList(tmpdt);
            c1FlexGrid1.Cols[2].DataMap = sl;


            //調査品目　調査担当部所
            discript = "BushokanriboKamei ";
            value = "GyoumuBushoCD ";
            table = "Mst_Busho ";
            where = "BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != '' " +
                                //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                                ////" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                                //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                                " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                                " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                                " ORDER BY BushoMadoguchiNarabijun ";
            //コンボボックスデータ取得
            tmpdt = new DataTable();
            tmpdt = GlobalMethod.getData(discript, value, table, where);
            if (tmpdt != null)
            {
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }
            src_Busho.DisplayMember = "Discript";
            src_Busho.ValueMember = "Value";
            src_Busho.DataSource = tmpdt;
            // Keyで並べたくないので、ListDictionary（詰めた順に表示）を利用する
            System.Collections.Specialized.ListDictionary ld = new System.Collections.Specialized.ListDictionary();
            ld = GlobalMethod.Get_ListDictionary(tmpdt);
            c1FlexGrid5.Cols[2].DataMap = ld;//担当部所タブGaroon追加宛先Grid担当部所

            //調査品目　Grid部所
            discript = "BushokanriboKameiRaku ";
            value = "GyoumuBushoCD ";
            table = "Mst_Busho ";
            where = "BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != '' " +
                                //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                                ////" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                                //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                                " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                                " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                                " ORDER BY BushoMadoguchiNarabijun ";
            //コンボボックスデータ取得
            tmpdt = new DataTable();
            tmpdt = GlobalMethod.getData(discript, value, table, where);

            ld = new ListDictionary();
            ld = GlobalMethod.Get_ListDictionary(tmpdt);
            //sl = new SortedList();
            //sl = GlobalMethod.Get_SortedList(tmpdt);
            //c1FlexGrid4.Cols[42].DataMap = ld;//調査品目タブGrid担当部所
            //c1FlexGrid4.Cols[44].DataMap = ld;//調査品目タブGrid担当部所
            //c1FlexGrid4.Cols[46].DataMap = ld;//調査品目タブGrid担当部所
            c1FlexGrid4.Cols["HinmokuRyakuBushoCD"].DataMap = ld;//調査品目タブGrid担当部所
            c1FlexGrid4.Cols["HinmokuRyakuBushoFuku1CD"].DataMap = ld;//調査品目タブGrid担当部所
            c1FlexGrid4.Cols["HinmokuRyakuBushoFuku2CD"].DataMap = ld;//調査品目タブGrid担当部所
        }
        private void get_data(int tab)
        {

            //レイアウトロジックを停止する
            this.SuspendLayout();
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    //調査概要タブ
                    if (tab == 1)
                    {
                        //窓口情報取得
                        cmd.CommandText = "SELECT " +
                            "MadoguchiJutakuBushoCD " + //受託課所支部 0
                            ",AnkenTantoushaMei " + //契約担当者 
                            ",mb1.BushoShozokuChou " +//受託部所所属長
                            ",MadoguchiTantoushaBushoCD " +//窓口部所
                            ",GyoumuKanrishaMei " +//業務担当者　
                            ",mb2.BushoShozokuChou " +//窓口部所所属長 5 
                            ",ChousainMei " +//窓口担当者名
                            ",MadoguchiTantoushaCD " +//窓口担当者CD
                            ",MadoguchiTourokuNendo " +//登録年度
                            ",KanriGijutsushaNM " +//管理技術者 9
                            ",CASE MadoguchiJutakuBangouEdaban WHEN ''  THEN MadoguchiJutakuBangou ELSE MadoguchiJutakuBangou + '-' + MadoguchiJutakuBangouEdaban END AS JutakuBangou  " +//受託番号 10
                            ",MadoguchiJutakuBangouEdaban " +//受託番号枝番
                            ",GyoumuKanrishaMei " +//業務管理者　　CD
                            ",MadoguchiGyoumuKanrishaCD " +//業務管理者の業務管理者CD
                            ",MadoguchiUketsukeBangou " +//特調番号
                            ",MadoguchiUketsukeBangouEdaban " +//特調番号枝番 15
                            ",MadoguchiHachuuKikanmei " +//発注者名・課名
                            ",MadoguchiKanriBangou " +//管理番号
                            ",MadoguchiGyoumuMeishou " +//業務名称 18
                            ",MadoguchiChousaKubunJibusho " +//調査区分
                            ",MadoguchiChousaKubunShibuShibu " +//調査区分 20
                            ",MadoguchiChousaKubunHonbuShibu " +//調査区分
                            ",MadoguchiChousaKubunShibuHonbu " +//調査区分
                            ",MadoguchiKoujiKenmei " +//工事件名 23
                            ",AnkenGyoumuKubun " +//契約区分 24
                            ",MadoguchiChousaShubetsu " +//調査種別
                            ",MadoguchiChousaHinmoku " +//調査品目
                            ",ISNULL(MadoguchiJiishiKubun,'1') " +//実施区分
                            ",MadoguchiBikou " +//備考 28
                            ",MadoguchiTourokubi " +//登録日
                            ",MadoguchiTankaTekiyou " +//単価適用地域
                            ",MadoguchiShimekiribi " +//調査担当者への締切日
                            ",MadoguchiNiwatashi " +//荷渡場所
                            ",MadoguchiHoukokuJisshibi " +//報告実施日 33
                            ",MadoguchiHikiwatsahi " +//遠隔地引渡承認
                            ",MadoguchiSaishuuKensa " +//遠隔地最終検査
                            ",MadoguchiShouninsha " +//遠隔地承認者
                            ",MadoguchiShouninnbi " +//遠隔地承認日
                            ",MadoguchiShukeiHyoFolder  " +//集計表フォルダ
                            ",MadoguchiHoukokuShoFolder " +//報告書フォルダ
                            ",MadoguchiShiryouHolder " +//調査資料フォルダ
                            ",MadoguchiHoukokuzumi " +
                            ",MadoguchiAnkenJouhouID " +//案件情報ID
                            ",MadoguchiJutakuTantoushaID " +//契約担当者CD
                            ",MadoguchiKanriGijutsusha " +//管理技術者CD
                            ",MadoguchiHonbuTanpinflg " +//本部単品
                            ",MadoguchiGaroonRenkei " +//Garoon連携
                            "FROM MadoguchiJouhou " +
                            //"LEFT JOIN AnkenJouhou ON AnkenJutakuBangou = replace(MadoguchiJutakuBangou,'-' + MadoguchiJutakuBangouEdaban,'') " +
                            //"AND MadoguchiJutakuBangouEdaban = AnkenJutakuBangouEda " +
                            "LEFT JOIN AnkenJouhou ON AnkenJouhou.AnkenJouhouID = MadoguchiJouhou.AnkenJouhouID " +
                            //"      AND AnkenJutakuBangou = MadoguchiJutakuBangou + '-' + MadoguchiJutakuBangouEdaban " +
                            "LEFT JOIN Mst_Chousain ON  MadoguchiTantoushaCD = KojinCD " +
                            "LEFT JOIN Mst_Busho mb1 ON JutakuBushoShozokuCD = mb1.GyoumuBushoCD " +
                            "LEFT JOIN Mst_Busho mb2 ON MadoguchiBushoShozokuCD = mb2.GyoumuBushoCD " +
                            "LEFT JOIN GyoumuJouhou gj ON AnkenJouhou.AnkenJouhouID = gj.AnkenJouhouID " +
                            "AND MadoguchiKanriGijutsusha = gj.KanriGijutsushaCD " +
                            "WHERE MadoguchiID = " + MadoguchiID + "";

                        var sda = new SqlDataAdapter(cmd);
                        MadoguchiData.Clear();
                        sda.Fill(MadoguchiData);
                    }

                    // 担当部所タブ
                    if (tab == 2)
                    {
                        //調査担当者の取得
                        cmd.CommandText = "SELECT " +
                            "CASE  " +
                                "WHEN MadoguchiHoukokuzumi = 1 THEN 8 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShinchokuJoukyou = 6 THEN 6 " +　//MadoguchiShinchokuJoukyouが2桁コードに変更
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShinchokuJoukyou = 3 THEN 5 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<GETDATE() THEN 1 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<GETDATE() +3 AND MadoguchiShinchokuJoukyou<> 3 THEN 2 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<GETDATE() +7 AND MadoguchiShinchokuJoukyou<> 3 THEN 3 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShinchokuJoukyou = 80 THEN 6 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShinchokuJoukyou = 70 THEN 5 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<'" + DateTime.Today + "' THEN 1 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<='" + DateTime.Today.AddDays(3) + "' AND MadoguchiShinchokuJoukyou<> 70 THEN 2 " +
                                         //"WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiShimekiribi<='" + DateTime.Today.AddDays(7) + "' AND MadoguchiShinchokuJoukyou<> 70 THEN 3 " +
                                         //"ELSE 4 " +
                                         "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShinchoku = 80 THEN 6 " +
                                         "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShinchoku = 70 THEN 5 " +
                                         "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShinchoku = 50 THEN 7 " +
                                         "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShinchoku = 60 THEN 7 " + // 一次検済
                                         "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShimekiribi<'" + DateTime.Today + "' THEN 1 " +
                                         "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShimekiribi<='" + DateTime.Today.AddDays(3) + "' AND MadoguchiL1ChousaShinchoku <> 70 THEN 2 " +
                                         "WHEN MadoguchiHoukokuzumi<> 1 AND MadoguchiL1ChousaShimekiribi<='" + DateTime.Today.AddDays(7) + "' AND MadoguchiL1ChousaShinchoku <> 70 THEN 3 " +
                                         "ELSE 4 " +
                            "END AS 'SortID'  " +
                            ", MadoguchiL1ChousaCD " +
                            ", MadoguchiL1ChousaBushoCD " +
                            ", MadoguchiL1ChousaTantoushaCD " +
                            ", MadoguchiL1ChousaShimekiribi " +
                            ", MadoguchiL1ChousaShinchoku " +
                            ", MadoguchiL1ChousaKakunin " +
                            ", BushokanriboKamei " +
                            "FROM MadoguchiJouhouMadoguchiL1Chou " +
                            "LEFT JOIN MadoguchiJouhou ON MadoguchiJouhouMadoguchiL1Chou.MadoguchiID = MadoguchiJouhou.MadoguchiID " +
                            "LEFT JOIN Mst_Busho ON MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaBushoCD = Mst_Busho.GyoumuBushoCD " +
                            //"WHERE MadoguchiJouhou.MadoguchiID = '" + MadoguchiID + "' ORDER BY SortID ";
                            // 638対応
                            "WHERE MadoguchiJouhou.MadoguchiID = '" + MadoguchiID + "' ORDER BY MadoguchiJouhouMadoguchiL1Chou.MadoguchiID,MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaCD ";

                        Console.WriteLine(cmd.CommandText);
                        var sda = new SqlDataAdapter(cmd);
                        DT_MadoguchiL1Chou.Clear();
                        sda.Fill(DT_MadoguchiL1Chou);

                        //Garoon追加宛先の取得
                        cmd.CommandText = "SELECT " +
                            "  GaroonTsuikaAtesakiID " +
                            ", GaroonTsuikaAtesakiBushoCD " +
                            ", GaroonTsuikaAtesakiTantoushaCD " +
                            //不具合No1332(1084) 画面から登録されたか否かのフラグ
                            ", GaroonTsuikaAtesakiGamenFlag " +
                            "FROM GaroonTsuikaAtesaki " +
                            "WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' AND GaroonTsuikaAtesakiDeleteFlag <> 1 ";

                        Console.WriteLine(cmd.CommandText);
                        sda = new SqlDataAdapter(cmd);
                        DT_GaroonTsuikaAtesaki.Clear();
                        sda.Fill(DT_GaroonTsuikaAtesaki);
                    }
                    else if (tab == 3)
                    {
                        //調査品目　件数の取得
                        cmd.CommandText = "SELECT " +
                                          "	  COUNT(ChousaHinmokuID) " +
                                           "FROM " +
                                           "	ChousaHinmoku  " +
                                           "WHERE " +
                                           "	MadoguchiID = " + MadoguchiID + " AND ChousaDeleteFlag <> 1 AND ChousaHinmokuID > 0 ";
                        Console.WriteLine(cmd.CommandText);
                        var sda = new SqlDataAdapter(cmd);
                        DataTable dt = new DataTable();
                        sda.Fill(dt);
                        if (dt != null)
                        {
                            item3_RegistrationRowCount.Text = dt.Rows[0][0].ToString();
                        }
                        else
                        {
                            item3_RegistrationRowCount.Text = "0";
                        }

                        string buf = "";
                        // 調査品目明細のリンク先表示フラグ 0:表示 1:非表示
                        buf = GlobalMethod.GetCommonValue1("CHOUSA_LINK_FLG");
                        if (buf != null && buf == "")
                        {
                            if (buf == "1")
                            {
                                chousaLinkFlg = "1";
                                // &i_LinkImage.Tooltiptext = "現在リンク先は非表示モードが設定されています。"
                            }
                            else if (buf == "0")
                            {
                                chousaLinkFlg = "0";
                            }
                            else
                            {
                                chousaLinkFlg = "0";
                            }
                        }


                        // 集計表フォルダのパスを取得
                        ShukeiHyoFolder = "";
                        cmd.CommandText = "SELECT " +
                                          "	  MadoguchiShukeiHyoFolder " +
                                           "FROM " +
                                           "	MadoguchiJouhou  " +
                                           "WHERE " +
                                           "	MadoguchiID = " + MadoguchiID + " ";
                        Console.WriteLine(cmd.CommandText);
                        sda = new SqlDataAdapter(cmd);
                        dt = new DataTable();
                        sda.Fill(dt);

                        //// 集計表フォルダ存在フラグ true：存在する false：存在しない
                        //Boolean existsFlg = false;
                        // 集計表フォルダ
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            ShukeiHyoFolder = dt.Rows[0][0].ToString();

                            //existsFlg = (Directory.Exists(ShukeiHyoFolder));
                        }

                        string where = "";

                        //調査品目の取得
                        cmd.CommandText = "SELECT " +
                                   //"   ChousaHinmokuID " + // 0
                                   "   0 " + // 0:未使用
                                   ", " +
                                    "CASE " +
                                    "WHEN MadoguchiHoukokuzumi = 1 THEN '8' " + //報告済み
                                    "WHEN MadoguchiHoukokuzumi != 1 THEN " +
                                    "     CASE " +
                                    "         WHEN ChousaChuushi = 1 THEN '6' " + //中止
                                    "         WHEN ChousaShinchokuJoukyou = 80 THEN '6' " + //中止
                                    "         WHEN ChousaShinchokuJoukyou = 70 THEN '5' " + //二次検済
                                    "         WHEN ChousaShinchokuJoukyou = 50 THEN '7' " + //担当者済
                                    "         WHEN ChousaShinchokuJoukyou = 60 THEN '7' " + // 一次検済
                                    "     ELSE " +
                                    "         CASE " +
                                    "              WHEN ChousaHinmokuShimekiribi < '" + DateTime.Today + "' THEN '1' " + //締切日超過
                                    "              WHEN ChousaHinmokuShimekiribi <= '" + DateTime.Today.AddDays(3) + "' THEN '2' " +//締切日3日以内
                                    "              WHEN ChousaHinmokuShimekiribi <= '" + DateTime.Today.AddDays(7) + "' THEN '3' " +//締切日7日以内
                                    "         ELSE '4' " +
                                    "         END " +
                                    "     END " +
                                    "END AS Shinchock " +
                                   " , ChousaZentaiJun " +
                                   " , ChousaKobetsuJun " +
                                   " , ChousaZaiKou " +
                                   " , ChousaHinmei " + // 5
                                   " , ChousaKikaku " +
                                   " , ChousaTanka " +
                                   " , ChousaSankouShitsuryou " +
                                   " , ChousaKakaku " +
                                   " , ChousaChuushi  " + // 10
                                   " , ChousaBikou2 " +
                                   " , ChousaBikou " +
                                   " , ChousaTankaTekiyouTiku " +
                                   " , ChousaZumenNo " +
                                   " , ChousaSuuryou " + //15
                                   " , ChousaMitsumorisaki " +
                                   " , ChousaBaseMakere " +
                                   " , ChousaBaseTanka " +
                                   " , ChousaKakeritsu " +
                                   " , ChousaObiMei " + // 20
                                   " , ChousaZenkaiTani " +
                                   " , ChousaZenkaiKakaku " +
                                   " , ChousaSankouti " +
                                   " , ChousaHinmokuJouhou1 " +
                                   " , ChousaHinmokuJouhou2 " + // 25
                                   " , ChousaFukuShizai  " +
                                   " , ChousaBunrui " +
                                   " , ChousaMemo2 " +
                                   " , ChousaTankaCD1 " +
                                   " , ChousaTikuWariCode " + // 30
                                   " , ChousaTikuCode " +
                                   " , ChousaTikuMei " +
                                   " , ChousaShougaku " +
                                   " , ChousaWebKen " +
                                   " , ChousaKonkyoCode "; // 35
                        // フォルダアイコン切り替え 0:グレー 1:イエロー 2:Excleアイコン
                        if (chousaLinkFlg != "1" && ShukeiHyoFolder != "")
                        {
                            cmd.CommandText += " , CASE WHEN ISNULL(ChousaLinkSakli,'') <> '' THEN 2 ELSE 1 END AS Folder ";
                        }
                        else
                        {
                            cmd.CommandText += " , 0 AS Folder ";
                        }

                        cmd.CommandText += " , ChousaLinkSakli " + // 37:フォルダ先リンク
                                   " , HinmokuRyakuBushoCD " +
                                   //"  , HinmokuRyakuBushoMei " +
                                   " , HinmokuChousainCD " +
                                   //" , HinmokuChousain " +
                                   " , HinmokuRyakuBushoFuku1CD " + // 40
                                                                    //" , HinmokuRyakuBushoMeiFuku1 " +
                                   " , HinmokuFukuChousainCD1 " +
                                   //" , HinmokuFukuChousain1 " +
                                   " , HinmokuRyakuBushoFuku2CD " +
                                   //" , HinmokuRyakuBushoMeiFuku2 " +
                                   " , HinmokuFukuChousainCD2 " +
                                   //" , HinmokuFukuChousain2 " +
                                   //" , ChousaHoukokuHonsuu " +
                                   //" , ChousaHoukokuRank " + // 45
                                   //" , ChousaIraiHonsuu " +
                                   //" , ChousaIraiRank " +
                                   " , ChousaHoukokuRank " +
                                   " , ChousaHoukokuHonsuu " +
                                   " , ChousaIraiRank " +
                                   " , ChousaIraiHonsuu " +
                                   " , ChousaHinmokuShimekiribi " +
                                   " , ChousaHoukokuzumi " +
                                   //" , ChousaTankaCD1 " + // 50:発注品目コード
                                   " , ChousaDeleteFlag " + // 50
                                   " , ChousaHinmokuID " + // 51:調査品目ID
                                   " , ChousaShinchokuJoukyou " + // 52:進捗状況
                                   " , 1 " + // 53:0:Insert/1:Select/2:Update
                                   " , ChousaTankaCD1 " + // 54:発注品目コード
                                   " , '' " + // 55:並び順
                                   //奉行エクセル
                                   ", ChousaShuukeihyouVer" +
                                   ", ChousaBunkatsuHouhou" +
                                   ", ChousaKoujiKouzoubutsumei" +
                                   ", ChousaHachushaTeikyouTani" +
                                   ", ChousaTaniAtariKakaku" +
                                   ", chousaTaniAtariSuuryou" +
                                   ", ChousaTaniAtariTanka" +
                                   ", ChousaNiwatashiJouken" +
                                   ", MadoguchiGroupMei" +
                                   " , ISNULL(MC0.RetireFlg, 0) AS RetireFlg " + //担当者退職フラグ
                                   " , ISNULL(MC1.RetireFlg, 0) AS RetireFlg1 " + //副担当者1退職フラグ
                                   " , ISNULL(MC2.RetireFlg, 0) AS RetireFlg2 " + //副担当者2退職フラグ
                                   " , MC0.ChousainMei " + // 調査員名
                                   " , MC1.ChousainMei AS FukuChousainMei1 " + // 副調査員名1
                                   " , MC2.ChousainMei AS FukuChousainMei2 " + // 副調査員名2
                                   "FROM " +
                                   " ChousaHinmoku  " +
                                   //奉行エクセル
                                   "LEFT JOIN MadoguchiGroupMaster ON ChousaHinmoku.MadoguchiID = MadoguchiGroupMaster.MadoguchiGroupMasterID " +
                                   "LEFT JOIN MadoguchiJouhou ON MadoguchiJouhou.MadoguchiID = ChousaHinmoku.MadoguchiID " +
                                   "LEFT JOIN Mst_Chousain MC0 ON HinmokuChousainCD = MC0.KojinCD " +
                                   "LEFT JOIN Mst_Chousain MC1 ON HinmokuFukuChousainCD1 = MC1.KojinCD " +
                                   "LEFT JOIN Mst_Chousain MC2 ON HinmokuFukuChousainCD2 = MC2.KojinCD " +
                                   "WHERE " +
                                   "MadoguchiJouhou.MadoguchiID = " + MadoguchiID + " AND ChousaDeleteFlag <> 1 AND ChousaHinmokuID > 0 ";

                        //調査品目検索条件
                        //調査担当部所
                        //if ((src_Busho.Text != "" || src_Busho.SelectedValue.ToString() != "") && src_ShuFuku.Text != "")
                        if ((src_Busho.Text != "" && src_Busho.SelectedValue.ToString() != "") && src_ShuFuku.Text != "")
                        {
                            //「本部」以外を選択している場合
                            if (src_ShuFuku.SelectedValue.ToString() != "127950")
                            {
                                //主＋副
                                if (src_ShuFuku.SelectedValue.ToString() == "0")
                                {
                                    cmd.CommandText += " AND (HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue + "' OR HinmokuRyakuBushoFuku1CD = '" + src_Busho.SelectedValue + "' OR HinmokuRyakuBushoFuku2CD = '" + src_Busho.SelectedValue + "' ) ";
                                    where += " AND (HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue + "' OR HinmokuRyakuBushoFuku1CD = '" + src_Busho.SelectedValue + "' OR HinmokuRyakuBushoFuku2CD = '" + src_Busho.SelectedValue + "' ) ";
                                }
                                //主
                                else if (src_ShuFuku.SelectedValue.ToString() == "1")
                                {
                                    cmd.CommandText += " AND (HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue + "' ) ";
                                    where += " AND (HinmokuRyakuBushoCD = '" + src_Busho.SelectedValue + "' ) ";
                                }
                                //副
                                else if (src_ShuFuku.SelectedValue.ToString() == "2")
                                {
                                    cmd.CommandText += " AND (HinmokuRyakuBushoFuku1CD = '" + src_Busho.SelectedValue + "' OR HinmokuRyakuBushoFuku2CD = '" + src_Busho.SelectedValue + "' ) ";
                                    where += " AND (HinmokuRyakuBushoFuku1CD = '" + src_Busho.SelectedValue + "' OR HinmokuRyakuBushoFuku2CD = '" + src_Busho.SelectedValue + "' ) ";
                                }
                            }
                            else
                            {
                                //「本部」は窓口表示フラグがFalseのため、選択されない
                            }
                        }

                        //調査担当者
                        if (src_HinmokuChousain.Text != "" && src_ShuFuku.Text != "")
                        {
                            string Chousain = GlobalMethod.ChangeSqlText(src_HinmokuChousain.Text, 1);
                            //主＋副
                            if (src_ShuFuku.SelectedValue.ToString() == "0")
                            {
                                cmd.CommandText += " AND (MC0.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' OR MC1.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' OR MC2.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' ) ";
                                where += " AND (MC0.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' OR MC1.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' OR MC2.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' ) ";
                            }
                            //主
                            else if (src_ShuFuku.SelectedValue.ToString() == "1")
                            {
                                cmd.CommandText += " AND (MC0.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' ) ";
                                where += " AND (MC0.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' ) ";
                            }
                            //副
                            else if (src_ShuFuku.SelectedValue.ToString() == "2")
                            {
                                cmd.CommandText += " AND (MC1.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' OR MC2.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' ) ";
                                where += " AND (MC1.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' OR MC2.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' ) ";
                            }
                        }

                        //品名
                        if (src_ChousaHinmei.Text != "")
                        {
                            cmd.CommandText += " AND (ChousaHinmei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_ChousaHinmei.Text, 1) + "%' ESCAPE'\\' ) ";
                            where += " AND (ChousaHinmei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_ChousaHinmei.Text, 1) + "%' ESCAPE'\\' ) ";
                        }

                        //規格
                        if (src_ChousaKikaku.Text != "")
                        {
                            cmd.CommandText += " AND (ChousaKikaku COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_ChousaKikaku.Text, 1) + "%' ESCAPE'\\' ) ";
                            where += " AND (ChousaKikaku COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + GlobalMethod.ChangeSqlText(src_ChousaKikaku.Text, 1) + "%' ESCAPE'\\' ) ";
                        }

                        //材工
                        if (src_Zaikou.Text != "")
                        {
                            //1：材のみ
                            if (src_Zaikou.SelectedValue.ToString() == "1")
                            {
                                cmd.CommandText += " AND ChousaZaiKou = 1 ";
                                where += " AND ChousaZaiKou = 1 ";
                            }
                            //2：工のみ
                            else if (src_Zaikou.SelectedValue.ToString() == "2")
                            {
                                cmd.CommandText += " AND (ChousaZaiKou = 2 OR ChousaZaiKou = 3) ";
                                where += " AND (ChousaZaiKou = 2 OR ChousaZaiKou = 3) ";
                            }
                            //3：材+D工
                            else if (src_Zaikou.SelectedValue.ToString() == "3")
                            {
                                cmd.CommandText += " AND (ChousaZaiKou = 1 OR ChousaZaiKou = 2) ";
                                where += " AND (ChousaZaiKou = 1 OR ChousaZaiKou = 2) ";
                            }
                            //4：E工のみ
                            else if (src_Zaikou.SelectedValue.ToString() == "4")
                            {
                                cmd.CommandText += " AND ChousaZaiKou = 3 ";
                                where += " AND ChousaZaiKou = 3 ";
                            }
                            //5：他
                            else if (src_Zaikou.SelectedValue.ToString() == "5")
                            {
                                cmd.CommandText += " AND ChousaZaiKou = 4 ";
                                where += " AND ChousaZaiKou = 4 ";
                            }
                        }

                        //担当者空白リスト
                        if (src_TantoushaKuuhaku.Text != "")
                        {
                            if (src_TantoushaKuuhaku.SelectedValue.ToString() == "1")
                            {
                                cmd.CommandText += " AND ISNULL(HinmokuChousainCD,'') = '' ";
                                where += " AND ISNULL(HinmokuChousainCD,'') = '' ";
                            }
                            else if (src_TantoushaKuuhaku.SelectedValue.ToString() == "2")
                            {
                                cmd.CommandText += " AND ISNULL(HinmokuChousainCD,'') <> '' ";
                                where += " AND ISNULL(HinmokuChousainCD,'') <> '' ";
                            }
                        }

                        // メモ1
                        if (item3_Memo1.Text != "")
                        {
                            cmd.CommandText += " AND ChousaBunrui COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(item3_Memo1.Text, 0) + "%' ESCAPE'\\' ";
                            where += " AND ChousaBunrui COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(item3_Memo1.Text, 0) + "%' ESCAPE'\\' ";
                        }

                        // メモ2
                        if (item3_Memo2.Text != "")
                        {
                            cmd.CommandText += " AND ChousaMemo2 COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(item3_Memo2.Text, 0) + "%' ESCAPE'\\' ";
                            where += " AND ChousaMemo2 COLLATE Japanese_XJIS_100_CI_AS_SC like N'%" + GlobalMethod.ChangeSqlText(item3_Memo2.Text, 0) + "%' ESCAPE'\\' ";
                        }

                        // 検索したときの条件を取得しておく
                        if (where != "")
                        {
                            // 頭の「 AND」以外を入れる
                            chousaHinmokuSearchWhere = where.Substring(4);
                        }
                        else
                        {
                            chousaHinmokuSearchWhere = "";
                        }

                        cmd.CommandText += " ORDER BY ChousaZentaiJun, ChousaKobetsuJun, ChousaHinmokuID ";
                        Console.WriteLine(cmd.CommandText);
                        sda = new SqlDataAdapter(cmd);
                        DT_ChousaHinmoku.Clear();
                        sda.Fill(DT_ChousaHinmoku);

                        //調査品目の表示初期化
                        c1FlexGrid4.Rows.Count = 2;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            set_data(tab);
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }


        private void set_data(int tab)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            if (tab == 1 && MadoguchiData != null)
            {
                //登録年度を変数に格納
                int.TryParse((string)MadoguchiData.Rows[0][8], out Nendo);
                // 登録年度
                item1_TourokuNendo.Text = MadoguchiData.Rows[0][8].ToString();

                // 窓口部所
                item1_MadoguchiTantoushaBushoCD.SelectedValue = MadoguchiData.Rows[0][3].ToString();
                // 特調番号
                item1_MadoguchiUketsukeBangou.Text = MadoguchiData.Rows[0][14].ToString();
                item1_MadoguchiUketsukeBangouEdaban.Text = MadoguchiData.Rows[0][15].ToString();
                // 業務区分
                item1_AnkenGyoumuKubun.SelectedValue = MadoguchiData.Rows[0][24].ToString();
                // 管理番号
                item1_MadoguchiKanriBangou.Text = MadoguchiData.Rows[0][17].ToString();
                // 調査種別
                item1_MadoguchiChousaShubetsu.SelectedValue = MadoguchiData.Rows[0][25].ToString();
                // 単価適用地域
                item1_MadoguchiTankaTekiyou.Text = MadoguchiData.Rows[0][30].ToString();
                // 荷渡場所
                item1_MadoguchiNiwatashi.Text = MadoguchiData.Rows[0][32].ToString();
                // 窓口担当者
                item1_MadoguchiTantousha.Text = MadoguchiData.Rows[0][4].ToString();
                // 発注者詳細名
                item1_MadoguchiHachuuKikanmei.Text = MadoguchiData.Rows[0][16].ToString();
                // 業務名称
                item1_MadoguchiGyoumuMeishou.Text = MadoguchiData.Rows[0][18].ToString();
                // 工事件名
                item1_MadoguchiKoujiKenmei.Text = MadoguchiData.Rows[0][23].ToString();
                // 調査品目
                item1_MadoguchiChousaHinmoku.Text = MadoguchiData.Rows[0][26].ToString();
                // 集計表
                item1_MadoguchiShukeiHyoFolder.Text = MadoguchiData.Rows[0][38].ToString();
                // 報告書
                item1_MadoguchiHoukokuShoFolder.Text = MadoguchiData.Rows[0][39].ToString();
                // 調査資料
                item1_MadoguchiShiryouHolder.Text = MadoguchiData.Rows[0][40].ToString();
                // 報告済
                MadoguchiHoukokuzumi = MadoguchiData.Rows[0][41].ToString();
                // 窓口担当者CD
                item1_MadoguchiTantoushaCD.Text = MadoguchiData.Rows[0][7].ToString();
                //締切日
                item1_MadoguchiShimekiribi.Text = MadoguchiData.Rows[0][31].ToString();
                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携の値を取得
                //Garoon連携
                item1_GaroonRenkei.Checked = bool_str(MadoguchiData.Rows[0][46].ToString());
            }

            //担当部所タブ
            if (tab == 2)
            {

                //担当部所の調査担当者を初期化
                c1FlexGrid1.Rows.Count = 1;
                //担当部所のGaroon追加宛先を初期化
                c1FlexGrid5.Rows.Count = 1;

                if (DT_MadoguchiL1Chou != null)
                {
                    //担当部署の部所一覧を初期化
                    for (int k = 1; k < 26; k++)
                    {
                        //((CheckBox)tableLayoutPanel17.Controls["KyoroykuBusho" + k]).Checked = false;
                        if (k == 1) { KyoroykuBusho1.Checked = false; }
                        if (k == 2) { KyoroykuBusho2.Checked = false; }
                        if (k == 3) { KyoroykuBusho3.Checked = false; }
                        if (k == 4) { KyoroykuBusho4.Checked = false; }
                        if (k == 5) { KyoroykuBusho5.Checked = false; }
                        if (k == 6) { KyoroykuBusho6.Checked = false; }
                        if (k == 7) { KyoroykuBusho7.Checked = false; }
                        if (k == 8) { KyoroykuBusho8.Checked = false; }
                        if (k == 9) { KyoroykuBusho9.Checked = false; }
                        if (k == 10) { KyoroykuBusho10.Checked = false; }
                        if (k == 11) { KyoroykuBusho11.Checked = false; }
                        if (k == 12) { KyoroykuBusho12.Checked = false; }
                        if (k == 13) { KyoroykuBusho13.Checked = false; }
                        if (k == 14) { KyoroykuBusho14.Checked = false; }
                        if (k == 15) { KyoroykuBusho15.Checked = false; }
                        if (k == 16) { KyoroykuBusho16.Checked = false; }
                        if (k == 17) { KyoroykuBusho17.Checked = false; }
                        if (k == 18) { KyoroykuBusho18.Checked = false; }
                        if (k == 19) { KyoroykuBusho19.Checked = false; }
                        if (k == 20) { KyoroykuBusho20.Checked = false; }
                        if (k == 21) { KyoroykuBusho21.Checked = false; }
                        if (k == 22) { KyoroykuBusho22.Checked = false; }
                        if (k == 23) { KyoroykuBusho23.Checked = false; }
                        if (k == 24) { KyoroykuBusho24.Checked = false; }
                        if (k == 25) { KyoroykuBusho25.Checked = false; }
                    }
                    for (int i = 0; i < DT_MadoguchiL1Chou.Rows.Count; i++)
                    {
                        //調査担当者をセット
                        c1FlexGrid1.Rows.Add();
                        // 1198 アイコン見切れ対策
                        c1FlexGrid1.Rows[i + 1].Height = 28;
                        for (int k = 0; k < c1FlexGrid1.Cols.Count; k++)
                        {
                            c1FlexGrid1.Rows[i + 1][k] = DT_MadoguchiL1Chou.Rows[i][k];
                        }
                        //部所一覧をセット
                        for (int k = 1; k < 26; k++)
                        {
                            //if (((CheckBox)tableLayoutPanel17.Controls["KyoroykuBusho" + k]).Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString()))
                            //{
                            //    ((CheckBox)tableLayoutPanel17.Controls["KyoroykuBusho" + k]).Checked = true;
                            //}
                            if (k == 1) { if (KyoroykuBusho1.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho1.Checked = true; } }
                            if (k == 2) { if (KyoroykuBusho2.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho2.Checked = true; } }
                            if (k == 3) { if (KyoroykuBusho3.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho3.Checked = true; } }
                            if (k == 4) { if (KyoroykuBusho4.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho4.Checked = true; } }
                            if (k == 5) { if (KyoroykuBusho5.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho5.Checked = true; } }
                            if (k == 6) { if (KyoroykuBusho6.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho6.Checked = true; } }
                            if (k == 7) { if (KyoroykuBusho7.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho7.Checked = true; } }
                            if (k == 8) { if (KyoroykuBusho8.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho8.Checked = true; } }
                            if (k == 9) { if (KyoroykuBusho9.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho9.Checked = true; } }
                            if (k == 10) { if (KyoroykuBusho10.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho10.Checked = true; } }
                            if (k == 11) { if (KyoroykuBusho11.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho11.Checked = true; } }
                            if (k == 12) { if (KyoroykuBusho12.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho12.Checked = true; } }
                            if (k == 13) { if (KyoroykuBusho13.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho13.Checked = true; } }
                            if (k == 14) { if (KyoroykuBusho14.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho14.Checked = true; } }
                            if (k == 15) { if (KyoroykuBusho15.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho15.Checked = true; } }
                            if (k == 16) { if (KyoroykuBusho16.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho16.Checked = true; } }
                            if (k == 17) { if (KyoroykuBusho17.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho17.Checked = true; } }
                            if (k == 18) { if (KyoroykuBusho18.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho18.Checked = true; } }
                            if (k == 19) { if (KyoroykuBusho19.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho19.Checked = true; } }
                            if (k == 20) { if (KyoroykuBusho20.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho20.Checked = true; } }
                            if (k == 21) { if (KyoroykuBusho21.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho21.Checked = true; } }
                            if (k == 22) { if (KyoroykuBusho22.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho22.Checked = true; } }
                            if (k == 23) { if (KyoroykuBusho23.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho23.Checked = true; } }
                            if (k == 24) { if (KyoroykuBusho24.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho24.Checked = true; } }
                            if (k == 25) { if (KyoroykuBusho25.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho25.Checked = true; } }
                        }
                    }
                }

                //Garoon
                if (DT_GaroonTsuikaAtesaki != null)
                {
                    for (int i = 0; i < DT_GaroonTsuikaAtesaki.Rows.Count; i++)
                    {
                        //調査担当者をセット
                        c1FlexGrid5.Rows.Add();
                        c1FlexGrid5.Rows[i + 1].Height = 28;
                        //不具合No1332(1084) 画面から登録されたかのフラグを追加で取得
                        c1FlexGrid5.Rows[i + 1].UserData = DT_GaroonTsuikaAtesaki.Rows[i][3];
                        for (int k = 1; k < c1FlexGrid5.Cols.Count; k++)
                        {
                            c1FlexGrid5.Rows[i + 1][k] = DT_GaroonTsuikaAtesaki.Rows[i][k - 1];
                        }
                    }
                }

                Resize_Grid("c1FlexGrid1");
                Resize_Grid("c1FlexGrid5");

            }
            if (tab == 3 && DT_ChousaHinmoku != null)
            {
                string discript = "";
                string value = "";
                string table = "";
                string where = "";

                //描画停止
                c1FlexGrid4.BeginUpdate();
                //for (int i = 0; i < DT_ChousaHinmoku.Rows.Count; i++)
                //{
                //    c1FlexGrid4.Rows.Add();

                //    for (int k = 0; k < c1FlexGrid4.Cols.Count - 4; k++)
                //    {
                //        // 14:中止、37:少額案件[10万/100万]、38:Web建、53:報告済
                //        if (k + 4 == 14 || k + 4 == 37 || k + 4 == 38 || k + 4 == 53)
                //        {
                //            if (DT_ChousaHinmoku.Rows[i][k].ToString() == "1")
                //            {
                //                c1FlexGrid4.SetCellCheck(i + 2, k + 4, C1.Win.C1FlexGrid.CheckEnum.Checked);
                //            }
                //            else
                //            {
                //                c1FlexGrid4.SetCellCheck(i + 2, k + 4, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                //            }
                //        }
                //        // 並び順
                //        else if (k + 4 == 58)
                //        {
                //            // エラー行を先頭にする為、
                //            // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                //            c1FlexGrid4[i + 2, k + 4] = "N" + zeroPadding(c1FlexGrid4[i + 2, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i + 2, 7].ToString());
                //        }
                //        else
                //        {
                //            c1FlexGrid4[i + 2, k + 4] = DT_ChousaHinmoku.Rows[i][k];
                //        }

                //        if (k + 4 == 42 || k + 4 == 44 || k + 4 == 46) { 
                //            // ここの処理が時間が掛かる
                //            //// 調査担当者のコンボを選択している部所で絞る
                //            //discript = "ChousainMei ";
                //            //value = "KojinCD ";
                //            //table = "Mst_Chousain ";
                //            //where = "RetireFLG = 0 AND TokuchoFLG = 1 ";
                //            //// 部所が空でない場合
                //            //if (c1FlexGrid4.Rows[i][k + 4] != null && c1FlexGrid4.Rows[i][k + 4].ToString() != "")
                //            //{
                //            //    where += "AND GyoumuBushoCD = '" + c1FlexGrid4.Rows[i][k + 4].ToString() + "' ";
                //            //}
                //            //else
                //            //{
                //            //    where += "AND GyoumuBushoCD in (select GyoumuBushoCD from Mst_Busho where BushoDeleteFlag = 0 AND BushoMadoguchiHyoujiFlg = 1) ";
                //            //}
                //            ////コンボボックスデータ取得
                //            //DataTable tmpdt11 = GlobalMethod.getData(discript, value, table, where);
                //            //SortedList sl = new SortedList();
                //            //sl = GlobalMethod.Get_SortedList(tmpdt11);

                //            //// 最初にパイプ文字"|"を記述してしまうと入力可になってしまう
                //            //String comboListValue = " ";

                //            //if (tmpdt11 != null && tmpdt11.Rows.Count > 0)
                //            //{
                //            //    for (int cnt = 0; cnt < tmpdt11.Rows.Count; cnt++)
                //            //    {
                //            //        comboListValue += "|" + tmpdt11.Rows[cnt][1].ToString();
                //            //    }
                //            //}

                //            //C1.Win.C1FlexGrid.CellStyle cs1 = c1FlexGrid4.Styles.Add("Combo" + i + 2 + "_" + k + 4); // スタイルを定義
                //            //cs1.ComboList = comboListValue; // ComboListを設定

                //            //C1.Win.C1FlexGrid.CellRange rg1 = c1FlexGrid4.GetCellRange(i, k + 4 + 1); // セルを選択
                //            //rg1.Style = cs1; // スタイルを割り当てる
                //        }

                //    }
                //    //if (c1FlexGrid4[i + 2, 40].ToString() == "1" && !Directory.Exists(c1FlexGrid4[i + 2, 41].ToString()))
                //    //{
                //    //    c1FlexGrid4[i + 2, 40] = "0";
                //    //}

                //    // リンク先アイコン
                //    if (chousaLinkFlg != "1")
                //    {
                //        // フォルダ or Excleアイコンの場合
                //        if (c1FlexGrid4[i + 2, 40].ToString() == "1" || c1FlexGrid4[i + 2, 40].ToString() == "2")
                //        {
                //            // リンク先に登録されているパスが、フォルダで存在するか
                //            if (Directory.Exists(c1FlexGrid4[i + 2, 41].ToString()))
                //            {
                //                c1FlexGrid4[i + 2, 40] = "1";
                //            }
                //            // リンク先に登録されているパスが、ファイルで存在するか
                //            else if (File.Exists(c1FlexGrid4[i + 2, 41].ToString()))
                //            {
                //                c1FlexGrid4[i + 2, 40] = "2";
                //            }
                //            // 上記がなかった場合、集計表フォルダが存在するか
                //            else if (Directory.Exists(ShukeiHyoFolder))
                //            {
                //                c1FlexGrid4[i + 2, 40] = "1";
                //            }
                //            else
                //            {
                //                c1FlexGrid4[i + 2, 40] = "0";
                //            }
                //        }
                //    }
                //    else
                //    {
                //        c1FlexGrid4[i + 2, 40] = "0";
                //    }

                //}

                // c1FlexGrid4の編集開始行の指定
                int RowCount = 2;

                // 取得した調査品目のレコードの行を回す
                for (int i = 0; i < DT_ChousaHinmoku.Rows.Count; i++)
                {
                    c1FlexGrid4.Rows.Add();

                    // 調査品目ID
                    c1FlexGrid4.Rows[RowCount]["ChousaHinmokuID"] = DT_ChousaHinmoku.Rows[i]["Column1"];
                    // 進捗状況
                    c1FlexGrid4.Rows[RowCount]["ShinchokuIcon"] = DT_ChousaHinmoku.Rows[i]["Shinchock"];
                    // 全体順
                    c1FlexGrid4.Rows[RowCount]["ChousaZentaiJun"] = DT_ChousaHinmoku.Rows[i]["ChousaZentaiJun"];
                    // 個別順
                    c1FlexGrid4.Rows[RowCount]["ChousaKobetsuJun"] = DT_ChousaHinmoku.Rows[i]["ChousaKobetsuJun"];
                    // 材工
                    c1FlexGrid4.Rows[RowCount]["ChousaZaiKou"] = DT_ChousaHinmoku.Rows[i]["ChousaZaiKou"];
                    // 品目
                    c1FlexGrid4.Rows[RowCount]["ChousaHinmei"] = DT_ChousaHinmoku.Rows[i]["ChousaHinmei"];
                    // 規格
                    c1FlexGrid4.Rows[RowCount]["ChousaKikaku"] = DT_ChousaHinmoku.Rows[i]["ChousaKikaku"];
                    // 単位
                    c1FlexGrid4.Rows[RowCount]["ChousaTanka"] = DT_ChousaHinmoku.Rows[i]["ChousaTanka"];
                    // 参考質量
                    c1FlexGrid4.Rows[RowCount]["ChousaSankouShitsuryou"] = DT_ChousaHinmoku.Rows[i]["ChousaSankouShitsuryou"];
                    // 価格
                    c1FlexGrid4.Rows[RowCount]["ChousaKakaku"] = DT_ChousaHinmoku.Rows[i]["ChousaKakaku"];
                    // 中止
                    c1FlexGrid4.SetCellCheck(RowCount, c1FlexGrid4.Cols["ChousaChuushi"].Index, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                    if (DT_ChousaHinmoku.Rows[i]["ChousaChuushi"].ToString() == "1")
                    {
                        c1FlexGrid4.SetCellCheck(RowCount, c1FlexGrid4.Cols["ChousaChuushi"].Index, C1.Win.C1FlexGrid.CheckEnum.Checked);
                    }
                    // 報告備考
                    c1FlexGrid4.Rows[RowCount]["ChousaBikou2"] = DT_ChousaHinmoku.Rows[i]["ChousaBikou2"];
                    // 依頼備考
                    c1FlexGrid4.Rows[RowCount]["ChousaBikou"] = DT_ChousaHinmoku.Rows[i]["ChousaBikou"];
                    // 単価適用地域
                    c1FlexGrid4.Rows[RowCount]["ChousaTankaTekiyouTiku"] = DT_ChousaHinmoku.Rows[i]["ChousaTankaTekiyouTiku"];
                    // 図面番号
                    c1FlexGrid4.Rows[RowCount]["ChousaZumenNo"] = DT_ChousaHinmoku.Rows[i]["ChousaZumenNo"];
                    // 数量
                    c1FlexGrid4.Rows[RowCount]["ChousaSuuryou"] = DT_ChousaHinmoku.Rows[i]["ChousaSuuryou"];
                    // 見積先
                    c1FlexGrid4.Rows[RowCount]["ChousaMitsumorisaki"] = DT_ChousaHinmoku.Rows[i]["ChousaMitsumorisaki"];
                    // ベースメーカー
                    c1FlexGrid4.Rows[RowCount]["ChousaBaseMakere"] = DT_ChousaHinmoku.Rows[i]["ChousaBaseMakere"];
                    // ベース単位
                    c1FlexGrid4.Rows[RowCount]["ChousaBaseTanka"] = DT_ChousaHinmoku.Rows[i]["ChousaBaseTanka"];
                    // 掛率
                    c1FlexGrid4.Rows[RowCount]["ChousaKakeritsu"] = DT_ChousaHinmoku.Rows[i]["ChousaKakeritsu"];
                    // 属性
                    c1FlexGrid4.Rows[RowCount]["ChousaObiMei"] = DT_ChousaHinmoku.Rows[i]["ChousaObiMei"];
                    // 前回単位
                    c1FlexGrid4.Rows[RowCount]["ChousaZenkaiTani"] = DT_ChousaHinmoku.Rows[i]["ChousaZenkaiTani"];
                    // 前回価格
                    c1FlexGrid4.Rows[RowCount]["ChousaZenkaiKakaku"] = DT_ChousaHinmoku.Rows[i]["ChousaZenkaiKakaku"];
                    // 発注者提供単価
                    c1FlexGrid4.Rows[RowCount]["ChousaSankouti"] = DT_ChousaHinmoku.Rows[i]["ChousaSankouti"];
                    // 品目情報1
                    c1FlexGrid4.Rows[RowCount]["ChousaHinmokuJouhou1"] = DT_ChousaHinmoku.Rows[i]["ChousaHinmokuJouhou1"];
                    // 品目情報2
                    c1FlexGrid4.Rows[RowCount]["ChousaHinmokuJouhou2"] = DT_ChousaHinmoku.Rows[i]["ChousaHinmokuJouhou2"];
                    // 前回質量
                    c1FlexGrid4.Rows[RowCount]["ChousaFukuShizai"] = DT_ChousaHinmoku.Rows[i]["ChousaFukuShizai"];
                    // メモ1
                    c1FlexGrid4.Rows[RowCount]["ChousaBunrui"] = DT_ChousaHinmoku.Rows[i]["ChousaBunrui"];
                    // メモ2
                    c1FlexGrid4.Rows[RowCount]["ChousaMemo2"] = DT_ChousaHinmoku.Rows[i]["ChousaMemo2"];
                    // 発注品目コード
                    c1FlexGrid4.Rows[RowCount]["ChousaTankaCD1"] = DT_ChousaHinmoku.Rows[i]["ChousaTankaCD1"];
                    // 地区割コード
                    c1FlexGrid4.Rows[RowCount]["ChousaTikuWariCode"] = DT_ChousaHinmoku.Rows[i]["ChousaTikuWariCode"];
                    // 地区コード
                    c1FlexGrid4.Rows[RowCount]["ChousaTikuCode"] = DT_ChousaHinmoku.Rows[i]["ChousaTikuCode"];
                    // 地区名
                    c1FlexGrid4.Rows[RowCount]["ChousaTikuMei"] = DT_ChousaHinmoku.Rows[i]["ChousaTikuMei"];
                    // 少額案件[10万/100万]
                    c1FlexGrid4.SetCellCheck(RowCount, c1FlexGrid4.Cols["ChousaShougaku"].Index, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                    if (DT_ChousaHinmoku.Rows[i]["ChousaShougaku"].ToString() == "1")
                    {
                        c1FlexGrid4.SetCellCheck(RowCount, c1FlexGrid4.Cols["ChousaShougaku"].Index, C1.Win.C1FlexGrid.CheckEnum.Checked);
                    }
                    // Web建
                    c1FlexGrid4.SetCellCheck(RowCount, c1FlexGrid4.Cols["ChousaWebKen"].Index, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                    if (DT_ChousaHinmoku.Rows[i]["ChousaWebKen"].ToString() == "1")
                    {
                        c1FlexGrid4.SetCellCheck(RowCount, c1FlexGrid4.Cols["ChousaWebKen"].Index, C1.Win.C1FlexGrid.CheckEnum.Checked);
                    }
                    // 根拠関連コード
                    c1FlexGrid4.Rows[RowCount]["ChousaKonkyoCode"] = DT_ChousaHinmoku.Rows[i]["ChousaKonkyoCode"];
                    // リンク先
                    c1FlexGrid4.Rows[RowCount]["ChousaLinkSakli"] = "0";
                    if (chousaLinkFlg != "1")
                    {
                        // フォルダ or Excelアイコンの場合
                        if (DT_ChousaHinmoku.Rows[i]["Folder"].ToString() == "1" || DT_ChousaHinmoku.Rows[i]["Folder"].ToString() == "2")
                        {
                            // リンク先に登録されているパスが、ファイルで存在するか
                            if (File.Exists(DT_ChousaHinmoku.Rows[i]["ChousaLinkSakli"].ToString()))
                            {
                                c1FlexGrid4.Rows[RowCount]["ChousaLinkSakli"] = "2";
                            }
                            // リンク先に登録されているパスが、フォルダで存在するか
                            else if (Directory.Exists(DT_ChousaHinmoku.Rows[i]["ChousaLinkSakli"].ToString()))
                            {
                                c1FlexGrid4.Rows[RowCount]["ChousaLinkSakli"] = "1";
                            }
                            // 上記がなかった場合、集計表フォルダが存在するか
                            else if (Directory.Exists(ShukeiHyoFolder))
                            {
                                c1FlexGrid4.Rows[RowCount]["ChousaLinkSakli"] = "1";
                            }
                        }
                    }
                    // リンク先パス
                    c1FlexGrid4.Rows[RowCount]["ChousaLinkSakliFolder"] = DT_ChousaHinmoku.Rows[i]["ChousaLinkSakli"];
                    // 調査担当部所
                    c1FlexGrid4.Rows[RowCount]["HinmokuRyakuBushoCD"] = DT_ChousaHinmoku.Rows[i]["HinmokuRyakuBushoCD"];
                    // 調査担当者（修正後）No1427　1201 嘱託に転籍された方が個人CDで表示される
                    //// 調査担当者（修正前）名前でなくコードで表示
                    //c1FlexGrid4.Rows[RowCount]["HinmokuChousainCD"] = DT_ChousaHinmoku.Rows[i]["HinmokuChousainCD"];
                    if ((bool)DT_ChousaHinmoku.Rows[i]["RetireFlg"])
                        c1FlexGrid4.Rows[RowCount]["HinmokuChousainCD"] = DT_ChousaHinmoku.Rows[i]["ChousainMei"] + "（退職）";
                    else
                        c1FlexGrid4.Rows[RowCount]["HinmokuChousainCD"] = DT_ChousaHinmoku.Rows[i]["HinmokuChousainCD"];
                    // 副調査担当部所1
                    c1FlexGrid4.Rows[RowCount]["HinmokuRyakuBushoFuku1CD"] = DT_ChousaHinmoku.Rows[i]["HinmokuRyakuBushoFuku1CD"];
                    // 副調査担当者1（修正後）No1427 1201 嘱託に転籍された方が個人CDで表示される
                    ////副調査担当者1（修正前）名前でなくコードで表示
                    //c1FlexGrid4.Rows[RowCount]["HinmokuFukuChousainCD1"] = DT_ChousaHinmoku.Rows[i]["HinmokuFukuChousainCD1"];
                    if ((bool)DT_ChousaHinmoku.Rows[i]["RetireFlg1"])
                        c1FlexGrid4.Rows[RowCount]["HinmokuFukuChousainCD1"] = DT_ChousaHinmoku.Rows[i]["FukuChousainMei1"] + "（退職）";
                    else
                        c1FlexGrid4.Rows[RowCount]["HinmokuFukuChousainCD1"] = DT_ChousaHinmoku.Rows[i]["HinmokuFukuChousainCD1"];
                    // 副調査担当部所2
                    c1FlexGrid4.Rows[RowCount]["HinmokuRyakuBushoFuku2CD"] = DT_ChousaHinmoku.Rows[i]["HinmokuRyakuBushoFuku2CD"];
                    //// 副調査担当者2（修正前）名前でなくコードで表示
                    //c1FlexGrid4.Rows[RowCount]["HinmokuFukuChousainCD2"] = DT_ChousaHinmoku.Rows[i]["HinmokuFukuChousainCD2"];
                    if ((bool)DT_ChousaHinmoku.Rows[i]["RetireFlg2"])
                        c1FlexGrid4.Rows[RowCount]["HinmokuFukuChousainCD2"] = DT_ChousaHinmoku.Rows[i]["FukuChousainMei2"] + "（退職）";
                    else
                        c1FlexGrid4.Rows[RowCount]["HinmokuFukuChousainCD2"] = DT_ChousaHinmoku.Rows[i]["HinmokuFukuChousainCD2"];
                    // 報告数
                    c1FlexGrid4.Rows[RowCount]["ChousaHoukokuHonsuu"] = DT_ChousaHinmoku.Rows[i]["ChousaHoukokuHonsuu"];
                    // 報告ランク
                    c1FlexGrid4.Rows[RowCount]["ChousaHoukokuRank"] = DT_ChousaHinmoku.Rows[i]["ChousaHoukokuRank"];
                    // 依頼数
                    c1FlexGrid4.Rows[RowCount]["ChousaIraiHonsuu"] = DT_ChousaHinmoku.Rows[i]["ChousaIraiHonsuu"];
                    // 依頼ランク
                    c1FlexGrid4.Rows[RowCount]["ChousaIraiRank"] = DT_ChousaHinmoku.Rows[i]["ChousaIraiRank"];
                    // 締切日
                    c1FlexGrid4.Rows[RowCount]["ChousaHinmokuShimekiribi"] = DT_ChousaHinmoku.Rows[i]["ChousaHinmokuShimekiribi"];
                    // 報告済
                    c1FlexGrid4.SetCellCheck(RowCount, c1FlexGrid4.Cols["ChousaHoukokuzumi"].Index, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                    if (DT_ChousaHinmoku.Rows[i]["ChousaHoukokuzumi"].ToString() == "1")
                    {
                        c1FlexGrid4.SetCellCheck(RowCount, c1FlexGrid4.Cols["ChousaHoukokuzumi"].Index, C1.Win.C1FlexGrid.CheckEnum.Checked);
                    }
                    // 削除フラグ
                    c1FlexGrid4.Rows[RowCount]["ChousaDeleteFlag"] = DT_ChousaHinmoku.Rows[i]["ChousaDeleteFlag"];
                    // 調査品目ID
                    c1FlexGrid4.Rows[RowCount]["ChousaHinmokuID2"] = DT_ChousaHinmoku.Rows[i]["ChousaHinmokuID"];
                    // 進捗状況
                    c1FlexGrid4.Rows[RowCount]["ChousaShinchokuJoukyou"] = DT_ChousaHinmoku.Rows[i]["ChousaShinchokuJoukyou"];
                    // 0:Insert/1:Select/2:Update
                    c1FlexGrid4.Rows[RowCount]["Mode"] = DT_ChousaHinmoku.Rows[i]["Column2"];
                    // 並び順
                    // エラー行を先頭にするため、並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                    c1FlexGrid4.Rows[RowCount]["ColumnSort"] = "N"
                                                             + zeroPadding(DT_ChousaHinmoku.Rows[i]["ChousaZentaiJun"].ToString())
                                                             + "-"
                                                             + zeroPadding(DT_ChousaHinmoku.Rows[i]["ChousaKobetsuJun"].ToString())
                                                             ;
                    // 奉行エクセル
                    // 集計表Ver
                    c1FlexGrid4.Rows[RowCount]["ShukeihyoVer"] = DT_ChousaHinmoku.Rows[i]["ChousaShuukeihyouVer"];

                    //分割方法（ファイル・シート）
                    c1FlexGrid4.Rows[RowCount]["BunkatsuHouhou"] = DT_ChousaHinmoku.Rows[i]["ChousaBunkatsuHouhou"];
                    if (DT_ChousaHinmoku.Rows[i]["ChousaShuukeihyouVer"].ToString() != "2")
                    {
                        //集計表Verが初期値であれば背景色をグレー
                        c1FlexGrid4.GetCellRange(RowCount, 58).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                    }

                    //グループ名
                    c1FlexGrid4.Rows[RowCount]["GroupMei"] = DT_ChousaHinmoku.Rows[i]["MadoguchiGroupMei"];
                    if (DT_ChousaHinmoku.Rows[i]["ChousaShuukeihyouVer"].ToString() != "2")
                    {
                        //集計表Verが初期値であれば背景色をグレー
                        c1FlexGrid4.GetCellRange(RowCount, 59).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                    }

                    //工事・構造物名
                    c1FlexGrid4.Rows[RowCount]["KojiKoubutsuMei"] = DT_ChousaHinmoku.Rows[i]["ChousaKoujiKouzoubutsumei"];
                    //単位当たり単価（単位）
                    c1FlexGrid4.Rows[RowCount]["TaniAtariTankaTani"] = DT_ChousaHinmoku.Rows[i]["ChousaTaniAtariTanka"];
                    //単位当たり単価（数量）
                    c1FlexGrid4.Rows[RowCount]["TaniAtariTankaSuryo"] = DT_ChousaHinmoku.Rows[i]["chousaTaniAtariSuuryou"];
                    //単位当たり単価（価格）
                    c1FlexGrid4.Rows[RowCount]["TaniAtariTankaKakaku"] = DT_ChousaHinmoku.Rows[i]["ChousaTaniAtariKakaku"];
                    //荷渡し条件
                    c1FlexGrid4.Rows[RowCount]["NiwatashiJoken"] = DT_ChousaHinmoku.Rows[i]["ChousaNiwatashiJouken"];
                    //発注者提供単位
                    c1FlexGrid4.Rows[RowCount]["HachushaTeikyoTani"] = DT_ChousaHinmoku.Rows[i]["ChousaHachushaTeikyouTani"];

                    RowCount += 1;
                }

                //不具合No1207
                //共通マスタの値により、行高さを固定にするか、自動高さ調整を行うか。
                gridRowHeightAutoResize(AutoSizeGridRowMode);

                //描画再開
                c1FlexGrid4.EndUpdate();

                // VIPS　20220203　課題管理表No797　ADD　表示件数「全件表示」対応
                // 表示件数
                int hyoujisuu = 0;
                // 「全件表示」の場合
                if (int.TryParse(item_Hyoujikensuu.Text, out hyoujisuu) == false)
                {
                    // かなり大きな数値をセット
                    hyoujisuu = 999999999;
                }

                // VIPS　20220203　課題管理表No797　CHANGE　表示件数「全件表示」対応
                Paging_all.Text = (Math.Ceiling(((double)c1FlexGrid4.Rows.Count - 2) / hyoujisuu)).ToString();
                Grid_Num.Text = "(" + DT_ChousaHinmoku.Rows.Count + ")";
                Grid_Visible(int.Parse(Paging_now.Text));

                //Grid編集許可状態の初期化
                ChousaHinmokuGrid_InputMode();
            }
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }
        //調査品目Gridの行表示切替
        private void Grid_Visible(int page)
        {

            //描画停止
            c1FlexGrid4.BeginUpdate();

            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            {
                // VIPS　20220203　課題管理表No797　CHANGE　表示件数「全件表示」対応
                //表示行フラグ true:表示 false:非表示
                int pagelimit = 0;
                // 「全件表示」の場合
                if (int.TryParse(item_Hyoujikensuu.Text, out pagelimit) == false)
                {
                    // かなり大きな値をセット
                    pagelimit = 999999999;
                }

                //
                if ((page - 1) * pagelimit + 1 < i && i < page * pagelimit + 2)
                {
                    c1FlexGrid4.Rows[i].Visible = true;
                }
                else
                {
                    c1FlexGrid4.Rows[i].Visible = false;
                }
            }
            set_page_enabled(int.Parse(Paging_now.Text), int.Parse(Paging_all.Text));

            //描画再開
            c1FlexGrid4.EndUpdate();

        }
        // ページングアイコンのON/OFF切替
        private void set_page_enabled(int now, int last)
        {
            GlobalMethod.outputLogger("Paging_ChousaHinomoku", "ページ:" + now, "GridAll", UserInfos[1]);
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

        private void Top_Page_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (1).ToString();
            item3_TargetPage.Text = Paging_now.Text;
            Grid_Visible(int.Parse(Paging_now.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void Previous_Page_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (int.Parse(Paging_now.Text) - 1).ToString();
            item3_TargetPage.Text = Paging_now.Text;
            Grid_Visible(int.Parse(Paging_now.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void After_Page_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (int.Parse(Paging_now.Text) + 1).ToString();
            item3_TargetPage.Text = Paging_now.Text;
            Grid_Visible(int.Parse(Paging_now.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void End_Page_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            item3_TargetPage.Text = Paging_now.Text;
            Grid_Visible(int.Parse(Paging_now.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }
        //調査品目タブ　入力開始ボタン押下時
        private void ChousaHinmokuGrid_InputMode()
        {
            Boolean EditMode = false;
            //入力開始
            if (ChousaHinmokuMode == 1)
            {
                EditMode = true;
                //c1FlexGrid4.Cols[0].Visible = true;
                //c1FlexGrid4.Cols[1].Visible = true;
                //c1FlexGrid4.Cols[2].Visible = false;
                //c1FlexGrid4.Cols[3].Visible = false;
                c1FlexGrid4.Cols["Add1"].Visible = true;
                c1FlexGrid4.Cols["Delete1"].Visible = true;
                c1FlexGrid4.Cols["Add2"].Visible = false;
                c1FlexGrid4.Cols["Delete2"].Visible = false;
                button3_Search.BackColor = Color.DimGray;
                button3_Search.Enabled = false;
                button3_Clear.BackColor = Color.DimGray;
                button3_Clear.Enabled = false;
                button3_RowAdd.BackColor = Color.FromArgb(42, 78, 122);
                button3_RowAdd.Enabled = true;
                button3_InputStatus.Text = "入力完了(更新)";

                button3_ExcelHoukokusho.BackColor = Color.DimGray;
                button3_ExcelHoukokusho.Enabled = false;
                button3_ExcelShukeihyo.BackColor = Color.DimGray;
                button3_ExcelShukeihyo.Enabled = false;
                button3_ExcelChousaHinmoku.BackColor = Color.DimGray;
                button3_ExcelChousaHinmoku.Enabled = false;
                //button3_ReadExcelResult.BackColor = Color.DimGray;
                //button3_ReadExcelResult.Enabled = false;
                button3_ReadExcelChousaHinmoku.BackColor = Color.DimGray;
                button3_ReadExcelChousaHinmoku.Enabled = false;

            }
            //入力完了
            else
            {
                //c1FlexGrid4.Cols[0].Visible = false;
                //c1FlexGrid4.Cols[1].Visible = false;
                //c1FlexGrid4.Cols[2].Visible = true;
                //c1FlexGrid4.Cols[3].Visible = true;
                c1FlexGrid4.Cols["Add1"].Visible = false;
                c1FlexGrid4.Cols["Delete1"].Visible = false;
                c1FlexGrid4.Cols["Add2"].Visible = true;
                c1FlexGrid4.Cols["Delete2"].Visible = true;
                button3_Search.BackColor = Color.FromArgb(42, 78, 122);
                button3_Search.Enabled = true;
                button3_Clear.BackColor = Color.FromArgb(42, 78, 122);
                button3_Clear.Enabled = true;
                button3_RowAdd.BackColor = Color.DimGray;
                button3_RowAdd.Enabled = false;
                button3_InputStatus.Text = "入力開始";

                button3_ExcelHoukokusho.BackColor = Color.FromArgb(42, 78, 122);
                button3_ExcelHoukokusho.Enabled = true;
                button3_ExcelShukeihyo.BackColor = Color.FromArgb(42, 78, 122);
                button3_ExcelShukeihyo.Enabled = true;
                button3_ExcelChousaHinmoku.BackColor = Color.FromArgb(42, 78, 122);
                button3_ExcelChousaHinmoku.Enabled = true;
                //button3_ReadExcelResult.BackColor = Color.FromArgb(42, 78, 122);
                //button3_ReadExcelResult.Enabled = true;
                button3_ReadExcelChousaHinmoku.BackColor = Color.FromArgb(42, 78, 122);
                button3_ReadExcelChousaHinmoku.Enabled = true;
            }

            // 調査品目明細のGrid
            //for (int i = 6; i < c1FlexGrid4.Cols.Count - 2; i++)
            for (int i = 0; i < c1FlexGrid4.Cols.Count; i++)
            {
                // 40:リンク先,53:報告済,41:リンク先パス
                // 42:調査担当部所,43:調査担当者,44:副調査担当部所1,45:副調査担当者1,46:副調査担当部所2,47:副調査担当者2,53:報告済
                //if (i != 40 && i != 42 && i != 43 && i != 44 && i != 45 && i != 46 && i != 47 && i != 53)
                //if (i != 40 && i != 53 && i != 41)
                if (c1FlexGrid4.Cols[i].Visible == true)
                {
                    if (i != c1FlexGrid4.Cols["RowChange"].Index
                        && i != c1FlexGrid4.Cols["Add1"].Index
                        && i != c1FlexGrid4.Cols["Delete1"].Index
                        && i != c1FlexGrid4.Cols["Add2"].Index
                        && i != c1FlexGrid4.Cols["Delete2"].Index
                        && i != c1FlexGrid4.Cols["ShinchokuIcon"].Index
                        && i != c1FlexGrid4.Cols["ChousaLinkSakli"].Index
                        && i != c1FlexGrid4.Cols["ChousaLinkSakliFolder"].Index
                        && i != c1FlexGrid4.Cols["ChousaHoukokuzumi"].Index
                        && i != c1FlexGrid4.Cols["HinmokuRyakuBushoCD"].Index
                        && i != c1FlexGrid4.Cols["HinmokuChousainCD"].Index
                        && i != c1FlexGrid4.Cols["HinmokuRyakuBushoFuku1CD"].Index
                        && i != c1FlexGrid4.Cols["HinmokuFukuChousainCD1"].Index
                        && i != c1FlexGrid4.Cols["HinmokuRyakuBushoFuku2CD"].Index
                        && i != c1FlexGrid4.Cols["HinmokuFukuChousainCD2"].Index
                        && i != c1FlexGrid4.Cols["ChousaHoukokuRank"].Index
                        && i != c1FlexGrid4.Cols["ChousaIraiRank"].Index
                        )
                    {
                        c1FlexGrid4.Cols[i].AllowEditing = EditMode;
                    }
                }
            }
        }
        // 戻るボタン
        private void btnReturn_Click(object sender, EventArgs e)
        {
            // I00013:一覧に戻ってもよろしいですか。
            if (MessageBox.Show(GlobalMethod.GetMessage("I00013", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                this.Owner.Show();
                this.Close();
            }
        }
        //ヘッダー特命課長ボタン
        private void btnTokumei_Click(object sender, EventArgs e)
        {
            //Tokumei form = new Tokumei();
            //form.UserInfos = this.UserInfos;
            //form.Show();
            //this.Close();

            Form f = null;
            Boolean openFlg = false;
            for (int i = 0; i < System.Windows.Forms.Application.OpenForms.Count; i++)
            {
                f = System.Windows.Forms.Application.OpenForms[i];
                if (f.Text.IndexOf("特命課長") >= 0 && f.Text.IndexOf("編集") <= -1)
                {
                    f.Show();
                    openFlg = true;
                    break;
                }
            }
            if (!openFlg)
            {
                Tokumei form = new Tokumei();
                form.UserInfos = this.UserInfos;
                form.Show();
                //this.Close();
            }
            this.Close();
        }
        //ヘッダー特徴野郎TOPボタン
        private void btnTokuchoyaro_Click(object sender, EventArgs e)
        {
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
            this.Close();
        }
        //ヘッダー窓口ボタン
        private void btnMadoguchi_Click(object sender, EventArgs e)
        {
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
            this.Close();
        }
        //ヘッダー自分大臣ボタン
        private void btnJibun_Click(object sender, EventArgs e)
        {
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
            this.Close();
        }
        // メッセージ出力
        private void set_error(string mes, int flg = 1)
        {
            // 第2引数が0ならメッセージクリア
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

        // フォルダーパスチェック
        private void FolderPathCheck()
        {
            // 集計表フォルダ(調査概要タブ　調査品目タブ)
            if (Directory.Exists(item1_MadoguchiShukeiHyoFolder.Text))
            {
                item1_MadoguchiShukeiHyoFolder_icon.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
                item3_MadoguchiShukeiHyoFolder_icon.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
                folderIcon = "1"; // 集計表フォルダアイコン 0:グレー 1:イエロー
            }
            else
            {
                item1_MadoguchiShukeiHyoFolder_icon.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
                item3_MadoguchiShukeiHyoFolder_icon.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
                folderIcon = "0"; // 集計表フォルダアイコン 0:グレー 1:イエロー
            }
            // 報告書フォルダ
            if (Directory.Exists(item1_MadoguchiHoukokuShoFolder.Text))
            {
                item1_MadoguchiHoukokuShoFolder_icon.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
            }
            else
            {
                item1_MadoguchiHoukokuShoFolder_icon.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
            }
            // 調査資料フォルダ
            if (Directory.Exists(item1_MadoguchiShiryouHolder.Text))
            {
                item1_MadoguchiShiryouHolder_icon.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
            }
            else
            {
                item1_MadoguchiShiryouHolder_icon.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
            }
        }
        // 集計表フォルダアイコン押下
        private void folderShukeiIcon_Click(object sender, EventArgs e)
        {
            if (item1_MadoguchiShukeiHyoFolder.Text == "")
            {
                System.Diagnostics.Process.Start("EXPLORER.EXE", "");
            }
            else
            {
                // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
                if (System.Text.RegularExpressions.Regex.IsMatch(item1_MadoguchiShukeiHyoFolder.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // 指定されたフォルダパスが存在するなら開く
                    if (item1_MadoguchiShukeiHyoFolder.Text != "" && item1_MadoguchiShukeiHyoFolder.Text != null && Directory.Exists(item1_MadoguchiShukeiHyoFolder.Text))
                    {
                        System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item1_MadoguchiShukeiHyoFolder.Text));
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
        // 報告書フォルダアイコン押下
        private void folderHoukokushoIcon_Click(object sender, EventArgs e)
        {
            if (item1_MadoguchiHoukokuShoFolder.Text == "")
            {
                System.Diagnostics.Process.Start("EXPLORER.EXE", "");
            }
            else
            {
                // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
                if (System.Text.RegularExpressions.Regex.IsMatch(item1_MadoguchiHoukokuShoFolder.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // 指定されたフォルダパスが存在するなら開く
                    if (item1_MadoguchiHoukokuShoFolder.Text != "" && item1_MadoguchiHoukokuShoFolder.Text != null && Directory.Exists(item1_MadoguchiHoukokuShoFolder.Text))
                    {
                        System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item1_MadoguchiHoukokuShoFolder.Text));
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
        // 調査資料・図面フォルダアイコン押下
        private void folderChousaIcon_Click(object sender, EventArgs e)
        {
            if (item1_MadoguchiShiryouHolder.Text == "")
            {
                System.Diagnostics.Process.Start("EXPLORER.EXE", "");
            }
            else
            {
                // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
                if (System.Text.RegularExpressions.Regex.IsMatch(item1_MadoguchiShiryouHolder.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // 指定されたフォルダパスが存在するなら開く
                    if (item1_MadoguchiShiryouHolder.Text != "" && item1_MadoguchiShiryouHolder.Text != null && Directory.Exists(item1_MadoguchiShiryouHolder.Text))
                    {
                        System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item1_MadoguchiShiryouHolder.Text));
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

        private void button3_Search_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            get_data(3);
        }

        //調査品目タブ　検索解除
        private void button3_Clear_Click(object sender, EventArgs e)
        {
            ClearHinmoku();
        }

        private void ClearHinmoku()
        {
            src_Busho.SelectedValue = UserInfos[2];
            src_HinmokuChousain.Text = "";
            src_ShuFuku.SelectedIndex = 0;
            src_ChousaHinmei.Text = "";
            src_ChousaKikaku.Text = "";
            src_Zaikou.SelectedIndex = 0;
            src_TantoushaKuuhaku.SelectedIndex = 0;
            //item_Hyoujikensuu.SelectedIndex = 1;
            item_Hyoujikensuu.SelectedIndex = 4; // 1000件対応

            item3_Memo1.Text = "";
            item3_Memo2.Text = "";

            //調査品目Grid初期化
            c1FlexGrid4.Rows.Count = 2;
            Paging_now.Text = "1";
            item3_TargetPage.Text = Paging_now.Text;
            Paging_all.Text = "0";
            Grid_Num.Text = "(0)";
            //ページングボタン初期化
            Top_Page.Enabled = false;
            Previous_Page.Enabled = false;
            After_Page.Enabled = false;
            End_Page.Enabled = false;

        }

        //「入力開始」「入力完了（更新）」押下時
        private void button3_InputStatus_Click(object sender, EventArgs e)
        {
            string methodName = ".button3_InputStatus_Click";

            // メッセージクリア
            set_error("", 0);
            if (ChousaHinmokuMode == 0)
            {
                //「調査品目を入力出来るようにしますがよろしいですか。」
                if (GlobalMethod.outputMessage("I20308", "", 1) == DialogResult.OK)
                {
                    // 調査品目の削除Keys
                    deleteChousaHinmokuIDs = "";

                    set_error("", 0);
                    // エラーフラグ false：正常 true：エラー
                    Boolean errorFlg = false;

                    //string table = "ChousaHinmoku";
                    //string UserID = "";
                    //string chousainMei = "";

                    //var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                    //using (var conn = new SqlConnection(connStr))
                    //{
                    //    conn.Open();
                    //    var cmd = conn.CreateCommand();
                    //    SqlTransaction transaction = conn.BeginTransaction();
                    //    cmd.Transaction = transaction;

                    //    try
                    //    {
                    //        // Lock情報取得
                    //        // 0:個人コード、1:氏名、2:部所CD、3:部所名、4：Role
                    //        chousainMei = UserInfos[1];

                    //        cmd.CommandText = "SELECT TOP 1 LOCK_USER_ID,LOCK_USER_MEI FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
                    //                          "AND LOCK_KEY = '" + MadoguchiID + "' ";
                    //        DataTable dt = new DataTable();

                    //        var sda = new SqlDataAdapter(cmd);
                    //        sda.Fill(dt);

                    //        if (dt.Rows.Count > 0)
                    //        {
                    //            // Lockテーブルにデータが存在した場合
                    //            UserID = dt.Rows[0][0].ToString();
                    //            chousainMei = dt.Rows[0][1].ToString();
                    //        }
                    //        else
                    //        {
                    //            // Lockテーブルにデータ存在しない場合
                    //            cmd.CommandText = "INSERT INTO T_LOCK(" +
                    //                             " LOCK_TABLE" +
                    //                             ",LOCK_KEY" +
                    //                             ",LOCK_USER_ID" +
                    //                             ",LOCK_USER_MEI" +
                    //                             ",LOCK_DATETIME" +
                    //                             ")VALUES(" +
                    //                             "'" + table + "' " +
                    //                             ",'" + MadoguchiID + "' " +
                    //                             ",'" + UserInfos[0] + "' " +
                    //                             ",'" + UserInfos[1] + "' " +
                    //                             ",SYSDATETIME() " +
                    //                             ")";
                    //            cmd.ExecuteNonQuery();
                    //            //transaction.Commit();
                    //            UserID = UserInfos[0];
                    //            chousainMei = UserInfos[1];
                    //        }
                    //        transaction.Commit();
                    //    }
                    //    catch
                    //    {
                    //        transaction.Rollback();
                    //        errorFlg = true;
                    //    }
                    //    finally
                    //    {
                    //        conn.Close();
                    //    }
                    //}

                    ChousaHinmokuMode = 1;
                    //button3_InputStatus.Text = "入力完了(更新)";
                    // 編集状態する
                    ChousaHinmokuGrid_InputMode();
                }
            }
            else
            {
                //「更新を行いますがよろしいですか。」
                if (GlobalMethod.outputMessage("I20309", "", 1) == DialogResult.OK)
                {
                    //ChousaHinmokuMode = 0;
                    //button3_InputStatus.Text = "入力開始";

                    writeHistory("【開始】調査品目明細の更新を開始します。 ID= " + MadoguchiID);

                    //// 調査品目明細が2件のみ（ヘッダーしかない場合）・・・全件削除
                    //if(c1FlexGrid4.Rows.Count == 2)
                    //{
                    //// 全件削除
                    //using (var conn = new SqlConnection(connStr))
                    //{
                    //    conn.Open();
                    //    var cmd = conn.CreateCommand();
                    //    // 調査品目全削除
                    //    cmd.CommandText = "DELETE FROM ChousaHinmoku " +
                    //        "WHERE MadoguchiID = '" + MadoguchiID + "' ";
                    //    cmd.ExecuteNonQuery();
                    //    conn.Close();

                    //    writeHistory("調査品目が全件削除されました。 ID = " + MadoguchiID + " ");
                    //}
                    //}

                    // 更新前のChousaHinmoku
                    DataTable beforeChousaHinmokuDT = new DataTable();
                    string beforeRyakuBushoCD = "";
                    string beforeChousainCD = "";
                    string afterRyakuBushoCD = "";
                    string afterChousainCD = "";
                    string afterFukuRyakuBushoCD1 = "";
                    string afterFukuChousainCD1 = "";
                    string afterFukuRyakuBushoCD2 = "";
                    string afterFukuChousainCD2 = "";
                    string historyMessage = "";

                    // Gridにデータが存在する場合
                    // １．調査品目の削除Key（ChousaHinmokuIDをカンマ区切りで連結したデータ）があれば削除
                    // ２．c1FlexGrid4 の 57:0:Insert/1:Select/2:Update があり、それで新規か更新、または処理なしを切り分ける
                    // ３．ChousaHinmokuから担当部所の連携を行う（支部備考も）


                    // 品目のCommit件数を取得
                    int i_RecodeCountMax = 0;
                    string w_RecodeCountMax = GlobalMethod.GetCommonValue1("HINMOKU_COMMIT_KENSU");
                    if (w_RecodeCountMax != null)
                    {
                        int.TryParse(w_RecodeCountMax, out i_RecodeCountMax);
                        if (i_RecodeCountMax == 0)
                        {
                            i_RecodeCountMax = 100;
                        }
                    }
                    else
                    {
                        i_RecodeCountMax = 100;
                    }

                    // 特調番号（ヘッダーにあるので利用する）
                    string tokuchoBangou = Header1.Text;

                    // メッセージフラグ1
                    int updmessage1 = 0; // 新規
                    int updmessage2 = 0; // 更新
                    int updmessage3 = 0; // 削除
                    // エラーメッセージフラグ
                    int errmessage1 = 0; // E20307:全体順、個別順が重複しています。
                    int errmessage2 = 0; // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい。
                    int errmessage3 = 0; // E20337:半角数字で入力してください。赤背景の項目を修正して下さい。
                    int errmessage4 = 0; // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    int cnt = 0;
                    Boolean errorFlg = false;
                    // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                    globalErrorFlg = "0";
                    string sysDateTimeStr = "";

                    string insertQuery = "Insert Into ChousaHinmoku( " +
                    "ChousaHinmokuID " +
                    ",MadoguchiID " +
                    ",ChousaZentaiJun " +
                    ",ChousaKobetsuJun " +
                    ",ChousaZaiKou " +
                    ",ChousaHinmei " +
                    ",ChousaKikaku " +
                    ",ChousaTanka " +
                    ",ChousaSankouShitsuryou " +
                    ",ChousaKakaku " +
                    ",ChousaChuushi " +
                    ",ChousaBikou2 " +
                    ",ChousaBikou " +
                    ",ChousaTankaTekiyouTiku " +
                    ",ChousaZumenNo " +
                    ",ChousaSuuryou " +
                    ",ChousaMitsumorisaki " +
                    ",ChousaBaseMakere " +
                    ",ChousaBaseTanka " +
                    ",ChousaKakeritsu " +
                    ",ChousaObiMei " +
                    ",ChousaZenkaiTani " +
                    ",ChousaZenkaiKakaku " +
                    ",ChousaSankouti " +
                    ",ChousaHinmokuJouhou1 " +
                    ",ChousaHinmokuJouhou2 " +
                    ",ChousaFukuShizai " +
                    ",ChousaBunrui " +
                    ",ChousaMemo2 " +
                    ",ChousaTankaCD1 " +
                    ",ChousaTikuWariCode " +
                    ",ChousaTikuCode " +
                    ",ChousaTikuMei " +
                    ",ChousaShougaku " +
                    ",ChousaWebKen " +
                    ",ChousaKonkyoCode " +
                    ",ChousaLinkSakli " +
                    ",HinmokuRyakuBushoCD " +
                    ",HinmokuChousainCD " +
                    ",HinmokuRyakuBushoFuku1CD " +
                    ",HinmokuFukuChousainCD1 " +
                    ",HinmokuRyakuBushoFuku2CD " +
                    ",HinmokuFukuChousainCD2 " +
                    ",ChousaHoukokuHonsuu " +
                    ",ChousaHoukokuRank " +
                    ",ChousaIraiHonsuu " +
                    ",ChousaIraiRank " +
                    ",ChousaHinmokuShimekiribi " +
                    ",ChousaHoukokuzumi " +
                    ",ChousaDeleteFlag " +
                    ",ChousaCreateDate " +
                    ",ChousaCreateUser " +
                    ",ChousaCreateProgram " +
                    ",ChousaUpdateDate " +
                    ",ChousaUpdateUser " +
                    ",ChousaUpdateProgram " +
                    ",ChousaShinchokuJoukyou " +
                    //奉行エクセル　清水検証
                    ",ChousaShuukeihyouVer" +
                    ",ChousaBunkatsuHouhou" +
                    ",ChousaKoujiKouzoubutsumei" +
                    ",ChousaTaniAtariTanka" +
                    ",chousaTaniAtariSuuryou" +
                    ",ChousaTaniAtariKakaku" +
                    ",ChousaNiwatashiJouken" +
                    ",ChousaHachushaTeikyouTani" +
                    ") VALUES ";
                    string valuesText = "";


                    var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                    using (var conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        var cmd = conn.CreateCommand();
                        SqlTransaction transaction = conn.BeginTransaction();
                        cmd.Transaction = transaction;

                        // エラーチェック

                        // 6:全体順、7:個別順が重複している場合
                        // E20307:全体順、個別順が重複しています。
                        // Gridの中で重複していないか確認、DBにある値とも重複を確認する

                        string zentai = "";
                        string kobetsu = "";
                        float zentaiF = 0;
                        float kobetsuF = 0;
                        float zentaiNextF = 0;
                        float kobetsuNextF = 0;
                        int recordCount = 0;
                        //// Grid内で重複していないか確認
                        //for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                        //{
                        //    zentai = c1FlexGrid4.Rows[i][6].ToString();
                        //    kobetsu = c1FlexGrid4.Rows[i][7].ToString();

                        //    float.TryParse(zentai, out zentaiF);
                        //    float.TryParse(kobetsu, out kobetsuF);

                        //    recordCount = 0;

                        //    // 色付けしても次のループでまた塗り直ししてしまうので、全件で回す
                        //    //for(int j = i + 1;j < c1FlexGrid4.Rows.Count; j++)
                        //    for (int j = 2; j < c1FlexGrid4.Rows.Count; j++)
                        //    {

                        //        if (c1FlexGrid4.Rows[j][6] != null && c1FlexGrid4.Rows[j][7] != null && c1FlexGrid4.Rows[j][6].ToString() != "" && c1FlexGrid4.Rows[j][7].ToString() != "")
                        //        {
                        //            zentai = c1FlexGrid4.Rows[j][6].ToString();
                        //            kobetsu = c1FlexGrid4.Rows[j][7].ToString();

                        //            float.TryParse(zentai, out zentaiNextF);
                        //            float.TryParse(kobetsu, out kobetsuNextF);

                        //            if (zentaiF == zentaiNextF && kobetsuF == kobetsuNextF)
                        //            {
                        //                recordCount += 1;
                        //            }
                        //        }
                        //    }
                        //    // 重複レコードがあった場合
                        //    if (recordCount > 1)
                        //    {
                        //        errmessage1 = 1;
                        //        errorFlg = true;
                        //        // ピンク背景
                        //        c1FlexGrid4.GetCellRange(i, 6).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                        //        c1FlexGrid4.GetCellRange(i, 7).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                        //        // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                        //        c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                        //    }
                        //    else
                        //    {
                        //        // クリーム色背景
                        //        c1FlexGrid4.GetCellRange(i, 6).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                        //        c1FlexGrid4.GetCellRange(i, 7).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                        //    }
                        //}

                        // 全体順と個別順のIndex（行番号）を取得する。
                        int ZentaiJunColIndex = c1FlexGrid4.Cols["ChousaZentaiJun"].Index;
                        int KobetsuJunColIndex = c1FlexGrid4.Cols["ChousaKobetsuJun"].Index;

                        // 全体順でソート
                        //c1FlexGrid4.Cols[6].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                        //c1FlexGrid4.Cols.Move(6, 1);
                        //c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 1, 6);
                        //c1FlexGrid4.Cols.Move(1, 6);
                        //c1FlexGrid4.Cols[6].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                        //c1FlexGrid4.Cols.Move(6, 1);
                        //c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 1, 6);
                        //c1FlexGrid4.Cols.Move(1, 6);
                        c1FlexGrid4.Cols[ZentaiJunColIndex].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                        c1FlexGrid4.Cols[KobetsuJunColIndex].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                        c1FlexGrid4.Cols.Move(ZentaiJunColIndex, 1);
                        c1FlexGrid4.Cols.Move(KobetsuJunColIndex, 2);
                        c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 1, 2);
                        c1FlexGrid4.Cols.Move(2, KobetsuJunColIndex);
                        c1FlexGrid4.Cols.Move(1, ZentaiJunColIndex);

                        int row = 0;
                        float zentaiBefore = 0;

                        // Grid内で重複していないか確認
                        //for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                        {
                            // 対象行の全体順、個別順を取得
                            //zentai = c1FlexGrid4.Rows[i][6].ToString();  // 全体順
                            //kobetsu = c1FlexGrid4.Rows[i][7].ToString(); // 個別順
                            //zentai = c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString();  // 全体順
                            //kobetsu = c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString(); // 個別順
                            zentai = (c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0");  // 全体順
                            kobetsu = (c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"); // 個別順

                            float.TryParse(zentai, out zentaiF);
                            float.TryParse(kobetsu, out kobetsuF);

                            recordCount = 0;

                            // 
                            //if(zentaiBefore == 0)
                            //{
                            //    zentaiBefore = zentaiF;
                            //    row = i;
                            //}
                            //// 直前の全体順の行数を取っておく
                            //if (zentaiBefore != zentaiF) 
                            //{
                            //    zentaiBefore = zentaiF;
                            //    row = i;
                            //}
                            if (zentaiBefore == 0 || zentaiBefore != zentaiF)
                            {
                                zentaiBefore = zentaiF;
                                row = i;
                            }

                            // 色付けしても次のループでまた塗り直ししてしまうので、全件で回す
                            //for(int j = i + 1;j < c1FlexGrid4.Rows.Count; j++)
                            //for (int j = 2; j < c1FlexGrid4.Rows.Count; j++)
                            // 同じ全体順の行数からループさせる
                            for (int j = row; j < c1FlexGrid4.Rows.Count; j++)
                            {

                                //if (c1FlexGrid4.Rows[j][6] != null && c1FlexGrid4.Rows[j][7] != null && c1FlexGrid4.Rows[j][6].ToString() != "" && c1FlexGrid4.Rows[j][7].ToString() != "")
                                if (c1FlexGrid4.Rows[j]["ChousaZentaiJun"] != null && c1FlexGrid4.Rows[j]["ChousaKobetsuJun"] != null
                                    && c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString() != "" && c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString() != "")
                                {
                                    //zentai = c1FlexGrid4.Rows[j][6].ToString();  // 全体順
                                    //kobetsu = c1FlexGrid4.Rows[j][7].ToString(); // 個別順
                                    //zentai = c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString();  // 全体順
                                    //kobetsu = c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString(); // 個別順
                                    zentai = (c1FlexGrid4.Rows[j]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString() : "0");  // 全体順
                                    kobetsu = (c1FlexGrid4.Rows[j]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString() : "0"); // 個別順

                                    float.TryParse(zentai, out zentaiNextF);
                                    float.TryParse(kobetsu, out kobetsuNextF);

                                    //// 重複しているか確認
                                    //if (zentaiF == zentaiNextF && kobetsuF == kobetsuNextF)
                                    //{
                                    //    recordCount += 1;
                                    //    // パフォーマンス確保の為、自分以外に重複するレコードがあった場合に終了する
                                    //    if(recordCount > 1) 
                                    //    { 
                                    //        break;
                                    //    }
                                    //}
                                    //else
                                    //{
                                    //    // 全体順が変わっていれば終わりとする
                                    //    if(zentaiF != zentaiNextF)
                                    //    {
                                    //        break;
                                    //    }
                                    //}

                                    // 重複していた場合、カウント
                                    if (zentaiF == zentaiNextF && kobetsuF == kobetsuNextF)
                                    {
                                        recordCount += 1;
                                    }

                                    // 全体順が変わった、または自分以外に重複するレコードがあった場合に終了する
                                    if (zentaiF != zentaiNextF || recordCount > 1)
                                    {
                                        break;
                                    }
                                }
                            }
                            // 重複レコードがあった場合
                            if (recordCount > 1)
                            {
                                errmessage1 = 1;
                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                                globalErrorFlg = "1";
                                errorFlg = true;
                                // ピンク背景
                                //c1FlexGrid4.GetCellRange(i, 6).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                //c1FlexGrid4.GetCellRange(i, 7).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                c1FlexGrid4.GetCellRange(i, ZentaiJunColIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                c1FlexGrid4.GetCellRange(i, KobetsuJunColIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                //c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "E"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                            else
                            {
                                // クリーム色背景
                                //c1FlexGrid4.GetCellRange(i, 6).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                                //c1FlexGrid4.GetCellRange(i, 7).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                                c1FlexGrid4.GetCellRange(i, ZentaiJunColIndex).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                                c1FlexGrid4.GetCellRange(i, KobetsuJunColIndex).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "N"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                        }

                        // 検索したときの条件 chousaHinmokuSearchWhere 
                        // 検索で絞った場合、DBにいるデータと重複しているとNG
                        if (chousaHinmokuSearchWhere != "")
                        {
                            // 検索で出てきていないデータを取得する
                            string where = "MadoguchiID = '" + MadoguchiID + "' AND not (" + chousaHinmokuSearchWhere + ")";
                            //コンボボックスデータ取得
                            DataTable combodt = GlobalMethod.getData("ChousaKobetsuJun", "ChousaZentaiJun", "ChousaHinmoku", where);

                            if (combodt != null && combodt.Rows.Count > 0)
                            {
                                for (int i = 0; i < combodt.Rows.Count; i++)
                                {
                                    zentai = combodt.Rows[i][0].ToString();
                                    kobetsu = combodt.Rows[i][1].ToString();

                                    float.TryParse(zentai, out zentaiF);
                                    float.TryParse(kobetsu, out kobetsuF);

                                    for (int j = 2; j < c1FlexGrid4.Rows.Count; j++)
                                    {
                                        //if (c1FlexGrid4.Rows[j][6] != null && c1FlexGrid4.Rows[j][7] != null && c1FlexGrid4.Rows[j][6].ToString() != "" && c1FlexGrid4.Rows[j][7].ToString() != "")
                                        if (c1FlexGrid4.Rows[j]["ChousaZentaiJun"] != null && c1FlexGrid4.Rows[j]["ChousaKobetsuJun"] != null
                                            && c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString() != "" && c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString() != "")
                                        {
                                            //zentai = c1FlexGrid4.Rows[j][6].ToString();
                                            //kobetsu = c1FlexGrid4.Rows[j][7].ToString();
                                            //zentai = c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString();
                                            //kobetsu = c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString();
                                            zentai = (c1FlexGrid4.Rows[j]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString() : "0");  // 全体順
                                            kobetsu = (c1FlexGrid4.Rows[j]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString() : "0"); // 個別順

                                            float.TryParse(zentai, out zentaiNextF);
                                            float.TryParse(kobetsu, out kobetsuNextF);

                                            if (zentaiF == zentaiNextF && kobetsuF == kobetsuNextF)
                                            {
                                                errmessage1 = 1;
                                                errorFlg = true;
                                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                                                globalErrorFlg = "1";
                                                // ピンク背景
                                                //c1FlexGrid4.GetCellRange(j, 6).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                                //c1FlexGrid4.GetCellRange(j, 7).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                                c1FlexGrid4.GetCellRange(j, ZentaiJunColIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                                c1FlexGrid4.GetCellRange(j, KobetsuJunColIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                                //c1FlexGrid4[j, 58] = "E" + zeroPadding(c1FlexGrid4[j, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[j, 7].ToString());
                                                c1FlexGrid4.Rows[j]["ColumnSort"] = "E"
                                                                                  + zeroPadding((c1FlexGrid4.Rows[j]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString() : "0"))
                                                                                  + "-"
                                                                                  + zeroPadding((c1FlexGrid4.Rows[j]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString() : "0"))
                                                                                  ;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // 34:地区割コードが空でなく、半角英数字でない場合、正規表現「^[0-9a-zA-Z]+$」
                        // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい。

                        // 35:地区コードが空でなく、半角英数字でない場合、正規表現「^[0-9a-zA-Z]+$」
                        // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい

                        // 13:価格が空でなく、「,」を空文字に置換し、Trimした値が、半角数字でない場合、正規表現「^-?[\d][\d.]*$」
                        // E20337:半角数字で入力してください。赤背景の項目を修正して下さい。

                        // 行番号の取得
                        int TikuWariCodeColIndex = c1FlexGrid4.Cols["ChousaTikuWariCode"].Index;    // 地区割りコード
                        int TikuCodeColIndex = c1FlexGrid4.Cols["ChousaTikuCode"].Index;            // 地区コード
                        int KakakuColIndex = c1FlexGrid4.Cols["ChousaKakaku"].Index;                // 価格
                        int ChousaZaiKouIndex = c1FlexGrid4.Cols["ChousaZaiKou"].Index;             // 材工
                        int ChousaZentaiJunIndex = c1FlexGrid4.Cols["ChousaZentaiJun"].Index;       // 全体順
                        int ChousaKobetsuJunIndex = c1FlexGrid4.Cols["ChousaKobetsuJun"].Index;     // 個別順

                        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                        {
                            // 地区割りコード
                            //if (c1FlexGrid4.Rows[i][34] != null && c1FlexGrid4.Rows[i][34].ToString() != ""
                            //    && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i][34].ToString().Replace(Environment.NewLine, ""), @"^[0-9a-zA-Z]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            if (c1FlexGrid4.Rows[i]["ChousaTikuWariCode"] != null && c1FlexGrid4.Rows[i]["ChousaTikuWariCode"].ToString() != ""
                                && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i]["ChousaTikuWariCode"].ToString().Replace(Environment.NewLine, ""), @"^[0-9a-zA-Z]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            {
                                errmessage2 = 1;
                                errorFlg = true;
                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                                globalErrorFlg = "1";
                                // ピンク背景
                                //c1FlexGrid4.GetCellRange(i, 34).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                c1FlexGrid4.GetCellRange(i, TikuWariCodeColIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                //c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "E"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                            else
                            {
                                // 白背景
                                //c1FlexGrid4.GetCellRange(i, 34).StyleNew.BackColor = Color.White;
                                c1FlexGrid4.GetCellRange(i, TikuWariCodeColIndex).StyleNew.BackColor = Color.White;
                            }
                            // 地区コード
                            //if (c1FlexGrid4.Rows[i][35] != null && c1FlexGrid4.Rows[i][35].ToString() != ""
                            //    && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i][35].ToString().Replace(Environment.NewLine, ""), @"^[0-9a-zA-Z]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            if (c1FlexGrid4.Rows[i]["ChousaTikuCode"] != null && c1FlexGrid4.Rows[i]["ChousaTikuCode"].ToString() != ""
                                && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i]["ChousaTikuCode"].ToString().Replace(Environment.NewLine, ""), @"^[0-9a-zA-Z]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            {
                                errmessage2 = 1;
                                errorFlg = true;
                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                                globalErrorFlg = "1";
                                // ピンク背景
                                //c1FlexGrid4.GetCellRange(i, 35).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                c1FlexGrid4.GetCellRange(i, TikuCodeColIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                //c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "E"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                            else
                            {
                                // 白背景
                                //c1FlexGrid4.GetCellRange(i, 35).StyleNew.BackColor = Color.White;
                                c1FlexGrid4.GetCellRange(i, TikuCodeColIndex).StyleNew.BackColor = Color.White;
                            }
                            // 価格
                            // C1FlexGridの制御で以下のチェックに入らない
                            //if (c1FlexGrid4.Rows[i][13] != null && c1FlexGrid4.Rows[i][13].ToString() != ""
                            //    && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i][13].ToString(), @"^-?[\d][\d.]*$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            if (c1FlexGrid4.Rows[i]["ChousaKakaku"] != null && c1FlexGrid4.Rows[i]["ChousaKakaku"].ToString() != ""
                                && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i]["ChousaKakaku"].ToString(), @"^-?[\d][\d.]*$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                            {
                                errmessage3 = 1;
                                errorFlg = true;
                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                                globalErrorFlg = "1";
                                // ピンク背景
                                //c1FlexGrid4.GetCellRange(i, 13).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                c1FlexGrid4.GetCellRange(i, KakakuColIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                //c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "E"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                            else
                            {
                                // 白背景
                                //c1FlexGrid4.GetCellRange(i, 13).StyleNew.BackColor = Color.White;
                                c1FlexGrid4.GetCellRange(i, KakakuColIndex).StyleNew.BackColor = Color.White;
                            }

                            // 材工
                            // 0:空欄
                            if (c1FlexGrid4.Rows[i]["ChousaZaiKou"] == null || c1FlexGrid4.Rows[i]["ChousaZaiKou"].ToString() == "" || c1FlexGrid4.Rows[i]["ChousaZaiKou"].ToString() == "0")
                            {
                                errmessage4 = 1;
                                errorFlg = true;
                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                                globalErrorFlg = "1";
                                // ピンク背景
                                c1FlexGrid4.GetCellRange(i, ChousaZaiKouIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "E"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                            else
                            {
                                // 必須背景薄黄色
                                c1FlexGrid4.GetCellRange(i, ChousaZaiKouIndex).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                            }
                            // 全体順
                            if (c1FlexGrid4.Rows[i]["ChousaZentaiJun"] == null || c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() == "" || c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() == "0")
                            {
                                errmessage4 = 1;
                                errorFlg = true;
                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                                globalErrorFlg = "1";
                                // ピンク背景
                                c1FlexGrid4.GetCellRange(i, ChousaZentaiJunIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "E"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                            else
                            {
                                // 必須背景薄黄色
                                c1FlexGrid4.GetCellRange(i, ChousaZentaiJunIndex).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                            }
                            // 個別順
                            if (c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] == null || c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() == "" || c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() == "0")
                            {
                                errmessage4 = 1;
                                errorFlg = true;
                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                                globalErrorFlg = "1";
                                // ピンク背景
                                c1FlexGrid4.GetCellRange(i, ChousaKobetsuJunIndex).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "E"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                            else
                            {
                                // 必須背景薄黄色
                                c1FlexGrid4.GetCellRange(i, ChousaKobetsuJunIndex).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                            }
                        }

                        GlobalMethod.outputLogger("Madoguhi button3_InputStatus_Click", "DB比較終了 ID:" + MadoguchiID, "update", "DEBUG");

                        // エラーが無ければ
                        if (errorFlg == false)
                        {
                            try
                            {
                                // 更新前のChousaHinmokuを取得する
                                cmd.CommandText = "SELECT " +
                                "ChousaHinmokuID " +
                                ",HinmokuRyakuBushoCD " +
                                ",HinmokuChousainCD " +
                                ",HinmokuRyakuBushoFuku1CD " +
                                ",HinmokuFukuChousainCD1 " +
                                ",HinmokuRyakuBushoFuku2CD " +
                                ",HinmokuFukuChousainCD2 " +
                                "FROM ChousaHinmoku " +
                                "WHERE MadoguchiID = '" + MadoguchiID + "' ";

                                var sda = new SqlDataAdapter(cmd);
                                beforeChousaHinmokuDT.Clear();
                                sda.Fill(beforeChousaHinmokuDT);

                                // １．調査品目の削除Key（ChousaHinmokuIDをカンマ区切りで連結したデータ）があれば削除
                                if (deleteChousaHinmokuIDs != "")
                                {
                                    // 削除
                                    // 調査品目全削除
                                    cmd.CommandText = "DELETE FROM ChousaHinmoku " +
                                        "WHERE ChousaHinmokuID in (" + deleteChousaHinmokuIDs + ") AND MadoguchiID = '" + MadoguchiID + "' ";
                                    cmd.ExecuteNonQuery();

                                    // T_History に登録する文言は256文字までなので、分割していれないと桁あふれする
                                    //writeHistory("調査品目が削除されました。調査品目ID in (" + deleteChousaHinmokuIDs + ")");

                                    string[] deleteID = deleteChousaHinmokuIDs.Split(',');

                                    updmessage3 = 1; // 削除

                                    for (int i = 0; i < deleteID.Length; i++)
                                    {
                                        writeHistory("調査品目が削除されました。調査品目ID = " + deleteID[i]);
                                    }

                                }

                                // ２．c1FlexGrid4 の 57:0:Insert/1:Select/2:Update があり、それで新規か更新、または処理なしを切り分ける
                                // ２－１．まずは新規を処理する

                                // ソートが効いてない、、、
                                //c1FlexGrid4.Cols[57].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending; // 57:0:Insert/1:Select/2:Update の昇順設定
                                //c1FlexGrid4.Cols[55].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending; // 55:ChousaHinmokuID の昇順設定
                                //c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 57, 55);  // 設定した内容で、ソートする

                                // 更新日付をあらかじめ所得しておく
                                sysDateTimeStr = DateTime.Now.ToString();

                                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                                {
                                    // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                    //c1FlexGrid4[i, 58] = "N" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                                    c1FlexGrid4.Rows[i]["ColumnSort"] = "N"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                      ;
                                    // 0:Insertを処理する
                                    //if (c1FlexGrid4.Rows[i][57] != null && c1FlexGrid4.Rows[i][57].ToString() == "0")
                                    if (c1FlexGrid4.Rows[i]["Mode"] != null && c1FlexGrid4.Rows[i]["Mode"].ToString() == "0")
                                    {
                                        //if (cnt >= 100)
                                        if (cnt >= i_RecodeCountMax)
                                        {
                                            // 登録をまとめて行う
                                            cmd.CommandText = insertQuery + valuesText;

                                            cmd.ExecuteNonQuery();
                                            // 追加メッセージ
                                            updmessage1 = 1;
                                            cnt = 0;
                                            valuesText = "";

                                        }
                                        cnt += 1;

                                        if (valuesText != "")
                                        {
                                            valuesText += ",";
                                        }

                                        valuesText += "(" +
                                        //" '" + c1FlexGrid4.Rows[i][55] + "' " +
                                        //",'" + MadoguchiID + "' " +
                                        //",'" + c1FlexGrid4.Rows[i][6] + "' " +   // 全体順
                                        //",'" + c1FlexGrid4.Rows[i][7] + "' " +   // 個別順
                                        //",'" + c1FlexGrid4.Rows[i][8] + "' " +   // 材工
                                        //",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][9].ToString(), 0, 0) + "' " +   // 品目
                                        //",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][10].ToString(), 0, 0) + "' " +  // 規格
                                        //",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][11].ToString(), 0, 0) + "' " +  // 単位
                                        //",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][12].ToString(), 0, 0) + "' ";   // 参考質量
                                        " '" + c1FlexGrid4.Rows[i]["ChousaHinmokuID2"] + "' " +
                                        ",'" + MadoguchiID + "' " +
                                        ",'" + c1FlexGrid4.Rows[i]["ChousaZentaiJun"] + "' " +   // 全体順
                                        ",'" + c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] + "' " +   // 個別順
                                        ",'" + c1FlexGrid4.Rows[i]["ChousaZaiKou"] + "' " +   // 材工
                                        ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmei"].ToString(), 0, 0) + "' " +   // 品目
                                        ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaKikaku"].ToString(), 0, 0) + "' " +  // 規格
                                        ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTanka"].ToString(), 0, 0) + "' " +  // 単位
                                        ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaSankouShitsuryou"].ToString(), 0, 0) + "' ";   // 参考質量

                                        // 価格
                                        //if (c1FlexGrid4.Rows[i][13] != null && c1FlexGrid4.Rows[i][13].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaKakaku"] != null && c1FlexGrid4.Rows[i]["ChousaKakaku"].ToString() != "")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][13] + "' ";
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["ChousaKakaku"] + "' ";
                                        }
                                        else
                                        {
                                            valuesText += ",null ";
                                        }

                                        // 中止
                                        //if (c1FlexGrid4.Rows[i][14] != null && c1FlexGrid4.Rows[i][14].ToString() == "True")
                                        if (c1FlexGrid4.Rows[i]["ChousaChuushi"] != null && c1FlexGrid4.Rows[i]["ChousaChuushi"].ToString() == "True")
                                        {
                                            valuesText += ",1 ";
                                        }
                                        else
                                        {
                                            valuesText += ",0 ";
                                        }

                                        //valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][15].ToString(), 0, 0) + "' " +  // 報告備考
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][16].ToString(), 0, 0) + "' " +            // 依頼備考
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][17].ToString(), 0, 0) + "' " +            // 単価適用地域
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][18].ToString(), 0, 0) + "' " +            // 図面番号
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][19].ToString(), 0, 0) + "' " +            // 数量
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][20].ToString(), 0, 0) + "' " +            // 見積先
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][21].ToString(), 0, 0) + "' ";             // ベースメーカー
                                        valuesText += ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBikou2"].ToString(), 0, 0) + "' " +  // 報告備考
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBikou"].ToString(), 0, 0) + "' " +             // 依頼備考
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTankaTekiyouTiku"].ToString(), 0, 0) + "' " +  // 単価適用地域
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaZumenNo"].ToString(), 0, 0) + "' " +           // 図面番号
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaSuuryou"].ToString(), 0, 0) + "' " +           // 数量
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaMitsumorisaki"].ToString(), 0, 0) + "' " +     // 見積先
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBaseMakere"].ToString(), 0, 0) + "' ";         // ベースメーカー

                                        // ベース単価
                                        //if (c1FlexGrid4.Rows[i][22] != null && c1FlexGrid4.Rows[i][22].ToString() != "" && c1FlexGrid4.Rows[i][22].ToString() != "0")
                                        if (c1FlexGrid4.Rows[i]["ChousaBaseTanka"] != null && c1FlexGrid4.Rows[i]["ChousaBaseTanka"].ToString() != "" && c1FlexGrid4.Rows[i]["ChousaBaseTanka"].ToString() != "0")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][22] + "' ";
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["ChousaBaseTanka"] + "' ";
                                        }
                                        else
                                        {
                                            valuesText += ",'           0.00' ";
                                        }

                                        // 掛率
                                        //if (c1FlexGrid4.Rows[i][23] != null && c1FlexGrid4.Rows[i][23].ToString() != "" && c1FlexGrid4.Rows[i][23].ToString() != "0")
                                        if (c1FlexGrid4.Rows[i]["ChousaKakeritsu"] != null && c1FlexGrid4.Rows[i]["ChousaKakeritsu"].ToString() != "" && c1FlexGrid4.Rows[i]["ChousaKakeritsu"].ToString() != "0")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][23] + "' ";
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["ChousaKakeritsu"] + "' ";
                                        }
                                        else
                                        {
                                            valuesText += ",'  0.00' ";
                                        }

                                        //valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][24].ToString(), 0, 0) + "' " +  // 属性
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][25].ToString(), 0, 0) + "' ";             // 前回単位
                                        valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaObiMei"].ToString(), 0, 0) + "' " +  // 属性
                                            ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaZenkaiTani"].ToString(), 0, 0) + "' ";         // 前回単位

                                        // 前回価格
                                        //if (c1FlexGrid4.Rows[i][26] != null && c1FlexGrid4.Rows[i][26].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaZenkaiKakaku"] != null && c1FlexGrid4.Rows[i]["ChousaZenkaiKakaku"].ToString() != "")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][26] + "' ";
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["ChousaZenkaiKakaku"] + "' ";
                                        }
                                        else
                                        {
                                            valuesText += ",0 ";
                                        }

                                        // 発注者提供単価
                                        //if (c1FlexGrid4.Rows[i][27] != null && c1FlexGrid4.Rows[i][27].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaSankouti"] != null && c1FlexGrid4.Rows[i]["ChousaSankouti"].ToString() != "")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][27] + "' ";
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["ChousaSankouti"] + "' ";
                                        }
                                        else
                                        {
                                            valuesText += ",0 ";
                                        }

                                        //valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][28].ToString(), 0, 0) + "' " +  // 品目情報1
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][29].ToString(), 0, 0) + "' " +  // 品目情報2
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][30].ToString(), 0, 0) + "' " +  // 前回質量
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][31].ToString(), 0, 0) + "' " +  // メモ1
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][32].ToString(), 0, 0) + "' " +  // メモ2
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][33].ToString(), 0, 0) + "' " +  // 発注品目コード
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][34].ToString(), 0, 0) + "' " +  // 地区割コード
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][35].ToString(), 0, 0) + "' " +  // 地区コード
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][36].ToString(), 0, 0) + "' ";   // 地区名
                                        valuesText += ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmokuJouhou1"].ToString(), 0, 0) + "' " +  // 品目情報1
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmokuJouhou2"].ToString(), 0, 0) + "' " +            // 品目情報2
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaFukuShizai"].ToString(), 0, 0) + "' " +                // 前回質量
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBunrui"].ToString(), 0, 0) + "' " +                    // メモ1
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaMemo2"].ToString(), 0, 0) + "' " +                     // メモ2
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTankaCD1"].ToString(), 0, 0) + "' " +                  // 発注品目コード
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTikuWariCode"].ToString(), 0, 0) + "' " +              // 地区割コード
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTikuCode"].ToString(), 0, 0) + "' " +                  // 地区コード
                                            ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTikuMei"].ToString(), 0, 0) + "' ";                    // 地区名

                                        // 少額案件[10万/100万]
                                        //if (c1FlexGrid4.Rows[i][37] != null && c1FlexGrid4.Rows[i][37].ToString() == "True")
                                        if (c1FlexGrid4.Rows[i]["ChousaShougaku"] != null && c1FlexGrid4.Rows[i]["ChousaShougaku"].ToString() == "True")
                                        {
                                            valuesText += ",1 ";
                                        }
                                        else
                                        {
                                            valuesText += ",0 ";
                                        }
                                        // Web建
                                        //if (c1FlexGrid4.Rows[i][38] != null && c1FlexGrid4.Rows[i][38].ToString() == "True")
                                        if (c1FlexGrid4.Rows[i]["ChousaWebKen"] != null && c1FlexGrid4.Rows[i]["ChousaWebKen"].ToString() == "True")
                                        {
                                            valuesText += ",1 ";
                                        }
                                        else
                                        {
                                            valuesText += ",0 ";
                                        }

                                        //valuesText += ",'" + c1FlexGrid4.Rows[i][39] + "' " +  // 根拠関連コード
                                        //                                                       // リンク先アイコン
                                        //    ",'" + c1FlexGrid4.Rows[i][41] + "' ";   // リンク先パス
                                        valuesText += ",N'" + c1FlexGrid4.Rows[i]["ChousaKonkyoCode"] + "' " +       // 根拠関連コード
                                                                                                                     // リンク先アイコン
                                                      ",N'" + c1FlexGrid4.Rows[i]["ChousaLinkSakliFolder"] + "' ";   // リンク先パス


                                        // 調査担当部所
                                        //if (c1FlexGrid4.Rows[i][42] != null && c1FlexGrid4.Rows[i][42].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"] != null && c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"].ToString() != "")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][42] + "' ";
                                            //afterRyakuBushoCD = c1FlexGrid4.Rows[i][42].ToString();
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"] + "' ";
                                            afterRyakuBushoCD = c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"].ToString();
                                        }
                                        else
                                        {
                                            valuesText += ",null ";
                                            afterRyakuBushoCD = "";
                                        }
                                        // 調査担当者
                                        //if (c1FlexGrid4.Rows[i][43] != null && c1FlexGrid4.Rows[i][43].ToString() != "" && c1FlexGrid4.Rows[i][43].ToString() != "0")
                                        if (c1FlexGrid4.Rows[i]["HinmokuChousainCD"] != null && c1FlexGrid4.Rows[i]["HinmokuChousainCD"].ToString() != "" && c1FlexGrid4.Rows[i]["HinmokuChousainCD"].ToString() != "0")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][43] + "' ";
                                            //afterChousainCD = c1FlexGrid4.Rows[i][43].ToString();
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["HinmokuChousainCD"] + "' ";
                                            afterChousainCD = c1FlexGrid4.Rows[i]["HinmokuChousainCD"].ToString();
                                        }
                                        else
                                        {
                                            valuesText += ",null ";
                                            afterChousainCD = "";
                                        }
                                        // 副調査担当部所1
                                        //if (c1FlexGrid4.Rows[i][44] != null && c1FlexGrid4.Rows[i][44].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"] != null && c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"].ToString() != "")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][44] + "' ";
                                            //afterFukuRyakuBushoCD1 = c1FlexGrid4.Rows[i][44].ToString();
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"] + "' ";
                                            afterFukuRyakuBushoCD1 = c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"].ToString();
                                        }
                                        else
                                        {
                                            valuesText += ",null ";
                                            afterFukuRyakuBushoCD1 = "";
                                        }
                                        // 副調査担当者1
                                        //if (c1FlexGrid4.Rows[i][45] != null && c1FlexGrid4.Rows[i][45].ToString() != "" && c1FlexGrid4.Rows[i][45].ToString() != "0")
                                        if (c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"] != null && c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"].ToString() != "" && c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"].ToString() != "0")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][45] + "' ";
                                            //afterFukuChousainCD1 = c1FlexGrid4.Rows[i][45].ToString();
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"] + "' ";
                                            afterFukuChousainCD1 = c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"].ToString();
                                        }
                                        else
                                        {
                                            valuesText += ",null ";
                                            afterFukuChousainCD1 = "";
                                        }
                                        // 副調査担当部所2
                                        //if (c1FlexGrid4.Rows[i][46] != null && c1FlexGrid4.Rows[i][46].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"] != null && c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"].ToString() != "")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][46] + "' ";
                                            //afterFukuRyakuBushoCD2 = c1FlexGrid4.Rows[i][46].ToString();
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"] + "' ";
                                            afterFukuRyakuBushoCD2 = c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"].ToString();
                                        }
                                        else
                                        {
                                            valuesText += ",null ";
                                            afterFukuRyakuBushoCD2 = "";
                                        }
                                        // 副調査担当者2
                                        //if (c1FlexGrid4.Rows[i][47] != null && c1FlexGrid4.Rows[i][47].ToString() != "" && c1FlexGrid4.Rows[i][47].ToString() != "0")
                                        if (c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"] != null && c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"].ToString() != "" && c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"].ToString() != "0")
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][47] + "' ";
                                            //afterFukuChousainCD2 = c1FlexGrid4.Rows[i][47].ToString();
                                            valuesText += ",'" + c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"] + "' ";
                                            afterFukuChousainCD2 = c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"].ToString();
                                        }
                                        else
                                        {
                                            valuesText += ",null ";
                                            afterFukuChousainCD1 = "";
                                        }

                                        //valuesText += ",'" + c1FlexGrid4.Rows[i][48] + "' " +  // 報告数
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][49].ToString(), 0, 0) + "' " +  // 報告ランク
                                        //    //",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][50].ToString(), 0, 0) + "' " +  // 依頼数
                                        //    ",'" + c1FlexGrid4.Rows[i][50] + "' " +  // 依頼数
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][51].ToString(), 0, 0) + "' " +  // 依頼ランク
                                        //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][52].ToString(), 0, 0) + "' ";  // 締切日

                                        //valuesText += ",'" + c1FlexGrid4.Rows[i][48] + "' ";  // 報告数
                                        valuesText += ",'" + c1FlexGrid4.Rows[i]["ChousaHoukokuHonsuu"] + "' ";  // 報告数
                                                                                                                 //if (c1FlexGrid4.Rows[i][49] != null && c1FlexGrid4.Rows[i][49].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaHoukokuRank"] != null && c1FlexGrid4.Rows[i]["ChousaHoukokuRank"].ToString() != "")
                                        {
                                            //valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][49].ToString(), 0, 0) + "' ";  // 報告ランク
                                            valuesText += ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHoukokuRank"].ToString(), 0, 0) + "' ";  // 報告ランク
                                        }
                                        else
                                        {
                                            valuesText += ",'' ";  // 報告ランク
                                        }

                                        //valuesText += ",'" + c1FlexGrid4.Rows[i][50] + "' ";  // 依頼数
                                        valuesText += ",'" + c1FlexGrid4.Rows[i]["ChousaIraiHonsuu"] + "' ";  // 依頼数

                                        //if (c1FlexGrid4.Rows[i][51] != null && c1FlexGrid4.Rows[i][51].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaIraiRank"] != null && c1FlexGrid4.Rows[i]["ChousaIraiRank"].ToString() != "")
                                        {
                                            //valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][51].ToString(), 0, 0) + "' ";  // 依頼ランク
                                            valuesText += ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaIraiRank"].ToString(), 0, 0) + "' ";  // 依頼ランク
                                        }
                                        else
                                        {
                                            valuesText += ",'' ";  // 依頼ランク
                                        }

                                        //valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][52].ToString(), 0, 0) + "' ";  // 締切日
                                        //valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmokuShimekiribi"].ToString(), 0, 0) + "' ";  // 締切日
                                        if (c1FlexGrid4.Rows[i]["ChousaHinmokuShimekiribi"] != null)
                                        {
                                            valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmokuShimekiribi"].ToString(), 0, 0) + "' ";  // 締切日
                                        }
                                        else
                                        {
                                            valuesText += ",null ";  // 締切日
                                        }

                                        // 報告済
                                        //if (c1FlexGrid4.Rows[i][53] != null && c1FlexGrid4.Rows[i][53].ToString() == "True")
                                        if (c1FlexGrid4.Rows[i]["ChousaHoukokuzumi"] != null && c1FlexGrid4.Rows[i]["ChousaHoukokuzumi"].ToString() == "True")
                                        {
                                            valuesText += ",1 ";
                                        }
                                        else
                                        {
                                            valuesText += ",0 ";
                                        }

                                        valuesText += ",'0' " +                          // 削除フラグ
                                            ",'" + sysDateTimeStr + "'" +                // 登録日時
                                            ",N'" + UserInfos[0] + "' " +                 // 登録ユーザー
                                            ",'" + pgmName + methodName + "' " +  // 登録プログラム
                                            ",'" + sysDateTimeStr + "'" +                // 更新日時
                                            ",N'" + UserInfos[0] + "' " +                 // 更新ユーザー
                                            ",'" + pgmName + methodName + "' ";  // 更新プログラム


                                        // 進捗状況
                                        // 価格 が入っている場合、50:担当者済とする
                                        //if (c1FlexGrid4.Rows[i][13] != null && c1FlexGrid4.Rows[i][13].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaKakaku"] != null && c1FlexGrid4.Rows[i]["ChousaKakaku"].ToString() != "")
                                        {
                                            valuesText += ",'50' ";
                                            // 進捗状況
                                            //c1FlexGrid4.Rows[i][56] = "50";
                                            c1FlexGrid4.Rows[i]["ChousaShinchokuJoukyou"] = "50";
                                            if (MadoguchiHoukokuzumi != "1")
                                            {
                                                // 進捗状況
                                                //c1FlexGrid4.Rows[i][5] = "7";
                                                c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "7";
                                                if (c1FlexGrid4.Rows[i]["ChousaChuushi"] != null && "True".Equals(c1FlexGrid4.Rows[i]["ChousaChuushi"].ToString()))
                                                {
                                                    c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "6";
                                                }
                                            }
                                        }
                                        else
                                        {
                                            //valuesText += ",'" + c1FlexGrid4.Rows[i][56] + "' ";       // 進捗状況
                                            valuesText += ",'20' ";

                                            // 価格 が入っていない場合、20:調査開始とする
                                            // 進捗状況
                                            //c1FlexGrid4.Rows[i][56] = "20";
                                            c1FlexGrid4.Rows[i]["ChousaShinchokuJoukyou"] = "20";
                                            if (MadoguchiHoukokuzumi != "1")
                                            {
                                                //DateTime dateTime = DateTime.Parse(c1FlexGrid4[i, 52].ToString());
                                                if (c1FlexGrid4.Rows[i]["ChousaHinmokuShimekiribi"] != null)
                                                {
                                                    DateTime dateTime = DateTime.Parse(c1FlexGrid4.Rows[i]["ChousaHinmokuShimekiribi"].ToString());
                                                    if (c1FlexGrid4.Rows[i]["ChousaChuushi"] != null && "True".Equals(c1FlexGrid4.Rows[i]["ChousaChuushi"].ToString()))
                                                    {
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "6";
                                                    }
                                                    else if (dateTime < DateTime.Today)
                                                    {
                                                        // 締切日経過
                                                        //c1FlexGrid4.Rows[i][5] = "1";
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "1";
                                                    }
                                                    else if (dateTime < DateTime.Today.AddDays(3))
                                                    {
                                                        // 締切日が3日以内、かつ2次検証が完了していない
                                                        //c1FlexGrid4.Rows[i][5] = "2";
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "2";
                                                    }
                                                    else if (dateTime < DateTime.Today.AddDays(7))
                                                    {
                                                        // 締切日が1週間以内、かつ2次検証が完了していない
                                                        //c1FlexGrid4.Rows[i][5] = "3";
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "3";
                                                    }
                                                    else
                                                    {
                                                        //c1FlexGrid4.Rows[i][5] = "4";
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "4";
                                                    }
                                                    //// 進捗状況
                                                    //c1FlexGrid4.Rows[i][5] = "7";
                                                }
                                                else
                                                {
                                                    c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "4";
                                                }
                                            }
                                        }
                                        valuesText += ")";

                                        // 履歴に登録する
                                        historyMessage = "調査員が追加されました。";
                                        if (afterChousainCD != "")
                                        {
                                            writeChousaHinmokuHistory(historyMessage, "", "", afterRyakuBushoCD, afterChousainCD);
                                        }
                                        historyMessage = "副調査員1が追加されました。";
                                        if (afterFukuChousainCD1 != "")
                                        {
                                            writeChousaHinmokuHistory(historyMessage, "", "", afterFukuRyakuBushoCD1, afterFukuChousainCD1);
                                        }
                                        historyMessage = "副調査員2が追加されました。";
                                        if (afterFukuChousainCD2 != "")
                                        {
                                            writeChousaHinmokuHistory(historyMessage, "", "", afterFukuRyakuBushoCD2, afterFukuChousainCD2);
                                        }
                                    }
                                }
                                // valuesTextが空でなければinsertを行う
                                if (valuesText != "")
                                {
                                    // 登録を行う
                                    cmd.CommandText = insertQuery + valuesText;
                                    cmd.ExecuteNonQuery();
                                    // 追加メッセージ
                                    updmessage1 = 1;
                                }


                                // ２－２．更新を処理する

                                // 更新日付をあらかじめ所得しておく
                                sysDateTimeStr = DateTime.Now.ToString();

                                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                                {
                                    // 2:Updateを処理する
                                    //if (c1FlexGrid4.Rows[i][57] != null && c1FlexGrid4.Rows[i][57].ToString() == "2")
                                    if (c1FlexGrid4.Rows[i]["Mode"] != null && c1FlexGrid4.Rows[i]["Mode"].ToString() == "2")
                                    {
                                        cmd.CommandText = "UPDATE ChousaHinmoku SET ";

                                        //cmd.CommandText += " ChousaZentaiJun = '" + c1FlexGrid4.Rows[i][6] + "' " +   // 全体順
                                        //    ",ChousaKobetsuJun = '" + c1FlexGrid4.Rows[i][7] + "' " +   // 個別順
                                        //    ",ChousaZaiKou = '" + c1FlexGrid4.Rows[i][8] + "' " +   // 材工
                                        //    ",ChousaHinmei = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][9].ToString(), 0, 0) + "' " +   // 品目
                                        //    ",ChousaKikaku = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][10].ToString(), 0, 0) + "' " +  // 規格
                                        //    ",ChousaTanka = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][11].ToString(), 0, 0) + "' " +  // 単位
                                        //    ",ChousaSankouShitsuryou = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][12].ToString(), 0, 0) + "' ";   // 参考質量
                                        cmd.CommandText += " ChousaZentaiJun = '" + c1FlexGrid4.Rows[i]["ChousaZentaiJun"] + "' " +                                             // 全体順
                                            ",ChousaKobetsuJun = '" + c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] + "' " +                                                          // 個別順
                                            ",ChousaZaiKou = '" + c1FlexGrid4.Rows[i]["ChousaZaiKou"] + "' " +                                                                  // 材工
                                            ",ChousaHinmei = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmei"].ToString(), 0, 0) + "' " +                     // 品目
                                            ",ChousaKikaku = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaKikaku"].ToString(), 0, 0) + "' " +                     // 規格
                                            ",ChousaTanka = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTanka"].ToString(), 0, 0) + "' " +                       // 単位
                                            ",ChousaSankouShitsuryou = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaSankouShitsuryou"].ToString(), 0, 0) + "' ";  // 参考質量

                                        // 価格
                                        //if (c1FlexGrid4.Rows[i][13] != null && c1FlexGrid4.Rows[i][13].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaKakaku"] != null && c1FlexGrid4.Rows[i]["ChousaKakaku"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",ChousaKakaku = '" + c1FlexGrid4.Rows[i][13] + "' ";
                                            cmd.CommandText += ",ChousaKakaku = '" + c1FlexGrid4.Rows[i]["ChousaKakaku"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaKakaku = null ";
                                        }

                                        // 中止
                                        //if (c1FlexGrid4.Rows[i][14] != null && c1FlexGrid4.Rows[i][14].ToString() == "True")
                                        if (c1FlexGrid4.Rows[i]["ChousaChuushi"] != null && c1FlexGrid4.Rows[i]["ChousaChuushi"].ToString() == "True")
                                        {
                                            cmd.CommandText += ",ChousaChuushi = 1 ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaChuushi = 0 ";
                                        }

                                        //cmd.CommandText += ",ChousaBikou2 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][15].ToString(), 0, 0) + "' " +  // 報告備考
                                        //    ",ChousaBikou = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][16].ToString(), 0, 0) + "' " +  // 依頼備考
                                        //    ",ChousaTankaTekiyouTiku = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][17].ToString(), 0, 0) + "' " +  // 単価適用地域
                                        //    ",ChousaZumenNo = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][18].ToString(), 0, 0) + "' " +  // 図面番号
                                        //    ",ChousaSuuryou = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][19].ToString(), 0, 0) + "' " +  // 数量
                                        //    ",ChousaMitsumorisaki = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][20].ToString(), 0, 0) + "' " +  // 見積先
                                        //    ",ChousaBaseMakere = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][21].ToString(), 0, 0) + "' " +  // ベースメーカー
                                        //    ",ChousaBaseTanka = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][22].ToString(), 0, 0) + "' " +  // ベース単位
                                        //    ",ChousaKakeritsu = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][23].ToString(), 0, 0) + "' " +  // 掛率
                                        //    ",ChousaObiMei = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][24].ToString(), 0, 0) + "' " +  // 属性
                                        //    ",ChousaZenkaiTani = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][25].ToString(), 0, 0) + "' " +  // 前回単位
                                        //    ",ChousaZenkaiKakaku = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][26].ToString(), 0, 0) + "' " +  // 前回価格
                                        //    ",ChousaSankouti = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][27].ToString(), 0, 0) + "' " +  // 発注者提供単価
                                        //    ",ChousaHinmokuJouhou1 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][28].ToString(), 0, 0) + "' " +  // 品目情報1
                                        //    ",ChousaHinmokuJouhou2 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][29].ToString(), 0, 0) + "' " +  // 品目情報2
                                        //    ",ChousaFukuShizai = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][30].ToString(), 0, 0) + "' " +  // 前回質量
                                        //    ",ChousaBunrui = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][31].ToString(), 0, 0) + "' " +  // メモ1
                                        //    ",ChousaMemo2 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][32].ToString(), 0, 0) + "' " +  // メモ2
                                        //    ",ChousaTankaCD1 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][33].ToString(), 0, 0) + "' " +  // 発注品目コード
                                        //    ",ChousaTikuWariCode = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][34].ToString(), 0, 0) + "' " +  // 地区割コード
                                        //    ",ChousaTikuCode = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][35].ToString(), 0, 0) + "' " +  // 地区コード
                                        //    ",ChousaTikuMei = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][36].ToString(), 0, 0) + "' ";   // 地区名
                                        cmd.CommandText += ",ChousaBikou2 = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBikou2"].ToString(), 0, 0) + "' " +          // 報告備考
                                            ",ChousaBikou = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBikou"].ToString(), 0, 0) + "' " +                           // 依頼備考
                                            ",ChousaTankaTekiyouTiku = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTankaTekiyouTiku"].ToString(), 0, 0) + "' " +     // 単価適用地域
                                            ",ChousaZumenNo = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaZumenNo"].ToString(), 0, 0) + "' " +                       // 図面番号
                                            ",ChousaSuuryou = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaSuuryou"].ToString(), 0, 0) + "' " +                       // 数量
                                            ",ChousaMitsumorisaki = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaMitsumorisaki"].ToString(), 0, 0) + "' " +           // 見積先
                                            ",ChousaBaseMakere = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBaseMakere"].ToString(), 0, 0) + "' " +                 // ベースメーカー
                                            //",ChousaBaseTanka = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBaseTanka"].ToString(), 0, 0) + "' " +                   // ベース単位
                                            ",ChousaBaseTanka = '" + GlobalMethod.ChangeSqlText((c1FlexGrid4.Rows[i]["ChousaBaseTanka"] != null ? c1FlexGrid4.Rows[i]["ChousaBaseTanka"].ToString() : "0"), 0, 0) + "' " + // ベース単位
                                            //",ChousaKakeritsu = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaKakeritsu"].ToString(), 0, 0) + "' " +                   // 掛率
                                            ",ChousaKakeritsu = '" + GlobalMethod.ChangeSqlText((c1FlexGrid4.Rows[i]["ChousaKakeritsu"] != null ? c1FlexGrid4.Rows[i]["ChousaKakeritsu"].ToString() : ""), 0, 0) + "' " + // 掛率
                                            ",ChousaObiMei = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaObiMei"].ToString(), 0, 0) + "' " +                         // 属性
                                            ",ChousaZenkaiTani = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaZenkaiTani"].ToString(), 0, 0) + "' " +                 // 前回単位
                                            //",ChousaZenkaiKakaku = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaZenkaiKakaku"].ToString(), 0, 0) + "' " +             // 前回価格
                                            ",ChousaZenkaiKakaku = '" + GlobalMethod.ChangeSqlText((c1FlexGrid4.Rows[i]["ChousaZenkaiKakaku"] != null ? c1FlexGrid4.Rows[i]["ChousaZenkaiKakaku"].ToString() : "0"), 0, 0) + "' " + // 前回価格
                                            //",ChousaSankouti = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaSankouti"].ToString(), 0, 0) + "' " +                     // 発注者提供単価
                                            ",ChousaSankouti = '" + GlobalMethod.ChangeSqlText((c1FlexGrid4.Rows[i]["ChousaSankouti"] != null ? c1FlexGrid4.Rows[i]["ChousaSankouti"].ToString() : "0"), 0, 0) + "' " + // 発注者提供単価
                                            ",ChousaHinmokuJouhou1 = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmokuJouhou1"].ToString(), 0, 0) + "' " +         // 品目情報1
                                            ",ChousaHinmokuJouhou2 = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmokuJouhou2"].ToString(), 0, 0) + "' " +         // 品目情報2
                                            ",ChousaFukuShizai = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaFukuShizai"].ToString(), 0, 0) + "' " +                 // 前回質量
                                            ",ChousaBunrui = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaBunrui"].ToString(), 0, 0) + "' " +                         // メモ1
                                            ",ChousaMemo2 = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaMemo2"].ToString(), 0, 0) + "' " +                           // メモ2
                                            ",ChousaTankaCD1 = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTankaCD1"].ToString(), 0, 0) + "' " +                     // 発注品目コード
                                            ",ChousaTikuWariCode = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTikuWariCode"].ToString(), 0, 0) + "' " +             // 地区割コード
                                            ",ChousaTikuCode = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTikuCode"].ToString(), 0, 0) + "' " +                     // 地区コード
                                            ",ChousaTikuMei = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTikuMei"].ToString(), 0, 0) + "' ";                        // 地区名

                                        // 少額案件[10万/100万]
                                        //if (c1FlexGrid4.Rows[i][37] != null && c1FlexGrid4.Rows[i][37].ToString() == "True")
                                        if (c1FlexGrid4.Rows[i]["ChousaShougaku"] != null && c1FlexGrid4.Rows[i]["ChousaShougaku"].ToString() == "True")
                                        {
                                            cmd.CommandText += ",ChousaShougaku = 1 ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaShougaku = 0 ";
                                        }
                                        // Web建
                                        //if (c1FlexGrid4.Rows[i][38] != null && c1FlexGrid4.Rows[i][38].ToString() == "True")
                                        if (c1FlexGrid4.Rows[i]["ChousaWebKen"] != null && c1FlexGrid4.Rows[i]["ChousaWebKen"].ToString() == "True")
                                        {
                                            cmd.CommandText += ",ChousaWebKen = 1 ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaWebKen = 0 ";
                                        }

                                        //cmd.CommandText += ",ChousaKonkyoCode = '" + c1FlexGrid4.Rows[i][39] + "' " +  // 根拠関連コード
                                        //                                                                               // リンク先アイコン
                                        //    ",ChousaLinkSakli = '" + c1FlexGrid4.Rows[i][41] + "' ";   // リンク先パス
                                        cmd.CommandText += ",ChousaKonkyoCode = N'" + c1FlexGrid4.Rows[i]["ChousaKonkyoCode"] + "' " +   // 根拠関連コード
                                                                                                                                         // リンク先アイコン
                                            ",ChousaLinkSakli = N'" + c1FlexGrid4.Rows[i]["ChousaLinkSakliFolder"] + "' ";               // リンク先パス

                                        // 調査担当部所
                                        //if (c1FlexGrid4.Rows[i][42] != null && c1FlexGrid4.Rows[i][42].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"] != null && c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",HinmokuRyakuBushoCD = '" + c1FlexGrid4.Rows[i][42] + "' ";
                                            cmd.CommandText += ",HinmokuRyakuBushoCD = '" + c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",HinmokuRyakuBushoCD = null ";
                                        }
                                        // 調査担当者
                                        //if (c1FlexGrid4.Rows[i][43] != null && c1FlexGrid4.Rows[i][43].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuChousainCD"] != null && c1FlexGrid4.Rows[i]["HinmokuChousainCD"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",HinmokuChousainCD = '" + c1FlexGrid4.Rows[i][43] + "' ";
                                            cmd.CommandText += ",HinmokuChousainCD = '" + c1FlexGrid4.Rows[i]["HinmokuChousainCD"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",HinmokuChousainCD = null ";
                                        }
                                        // 副調査担当部所1
                                        //if (c1FlexGrid4.Rows[i][44] != null && c1FlexGrid4.Rows[i][44].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"] != null && c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",HinmokuRyakuBushoFuku1CD = '" + c1FlexGrid4.Rows[i][44] + "' ";
                                            cmd.CommandText += ",HinmokuRyakuBushoFuku1CD = '" + c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",HinmokuRyakuBushoFuku1CD = null ";
                                        }
                                        // 副調査担当者1
                                        //if (c1FlexGrid4.Rows[i][45] != null && c1FlexGrid4.Rows[i][45].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"] != null && c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",HinmokuFukuChousainCD1 = '" + c1FlexGrid4.Rows[i][45] + "' ";
                                            cmd.CommandText += ",HinmokuFukuChousainCD1 = '" + c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",HinmokuFukuChousainCD1 = null ";
                                        }
                                        // 副調査担当部所2
                                        //if (c1FlexGrid4.Rows[i][46] != null && c1FlexGrid4.Rows[i][46].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"] != null && c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",HinmokuRyakuBushoFuku2CD = '" + c1FlexGrid4.Rows[i][46] + "' ";
                                            cmd.CommandText += ",HinmokuRyakuBushoFuku2CD = '" + c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",HinmokuRyakuBushoFuku2CD = null ";
                                        }
                                        // 副調査担当者2
                                        //if (c1FlexGrid4.Rows[i][47] != null && c1FlexGrid4.Rows[i][47].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"] != null && c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",HinmokuFukuChousainCD2 = '" + c1FlexGrid4.Rows[i][47] + "' ";
                                            cmd.CommandText += ",HinmokuFukuChousainCD2 = '" + c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",HinmokuFukuChousainCD2 = null ";
                                        }

                                        //cmd.CommandText += ",ChousaHoukokuHonsuu = '" + c1FlexGrid4.Rows[i][48] + "' " +  // 報告数
                                        //    ",ChousaHoukokuRank = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][49].ToString(), 0, 0) + "' " +  // 報告ランク
                                        //    //",ChousaIraiHonsuu = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][50].ToString(), 0, 0) + "' " +  // 依頼数
                                        //    ",ChousaIraiHonsuu = '" + c1FlexGrid4.Rows[i][50]+ "' " +  // 依頼数
                                        //    ",ChousaIraiRank = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][51].ToString(), 0, 0) + "' " +  // 依頼ランク
                                        //    ",ChousaHinmokuShimekiribi = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][52].ToString(), 0, 0) + "' ";  // 締切日

                                        //cmd.CommandText += ",ChousaHoukokuHonsuu = '" + c1FlexGrid4.Rows[i][48] + "' ";  // 報告数
                                        cmd.CommandText += ",ChousaHoukokuHonsuu = '" + c1FlexGrid4.Rows[i]["ChousaHoukokuHonsuu"] + "' ";  // 報告数
                                                                                                                                            //if (c1FlexGrid4.Rows[i][49] != null && c1FlexGrid4.Rows[i][49].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaHoukokuRank"] != null && c1FlexGrid4.Rows[i]["ChousaHoukokuRank"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",ChousaHoukokuRank = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][49].ToString(), 0, 0) + "' ";  // 報告ランク
                                            cmd.CommandText += ",ChousaHoukokuRank = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHoukokuRank"].ToString(), 0, 0) + "' ";  // 報告ランク
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaHoukokuRank = '' ";  // 報告ランク
                                        }

                                        //cmd.CommandText += ",ChousaIraiHonsuu = '" + c1FlexGrid4.Rows[i][50] + "' ";  // 依頼数
                                        cmd.CommandText += ",ChousaIraiHonsuu = '" + c1FlexGrid4.Rows[i]["ChousaIraiHonsuu"] + "' ";  // 依頼数

                                        //if (c1FlexGrid4.Rows[i][51] != null && c1FlexGrid4.Rows[i][51].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaIraiRank"] != null && c1FlexGrid4.Rows[i]["ChousaIraiRank"].ToString() != "")
                                        {
                                            //cmd.CommandText += ",ChousaIraiRank = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][51].ToString(), 0, 0) + "' ";  // 依頼ランク
                                            cmd.CommandText += ",ChousaIraiRank = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaIraiRank"].ToString(), 0, 0) + "' ";  // 依頼ランク
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaIraiRank = '' ";  // 依頼ランク
                                        }

                                        //cmd.CommandText += ",ChousaHinmokuShimekiribi = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][52].ToString(), 0, 0) + "' ";  // 締切日
                                        if (c1FlexGrid4.Rows[i]["ChousaHinmokuShimekiribi"] != null)
                                        {
                                            cmd.CommandText += ",ChousaHinmokuShimekiribi = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmokuShimekiribi"].ToString(), 0, 0) + "' ";  // 締切日
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaHinmokuShimekiribi = null ";  // 締切日
                                        }


                                        // 報告済
                                        //if (c1FlexGrid4.Rows[i][53] != null && c1FlexGrid4.Rows[i][53].ToString() == "True")
                                        if (c1FlexGrid4.Rows[i]["ChousaHoukokuzumi"] != null && c1FlexGrid4.Rows[i]["ChousaHoukokuzumi"].ToString() == "True")
                                        {
                                            cmd.CommandText += ",ChousaHoukokuzumi = 1 ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaHoukokuzumi = 0 ";
                                        }

                                        cmd.CommandText += ",ChousaDeleteFlag = '0' " +                          // 削除フラグ
                                            ",ChousaUpdateDate = '" + sysDateTimeStr + "'" +                // 更新日時
                                            ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +                 // 更新ユーザー
                                            ",ChousaUpdateProgram = '" + pgmName + methodName + "' ";  // 更新プログラム

                                        // 進捗状況
                                        // 価格が入力されている場合、50:担当者済とする
                                        //if (c1FlexGrid4.Rows[i][13] != null && c1FlexGrid4.Rows[i][13].ToString() != "")
                                        if (c1FlexGrid4.Rows[i]["ChousaKakaku"] != null && c1FlexGrid4.Rows[i]["ChousaKakaku"].ToString() != "")
                                        {
                                            cmd.CommandText += ",ChousaShinchokuJoukyou = '50' ";       // 進捗状況

                                            // 進捗状況
                                            //c1FlexGrid4.Rows[i][56] = "50";
                                            c1FlexGrid4.Rows[i]["ChousaShinchokuJoukyou"] = "50";
                                            if (MadoguchiHoukokuzumi != "1")
                                            {
                                                // 進捗状況
                                                //c1FlexGrid4.Rows[i][5] = "7";
                                                c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "7";
                                                if (c1FlexGrid4.Rows[i]["ChousaChuushi"] != null && "True".Equals(c1FlexGrid4.Rows[i]["ChousaChuushi"].ToString()))
                                                {
                                                    c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "6";
                                                }

                                            }
                                        }
                                        else
                                        {
                                            //cmd.CommandText += ",ChousaShinchokuJoukyou = '" + c1FlexGrid4.Rows[i][56] + "' ";       // 進捗状況
                                            cmd.CommandText += ",ChousaShinchokuJoukyou = '20' ";       // 進捗状況

                                            // 価格 が入っていない場合、20:調査開始とする
                                            // 進捗状況
                                            //c1FlexGrid4.Rows[i][56] = "20";
                                            c1FlexGrid4.Rows[i]["ChousaShinchokuJoukyou"] = "20";
                                            if (MadoguchiHoukokuzumi != "1")
                                            {
                                                //DateTime dateTime = DateTime.Parse(c1FlexGrid4[i, 52].ToString());
                                                if (c1FlexGrid4[i, "ChousaHinmokuShimekiribi"] != null)
                                                {
                                                    DateTime dateTime = DateTime.Parse(c1FlexGrid4[i, "ChousaHinmokuShimekiribi"].ToString());
                                                    if (c1FlexGrid4.Rows[i]["ChousaChuushi"] != null && "True".Equals(c1FlexGrid4.Rows[i]["ChousaChuushi"].ToString()))
                                                    {
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "6";
                                                    }
                                                    else if (dateTime < DateTime.Today)
                                                    {
                                                        // 締切日経過
                                                        //c1FlexGrid4.Rows[i][5] = "1";
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "1";
                                                    }
                                                    else if (dateTime < DateTime.Today.AddDays(3))
                                                    {
                                                        // 締切日が3日以内、かつ2次検証が完了していない
                                                        //c1FlexGrid4.Rows[i][5] = "2";
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "2";
                                                    }
                                                    else if (dateTime < DateTime.Today.AddDays(7))
                                                    {
                                                        // 締切日が1週間以内、かつ2次検証が完了していない
                                                        //c1FlexGrid4.Rows[i][5] = "3";
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "3";
                                                    }
                                                    else
                                                    {
                                                        //c1FlexGrid4.Rows[i][5] = "4";
                                                        c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "4";
                                                    }
                                                    //// 進捗状況
                                                    //c1FlexGrid4.Rows[i][5] = "7";
                                                }
                                                else
                                                {
                                                    c1FlexGrid4.Rows[i]["ShinchokuIcon"] = "4";
                                                }
                                            }
                                        }
                                        ////奉行エクセル
                                        //集計表Ver
                                        if (c1FlexGrid4.Rows[i]["ShukeihyoVer"] != null && c1FlexGrid4.Rows[i]["ShukeihyoVer"].ToString() != "")
                                        {
                                            cmd.CommandText += ",ChousaShuukeihyouVer = '" + c1FlexGrid4.Rows[i]["ShukeihyoVer"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaShuukeihyouVer = null ";
                                        }
                                        //分割方法
                                        if (c1FlexGrid4.Rows[i]["BunkatsuHouhou"] != null && c1FlexGrid4.Rows[i]["BunkatsuHouhou"].ToString() != "")
                                        {
                                            cmd.CommandText += ",ChousaBunkatsuHouhou = '" + c1FlexGrid4.Rows[i]["BunkatsuHouhou"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaBunkatsuHouhou = null ";
                                        }
                                        //工事・構造物名
                                        if (c1FlexGrid4.Rows[i]["KojiKoubutsuMei"] != null && c1FlexGrid4.Rows[i]["KojiKoubutsuMei"].ToString() != "")
                                        {
                                            cmd.CommandText += ",ChousaKoujiKouzoubutsumei = '" + c1FlexGrid4.Rows[i]["KojiKoubutsuMei"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaKoujiKouzoubutsumei = null ";
                                        }
                                        // 単位当たり単価（単位）
                                        if (c1FlexGrid4.Rows[i]["TaniAtariTankaTani"] != null && c1FlexGrid4.Rows[i]["TaniAtariTankaTani"].ToString() != "")
                                        {
                                            cmd.CommandText += ",ChousaTaniAtariTanka = '" + c1FlexGrid4.Rows[i]["TaniAtariTankaTani"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaTaniAtariTanka = null ";
                                        }
                                        //単位当たり単価（数量）
                                        if (c1FlexGrid4.Rows[i]["TaniAtariTankaSuryo"] != null && c1FlexGrid4.Rows[i]["TaniAtariTankaSuryo"].ToString() != "")
                                        {
                                            cmd.CommandText += ",chousaTaniAtariSuuryou = '" + c1FlexGrid4.Rows[i]["TaniAtariTankaSuryo"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",chousaTaniAtariSuuryou = null ";
                                        }
                                        // 単位当たり単価（価格）
                                        if (c1FlexGrid4.Rows[i]["TaniAtariTankaKakaku"] != null && c1FlexGrid4.Rows[i]["TaniAtariTankaKakaku"].ToString() != "")
                                        {
                                            cmd.CommandText += ",ChousaTaniAtariKakaku = '" + c1FlexGrid4.Rows[i]["TaniAtariTankaKakaku"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaTaniAtariKakaku = null ";
                                        }
                                        //発注者提供単位
                                        if (c1FlexGrid4.Rows[i]["HachusyaTeikyoTani"] != null && c1FlexGrid4.Rows[i]["HachusyaTeikyoTani"].ToString() != "")
                                        {
                                            cmd.CommandText += ",ChousaHachushaTeikyouTani = '" + c1FlexGrid4.Rows[i]["HachusyaTeikyoTani"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaHachushaTeikyouTani = null ";
                                        }
                                        //荷渡し条件
                                        if (c1FlexGrid4.Rows[i]["NiwatashiJoken"] != null && c1FlexGrid4.Rows[i]["NiwatashiJoken"].ToString() != "")
                                        {
                                            cmd.CommandText += ",ChousaNiwatashiJouken = '" + c1FlexGrid4.Rows[i]["NiwatashiJoken"] + "' ";
                                        }
                                        else
                                        {
                                            cmd.CommandText += ",ChousaNiwatashiJouken = null ";
                                        }


                                        //cmd.CommandText += "WHERE ChousaHinmokuID ='" + c1FlexGrid4.Rows[i][55] + "' AND MadoguchiID ='" + MadoguchiID + "' ";
                                        cmd.CommandText += "WHERE ChousaHinmokuID ='" + c1FlexGrid4.Rows[i]["ChousaHinmokuID2"] + "' AND MadoguchiID ='" + MadoguchiID + "' ";

                                        cmd.ExecuteNonQuery();
                                        // 更新メッセージ
                                        updmessage2 = 1;

                                        // 履歴を登録する
                                        if (beforeChousaHinmokuDT != null)
                                        {
                                            afterRyakuBushoCD = "";
                                            afterChousainCD = "";

                                            for (int j = 0; j < beforeChousaHinmokuDT.Rows.Count; j++)
                                            {
                                                // ChousaHinmokuIDが見つかるまで回す 55:調査品目ID
                                                //if (beforeChousaHinmokuDT.Rows[j][0].ToString() == c1FlexGrid4.Rows[i][55].ToString())
                                                if (beforeChousaHinmokuDT.Rows[j][0].ToString() == c1FlexGrid4.Rows[i]["ChousaHinmokuID2"].ToString())
                                                {
                                                    // 43:調査担当者 42:調査担当部所  beforeChousaHinmokuDT 2:HinmokuChousainCD 1:HinmokuRyakuBushoCD
                                                    //if ((c1FlexGrid4.Rows[i][43] != null && beforeChousaHinmokuDT.Rows[j][2].ToString() != c1FlexGrid4.Rows[i][43].ToString())
                                                    //    ||
                                                    //    (c1FlexGrid4.Rows[i][42] != null && beforeChousaHinmokuDT.Rows[j][1].ToString() != c1FlexGrid4.Rows[i][42].ToString())
                                                    //    )
                                                    if ((c1FlexGrid4.Rows[i]["HinmokuChousainCD"] != null && beforeChousaHinmokuDT.Rows[j][2].ToString() != c1FlexGrid4.Rows[i]["HinmokuChousainCD"].ToString())
                                                        ||
                                                        (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"] != null && beforeChousaHinmokuDT.Rows[j][1].ToString() != c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"].ToString())
                                                        )
                                                    {
                                                        beforeRyakuBushoCD = beforeChousaHinmokuDT.Rows[j][1].ToString();
                                                        beforeChousainCD = beforeChousaHinmokuDT.Rows[j][2].ToString();
                                                        //if(c1FlexGrid4.Rows[i][42] != null) 
                                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"] != null)
                                                        {
                                                            //afterRyakuBushoCD = c1FlexGrid4.Rows[i][42].ToString();
                                                            afterRyakuBushoCD = c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"].ToString();
                                                        }
                                                        //if (c1FlexGrid4.Rows[i][43] != null)
                                                        if (c1FlexGrid4.Rows[i]["HinmokuChousainCD"] != null)
                                                        {
                                                            //afterChousainCD = c1FlexGrid4.Rows[i][43].ToString();
                                                            afterChousainCD = c1FlexGrid4.Rows[i]["HinmokuChousainCD"].ToString();
                                                        }
                                                        if (beforeRyakuBushoCD != "" && beforeChousainCD != "")
                                                        {
                                                            historyMessage = "調査員が更新されました。";
                                                        }
                                                        else
                                                        {
                                                            historyMessage = "調査員が追加されました。";
                                                        }
                                                        writeChousaHinmokuHistory(historyMessage, beforeRyakuBushoCD, beforeChousainCD, afterRyakuBushoCD, afterChousainCD);
                                                    }
                                                    // 45:副調査担当者1 44:副調査担当部所1  beforeChousaHinmokuDT 4:HinmokuFukuChousainCD1 3:HinmokuRyakuBushoFuku1CD
                                                    //if ((c1FlexGrid4.Rows[i][45] != null && beforeChousaHinmokuDT.Rows[j][4].ToString() != c1FlexGrid4.Rows[i][45].ToString())
                                                    //    ||
                                                    //    (c1FlexGrid4.Rows[i][44] != null && beforeChousaHinmokuDT.Rows[j][3].ToString() != c1FlexGrid4.Rows[i][44].ToString())
                                                    //    )
                                                    if ((c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"] != null && beforeChousaHinmokuDT.Rows[j][4].ToString() != c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"].ToString())
                                                        ||
                                                        (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"] != null && beforeChousaHinmokuDT.Rows[j][3].ToString() != c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"].ToString())
                                                        )
                                                    {
                                                        beforeRyakuBushoCD = beforeChousaHinmokuDT.Rows[j][3].ToString();
                                                        beforeChousainCD = beforeChousaHinmokuDT.Rows[j][4].ToString();
                                                        //if (c1FlexGrid4.Rows[i][44] != null)
                                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"] != null)
                                                        {
                                                            //afterRyakuBushoCD = c1FlexGrid4.Rows[i][44].ToString();
                                                            afterRyakuBushoCD = c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku1CD"].ToString();
                                                        }
                                                        //if (c1FlexGrid4.Rows[i][45] != null)
                                                        if (c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"] != null)
                                                        {
                                                            //afterChousainCD = c1FlexGrid4.Rows[i][45].ToString();
                                                            afterChousainCD = c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD1"].ToString();
                                                        }
                                                        if (beforeRyakuBushoCD != "" && beforeChousainCD != "")
                                                        {
                                                            historyMessage = "副調査員1が更新されました。";
                                                        }
                                                        else
                                                        {
                                                            historyMessage = "副調査員1が追加されました。";
                                                        }
                                                        writeChousaHinmokuHistory(historyMessage, beforeRyakuBushoCD, beforeChousainCD, afterRyakuBushoCD, afterChousainCD);
                                                    }
                                                    // 47:副調査担当者2 46:副調査担当部所1  beforeChousaHinmokuDT 6:HinmokuFukuChousainCD2 5:HinmokuRyakuBushoFuku2CD
                                                    //if ((c1FlexGrid4.Rows[i][47] != null && beforeChousaHinmokuDT.Rows[j][6].ToString() != c1FlexGrid4.Rows[i][47].ToString())
                                                    //    ||
                                                    //    (c1FlexGrid4.Rows[i][46] != null && beforeChousaHinmokuDT.Rows[j][5].ToString() != c1FlexGrid4.Rows[i][46].ToString())
                                                    //    )
                                                    if ((c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"] != null && beforeChousaHinmokuDT.Rows[j][6].ToString() != c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"].ToString())
                                                        ||
                                                        (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"] != null && beforeChousaHinmokuDT.Rows[j][5].ToString() != c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"].ToString())
                                                        )
                                                    {
                                                        beforeRyakuBushoCD = beforeChousaHinmokuDT.Rows[j][5].ToString();
                                                        beforeChousainCD = beforeChousaHinmokuDT.Rows[j][6].ToString();
                                                        //if (c1FlexGrid4.Rows[i][46] != null)
                                                        if (c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"] != null)
                                                        {
                                                            //afterRyakuBushoCD = c1FlexGrid4.Rows[i][46].ToString();
                                                            afterRyakuBushoCD = c1FlexGrid4.Rows[i]["HinmokuRyakuBushoFuku2CD"].ToString();
                                                        }
                                                        //if (c1FlexGrid4.Rows[i][47] != null)
                                                        if (c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"] != null)
                                                        {
                                                            //afterChousainCD = c1FlexGrid4.Rows[i][47].ToString();
                                                            afterChousainCD = c1FlexGrid4.Rows[i]["HinmokuFukuChousainCD2"].ToString();
                                                        }
                                                        if (beforeRyakuBushoCD != "" && beforeChousainCD != "")
                                                        {
                                                            historyMessage = "副調査員2が更新されました。";
                                                        }
                                                        else
                                                        {
                                                            historyMessage = "副調査員2が追加されました。";
                                                        }
                                                        writeChousaHinmokuHistory(historyMessage, beforeRyakuBushoCD, beforeChousainCD, afterRyakuBushoCD, afterChousainCD);
                                                    }
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                                transaction.Commit();

                                // ３．ChousaHinmokuから担当部所の連携を行う（支部備考も）
                                String resultMessage = "";
                                GlobalMethod.MadoguchiHinmokuRenkeiUpdate_SQL(MadoguchiID, "Jibun", UserInfos[0], out resultMessage);

                                // メッセージがあれば画面に表示
                                if (resultMessage != "")
                                {
                                    set_error(resultMessage);
                                }

                                transaction = conn.BeginTransaction();
                                cmd.Transaction = transaction;

                                //string table = "ChousaHinmoku";
                                //// 編集ロック開放
                                //// Lockテーブル更新
                                //cmd.CommandText = "DELETE FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
                                //                    "AND LOCK_KEY = '" + MadoguchiID + "' " +
                                //                    "AND LOCK_USER_ID = '" + UserInfos[0] + "' ";
                                //cmd.ExecuteNonQuery();

                                GlobalMethod.outputLogger("Madoguchi button3_InputStatus_Click", "DB更新終了 ID:" + MadoguchiID, "update", "DEBUG");

                                // 画面のメッセージ表示 （新規 or 更新）+ 削除
                                if (updmessage1 == 1)
                                {
                                    // I20302:調査品目明細を追加しました。
                                    set_error(GlobalMethod.GetMessage("I20302", ""));
                                }
                                else if (updmessage2 == 1)
                                {
                                    // I20301:調査品目明細を更新しました。
                                    set_error(GlobalMethod.GetMessage("I20301", ""));
                                }
                                if (updmessage3 == 1)
                                {
                                    // I20303:調査品目明細を削除しました。
                                    set_error(GlobalMethod.GetMessage("I20303", ""));
                                }

                                transaction.Commit();
                            }
                            catch (Exception)
                            {
                                transaction.Rollback();
                                throw;
                            }
                            finally
                            {
                                conn.Close();
                            }
                            // 調査品目の削除Keys
                            deleteChousaHinmokuIDs = "";
                            ChousaHinmokuMode = 0;
                            // 編集状態を解除する
                            ChousaHinmokuGrid_InputMode();

                            // 背景色を通常色に戻す
                            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                            {
                                //c1FlexGrid4.GetCellRange(i, 6).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                                //c1FlexGrid4.GetCellRange(i, 7).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                                //c1FlexGrid4.GetCellRange(i, 34).StyleNew.BackColor = Color.White;
                                //c1FlexGrid4.GetCellRange(i, 34).StyleNew.BackColor = Color.White;
                                //c1FlexGrid4.GetCellRange(i, 13).StyleNew.BackColor = Color.White;
                                c1FlexGrid4.GetCellRange(i, ZentaiJunColIndex).StyleNew.BackColor = Color.FromArgb(245, 245, 220);      // 全体順
                                c1FlexGrid4.GetCellRange(i, KobetsuJunColIndex).StyleNew.BackColor = Color.FromArgb(245, 245, 220);     // 個別順
                                c1FlexGrid4.GetCellRange(i, TikuWariCodeColIndex).StyleNew.BackColor = Color.White;
                                c1FlexGrid4.GetCellRange(i, TikuCodeColIndex).StyleNew.BackColor = Color.White;
                                c1FlexGrid4.GetCellRange(i, KakakuColIndex).StyleNew.BackColor = Color.White;
                                // 更新が通ったので、
                                c1FlexGrid4.Rows[i]["Mode"] = "1"; // 0:Insert/1:Select/2:Update
                            }
                        }
                        // errorFlg がtrueの場合
                        else
                        {
                            // エラーメッセージ表示
                            if (errmessage1 == 1)
                            {
                                // E20307: 全体順、個別順が重複しています。
                                set_error(GlobalMethod.GetMessage("E20307", ""));
                            }
                            else if (errmessage2 == 1)
                            {
                                // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい。
                                set_error(GlobalMethod.GetMessage("E20336", ""));
                            }
                            if (errmessage3 == 1)
                            {
                                // E20337:半角数字で入力してください。赤背景の項目を修正して下さい。
                                set_error(GlobalMethod.GetMessage("E20337", ""));
                            }
                            if (errmessage4 == 1)
                            {
                                // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                                set_error(GlobalMethod.GetMessage("E10010", ""));
                            }
                        }

                        // 並び順列のIndex（行番号）を取得する。
                        int ColumnSortColIndex = c1FlexGrid4.Cols["ColumnSort"].Index;

                        // ソート
                        //c1FlexGrid4.Cols[58].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                        //c1FlexGrid4.Cols[6].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                        //c1FlexGrid4.Cols.Move(58, 6);
                        //c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 58, 6);
                        //c1FlexGrid4.Cols.Move(6, 58);
                        c1FlexGrid4.Cols[ColumnSortColIndex].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                        c1FlexGrid4.Cols[ZentaiJunColIndex].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                        c1FlexGrid4.Cols.Move(ColumnSortColIndex, 1);
                        c1FlexGrid4.Cols.Move(ZentaiJunColIndex, 2);
                        c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 1, 2);
                        c1FlexGrid4.Cols.Move(2, ZentaiJunColIndex);
                        c1FlexGrid4.Cols.Move(1, ColumnSortColIndex);
                    }

                    // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック、Garoon連携対象チェック
                    //Garoon連携対象である場合、かつ、更新処理でエラーが出ていない場合に連携処理を行う。
                    if (item1_GaroonRenkei.Checked == true && globalErrorFlg == "0")
                    {
                        // VIPS　20220302　課題管理表No1275(969)　ADD　「Garoon連携処理」追加　対応
                        GaroonBtn_Click(sender, e);
                    }
                }
            }

        }

        //調査品目Gridソート時
        private void c1FlexGrid4_AfterSort(object sender, C1.Win.C1FlexGrid.SortColEventArgs e)
        {
            Grid_Visible(int.Parse(Paging_now.Text));
        }
        private void c1FlexGrid4_OwnerDrawCell(object sender, C1.Win.C1FlexGrid.OwnerDrawCellEventArgs e)
        {
            //if (e.Row > 1 && e.Col == 0)
            //{
            //    e.Image = Img_AddRow;
            //}
            //if (e.Row > 1 && e.Col == 1)
            //{
            //    e.Image = Img_DeleteRow;
            //}
            //if (e.Row > 1 && e.Col == 2)
            //{
            //    e.Image = Img_AddRowNonactive;
            //}
            //if (e.Row > 1 && e.Col == 3)
            //{
            //    e.Image = Img_DeleteRowNonactive;
            //}
            //if (e.Row == 0 && e.Col >= 5)
            //{
            //    e.Image = Img_Sort;
            //}

            switch (c1FlexGrid4.Cols[e.Col].Name)
            {
                case "RowChange":
                    break;
                case "Add1":
                    if (e.Row > 1)
                    {
                        e.Image = Img_AddRow;
                    }
                    break;
                case "Delete1":
                    if (e.Row > 1)
                    {
                        e.Image = Img_DeleteRow;
                    }
                    break;
                case "Add2":
                    if (e.Row > 1)
                    {
                        e.Image = Img_AddRowNonactive;
                    }
                    break;
                case "Delete2":
                    if (e.Row > 1)
                    {
                        e.Image = Img_DeleteRowNonactive;
                    }
                    break;
                default:
                    // ヘッダーにはソートの画像をセットする
                    if (e.Row == 0 && e.Col != c1FlexGrid4.Cols["ChousaLinkSakli"].Index && e.Col != c1FlexGrid4.Cols["ChousaLinkSakliFolder"].Index)
                    {
                        C1.Win.C1FlexGrid.CellRange cr;
                        cr = c1FlexGrid4.GetCellRange(0, e.Col);
                        cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
                        cr.Image = Img_Sort;
                        //e.Image = Img_Sort;
                    }
                    break;
            }
        }

        public void Resize_Grid(string name)
        {
            Control[] cs;
            cs = this.Controls.Find(name, true);
            if (cs.Length > 0)
            {
                var fx = (C1.Win.C1FlexGrid.C1FlexGrid)cs[0];
                int h = 0;
                for (int i = 0; i < fx.Rows.Count; i++)
                {
                    if (fx.Rows[i].Height == -1)
                    {
                        h += 22;
                    }
                    else
                    {
                        h += fx.Rows[i].Height;
                    }
                }
                fx.Height = 4 + h;
                //c1FlexGrid1.Height = 4 + h;

                int w = 0;
                for (int i = 0; i < fx.Cols.Count; i++)
                {
                    if (fx.Cols[i].Width == -1)
                    {
                        w += 100;
                    }
                    else
                    {
                        w += fx.Cols[i].Width;
                    }
                }
                if (fx.Width < 4 + w)
                {
                    fx.Height += 18;
                }
            }
        }

        // 割振訂正ボタン
        private void button2_InputMode_Click(object sender, EventArgs e)
        {

            //レコードがある場合
            if (c1FlexGrid1.Rows[1][1] != null)
            {
                //割振訂正で編集許可
                c1FlexGrid1.Cols[4].AllowEditing = true;
                c1FlexGrid1.Cols[5].AllowEditing = true;
                c1FlexGrid1.Cols[6].AllowEditing = true;
            }

            //そのボタンを押せなくする
            button2_InputMode.Enabled = false;
            button2_InputMode.BackColor = Color.DarkGray;

            //更新ボタンを押せるようにする
            button2_Update.Enabled = true;
            button2_Update.BackColor = Color.FromArgb(42, 78, 122);
            button2_Update.ForeColor = Color.FromArgb(255, 255, 255);

        }

        private void button2_Update_Click(object sender, EventArgs e)
        {
            //更新ボタン処理
            if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                // エラーフラグ true：エラー false：正常
                Boolean errorFlg = false;
                // エラーメッセージフラグ
                Boolean messageFlg1 = false;
                Boolean messageFlg2 = false;

                string tantousha = "";

                set_error("", 0);
                // エラーチェック
                for (int i = 1; i < c1FlexGrid5.Rows.Count; i++)
                {
                    // 担当者が重複していないか確認
                    if (c1FlexGrid5.Rows[i][3] != null && !"".Equals(c1FlexGrid5.Rows[i][3].ToString()) && !"0".Equals(c1FlexGrid5.Rows[i][3].ToString()) && !"".Equals(tantousha))
                    {
                        if (tantousha.IndexOf(c1FlexGrid5.Rows[i][3].ToString()) > -1)
                        {
                            messageFlg2 = true;
                            errorFlg = true;
                        }
                    }

                    //// 部所が選択されており、担当者が空の場合、エラー
                    //// 2:担当部所 3:担当者
                    //if (c1FlexGrid5.Rows[i][2] != null && c1FlexGrid5.Rows[i][3] != null && !"".Equals(c1FlexGrid5.Rows[i][2].ToString()) && !"0".Equals(c1FlexGrid5.Rows[i][2].ToString()) && "".Equals(c1FlexGrid5.Rows[i][3].ToString()))
                    //{
                    //    messageFlg1 = true;
                    //    errorFlg = true;
                    //}
                    if (c1FlexGrid5.Rows[i][2] != null && !"".Equals(c1FlexGrid5.Rows[i][2].ToString()) && !"0".Equals(c1FlexGrid5.Rows[i][2].ToString()))
                    {
                        if (c1FlexGrid5.Rows[i][3] == null || "".Equals(c1FlexGrid5.Rows[i][3].ToString()))
                        {
                            messageFlg1 = true;
                            errorFlg = true;
                        }
                    }

                    // 担当者が空でない場合、変数に格納する
                    if (c1FlexGrid5.Rows[i][3] != null && !"".Equals(c1FlexGrid5.Rows[i][3].ToString()) && !"0".Equals(c1FlexGrid5.Rows[i][3].ToString()))
                    {
                        tantousha += c1FlexGrid5.Rows[i][3] + ",";
                    }
                }
                if (errorFlg == false)
                {
                    //更新ボタン処理
                    //調査担当者Gridのデータ
                    string[,] SQLData = new string[c1FlexGrid1.Rows.Count - 1, 7];
                    for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                    {
                        if (c1FlexGrid1.Rows[1][1] == null)
                        {
                            break;
                        }
                        SQLData[i - 1, 0] = c1FlexGrid1.Rows[i][1].ToString();
                        SQLData[i - 1, 1] = c1FlexGrid1.Rows[i][2].ToString();
                        SQLData[i - 1, 2] = c1FlexGrid1.Rows[i][3].ToString();
                        if (c1FlexGrid1.Rows[i][4] != null && c1FlexGrid1.Rows[i][4].ToString() != "")
                        {
                            SQLData[i - 1, 3] = c1FlexGrid1.Rows[i][4].ToString();
                        }
                        //SQLData[i - 1, 3] = c1FlexGrid1.Rows[i][4].ToString();
                        SQLData[i - 1, 4] = c1FlexGrid1.Rows[i][5].ToString();
                        if (c1FlexGrid1.GetCellCheck(i, 6) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                        {
                            SQLData[i - 1, 5] = "1";
                        }
                        else
                        {
                            SQLData[i - 1, 5] = "0";
                        }
                        if (checkBox28.Checked == true)
                        {
                            SQLData[i - 1, 6] = "1";
                        }
                        else
                        {
                            SQLData[i - 1, 6] = "0";
                        }
                    }
                    //不具合No1332(1084)　グリッドの行UserDataに画面更新フラグ0、1をセットするようにしたので、6にセットするよう修正
                    //Groon追加宛先Gridのデータ
                    string[,] SQLData2 = new string[c1FlexGrid5.Rows.Count - 1, 6];
                    //string[,] SQLData2 = new string[c1FlexGrid5.Rows.Count - 1, 5];
                    for (int i = 1; i < c1FlexGrid5.Rows.Count; i++)
                    {
                        if (c1FlexGrid5.Rows[i][1] != null)
                        {
                            SQLData2[i - 1, 0] = c1FlexGrid5.Rows[i][1].ToString();
                        }
                        if (c1FlexGrid5.Rows[i][2] != null)
                        {
                            SQLData2[i - 1, 1] = c1FlexGrid5.Rows[i][2].ToString();
                            SQLData2[i - 1, 2] = c1FlexGrid5.GetDataDisplay(i, 2).ToString();
                        }
                        if (c1FlexGrid5.Rows[i][3] != null)
                        {
                            SQLData2[i - 1, 3] = c1FlexGrid5.Rows[i][3].ToString();
                            SQLData2[i - 1, 4] = c1FlexGrid5.GetDataDisplay(i, 3).ToString();
                        }
                        //不具合No1332(1084)
                        SQLData2[i - 1, 5] = c1FlexGrid5.Rows[i].UserData.ToString();
                    }

                    string mes = "";
                    set_error("", 0);
                    GlobalMethod.MadoguchiUpdate_SQL(2, MadoguchiID, SQLData, out mes, UserInfos, SQLData2, "Tokumei");
                    set_error(mes);
                    get_data(2);

                    //　担当部所の更新ボタンを非活性
                    button2_Update.Enabled = false;
                    button2_Update.BackColor = Color.DarkGray;

                    // 担当部所の割振訂正ボタンを活性
                    button2_InputMode.Enabled = true;
                    button2_InputMode.BackColor = Color.FromArgb(42, 78, 122);
                    button2_InputMode.ForeColor = Color.FromArgb(255, 255, 255);
                }
                else
                {
                    if (messageFlg1 == true)
                    {
                        // E70054:Garoon追加宛先の担当者が未選択です。
                        set_error(GlobalMethod.GetMessage("E70054", ""));
                    }
                    if (messageFlg2 == true)
                    {
                        // E70055:Garoon追加宛先の担当者が重複しています。
                        set_error(GlobalMethod.GetMessage("E70055", ""));
                    }
                }
            }
         }

        // 調査品目明細の調査担当者のプロンプト
        private void item3_ChousaTantouPrompt(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //選択されている年度を条件に調査員プロンプトを表示
            if (item1_TourokuNendo.Text != "")
            {
                form.nendo = item1_TourokuNendo.Text;
            }
            form.program = "madoguchi";
            if (src_Busho.Text != "")
            {
                form.Busho = src_Busho.SelectedValue.ToString();
            }
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                src_HinmokuChousain.Text = form.ReturnValue[1];

                src_Busho.SelectedValue = form.ReturnValue[2];
            }
        }

        // 担当部所タブのGaroon追加宛先
        private void c1FlexGrid5_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid5.HitTest(new Point(e.X, e.Y));
            //担当者列をクリック
            if (hti.Row > 0 && hti.Column == 3)
            {
                Popup_ChousainList form = new Popup_ChousainList();
                if (c1FlexGrid5[hti.Row, 2] != null && c1FlexGrid5[hti.Row, 2].ToString() != "")
                {
                    form.Busho = c1FlexGrid5[hti.Row, 2].ToString();
                }
                else
                {
                    form.Busho = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString();
                }
                form.program = "madoguchi";
                if (item1_TourokuNendo.Text != "")
                {
                    form.nendo = item1_TourokuNendo.ToString();
                }
                form.ShowDialog();
                if (form.ReturnValue != null && form.ReturnValue[0] != null)
                {
                    c1FlexGrid5[hti.Row, 2] = form.ReturnValue[2];
                    c1FlexGrid5[hti.Row, 3] = form.ReturnValue[0];
                }
            }
            //削除列をクリック
            if (hti.Row > 0 && hti.Column == 0)
            {
                if (GlobalMethod.outputMessage("I10002", "", 1) == DialogResult.OK)
                {
                    c1FlexGrid5.Rows.Remove(hti.Row);
                    Resize_Grid("c1FlexGrid5");
                }
            }
        }

        private void button_GaroonAtesakiGridAdd_Click(object sender, EventArgs e)
        {
            c1FlexGrid5.Rows.Add();
            c1FlexGrid5.Rows[c1FlexGrid5.Rows.Count - 1].Height = 28;
            //不具合No1332(1084)　画面から追加されたよフラグをつける
            c1FlexGrid5.Rows[c1FlexGrid5.Rows.Count - 1].UserData = "1";
            Resize_Grid("c1FlexGrid5");
        }

        // 調査品目明細 調査品目一覧からの取込ボタン
        private void button3_ReadExcelChousaHinmoku_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            // エラーフラグ false：正常 true：エラー
            Boolean errorFlg = false;
            button3_ReadExcelResult.BackColor = Color.DarkGray;

            //string table = "ChousaHinmoku";
            string UserID = "";
            //string chousainMei = "";

            //// Lockテーブル更新
            //var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            //using (var conn = new SqlConnection(connStr))
            //{
            //    conn.Open();
            //    var cmd = conn.CreateCommand();
            //    SqlTransaction transaction = conn.BeginTransaction();
            //    cmd.Transaction = transaction;

            //    try
            //    {
            //        // Lock情報取得
            //        // 0:個人コード、1:氏名、2:部所CD、3:部所名、4：Role
            //        chousainMei = UserInfos[1];

            //        cmd.CommandText = "SELECT TOP 1 LOCK_USER_ID,LOCK_USER_MEI FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
            //                          "AND LOCK_KEY = '" + MadoguchiID + "' ";
            //        DataTable dt = new DataTable();

            //        var sda = new SqlDataAdapter(cmd);
            //        sda.Fill(dt);

            //        if (dt.Rows.Count > 0)
            //        {
            //            // Lockテーブルにデータが存在した場合
            //            UserID = dt.Rows[0][0].ToString();
            //            chousainMei = dt.Rows[0][1].ToString();
            //        }
            //        else
            //        {
            //            // Lockテーブルにデータ存在しない場合
            //            cmd.CommandText = "INSERT INTO T_LOCK(" +
            //                             " LOCK_TABLE" +
            //                             ",LOCK_KEY" +
            //                             ",LOCK_USER_ID" +
            //                             ",LOCK_USER_MEI" +
            //                             ",LOCK_DATETIME" +
            //                             ")VALUES(" +
            //                             "'" + table + "' " +
            //                             ",'" + MadoguchiID + "' " +
            //                             ",'" + UserInfos[0] + "' " +
            //                             ",'" + UserInfos[1] + "' " +
            //                             ",SYSDATETIME() " +
            //                             ")";
            //            cmd.ExecuteNonQuery();
            //            transaction.Commit();
            //            UserID = UserInfos[0];
            //            chousainMei = UserInfos[1];
            //        }
            //    }
            //    catch
            //    {
            //        transaction.Rollback();
            //        errorFlg = true;
            //    }
            //    finally
            //    {
            //        conn.Close();
            //    }
            //}

            Popup_Loading Loading = new Popup_Loading();
            Loading.StartPosition = FormStartPosition.CenterScreen;
            Loading.Show();

            //ファイル
            OpenFileDialog Dialog1 = new OpenFileDialog();
            Dialog1.InitialDirectory = @"C:";
            Dialog1.Title = "インポートするファイルを選択してください。";

            // ファイルが開けたか
            if (Dialog1.ShowDialog() == DialogResult.OK)
            {
                // ファイル名チェック 調査品目一覧 .xlsm _ をreplaceし、特調番号を取得
                String tokuchoBangou = Dialog1.FileName;

                // ファイルの開始前の\ 位置を取得
                int filePath = tokuchoBangou.LastIndexOf(@"\");
                // 最後の\より1つ後ろから
                tokuchoBangou = tokuchoBangou.Substring(filePath + 1, tokuchoBangou.Length - (filePath + 1));

                tokuchoBangou = tokuchoBangou.Replace("調査品目一覧", "");
                tokuchoBangou = tokuchoBangou.Replace(".xlsm", "");
                tokuchoBangou = tokuchoBangou.Replace("_", "");

                if (tokuchoBangou == "")
                {
                    // E20349:調査品目明細一覧ファイル名に特調番号が付与されていない可能性があります。例のようにファイル名に特調番号-枝番を付与して下さい。例：調査品目一覧_T18019999-Z999.xlsm
                    set_error(GlobalMethod.GetMessage("E20349", ""));
                    errorFlg = true;
                }
                else
                {
                    // 特調番号が一致するか
                    if (!tokuchoBangou.Equals(item1_MadoguchiUketsukeBangou.Text + "-" + item1_MadoguchiUketsukeBangouEdaban.Text))
                    {
                        // E20348:調査品目明細一覧ファイルの特調番号が一致しませんでした。
                        set_error(GlobalMethod.GetMessage("E20348", ""));
                        errorFlg = true;
                    }
                }

                // 正常の場合
                if (errorFlg == false)
                {
                    //// ロック所有か自分がどうか
                    //if (UserID == UserInfos[0])
                    //{
                    string[] result = GlobalMethod.InsertHinmoku(Dialog1.FileName, MadoguchiID, UserInfos[0], UserInfos[2]);

                    // result
                    // 成否判定 0:正常 1：エラー
                    // T_ReadFileErrorテーブルのエラーカウント（FileReadErrorReadCount）
                    // メッセージ（主にエラー用）
                    if (result != null && result.Length >= 1)
                    {
                        // 改行コードがあるので、削る
                        result[0] = result[0].Replace(@"\r\n", "");

                        if (result[0].Trim() == "1")
                        {
                            //set_error(result[1]);
                            // エラーが発生しました
                            set_error(GlobalMethod.GetMessage("E00091", ""));
                            set_error(result[2]);
                            int count = 0;
                            // T_ReadFileErrorテーブルのエラーカウントをセット
                            if (result[1] != null && int.TryParse(result[1].ToString(), out count))
                            {
                                errorCnt = count;
                            }
                            button3_ReadExcelResult.BackColor = Color.FromArgb(42, 78, 122);
                        }
                        else if (result[0].Trim() == "0")
                        {
                                // 正常 データ取り直し
                                get_data(3);
                                // E20321:取込が完了しました。
                                set_error(GlobalMethod.GetMessage("E20321", ""));

                                // 調査品目明細から担当部所への連携 + 担当部所から窓口情報への連携
                                // ProUpdateHinmokuRenkei.Call(&p_MadoguchiID,&TabCode,&pRes)

                                String resultMessage = "";
                                Boolean hinmokuRenkeiResult = true;
                                hinmokuRenkeiResult = GlobalMethod.MadoguchiHinmokuRenkeiUpdate_SQL(MadoguchiID, "Tokumei", UserInfos[0], out resultMessage);

                                // メッセージがあれば画面に表示
                                if (resultMessage != "")
                                {
                                    set_error(resultMessage);
                                }

                            // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック、Garoon連携対象チェック
                            //Garoon連携対象である場合、かつ、下記SQLの処理が正常終了した場合、Garoon連携処理を行う
                            if (item1_GaroonRenkei.Checked == true && hinmokuRenkeiResult == true)
                            {
                                    // VIPS　20220302　課題管理表No1275(969)　ADD　「Garoon連携処理」追加　対応
                                    GaroonBtn_Click(sender, e);
                                }

                            //// 編集ロック開放
                            //// Lockテーブル更新
                            //using (var conn = new SqlConnection(connStr))
                            //{
                            //    conn.Open();
                            //    var cmd = conn.CreateCommand();
                            //    SqlTransaction transaction = conn.BeginTransaction();
                            //    cmd.Transaction = transaction;

                            //    try
                            //    {
                            //        // Lock情報取得
                            //        // 0:個人コード、1:氏名、2:部所CD、3:部所名、4：Role
                            //        chousainMei = UserInfos[1];

                            //        cmd.CommandText = "DELETE FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
                            //                          "AND LOCK_KEY = '" + MadoguchiID + "' " +
                            //                          "AND LOCK_USER_ID = '" + UserInfos[0] + "' ";

                            //        cmd.ExecuteNonQuery();
                            //        transaction.Commit();
                            //    }
                            //    catch
                            //    {
                            //        transaction.Rollback();
                            //        errorFlg = true;
                            //    }
                            //    finally
                            //    {
                            //        conn.Close();
                            //    }
                            //}
                            button3_ReadExcelResult.BackColor = Color.DarkGray;
                        }
                        else
                        {
                            // エラーが発生しました
                            set_error(GlobalMethod.GetMessage("E00091", ""));
                            GlobalMethod.outputLogger("Madoguchi_Input", "調査品目明細取込 :exeファイルの呼び出しに失敗しました", "insert", "DEBUG");
                            button3_ReadExcelResult.BackColor = Color.FromArgb(42, 78, 122);
                        }
                    }
                    else
                    {
                        // E20322:取込ファイルにエラーがありました。
                        set_error(GlobalMethod.GetMessage("E20322", ""));
                        button3_ReadExcelResult.BackColor = Color.FromArgb(42, 78, 122);
                    }
                    //}
                    //// 自身がロックしていない場合
                    //else
                    //{
                    //    set_error("現在編集にロックがかかっています");
                    //}
                }
            }
            else
            {
                // E70039:ファイルが読み込まれていません。
                set_error(GlobalMethod.GetMessage("E70039", ""));
            }
            Dialog1.Dispose();
            Loading.Close();
        }

        // 調査品目一覧出力ボタン
        private void button3_ExcelChousaHinmoku_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            // string[]
            // 0:MadoguchiID        窓口ID           MadoguchiID
            // 1:Shozoku            調査担当部所     src_Busho
            // 2:HinmokuChousain    調査担当者       src_HinmokuChousain
            // 3:ShuFuku            主+副            src_ShuFuku
            // 4:ChousaHinmei       品名             src_ChousaHinmei
            // 5:ChousaKikaku       規格             src_ChousaKikaku
            // 6:Zaikou             材工             src_Zaikou
            // 7:TantoushaKuuhaku   担当者空白リスト src_TantoushaKuuhaku
            // 8:PrintGamen         呼び出し元画面（0:窓口ミハル、1:特命課長）

            // 9個用意
            string[] report_data = new string[9] { "", "", "", "", "", "", "", "", "" };

            report_data[0] = MadoguchiID;
            // 調査担当部所 
            if (src_Busho.Text != null && src_Busho.Text != "")
            {
                report_data[1] = src_Busho.SelectedValue.ToString();
            }
            else
            {
                report_data[1] = "";
            }
            report_data[2] = src_HinmokuChousain.Text;
            // 主+副 0:主+副 1:主のみ 2:副のみ
            report_data[3] = src_ShuFuku.SelectedValue.ToString();

            report_data[4] = src_ChousaHinmei.Text;
            report_data[5] = src_ChousaKikaku.Text;
            // 材工
            if (src_Zaikou.Text != null && src_Zaikou.Text != "")
            {
                report_data[6] = src_Zaikou.SelectedValue.ToString();
            }
            else
            {
                report_data[6] = "";
            }
            // 担当者空白リスト
            if (src_TantoushaKuuhaku.Text != null && src_TantoushaKuuhaku.Text != "")
            {
                report_data[7] = src_TantoushaKuuhaku.SelectedValue.ToString();
            }
            else
            {
                report_data[7] = "";
            }
            // 呼び出し元画面
            report_data[8] = "1";   // 1:特命課長

            // 44:調査品目一覧出力
            string[] result = GlobalMethod.InsertReportWork(44, UserInfos[0], report_data);

            // result
            // 成否判定 0:正常 1：エラー
            // メッセージ（主にエラー用）
            // ファイル物理パス（C:\Work\xxxx\0000000111_エントリーシート.xlsx）
            // ダウンロード時のファイル名（エントリーシート.xlsx）
            if (result != null && result.Length >= 4)
            {
                if (result[0].Trim() == "1")
                {
                    set_error(result[1]);
                }
                else
                {
                    string fileName = result[3];

                    fileName = fileName.Replace("特調番号", item1_MadoguchiUketsukeBangou.Text + "-" + item1_MadoguchiUketsukeBangouEdaban.Text);

                    Popup_Download form = new Popup_Download();
                    form.TopLevel = false;
                    this.Controls.Add(form);
                    form.ExcelName = Path.GetFileName(fileName);
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

        // 調査品目取込結果ボタン
        private void button3_ReadExcelResult_Click(object sender, EventArgs e)
        {
            if (button3_ReadExcelResult.BackColor == Color.FromArgb(42, 78, 122))
            {
                Popup_FileError form = new Popup_FileError(MadoguchiID, errorCnt);
                form.ShowDialog();
            }
        }

        private void button3_ExcelShukeihyo_Click(object sender, EventArgs e)
        {
            //集計表プロンプト
            Popup_ShukeiHyou form = new Popup_ShukeiHyou();
            //form.nendo = item1_3.SelectedValue.ToString();
            form.MadoguchiID = MadoguchiID;
            form.Busho = UserInfos[2];
            form.TokuhoBangou = item1_MadoguchiUketsukeBangou.Text;
            form.TokuhoBangouEda = item1_MadoguchiUketsukeBangouEdaban.Text;
            form.KanriBangou = item1_MadoguchiKanriBangou.Text;
            form.UserInfos = UserInfos;
            form.PrintGamen = "Tokumei";
            form.ShowDialog();

            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                //item1_MadoguchiTantoushaCD.Text = form.ReturnValue[0];
                //item1_MadoguchiTantousha.Text = form.ReturnValue[1];

                //item_Hyoujikensuu.SelectedIndex = 1;
                item_Hyoujikensuu.SelectedIndex = 4; // 1000件対応
                // データ取り直し
                get_data(3);
            }
        }

        // 担当部所タブの調査担当者Grid
        private void c1FlexGrid1_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            //DateTime dateTime = DateTime.Today;

            //// 52:締切日
            ////if (e.Row > 0 && e.Col == 52)
            ////if (e.Row > 0 && e.Col == c1FlexGrid4.Cols["ChousaHinmokuShimekiribi"].Index)
            //// 4:締切日、5:担当者状況変更
            //if (e.Row > 0 && (e.Col == 4 || e.Col == 5))
            //{
            //    // 1:報告済みの場合
            //    if ("1".Equals(MadoguchiHoukokuzumi))
            //    {
            //        // 報告済み
            //        //c1FlexGrid4[e.Row, 5] = "8";
            //        c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "8";
            //    }
            //    else
            //    {
            //        // ▼担当者状況
            //        // 10:依頼
            //        // 20:調査開始
            //        // 30:見積中
            //        // 40:集計中
            //        // 50:担当者済
            //        // 60:一次検済
            //        // 70:二次検済
            //        // 80:中止
            //        //if ("80".Equals(c1FlexGrid4[e.Row, 55].ToString()))
            //        if ("80".Equals(c1FlexGrid4.Rows[e.Row]["ChousaShinchokuJoukyou"].ToString()))
            //        {
            //            // 二次検証済み、または中止（中止）
            //            //c1FlexGrid4.Rows[e.Row][5] = "6";
            //            c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "6";
            //        }
            //        //else if ("70".Equals(c1FlexGrid4[e.Row, 55].ToString()))
            //        else if ("70".Equals(c1FlexGrid4.Rows[e.Row]["ChousaShinchokuJoukyou"].ToString()))
            //        {
            //            // 二次検証済み、または中止（二次検証済み）
            //            //c1FlexGrid4.Rows[e.Row][5] = "5";
            //            c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "5";
            //        }
            //        //else if ("50".Equals(c1FlexGrid4[e.Row, 55].ToString()) || "60".Equals(c1FlexGrid4[e.Row, 55].ToString()))
            //        else if ("50".Equals(c1FlexGrid4.Rows[e.Row]["ChousaShinchokuJoukyou"].ToString()) || "60".Equals(c1FlexGrid4.Rows[e.Row]["ChousaShinchokuJoukyou"].ToString()))
            //        {
            //            // 担当者済み or 一次検済
            //            //c1FlexGrid4.Rows[e.Row][5] = "7";
            //            c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "7";
            //        }
            //        //else if (c1FlexGrid4[e.Row, 52] != null)
            //        else if (c1FlexGrid4.Rows[e.Row]["ChousaHinmokuShimekiribi"] != null)
            //        {
            //            try
            //            {
            //                //dateTime = DateTime.Parse(c1FlexGrid4[e.Row, 52].ToString());
            //                dateTime = DateTime.Parse(c1FlexGrid4.Rows[e.Row]["ChousaHinmokuShimekiribi"].ToString());
            //                if (dateTime < DateTime.Today)
            //                {
            //                    // 締切日経過
            //                    //c1FlexGrid4.Rows[e.Row][5] = "1";
            //                    c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "1";
            //                }
            //                else if (dateTime < DateTime.Today.AddDays(3))
            //                {
            //                    // 締切日が3日以内、かつ2次検証が完了していない
            //                    //c1FlexGrid4.Rows[e.Row][5] = "2";
            //                    c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "2";
            //                }
            //                else if (dateTime < DateTime.Today.AddDays(7))
            //                {
            //                    // 締切日が1週間以内、かつ2次検証が完了していない
            //                    //c1FlexGrid4.Rows[e.Row][5] = "3";
            //                    c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "3";
            //                }
            //                else
            //                {
            //                    //c1FlexGrid4.Rows[e.Row][5] = "4";
            //                    c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "4";
            //                }
            //            }
            //            catch
            //            {
            //                // 日付変換エラー
            //                throw;
            //            }
            //        }
            //    }
            //}

            //// 報告数、依頼数
            ////if (e.Col == 48 || e.Col == 50)
            //if (e.Col == c1FlexGrid4.Cols["ChousaHoukokuHonsuu"].Index || e.Col == c1FlexGrid4.Cols["ChousaIraiHonsuu"].Index)
            //{
            //    if (c1FlexGrid4.Rows[e.Row][e.Col] != null)
            //    {
            //        if (!Regex.IsMatch(c1FlexGrid4.Rows[e.Row][e.Col].ToString(), @"^-?[\d][\d.]*$", RegexOptions.ECMAScript))
            //        {
            //            c1FlexGrid4.Rows[e.Row][e.Col] = "";
            //        }
            //    }
            //}
            DateTime dateTime = DateTime.Today;

            // 4:締切日、5:担当者状況変更
            if (e.Row > 0 && (e.Col == 4 || e.Col == 5))
            {
                // 1:報告済みの場合
                if ("1".Equals(MadoguchiHoukokuzumi))
                {
                    // 報告済み
                    c1FlexGrid1[e.Row, 0] = "8";
                }
                else
                {
                    // ▼担当者状況
                    // 10:依頼
                    // 20:調査開始
                    // 30:見積中
                    // 40:集計中
                    // 50:担当者済
                    // 60:一次検済
                    // 70:二次検済
                    // 80:中止
                    if ("80".Equals(c1FlexGrid1[e.Row, 5].ToString()))
                    {
                        // 二次検証済み、または中止（中止）
                        c1FlexGrid1.Rows[e.Row][0] = "6";
                    }
                    else if ("70".Equals(c1FlexGrid1[e.Row, 5].ToString()))
                    {
                        // 二次検証済み、または中止（二次検証済み）
                        c1FlexGrid1.Rows[e.Row][0] = "5";
                    }
                    else if ("50".Equals(c1FlexGrid1[e.Row, 5].ToString()) || "60".Equals(c1FlexGrid1[e.Row, 5].ToString()))
                    {
                        // 担当者済み or 一次検済
                        c1FlexGrid1.Rows[e.Row][0] = "7";
                    }
                    else if (c1FlexGrid1[e.Row, 4] != null)
                    {
                        try
                        {
                            dateTime = DateTime.Parse(c1FlexGrid1[e.Row, 4].ToString());
                            if (dateTime < DateTime.Today)
                            {
                                // 締切日経過
                                c1FlexGrid1.Rows[e.Row][0] = "1";
                            }
                            else if (dateTime < DateTime.Today.AddDays(3))
                            {
                                // 締切日が3日以内、かつ2次検証が完了していない
                                c1FlexGrid1.Rows[e.Row][0] = "2";
                            }
                            else if (dateTime < DateTime.Today.AddDays(7))
                            {
                                // 締切日が1週間以内、かつ2次検証が完了していない
                                c1FlexGrid1.Rows[e.Row][0] = "3";
                            }
                            else
                            {
                                c1FlexGrid1.Rows[e.Row][0] = "4";
                            }
                        }
                        catch
                        {
                            // 日付変換エラー
                            throw;
                        }
                    }

                }
            }
        }
        // Garoon送信ボタン
        private void GaroonBtn_Click(object sender, EventArgs e)
        {
            string methodName = ".GaroonBtn_Click";

            // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携以前の処理メッセージを残す
            //set_error("", 0);
            // I20003:Garoonとの連携を行います。
            if (MessageBox.Show(GlobalMethod.GetMessage("I20003", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                // エラーフラグ true:エラー false:正常
                Boolean errorFlg = false;
                // Garoon連携対象チェック
                if (!item1_GaroonRenkei.Checked)
                {
                    // I20005:Garoonとの連携対象ではありません。
                    set_error(GlobalMethod.GetMessage("I20005", ""));
                    errorFlg = true;
                }
                // 窓口担当者チェック
                if (item1_MadoguchiTantoushaCD.Text == "")
                {
                    // E20011:窓口担当者が未登録のため、Garoon連携ができません。
                    set_error(GlobalMethod.GetMessage("E20011", ""));
                    errorFlg = true;
                }

                // 連携処理
                if (errorFlg == false)
                {
                    string w_MadoguchiMailGaRenkeiKubun = "";
                    string w_MadoguchiMailMessageID = "";
                    string w_MadoguchiUketsukeBangou = "";
                    string w_MadoguchiUketsukeBangouEdaban = "";
                    string w_MadoguchiTantoushaCD = "";
                    string w_TokuchoBangou = "";
                    string w_MailInfoCSVWorkAtesakiUser = "";
                    string w_KojinCD = "";
                    string w_MadoguchiKanriGijutsusha = "";
                    string w_MadoguchiL1ChousaBushoCD = "";
                    string w_MadoguchiL1ChousaTantoushaCD = "";

                    string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                    using (var conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        var cmd = conn.CreateCommand();
                        SqlTransaction transaction = conn.BeginTransaction();
                        cmd.Transaction = transaction;
                        try
                        {
                            string historyMessage = "";
                            string GaroonRenkei = "";
                            if (item1_GaroonRenkei.Checked)
                            {
                                GaroonRenkei = "1";
                            }
                            else
                            {
                                GaroonRenkei = "0";
                            }

                            // I20006:Garoon送信ボタンからOKが押下されました。
                            historyMessage = GlobalMethod.GetMessage("I20006", "") + " ID:" + MadoguchiID + " Garoon連携区分:" + GaroonRenkei;

                            //履歴登録
                            cmd.CommandText = "INSERT INTO T_HISTORY(" +
                                "H_DATE_KEY " +
                                ",H_NO_KEY " +
                                ",H_OPERATE_DT " +
                                ",H_OPERATE_USER_ID " +
                                ",H_OPERATE_USER_MEI " +
                                ",H_OPERATE_USER_BUSHO_CD " +
                                ",H_OPERATE_USER_BUSHO_MEI " +
                                ",H_OPERATE_NAIYO " +
                                ",H_ProgramName " +
                                ",MadoguchiID " +
                                ",HistoryBeforeTantoubushoCD " +
                                ",HistoryBeforeTantoushaCD " +
                                ",HistoryAfterTantoubushoCD " +
                                ",HistoryAfterTantoushaCD " +
                                ",H_TOKUCHOBANGOU " +
                                ")VALUES(" +
                                "SYSDATETIME() " + 
                                ", " + GlobalMethod.getSaiban("HistoryID") + " " +
                                ",SYSDATETIME() " +
                                ",'" + UserInfos[0] + "' " +
                                ",N'" + UserInfos[1] + "' " +
                                ",'" + UserInfos[2] + "' " +
                                ",N'" + UserInfos[3] + "' " +
                                ",N'" + historyMessage + "'" +
                                ",'" + pgmName + methodName + "' " +
                                "," + MadoguchiID + " " +
                                ",NULL " +
                                ",NULL " +
                                ",NULL " +
                                ",NULL " +
                                ",N'" + Header1.Text + "' " +
                                ")";

                            cmd.ExecuteNonQuery();

                            var Dt = new System.Data.DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "MadoguchiGaroonRenkei,MadoguchiUketsukeBangou,MadoguchiUketsukeBangouEdaban,MadoguchiTantoushaCD,MadoguchiKanriGijutsusha " +
                              "FROM MadoguchiJouhou " +
                              "WHERE MadoguchiID = '" + MadoguchiID + "' ";

                            //データ取得
                            var sda = new SqlDataAdapter(cmd);
                            sda.Fill(Dt);

                            // MadoguchiJouhouに登録されているデータを取得
                            if (Dt != null && Dt.Rows.Count > 0)
                            {
                                w_MadoguchiMailGaRenkeiKubun = Dt.Rows[0][0].ToString();
                                w_MadoguchiUketsukeBangou = Dt.Rows[0][1].ToString();
                                w_MadoguchiUketsukeBangouEdaban = Dt.Rows[0][2].ToString();
                                w_MadoguchiTantoushaCD = Dt.Rows[0][3].ToString();
                                w_MadoguchiKanriGijutsusha = Dt.Rows[0][4].ToString();
                            }

                            // MadoguchiMailのIDを取得
                            string discript = "MadoguchiMailMessageID ";
                            string value = "TOP 1 MadoguchiMailMessageID ";
                            string table = "MadoguchiMail ";
                            string where = "MadoguchiMailTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_MadoguchiUketsukeBangou + "' AND MadoguchiMailTokuchoBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_MadoguchiUketsukeBangouEdaban + "' ";

                            // データ取得
                            DataTable tmpdt = GlobalMethod.getData(discript, value, table, where);
                            if (tmpdt != null && tmpdt.Rows.Count > 0)
                            {
                                w_MadoguchiMailMessageID = tmpdt.Rows[0][0].ToString();
                            }
                            else
                            {
                                // 取得できなかった場合は、0をセット（MailInfoCSVWorkMessageID は数値型の為、insert時に空文字をセットしようとするとエラーになる）
                                w_MadoguchiMailMessageID = "0";
                            }

                            // 特調番号
                            w_TokuchoBangou = w_MadoguchiUketsukeBangou + "-" + w_MadoguchiUketsukeBangouEdaban;

                            // 調査員取得
                            w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiTantoushaCD, w_MailInfoCSVWorkAtesakiUser);

                            // 管理技術者が存在すれば
                            if (w_MadoguchiKanriGijutsusha != "" && w_MadoguchiKanriGijutsusha != "0")
                            {
                                w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiKanriGijutsusha, w_MailInfoCSVWorkAtesakiUser);
                            }

                            Dt = new System.Data.DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "MadoguchiL1ChousaBushoCD,MadoguchiL1ChousaTantoushaCD " +
                              "FROM MadoguchiJouhouMadoguchiL1Chou " +
                              "WHERE MadoguchiID = '" + MadoguchiID + "' Order By MadoguchiL1ChousaCD";

                            //データ取得
                            sda = new SqlDataAdapter(cmd);
                            sda.Fill(Dt);

                            // MadoguchiJouhouに登録されているデータを取得
                            if (Dt != null && Dt.Rows.Count > 0)
                            {
                                for (int i = 0; i < Dt.Rows.Count; i++)
                                {
                                    //w_MadoguchiL1ChousaBushoCD = Dt.Rows[0][0].ToString();
                                    w_MadoguchiL1ChousaTantoushaCD = Dt.Rows[i][1].ToString();

                                    // 調査担当者が存在すれば
                                    if (w_MadoguchiL1ChousaTantoushaCD != "" && w_MadoguchiL1ChousaTantoushaCD != "0")
                                    {
                                        w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiL1ChousaTantoushaCD, w_MailInfoCSVWorkAtesakiUser);
                                    }
                                    // 支部応援マスタから担当調査員の部所に該当する調査員を設定する
                                    if (w_MadoguchiL1ChousaBushoCD != Dt.Rows[i][0].ToString())
                                    {
                                        w_MadoguchiL1ChousaBushoCD = Dt.Rows[i][0].ToString();
                                        w_MailInfoCSVWorkAtesakiUser = GetShibuouen(w_MadoguchiL1ChousaBushoCD, w_MailInfoCSVWorkAtesakiUser);
                                    }
                                }
                            }

                            // Garoon追加追加宛先の調査員も追加する
                            DataTable GaroonDt = new System.Data.DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "GaroonTsuikaAtesakiBushoCD,GaroonTsuikaAtesakiTantoushaCD " +
                              "FROM GaroonTsuikaAtesaki " +
                              "WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' ";

                            //データ取得
                            sda = new SqlDataAdapter(cmd);
                            sda.Fill(GaroonDt);

                            // GaroonTsuikaAtesakiに登録されているデータを取得
                            if (GaroonDt != null && GaroonDt.Rows.Count > 0)
                            {
                                for (int i = 0; i < GaroonDt.Rows.Count; i++)
                                {
                                    //w_MadoguchiL1ChousaBushoCD = Dt.Rows[0][0].ToString();
                                    w_MadoguchiL1ChousaTantoushaCD = GaroonDt.Rows[i][1].ToString();

                                    // 調査担当者が存在すれば
                                    if (w_MadoguchiL1ChousaTantoushaCD != "" && w_MadoguchiL1ChousaTantoushaCD != "0")
                                    {
                                        w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiL1ChousaTantoushaCD, w_MailInfoCSVWorkAtesakiUser);
                                    }
                                }
                            }

                            // メール情報CSVに追加するユーザーが空でない場合
                            if (w_MailInfoCSVWorkAtesakiUser != "")
                            {
                                // メール情報ワークの取得
                                Dt = new System.Data.DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "MailInfoCSVWorkID " +
                                  "FROM MailInfoCSVWork " +
                                  "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_TokuchoBangou + "' AND MailInfoCSVWorkCSVOutFlg = 0 AND MailInfoCSVWorkGaRenkeiFlg = 0 ";

                                //データ取得
                                sda = new SqlDataAdapter(cmd);
                                sda.Fill(Dt);

                                string MailInfoCSVWorkID = "";
                                // データの存在確認
                                if (Dt != null && Dt.Rows.Count > 0)
                                {
                                    MailInfoCSVWorkID = Dt.Rows[0][0].ToString();
                                    // 連携フラグにより、更新、削除を振り分ける
                                    if (w_MadoguchiMailGaRenkeiKubun == "1")
                                    {
                                        // 宛先を更新
                                        cmd.CommandText = "UPDATE MailInfoCSVWork SET " +
                                        "MailInfoCSVWorkAtesakiUser = '" + w_MailInfoCSVWorkAtesakiUser + "' " +
                                        ",MailInfoCSVWorkUpdateDate = SYSDATETIME() " +
                                        ",MailInfoCSVWorkUpdateUser = N'" + UserInfos[0] + "' " +
                                        ",MailInfoCSVWorkUpdateProgram = '" + pgmName + methodName + "' " +
                                        "Where MailInfoCSVWorkID = '" + MailInfoCSVWorkID + "' ";
                                        cmd.ExecuteNonQuery();
                                    }
                                    else
                                    {
                                        // 連携フラグがないので削除
                                        cmd.CommandText = "DELETE FROM MailInfoCSVWork " +
                                            "WHERE MailInfoCSVWorkID = '" + MailInfoCSVWorkID + "' ";
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                                else
                                {
                                    if (w_MadoguchiMailGaRenkeiKubun == "1")
                                    {
                                        // 存在しない場合、Insert
                                        cmd.CommandText = "INSERT INTO MailInfoCSVWork (" +
                                        "MailInfoCSVWorkID " +
                                        ",MailInfoCSVWorkMadoguchiID " +
                                        ",MailInfoCSVWorkTokuchoBangou " +
                                        ",MailInfoCSVWorkMessageID " +
                                        ",MailInfoCSVWorkAtesakiUser " +
                                        ",MailInfoCSVWorkCSVOutFlg " +
                                        ",MailInfoCSVWorkGaRenkeiFlg " +
                                        ",MailInfoCSVWorkCreateDate " +
                                        ",MailInfoCSVWorkCreateUser " +
                                        ",MailInfoCSVWorkCreateProgram " +
                                        ",MailInfoCSVWorkUpdateDate " +
                                        ",MailInfoCSVWorkUpdateUser " +
                                        ",MailInfoCSVWorkUpdateProgram " +
                                        ",MailInfoCSVWorkDeleteFlag " +
                                        ") VALUES (" +
                                        GlobalMethod.getSaiban("MailInfoCSVWorkID") +
                                        ",'" + MadoguchiID + "' " +
                                        ",N'" + w_TokuchoBangou + "' " +
                                        ",'" + w_MadoguchiMailMessageID + "' " +
                                        ",'" + w_MailInfoCSVWorkAtesakiUser + "' " +
                                        ",'0'" +
                                        ",'0'" +
                                        ",SYSDATETIME() " +                             // 登録日時
                                        ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                        ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                        ",SYSDATETIME() " +                             // 更新日時
                                        ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                        ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                        ",0 " +                                         // 削除フラグ
                                        ")";
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                            else
                            {
                                // 宛先が存在しない場合、削除する
                                cmd.CommandText = "DELETE FROM MailInfoCSVWork " +
                                                 "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_TokuchoBangou + "' AND MailInfoCSVWorkCSVOutFlg = 0 AND MailInfoCSVWorkGaRenkeiFlg = 0";
                                cmd.ExecuteNonQuery();
                            }

                            // 窓口情報の連携実行日時を更新
                            string datetTime = DateTime.Now.ToString();

                            cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                            "MadoguchiGaroonRenkeiJikouDate = '" + datetTime + "'" +
                            ",MadoguchiUpdateDate = SYSDATETIME()" +
                            ",MadoguchiUpdateUser = N'" + UserInfos[0] + "' " +
                            ",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
                            "Where MadoguchiID = '" + MadoguchiID + "' ";
                            cmd.ExecuteNonQuery();

                            // 更新日時の表記を更新
                            item1_GaroonUpdateDisp.Text = datetTime;

                            transaction.Commit();

                            GlobalMethod.outputLogger("GaroonBtn_Click", GlobalMethod.GetMessage("I20006", "") + " ID:" + MadoguchiID + " Garoon連携区分:" + item1_GaroonRenkei.ToString(), "insert", "DEBUG");
                            // I20004:Garoonとの連携を行いました。
                            set_error(GlobalMethod.GetMessage("I20004", ""));

                            conn.Close();
                        }
                        catch
                        {
                            transaction.Rollback();
                            throw;
                        }
                        finally
                        {
                            conn.Close();
                        }
                        cmd.Transaction = transaction;
                    }
                }
            }
        }

        // 個人コードを検索し、連結して返す
        private string SetChousain(string KojinCD, string MailInfoCSVWorkAtesakiUser)
        {
            string discript = "KojinCD ";
            string value = "KojinCD ";
            string table = "Mst_Chousain ";
            string where = "KojinCD = '" + KojinCD + "' ";
            string w_KojinCD = "";

            // データ取得
            var tmpdt = GlobalMethod.getData(discript, value, table, where);
            if (tmpdt != null && tmpdt.Rows.Count > 0)
            {
                if (MailInfoCSVWorkAtesakiUser == "")
                {
                    MailInfoCSVWorkAtesakiUser = tmpdt.Rows[0][0].ToString();
                }
                else
                {
                    w_KojinCD = tmpdt.Rows[0][0].ToString();
                    // MailInfoCSVWorkAtesakiUser が2048文字までなので、OVERする場合は、セットしない
                    if ((MailInfoCSVWorkAtesakiUser.Length + w_KojinCD.Length) <= 2048)
                    {
                        // 既に存在する場合は追加しない
                        if (MailInfoCSVWorkAtesakiUser.IndexOf(w_KojinCD) == -1)
                        {
                            MailInfoCSVWorkAtesakiUser = MailInfoCSVWorkAtesakiUser + ";" + w_KojinCD;
                        }
                    }
                    else
                    {
                        GlobalMethod.outputLogger("SetChousain", "ID:" + MadoguchiID + " Garoon連携で宛先ユーザーの文字数が2048を超える為、KojinCD:" + w_KojinCD + " を追加できませんでした。", "insert", "DEBUG");
                    }
                }
            }
            return MailInfoCSVWorkAtesakiUser;
        }

        // 支部応援の取得
        private string GetShibuouen(string w_MadoguchiL1ChousaBushoCD, string MailInfoCSVWorkAtesakiUser)
        {
            string discript = "ShibuouenKojinCD ";
            string value = "ShibuouenKojinCD ";
            string table = "Mst_Shibuouen ";
            //string where = "ShibuouenDeleteFlag = 0 Order By ShibuouenKojinCD ";
            string where = "(ShibuouenDeleteFlag = 0 or ShibuouenDeleteFlag = 1) Order By ShibuouenKojinCD ";
            string w_ShibuouenKojinCD = "";

            // データ取得
            var tmpdt = GlobalMethod.getData(discript, value, table, where);
            DataTable dt = new DataTable();

            if (tmpdt != null && tmpdt.Rows.Count > 0)
            {
                for (int i = 0; i < tmpdt.Rows.Count; i++)
                {
                    w_ShibuouenKojinCD = tmpdt.Rows[i][0].ToString();

                    discript = "KojinCD ";
                    value = "KojinCD ";
                    table = "Mst_Chousain ";
                    where = "GyoumuBushoCD = '" + w_MadoguchiL1ChousaBushoCD + "' AND KojinCD = '" + w_ShibuouenKojinCD + "' ";
                    dt = GlobalMethod.getData(discript, value, table, where);

                    // 存在する場合のみ追加する
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        if (MailInfoCSVWorkAtesakiUser == "")
                        {
                            MailInfoCSVWorkAtesakiUser = dt.Rows[0][0].ToString();
                        }
                        else
                        {
                            // 既に存在する場合は追加しない
                            if (MailInfoCSVWorkAtesakiUser.IndexOf(dt.Rows[0][0].ToString()) == -1)
                            {
                                MailInfoCSVWorkAtesakiUser = MailInfoCSVWorkAtesakiUser + ";" + dt.Rows[0][0].ToString();
                            }
                        }
                    }
                }
            }
            return MailInfoCSVWorkAtesakiUser;
        }

        private Boolean bool_str(String str)
        {
            //文字列0と1をBooleanに変換
            Boolean checkValue = true;

            //strがnullか空なら0扱い
            if (String.IsNullOrEmpty(str))
            {
                str = "0";
            }

            //zeroOneが0のとき
            if ("0".Equals(str))
            {
                checkValue = false;
            }
            else
            {
                checkValue = true;
            }
            return checkValue;
        }

        private void button3_ExcelHoukokusho_Click(object sender, EventArgs e)
        {
            // 報告書プロンプト
            Popup_HoukokuSho form = new Popup_HoukokuSho();
            form.MadoguchiID = MadoguchiID;
            //form.MENU_ID = 303;
            form.UserInfos = UserInfos;
            form.PrintGamen = "Tokumei";
            form.BushoCD = "";
            if (src_Busho.Text != "")
            {
                form.BushoCD = src_Busho.SelectedValue.ToString();
            }
            form.Chousain = src_HinmokuChousain.Text;
            form.ShuFuku = src_ShuFuku.SelectedIndex;
            form.Hinmei = src_ChousaHinmei.Text;
            form.Kikaku = src_ChousaKikaku.Text;
            form.Zaikou = int.Parse(src_Zaikou.SelectedValue.ToString());
            form.KuhakuList = int.Parse(src_TantoushaKuuhaku.SelectedValue.ToString());
            form.Memo1 = item3_Memo1.Text;
            form.Memo2 = item3_Memo2.Text;
            form.ShowDialog();

            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {

                //item_Hyoujikensuu.SelectedIndex = 1;
                item_Hyoujikensuu.SelectedIndex = 4; // 1000件対応
                // データ取り直し
                get_data(3);
            }
        }

        // 調査品目明細タブ
        private void c1FlexGrid4_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid4.HitTest(new Point(e.X, e.Y));

            // 列名の取得
            string ColName = c1FlexGrid4.Cols[hti.Column].Name;

            if (hti.Row > 1 & e.Button == MouseButtons.Right)
            {
                //c1FlexGrid4.Select(hti.Row, hti.Column);
                BushoTantouRow = hti.Row;
                BushoTantouColumn = hti.Column;

                contextMenuStrip1.Items.Clear();

                // 編集状態が1:編集の場合
                if (ChousaHinmokuMode == 1)
                {
                    //if (hti.Column == 42 || hti.Column == 44 || hti.Column == 46)
                    if (ColName == "HinmokuRyakuBushoCD" || ColName == "HinmokuRyakuBushoFuku1CD" || ColName == "HinmokuRyakuBushoFuku2CD")
                    {
                        DataTable dt = new DataTable();
                        //受託課所支部
                        using (var conn = new SqlConnection(connStr))
                        {
                            var cmd = conn.CreateCommand();

                            //データ取得時に年度がいない場合、当年度とする
                            int Nendo;
                            int ToNendo;
                            if (item1_TourokuNendo.Text == "")
                            {
                                Nendo = DateTime.Today.Year;
                                ToNendo = DateTime.Today.AddYears(1).Year;
                            }
                            else
                            {
                                int.TryParse(item1_TourokuNendo.Text.ToString(), out Nendo);
                                ToNendo = Nendo + 1;
                            }
                            //cmd.CommandText = "SELECT " +
                            //"GyoumuBushoCD  " +
                            //",BushokanriboKameiRaku  " +
                            //"FROM Mst_Busho  " +
                            //"WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                            //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                            //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                            //"ORDER BY BushoMadoguchiNarabijun";

                            cmd.CommandText = "SELECT " +
                            "GyoumuBushoCD  " +
                            ",BushokanriboKameiRaku  " +
                            "FROM Mst_Busho  " +
                            "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  ";
                            //// 今日日付の年度データを検索する際は、今日有効な部所を表示
                            //if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                            //{
                            //    cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                            //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today + "' ) ";
                            //}
                            //else
                            //{
                            //    cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                            //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            //}
                            //cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                            ////"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                            cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                            "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                            cmd.CommandText += "ORDER BY BushoMadoguchiNarabijun";

                            var sda = new SqlDataAdapter(cmd);
                            dt.Clear();
                            sda.Fill(dt);
                            conn.Close();
                        }
                        contextMenuBusho = new ToolStripMenuItem();

                        contextMenuBusho.Text = "部所";
                        Set_ContextMenu(contextMenuBusho, dt);

                        // 部所
                        contextMenuStrip1.Items.Add(contextMenuBusho);
                        contextMenuStrip1.Items.Add(contextMenuBushoClear);
                    }
                    //else if (hti.Column == 41 || hti.Column == 43 || hti.Column == 45)
                    //else if (hti.Column == 43 || hti.Column == 45 || hti.Column == 47)
                    else if (ColName == "HinmokuChousainCD" || ColName == "HinmokuFukuChousainCD1" || ColName == "HinmokuFukuChousainCD2")
                    {
                        contextMenuTantoushaBusho = new ToolStripMenuItem();
                        contextMenuTantousha = new ToolStripMenuItem();

                        contextMenuTantoushaBusho.Text = "部所";
                        contextMenuTantousha.Text = "担当者";

                        string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                        using (var conn = new SqlConnection(connStr))
                        {
                            var cmd = conn.CreateCommand();
                            DataTable dt = new DataTable();

                            String bushoQuery = "";

                            //データ取得時に年度がいない場合、当年度とする
                            int Nendo;
                            int ToNendo;
                            if (item1_TourokuNendo.Text == "")
                            {
                                Nendo = DateTime.Today.Year;
                                ToNendo = DateTime.Today.AddYears(1).Year;
                            }
                            else
                            {
                                int.TryParse(item1_TourokuNendo.Text.ToString(), out Nendo);
                                ToNendo = Nendo + 1;
                            }
                            // 一つ左の部所を見る
                            //if (c1FlexGrid4.Rows[hti.Row][hti.Column - 1] != null && c1FlexGrid4.Rows[hti.Row][hti.Column - 1].ToString() != "")
                            //{
                            //    // 空でない場合
                            //    cmd.CommandText = "SELECT " +
                            //        "GyoumuBushoCD  " +
                            //        ",BushokanriboKameiRaku  " +
                            //        "FROM Mst_Busho  " +
                            //        "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                            //        " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                            //        " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                            //        " AND GyoumuBushoCD = '" + c1FlexGrid4.Rows[hti.Row][hti.Column - 1].ToString() + "' " +
                            //        "ORDER BY BushoMadoguchiNarabijun";
                            //}
                            //else
                            //{
                            //    // 空の場合
                            //    cmd.CommandText = "SELECT " +
                            //        "GyoumuBushoCD  " +
                            //        ",BushokanriboKameiRaku  " +
                            //        "FROM Mst_Busho  " +
                            //        "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                            //        " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                            //        " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                            //        "ORDER BY BushoMadoguchiNarabijun";

                            //    bushoQuery = "SELECT " +
                            //        "GyoumuBushoCD  " +
                            //        "FROM Mst_Busho  " +
                            //        "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                            //        " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                            //        " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            //}

                            // 部所は全部出す
                            //cmd.CommandText = "SELECT " +
                            //    "GyoumuBushoCD  " +
                            //    ",BushokanriboKameiRaku  " +
                            //    "FROM Mst_Busho  " +
                            //    "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                            //    " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                            //    " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                            //    "ORDER BY BushoMadoguchiNarabijun";

                            cmd.CommandText = "SELECT " +
                                "GyoumuBushoCD  " +
                                ",BushokanriboKameiRaku  " +
                                "FROM Mst_Busho  " +
                                "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  ";

                            //    // 今日日付の年度データを検索する際は、今日有効な部所を表示
                            //    if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                            //    {
                            //        cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                            //        "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today + "' ) ";
                            //    }
                            //    else
                            //    {
                            //        cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                            //        "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            //}
                            //cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                            ////"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                            cmd.CommandText += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                            "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                            cmd.CommandText += "ORDER BY BushoMadoguchiNarabijun";

                            //bushoQuery = "SELECT " +
                            //    "GyoumuBushoCD  " +
                            //    "FROM Mst_Busho  " +
                            //    "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                            //    " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                            //    " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";

                            bushoQuery = "SELECT " +
                                "GyoumuBushoCD  " +
                                //",BushokanriboKameiRaku  " +
                                "FROM Mst_Busho  " +
                                "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  ";

                            //// 今日日付の年度データを検索する際は、今日有効な部所を表示
                            //if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                            //{
                            //    bushoQuery += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                            //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today + "' ) ";
                            //}
                            //else
                            //{
                            //    bushoQuery += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                            //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            //}
                            //bushoQuery += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                            ////"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                            bushoQuery += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                            "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                            //bushoQuery += "ORDER BY BushoMadoguchiNarabijun";

                            var sda = new SqlDataAdapter(cmd);
                            dt.Clear();
                            sda.Fill(dt);
                            Set_ContextBushoMenu(contextMenuTantoushaBusho, dt);

                            // 一つ左の部所を見る
                            if (c1FlexGrid4.Rows[hti.Row][hti.Column - 1] != null && c1FlexGrid4.Rows[hti.Row][hti.Column - 1].ToString() != "")
                            {
                                // 空でない場合
                                cmd.CommandText = "SELECT " +
                                    "KojinCD " +
                                    ",ChousainMei " +
                                    "FROM Mst_Chousain " +
                                    "WHERE RetireFLG = 0 AND TokuchoFLG = 1 " +
                                    "AND GyoumuBushoCD = '" + c1FlexGrid4.Rows[hti.Row][hti.Column - 1].ToString() + "' ";

                                // 今日日付の年度データを検索する際は、今日有効な調査員を表示
                                if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                                {
                                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today + "' ) ";
                                }
                                else
                                {
                                    //cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                                    //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                                }
                                //"ORDER BY ChousainMei ";
                                cmd.CommandText += "ORDER BY KojinCD ";
                            }
                            else
                            {
                                // 空の場合
                                cmd.CommandText = "SELECT " +
                                    "KojinCD " +
                                    ",ChousainMei " +
                                    "FROM Mst_Chousain " +
                                    "WHERE RetireFLG = 0 AND TokuchoFLG = 1 " +
                                    "AND GyoumuBushoCD in (" + bushoQuery + ") ";

                                // 今日日付の年度データを検索する際は、今日有効な調査員を表示
                                if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                                {
                                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today + "' ) ";
                                }
                                else
                                {
                                    //cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                                    //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                                    cmd.CommandText += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                                    "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                                }
                                //"ORDER BY ChousainMei ";
                                cmd.CommandText += "ORDER BY KojinCD ";
                            }
                            DataTable dt2 = new DataTable();

                            sda = new SqlDataAdapter(cmd);
                            sda.Fill(dt2);
                            Set_ContextMenu(contextMenuTantousha, dt2);
                            conn.Close();
                        }

                        // 調査担当者
                        contextMenuStrip1.Items.Add(contextMenuTantoushaBusho); // 部所
                        contextMenuStrip1.Items.Add(contextMenuTantousha);      // 担当者
                        contextMenuStrip1.Items.Add(contextMenuTantoushaBushoClear);     // 部所クリア
                        contextMenuStrip1.Items.Add(contextMenuTantoushaTantoushaClear); // 担当者クリア
                    }
                    //else if (hti.Column == 47)
                    //else if (hti.Column == 49)
                    else if (ColName == "ChousaHoukokuRank")
                    {
                        // 報告ランク
                        contextMenuStrip1.Items.Add(contextMenuHoukoku);
                        contextMenuStrip1.Items.Add(contextMenuHoukokuClear);
                    }
                    //else if (hti.Column == 49)
                    //else if (hti.Column == 51)
                    else if (ColName == "ChousaIraiRank")
                    {
                        // 依頼ランク
                        contextMenuStrip1.Items.Add(contextMenuIrai);
                        contextMenuStrip1.Items.Add(contextMenuIraiClear);
                    }
                    contextMenuStrip1.Items.Add(contextMenuCopy);     // コピー
                    contextMenuStrip1.Items.Add(contextMenuPaste); // 貼り付け
                    contextMenuStrip1.Show(Cursor.Position.X, Cursor.Position.Y);
                }
                else
                {
                    contextMenuStrip1.Items.Add(contextMenuCopy);     // コピー
                    contextMenuStrip1.Show(Cursor.Position.X, Cursor.Position.Y);
                }
            }

            // クリック時
            if (hti.Row > 1 & e.Button == MouseButtons.Left)
            {
                // 編集状態が1:編集の場合
                if (ChousaHinmokuMode == 1)
                {
                    // 追加
                    //if (hti.Column == 0)
                    if (ColName == "Add1")
                    {
                        // I20311:行をコピーし追加しますがよろしいですか。
                        if (MessageBox.Show(GlobalMethod.GetMessage("I20311", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                        {
                            GlobalMethod.outputLogger("Madoguhi_ChousaHinmoku lineAdd", "追加行(表示):" + (hti.Row + 1), "AddLine", "DEBUG");
                            // 行挿入
                            c1FlexGrid4.Rows.Insert(hti.Row + 1);

                            // 値をコピーする
                            // hti.Row コピー元の行
                            // hti.Row + 1 追加した行
                            for (int i = 0; i < c1FlexGrid4.Cols.Count; i++)
                            {
                                //// ChousaHinmokuID
                                //if (i == 55)
                                //{
                                //    // 追加時にIDを振っておく
                                //    c1FlexGrid4.Rows[hti.Row + 1][i] = GlobalMethod.getSaiban("HinmokuMeisaiID");
                                //}
                                //// 個別順
                                //else if (i == 7)
                                //{
                                //    // 行のコピーをした際に、
                                //    // 全体順の括りの中で、個別順の最大値 + 1を取得し、コピーした行の個別順にセットする
                                //    float num = 0;
                                //    // コピー元の全体順を取得
                                //    string zentai = c1FlexGrid4.Rows[hti.Row][i - 1].ToString();
                                //    float kobetsuMaxNum = 0;

                                //    for (int j = 2; j < c1FlexGrid4.Rows.Count; j++)
                                //    {
                                //        if (zentai == c1FlexGrid4.Rows[j][i - 1].ToString())
                                //        {
                                //            // 個別順の最大値を取り出す
                                //            if (c1FlexGrid4.Rows[j][i] != null && float.TryParse(c1FlexGrid4.Rows[j][i].ToString(), out num))
                                //            {
                                //                if (num > kobetsuMaxNum)
                                //                {
                                //                    kobetsuMaxNum = num;
                                //                }
                                //            }
                                //        }
                                //    }
                                //    // 個別順に+1した値をセット
                                //    c1FlexGrid4.Rows[hti.Row + 1][i] = kobetsuMaxNum + 1;
                                //}
                                //// 価格
                                //else if (i == 13)
                                //{
                                //    //c1FlexGrid4.Rows[hti.Row + 1][i] = 0;
                                //    c1FlexGrid4.Rows[hti.Row + 1][i] = "";
                                //}
                                //// ChousaShinchokuJoukyou
                                //else if (i == 56)
                                //{
                                //    // 進捗状況は、20:調査開始
                                //    c1FlexGrid4.Rows[hti.Row + 1][i] = 20;
                                //}
                                //// 0:Insert/1:Select/2:Update
                                //else if (i == 57)
                                //{
                                //    c1FlexGrid4.Rows[hti.Row + 1][i] = "0";
                                //}
                                //else
                                //{
                                //    // コピー
                                //    c1FlexGrid4.Rows[hti.Row + 1][i] = c1FlexGrid4.Rows[hti.Row][i];
                                //}

                                switch (c1FlexGrid4.Cols[i].Name)
                                {
                                    // ChousaHinmokuID
                                    case "ChousaHinmokuID2":
                                        // 追加時にIDを振っておく
                                        c1FlexGrid4.Rows[hti.Row + 1]["ChousaHinmokuID2"] = GlobalMethod.getSaiban("HinmokuMeisaiID");
                                        break;
                                    // 個別順
                                    case "ChousaKobetsuJun":
                                        // 行のコピーをした際に、
                                        // 全体順の括りの中で、個別順の最大値 + 1を取得し、コピーした行の個別順にセットする
                                        float num = 0;
                                        // コピー元の全体順を取得
                                        string zentai = c1FlexGrid4.Rows[hti.Row]["ChousaZentaiJun"].ToString();
                                        float kobetsuMaxNum = 0;

                                        for (int j = 2; j < c1FlexGrid4.Rows.Count; j++)
                                        {
                                            if (zentai == c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString())
                                            {
                                                // 個別順の最大値を取り出す
                                                if (c1FlexGrid4.Rows[j]["ChousaKobetsuJun"] != null && float.TryParse(c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString(), out num))
                                                {
                                                    if (num > kobetsuMaxNum)
                                                    {
                                                        kobetsuMaxNum = num;
                                                    }
                                                }
                                            }
                                        }
                                        // 個別順に+1した値をセット
                                        c1FlexGrid4.Rows[hti.Row + 1]["ChousaKobetsuJun"] = kobetsuMaxNum + 1;
                                        break;
                                    // 価格
                                    case "ChousaKakaku":
                                        c1FlexGrid4.Rows[hti.Row + 1]["ChousaKakaku"] = "";
                                        break;
                                    // 進捗状況
                                    case "ChousaShinchokuJoukyou":
                                        // 進捗状況は、20:調査開始
                                        c1FlexGrid4.Rows[hti.Row + 1]["ChousaShinchokuJoukyou"] = 20;
                                        break;
                                    // 0:Insert/1:Select/2:Update
                                    case "Mode":
                                        c1FlexGrid4.Rows[hti.Row + 1]["Mode"] = "0";
                                        break;
                                    default:
                                        // コピー
                                        c1FlexGrid4.Rows[hti.Row + 1][i] = c1FlexGrid4.Rows[hti.Row][i];
                                        break;
                                }

                            }
                            // 進捗状況を判定する（進捗状況は20:調査開始固定なので、進捗状況の以外を判定する）
                            DateTime dateTime = DateTime.Today;
                            // 1:報告済みの場合
                            if ("1".Equals(MadoguchiHoukokuzumi))
                            {
                                // 報告済み
                                //c1FlexGrid4.Rows[hti.Row + 1][5] = "8";
                                c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "8";
                            }
                            else
                            {
                                try
                                {
                                    // 締切日
                                    //if (c1FlexGrid4[hti.Row, 52] != null) {
                                    if (c1FlexGrid4.Rows[hti.Row]["ChousaHinmokuShimekiribi"] != null)
                                    {
                                        //dateTime = DateTime.Parse(c1FlexGrid4[hti.Row, 52].ToString());
                                        dateTime = DateTime.Parse(c1FlexGrid4.Rows[hti.Row]["ChousaHinmokuShimekiribi"].ToString());
                                        if (dateTime < DateTime.Today)
                                        {
                                            // 締切日経過
                                            //c1FlexGrid4.Rows[hti.Row + 1][5] = "1";
                                            c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "1";
                                        }
                                        else if (dateTime < DateTime.Today.AddDays(3))
                                        {
                                            // 締切日が3日以内、かつ2次検証が完了していない
                                            //c1FlexGrid4.Rows[hti.Row + 1][5] = "2";
                                            c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "2";
                                        }
                                        else if (dateTime < DateTime.Today.AddDays(7))
                                        {
                                            // 締切日が1週間以内、かつ2次検証が完了していない
                                            //c1FlexGrid4.Rows[hti.Row + 1][5] = "3";
                                            c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "3";
                                        }
                                        else
                                        {
                                            //c1FlexGrid4.Rows[hti.Row + 1][5] = "4";
                                            c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "4";
                                        }
                                    }
                                }
                                catch
                                {
                                    // 日付変換エラー
                                    throw;
                                }
                            }

                        }
                    }
                    // 削除
                    //if (hti.Column == 1)
                    if (ColName == "Delete1")
                    {
                        // I20312:行を削除しますがよろしいですか。
                        if (MessageBox.Show(GlobalMethod.GetMessage("I20312", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                        {
                            writeHistory("【開始】調査品目明細の個別削除を開始します。 ID= :" + MadoguchiID);

                            GlobalMethod.outputLogger("Madoguhi_ChousaHinmoku lineDelete", "削除行(表示):" + (hti.Row), "linedelete", "DEBUG");

                            //if (c1FlexGrid4.Rows[hti.Row][55] != null && c1FlexGrid4.Rows[hti.Row][55].ToString() != "")
                            if (c1FlexGrid4.Rows[hti.Row]["ChousaHinmokuID2"] != null && c1FlexGrid4.Rows[hti.Row]["ChousaHinmokuID2"].ToString() != "")
                            {
                                if (deleteChousaHinmokuIDs == "")
                                {
                                    //deleteChousaHinmokuIDs = c1FlexGrid4.Rows[hti.Row][55].ToString();
                                    deleteChousaHinmokuIDs = c1FlexGrid4.Rows[hti.Row]["ChousaHinmokuID2"].ToString();
                                }
                                else
                                {
                                    //deleteChousaHinmokuIDs += "," + c1FlexGrid4.Rows[hti.Row][55].ToString();
                                    deleteChousaHinmokuIDs += "," + c1FlexGrid4.Rows[hti.Row]["ChousaHinmokuID2"].ToString();
                                }
                            }

                            // 行の削除
                            c1FlexGrid4.RemoveItem(hti.Row);
                        }
                    }

                    // 表示件数（ヘッダー2行分を引く）
                    Grid_Num.Text = "(" + (c1FlexGrid4.Rows.Count - 2) + ")";
                }

                // フォルダアイコン
                if (ColName == "ChousaLinkSakli")
                {
                    if (c1FlexGrid4.Rows[hti.Row]["ChousaLinkSakli"] != null)
                    {
                        switch (c1FlexGrid4.Rows[hti.Row]["ChousaLinkSakli"].ToString())
                        {
                            // アイコン 0:グレー 1:イエロー 2:Excleアイコン
                            case "1":
                                // 集計表フォルダが存在すれば開く
                                if (Directory.Exists(ShukeiHyoFolder))
                                {
                                    System.Diagnostics.Process.Start(ShukeiHyoFolder);
                                }
                                break;
                            case "2":
                                if (File.Exists(c1FlexGrid4[hti.Row, hti.Column + 1].ToString()))
                                {
                                    System.Diagnostics.Process.Start(c1FlexGrid4[hti.Row, hti.Column + 1].ToString());
                                }
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            // フォルダアイコン
            //if (hti.Row > 1 & hti.Column == 38)
            //if (hti.Row > 1 & hti.Column == 40)
            //{
            //    // アイコン 0:グレー 1:イエロー 2:Excleアイコン
            //    // 集計表フォルダ表示
            //    if (c1FlexGrid4[hti.Row, hti.Column].ToString() == "1")
            //    {
            //        // ChousaLinkSakli のフォルダが存在すればフォルダを開く
            //        if (Directory.Exists(c1FlexGrid4[hti.Row, hti.Column + 1].ToString()))
            //        {
            //            System.Diagnostics.Process.Start(ShukeiHyoFolder);
            //        }
            //        // 集計表フォルダが存在すれば開く
            //        else if (Directory.Exists(ShukeiHyoFolder))
            //        {
            //            System.Diagnostics.Process.Start(ShukeiHyoFolder);
            //        }
            //        else
            //        {
            //            c1FlexGrid4[hti.Row, hti.Column] = 0;
            //        }
            //    }
            //    // Excleアイコン
            //    if (c1FlexGrid4[hti.Row, hti.Column].ToString() == "2")
            //    {
            //        if (File.Exists(c1FlexGrid4[hti.Row, hti.Column + 1].ToString()))
            //        {
            //            System.Diagnostics.Process.Start(c1FlexGrid4[hti.Row, hti.Column + 1].ToString());
            //        }
            //        else
            //        {
            //            c1FlexGrid4[hti.Row, hti.Column] = 0;
            //        }
            //    }
            //}
        }

        // 調査品目明細Gridの値変更後イベント
        private void c1FlexGrid4_CellChanged(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            // コピーアンドペーストの貼り付けは当イベントぐらいでしか拾えない
            // 調査品目編集モード 0:表示 1:編集
            if (ChousaHinmokuMode == 1)
            {
                // 列名の取得
                string ColName = c1FlexGrid4.Cols[e.Col].Name;

                // コピペで空文字を貼り付けた場合、NULL扱いとなっており、更新処理でエラーが発生するので、NULLを空文字に変換しておく
                // 6:全体順 53:報告済
                //if (e.Col >= 6 && e.Col <= 53)
                // コピペ対象の項目の場合に処理する
                if (ColName == "RowChange" || ColName == "Add1" || ColName == "Delete1" || ColName == "Add2" || ColName == "Delete2"
                    || ColName == "ChousaHinmokuID" || ColName == "ShinchokuIcon" || ColName == "ChousaDeleteFlag" || ColName == "ChousaHinmokuID2"
                    || ColName == "ChousaShinchokuJoukyou" || ColName == "Mode" || ColName == "ColumnSort")
                {
                }
                else
                {
                    if (c1FlexGrid4.Rows[e.Row][e.Col] == null)
                    {
                        // NULL
                        c1FlexGrid4.Rows[e.Row][e.Col] = "";
                    }
                }

                // 57:0:Insert/1:Select/2:Update
                //if (c1FlexGrid4.Rows[e.Row][57] != null && c1FlexGrid4.Rows[e.Row][57].ToString() == "1")
                if (c1FlexGrid4.Rows[e.Row]["Mode"] != null && c1FlexGrid4.Rows[e.Row]["Mode"].ToString() == "1")
                {
                    // 更新フラグを立てておく
                    //c1FlexGrid4.Rows[e.Row][57] = "2";
                    c1FlexGrid4.Rows[e.Row]["Mode"] = "2";
                }

                DataTable combodt = new DataTable();
                string discript = "GyoumuBushoCD";
                string value = "GyoumuBushoCD";
                string table = "Mst_Busho";
                string where = "";

                //データ取得時に年度がいない場合、当年度とする
                int Nendo;
                int ToNendo;
                if (item1_TourokuNendo.Text == "")
                {
                    Nendo = DateTime.Today.Year;
                    ToNendo = DateTime.Today.AddYears(1).Year;
                }
                else
                {
                    int.TryParse(item1_TourokuNendo.Text.ToString(), out Nendo);
                    ToNendo = Nendo + 1;
                }

                // 42:調査担当部所、44:副調査担当部所1、46:副調査担当部所2
                //if (e.Col == 42 || e.Col == 44 || e.Col == 46)
                if (ColName == "HinmokuRyakuBushoCD" || ColName == "HinmokuRyakuBushoFuku1CD" || ColName == "HinmokuRyakuBushoFuku2CD")
                {
                    if (c1FlexGrid4.Rows[e.Row][e.Col] != null && c1FlexGrid4.Rows[e.Row][e.Col].ToString() != "")
                    {
                        if (System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[e.Row][e.Col].ToString().Replace(Environment.NewLine, ""), @"^[0-9]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                        {
                            where = "GyoumuBushoCD = '" + c1FlexGrid4.Rows[e.Row][e.Col].ToString() + "' ";
                        }
                        else
                        {
                            where = "BushokanriboKameiRaku COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + c1FlexGrid4.Rows[e.Row][e.Col].ToString() + "' ";
                        }

                        where += " AND BushoDeleteFlag = 0 AND BushoMadoguchiHyoujiFlg = 1 ";
                        //// 今日日付の年度データを検索する際は、今日有効な部所を表示
                        //if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                        //{
                        //    where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                        //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + DateTime.Today + "' ) ";
                        //}
                        //else
                        //{
                        //    where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                        //    "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                        //}
                        //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                        ////"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                        //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                        where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                        "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                        combodt = new DataTable();
                        combodt = GlobalMethod.getData(discript, value, table, where);
                        if (combodt != null && combodt.Rows.Count > 0)
                        {
                            // 取得した部所をセット
                            c1FlexGrid4.Rows[e.Row][e.Col] = combodt.Rows[0][0].ToString();
                        }
                        else
                        {
                            // 部所をクリア
                            c1FlexGrid4.Rows[e.Row][e.Col] = "";
                        }
                    }
                    // 部所に対する調査員をセットする
                    string ColName2 = "";
                    switch (ColName)
                    {
                        case "HinmokuRyakuBushoCD":
                            ColName2 = "HinmokuChousainCD";
                            break;
                        case "HinmokuRyakuBushoFuku1CD":
                            ColName2 = "HinmokuFukuChousainCD1";
                            break;
                        case "HinmokuRyakuBushoFuku2CD":
                            ColName2 = "HinmokuFukuChousainCD2";
                            break;
                        default:
                            break;
                    }

                    // 部所変更時にユーザーが部所に所属していなければクリア
                    //if (c1FlexGrid4.Rows[e.Row][e.Col + 1] != null && c1FlexGrid4.Rows[e.Row][e.Col + 1].ToString() != "")
                    if (c1FlexGrid4.Rows[e.Row][ColName2] != null && c1FlexGrid4.Rows[e.Row][ColName2].ToString() != "")
                    {
                        combodt = new DataTable();
                        discript = "KojinCD";
                        value = "GyoumuBushoCD";
                        table = "Mst_Chousain";
                        where += "GyoumuBushoCD = '" + c1FlexGrid4.Rows[e.Row][e.Col].ToString() + "' ";

                        // 今日日付の年度データを検索する際は、今日有効な調査員を表示
                        if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                        {
                            where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                            "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today + "' ) ";
                        }
                        else
                        {
                            //where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                            //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                            "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                        }

                        combodt = new DataTable();
                        combodt = GlobalMethod.getData(discript, value, table, where);
                        if (combodt != null && combodt.Rows.Count > 0)
                        {
                            // 部所でユーザーが取得できるので問題ない
                        }
                        else
                        {
                            // 部所に所属したユーザーではないのでクリア
                            //c1FlexGrid4.Rows[e.Row][e.Col + 1] = "";
                            c1FlexGrid4.Rows[e.Row][ColName2] = "";
                        }
                    }
                    // 調査担当者のコンボを選択している部所で絞る
                    discript = "ChousainMei ";
                    value = "KojinCD ";
                    table = "Mst_Chousain ";
                    where = "RetireFLG = 0 AND TokuchoFLG = 1 ";
                    // 部所が空でない場合
                    if (c1FlexGrid4.Rows[e.Row][e.Col] != null && c1FlexGrid4.Rows[e.Row][e.Col].ToString() != "")
                    {
                        where += "AND GyoumuBushoCD = '" + c1FlexGrid4.Rows[e.Row][e.Col].ToString() + "' ";
                    }
                    else
                    {
                        where += "AND GyoumuBushoCD in (select GyoumuBushoCD from Mst_Busho where BushoDeleteFlag = 0 AND BushoMadoguchiHyoujiFlg = 1) ";
                    }
                    // 今日日付の年度データを検索する際は、今日有効な調査員を表示
                    if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                    {
                        where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                        "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today + "' ) ";
                    }
                    else
                    {
                        //where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                        //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                        where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                        "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                    }
                    //コンボボックスデータ取得
                    DataTable tmpdt11 = GlobalMethod.getData(discript, value, table, where);
                    SortedList sl = new SortedList();
                    sl = GlobalMethod.Get_SortedList(tmpdt11);

                    // 最初にパイプ文字"|"を記述してしまうと入力可になってしまう
                    String comboListValue = " ";

                    if (tmpdt11 != null && tmpdt11.Rows.Count > 0)
                    {
                        for (int i = 0; i < tmpdt11.Rows.Count; i++)
                        {
                            comboListValue += "|" + tmpdt11.Rows[i][1].ToString();
                        }
                    }

                    int TantoushaIndex = c1FlexGrid4.Cols[ColName2].Index;

                    //C1.Win.C1FlexGrid.CellStyle cs1 = c1FlexGrid4.Styles.Add("Combo" + e.Row + "_" + e.Col + 1); // スタイルを定義
                    C1.Win.C1FlexGrid.CellStyle cs1 = c1FlexGrid4.Styles.Add("Combo" + e.Row + "_" + TantoushaIndex); // スタイルを定義
                    cs1.ComboList = comboListValue; // ComboListを設定

                    //C1.Win.C1FlexGrid.CellRange rg1 = c1FlexGrid4.GetCellRange(e.Row, e.Col + 1); // セルを選択
                    C1.Win.C1FlexGrid.CellRange rg1 = c1FlexGrid4.GetCellRange(e.Row, TantoushaIndex); // セルを選択
                    rg1.Style = cs1; // スタイルを割り当てる
                }

                combodt = new DataTable();
                discript = "KojinCD";
                value = "GyoumuBushoCD";
                table = "Mst_Chousain";
                where = "";

                // コピーペースト でユーザーを貼り付けた際に部所を正しい部所に変更する
                // 右クリックから調査員を選択する場合は、調査員名がそのままくる
                // コピーアンドペーストだとCDが渡ってくる
                // 43:調査担当者 45:副調査担当者1 47:副調査担当者2
                //if (e.Col == 43 || e.Col == 45 || e.Col == 47) 
                if (ColName == "HinmokuChousainCD" || ColName == "HinmokuFukuChousainCD1" || ColName == "HinmokuFukuChousainCD2")
                {

                    // 調査員に対する部所をセットする
                    string ColName2 = "";
                    switch (ColName)
                    {
                        case "HinmokuChousainCD":
                            ColName2 = "HinmokuRyakuBushoCD";
                            break;
                        case "HinmokuFukuChousainCD1":
                            ColName2 = "HinmokuRyakuBushoFuku1CD";
                            break;
                        case "HinmokuFukuChousainCD2":
                            ColName2 = "HinmokuRyakuBushoFuku2CD";
                            break;
                        default:
                            break;
                    }
                    if (c1FlexGrid4.Rows[e.Row][e.Col] != null && c1FlexGrid4.Rows[e.Row][e.Col].ToString() != "")
                    {
                        if (System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[e.Row][e.Col].ToString().Replace(Environment.NewLine, ""), @"^[0-9]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                        {
                            where = "KojinCD = '" + c1FlexGrid4.Rows[e.Row][e.Col].ToString() + "' ";
                        }
                        else
                        {
                            where = "ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + c1FlexGrid4.Rows[e.Row][e.Col].ToString() + "' ";
                        }

                        where += " AND RetireFLG = 0 AND TokuchoFLG = 1";
                        // 今日日付の年度データを検索する際は、今日有効な調査員を表示
                        if (DateTime.Today <= new DateTime(ToNendo, 3, 31) && DateTime.Today >= new DateTime(Nendo, 4, 1))
                        {
                            where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + DateTime.Today + "' ) " +
                            "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + DateTime.Today + "' ) ";
                        }
                        else
                        {
                            //where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + Nendo + "/4/1' ) " +
                            //"AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                            where += "AND (ChousainYukoukikanFrom IS NULL OR ChousainYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                            "AND (ChousainYukoukikanTo IS NULL OR ChousainYukoukikanTo >= '" + Nendo + "/4/1' ) ";
                        }
                        combodt = new DataTable();
                        combodt = GlobalMethod.getData(discript, value, table, where);
                        if (combodt != null && combodt.Rows.Count > 0)
                        {
                            // 取得した部所をセット
                            //c1FlexGrid4.Rows[e.Row][e.Col - 1] = combodt.Rows[0][0].ToString();
                            //c1FlexGrid4.Rows[e.Row][e.Col] = combodt.Rows[0][1].ToString();
                            c1FlexGrid4.Rows[e.Row][ColName2] = combodt.Rows[0][0].ToString();
                            c1FlexGrid4.Rows[e.Row][ColName] = combodt.Rows[0][1].ToString();
                        }
                        else
                        {
                            // 取得できなかったらクリア
                            //c1FlexGrid4.Rows[e.Row][e.Col] = "";
                            c1FlexGrid4.Rows[e.Row][ColName] = "";
                        }
                    }
                }

                // 材工
                if (ColName == "ChousaZaiKou")
                {
                    // 別の値がペーストされる対策
                    if (c1FlexGrid4.Rows[e.Row][ColName] != null)
                    {
                        if (c1FlexGrid4.Rows[e.Row][ColName].ToString() == "1" || c1FlexGrid4.Rows[e.Row][ColName].ToString() == "材")
                        {
                            c1FlexGrid4.Rows[e.Row][ColName] = "1";
                        }
                        else if (c1FlexGrid4.Rows[e.Row][ColName].ToString() == "2" || c1FlexGrid4.Rows[e.Row][ColName].ToString() == "D工")
                        {
                            c1FlexGrid4.Rows[e.Row][ColName] = "2";
                        }
                        else if (c1FlexGrid4.Rows[e.Row][ColName].ToString() == "3" || c1FlexGrid4.Rows[e.Row][ColName].ToString() == "E工")
                        {
                            c1FlexGrid4.Rows[e.Row][ColName] = "3";
                        }
                        else if (c1FlexGrid4.Rows[e.Row][ColName].ToString() == "4" || c1FlexGrid4.Rows[e.Row][ColName].ToString() == "他")
                        {
                            c1FlexGrid4.Rows[e.Row][ColName] = "4";
                        }
                        else
                        {
                            c1FlexGrid4.Rows[e.Row][ColName] = "";
                        }
                    }
                    else
                    {
                        c1FlexGrid4.Rows[e.Row][ColName] = "";
                    }
                }

                // 属性
                if (ColName == "ChousaObiMei")
                {
                    // 別の値がペーストされる対策
                    if (c1FlexGrid4.Rows[e.Row][ColName] != null)
                    {
                        discript = "ZokuseiMeishou ";
                        value = "ZokuseiMeishou ";
                        table = "Mst_Zokusei ";
                        where = "ZokuseiID IS NOT NULL AND ZokuseiMeishou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[e.Row][ColName].ToString(), 0, 0) + "' ";

                        combodt = new DataTable();
                        combodt = GlobalMethod.getData(discript, value, table, where);
                        if (combodt != null && combodt.Rows.Count > 0)
                        {
                            // 取得出来た場合はスルー
                        }
                        else
                        {
                            // 取得できなかったらクリア
                            c1FlexGrid4.Rows[e.Row][ColName] = "";
                        }
                    }
                    else
                    {
                        c1FlexGrid4.Rows[e.Row][ColName] = "";
                    }
                }

                // 報告ランク,依頼ランク
                if (ColName == "ChousaHoukokuRank" || ColName == "ChousaIraiRank")
                {
                    // 別の値がペーストされる対策
                    if (c1FlexGrid4.Rows[e.Row][ColName] != null)
                    {
                        discript = "TankaRankHinmoku ";
                        value = "TankaRankHinmoku ";
                        table = "TankaKeiyakuRank ";
                        where = "TankaRankDeleteFlag != 1 AND TankaKeiyakuID = (SELECT TanpinGyoumuCD FROM TanpinNyuuryoku WHERE MadoguchiID = " + MadoguchiID + ") " +
                               "AND TankaRankHinmoku COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[e.Row][ColName].ToString(), 0, 0) + "' ";
                        //No.1443
                        string sRank = c1FlexGrid4.Rows[e.Row][ColName].ToString();
                        if (c1FlexGrid4.Rows[e.Row][ColName].ToString() == "半角ハイフン ")
                        {
                            sRank = "-";
                            where = "TankaRankDeleteFlag != 1 AND TankaKeiyakuID = (SELECT TanpinGyoumuCD FROM TanpinNyuuryoku WHERE MadoguchiID = " + MadoguchiID + ") " +
                               "AND TankaRankHinmoku COLLATE Japanese_XJIS_100_CI_AS_SC = N'-'";
                        }
                        combodt = new DataTable();
                        combodt = GlobalMethod.getData(discript, value, table, where);
                        if (combodt != null && combodt.Rows.Count > 0)
                        {
                            // 取得出来た場合はスルー
                            //No.1443
                            if (sRank == "-")
                            {
                                c1FlexGrid4.Rows[e.Row][ColName] = sRank;
                            }
                        }
                        else
                        {
                            // 取得できなかったらクリア
                            c1FlexGrid4.Rows[e.Row][ColName] = "";
                        }
                    }
                    else
                    {
                        c1FlexGrid4.Rows[e.Row][ColName] = "";
                    }
                }

                // 報告数、依頼数
                if (ColName == "ChousaHoukokuHonsuu" || ColName == "ChousaIraiHonsuu")
                {
                    if (c1FlexGrid4.Rows[e.Row][ColName] != null)
                    {
                        if (!Regex.IsMatch(c1FlexGrid4.Rows[e.Row][ColName].ToString(), @"^-?[\d][\d.]*$", RegexOptions.ECMAScript))
                        {
                            c1FlexGrid4.Rows[e.Row][ColName] = "0";
                        }
                        long maxNum = 99999999;
                        long inputNum = 0;

                        if (long.TryParse(c1FlexGrid4.Rows[e.Row][ColName].ToString(), out inputNum))
                        {
                            if (maxNum < inputNum || 0 > inputNum)
                            {
                                c1FlexGrid4.Rows[e.Row][ColName] = "0";
                            }
                        }
                        else
                        {
                            c1FlexGrid4.Rows[e.Row][ColName] = "0";
                        }
                    }
                }

                int strIndex = 0;
                Boolean strCheck = false;

                // 文字数制御
                // 品名
                if (e.Col == c1FlexGrid4.Cols["ChousaHinmei"].Index)
                {
                    strIndex = 200;
                }
                // 規格
                if (e.Col == c1FlexGrid4.Cols["ChousaKikaku"].Index)
                {
                    strIndex = 600;
                }
                // 単位
                if (e.Col == c1FlexGrid4.Cols["ChousaTanka"].Index)
                {
                    strIndex = 50;
                }
                // 参考質量
                if (e.Col == c1FlexGrid4.Cols["ChousaSankouShitsuryou"].Index)
                {
                    strIndex = 40;
                }
                // 価格
                if (e.Col == c1FlexGrid4.Cols["ChousaKakaku"].Index)
                {
                    strIndex = 50;
                }
                // 報告備考
                if (e.Col == c1FlexGrid4.Cols["ChousaBikou2"].Index)
                {
                    strIndex = 2000;
                }
                // 依頼備考
                if (e.Col == c1FlexGrid4.Cols["ChousaBikou"].Index)
                {
                    strIndex = 2000;
                }
                // 単価適用地域
                if (e.Col == c1FlexGrid4.Cols["ChousaTankaTekiyouTiku"].Index)
                {
                    strIndex = 100;
                }
                // 図面番号
                if (e.Col == c1FlexGrid4.Cols["ChousaZumenNo"].Index)
                {
                    strIndex = 200;
                }
                // 数量
                if (e.Col == c1FlexGrid4.Cols["ChousaSuuryou"].Index)
                {
                    strIndex = 50;
                }
                // 見積先
                if (e.Col == c1FlexGrid4.Cols["ChousaMitsumorisaki"].Index)
                {
                    strIndex = 32767;
                }
                // ベースメーカー
                if (e.Col == c1FlexGrid4.Cols["ChousaBaseMakere"].Index)
                {
                    strIndex = 50;
                }
                // 前回単位
                if (e.Col == c1FlexGrid4.Cols["ChousaZenkaiTani"].Index)
                {
                    strIndex = 50;
                }
                // 品目情報1
                if (e.Col == c1FlexGrid4.Cols["ChousaHinmokuJouhou1"].Index)
                {
                    strIndex = 100;
                }
                // 品目情報2
                if (e.Col == c1FlexGrid4.Cols["ChousaHinmokuJouhou2"].Index)
                {
                    strIndex = 100;
                }
                // 前回質量
                if (e.Col == c1FlexGrid4.Cols["ChousaFukuShizai"].Index)
                {
                    strIndex = 100;
                }
                // メモ1
                if (e.Col == c1FlexGrid4.Cols["ChousaBunrui"].Index)
                {
                    strIndex = 2000;
                }
                // メモ2
                if (e.Col == c1FlexGrid4.Cols["ChousaMemo2"].Index)
                {
                    strIndex = 2000;
                }
                // 発注品目コード
                if (e.Col == c1FlexGrid4.Cols["ChousaTankaCD1"].Index)
                {
                    strIndex = 50;
                }
                // 地区割コード
                if (e.Col == c1FlexGrid4.Cols["ChousaTikuWariCode"].Index)
                {
                    strIndex = 10;
                }
                // 地区コード
                if (e.Col == c1FlexGrid4.Cols["ChousaTikuCode"].Index)
                {
                    strIndex = 10;
                }
                // 地区名
                if (e.Col == c1FlexGrid4.Cols["ChousaTikuMei"].Index)
                {
                    strIndex = 50;
                }
                // 根拠関連コード
                if (e.Col == c1FlexGrid4.Cols["ChousaKonkyoCode"].Index)
                {
                    strIndex = 200;
                }

                if (strIndex != 0)
                {
                    if (c1FlexGrid4.Rows[e.Row][e.Col] != null && c1FlexGrid4.Rows[e.Row][e.Col].ToString().Length > strIndex)
                    {
                        strCheck = true;
                    }
                }

                if (strCheck == true)
                {
                    c1FlexGrid4.Rows[e.Row][e.Col] = c1FlexGrid4.Rows[e.Row][e.Col].ToString().Substring(0, strIndex);
                }


                // 全体順、個別順
                if (e.Col == c1FlexGrid4.Cols["ChousaZentaiJun"].Index || e.Col == c1FlexGrid4.Cols["ChousaKobetsuJun"].Index)
                {
                    if (c1FlexGrid4.Rows[e.Row][e.Col] != null)
                    {
                        // smallmoney の最大値
                        Double maxNum = 214748.3647;
                        Double inputNum = 0;

                        if (Double.TryParse(c1FlexGrid4.Rows[e.Row][e.Col].ToString(), out inputNum))
                        {
                            if (maxNum < inputNum || 0 > inputNum)
                            {
                                c1FlexGrid4.Rows[e.Row][e.Col] = "0";
                            }
                        }
                        else
                        {
                            c1FlexGrid4.Rows[e.Row][e.Col] = "0";
                        }
                    }
                }


                // ベース単価、前回価格、発注者提供単価、価格
                if (e.Col == c1FlexGrid4.Cols["ChousaBaseTanka"].Index || e.Col == c1FlexGrid4.Cols["ChousaZenkaiKakaku"].Index || e.Col == c1FlexGrid4.Cols["ChousaSankouti"].Index || e.Col == c1FlexGrid4.Cols["ChousaKakaku"].Index)
                {
                    if (c1FlexGrid4.Rows[e.Row][e.Col] != null)
                    {
                        Double maxNum = 999999999999.99;
                        Double inputNum = 0;
                        Double minNum = -999999999999.99;

                        if (Double.TryParse(c1FlexGrid4.Rows[e.Row][e.Col].ToString(), out inputNum))
                        {
                            if (maxNum < inputNum || minNum > inputNum)
                            {
                                c1FlexGrid4.Rows[e.Row][e.Col] = "0";
                            }
                        }
                        else
                        {
                            c1FlexGrid4.Rows[e.Row][e.Col] = "0";
                        }
                    }
                }

                // 掛け率
                if (e.Col == c1FlexGrid4.Cols["ChousaKakeritsu"].Index)
                {
                    if (c1FlexGrid4.Rows[e.Row][e.Col] != null)
                    {
                        Double maxNum = 999.99;
                        Double inputNum = 0;

                        if (Double.TryParse(c1FlexGrid4.Rows[e.Row][e.Col].ToString(), out inputNum))
                        {
                            if (maxNum < inputNum || 0 > inputNum)
                            {
                                c1FlexGrid4.Rows[e.Row][e.Col] = "0";
                            }
                        }
                        else
                        {
                            c1FlexGrid4.Rows[e.Row][e.Col] = "0";
                        }
                    }
                }
                ////奉行エクセル
                //// グループ名
                if (e.Col == c1FlexGrid4.Cols["GroupMei"].Index)
                {
                    strIndex = 15;
                }
            }
        }

        // 調査品目明細タブ
        private void c1FlexGrid4_KeyDownEdit(object sender, C1.Win.C1FlexGrid.KeyEditEventArgs e)
        {
            // 調査品目編集モード 0:表示 1:編集
            if (ChousaHinmokuMode == 1)
            {
                // Enterだけで改行できるように
                // 09:品名、10:規格、11:単位、12:参考質量、15:報告備考、16:依頼備考、17:単価適用地域、18:図面番号、
                // 20:見積先、21:ベースメーカ、25:前回単位、28:品目情報1、29:品目情報2、30:前回質量、31:メモ1、32:メモ2
                //if ((e.Col == 9 || e.Col == 10 || e.Col == 11 || e.Col == 12 || e.Col == 15 || e.Col == 16 || e.Col == 17 || e.Col == 18 ||
                //    e.Col == 20 || e.Col == 21 || e.Col == 25 || e.Col == 28 || e.Col == 29 || e.Col == 30 || e.Col == 31 || e.Col == 32) && e.Row >= 2)
                if ((e.Col == c1FlexGrid4.Cols["ChousaHinmei"].Index || e.Col == c1FlexGrid4.Cols["ChousaKikaku"].Index
                    || e.Col == c1FlexGrid4.Cols["ChousaTanka"].Index || e.Col == c1FlexGrid4.Cols["ChousaSankouShitsuryou"].Index
                    || e.Col == c1FlexGrid4.Cols["ChousaBikou2"].Index || e.Col == c1FlexGrid4.Cols["ChousaBikou"].Index
                    || e.Col == c1FlexGrid4.Cols["ChousaTankaTekiyouTiku"].Index || e.Col == c1FlexGrid4.Cols["ChousaZumenNo"].Index
                    || e.Col == c1FlexGrid4.Cols["ChousaMitsumorisaki"].Index || e.Col == c1FlexGrid4.Cols["ChousaBaseMakere"].Index
                    || e.Col == c1FlexGrid4.Cols["ChousaZenkaiTani"].Index || e.Col == c1FlexGrid4.Cols["ChousaHinmokuJouhou1"].Index
                    || e.Col == c1FlexGrid4.Cols["ChousaHinmokuJouhou2"].Index || e.Col == c1FlexGrid4.Cols["ChousaFukuShizai"].Index
                    || e.Col == c1FlexGrid4.Cols["ChousaBunrui"].Index || e.Col == c1FlexGrid4.Cols["ChousaMemo2"].Index
                    ) && e.Row >= 2)
                {
                    // Enterキー押下で改行させる
                    if ((e.KeyCode == Keys.Enter))
                    {
                        if ((e.Modifiers != Keys.Alt))
                        {
                            SendKeys.Send("%{ENTER}");
                            e.Handled = true;
                        }
                    }
                }
            }
        }

        // 履歴登録
        private void writeHistory(string historyMessage)
        {
            string methodName = ".writeHistory";
            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                //履歴登録
                cmd.CommandText = "INSERT INTO T_HISTORY(" +
                    "H_DATE_KEY " +
                    ",H_NO_KEY " +
                    ",H_OPERATE_DT " +
                    ",H_OPERATE_USER_ID " +
                    ",H_OPERATE_USER_MEI " +
                    ",H_OPERATE_USER_BUSHO_CD " +
                    ",H_OPERATE_USER_BUSHO_MEI " +
                    ",H_OPERATE_NAIYO " +
                    ",H_ProgramName " +
                    ",MadoguchiID " +
                    ",HistoryBeforeTantoubushoCD " +
                    ",HistoryBeforeTantoushaCD " +
                    ",HistoryAfterTantoubushoCD " +
                    ",HistoryAfterTantoushaCD " +
                    ",H_TOKUCHOBANGOU " +
                    ")VALUES(" +
                    "SYSDATETIME() " + 
                    ", " + GlobalMethod.getSaiban("HistoryID") + " " +
                    ",SYSDATETIME() " +
                    ",'" + UserInfos[0] + "' " +
                    ",N'" + UserInfos[1] + "' " +
                    ",'" + UserInfos[2] + "' " +
                    ",N'" + UserInfos[3] + "' " +
                    ",N'" + historyMessage + "' " +
                    ",'" + pgmName + methodName + "' " +
                    "," + MadoguchiID + " " +
                    ",NULL " +
                    ",NULL " +
                    ",NULL " +
                    ",NULL " +
                    ",N'" + Header1.Text + "' " +
                    ")";

                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        // 11桁になるようゼロパディングした文字を返却する
        private string zeroPadding(string str)
        {
            string moji = "";

            double.TryParse(str, out double num);
            moji = string.Format("{0:000000.0000}", num);
            return moji;
        }

        private void button3_RowAdd_Click(object sender, EventArgs e)
        {

            // 行の追加プロンプト
            Popup_GyouTsuika form = new Popup_GyouTsuika();
            form.UserInfos = UserInfos;
            form.Nendo = DateTime.Today.Year;
            if (item1_TourokuNendo.Text != "")
            {
                form.Nendo = int.Parse(item1_TourokuNendo.Text);
            }
            form.ShowDialog();

            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                addline(form.ReturnValue);
            }
            // 表示件数（ヘッダー2行分を引く）
            Grid_Num.Text = "(" + (c1FlexGrid4.Rows.Count - 2) + ")";
        }

        private void addline(string[] Value)
        {
            string w_ChousaBushoCD = Value[0];                                      // 調査担当部所CD
            string w_ChousaBusho = Value[1];                                        // 調査担当部所
            string w_ChousaTantoushaCD = Value[2];                                  // 調査担当者CD
            string w_ChousaTantousha = Value[3];                                    // 調査担当者
            string w_TankaTekiyouChiiki = Value[4];                                 // 単価適用地域
            int.TryParse(Value[5].ToString(), out int w_TuikaGyousuu);              // 追加行数
            double.TryParse(Value[6].ToString(), out double w_ZentaiJunKaishiNo);   // 全体順開始番号

            //レイアウトロジックを停止する
            this.SuspendLayout();

            c1FlexGrid4.Visible = false;

            double num = 0;
            double w_ZentaiJun = w_ZentaiJunKaishiNo;
            // 全体順が指定されていなかった場合、最大値を取得する
            if (w_ZentaiJun == 0)
            {
                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                {
                    // 全体順の最大値を取り出す
                    //if (c1FlexGrid4.Rows[i][6] != null && double.TryParse(c1FlexGrid4.Rows[i][6].ToString(), out num))
                    if (c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null && double.TryParse(c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString(), out num))
                    {
                        if (num > w_ZentaiJun)
                        {
                            w_ZentaiJun = num;
                        }
                    }
                }
                // 取得した値の次から追加するので+1しておく
                w_ZentaiJun += 1;
            }

            // 全体順の中での最大の個別順を取得する
            double w_KobetsuJun = 0;
            num = 0;
            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            {
                //if (w_ZentaiJun == double.Parse(c1FlexGrid4.Rows[i][6].ToString()))
                if (w_ZentaiJun == double.Parse(c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString()))
                {
                    // 個別順の最大値を取り出す
                    //if (c1FlexGrid4.Rows[i][7] != null && double.TryParse(c1FlexGrid4.Rows[i][7].ToString(), out num))
                    if (c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null && double.TryParse(c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString(), out num))
                    {
                        if (num > w_KobetsuJun)
                        {
                            w_KobetsuJun = num;
                        }
                    }
                }
            }
            // 取得した値の次から追加するので+1しておく(小数桁切り捨て)
            w_KobetsuJun = Math.Truncate(w_KobetsuJun) + 1;

            int rowCount = c1FlexGrid4.Rows.Count;
            // 追加行数分処理をおこなう
            for (int i = 0; i < w_TuikaGyousuu; i++)
            {
                // 行挿入
                c1FlexGrid4.Rows.Insert(rowCount);

                // グリッドに値をセットする
                for (int j = 0; j < c1FlexGrid4.Cols.Count; j++)
                {

                    // ChousaHinmokuID
                    //if (j == 55)
                    if (j == c1FlexGrid4.Cols["ChousaHinmokuID2"].Index)
                    {
                        // 追加時にIDを振っておく
                        c1FlexGrid4.Rows[rowCount][j] = GlobalMethod.getSaiban("HinmokuMeisaiID");
                    }
                    // 全体順
                    //else if (j == 6)
                    else if (j == c1FlexGrid4.Cols["ChousaZentaiJun"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = w_ZentaiJun;
                    }
                    // 個別順
                    //else if (j == 7)
                    else if (j == c1FlexGrid4.Cols["ChousaKobetsuJun"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = w_KobetsuJun + i;
                    }
                    // 材工
                    else if (j == c1FlexGrid4.Cols["ChousaZaiKou"].Index)
                    {
                        // 1:材
                        c1FlexGrid4.Rows[rowCount][j] = 1;
                    }
                    // 中止
                    else if (j == c1FlexGrid4.Cols["ChousaChuushi"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = false;
                    }
                    // 単価適用地域
                    //else if (j == 17)
                    else if (j == c1FlexGrid4.Cols["ChousaTankaTekiyouTiku"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = w_TankaTekiyouChiiki;
                    }
                    // ベース単価
                    //else if (j == 22)
                    else if (j == c1FlexGrid4.Cols["ChousaBaseTanka"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = 0;
                    }
                    // 掛率
                    //else if (j == 23)
                    else if (j == c1FlexGrid4.Cols["ChousaKakeritsu"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = 0;
                    }
                    // 前回価格
                    //else if (j == 26)
                    else if (j == c1FlexGrid4.Cols["ChousaZenkaiKakaku"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = 0;
                    }
                    // 発注者提供単価
                    //else if (j == 27)
                    else if (j == c1FlexGrid4.Cols["ChousaSankouti"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = 0;
                    }
                    // フォルダアイコン
                    //else if (j == 40)
                    else if (j == c1FlexGrid4.Cols["ChousaLinkSakli"].Index)
                    {
                        // 集計表フォルダアイコン 0:グレー 1:イエロー
                        if (folderIcon == "1")
                        {
                            c1FlexGrid4.Rows[rowCount][j] = "1";
                        }
                        else
                        {
                            c1FlexGrid4.Rows[rowCount][j] = "0";
                        }
                    }
                    // 調査担当部所
                    //else if (j == 42)
                    else if (j == c1FlexGrid4.Cols["HinmokuRyakuBushoCD"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = w_ChousaBushoCD;
                    }
                    // 調査担当者
                    //else if (j == 43)
                    else if (j == c1FlexGrid4.Cols["HinmokuChousainCD"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = w_ChousaTantoushaCD;
                    }
                    // 報告数
                    //else if (j == 48)
                    else if (j == c1FlexGrid4.Cols["ChousaHoukokuHonsuu"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = 0;
                    }
                    // 依頼数
                    //else if (j == 50)
                    else if (j == c1FlexGrid4.Cols["ChousaIraiHonsuu"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = 0;
                    }
                    // 締切日
                    //else if (j == 52)
                    else if (j == c1FlexGrid4.Cols["ChousaHinmokuShimekiribi"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = item1_MadoguchiShimekiribi.Text;
                    }
                    // ChousaShinchokuJoukyou
                    //else if (j == 56)
                    else if (j == c1FlexGrid4.Cols["ChousaShinchokuJoukyou"].Index)
                    {
                        // 進捗状況は、20:調査開始
                        c1FlexGrid4.Rows[rowCount][j] = 20;
                    }
                    // 0:Insert/1:Select/2:Update
                    //else if (j == 57)
                    else if (j == c1FlexGrid4.Cols["Mode"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = "0";
                    }
                    // ソートキー
                    //else if (j == 58)
                    else if (j == c1FlexGrid4.Cols["ColumnSort"].Index)
                    {
                        // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                        //c1FlexGrid4.Rows[rowCount][j] = "N" + string.Format("{0:000000.0000}", w_ZentaiJun) + "-" + string.Format("{0:000000.0000}", w_KobetsuJun + i);
                        //c1FlexGrid4.Rows[rowCount][j] = "N" + zeroPadding(c1FlexGrid4.Rows[rowCount][6].ToString()) + "-" + zeroPadding(c1FlexGrid4.Rows[rowCount][7].ToString());
                        c1FlexGrid4.Rows[rowCount][j] = "N"
                                                      + zeroPadding(c1FlexGrid4.Rows[rowCount]["ChousaZentaiJun"].ToString())
                                                      + "-"
                                                      + zeroPadding(c1FlexGrid4.Rows[rowCount]["ChousaKobetsuJun"].ToString())
                                                      ;
                    }
                    else
                    {
                        c1FlexGrid4.Rows[rowCount][j] = "";
                    }
                }
                // 進捗状況を判定する（進捗状況は20:調査開始固定なので、進捗状況の以外を判定する）
                DateTime dateTime = DateTime.Today;
                // 1:報告済みの場合
                if ("1".Equals(MadoguchiHoukokuzumi))
                {
                    // 報告済み
                    //c1FlexGrid4.Rows[rowCount][5] = "8";
                    c1FlexGrid4.Rows[rowCount]["ShinchokuIcon"] = "8";
                }
                else
                {
                    try
                    {
                        // 締切日
                        //if (c1FlexGrid4[rowCount, 52] != null)
                        if (c1FlexGrid4.Rows[rowCount]["ChousaHinmokuShimekiribi"] != null)
                        {
                            //dateTime = DateTime.Parse(c1FlexGrid4[rowCount, 52].ToString());
                            dateTime = DateTime.Parse(c1FlexGrid4.Rows[rowCount]["ChousaHinmokuShimekiribi"].ToString());
                            if (dateTime < DateTime.Today)
                            {
                                // 締切日経過
                                //c1FlexGrid4.Rows[rowCount][5] = "1";
                                c1FlexGrid4.Rows[rowCount]["ShinchokuIcon"] = "1";
                            }
                            else if (dateTime < DateTime.Today.AddDays(3))
                            {
                                // 締切日が3日以内、かつ2次検証が完了していない
                                //c1FlexGrid4.Rows[rowCount][5] = "2";
                                c1FlexGrid4.Rows[rowCount]["ShinchokuIcon"] = "2";
                            }
                            else if (dateTime < DateTime.Today.AddDays(7))
                            {
                                // 締切日が1週間以内、かつ2次検証が完了していない
                                //c1FlexGrid4.Rows[rowCount][5] = "3";
                                c1FlexGrid4.Rows[rowCount]["ShinchokuIcon"] = "3";
                            }
                            else
                            {
                                //c1FlexGrid4.Rows[rowCount][5] = "4";
                                c1FlexGrid4.Rows[rowCount]["ShinchokuIcon"] = "4";
                            }
                        }
                    }
                    catch
                    {
                        // 日付変換エラー
                        throw;
                    }
                }
                rowCount += 1;

            }
            // 並び順列のIndex（行番号）を取得する。
            int ColumnSortColIndex = c1FlexGrid4.Cols["ColumnSort"].Index;
            int ZentaiJunColIndex = c1FlexGrid4.Cols["ChousaZentaiJun"].Index;

            // ソート
            //c1FlexGrid4.Cols[58].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
            //c1FlexGrid4.Cols[6].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
            //c1FlexGrid4.Cols.Move(58, 6);
            //c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 58, 6);
            //c1FlexGrid4.Cols.Move(6, 58);
            c1FlexGrid4.Cols[ColumnSortColIndex].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
            c1FlexGrid4.Cols[ZentaiJunColIndex].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
            c1FlexGrid4.Cols.Move(ColumnSortColIndex, 1);
            c1FlexGrid4.Cols.Move(ZentaiJunColIndex, 2);
            c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 1, 2);
            c1FlexGrid4.Cols.Move(2, ZentaiJunColIndex);
            c1FlexGrid4.Cols.Move(1, ColumnSortColIndex);

            c1FlexGrid4.Visible = true;

            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void c1FlexGrid4_ValidateEdit(object sender, C1.Win.C1FlexGrid.ValidateEditEventArgs e)
        {
            // 報告数、依頼数
            //if (e.Col == 48 || e.Col == 50)
            if (e.Col == c1FlexGrid4.Cols["ChousaHoukokuHonsuu"].Index || e.Col == c1FlexGrid4.Cols["ChousaIraiHonsuu"].Index)
            {
                if (c1FlexGrid4.Editor != null)
                {
                    if (!Regex.IsMatch(c1FlexGrid4.Editor.Text, @"^-?[\d][\d.]*$", RegexOptions.ECMAScript))
                    {
                        c1FlexGrid4.Editor.Text = "0";
                    }
                    long maxNum = 99999999;
                    long inputNum = 0;

                    if (long.TryParse(c1FlexGrid4.Editor.Text, out inputNum))
                    {
                        if (maxNum < inputNum || 0 > inputNum)
                        {
                            c1FlexGrid4.Editor.Text = "0";
                        }
                    }
                    else
                    {
                        c1FlexGrid4.Editor.Text = "0";
                    }

                }
            }

            // 全体順、個別順
            if (e.Col == c1FlexGrid4.Cols["ChousaZentaiJun"].Index || e.Col == c1FlexGrid4.Cols["ChousaKobetsuJun"].Index)
            {
                if (c1FlexGrid4.Editor != null)
                {
                    // smallmoney の最大値
                    Double maxNum = 214748.3647;
                    Double inputNum = 0;

                    if (Double.TryParse(c1FlexGrid4.Editor.Text, out inputNum))
                    {
                        if (maxNum < inputNum || 0 > inputNum)
                        {
                            c1FlexGrid4.Editor.Text = "0";
                        }
                    }
                    else
                    {
                        c1FlexGrid4.Editor.Text = "0";
                    }
                }
            }


            // ベース単価、前回価格、発注者提供単価、価格
            if (e.Col == c1FlexGrid4.Cols["ChousaBaseTanka"].Index || e.Col == c1FlexGrid4.Cols["ChousaZenkaiKakaku"].Index || e.Col == c1FlexGrid4.Cols["ChousaSankouti"].Index || e.Col == c1FlexGrid4.Cols["ChousaKakaku"].Index)
            {
                if (c1FlexGrid4.Editor != null)
                {
                    Double maxNum = 999999999999.99;
                    Double inputNum = 0;
                    Double minNum = -999999999999.99;

                    if (Double.TryParse(c1FlexGrid4.Editor.Text, out inputNum))
                    {
                        if (maxNum < inputNum || minNum > inputNum)
                        {
                            c1FlexGrid4.Editor.Text = "0";
                        }
                    }
                    else
                    {
                        c1FlexGrid4.Editor.Text = "0";
                    }
                }
            }

            // 掛け率
            if (e.Col == c1FlexGrid4.Cols["ChousaKakeritsu"].Index)
            {
                if (c1FlexGrid4.Editor != null)
                {
                    Double maxNum = 999.99;
                    Double inputNum = 0;

                    if (Double.TryParse(c1FlexGrid4.Editor.Text, out inputNum))
                    {
                        if (maxNum < inputNum || 0 > inputNum)
                        {
                            c1FlexGrid4.Editor.Text = "0";
                        }
                    }
                    else
                    {
                        c1FlexGrid4.Editor.Text = "0";
                    }
                }
            }

            // 報告数、依頼数
            if (e.Col == c1FlexGrid4.Cols["ChousaHoukokuHonsuu"].Index || e.Col == c1FlexGrid4.Cols["ChousaIraiHonsuu"].Index)
            {
                if (c1FlexGrid4.Editor != null)
                {
                    Double maxNum = 99999999;
                    Double inputNum = 0;

                    if (Double.TryParse(c1FlexGrid4.Editor.Text, out inputNum))
                    {
                        if (maxNum < inputNum || 0 > inputNum)
                        {
                            c1FlexGrid4.Editor.Text = "0";
                        }
                    }
                    else
                    {
                        c1FlexGrid4.Editor.Text = "0";
                    }
                }
            }
        }

        private void c1FlexGrid4_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            DateTime dateTime = DateTime.Today;

            // 52:締切日
            //if (e.Row > 0 && e.Col == 52)
            if (e.Row > 0 && e.Col <= c1FlexGrid4.Cols["ChousaHinmokuShimekiribi"].Index)
            {
                // 1:報告済みの場合
                if ("1".Equals(MadoguchiHoukokuzumi))
                {
                    // 報告済み
                    //c1FlexGrid4[e.Row, 5] = "8";
                    c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "8";
                }
                else
                {
                    // ▼担当者状況
                    // 10:依頼
                    // 20:調査開始
                    // 30:見積中
                    // 40:集計中
                    // 50:担当者済
                    // 60:一次検済
                    // 70:二次検済
                    // 80:中止
                    //if ("80".Equals(c1FlexGrid4[e.Row, 55].ToString()))
                    if (c1FlexGrid4.Rows[e.Row]["ChousaChuushi"] != null && "True".Equals(c1FlexGrid4.Rows[e.Row]["ChousaChuushi"].ToString()))
                    {
                        c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "6";
                    }
                    else if ("80".Equals(c1FlexGrid4.Rows[e.Row]["ChousaShinchokuJoukyou"].ToString()))
                    {
                        // 二次検証済み、または中止（中止）
                        //c1FlexGrid4.Rows[e.Row][5] = "6";
                        c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "6";
                    }
                    //else if ("70".Equals(c1FlexGrid4[e.Row, 55].ToString()))
                    else if ("70".Equals(c1FlexGrid4.Rows[e.Row]["ChousaShinchokuJoukyou"].ToString()))
                    {
                        // 二次検証済み、または中止（二次検証済み）
                        //c1FlexGrid4.Rows[e.Row][5] = "5";
                        c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "5";
                    }
                    //else if ("50".Equals(c1FlexGrid4[e.Row, 55].ToString()) || "60".Equals(c1FlexGrid4[e.Row, 55].ToString()))
                    else if ("50".Equals(c1FlexGrid4.Rows[e.Row]["ChousaShinchokuJoukyou"].ToString()) || "60".Equals(c1FlexGrid4.Rows[e.Row]["ChousaShinchokuJoukyou"].ToString()))
                    {
                        // 担当者済み or 一次検済
                        //c1FlexGrid4.Rows[e.Row][5] = "7";
                        c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "7";
                    }
                    //else if (c1FlexGrid4[e.Row, 52] != null)
                    else if (c1FlexGrid4.Rows[e.Row]["ChousaHinmokuShimekiribi"] != null)
                    {
                        try
                        {
                            //dateTime = DateTime.Parse(c1FlexGrid4[e.Row, 52].ToString());
                            dateTime = DateTime.Parse(c1FlexGrid4.Rows[e.Row]["ChousaHinmokuShimekiribi"].ToString());
                            if (dateTime < DateTime.Today)
                            {
                                // 締切日経過
                                //c1FlexGrid4.Rows[e.Row][5] = "1";
                                c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "1";
                            }
                            else if (dateTime < DateTime.Today.AddDays(3))
                            {
                                // 締切日が3日以内、かつ2次検証が完了していない
                                //c1FlexGrid4.Rows[e.Row][5] = "2";
                                c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "2";
                            }
                            else if (dateTime < DateTime.Today.AddDays(7))
                            {
                                // 締切日が1週間以内、かつ2次検証が完了していない
                                //c1FlexGrid4.Rows[e.Row][5] = "3";
                                c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "3";
                            }
                            else
                            {
                                //c1FlexGrid4.Rows[e.Row][5] = "4";
                                c1FlexGrid4.Rows[e.Row]["ShinchokuIcon"] = "4";
                            }
                        }
                        catch
                        {
                            // 日付変換エラー
                            throw;
                        }
                    }
                }
            }

            // 報告数、依頼数
            if (e.Col == 48 || e.Col == 50)
            {
                if (c1FlexGrid4.Rows[e.Row][e.Col] != null)
                {
                    if (!Regex.IsMatch(c1FlexGrid4.Rows[e.Row][e.Col].ToString(), @"^-?[\d][\d.]*$", RegexOptions.ECMAScript))
                    {
                        c1FlexGrid4.Rows[e.Row][e.Col] = "";
                    }
                }
            }
        }

        // 調査品目変更履歴登録
        private void writeChousaHinmokuHistory(string historyMessage, string beforeRyakuBushoCD, string beforeChousainCD, string afterRyakuBushoCD, string afterChousainCD)
        {
            string methodName = ".writeHinmokuHistory";
            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                //履歴登録
                cmd.CommandText = "INSERT INTO T_HISTORY(" +
                    "H_DATE_KEY " +
                    ",H_NO_KEY " +
                    ",H_OPERATE_DT " +
                    ",H_OPERATE_USER_ID " +
                    ",H_OPERATE_USER_MEI " +
                    ",H_OPERATE_USER_BUSHO_CD " +
                    ",H_OPERATE_USER_BUSHO_MEI " +
                    ",H_OPERATE_NAIYO " +
                    ",H_ProgramName " +
                    ",H_TOKUCHOBANGOU " +
                    ",MadoguchiID " +
                    ",HistoryBeforeTantoubushoCD " +   // nvarchar(6)
                    ",HistoryBeforeTantoubushoMei " +  // nvarchar(50)
                    ",HistoryBeforeTantoushaCD " +     // nvarchar(6)
                    ",HistoryBeforeTantoushaMei " +    // nvarchar(50)
                    ",HistoryAfterTantoubushoCD " +    // nvarchar(6)
                    ",HistoryAfterTantoubushoMei " +   // nvarchar(50)
                    ",HistoryAfterTantoushaCD " +      // int
                    ",HistoryAfterTantoushaMei " +     // nvarchar(50)
                    ")VALUES(" +
                    "SYSDATETIME() " + 
                    ", " + GlobalMethod.getSaiban("HistoryID") + " " +
                    ",SYSDATETIME() " +
                    ",'" + UserInfos[0] + "' " +
                    ",N'" + UserInfos[1] + "' " +
                    ",'" + UserInfos[2] + "' " +
                    ",N'" + UserInfos[3] + "' " +
                    ",N'" + historyMessage + "' " +
                    ",'" + pgmName + methodName + "' " +
                    ",N'" + Header1.Text + "' " +
                    "," + MadoguchiID + " ";

                // 変更前部所
                if (beforeRyakuBushoCD != "")
                {
                    cmd.CommandText += ",'" + beforeRyakuBushoCD + "' " +
                            ",(SELECT TOP 1 ShibuMei FROM Mst_Busho WHERE GyoumuBushoCD = '" + beforeRyakuBushoCD + "' AND BushoDeleteFlag != 1) ";
                }
                else
                {
                    cmd.CommandText += ",null " +
                        ",null ";
                }
                // 変更前担当者
                if (beforeChousainCD != "")
                {
                    cmd.CommandText += ",'" + beforeChousainCD + "' " +
                            ",(SELECT TOP 1 ChousainMei FROM Mst_Chousain WHERE KojinCD = '" + beforeChousainCD + "' AND ChousainDeleteFlag != 1) ";
                }
                else
                {
                    cmd.CommandText += ",null " +
                        ",null ";
                }
                // 変更後部所
                if (afterRyakuBushoCD != "")
                {
                    cmd.CommandText += ",'" + afterRyakuBushoCD + "' " +
                            ",(SELECT TOP 1 ShibuMei FROM Mst_Busho WHERE GyoumuBushoCD = '" + afterRyakuBushoCD + "' AND BushoDeleteFlag != 1) ";
                }
                else
                {
                    cmd.CommandText += ",null " +
                        ",null ";
                }
                // 変更後担当者
                if (afterChousainCD != "")
                {
                    cmd.CommandText += ",'" + afterChousainCD + "' " +
                            ",(SELECT TOP 1 ChousainMei FROM Mst_Chousain WHERE KojinCD = '" + afterChousainCD + "' AND ChousainDeleteFlag != 1) ";
                }
                else
                {
                    cmd.CommandText += ",null " +
                        ",null ";
                }

                cmd.CommandText +=
                    ")";

                cmd.ExecuteNonQuery();
                conn.Close();
            }
        }

        //備考画面更新ボタン押下時
        private void BikoUpdateBtn_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {                // ShibuBikou 初期化
                ShibuBikoManager sbm = new ShibuBikoManager();
                set_error("", 0);
                set_error(sbm.UpdateShibuBiko(BikoGrid, Decimal.Parse(MadoguchiID)));
                sbm.ShibuBikoInit(this, Decimal.Parse(MadoguchiID));
                //更新履歴
                writeHistory("支部備考が変更されました。 ID= :" + MadoguchiID);
            }
        }

        // 備考Grid編集前判定
        private void BikoGrid_StartEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            string loginBCD = UserInfos[2].ToString();           // MadoguchiID
            String sbkanri = BikoGrid.GetData(BikoGrid.Row, 2).ToString(); // ShibuBikouBushoKanriboBushoCD

            if (loginBCD.Equals(sbkanri))
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }
        // 調査品目明細
        private void c1FlexGrid4_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            // ペースト時にコピーしたセルの改行が入らない対応
            // ListにC1のデータを保持しておき、
            // ペースト時にクリップボードではなく、Listに保持したデータをセットする

            // Ctrl + C
            if (e.KeyCode == Keys.C && e.Modifiers == Keys.Control)
            {
                if (copyData == null)
                {
                    copyData = new List<List<string>>();
                }
                copyData.Clear();
                for (int rowIndex = c1FlexGrid4.Selection.TopRow; rowIndex <= c1FlexGrid4.Selection.BottomRow; rowIndex++)
                {
                    copyData.Add(new List<string>());
                    for (int colIndex = c1FlexGrid4.Selection.LeftCol; colIndex <= c1FlexGrid4.Selection.RightCol; colIndex++)
                    {
                        // c1FlexGridのセルがNullの場場合、エラーとなるので、切り分ける
                        if (c1FlexGrid4[rowIndex, colIndex] != null)
                        {
                            DataTable combodt = new DataTable();
                            string discript = "BushokanriboKameiRaku";
                            string value = "BushokanriboKameiRaku";
                            string table = "Mst_Busho";
                            string where = "";

                            // 42:調査担当部所、44:副調査担当部所1、46:副調査担当部所2
                            //if (colIndex == 42 || colIndex == 44 || colIndex == 46)
                            if (colIndex == c1FlexGrid4.Cols["HinmokuRyakuBushoCD"].Index
                                || colIndex == c1FlexGrid4.Cols["HinmokuRyakuBushoFuku1CD"].Index
                                || colIndex == c1FlexGrid4.Cols["HinmokuRyakuBushoFuku2CD"].Index)
                            {
                                if (c1FlexGrid4.Rows[rowIndex][colIndex] != null && c1FlexGrid4.Rows[rowIndex][colIndex].ToString() != "")
                                {
                                    if (System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[rowIndex][colIndex].ToString().Replace(Environment.NewLine, ""), @"^[0-9]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                                    {
                                        where = "GyoumuBushoCD = '" + c1FlexGrid4.Rows[rowIndex][colIndex].ToString() + "' ";
                                    }
                                    else
                                    {
                                        where = "BushokanriboKameiRaku COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + c1FlexGrid4.Rows[rowIndex][colIndex].ToString() + "' ";
                                    }
                                    where += " AND BushoDeleteFlag = 0 AND BushoMadoguchiHyoujiFlg = 1 ";
                                    combodt = new DataTable();
                                    combodt = GlobalMethod.getData(discript, value, table, where);
                                    if (combodt != null && combodt.Rows.Count > 0)
                                    {
                                        // 取得した部所をセット
                                        copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add(combodt.Rows[0][0].ToString());
                                    }
                                    else
                                    {
                                        // 取得できなかったら
                                        copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("");
                                    }
                                }
                                else
                                {
                                    copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("");
                                }

                            }
                            //else if (colIndex == 43 || colIndex == 45 || colIndex == 47)
                            else if (colIndex == c1FlexGrid4.Cols["HinmokuChousainCD"].Index
                                || colIndex == c1FlexGrid4.Cols["HinmokuFukuChousainCD1"].Index
                                || colIndex == c1FlexGrid4.Cols["HinmokuFukuChousainCD2"].Index)
                            {
                                combodt = new DataTable();
                                discript = "ChousainMei";
                                value = "ChousainMei";
                                table = "Mst_Chousain";
                                where = "";

                                if (c1FlexGrid4.Rows[rowIndex][colIndex] != null && c1FlexGrid4.Rows[rowIndex][colIndex].ToString() != "")
                                {
                                    if (System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[rowIndex][colIndex].ToString().Replace(Environment.NewLine, ""), @"^[0-9]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                                    {
                                        where = "KojinCD = '" + c1FlexGrid4.Rows[rowIndex][colIndex].ToString() + "' ";
                                    }
                                    else
                                    {
                                        where = "ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + c1FlexGrid4.Rows[rowIndex][colIndex].ToString() + "' ";
                                    }

                                    where += " AND RetireFLG = 0 AND TokuchoFLG = 1";
                                    combodt = new DataTable();
                                    combodt = GlobalMethod.getData(discript, value, table, where);
                                    if (combodt != null && combodt.Rows.Count > 0)
                                    {
                                        copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add(combodt.Rows[0][0].ToString());
                                    }
                                    else
                                    {
                                        // 取得できなかったら
                                        copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("");
                                    }
                                }
                                else
                                {
                                    copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("");
                                }
                            }
                            // 材工
                            else if (colIndex == c1FlexGrid4.Cols["ChousaZaiKou"].Index)
                            {
                                if (c1FlexGrid4.Rows[rowIndex][colIndex].ToString() == "1")
                                {
                                    copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("材");
                                }
                                else if (c1FlexGrid4.Rows[rowIndex][colIndex].ToString() == "2")
                                {
                                    copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("D工");
                                }
                                else if (c1FlexGrid4.Rows[rowIndex][colIndex].ToString() == "3")
                                {
                                    copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("E工");
                                }
                                else if (c1FlexGrid4.Rows[rowIndex][colIndex].ToString() == "4")
                                {
                                    copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("他");
                                }
                                else
                                {
                                    copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("");
                                }
                            }
                            else
                            {
                                copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add(c1FlexGrid4[rowIndex, colIndex].ToString());
                            }
                        }
                        else
                        {
                            copyData[rowIndex - c1FlexGrid4.Selection.TopRow].Add("");
                        }
                    }
                }
            }
            // Ctrl + V
            else if (e.KeyCode == Keys.V && e.Modifiers == Keys.Control)
            {
                // No.1417 1182 奉行上でのコピペについて
                //if (copyData == null || ChousaHinmokuMode != 1)
                //{
                //    return;
                //}
                //for (int rowIndex = 0; rowIndex < copyData.Count && c1FlexGrid4.Selection.TopRow + rowIndex < c1FlexGrid4.Rows.Count; rowIndex++)
                //{
                //    //for (int colIndex = 0; colIndex < copyData[rowIndex].Count && c1FlexGrid4.Selection.RightCol + colIndex < c1FlexGrid4.Cols.Count; colIndex++)
                //    //{
                //    //    c1FlexGrid4[c1FlexGrid4.Selection.TopRow + rowIndex, c1FlexGrid4.Selection.RightCol + colIndex] = copyData[rowIndex][colIndex];
                //    //}
                //    for (int colIndex = 0; colIndex < copyData[rowIndex].Count && c1FlexGrid4.Selection.LeftCol + colIndex < c1FlexGrid4.Cols.Count; colIndex++)
                //    {
                //        c1FlexGrid4[c1FlexGrid4.Selection.TopRow + rowIndex, c1FlexGrid4.Selection.LeftCol + colIndex] = copyData[rowIndex][colIndex];
                //    }
                //}
            }
        }

        #region 廃棄
        // 調査品目明細GridのKeyDown
        private void c1FlexGrid4_KeyDown_1(object sender, KeyEventArgs e)
        {
            //調査品目編集モード 0:表示 1:編集
            if (ChousaHinmokuMode == 1)
            {
                // Ctrl + V
                if (e.KeyData == (Keys.Control | Keys.V))
                {
                    // 範囲選択時のペースト
                    // 750 行を複数行選択しペーストしたが1行分のみしかペースト出来ない対応

                    // 範囲選択が1列のみを選択している場合
                    //if (HinmokuCol == HinmokuColSel)
                    // 1列に限らずに貼り付ける
                    //if (HinmokuCol <= HinmokuColSel)
                    // 絶対にこの処理を通す
                    if (HinmokuCol <= HinmokuColSel || HinmokuCol >= HinmokuColSel)
                    {
                        //No.1417 1182 奉行上でもコピペについて
                        bool isWinCopy = false;
                        IDataObject data = Clipboard.GetDataObject();
                        string strWinCopyText = "";
                        if (data.GetDataPresent(DataFormats.Text))
                        {
                            string str;
                            //クリップボードからデータを取得
                            str = (string)data.GetData(DataFormats.Text);
                            //クリップボードにある最後の開業コードを削除
                            strWinCopyText = str.TrimEnd('\r', '\n');
                            isWinCopy = true;
                        }
                        if (isWinCopy)
                        {
                            //選択範囲、データを貼り付け
                            c1FlexGrid4.Select(HinmokuRow, HinmokuCol, HinmokuRowSel, HinmokuColSel, false);
                            c1FlexGrid4.Clip = strWinCopyText;
                            c1FlexGrid4.Select(HinmokuRow, HinmokuCol, HinmokuRowSel, HinmokuColSel);
                        }
                        else
                        {
                            if (copyData == null)
                            {
                                return;
                            }

                            // ペースト時に取り直されるので、このタイミングで保持しておく
                            int row = HinmokuRow; // 選択している開始行
                            int rowSel = HinmokuRowSel; // 選択している最終行
                            int col = HinmokuCol; // 選択している開始列
                            int colSel = HinmokuColSel; // 選択している最終列

                            int num = 0;

                            // 下から上対応（右から左に複数選択には未対応）
                            if (row > rowSel)
                            {
                                num = row;
                                row = rowSel;
                                rowSel = num;
                                num = 0;
                            }

                            // 右から左対応
                            if (col > colSel)
                            {
                                num = col;
                                col = colSel;
                                colSel = num;
                                num = 0;
                            }

                            // ▼行
                            // 範囲選択している行までペーストを繰り返す
                            //  i < rowSel ・・・選択開始行から終了行まで
                            //  i < c1FlexGrid4.Rows.Count・・・Gridの最終行まで
                            for (int i = row; i <= rowSel && i < c1FlexGrid4.Rows.Count; i++)
                            {
                                // ▼列
                                //for (int colIndex = 0; colIndex < copyData[num].Count && HinmokuColSel + colIndex < c1FlexGrid4.Cols.Count; colIndex++)
                                for (int j = 0; j + HinmokuCol < c1FlexGrid4.Cols.Count; j++)
                                {
                                    // コピー列が超えたら
                                    if (j > copyData[0].Count - 1)
                                    {
                                        break;
                                    }

                                    // 列データが存在しない場合、終了
                                    if (copyData[0][j] == null)
                                    {
                                        break;
                                    }
                                    c1FlexGrid4[i, col + j] = copyData[0][j];
                                }
                            }
                        }
                    }
                    // イベントをキャンセルする
                    e.Handled = true;
                }

                // Ctrl + Z
                if (e.KeyData == (Keys.Control | Keys.Z))
                {

                }
            }
            else
            {
                return;
            }
        }
        #endregion

        #region No.1417戻し
        // 調査品目明細GridのKeyDown
        private void c1FlexGrid4_KeyDown_0(object sender, KeyEventArgs e)
        {
            //調査品目編集モード 0:表示 1:編集
            if (ChousaHinmokuMode == 1)
            {
                // Ctrl + V
                if (e.KeyData == (Keys.Control | Keys.V))
                {
                    // 範囲選択時のペースト
                    // 750 行を複数行選択しペーストしたが1行分のみしかペースト出来ない対応

                    // 範囲選択が1列のみを選択している場合
                    //if (HinmokuCol == HinmokuColSel)
                    // 1列に限らずに貼り付ける
                    //if (HinmokuCol <= HinmokuColSel)
                    // 絶対にこの処理を通す
                    if (HinmokuCol <= HinmokuColSel || HinmokuCol >= HinmokuColSel)
                    {
                        if (copyData == null)
                        {
                            return;
                        }
                        // ペースト時に取り直されるので、このタイミングで保持しておく
                        int row = HinmokuRow; // 選択している開始行
                        int rowSel = HinmokuRowSel; // 選択している最終行
                        int col = HinmokuCol; // 選択している開始列
                        int colSel = HinmokuColSel; // 選択している最終列

                        int num = 0;

                        // 下から上対応（右から左に複数選択には未対応）
                        if (row > rowSel)
                        {
                            num = row;
                            row = rowSel;
                            rowSel = num;
                            num = 0;
                        }

                        // 右から左対応
                        if (col > colSel)
                        {
                            num = col;
                            col = colSel;
                            colSel = num;
                            num = 0;
                        }

                        // ▼行
                        // 範囲選択している行までペーストを繰り返す
                        //  i < rowSel ・・・選択開始行から終了行まで
                        //  i < c1FlexGrid4.Rows.Count・・・Gridの最終行まで
                        for (int i = row; i <= rowSel && i < c1FlexGrid4.Rows.Count; i++)
                        {
                            // ▼列
                            //for (int colIndex = 0; colIndex < copyData[num].Count && HinmokuColSel + colIndex < c1FlexGrid4.Cols.Count; colIndex++)
                            for (int j = 0; j + HinmokuCol < c1FlexGrid4.Cols.Count; j++)
                            {
                                // コピー列が超えたら
                                if (j > copyData[0].Count - 1)
                                {
                                    break;
                                }

                                // 列データが存在しない場合、終了
                                if (copyData[0][j] == null)
                                {
                                    break;
                                }
                                c1FlexGrid4[i, col + j] = copyData[0][j];
                            }
                        }
                    }
                    // イベントをキャンセルする
                    e.Handled = true;
                }

                // Ctrl + Z
                if (e.KeyData == (Keys.Control | Keys.Z))
                {

                }
            }
            else
            {
                return;
            }
        }
        #endregion

        #region 最新
        // 調査品目明細GridのKeyDown
        private void c1FlexGrid4_KeyDown(object sender, KeyEventArgs e)
        {
            //調査品目編集モード 0:表示 1:編集
            if (ChousaHinmokuMode == 1)
            {
                // Ctrl + V
                if (e.KeyData == (Keys.Control | Keys.V))
                {
                    // 範囲選択時のペースト
                    // 750 行を複数行選択しペーストしたが1行分のみしかペースト出来ない対応

                    // 範囲選択が1列のみを選択している場合
                    //if (HinmokuCol == HinmokuColSel)
                    // 1列に限らずに貼り付ける
                    //if (HinmokuCol <= HinmokuColSel)
                    // 絶対にこの処理を通す
                    if (HinmokuCol <= HinmokuColSel || HinmokuCol >= HinmokuColSel)
                    {
                        //No.1417 1182 奉行上でもコピペについて
                        bool isWinCopy = false;
                        IDataObject data = Clipboard.GetDataObject();
                        string strWinCopyText = "";
                        if (data.GetDataPresent(DataFormats.Text))
                        {
                            string str;
                            //クリップボードからデータを取得
                            str = (string)data.GetData(DataFormats.Text);
                            //クリップボードにある最後の開業コードを削除
                            strWinCopyText = str.TrimEnd('\r', '\n');

                            string strGridCopyText = "";
                            if (copyData != null && copyData.Count > 0)
                            {
                                for (int i = 0; i < copyData.Count; i++)
                                {
                                    string sLineText = "";
                                    for (int j = 0; j < copyData[i].Count; j++)
                                    {
                                        sLineText = sLineText + copyData[i][j];
                                    }
                                    strGridCopyText = strGridCopyText + sLineText;
                                }
                                strGridCopyText = strGridCopyText.Replace(Environment.NewLine, "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Replace(" ", "");
                            }
                            string strWin = strWinCopyText.Replace(Environment.NewLine, "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Replace(" ", "");
                            if (!strWin.Equals(strGridCopyText))
                                isWinCopy = true;
                        }
                        if (isWinCopy)
                        {
                            if (copyData == null)
                            {
                                copyData = new List<List<string>>();
                            }
                            copyData.Clear();
                            string[] lines = strWinCopyText.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                            for (int i = 0; i < lines.Length; i++)
                            {
                                copyData.Add(new List<string>());
                                string[] cols = lines[i].Split('\t');
                                for (int j = 0; j < cols.Length; j++)
                                {
                                    copyData[i].Add(cols[j]);
                                }
                            }
                        }
                        if (copyData == null)
                        {
                            return;
                        }
                        // ペースト時に取り直されるので、このタイミングで保持しておく
                        int row = HinmokuRow; // 選択している開始行
                        int rowSel = HinmokuRowSel; // 選択している最終行
                        int col = HinmokuCol; // 選択している開始列
                        int colSel = HinmokuColSel; // 選択している最終列

                        int num = 0;

                        // 下から上対応（右から左に複数選択には未対応）
                        if (row > rowSel)
                        {
                            num = row;
                            row = rowSel;
                            rowSel = num;
                            num = 0;
                        }

                        // 右から左対応
                        if (col > colSel)
                        {
                            num = col;
                            col = colSel;
                            colSel = num;
                            num = 0;
                        }
                        bool isRowBreak = false;
                        bool isColBreak = false;
                        // 貼り付け列数＜＝コピー列数
                        num = copyData[0].Count;
                        if ((colSel - col + 1) % num == 0)
                        {
                            isColBreak = true;
                        }
                        else
                        {
                            colSel = num + col - 1;
                        }

                        // 貼り付け行数＜＝コピー行数
                        num = copyData.Count;
                        // 貼り付け行数＞コピー行数
                        if ((rowSel - row + 1) % num == 0)
                        {
                            // 倍数行を選択する場合
                            isRowBreak = true;
                        }
                        else
                        {
                            rowSel = num + row - 1;
                        }

                        // ▼行
                        // 範囲選択している行までペーストを繰り返す
                        //  i < rowSel ・・・選択開始行から終了行まで
                        //  i < c1FlexGrid4.Rows.Count・・・Gridの最終行まで
                        num = 0;
                        for (int i = row; i <= rowSel && i < c1FlexGrid4.Rows.Count; i++)
                        {
                            // 選択行は
                            if (num >= copyData.Count)
                            {
                                if (isRowBreak) num = 0;
                                else break;
                            }
                            // ▼列
                            // コピー列の倍数を選択する場合
                            int iCol = 0;
                            for (int j = col; j <= colSel && j < c1FlexGrid4.Cols.Count; j++)
                            {
                                // コピー列が超えたら
                                if (iCol > copyData[num].Count - 1)
                                {
                                    if (isColBreak) iCol = 0;
                                    else break;
                                }

                                // 列データが存在しない場合、終了
                                if (copyData[num][iCol] == null)
                                {
                                    break;
                                }
                                c1FlexGrid4[i, j] = copyData[num][iCol];
                                iCol++;
                            }
                            num++;
                        }
                        //}
                    }
                    // イベントをキャンセルする
                    e.Handled = true;
                }

                // Ctrl + Z
                if (e.KeyData == (Keys.Control | Keys.Z))
                {

                }

            }
            else
            {
                return;
            }
        }
        #endregion

        // 調査品目明細の範囲選択
        private void c1FlexGrid4_SelChange(object sender, EventArgs e)
        {
            HinmokuRow = c1FlexGrid4.Row;       // 選択範囲の上端行番号
            HinmokuRowSel = c1FlexGrid4.RowSel; // 選択範囲の下端行番号
            HinmokuCol = c1FlexGrid4.Col;       // 選択範囲の上端列番号
            HinmokuColSel = c1FlexGrid4.ColSel; // 選択範囲の下端列番号
        }
        private void button2_PrintBushoBetsuTeishutsu_Click(object sender, EventArgs e)
        {
            set_error("", 0);

            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var Dt = new System.Data.DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                  "PrintDataPattern,PrintKikanFlg " +
                  "FROM " + "Mst_PrintList " +
                  "WHERE PrintListID = 11";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(Dt);
                //Boolean errorFLG = false;

                if (Dt.Rows.Count > 0)
                {
                    set_error("", 0);
                    // 11:部所別提出状況一覧表
                    if (Dt.Rows[0][0].ToString() == "8")
                    {
                        // string[]
                        // 5個分先に用意
                        string[] report_data = new string[3] { "", "", "" };

                        // 0.窓口ID
                        report_data[0] = MadoguchiID;
                        // 1.部所CD
                        report_data[1] = UserInfos[2];
                        // 2.呼び出し元画面
                        report_data[2] = "1";           // 呼び出し元画面（0:窓口ミハル、1:特命課長、2:自分大臣）

                        string[] result = GlobalMethod.InsertMadoguchiReportWork(11, UserInfos[0], report_data, "BushoBetsu");

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
                conn.Close();
            }

        }

        // ページ移動
        private void btnGoPage_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();

            set_error("", 0);

            int go_page = 0;
            int max_page = 0;
            int.TryParse(item3_TargetPage.Text, out go_page);
            int.TryParse(Paging_all.Text, out max_page);

            // ページ移動が数値化出来て、最大ページ数よりも小さい場合
            if (go_page != 0 && go_page <= max_page)
            {
                Paging_now.Text = (go_page).ToString();
                Grid_Visible(int.Parse(Paging_now.Text));
            }
            else
            {
                // E70062:移動ページ数は、最大ページ数以下を入力してください。
                set_error(GlobalMethod.GetMessage("E70062", ""));
            }

            //レイアウトロジックを再開する
            this.ResumeLayout();


        }

        // ページ数
        private void item3_TargetPage_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 0～9と、バックスペース以外の時は、イベントをキャンセルする
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        // Garoon宛先追加
        private void c1FlexGrid5_CellChanged(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            DataTable combodt = new DataTable();
            string discript = "GyoumuBushoCD";
            string value = "GyoumuBushoCD";
            string table = "Mst_Busho";
            string where = "";

            // 2:調担当部所
            if (e.Col == 2)
            {
                // 部所変更時にユーザーが部所に所属していなければクリア
                if (c1FlexGrid5.Rows[e.Row][e.Col + 1] != null && c1FlexGrid5.Rows[e.Row][e.Col + 1].ToString() != "")
                {
                    combodt = new DataTable();
                    discript = "KojinCD";
                    value = "GyoumuBushoCD";
                    table = "Mst_Chousain";
                    where += "GyoumuBushoCD = '" + c1FlexGrid5.Rows[e.Row][e.Col].ToString() + "' AND KojinCD = '" + c1FlexGrid5.Rows[e.Row][e.Col + 1].ToString() + "' ";
                    combodt = new DataTable();
                    combodt = GlobalMethod.getData(discript, value, table, where);
                    if (combodt != null && combodt.Rows.Count > 0)
                    {
                        // 部所でユーザーが取得できるので問題ない
                    }
                    else
                    {
                        // 部所に所属したユーザーではないのでクリア
                        c1FlexGrid5.Rows[e.Row][e.Col + 1] = "";
                    }
                }
            }
        }

        private void btnGridSize_Click(object sender, EventArgs e)
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:457 → 914
            //    // width:1876 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid4.Height = 914;
            //    c1FlexGrid4.Width = 3752;
            //}
            //else
            //{
            //    // height:914 → 457
            //    // width:3752 → 1876
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid4.Height = 457;
            //    c1FlexGrid4.Width = 1876;
            //}
            gridSizeChange();
        }
        private void gridSizeChange()
        {
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:457 → 914
            //    // width:1820 → 3752
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid4.Height = 914;
            //    c1FlexGrid4.Width = 3752;
            //}
            //else
            //{
            //    // height:914 → 457
            //    // width:3752 → 1820
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid4.Height = 457;
            //    c1FlexGrid4.Width = 1820;
            //}
            //if (btnGridSize.Text == "一覧拡大")
            //{
            //    // height:457 → 914
            //    // width:1820 → 3752
            //    c1FlexGrid4.MaximumSize = new System.Drawing.Size(0, 914);
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid4.Height = 914;
            //    c1FlexGrid4.Width = this.Size.Width - 40;
            //    //c1FlexGrid4.Width = 3752;
            //    //c1FlexGrid4.Dock = DockStyle.Fill;
            //}
            //else
            //{
            //    // height:914 → 457
            //    // width:3752 → 1820
            //    c1FlexGrid4.MaximumSize = new System.Drawing.Size(1820, 457);
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid4.Height = 457;
            //    c1FlexGrid4.Width = 1820;
            //    //c1FlexGrid4.Dock = DockStyle. None;
            //}

            string num = "";
            int smallWidth = 0;
            int smallHeight = 0;
            int bigHeight = 0;
            int maximumWidth = 0;
            int minimumWidth = 0;
            int padding = 0;

            num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_HINMOKU_GRID_MAX_WIDTH");
            if (num != null)
            {
                Int32.TryParse(num, out maximumWidth);
                if (maximumWidth == 0)
                {
                    maximumWidth = 1820;
                }
            }
            num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_HINMOKU_GRID_MIN_WIDTH");
            if (num != null)
            {
                Int32.TryParse(num, out minimumWidth);
                if (minimumWidth == 0)
                {
                    minimumWidth = 1820;
                }
            }
            if (btnGridSize.Text == "一覧拡大")
            {
                num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_HINMOKU_GRID_BIG_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out bigHeight);
                    if (bigHeight == 0)
                    {
                        bigHeight = 914;
                    }
                }
                num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_HINMOKU_GRID_PADDING");
                if (num != null)
                {
                    Int32.TryParse(num, out padding);
                    if (padding == 0)
                    {
                        padding = 40;
                    }
                }

                // height:457 → 914
                // width:1820 → 3752
                //c1FlexGrid4.MaximumSize = new System.Drawing.Size(0, 914);
                c1FlexGrid4.MaximumSize = new System.Drawing.Size(0, bigHeight);
                c1FlexGrid4.MinimumSize = new System.Drawing.Size(minimumWidth, bigHeight);
                btnGridSize.Text = "一覧縮小";
                //c1FlexGrid4.Height = 914;
                c1FlexGrid4.Height = bigHeight;
                //c1FlexGrid4.Width = this.Size.Width - 40;
                c1FlexGrid4.Width = this.Size.Width - padding;
                //c1FlexGrid4.Width = 3752;

                // 1201 最大化する
                // 調査品目明細タブを初めて開いた
                // 0:調査品目明細を開いたことが無い 1:調査品目明細を開いたことがある
                if (tabChousahinmokuFlg == "0")
                {
                    // タブ移動時の最大化 1:最大化する それ以外:最大化しない
                    num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_HINMOKU_TAB_MOVE_MAXIMIZE");
                    if ("1".Equals(num))
                    {
                        this.WindowState = FormWindowState.Maximized;
                    }
                }
                else
                {
                    // タブ移動ではない場合(拡大ボタンを押した)
                    // // 0:タブ移動してない 1:タブ移動した
                    if (tabChangeFlg == "0")
                    {
                        num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_HINMOKU_TAB_MOVE_MAXIMIZE");
                        if ("1".Equals(num))
                        {
                            this.WindowState = FormWindowState.Maximized;
                        }
                    }
                }
            }
            else
            {
                num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_HINMOKU_GRID_SMALL_WIDTH");
                if (num != null)
                {
                    Int32.TryParse(num, out smallWidth);
                    if (smallWidth == 0)
                    {
                        smallWidth = 1820;
                    }
                }
                num = GlobalMethod.GetCommonValue1("TOKUMEIKACHO_HINMOKU_GRID_SMALL_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out smallHeight);
                    if (smallHeight == 0)
                    {
                        smallHeight = 457;
                    }
                }

                // height:914 → 457
                // width:3752 → 1820
                //c1FlexGrid4.MaximumSize = new System.Drawing.Size(1820, 457);
                c1FlexGrid4.MaximumSize = new System.Drawing.Size(maximumWidth, smallHeight);
                c1FlexGrid4.MinimumSize = new System.Drawing.Size(minimumWidth, smallHeight);
                btnGridSize.Text = "一覧拡大";
                //c1FlexGrid4.Height = 457;
                c1FlexGrid4.Height = smallHeight;
                //c1FlexGrid4.Width = 1820;
                c1FlexGrid4.Width = smallWidth;
            }
        }

        private void Tokumei_Input_KeyDown(object sender, KeyEventArgs e)
        {
            Control c = this.ActiveControl;

            //レイアウトロジックを停止する
            this.SuspendLayout();
            // コンボボックス以外で
            if (c != null && (c.GetType().Equals(typeof(ComboBox))
                //|| c.GetType().Equals(typeof(TextBox)) 
                || c.GetType().Equals(typeof(C1.Win.C1FlexGrid.C1FlexGrid))
                || c.GetType().ToString().Equals("C1.Win.C1FlexGrid.GridEditorTextBox")
                || c.GetType().ToString().Equals("C1.Win.C1FlexGrid.GridEditorComboBox")
                || c.GetType().ToString().Equals("C1.Win.C1FlexGrid.GridEditorNumericTextBox")
                || c.GetType().ToString().Equals("C1.Win.C1FlexGrid.GridEditorDatePicker")
                ))
            {
                // ↑↓押下時、コンボボックスがアクティブだった場合は、コンボの値変更を優先し、
                // 画面スクロールは動かさない
                // タブのタイトルを取得 引合、入札、契約、技術者評価
                string tabName = this.tab.SelectedTab.Text;
                if (e.KeyCode == Keys.PageDown)
                {
                    if ("調査概要".Equals(tabName))
                    {
                        this.tabPage1.AutoScrollPosition = new System.Drawing.Point(-this.tabPage1.AutoScrollPosition.X, -this.tabPage1.AutoScrollPosition.Y + 600);
                    }
                    if ("担当部所".Equals(tabName))
                    {
                        this.tabPage2.AutoScrollPosition = new System.Drawing.Point(-this.tabPage2.AutoScrollPosition.X, -this.tabPage2.AutoScrollPosition.Y + 600);
                    }
                    if ("調査品目明細".Equals(tabName))
                    {
                        this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y + 600);
                    }
                }
                if (e.KeyCode == Keys.PageUp)
                {
                    if ("調査概要".Equals(tabName))
                    {
                        this.tabPage1.AutoScrollPosition = new System.Drawing.Point(-this.tabPage1.AutoScrollPosition.X, -this.tabPage1.AutoScrollPosition.Y - 600);
                    }
                    if ("担当部所".Equals(tabName))
                    {
                        this.tabPage2.AutoScrollPosition = new System.Drawing.Point(-this.tabPage2.AutoScrollPosition.X, -this.tabPage2.AutoScrollPosition.Y - 600);
                    }
                    if ("調査品目明細".Equals(tabName))
                    {
                        this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y - 600);
                    }
                }
            }
            else
            {
                // タブのタイトルを取得 引合、入札、契約、技術者評価
                string tabName = this.tab.SelectedTab.Text;
                if (e.KeyCode == Keys.PageDown || e.KeyCode == Keys.Down)
                {
                    if ("調査概要".Equals(tabName))
                    {
                        this.tabPage1.AutoScrollPosition = new System.Drawing.Point(-this.tabPage1.AutoScrollPosition.X, -this.tabPage1.AutoScrollPosition.Y + 600);
                    }
                    if ("担当部所".Equals(tabName))
                    {
                        this.tabPage2.AutoScrollPosition = new System.Drawing.Point(-this.tabPage2.AutoScrollPosition.X, -this.tabPage2.AutoScrollPosition.Y + 600);
                    }
                    if ("調査品目明細".Equals(tabName))
                    {
                        this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y + 600);
                    }
                }
                if (e.KeyCode == Keys.PageUp || e.KeyCode == Keys.Up)
                {
                    if ("調査概要".Equals(tabName))
                    {
                        this.tabPage1.AutoScrollPosition = new System.Drawing.Point(-this.tabPage1.AutoScrollPosition.X, -this.tabPage1.AutoScrollPosition.Y - 600);
                    }
                    if ("担当部所".Equals(tabName))
                    {
                        this.tabPage2.AutoScrollPosition = new System.Drawing.Point(-this.tabPage2.AutoScrollPosition.X, -this.tabPage2.AutoScrollPosition.Y - 600);
                    }
                    if ("調査品目明細".Equals(tabName))
                    {
                        this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y - 600);
                    }
                }
            }
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        // 調査品目Gridのマウスホイールイベント
        private void c1FlexGrid4_MouseWheel(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            // VIPS 20220414 コンポーネント最新化にあたり修正
            // e.Deltaがマイナス値だと↓、プラス値だと↑
            //this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y - e.Delta);
        }

        //不具合No1207(903)
        //mode:0 デフォルト行の高さ、mode:0以外 行幸自動調整
        private void gridRowHeightAutoResize(string mode)
        {
            //タブ選択時に現在の選択値を保持する。//タブ選択する度にDBのデフォルト設定に戻してよい場合はこの1行コメントアウト
            AutoSizeGridRowMode = mode;
            //固定行高さにする場合、ボタンキャプションは自動調整にする
            if (mode == "0")
            {
                btnRowSizeChange.Text = GRID_ROW_AUTO_SIZE;
            }
            else
            {
                btnRowSizeChange.Text = GRID_ROW_FIX_SIZE;
            }
            //Gridのヘッダの次の１行目から行の最後まで。
            for (int i = 1; i < c1FlexGrid4.Rows.Count; i++)
            {
                //デフォルト行高さ
                if (mode == "0")
                {
                    c1FlexGrid4.Rows[i].Height = -1;    //-1 にすると、グリッドのデフォルトの行高になる
                }
                //自動行高調整
                else
                {
                    c1FlexGrid4.AutoSizeRow(i);
                }
            }
        }

        private void btnRowSizeChange_Click(object sender, EventArgs e)
        {
            if (btnRowSizeChange.Text == GRID_ROW_AUTO_SIZE)
            {
                gridRowHeightAutoResize("1");
            }
            else
            {
                gridRowHeightAutoResize("0");
            }
        }

        private void tab_DrawItem(object sender, DrawItemEventArgs e)
        {
            GlobalMethod.tabDisplaySet(tab, sender, e);
        }

        private void button4_2_Click(object sender, EventArgs e)
        {
            //奉行エクセル　
            //Popup_GroupMei form = new Popup_GroupMei();
            //form.ShowDialog();
        }

        private void button6_2_Click(object sender, EventArgs e)
        {
            //奉行エクセル　グループ名まで移動
            c1FlexGrid4.LeftCol = 59;
            Console.WriteLine(c1FlexGrid4.BottomRow);
        }

        private void c1FlexGrid4_BeforeEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            //奉行エクセル　集計表VerがVer1の場合に選択不可
            if (e.Col == c1FlexGrid4.Cols["BunkatsuHouhou"].Index)
            {
                if (c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() == "-")
                {
                    e.Cancel = true;
                }
            }
            if (e.Col == c1FlexGrid4.Cols["GroupMei"].Index)
            {
                if (c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() == "-")
                {
                    e.Cancel = true;
                }
            }
        }
    }
}
