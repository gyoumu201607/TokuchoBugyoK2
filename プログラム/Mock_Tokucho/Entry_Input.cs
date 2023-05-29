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
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Configuration;
using System.Diagnostics;
using C1.Win.C1FlexGrid;
using DataTable = System.Data.DataTable;
using System.Collections;
using TokuchoBugyoK2.TokuchoBugyoKDataSetTableAdapters;
using System.Drawing.Text;

namespace TokuchoBugyoK2
{
    public partial class Entry_Input : Form
    {
        private String pgmName = "Entry_Input";

        private System.Data.DataTable AnkenData_H = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_N = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_K = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_G = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_Grid1 = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_Grid2 = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_Grid3 = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_Grid4 = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_Grid5 = new System.Data.DataTable();
        // VIPS 20220415 コンポーネント最新化にあたり修正
        private Image Img_DeleteRowNonactive;

        private string Message = "";
        public string[] UserInfos;
        GlobalMethod GlobalMethod = new GlobalMethod();
        private Boolean ownerflg = true;
        private int saishinFLG;
        private Boolean KianKaijoFLG = false;
        private Boolean KianFLG = false;
        private string BushoCD = "";
        private string c1FlexGrid2Data = "";
        private string beforeKeikakuBangou = "";

        public string mode = "";
        public string AnkenID = "";
        public string AnkenbaBangou = "";
        public string KeikakuID = "";
        // 変更伝票がどのボタンから遷移したかのフラグ
        public int ChangeFlag = 0;

        // えんとり君修正STEP2
        private string sFolderRenameBef = "";    //ファイルを移動するため、変更前のフォルダを保存する
        private string sFolderBushoCDRenameBef = "";    //受託課所支部（契約部所）
        private string sFolderYearRenameBef = "";   // 工期開始年度
        private string sFolderGyomuRenameBef = "";   // 業務名称
        private string sFolderOrderRenameBef = "";//発注者名
        private string sItem1_10_ori = ""; //受託課所支部（契約部所）DB値
        private string sItem1_2_KoukiNendo_ori = ""; //工期開始年度DB値
        private string sAnkenSakuseiKubun_ori = ""; // 案件区分変更前の値
        private string sJigyoubuHeadCD_ori = "";    // 事業部ヘッダーコード

        // VIPS 20220221 課題管理表No.1273(967) ADD 計画番号コピー制御用
        // この業務を元に新規登録ボタン                ：1
        // この案件番号の枝番で受託番号を作成するボタン：2
        public string CopyMode = "";

        //エントリ君修正STEP1
        //計画詳細画面の「前回案件番号を元に新規登録」ボタンから遷移してきたときTrue
        public bool isKeikakuAnkenNew = false;

        public Entry_Input()
        {
            InitializeComponent();

            // 契約タブの契約図書のテキスト非表示
            item3_1_26.Visible = false;
            // 技術者評価タブの請求書のテキスト非表示
            item4_1_8.Visible = false;
            // 工期開始年度にマウスホイールイベントを付与
            this.item1_2_KoukiNendo.MouseWheel += item1_3_MouseWheel;
            // 売上年度にマウスホイールイベントを付与
            this.item1_3.MouseWheel += item1_3_MouseWheel;
            // 受託課所支部にマウスホイールイベントを付与
            this.item1_10.MouseWheel += item_MouseWheel;
            // 契約タブの売上年度にマウスホイールイベントを付与
            this.item3_1_5.MouseWheel += item_MouseWheel;

            this.item1_1.MouseWheel += item1_3_MouseWheel; // 引合状況
            this.item1_2.MouseWheel += item1_3_MouseWheel; // 案件区分
            this.item1_14.MouseWheel += item1_3_MouseWheel; // 契約区分
            this.item1_15.MouseWheel += item1_3_MouseWheel; // 入札方式
            this.item1_17.MouseWheel += item1_3_MouseWheel; // 入札状況
            this.item1_34.MouseWheel += item1_3_MouseWheel; // 参考見積
            this.item1_35.MouseWheel += item1_3_MouseWheel; // 受注意欲
            this.item2_1_1.MouseWheel += item1_3_MouseWheel; // 入札状況
            this.item2_2_1.MouseWheel += item1_3_MouseWheel; // 当会応札
            this.item2_2_3.MouseWheel += item1_3_MouseWheel; // 受注意欲
            this.item2_1_4.MouseWheel += item1_3_MouseWheel; // 引合状況
            this.item2_2_2.MouseWheel += item1_3_MouseWheel; // 参考見積
            this.item2_1_5.MouseWheel += item1_3_MouseWheel; // 案件区分
            this.item2_3_1.MouseWheel += item1_3_MouseWheel; // 落札者状況
            this.item2_3_2.MouseWheel += item1_3_MouseWheel; // 落札額状況
            this.item3_1_1.MouseWheel += item1_3_MouseWheel; // 案件区分
            this.item3_1_8.MouseWheel += item1_3_MouseWheel; // 契約区分

            //エントリ君修正STEP2
            this.ErrorMessage.Font = new System.Drawing.Font(this.ErrorMessage.Font.Name, float.Parse(GlobalMethod.GetCommonValue1("DSP_ERROR_FONTSIZE")));
        }

        //不具合No1388対応
        private void Entry_Input_Shown(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.Refresh();
            this.TopMost = false;
        }

        private void Entry_Input_Load(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();

            //不具合No1017（751）
            //タブの文字装飾変更対応
            //文字表示を大きくする場合は、デザイナでTabのItemSize.widthを変更する。窓口、特命課長、自分大臣は、125で設定すると、14ポイントぐらいのサイズでいける
            tab.DrawMode = TabDrawMode.OwnerDrawFixed;

            //GlobalMethod.outputLogger("entory_input", "entory_input load 開始 " + DateTime.Now, "GetAnkenJouhou", UserInfos[1]);
            //不具合No1355（1123）
            lblVersion.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                lblBootMode.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            //ユーザ名を設定
            label7.Text = UserInfos[3] + "：" + UserInfos[1];

            // 昇順降順アイコン設定
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid1.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid2.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid2.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid3.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid3.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid4.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid4.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid5.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid5.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            //コンボ内容の設定
            set_combo();

            c1FlexGrid4.Rows[0].AllowMerging = true;
            c1FlexGrid4.Rows[2][0] = "1回目";
            c1FlexGrid4.Rows[1][1] = "工期末日付";
            c1FlexGrid4.Rows[1][2] = "計上月";
            c1FlexGrid4.Rows[1][3] = "計上額";
            c1FlexGrid4.Rows[1][9] = "工期末日付";
            c1FlexGrid4.Rows[1][10] = "計上月";
            c1FlexGrid4.Rows[1][11] = "計上額";
            c1FlexGrid4.Rows[1][17] = "工期末日付";
            c1FlexGrid4.Rows[1][18] = "計上月";
            c1FlexGrid4.Rows[1][19] = "計上額";
            c1FlexGrid4.Rows[1][25] = "工期末日付";
            c1FlexGrid4.Rows[1][26] = "計上月";
            c1FlexGrid4.Rows[1][27] = "計上額";

            // VIPS 20220415 コンポーネント最新化にあたり修正
            Img_DeleteRowNonactive = Image.FromFile("Resource/Image/DeleteRow.gif");


            int num = 2;
            // 20210412 不具合があったため、一度コメントアウト
            // 最初に12月分表示する
            for (int i = 0; i < 11; i++)
            {
                num = i + 2;
                c1FlexGrid4.Rows.Add();
                c1FlexGrid4.Rows[num + 1][0] = num + "回目";
                Resize_Grid("c1FlexGrid4");
            }

            //新規モード以外ではデータを読み込み
            if (AnkenID != "")
            {
                get_date();
            }

            C1.Win.C1FlexGrid.CellRange rng = c1FlexGrid4.GetCellRange(0, 0, 0, 27);
            rng.Style = c1FlexGrid4.Styles["FixedBumon"];

            this.Owner.Hide();

            // button1  この業務を元に新規登録ボタン
            // button10 この案件番号の枝番で受託番号を作成するボタン
            // 上記二つは当面非表示で


            // IMEモード変更
            // 引合タブ
            item1_27.ImeMode = ImeMode.Disable; //  電話
            item1_28.ImeMode = ImeMode.Disable; //  FAX
            item1_29.ImeMode = ImeMode.Disable; //  E-Mail
            item1_30.ImeMode = ImeMode.Disable; //  郵便番号
            item1_36.ImeMode = ImeMode.Disable; // 参考見積額（税抜）
            item1_7_1_1_1.ImeMode = ImeMode.Disable; // 部門配分 引合・入札 配分率 調査部
            item1_7_1_2_1.ImeMode = ImeMode.Disable; // 部門配分 引合・入札 配分率 事業普及部
            item1_7_1_3_1.ImeMode = ImeMode.Disable; // 部門配分 引合・入札 配分率 情報システム部
            item1_7_1_4_1.ImeMode = ImeMode.Disable; // 部門配分 引合・入札 配分率 総合研究所
            item1_7_2_1_1.ImeMode = ImeMode.Disable;
            item1_7_2_2_1.ImeMode = ImeMode.Disable;
            item1_7_2_3_1.ImeMode = ImeMode.Disable;
            item1_7_2_4_1.ImeMode = ImeMode.Disable;
            item1_7_2_5_1.ImeMode = ImeMode.Disable;
            item1_7_2_6_1.ImeMode = ImeMode.Disable;
            item1_7_2_7_1.ImeMode = ImeMode.Disable;
            item1_7_2_8_1.ImeMode = ImeMode.Disable;
            item1_7_2_9_1.ImeMode = ImeMode.Disable;
            item1_7_2_10_1.ImeMode = ImeMode.Disable;
            item1_7_2_11_1.ImeMode = ImeMode.Disable;
            item1_7_2_12_1.ImeMode = ImeMode.Disable;

            // 入札タブ
            item2_3_5.ImeMode = ImeMode.Disable; // 予定価格（税抜）
            item2_4_1_1_1.ImeMode = ImeMode.Disable; // 部門配分 引合・入札 配分率 調査部
            item2_4_1_2_1.ImeMode = ImeMode.Disable; // 部門配分 引合・入札 配分率 事業普及部
            item2_4_1_3_1.ImeMode = ImeMode.Disable; // 部門配分 引合・入札 配分率 情報システム部
            item2_4_1_4_1.ImeMode = ImeMode.Disable; // 部門配分 引合・入札 配分率 総合研究所
            item2_4_2_1_1.ImeMode = ImeMode.Disable;
            item2_4_2_2_1.ImeMode = ImeMode.Disable;
            item2_4_2_3_1.ImeMode = ImeMode.Disable;
            item2_4_2_4_1.ImeMode = ImeMode.Disable;
            item2_4_2_5_1.ImeMode = ImeMode.Disable;
            item2_4_2_6_1.ImeMode = ImeMode.Disable;
            item2_4_2_7_1.ImeMode = ImeMode.Disable;
            item2_4_2_8_1.ImeMode = ImeMode.Disable;
            item2_4_2_9_1.ImeMode = ImeMode.Disable;
            item2_4_2_10_1.ImeMode = ImeMode.Disable;
            item2_4_2_11_1.ImeMode = ImeMode.Disable;
            item2_4_2_12_1.ImeMode = ImeMode.Disable;

            // 契約タブ
            item3_1_10.ImeMode = ImeMode.Disable; // 消費税率
            item3_1_12.ImeMode = ImeMode.Disable; // 税抜（自動計算用）
            item3_1_13.ImeMode = ImeMode.Disable; // 税込
            item3_1_14.ImeMode = ImeMode.Disable; // 内消費税
            item3_1_15.ImeMode = ImeMode.Disable; // 受託金額（税込）
            item3_1_16.ImeMode = ImeMode.Disable; // 受託外金額（税込）
            //えんとり君修正STEP2
            item3_1_27.ImeMode = ImeMode.Disable;

            item3_2_1_1.ImeMode = ImeMode.Disable; // 受託金額配分 配分額（税抜） 調査部
            item3_2_2_1.ImeMode = ImeMode.Disable; // 受託金額配分 配分額（税抜） 事業普及部
            item3_2_3_1.ImeMode = ImeMode.Disable; // 受託金額配分 配分額（税抜） 情報システム部
            item3_2_4_1.ImeMode = ImeMode.Disable; // 受託金額配分 配分額（税抜） 総合研究所
            item3_2_1_2.ImeMode = ImeMode.Disable; // 受託金額配分 配分額（税込） 調査部
            item3_2_2_2.ImeMode = ImeMode.Disable; // 受託金額配分 配分額（税込） 事業普及部
            item3_2_3_2.ImeMode = ImeMode.Disable; // 受託金額配分 配分額（税込） 情報システム部
            item3_2_4_2.ImeMode = ImeMode.Disable; // 受託金額配分 配分額（税込） 総合研究所

            item3_3_1.ImeMode = ImeMode.Disable; // 調査部
            item3_3_2.ImeMode = ImeMode.Disable; // 事業普及部
            item3_3_3.ImeMode = ImeMode.Disable; // 情報システム部
            item3_3_4.ImeMode = ImeMode.Disable; // 総合研究所

            item3_7_1.ImeMode = ImeMode.Disable; // 調査部
            item3_7_2.ImeMode = ImeMode.Disable; // 事業普及部
            item3_7_3.ImeMode = ImeMode.Disable; // 情報システム部
            item3_7_4.ImeMode = ImeMode.Disable; // 総合研究所

            item3_6_1.ImeMode = ImeMode.Disable;
            item3_6_2.ImeMode = ImeMode.Disable;
            item3_6_3.ImeMode = ImeMode.Disable;
            item3_6_4.ImeMode = ImeMode.Disable;
            item3_6_5.ImeMode = ImeMode.Disable;
            item3_6_6.ImeMode = ImeMode.Disable;
            item3_6_7.ImeMode = ImeMode.Disable;
            item3_6_8.ImeMode = ImeMode.Disable;
            item3_6_9.ImeMode = ImeMode.Disable;
            item3_6_10.ImeMode = ImeMode.Disable;
            item3_6_11.ImeMode = ImeMode.Disable;
            item3_6_12.ImeMode = ImeMode.Disable;

            item3_7_2_14_1.ImeMode = ImeMode.Disable;
            item3_7_2_15_1.ImeMode = ImeMode.Disable;
            item3_7_2_16_1.ImeMode = ImeMode.Disable;
            item3_7_2_17_1.ImeMode = ImeMode.Disable;
            item3_7_2_18_1.ImeMode = ImeMode.Disable;
            item3_7_2_19_1.ImeMode = ImeMode.Disable;
            item3_7_2_20_1.ImeMode = ImeMode.Disable;
            item3_7_2_21_1.ImeMode = ImeMode.Disable;
            item3_7_2_22_1.ImeMode = ImeMode.Disable;
            item3_7_2_23_1.ImeMode = ImeMode.Disable;
            item3_7_2_24_1.ImeMode = ImeMode.Disable;
            item3_7_2_25_1.ImeMode = ImeMode.Disable;

            // 技術者評価タブ
            item4_1_1.ImeMode = ImeMode.Disable; // 評点（業務）
            item4_1_3.ImeMode = ImeMode.Disable; // 管理技術者の評点
            item4_1_5.ImeMode = ImeMode.Disable; // 照査技術者の評点

            // Role:2システム管理者 以外の場合、起案解除は非表示
            if (!UserInfos[4].Equals("2"))
            {
                button17.Visible = false;
            }

            //新規モード
            if (mode == "insert" || mode == "keikaku")
            {
                item1_2.SelectedValue = "01";
                item1_2.Enabled = false;
                //登録日に当日をセット
                item1_9.CustomFormat = "";
                item1_9.Text = DateTime.Today.ToString();

                //// 過去案件情報の前回受託番号IDに1を入れておく
                //c1FlexGrid1.Rows[1][16] = 1;

                //売上年度の設定
                /*
                String discript = "NendoSeireki";
                String value = "NendoID ";
                String table = "Mst_Nendo";
                String where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
                //コンボボックスデータ取得
                DataTable dt = GlobalMethod.getData(discript, value, table, where);
                if (dt != null)
                {
                    item1_3.SelectedValue = dt.Rows[0][0].ToString();
                }
                else
                {
                    item1_3.SelectedValue = System.DateTime.Now.Year;
                }
                */
                item1_3.SelectedValue = GlobalMethod.GetTodayNendo();
                item1_2_KoukiNendo.SelectedValue = GlobalMethod.GetTodayNendo();

                // 新規時 受託課所支部を自部所にセット
                //item1_10.SelectedValue = UserInfos[2];
                if (AnkenID != null && AnkenID != "")
                {
                    DataTable ankenDt = GlobalMethod.getData("AnkenJutakubushoCD", "AnkenJutakubushoCD", "AnkenJouhou", "AnkenJouhouID = " + AnkenID);
                    if (ankenDt != null && ankenDt.Rows.Count > 0)
                    {
                        item1_10.SelectedValue = ankenDt.Rows[0][0].ToString();
                    }
                }
                else
                {
                    item1_10.SelectedValue = UserInfos[2];
                }

                // 案件（受託）フォルダ初期値設定
                if (AnkenbaBangou == "")
                {
                    String discript = "FolderPath";
                    String value = "FolderPath ";
                    String table = "M_Folder";
                    String where = "MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + UserInfos[2] + "' ";

                    // //xxxx/00Cyousa/00調査情報部門共有/$NENDO$/200受託調査関連
                    //string FolderBase = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", item1_3.SelectedValue.ToString());
                    // フォルダ関連は工期開始年度で作成する
                    string FolderBase = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", item1_2_KoukiNendo.SelectedValue.ToString());
                    string FolderPath = "";

                    DataTable dt = new System.Data.DataTable();
                    dt = GlobalMethod.getData(discript, value, table, where);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        // $FOLDER_BASE$/004 本部
                        FolderPath = dt.Rows[0][0].ToString();
                    }
                    if (FolderBase != "" && FolderPath != "")
                    {
                        FolderPath = FolderPath.Replace(@"$FOLDER_BASE$", FolderBase);
                        FolderPath = FolderPath.Replace("/", @"\");
                    }

                    // 案件（受託）フォルダ
                    item1_12.Text = FolderPath;

                }

                //元にする案件あり
                if (AnkenID != "")
                {
                    if (AnkenbaBangou == "")
                    {
                        //「この案件を元に新規登録」
                        item1_1.SelectedValue = "1";
                        item1_6.Text = "";
                        item1_7.Text = "";
                        item1_8.Text = "";
                        item1_9.CustomFormat = "";
                        item1_9.Text = DateTime.Today.ToString();
                        //item1_10.SelectedValue = UserInfos[2]; // 上でセット
                        item1_16.Text = "";
                        item1_16.CustomFormat = " ";
                        item1_17.SelectedValue = "1";
                        //item1_18.Text = "";
                        item1_35.SelectedValue = "1";
                    }
                    else
                    {
                        //「この案件番号の枝番で受託番号を作成する」
                        item1_36.Text = GetMoneyTextLong(0);

                        c1FlexGrid1.Rows.Count = 2;
                        Resize_Grid("c1FlexGrid1");

                        // 過去案件情報のRakusatushaIDに1をセット
                        c1FlexGrid1.Rows[1][16] = 1;
                    }
                    //「この案件を元に新規登録」「この案件番号の枝番で受託番号を作成する」共通初期化
                    //item1_12.Text = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", item1_3.SelectedValue.ToString());
                    //item1_19.Text = "";
                    //item1_20.Text = "";
                    //item1_21.Text = "";
                    //item1_22.Text = "";
                    //item1_23.Text = "";
                    //item1_24.Text = ""; //発注者課名
                    item1_34.SelectedValue = "1"; //参考見積
                    //item1_7_1_1_1.Text = "0.00%"; //部門配分比率
                    //item1_7_1_2_1.Text = "0.00%";
                    //item1_7_1_3_1.Text = "0.00%";
                    //item1_7_1_4_1.Text = "0.00%";
                    //item1_7_1_5_1.Text = "0.00%";
                }
                else
                {
                    // 新規時
                    // 過去案件情報のRakusatushaIDに1をセット
                    c1FlexGrid1.Rows[1][16] = 1;
                }
                //計画詳細の「新規案件」ボタン押下時
                if (mode == "keikaku")
                {
                    DataTable Keikakudt = GlobalMethod.getData("KeikakuBangou", "KeikakuAnkenMei", "KeikakuJouhou", "KeikakuID = " + KeikakuID);
                    if (Keikakudt != null && Keikakudt.Rows.Count > 0)
                    {
                        item1_4.Text = Keikakudt.Rows[0][1].ToString();
                        item1_5.Text = Keikakudt.Rows[0][0].ToString();
                    }
                }
                //「この案件番号の枝番で受託番号を作成」ボタン押下時
                if (mode == "insert" && AnkenbaBangou != null && AnkenbaBangou != "")
                {
                    int Eda = 0;
                    string EdaStr = "";
                    // 基にした案件番号の枝番に+1しただけだと重複が発生するので、
                    // 案件番号で枝番の最大値を取得し、+1したものを枝番として採用する
                    //if (item1_8.Text != "")
                    //{
                    //    // 数値化する
                    //    if (int.TryParse(item1_8.Text, out Eda))
                    //    {
                    //        Eda++;
                    //    }
                    //    else
                    //    {
                    //        // 枝番が数値以外場合（移行データの場合、Bとかがありうる）
                    //        Eda = 1;
                    //    }
                    //    Eda = int.Parse(item1_8.Text);
                    //}
                    //else
                    //{
                    //    // 枝番が空の場合（移行データの場合）
                    //    Eda = 1;
                    //}

                    // 案件番号の中で、枝番が数値で構成されていて、最大の物を取得する
                    DataTable AnkenEdadt = GlobalMethod.getData("' '", "max(AnkenJutakuBangouEda)", "AnkenJouhou", "AnkenAnkenBangou = '" + item1_6.Text + "' AND AnkenJutakuBangouEda LIKE '%[0-9]%' AND AnkenDeleteFlag = 0");
                    if (AnkenEdadt != null && AnkenEdadt.Rows.Count > 0)
                    {
                        // 取得できた場合、最大値+1する
                        EdaStr = AnkenEdadt.Rows[0][0].ToString();
                        if (int.TryParse(EdaStr, out Eda))
                        {
                            Eda += 1;
                        }
                        else
                        {
                            // 変換できなかった場合、1とする（MAXで取ってるので、Whereで引っかかるものがなければNULL）
                            Eda = 1;
                        }
                    }
                    else
                    {
                        // 取得できなかった場合、1とする
                        Eda = 1;
                    }

                    item1_8.Text = string.Format("{0:D2}", Eda);
                    item1_7.Text = item1_6.Text + "-" + item1_8.Text;
                }

                // フォルダチェック
                FolderPathCheck();
            }
            else
            {
                // 売上年度編集不可
                item1_3.Enabled = false;
                item3_1_5.Enabled = false;
            }

            //変更伝票以外
            if (mode != "change")
            {
                item3_1_1.Text = "";
                //item3_1_2.Checked = false;
                //item3_1_3.CustomFormat = " ";
                //item3_1_4.CustomFormat = " ";
                //えんとり君修正STEP2
                item3_1_27.Visible = false;
                label44.Visible = false;
                label50.Visible = false;
                if (item1_7.Text == "")
                {
                    if (item3_1_2.Checked)
                    {
                        label51.Visible = false;
                        txt_renamedfolder.Visible = false;
                        label115.Visible = false;
                    }
                    else
                    {
                        label51.Visible = true;
                        txt_renamedfolder.Visible = true;
                        label115.Visible = true;
                    }
                }
                else
                {
                    label51.Visible = false;
                    txt_renamedfolder.Visible = false;
                    label115.Visible = false;
                }
            }
            else
            {
                label584.Text = (Convert.ToInt32(AnkenData_H.Rows[0][52]) + 1).ToString();
                item3_1_4.Text = "";
                item3_1_4.CustomFormat = " "; // 741:赤黒作成時に起案日をクリア対応
                //えんとり君修正STEP2
                item3_1_27.Visible = true;
                label44.Visible = true;
                label50.Visible = true;
            }

            //新規登録後
            if (mode == "update")
            {
                set_error(GlobalMethod.GetMessage("I00004", ""));

                // 売上年度は編集不可
                item3_1_5.Enabled = false;

            }

            // 起案されてない
            if (!item3_1_2.Checked && mode != "change" && mode != "insert" && mode != "keikaku")
            {
                // 1073 赤黒の案件区分修正対応
                // 01:新規
                // 02:契約変更(赤伝)
                // 03:契約変更(黒伝)
                // 04:中止
                // 05:計画
                // 06:契約変更(黒伝・金額変更)
                // 07:契約変更(黒伝・工期変更)
                // 08:契約変更(黒伝・金額工期変更)
                // 09:契約変更(黒伝・その他)
                // 01以外の場合、変更伝票を行ったと判定
                if (AnkenData_K.Rows[0][0].ToString() != "01")
                {
                    item3_1_1.Enabled = true; // 案件区分
                }
                else
                {
                    item3_1_1.Enabled = false; // 案件区分
                }
            }

            //起案時(変更伝票・新規登録以外)
            if (item3_1_2.Checked && mode != "change" && mode != "insert" && mode != "keikaku")
            {
                //ヘッダーボタン
                //button1.Visible = false;
                //button10.Visible = false;
                button12.Visible = true;

                // 起案後は変更出来る項目が制限される

                //引合
                //引合状況
                //item1_1.Enabled = false;
                //基本情報
                item1_1.Enabled = false; // 引合状況
                item1_3.Enabled = false; // 売上年度
                item1_2_KoukiNendo.Enabled = false; // 工期開始年度
                item1_9.Enabled = false; // 登録日
                item1_10.Enabled = false; // 受託課所支部
                pictureBox2.Visible = false; // 契約担当者プロンプト
                //pictureBox16.Enabled = false; // 計画プロンプト

                //案件情報
                item1_13.ReadOnly = true; // 業務名称
                item1_14.Enabled = false; // 契約区分
                item1_15.Enabled = false; // 入力方式
                item1_17.Enabled = false; // 入札状況
                item1_18.ReadOnly = true; // 案件メモ
                //発注者情報
                pictureBox4.Visible = false; // 発注者コードプロンプト
                item1_23.ReadOnly = true; // 発注者名
                item1_24.ReadOnly = true; // 発注者名課名
                //当会対応
                item1_34.Enabled = false; // 参考見積
                item1_35.Enabled = false; // 受注意欲
                //item1_36.ReadOnly = true; // 参考見積額（税抜）・・・参考見積額は起案後も変更可能
                //業務内容
                //部門配分
                item1_7_1_1_1.ReadOnly = true; // 部門配分 引合・入札 配分率 調査部
                item1_7_1_2_1.ReadOnly = true; // 部門配分 引合・入札 配分率 事業普及部
                item1_7_1_3_1.ReadOnly = true; // 部門配分 引合・入札 配分率 情報システム部
                item1_7_1_4_1.ReadOnly = true; // 部門配分 引合・入札 配分率 総合研究所
                //業務配分
                item1_7_2_1_1.ReadOnly = true;
                item1_7_2_2_1.ReadOnly = true;
                item1_7_2_3_1.ReadOnly = true;
                item1_7_2_4_1.ReadOnly = true;
                item1_7_2_5_1.ReadOnly = true;
                item1_7_2_6_1.ReadOnly = true;
                item1_7_2_7_1.ReadOnly = true;
                item1_7_2_8_1.ReadOnly = true;
                item1_7_2_9_1.ReadOnly = true;
                item1_7_2_10_1.ReadOnly = true;
                item1_7_2_11_1.ReadOnly = true;
                item1_7_2_12_1.ReadOnly = true;

                //入札
                //入札状況
                item2_1_1.Enabled = false; // 入札状況
                //item2_1_3.Enabled = false; // 入札（予定）日・・・データ項目定義に記載がない
                item2_1_4.Enabled = false; // 引合状況
                item2_1_5.Enabled = false; // 案件区分
                item2_1_6.Enabled = false; // 案件メモ
                //当会対応
                item2_2_1.Enabled = false; // 当会応札
                item2_2_2.Enabled = false; // 参考見積・・・データ項目定義だと、空欄
                item2_2_3.Enabled = false; // 受注意欲
                item2_2_4.Enabled = false; // 参考見積額（税抜）・・・参考見積額は起案後も変更可能、だが、ここは表示のみ
                //入札参加者
                c1FlexGrid2.Cols[3].AllowEditing = false;
                //業務内容
                //部門配分
                item2_4_1_1_1.Enabled = false;
                item2_4_1_2_1.Enabled = false;
                item2_4_1_3_1.Enabled = false;
                item2_4_1_4_1.Enabled = false;
                //業務配分
                item2_4_2_1_1.Enabled = false;
                item2_4_2_2_1.Enabled = false;
                item2_4_2_3_1.Enabled = false;
                item2_4_2_4_1.Enabled = false;
                item2_4_2_5_1.Enabled = false;
                item2_4_2_6_1.Enabled = false;
                item2_4_2_7_1.Enabled = false;
                item2_4_2_8_1.Enabled = false;
                item2_4_2_9_1.Enabled = false;
                item2_4_2_10_1.Enabled = false;
                item2_4_2_11_1.Enabled = false;
                item2_4_2_12_1.Enabled = false;

                //契約
                // VIPS 20220221 課題管理表No.1271(965) DEL チェック用帳票出力・内容確認ボタンを起案後も使用可とする
                // チェック用帳票出力・内容確認ボタンは常時使用可とするためコメントアウト
                //button11.Enabled = false;
                //button11.BackColor = Color.DimGray;
                button14.Enabled = false;
                button14.BackColor = Color.DimGray;
                button15.Enabled = true;
                button15.BackColor = Color.FromArgb(42, 78, 122);
                button15.ForeColor = Color.White;
                button16.Enabled = true;
                button16.BackColor = Color.FromArgb(42, 78, 122);
                button16.ForeColor = Color.White;

                // Role:2システム管理者 以外の場合、起案解除は非表示
                if (UserInfos[4].Equals("2"))
                {
                    // 起案解除
                    button17.Enabled = true;
                    button17.BackColor = Color.FromArgb(42, 78, 122);
                    button17.ForeColor = Color.White;
                }
                else
                {
                    button17.Visible = false;
                }

                item3_1_20_akaden.Visible = false;
                item3_1_20_kuroden.Visible = false;
                //pictureBox6.Enabled = false; // 管理技術者プロンプト
                //pictureBox18.Enabled = false; // 管理技術者プロンプト×
                //pictureBox8.Enabled = false; // 照査技術者プロンプト
                //pictureBox19.Enabled = false; // 照査技術者プロンプト×
                //pictureBox9.Enabled = false; // 審査担当者プロンプト
                //pictureBox20.Enabled = false; // 審査担当者プロンプト×
                //pictureBox10.Enabled = false; // 業務担当者プロンプト
                //pictureBox21.Enabled = false; // 業者担当者プロンプト×
                //pictureBox11.Enabled = false; // 窓口担当者プロンプト
                //pictureBox22.Enabled = false; // 窓口担当者プロンプト×
                //c1FlexGrid3.Enabled = false;
                //契約情報
                item3_1_1.Enabled = false; // 案件区分

                //item3_1_3.Enabled = false; // 契約締結（変更）日 // 741:起案後の変更可に変更
                item3_1_4.Enabled = false; // 起案日
                //item3_1_5.Enabled = false; // 売上年度
                //item3_1_6.Enabled = false; // 契約工期自 // 741:起案後の変更可に変更
                item3_1_7.Enabled = false; // 契約工期至
                item3_1_8.Enabled = false; // 契約区分
                button2.Visible = false; // 工期末日付、及び、請求（1回目）に設定
                item3_1_10.Enabled = false; // 消費税率
                item3_1_12.Enabled = false; // 税抜（自動計算用）
                item3_1_13.Enabled = false; // 税込
                item3_1_14.Enabled = false; // 内消費税
                item3_1_15.Enabled = false; // 受託金額（税込）
                item3_1_16.Enabled = false; // 受託外金額（税込）
                item3_1_17.Enabled = false; // 変更・中止理由
                item3_1_18.Enabled = false; // 案件メモ
                item3_1_19.Enabled = false; // 備考
                item3_1_20.Enabled = false; // 契約書写
                item3_1_21.Enabled = false; // 特記仕様書
                item3_1_22.Enabled = false; // 見積書
                item3_1_23.Enabled = false; // 単品調査内訳書
                item3_1_24.Enabled = false; // その他
                item3_1_25.Enabled = false; // その他備考
                item3_1_26.Enabled = false; // 契約図書
                // えんとり君修正STEP2
                item3_ribc_price.Enabled = false;//RIBC用単価データ
                item3_sa_commpany.Enabled = false;//サ社経由
                item3_1_ribc.Enabled = false;//RIBC用単価契約書
                //item3_1_
                //配分情報
                item3_2_1_1.Enabled = false;
                item3_2_2_1.Enabled = false;
                item3_2_3_1.Enabled = false;
                item3_2_4_1.Enabled = false;
                item3_2_1_2.Enabled = false;
                item3_2_2_2.Enabled = false;
                item3_2_3_2.Enabled = false;
                item3_2_4_2.Enabled = false;

                // 20210511 起案後も入力できるよう修正
                //単契
                //item3_3_1.Enabled = false;
                //item3_3_2.Enabled = false;
                //item3_3_3.Enabled = false;
                //item3_3_4.Enabled = false;

                //管理者・技術者
                item3_4_1.Enabled = false;
                item3_4_2.Enabled = false;
                item3_4_3.Enabled = false;
                item3_4_4.Enabled = false;
                //label548.Visible = false;
                item3_4_5.Enabled = false;
                //売上計上
                label111.Visible = false;
                c1FlexGrid4.Cols[1].AllowEditing = false;
                c1FlexGrid4.Cols[3].AllowEditing = false;
                c1FlexGrid4.Cols[9].AllowEditing = false;
                c1FlexGrid4.Cols[11].AllowEditing = false;
                c1FlexGrid4.Cols[17].AllowEditing = false;
                c1FlexGrid4.Cols[19].AllowEditing = false;
                c1FlexGrid4.Cols[25].AllowEditing = false;
                c1FlexGrid4.Cols[27].AllowEditing = false;
                //請求書情報
                label586.Visible = false;
                //エントリ君修正STEP2
                label43.Visible = false;
                label326.Visible = false;

                item3_6_1.Enabled = false;
                item3_6_2.Enabled = false;
                item3_6_3.Enabled = false;
                item3_6_4.Enabled = false;
                item3_6_5.Enabled = false;
                item3_6_6.Enabled = false;
                item3_6_7.Enabled = false;
                item3_6_8.Enabled = false;
                item3_6_9.Enabled = false;
                item3_6_10.Enabled = false;
                item3_6_11.Enabled = false;
                item3_6_12.Enabled = false;

                // 調査部 業務別配分
                item3_7_2_14_1.Enabled = false;
                item3_7_2_15_1.Enabled = false;
                item3_7_2_16_1.Enabled = false;
                item3_7_2_17_1.Enabled = false;
                item3_7_2_18_1.Enabled = false;
                item3_7_2_19_1.Enabled = false;
                item3_7_2_20_1.Enabled = false;
                item3_7_2_21_1.Enabled = false;
                item3_7_2_22_1.Enabled = false;
                item3_7_2_23_1.Enabled = false;
                item3_7_2_24_1.Enabled = false;
                item3_7_2_25_1.Enabled = false;

                // 売上年度は編集不可
                item3_1_5.Enabled = false;
            }

            //受託番号が採番されていない場合、「この案件番号の枝番で受託番号を作成する」ボタンを非表示
            //if (item1_7.Text == "")
            if (item1_8.Text == "")
            {
                button10.Visible = false;
            }
            //「この案件番号の枝番で受託番号を作成する」ボタンは押下時に受託番号をチェックするように変更
            ////受託番号が採番されていない場合、「この案件番号の枝番で受託番号を作成する」ボタンを無効化
            //if (item1_8.Text == "")
            //{
            //    //button10.Visible = false;
            //    button10.BackColor = Color.DarkGray;
            //    button10.Enabled = false;
            //}

            // 起案 or 起案解除した場合に、契約タブに移動する
            if (KianKaijoFLG | KianFLG)
            {
                tab.SelectedIndex = 2;
                KianKaijoFLG = false;
                KianFLG = false;
            }

            c1FlexGrid1.Height = 4 + 22 * c1FlexGrid1.Rows.Count;
            c1FlexGrid2.Height = 4 + 22 * c1FlexGrid2.Rows.Count;
            c1FlexGrid3.Height = 4 + 22 * c1FlexGrid3.Rows.Count;
            c1FlexGrid4.Height = 4 + 22 * c1FlexGrid4.Rows.Count;
            c1FlexGrid5.Height = 4 + 22 * c1FlexGrid5.Rows.Count;


            //モード別処理
            if (mode == "insert" || mode == "keikaku")
            {
                //不要タブの非表示化
                this.tab.TabPages.Remove(this.tabPage3);
                this.tab.TabPages.Remove(this.tabPage4);
                this.tab.TabPages.Remove(this.tabPage6);
                //一部名称の変更
                button6.Text = "新規登録";
                label4.Text = "■エントリくん 新規追加";
                //一部ボタンの非表示化
                tableLayoutPanel3.Visible = false;
                button1.Visible = false;
                button10.Visible = false;
                button12.Visible = false;
                //えんとり君修正STEP2
                label51.Visible = false;
                txt_renamedfolder.Visible = false;
                label115.Visible = false;
                item1_37.Visible = false;
                item1_38.Visible = false;
                // No.1422 1196 案件番号の変更履歴を保存する
                item1_39.Visible = false;
                label82.Visible = false;
                label84.Visible = false;
                // No.1434 不要な項目名「案件番号変更履歴」を非表示にする
                label124.Visible = false;

                //不具合No1310(1028)
                //コピペテキストと反映するボタンが配置されたテールブルレイアウトパネルを表示有効化する
                tableLayoutPanel8.Visible = true;
                //案件番号などの表が乗ってる列をサイズゼロにする
                tableLayoutPanel2.ColumnStyles[4] = new ColumnStyle(SizeType.Absolute, 0.0F);
                //削除ボタン乗ってる列のサイズをゼロパーセントにする。
                tableLayoutPanel2.ColumnStyles[5] = new ColumnStyle(SizeType.Percent, 0);
                //else追加　新規・計画以外は反映用のテキストとボタンを消す
            }
            else
            {
                //新規・計画以外は反映用のテキストとボタンを消す
                txtTayoriData.Visible = false;
                btnHanei.Visible = false;
                tableLayoutPanel8.Visible = false;
                //コピペテキストと反映するボタンの配置された列幅をゼロにする
                tableLayoutPanel2.ColumnStyles[6] = new ColumnStyle(SizeType.Absolute, 0.0F);
            }

            if (mode == "change")
            {
                //不要タブの非表示化
                this.tab.TabPages.Remove(this.tabPage1);
                this.tab.TabPages.Remove(this.tabPage3);
                this.tab.TabPages.Remove(this.tabPage6);
                //一部名称の変更
                label4.Text = "■エントリくん 変更伝票";
                //一部ボタンの非表示化
                button1.Visible = false;
                button6.Visible = false;
                button10.Visible = false;
                //button12.Visible = true; 
                button12.Visible = false; // 削除ボタン
                tableLayoutPanel120.Visible = true;
                tableLayoutPanel98.Visible = false;
                tableLayoutPanel99.Visible = true;
                item3_1_1.Enabled = true;
                item3_1_11.ReadOnly = false;
                label324.BackColor = Color.FromArgb(252, 228, 214);

                item3_1_20_akaden.Visible = false;
                item3_1_20_kuroden.Visible = false;

                // 売上年度は編集不可
                item3_1_5.Enabled = false;
            }
            if (mode == "view")
            {
                //一部名称の変更
                label4.Text = "■エントリくん 参照";
                //一部ボタンの非表示化
                button1.Visible = false;
                button6.Visible = false;
                button10.Visible = false;
                button12.Visible = false;
                button2.Visible = false;
                tableLayoutPanel120.Visible = true;
                tableLayoutPanel97.Visible = false;
                tableLayoutPanel98.Visible = false;
                tableLayoutPanel99.Visible = true;
                item3_1_1.Enabled = true;
                item3_1_11.ReadOnly = false;
                label324.BackColor = Color.FromArgb(252, 228, 214);

                item3_1_20_akaden.Visible = false;
                item3_1_20_kuroden.Visible = false;
            }

            if (Message != "")
            {
                // 画面呼びなおし時にメッセージを表示
                set_error(Message);
                // メッセージをクリア
                Message = "";
            }
            //「この案件を元に新規登録」「この案件番号の枝番で受託番号を作成する」ボタンを一時凍結
            if (GlobalMethod.GetCommonValue1("COPYBUTTON_FLAG", "1") == "0")
            {
                button1.Visible = false;
            }
            if (GlobalMethod.GetCommonValue1("COPYBUTTON_FLAG", "2") == "0")
            {
                button10.Visible = false;
            }

            // えんとり君修正STEP2
            sFolderRenameBef = item1_12.Text;
            sFolderBushoCDRenameBef = UserInfos[2];    //受託課所支部（契約部所）
            if (item1_2_KoukiNendo.SelectedValue == null)
            {
                sFolderYearRenameBef = item1_2_KoukiNendo.Text.Substring(0, 4);
            }
            else
            {
                sFolderYearRenameBef = item1_2_KoukiNendo.SelectedValue.ToString();   // 工期開始年度
            }
            sFolderGyomuRenameBef = item1_13.Text;   // 業務名称
            sFolderOrderRenameBef = item1_23.Text;//発注者名

            sItem1_10_ori = item1_10.SelectedValue == null ? "" : item1_10.SelectedValue.ToString(); //受託課所支部（契約部所）DB値
            sItem1_2_KoukiNendo_ori = sFolderYearRenameBef; //工期開始年度DB値

            bool bVisible = UserInfos[4].Equals("2");
            label66.Visible = bVisible;
            label70.Visible = bVisible;
            item3_sa_commpany.Visible = bVisible;

            //GlobalMethod.outputLogger("entory_input", "entory_input load 終了 " + DateTime.Now, "GetAnkenJouhou", UserInfos[1]);
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void set_combo()
        {
            GlobalMethod GlobalMethod = new GlobalMethod();
            //コンボボックスの内容を設定
            var combodt1 = new System.Data.DataTable();
            var combodt2 = new System.Data.DataTable();
            var combodt3 = new System.Data.DataTable();
            var combodt4 = new System.Data.DataTable();
            var combodt5 = new System.Data.DataTable();
            var combodt6 = new System.Data.DataTable();
            var combodt7 = new System.Data.DataTable();
            var combodt8 = new System.Data.DataTable();
            var combodt9 = new System.Data.DataTable();

            /*
            //受託課所支部
            //SQL変数
            String discript = "Mst_Busho.ShibuMei + ' ' + ISNULL(Mst_Busho.KaMei,'')";
            String value = "Mst_Busho.GyoumuBushoCD";
            String table = "Mst_Busho";
            String where = "JutakubuBushoHyoujiFlg = 1 AND GyoumuBushoCD < '999990' AND BushoNewOld <= 1 " +
                            " AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' ";
            //コンボボックスデータ取得
            combodt1 = GlobalMethod.getData(discript, value, table, where);
            item1_10.DataSource = combodt1;
            item1_10.DisplayMember = "Discript";
            item1_10.ValueMember = "Value";
            */
            //事業部
            //SQL変数
            String discript = "JigyoubuMei";
            String value = "JigyoubuHeadCD";
            String table = "Mst_Jigyoubu";
            String where = "";
            //コンボボックスデータ取得
            //DataTable combodt = GetComboData.getData(discript, value, table, where);
            //src_5.Items.Clear();
            //src_5.Items.AddRange(combodt.AsEnumerable().Select(row => row["JigyoubuMei"].ToString()).ToArray<string>());

            //案件区分
            //SQL変数
            discript = "SakuseiKubun";
            value = "SakuseiKubunID";
            table = "Mst_SakuseiKubun";
            if (mode == "change")
            {
                where = " SakuseiKubunID >= '04' AND SakuseiKubunID <> '05' ";
            }
            else
            {
                // 案件情報IDが存在する（更新等）
                if (AnkenID != "" && mode != "insert" && mode != "keikaku")
                {
                    // 01以外の場合、変更伝票を行ったと判定
                    // 1073 赤黒の案件区分修正対応
                    // 01:新規
                    // 02:契約変更(赤伝)
                    // 03:契約変更(黒伝)
                    // 04:中止
                    // 05:計画
                    // 06:契約変更(黒伝・金額変更)
                    // 07:契約変更(黒伝・工期変更)
                    // 08:契約変更(黒伝・金額工期変更)
                    // 09:契約変更(黒伝・その他)
                    DataTable dt = GlobalMethod.getData("AnkenSakuseiKubun", "AnkenSakuseiKubun", "AnkenJouhou", "AnkenJouhouID = " + AnkenID);
                    if (dt != null && dt.Rows[0][0].ToString() != "01" && dt.Rows[0][0].ToString() != "02")
                    {
                        where = " SakuseiKubunID >= '04' AND SakuseiKubunID <> '05'";
                    }
                    else if (dt != null && dt.Rows[0][0].ToString() == "02")
                    {
                        // 02:契約変更(赤伝)が表示されない対応
                        where = " SakuseiKubunID = '02'";
                        // 編集不可に
                        item3_1_1.Enabled = false;
                    }
                    else
                    {
                        where = "";
                    }
                }
                else
                {
                    where = "";
                }
            }
            //コンボボックスデータ取得
            combodt2 = GlobalMethod.getData(discript, value, table, where);
            item1_2.DataSource = combodt2;
            item1_2.DisplayMember = "Discript";
            item1_2.ValueMember = "Value";
            item2_1_5.DataSource = combodt2;
            item2_1_5.DisplayMember = "Discript";
            item2_1_5.ValueMember = "Value";
            item3_1_1.DataSource = combodt2;
            item3_1_1.DisplayMember = "Discript";
            item3_1_1.ValueMember = "Value";

            //契約区分
            //SQL変数
            discript = "GyoumuKubunHyouji";
            value = "GyoumuNarabijunCD";
            table = "Mst_GyoumuKubun";
            where = "GyoumuNarabijunCD < 100 ";
            //コンボボックスデータ取得
            combodt3 = GlobalMethod.getData(discript, value, table, where);
            //えんとり君修正STEP2
            DataTable combodt3C = new DataTable();
            if (combodt3 != null)
            {
                combodt3C = combodt3.Copy();
                DataRow dr = combodt3.NewRow();
                combodt3.Rows.InsertAt(dr, 0);
            }
            else
            {
                combodt3C = null;
            }
            item1_14.DataSource = combodt3;
            item1_14.DisplayMember = "Discript";
            item1_14.ValueMember = "Value";
            item3_1_8.DataSource = combodt3C;
            item3_1_8.DisplayMember = "Discript";
            item3_1_8.ValueMember = "Value";

            //入札方式
            //SQL変数
            discript = "KeiyakuKeitai";
            value = "KeiyakuKeitaiCD";
            table = "Mst_KeiyakuKeitai";
            where = "KeiyakuKeitaiNarabijun < 20";
            //コンボボックスデータ取得
            combodt4 = GlobalMethod.getData(discript, value, table, where);
            //えんとり君修正STEP2
            if (combodt4 != null)
            {
                DataRow dr = combodt4.NewRow();
                combodt4.Rows.InsertAt(dr, 0);
            }
            item1_15.DataSource = combodt4;
            item1_15.DisplayMember = "Discript";
            item1_15.ValueMember = "Value";

            //入札状況
            //SQL変数
            discript = "RakusatsuShaMei";
            value = "RakusatsuShaID";
            table = "Mst_RakusatsuSha";
            where = "RakusatsuShaNarabijun > 0";
            //コンボボックスデータ取得
            combodt5 = GlobalMethod.getData(discript, value, table, where);
            item1_17.DataSource = combodt5;
            item1_17.DisplayMember = "Discript";
            item1_17.ValueMember = "Value";
            item2_1_1.DataSource = combodt5;
            item2_1_1.DisplayMember = "Discript";
            item2_1_1.ValueMember = "Value";
            //グリッドのコンボボックス用リスト
            SortedList sl = new SortedList();
            //行の数だけの数だけSortedListにIDとSakuseiKubunをadd
            sl = GlobalMethod.Get_SortedList(combodt5);
            //該当グリッドのセルにセット
            c1FlexGrid1.Cols[7].DataMap = sl;

            //引合状況

            System.Data.DataTable tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "未確定");
            tmpdt.Rows.Add(2, "発注確定");
            tmpdt.Rows.Add(3, "発注無し");

            item1_1.DataSource = tmpdt;
            item1_1.DisplayMember = "Discript";
            item1_1.ValueMember = "Value";
            item2_1_4.DataSource = tmpdt;
            item2_1_4.DisplayMember = "Discript";
            item2_1_4.ValueMember = "Value";


            //参考見積
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "未提出");
            tmpdt.Rows.Add(2, "提出");
            tmpdt.Rows.Add(3, "依頼無し");
            tmpdt.Rows.Add(4, "辞退");

            item1_34.DataSource = tmpdt;
            item1_34.DisplayMember = "Discript";
            item1_34.ValueMember = "Value";
            item2_2_2.DataSource = tmpdt;
            item2_2_2.DisplayMember = "Discript";
            item2_2_2.ValueMember = "Value";

            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "対応前");
            tmpdt.Rows.Add(2, "応札");
            tmpdt.Rows.Add(3, "不参加");
            tmpdt.Rows.Add(4, "辞退");
            item2_2_1.DataSource = tmpdt;
            item2_2_1.DisplayMember = "Discript";
            item2_2_1.ValueMember = "Value";

            //tmpdt = new System.Data.DataTable();
            //tmpdt.Columns.Add("Value", typeof(int));
            //tmpdt.Columns.Add("Discript", typeof(string));
            //tmpdt.Rows.Add(1, "未提出");
            //tmpdt.Rows.Add(2, "提出");
            //tmpdt.Rows.Add(3, "依頼無し");
            //tmpdt.Rows.Add(4, "辞退");
            //item2_2_2.DataSource = tmpdt;
            //item2_2_2.DisplayMember = "Discript";
            //item2_2_2.ValueMember = "Value";


            //受注意欲
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "フラット");
            tmpdt.Rows.Add(2, "あり");
            tmpdt.Rows.Add(3, "なし");

            item1_35.DataSource = tmpdt;
            item1_35.DisplayMember = "Discript";
            item1_35.ValueMember = "Value";
            item2_2_3.DataSource = tmpdt;
            item2_2_3.DisplayMember = "Discript";
            item2_2_3.ValueMember = "Value";

            //売上年度
            discript = "NendoSeireki";
            value = "NendoID";
            table = "Mst_Nendo";
            where = "";
            //コンボボックスデータ取得
            combodt9 = GlobalMethod.getData(discript, value, table, where);
            item1_3.DataSource = combodt9;
            item1_3.DisplayMember = "Discript";
            item1_3.ValueMember = "Value";
            item1_2_KoukiNendo.DataSource = combodt9;
            item1_2_KoukiNendo.DisplayMember = "Discript";
            item1_2_KoukiNendo.ValueMember = "Value";
            item3_1_5.DataSource = combodt9;
            item3_1_5.DisplayMember = "Discript";
            item3_1_5.ValueMember = "Value";

            DataTable koukiNendoCombodt = GlobalMethod.getData(discript, value, table, where);

            item1_2_KoukiNendo.DataSource = koukiNendoCombodt;
            item1_2_KoukiNendo.DisplayMember = "Discript";
            item1_2_KoukiNendo.ValueMember = "Value";

            //落札者状況・落札額状況
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "判明");
            tmpdt.Rows.Add(2, "不明");
            tmpdt.Rows.Add(3, "推定");
            item2_3_1.DataSource = tmpdt;
            item2_3_1.DisplayMember = "Discript";
            item2_3_1.ValueMember = "Value";
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "判明");
            tmpdt.Rows.Add(2, "不明");
            tmpdt.Rows.Add(3, "推定");
            item2_3_2.DataSource = tmpdt;
            item2_3_2.DisplayMember = "Discript";
            item2_3_2.ValueMember = "Value";
        }

        private void set_combo_shibu(string nendo)
        {
            //受託課所支部
            string SelectedValue = "";
            if (item1_10.Text != "")
            {
                SelectedValue = item1_10.SelectedValue.ToString();
            }
            //SQL変数
            string discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            string where = "GyoumuBushoCD < '999990' AND BushoNewOld <= 1 AND BushoEntryHyoujiFlg = 1 AND ISNULL(BushoDeleteFlag,0) = 0 " +
                    //"AND NOT GyoumuBushoCD LIKE '1502%' AND NOT GyoumuBushoCD LIKE '1504%' AND NOT GyoumuBushoCD LIKE '121%' " +
                    "AND NOT GyoumuBushoCD LIKE '121%' " +

                    "AND ISNULL(KashoShibuCD,'') <> ''  ";
            int FromNendo;
            if (int.TryParse(nendo, out FromNendo))
            {
                int ToNendo = int.Parse(nendo) + 1;
                //where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + FromNendo + "/4/1' ) " +
                //"AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";
                where += "AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                "AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + FromNendo + "/4/1' ) ";

                // 工期開始年度が2021年度未満の場合、旧積シス（127910）をコンボに追加する
                if (FromNendo < 2021)
                {
                    // イレギュラー対応の為、以下の条件を付与
                    where += " OR (GyoumuBushoCD = '127910') ";
                }
            }

            Console.WriteLine(where);
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            item1_10.DataSource = combodt;
            item1_10.DisplayMember = "Discript";
            item1_10.ValueMember = "Value";
            if (SelectedValue != "")
            {
                item1_10.SelectedValue = SelectedValue;
            }
        }

        // 過去案件情報の行追加
        private void button7_Click(object sender, EventArgs e)
        {
            if (c1FlexGrid1.Rows.Count < 6)
            {
                // 前回受託番号ID
                string AnkenZenkaiRakusatsuID = "1";
                int num = 0;
                int maxNum = 0;

                // ヘッダーを除いて回し、前回受託番号IDの最大値を取得する
                for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                {
                    AnkenZenkaiRakusatsuID = c1FlexGrid1.Rows[i][16].ToString();
                    if (int.TryParse(AnkenZenkaiRakusatsuID, out num))
                    {
                        if (maxNum < num)
                        {
                            maxNum = num;
                        }
                    }
                }
                // 最大値 + 1
                maxNum = maxNum + 1;

                c1FlexGrid1.Rows.Add();
                // 追加した行にセット
                c1FlexGrid1.Rows[c1FlexGrid1.Rows.Count - 1][16] = maxNum;
            }
            Resize_Grid("c1FlexGrid1");
        }


        private void dateTimePicker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Delete)
            {
                ((DateTimePicker)sender).Text = "";
                ((DateTimePicker)sender).CustomFormat = " ";
            }
        }
        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";


        }

        private void textBox_Enter(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            tmp = tmp.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty);
            if (tmp == "")
            {
                tmp = "0";
            }
            ((System.Windows.Forms.TextBox)sender).Text = tmp;
        }

        private void textBox_Validated(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            ((System.Windows.Forms.TextBox)sender).Text = GetMoneyTextLong(GetLong(tmp));
        }
        // 応札数
        private void ousatsusuu_Validated(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            ((System.Windows.Forms.TextBox)sender).Text = GetInt(tmp).ToString();
        }
        private void textBox_ValidatedPercent(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            ((System.Windows.Forms.TextBox)sender).Text = GetPercentText(GetDouble(tmp));
        }

        // 消費税率で仕様するFormat
        private void textBox_NumericValidated(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            tmp = tmp.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty);
            if (tmp == "")
            {
                tmp = "0";
            }
            int num = 0;
            if (Int32.TryParse(tmp, out num))
            {
                ((System.Windows.Forms.TextBox)sender).Text = Int32.Parse(tmp).ToString();
            }
        }

        private void textbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        private void textbox_KeyPressPercent(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (c1FlexGrid2.Rows.Count < 11)
            {
                c1FlexGrid2.AllowAddNew = true;
                c1FlexGrid2.Rows.Add();
                Resize_Grid("c1FlexGrid2");
                c1FlexGrid2.AllowAddNew = false;
            }
        }

        // 落札者の入力方式を更新する
        private void button9_Click(object sender, EventArgs e)
        {
            get_guidance();
        }



        private void get_guidance()
        {
            string WHERE = " RakushijiHikiaijhokyo COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + item2_1_4.SelectedValue.ToString() + "%' " +
                            " AND RakushijiNyusatsuJhokyo COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + item2_1_1.SelectedValue.ToString() + "%' " +
                            " AND RakushijiKeiyakuKubun COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + item1_14.SelectedValue.ToString() + "%' " +
                            " AND RakushijiSankouMitsumori COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + item2_2_2.SelectedValue.ToString() + "%' " +
                            " AND RakushijiToukaiOusatu COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + item2_2_1.SelectedValue.ToString() + "%' " +
                            " AND RakushijiRakusatsuShaJokyou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + item2_3_1.SelectedValue.ToString() + "%' " +
                            " AND RakushijiRakusatsuGakuJokyou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + item2_3_2.SelectedValue.ToString() + "%' ";



            DataTable dt = GlobalMethod.getData("RakushijiOusatsuKingaku,RakushijiComment", "RakushijiOusatsusha", "Mst_RakusatsushaNyuryokushiji", WHERE);
            if (dt != null && dt.Rows.Count > 0)
            {
                item2_3_9.Text = dt.Rows[0][0].ToString();
                item2_3_10.Text = dt.Rows[0][1].ToString();
                item2_3_11.Text = dt.Rows[0][2].ToString();
            }
            else
            {
                item2_3_9.Text = "";
                item2_3_10.Text = "";
                item2_3_11.Text = "";
            }
        }

        private void c1FlexGrid2_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid2.HitTest(new System.Drawing.Point(e.X, e.Y));
            if (hti.Column == 5 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;
                Popup_Kyougou form = new Popup_Kyougou();
                form.ShowDialog();
                if (form.ReturnValue != null && form.ReturnValue[0] != null)
                {
                    c1FlexGrid2.Rows[_row][_col] = form.ReturnValue[1];
                    c1FlexGrid2.Rows[_row][_col - 1] = form.ReturnValue[2];
                    c1FlexGrid2.Rows[_row][8] = form.ReturnValue[0];
                    if (c1FlexGrid2.Rows[_row][6] == null || c1FlexGrid2.Rows[_row][6].ToString() == "")
                    {
                        c1FlexGrid2.Rows[_row][6] = 0;
                    }
                    if (c1FlexGrid2.Rows[_row][7] == null || c1FlexGrid2.Rows[_row][7].ToString() == "")
                    {
                        c1FlexGrid2.Rows[_row][7] = "";
                    }
                    if (c1FlexGrid2.GetCellCheck(_row, 3) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                    {
                        item2_3_7.Text = c1FlexGrid2.Rows[_row][5].ToString();
                        item2_3_8.Text = GetMoneyTextLong(GetLong(c1FlexGrid2.Rows[_row][6].ToString()));
                    }
                    int nyusatsuCnt = 0;
                    for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
                    {
                        if (c1FlexGrid2.Rows[i][5] != null && c1FlexGrid2.Rows[i][5].ToString() != "")
                        {
                            nyusatsuCnt++;
                        }
                    }
                    item2_3_6.Text = nyusatsuCnt.ToString();
                }
            }
            if (hti.Column == 1 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;
                if (MessageBox.Show(GlobalMethod.GetMessage("I10002", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    if (c1FlexGrid2.GetCellCheck(_row, 3) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                    {
                        item2_3_7.Text = "";
                        item2_3_8.Text = GetMoneyTextLong(0);
                    }

                    c1FlexGrid2.RemoveItem(_row);
                    Resize_Grid("c1FlexGrid2");


                    int nyusatsuCnt = 0;
                    for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
                    {
                        if (c1FlexGrid2.Rows[i][5] != null && c1FlexGrid2.Rows[i][5].ToString() != "")
                        {
                            nyusatsuCnt++;
                        }
                    }
                    item2_3_6.Text = nyusatsuCnt.ToString();
                }
            }
        }

        private void ChousainList_Click(object sender, EventArgs e)
        {

        }

        // 過去案件情報 の 案件選択プロンプト
        private void c1FlexGrid1_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.c1FlexGrid1.HitTest(new System.Drawing.Point(e.X, e.Y));

            //if (hti.Column == 3 & hti.Row != 0)
            if (hti.Column == 3 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                Popup_Anken form = new Popup_Anken();
                form.mode = "";
                int nendo = DateTime.Today.Year;
                //if (int.TryParse(item1_3.SelectedValue.ToString(), out nendo))
                if (int.TryParse(item1_2_KoukiNendo.SelectedValue.ToString(), out nendo))
                {
                    nendo--;
                }
                form.nendo = nendo.ToString();
                form.hachuushaKaMei = item1_23.Text.Trim() + "　" + item1_24.Text.Trim();
                form.gyoumuMei = item1_13.Text.Trim();
                form.gyoumuBushoCD = UserInfos[2];
                form.ShowDialog();
                if (form.ReturnValue != null && form.ReturnValue[0] != null)
                {
                    c1FlexGrid1.Rows[_row][2] = form.ReturnValue[0];   // AnkenJouhou.AnkenJouhouID
                    c1FlexGrid1.Rows[_row][3] = form.ReturnValue[1];   // AnkenAnkenBangou
                    c1FlexGrid1.Rows[_row][4] = form.ReturnValue[2];   // AnkenJutakuBangouALL
                    c1FlexGrid1.Rows[_row][5] = form.ReturnValue[3];   // AnkenJutakuBangouEda
                    c1FlexGrid1.Rows[_row][6] = form.ReturnValue[4];   // AnkenGyoumuMei
                    c1FlexGrid1.Rows[_row][7] = form.ReturnValue[5];   // NyuusatsuRakusatsusha
                    c1FlexGrid1.Rows[_row][8] = form.ReturnValue[6];   // NyuusatsuRakusatsushaID
                    c1FlexGrid1.Rows[_row][9] = form.ReturnValue[7];   // NyuusatsuRakusatugaku
                    c1FlexGrid1.Rows[_row][10] = form.ReturnValue[8];  // NyuusatsuOusatugaku
                    c1FlexGrid1.Rows[_row][11] = form.ReturnValue[9];  // NyuusatsuMitsumorigaku
                    c1FlexGrid1.Rows[_row][12] = form.ReturnValue[10]; // KeiyakuZeikomiKingaku
                    c1FlexGrid1.Rows[_row][13] = form.ReturnValue[11]; // Keiyakukeiyakukingakukei // 前回受託金額（税抜）
                    c1FlexGrid1.Rows[_row][14] = form.ReturnValue[12]; // NyuusatsuKyougouTashaID
                    c1FlexGrid1.Rows[_row][15] = form.ReturnValue[13]; // KyougouKigyouCD
                }
            }
            if (hti.Column == 1 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                if (MessageBox.Show("行を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    c1FlexGrid1.RemoveItem(_row);
                    Resize_Grid("c1FlexGrid1");
                }
            }
        }

        // 引合タブの2.案件情報の計画番号プロンプト
        private void pictureBox16_Click(object sender, EventArgs e)
        {
            Popup_Keikaku form = new Popup_Keikaku();
            form.gyoumuBushoCD = item1_10.SelectedValue.ToString();
            form.nendo = item1_2_KoukiNendo.SelectedValue.ToString();
            form.ShowDialog();

            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item1_4.Text = form.ReturnValue[0];
                item1_5.Text = form.ReturnValue[1];
                //// 受託課所支部
                //item1_10.SelectedValue = form.ReturnValue[2];
            }
            item1_4.Focus();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //form.nendo = item1_3.SelectedValue.ToString();
            form.nendo = item1_2_KoukiNendo.SelectedValue.ToString();
            // 受託課所支部が空じゃない場合、工期開始年度を変更時にコンボから値が無くなった場合にエラーとならないように
            if (item1_10.SelectedValue != null)
            {
                form.Busho = item1_10.SelectedValue.ToString();
            }

            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item1_11_CD.Text = form.ReturnValue[0];
                item1_11.Text = form.ReturnValue[1];
                item1_11_Busho.Text = form.ReturnValue[2];
                item1_10.SelectedValue = form.ReturnValue[2];
            }
            item1_11.Focus();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            Popup_Yubin form = new Popup_Yubin();
            form.Yubin = item1_30.Text;
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item1_30.Text = form.ReturnValue[0];
                item1_31.Text = form.ReturnValue[1];
            }
            item1_30.Focus();
        }

        // チェック用帳票出力・内容確認
        private void button11_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                // VIPS 20220221 課題管理表No.1271(965) ADD チェック用帳票出力・内容確認ボタンを起案後も使用可とする
                //起案済みの場合は帳票出力のみ実行する
                if (item3_1_2.Checked && mode != "change" && mode != "insert" && mode != "keikaku")
                {
                    ErrorMessage.Text = "";
                    // えんとり君修正STEP2
                    //string[] result = GlobalMethod.InsertReportWork(2, UserInfos[0], new string[] { AnkenID, Header1.Text, "1", "0" });
                    if (!ErrorFLG(2))
                    {
                        Execute_SQL(2);
                        int ListID = 2;
                        if (item3_1_1.Text != null && item3_1_1.Text != "" && (item3_1_1.SelectedValue.ToString() == "03" || int.Parse(item3_1_1.SelectedValue.ToString()) > 5))
                        {
                            // 01:新規
                            // 02:契約変更(赤伝)
                            // 03:契約変更(黒伝)
                            // 04:中止
                            // 05:計画
                            // 06:契約変更(黒伝・金額変更)
                            // 07:契約変更(黒伝・工期変更)
                            // 08:契約変更(黒伝・金額工期変更)
                            // 09:契約変更(黒伝・その他)
                            ListID = 353;
                        }
                        string[] result = GlobalMethod.InsertReportWork(ListID, UserInfos[0], new string[] { AnkenID, Header1.Text, "1", "0" });
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
                                Popup_Download form = new Popup_Download();
                                form.TopLevel = false;
                                this.Controls.Add(form);
                                form.ExcelName = Path.GetFileName(result[3]);
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
                else
                {
                    //if (ErrorFLG(1) && KianError(0))
                    if (!ErrorFLG(2))
                    {
                        Execute_SQL(2);
                        //string[] result = GlobalMethod.InsertReportWork(2, UserInfos[0], new string[] { AnkenID, Header1.Text, "1", "0" });
                        int ListID = 2;
                        if (item3_1_1.Text != null && item3_1_1.Text != "" && (item3_1_1.SelectedValue.ToString() == "03" || int.Parse(item3_1_1.SelectedValue.ToString()) > 5))
                        {
                            // 01:新規
                            // 02:契約変更(赤伝)
                            // 03:契約変更(黒伝)
                            // 04:中止
                            // 05:計画
                            // 06:契約変更(黒伝・金額変更)
                            // 07:契約変更(黒伝・工期変更)
                            // 08:契約変更(黒伝・金額工期変更)
                            // 09:契約変更(黒伝・その他)
                            ListID = 353;
                        }
                        string[] result = GlobalMethod.InsertReportWork(ListID, UserInfos[0], new string[] { AnkenID, Header1.Text, "1", "0" });
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
                                Popup_Download form = new Popup_Download();
                                form.TopLevel = false;
                                this.Controls.Add(form);
                                form.ExcelName = Path.GetFileName(result[3]);
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
            }
        }

        // えんとり君修正STEP2　変更伝票画面から確認用エントリーチェックシートを出力出来ます。　「赤伝作成・出力」ボタンを流用する
        private void button20_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                ErrorMessage.Text = "";
                if (String.IsNullOrEmpty(item3_1_1.Text))
                {
                    set_error("案件区分を選択してください。");
                    return;
                }
                KianError(1);
                // エラーでも確認シートを出力する    
                string sKubun = item3_1_1.SelectedValue.ToString();
                int ListID = 2;
                if ((sKubun == "03" || int.Parse(sKubun) > 5))
                {
                    // 01:新規
                    // 02:契約変更(赤伝)
                    // 03:契約変更(黒伝)
                    // 04:中止
                    // 05:計画
                    // 06:契約変更(黒伝・金額変更)
                    // 07:契約変更(黒伝・工期変更)
                    // 08:契約変更(黒伝・金額工期変更)
                    // 09:契約変更(黒伝・その他)
                    ListID = 353;
                }
                int ankenJouhouID = Create_DummyData();
                if (ankenJouhouID > 0)
                {
                    string[] result = GlobalMethod.InsertReportWork(ListID, UserInfos[0], new string[] { ankenJouhouID.ToString(), Header1.Text, "1", "0" });
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
                            Popup_Download form = new Popup_Download();
                            form.TopLevel = false;
                            this.Controls.Add(form);
                            form.ExcelName = Path.GetFileName(result[3]);
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
                else
                {
                    // エラーが発生しました
                    set_error("出力用データを作成する時にエラーが発生しました。");
                }
            }
        }

        /*
        // 赤伝作成・出力
        private void button20_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string[] result = GlobalMethod.InsertReportWork(1, UserInfos[0], new string[] { item3_1_20_akaden.Text, Header1.Text, "0", "2" });

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
                        Popup_Download form = new Popup_Download();
                        form.TopLevel = false;
                        this.Controls.Add(form);

                        String fileName = Path.GetFileName(result[3]);
                        // VIPS 20220303 課題管理表No.1262(955) DEL 帳票出力EXEで出力ファイル名を設定しているため削除
                        //fileName = fileName.Replace(".xlsx", "(赤伝).xlsx");
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
        */

        // 黒伝・中止伝票作成・出力
        private void button21_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                //えんとり君修正STEP2 STT
                int ListID = 1;
                if (item3_1_1.Text != null && item3_1_1.Text != "" && (item3_1_1.SelectedValue.ToString() == "03" || int.Parse(item3_1_1.SelectedValue.ToString()) > 5))
                {
                    // 01:新規
                    // 02:契約変更(赤伝)
                    // 03:契約変更(黒伝)
                    // 04:中止
                    // 05:計画
                    // 06:契約変更(黒伝・金額変更)
                    // 07:契約変更(黒伝・工期変更)
                    // 08:契約変更(黒伝・金額工期変更)
                    // 09:契約変更(黒伝・その他)
                    ListID = 352;
                }
                //えんとり君修正STEP2 END
                // 帳票に渡すAnkenJouhouID
                string ankenJouhouID = "";
                
                if (item3_1_1.Text != null && item3_1_1.Text != "" && item3_1_1.SelectedValue.ToString() == "04")
                {
                    // 案件区分が04：中止の場合、帳票プログラムには赤伝のAnkenJouhouIDを渡す
                    ankenJouhouID = item3_1_20_akaden.Text;
                }
                else
                {
                    // 黒伝のAnkenJouhouID
                    ankenJouhouID = item3_1_20_kuroden.Text;
                }

                //えんとり君修正STEP2
                //string[] result = GlobalMethod.InsertReportWork(1, UserInfos[0], new string[] { ankenJouhouID, Header1.Text, "0", "1" });
                string[] result = GlobalMethod.InsertReportWork(ListID, UserInfos[0], new string[] { ankenJouhouID, Header1.Text, "0", "1" });

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
                        Popup_Download form = new Popup_Download();
                        form.TopLevel = false;
                        this.Controls.Add(form);
                        String fileName = Path.GetFileName(result[3]);
                        // VIPS 20220303 課題管理表No.1262(955) DEL 帳票出力EXEで出力ファイル名を設定しているため削除
                        //fileName = fileName.Replace(".xlsx", "(黒伝・中止).xlsx");
                        form.ExcelName = fileName;
                        //form.ExcelName = Path.GetFileName(result[3]) + "(黒伝・中止)";
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

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Popup_Hachusha form = new Popup_Hachusha();
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item1_19.Text = form.ReturnValue[0];
                item1_20.Text = form.ReturnValue[1];
                item1_21.Text = form.ReturnValue[2];
                item1_22.Text = form.ReturnValue[3];
                item1_23.Text = form.ReturnValue[4];
            }
            item1_19.Focus();
        }

        // 案件（受託）フォルダアイコン押下
        private void pictureBox3_Click(object sender, EventArgs e)
        {
            if (item1_12.Text == "")
            {
                // No.300対応 エクスプローラーを表示
                //FolderBrowserDialog fdb = new FolderBrowserDialog();
                //fdb.Description = "フォルダを選択してください。";
                ////fdb.SelectedPath = @"\\PC34-013\00Cyousa\00調査 情報部門共有";
                //string SelectedPath = "";
                //GlobalMethod GM = new GlobalMethod();

                //// M_CommonMasterからFolderPathを取得 例）//HPNJPH8451DXGA/00Cyousa/00調査情報部門共有/$NENDO$/200受託調査関連
                //string folderBase = GM.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", item1_3.SelectedValue.ToString());

                //string folderPath = "";
                //string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

                //using (var conn = new SqlConnection(connStr))
                //{
                //    conn.Open();
                //    var cmd = conn.CreateCommand();
                //    var dt = new System.Data.DataTable();
                //    //SQL生成
                //    cmd.CommandText = "SELECT " +
                //      "FolderPath " +
                //      "FROM " + "M_Folder " +
                //      "WHERE MENU_ID = '100' " +
                //      "and FolderBushoCD = '" + item1_10.SelectedValue + "' " +
                //      "and FolderBunruiCD = '1'";

                //    //データ取得
                //    var sda = new SqlDataAdapter(cmd);
                //    sda.Fill(dt);

                //    if (dt.Rows.Count > 0)
                //    {
                //        folderPath = dt.Rows[0][0].ToString();
                //    }
                //}

                //if (UserInfos[4] == "35")
                //{
                //    //SelectedPath = (@"\\00Cyousa\00調査 情報部門共有\派遣者用フォルダ");
                //    folderPath = folderPath.Replace(@"$FOLDER_BASE$", folderBase);
                //    SelectedPath = folderPath + @"\派遣者用フォルダ";
                //}
                //else
                //{
                //    //SelectedPath = (@"\\00Cyousa\00調査 情報部門共有");
                //    folderPath = folderPath.Replace(@"$FOLDER_BASE$", folderBase);
                //}
                ////MessageBox.Show(GM.GetPathValid(SelectedPath));
                //folderPath = @folderPath.Replace("/", @"\");
                //fdb.SelectedPath = folderPath;

                //if (fdb.ShowDialog(this) == DialogResult.OK)
                //{
                //    item1_12.Text = fdb.SelectedPath;
                //}
                System.Diagnostics.Process.Start("EXPLORER.EXE", "");
            }
            else
            {
                // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
                if (System.Text.RegularExpressions.Regex.IsMatch(item1_12.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // 指定されたフォルダパスが存在するなら開く
                    if (item1_12.Text != "" && item1_12.Text != null && Directory.Exists(item1_12.Text))
                    {
                        System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item1_12.Text));
                    }
                    else
                    {
                        // No.300対応
                        //// 存在しないパスを指定された場合
                        //FolderBrowserDialog fdb = new FolderBrowserDialog();
                        //if (fdb.ShowDialog(this) == DialogResult.OK)
                        //{
                        //    item1_12.Text = fdb.SelectedPath;
                        //}
                        System.Diagnostics.Process.Start("EXPLORER.EXE", "");
                    }
                }
                else
                {
                    // No.300対応
                    // 書かれているファイルパスとして認識できない場合のみ、エクスプローラーで表示する
                    //FolderBrowserDialog fdb = new FolderBrowserDialog();
                    //if (fdb.ShowDialog(this) == DialogResult.OK)
                    //{
                    //    item1_12.Text = fdb.SelectedPath;
                    //}
                    System.Diagnostics.Process.Start("EXPLORER.EXE", "");
                }
            }
        }

        //private Boolean ErrorFLG(int flg)
        //{
        //    //0:新規登録　1:更新処理 2:チェック用帳票前の更新　3:起案前の更新
        //    // エラーフラグ true：エラー、false：正常
        //    Boolean ErrorFLG = false;
        //    //必須チェック デフォルト　requiredFlag:true /　falseだとエラー
        //    Boolean requiredFlag = true;
        //    // データチェックフラグ true：正常、false：エラー
        //    Boolean varidateFlag = true;

        //    // 入札エラーON/OFFフラグ true：ON（警告）、false：OFF（無警告）
        //    Boolean errorNyuusatsuFlag = false;
        //    // 契約エラーON/OFFフラグ true：ON（警告）、false：OFF（無警告）
        //    Boolean errorKeiyakuFlag = false;

        //    //入札・契約チェック デフォルト　true:警告 /　false:警告なし
        //    Boolean requiredNyuusatsuKeiyakuFlag = false;

        //    //新規登録チェック　　true:エラー終了 /　false:警告なし
        //    Boolean InsertErrorFlag = false;

        //    // 起案の場合は、エラーチェックを行う
        //    if (flg >= 3)
        //    {
        //        errorNyuusatsuFlag = true;
        //        errorKeiyakuFlag = true;
        //    }
        //    else
        //    {
        //        // ①入札タブ：入札状況が入札成立となったら入札のチェックを行う（更新可能）
        //        // item2_1_1:入札タブの1.入札状況の入札状況
        //        // NYUUSATSU_SEIRITSU:2(入札成立)
        //        if (item2_1_1.SelectedValue == GlobalMethod.GetCommonValue2("NYUUSATSU_SEIRITSU"))
        //        {
        //            errorNyuusatsuFlag = true;
        //        }
        //        // ②契約タブ：調査会様での入札が成立したら契約のチェック処理を行う（更新可能）
        //        // item2_3_7:入札タブの3.入札結果の落札者
        //        // ENTORY_TOUKAI:建設物価調査会
        //        if (item2_3_7.Text == GlobalMethod.GetCommonValue2("ENTORY_TOUKAI"))
        //        {
        //            errorKeiyakuFlag = true;
        //        }
        //    }

        //    set_error("", 0);
        //    //0:新規登録 1:更新 2:チェック用出力(赤伝・黒伝) 3:起案

        //    //必須チェックを行う項目の背景を白に戻す。
        //    item1_3.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_9.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_10.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_11.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_12.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_13.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_14.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_16.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_19.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_20.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_21.BackColor = Color.FromArgb(255, 255, 255);
        //    item1_36.BackColor = Color.FromArgb(255, 255, 255);
        //    item2_2_1.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_3.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_4.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_6.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_3.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_4.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_6.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_7.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_13.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_15.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_16.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_4_1.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_4_4.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_4_5.BackColor = Color.FromArgb(255, 255, 255);
        //    if (c1FlexGrid3.Rows.Count == 1)
        //    {
        //        c1FlexGrid3.Rows.Add();
        //        c1FlexGrid5.Rows.Add();
        //    }
        //    c1FlexGrid3.GetCellRange(1, 1).StyleNew.BackColor = Color.FromArgb(255, 255, 255);
        //    c1FlexGrid3.GetCellRange(1, 2).StyleNew.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_1.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_3.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_5.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_7.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_9.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_2.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_4.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_6.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_8.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_6_10.BackColor = Color.FromArgb(255, 255, 255);
        //    label77.BackColor = Color.FromArgb(252, 228, 214);
        //    label22.BackColor = Color.FromArgb(252, 228, 214);
        //    label333.BackColor = Color.FromArgb(252, 228, 214);
        //    label335.BackColor = Color.FromArgb(252, 228, 214);
        //    label328.BackColor = Color.FromArgb(252, 228, 214);
        //    label325.BackColor = Color.FromArgb(252, 228, 214);
        //    label86.BackColor = Color.DarkGray;

        //    // 引合タブのチェックは必ず行う

        //    // 0:新規登録 1:更新
        //    //if (flg <= 1)
        //    //{
        //    //引合タブ　2.基本情報　
        //    //売上年度
        //    if (String.IsNullOrEmpty(item1_3.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_3.BackColor = Color.FromArgb(255, 204, 255);
        //    }
        //    //登録日
        //    if (item1_9.CustomFormat != "")
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_9.BackColor = Color.FromArgb(255, 204, 255);
        //        label77.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //受託課所支部
        //    if (String.IsNullOrEmpty(item1_10.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_10.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //契約担当者
        //    if (String.IsNullOrEmpty(item1_11.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_11.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //案件(受託)フォルダ
        //    if (String.IsNullOrEmpty(item1_12.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_12.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //引合タブ　3.案件情報
        //    //業務名称	
        //    if (String.IsNullOrEmpty(item1_13.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_13.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //契約区分
        //    if (String.IsNullOrEmpty(item1_14.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_14.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //入札状況が、1:入札前以外の場合、
        //    if (item1_17.SelectedValue.ToString() != "1")
        //    {
        //        //入札(予定)日
        //        if (item1_16.CustomFormat == " ")
        //        {
        //            requiredFlag = false;
        //            InsertErrorFlag = true;
        //            item1_16.BackColor = Color.FromArgb(255, 204, 255);
        //            label22.BackColor = Color.FromArgb(255, 204, 255);
        //        }
        //    }

        //    //引合タブ　4.発注者情報
        //    //発注者コード
        //    if (String.IsNullOrEmpty(item1_19.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_19.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //発注者区分1
        //    if (String.IsNullOrEmpty(item1_20.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_20.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //発注者区分2
        //    if (String.IsNullOrEmpty(item1_21.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_21.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //引合タブ　6.当会対応 
        //    //参考見積額(税抜)
        //    if (String.IsNullOrEmpty(item1_36.Text))
        //    {
        //        requiredFlag = false;
        //        InsertErrorFlag = true;
        //        item1_36.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    ////入札タブ 2.当会対応
        //    ////当会応札
        //    //if (String.IsNullOrEmpty(item2_2_1.Text))
        //    //{
        //    //    // 入札状況が入札成立の場合
        //    //    if (errorNyuusatsuFlag == true)
        //    //    {
        //    //        //requiredFlag = false;
        //    //    InsertErrorFlag = true;
        //    //        item2_2_1.BackColor = Color.FromArgb(255, 204, 255);
        //    //    }
        //    //}

        //    // 1:更新 2:チェック用出力 3:起案
        //    if (flg >= 1)
        //    {
        //        //入札タブ 2.当会対応
        //        // 入札状況が入札成立の場合
        //        if (errorNyuusatsuFlag == true)
        //        {
        //            //当会応札
        //            if (String.IsNullOrEmpty(item2_2_1.Text))
        //            {
        //                requiredFlag = false;
        //                item2_2_1.BackColor = Color.FromArgb(255, 204, 255);
        //            }
        //        }

        //        // 入札タブの3.入札結果の落札者の場合
        //        if (errorKeiyakuFlag == true)
        //        {
        //            //契約タブ 1.契約情報緒
        //            //契約締結(変更)日
        //            if (item3_1_3.CustomFormat == " ")
        //            {
        //                requiredFlag = false;
        //                item3_1_3.BackColor = Color.FromArgb(255, 204, 255);
        //                label335.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }

        //            //起業日
        //            if (item3_1_4.CustomFormat == " ")
        //            {
        //                requiredFlag = false;
        //                item3_1_4.BackColor = Color.FromArgb(255, 204, 255);
        //                label333.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }

        //            //契約工期自
        //            if (item3_1_6.CustomFormat == " ")
        //            {
        //                requiredFlag = false;
        //                item3_1_6.BackColor = Color.FromArgb(255, 204, 255);
        //                label328.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }

        //            //契約工期至
        //            if (item3_1_7.CustomFormat == " ")
        //            {
        //                requiredFlag = false;
        //                item3_1_7.BackColor = Color.FromArgb(255, 204, 255);
        //                label325.BackColor = Color.FromArgb(255, 204, 255);
        //            }

        //            //契約金額の税込
        //            if (String.IsNullOrEmpty(item3_1_13.Text))
        //            {
        //                requiredFlag = false;
        //                item3_1_13.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }

        //            //受託金額(税込)
        //            if (String.IsNullOrEmpty(item3_1_15.Text))
        //            {
        //                requiredFlag = false;
        //                item3_1_15.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }

        //            //受託外金額(税込)
        //            if (String.IsNullOrEmpty(item3_1_16.Text))
        //            {
        //                requiredFlag = false;
        //                item3_1_16.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }


        //            //契約タブ 4.管理者・担当者
        //            //管理技術者 
        //            if (String.IsNullOrEmpty(item3_4_1.Text))
        //            {
        //                requiredFlag = false;
        //                item3_4_1.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }

        //            //業務担当者
        //            if (String.IsNullOrEmpty(item3_4_4.Text))
        //            {
        //                requiredFlag = false;
        //                item3_4_4.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }

        //            //窓口担当者
        //            if (String.IsNullOrEmpty(item3_4_5.Text))
        //            {
        //                requiredFlag = false;
        //                item3_4_5.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }

        //            //担当技術者 c1FlexGrid3が2行だったら 1行目ヘッダー
        //            if (c1FlexGrid3.Rows.Count < 2)
        //            {
        //                c1FlexGrid3.Rows.Add();
        //            }
        //            //2行目がnullでないことを確認する
        //            if (c1FlexGrid3[1, 1] == null || c1FlexGrid3[1, 1].ToString() == "")
        //            {
        //                c1FlexGrid3.GetCellRange(1, 1).StyleNew.BackColor = Color.FromArgb(255, 204, 255);
        //                c1FlexGrid3.GetCellRange(1, 2).StyleNew.BackColor = Color.FromArgb(255, 204, 255);
        //                requiredNyuusatsuKeiyakuFlag = true;
        //            }
        //        }
        //    }

        //    //必須項目エラーの出力
        //    if (!requiredFlag || requiredNyuusatsuKeiyakuFlag)
        //    {
        //        set_error(GlobalMethod.GetMessage("E10010", ""));
        //    }

        //    //データチェック
        //    //引合タブ 2.基本情報 
        //    // 1:更新 2:チェック用出力 3:起案 4:変更伝票後の起案
        //    if (flg <= 3)
        //    {
        //        //契約担当者の部所CDと、受託課所支部（item1_10）が違う(未入力の場合チェックしない)
        //        if (item1_11.Text != "" && !item1_11_Busho.Text.Equals(item1_10.SelectedValue))
        //        {
        //            set_error(GlobalMethod.GetMessage("W10602", ""));
        //        }

        //        //案件受託フォルダのフォーマット: ^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$ が違う
        //        if (item1_12.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(item1_12.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
        //        {
        //            InsertErrorFlag = false;
        //            set_error(GlobalMethod.GetMessage("E10017", ""));
        //            varidateFlag = false;
        //        }

        //        //引合タブ 5.発注担当者情報
        //        //メールアドレスのフォーマット：が違う
        //        if (item1_29.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(item1_29.Text, @"^((\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
        //        {
        //            InsertErrorFlag = false;
        //            set_error(GlobalMethod.GetMessage("E10605", ""));
        //            varidateFlag = false;

        //        }

        //    }
        //    //　1:更新
        //    if (flg >= 1)
        //    {
        //        //入札成立の場合のみチェック
        //        if (errorNyuusatsuFlag)
        //        {
        //            //引合タブ 7.業務内容
        //            //調査部、事業普及部、情報システム部、総合研究所の合計が100でない場合
        //            if (GetDouble(item1_7_1_5_1.Text).ToString("F2") != "100.00")
        //            {
        //                set_error("部門配分の合計が不正です。");
        //                varidateFlag = false;
        //            }

        //            //調査部の配分がある場合
        //            if (GetDouble(item1_7_1_1_1.Text) > 0)
        //            {
        //                //調査部　業務別配分率の合計が100%ではない
        //                if (GetDouble(item1_7_2_13_1.Text).ToString("F2") != "100.00")
        //                {
        //                    set_error(GlobalMethod.GetMessage("E10604", "(調査部)"));
        //                    varidateFlag = false;
        //                }

        //                //調査部　業務別配分率の資材調査、営繕調査、機器類調査、工事費調査のどれかに配分率が入っていない
        //                if (GetDouble(item1_7_2_1_1.Text) > 0 || GetDouble(item1_7_2_1_1.Text) > 0 || GetDouble(item1_7_2_1_1.Text) > 0 || GetDouble(item1_7_2_1_1.Text) > 0)
        //                {

        //                }
        //                else
        //                {
        //                    set_error(GlobalMethod.GetMessage("E10604", "(調査部)"));
        //                    varidateFlag = false;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            //引合タブ 7.業務内容
        //            //調査部、事業普及部、情報システム部、総合研究所の合計が0か100でない場合
        //            if (GetDouble(item1_7_1_5_1.Text).ToString("F2") != "100.00" && GetDouble(item1_7_1_5_1.Text).ToString("F2") != "0.00")
        //            {
        //                set_error("部門配分の合計が不正です。(引合タブ)");
        //                varidateFlag = false;
        //            }

        //            //調査部の配分がある場合
        //            if (GetDouble(item1_7_1_1_1.Text) > 0)
        //            {
        //                //調査部　業務別配分率の合計が0か100%ではない
        //                if (GetDouble(item1_7_2_13_1.Text).ToString("F2") != "100.00")
        //                {
        //                    set_error(GlobalMethod.GetMessage("E10604", "(調査部)"));
        //                    varidateFlag = false;
        //                }

        //                //調査部　業務別配分率の資材調査、営繕調査、機器類調査、工事費調査のどれかに配分率が入っていない
        //                if (GetDouble(item1_7_2_1_1.Text) > 0 || GetDouble(item1_7_2_1_1.Text) > 0 || GetDouble(item1_7_2_1_1.Text) > 0 || GetDouble(item1_7_2_1_1.Text) > 0)
        //                {

        //                }
        //                else
        //                {
        //                    set_error("業務別配分（資材調査・営繕調査、機器類調査、工事費調査）の合計が不正です。");
        //                    varidateFlag = false;
        //                }
        //            }
        //        }

        //        //事業部コード（案件番号の頭文字1つ）がAのとき
        //        // ⇒Tが正
        //        String jigyoCd = Header1.Text.Substring(0, 1);
        //        if ("T".Equals(jigyoCd))
        //        {
        //            //契約図書が空
        //            if (String.IsNullOrEmpty(item3_1_26.Text))
        //            {
        //                set_error(GlobalMethod.GetMessage("W10601", ""));
        //            }
        //            else
        //            {
        //                //契約図書のフォーマット：^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$　が違う
        //                if (!System.Text.RegularExpressions.Regex.IsMatch(item3_1_26.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
        //                {
        //                    set_error(GlobalMethod.GetMessage("E10017", ""));
        //                    varidateFlag = false;
        //                }
        //            }
        //        }


        //        //技術者評価タブ 評価・評点
        //        if (item4_1_1.Text != "" && (int.Parse(item4_1_1.Text) < 0 || int.Parse(item4_1_1.Text) > 100))
        //        {
        //            set_error(GlobalMethod.GetMessage("E10913", "(業務評点)"));
        //            varidateFlag = false;
        //        }
        //        if (item4_1_3.Text != "" && (int.Parse(item4_1_3.Text) < 0 || int.Parse(item4_1_3.Text) > 100))
        //        {
        //            set_error(GlobalMethod.GetMessage("E10913", "(管理技術者評点)"));
        //            varidateFlag = false;
        //        }
        //        if (item4_1_5.Text != "" && (int.Parse(item4_1_5.Text) < 0 || int.Parse(item4_1_5.Text) > 100))
        //        {
        //            set_error(GlobalMethod.GetMessage("E10913", "(協力担当者評点)"));
        //            varidateFlag = false;
        //        }
        //        //請求書のパスのフォーマット：^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$　がちがう
        //        if (item4_1_8.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(item4_1_8.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
        //        {
        //            set_error(GlobalMethod.GetMessage("E10017", ""));
        //            varidateFlag = false;
        //        }
        //    }

        //    //新規登録時のエラーチェックにひっかかった場合のみ、エラーとして終了
        //    if (InsertErrorFlag || (flg == 3 && (!varidateFlag || !requiredFlag)))
        //    {
        //        //　エラーでも1:更新
        //        //0:新規登録 1:更新 2:チェック用出力(赤伝・黒伝) 3:起案
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}

        // エラーチェック true:エラー false：正常
        private Boolean ErrorFLG(int flg)
        {
            // エラーフラグ true：エラー、false：正常
            Boolean ErrorFLG = false;
            //必須チェック デフォルト　requiredFlag:true /　falseだとエラー
            //Boolean requiredFlag = true;
            // データチェックフラグ true：正常、false：エラー
            //Boolean varidateFlag = true;

            // 入札エラーON/OFFフラグ true：ON（警告）、false：OFF（無警告）
            Boolean errorNyuusatsuFlag = false;
            // 契約エラーON/OFFフラグ true：ON（警告）、false：OFF（無警告）
            Boolean errorKeiyakuFlag = false;

            //入札・契約チェック デフォルト　true:警告 /　false:警告なし
            //Boolean requiredNyuusatsuKeiyakuFlag = false;

            // 引合エラーフラグ true:エラー false:正常
            //Boolean hikiaiErrorFlg = false;
            // 入札エラーフラグ true:エラー false:正常
            //Boolean nyuusatsuErrorFlg = false;
            // 契約エラーフラグ true:エラー false:正常
            //Boolean KeiyakuErrorFlg = false;

            // 起案の場合は、エラーチェックを行う
            if (flg >= 3)
            {
                errorNyuusatsuFlag = true;
                errorKeiyakuFlag = true;
            }
            else
            {
                // ①入札タブ：入札状況が入札成立となったら入札のチェックを行う（更新可能）
                // item2_1_1:入札タブの1.入札状況の入札状況
                // NYUUSATSU_SEIRITSU:2(入札成立)
                if (item2_1_1.Text == GlobalMethod.GetCommonValue2("NYUUSATSU_SEIRITSU"))
                {
                    errorNyuusatsuFlag = true;
                }
                // ②契約タブ：調査会様での入札が成立したら契約のチェック処理を行う（更新可能）
                // item2_3_7:入札タブの3.入札結果の落札者
                // ENTORY_TOUKAI:建設物価調査会
                if (item2_3_7.Text == GlobalMethod.GetCommonValue2("ENTORY_TOUKAI"))
                {
                    errorKeiyakuFlag = true;
                }
            }

            set_error("", 0);
            //0:新規登録 1:更新 2:チェック用出力(赤伝・黒伝) 3:起案

            //必須チェックを行う項目の背景を白に戻す。
            item1_3.BackColor = Color.FromArgb(255, 255, 255);
            item1_2_KoukiNendo.BackColor = Color.FromArgb(255, 255, 255);
            item1_9.BackColor = Color.FromArgb(255, 255, 255);
            item1_10.BackColor = Color.FromArgb(255, 255, 255);
            item1_11.BackColor = Color.FromArgb(255, 255, 255);
            item1_12.BackColor = Color.FromArgb(255, 255, 255);
            item1_13.BackColor = Color.FromArgb(255, 255, 255);
            item1_14.BackColor = Color.FromArgb(255, 255, 255);
            item1_15.BackColor = Color.FromArgb(255, 255, 255); //えんとり君修正STEP2　ご指摘：1392
            item1_16.BackColor = Color.FromArgb(255, 255, 255);
            item1_19.BackColor = Color.FromArgb(255, 255, 255);
            item1_20.BackColor = Color.FromArgb(255, 255, 255);
            item1_21.BackColor = Color.FromArgb(255, 255, 255);
            item1_36.BackColor = Color.FromArgb(255, 255, 255);
            item2_2_1.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_3.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_4.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_6.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_3.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_4.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_6.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_7.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_13.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_15.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_16.BackColor = Color.FromArgb(255, 255, 255);
            item3_4_1.BackColor = Color.FromArgb(255, 255, 255);
            item3_4_4.BackColor = Color.FromArgb(255, 255, 255);
            item3_4_5.BackColor = Color.FromArgb(255, 255, 255);
            if (c1FlexGrid3.Rows.Count == 1)
            {
                c1FlexGrid3.Rows.Add();
                c1FlexGrid5.Rows.Add();
            }
            c1FlexGrid3.GetCellRange(1, 1).StyleNew.BackColor = Color.FromArgb(255, 255, 255);
            c1FlexGrid3.GetCellRange(1, 2).StyleNew.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_1.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_3.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_5.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_7.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_9.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_2.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_4.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_6.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_8.BackColor = Color.FromArgb(255, 255, 255);
            item3_6_10.BackColor = Color.FromArgb(255, 255, 255);
            label77.BackColor = Color.FromArgb(252, 228, 214);
            label22.BackColor = Color.FromArgb(252, 228, 214);
            label333.BackColor = Color.FromArgb(252, 228, 214);
            label335.BackColor = Color.FromArgb(252, 228, 214);
            label328.BackColor = Color.FromArgb(252, 228, 214);
            label325.BackColor = Color.FromArgb(252, 228, 214);
            label86.BackColor = Color.DarkGray;
            //えんとり君修正STEP2
            item3_7_2_26_1.BackColor = Color.FromArgb(255, 255, 255);
            label502.BackColor = Color.FromArgb(252, 228, 214);
            c1FlexGrid4.GetCellRange(2, 3).StyleNew.BackColor = Color.FromArgb(255, 255, 255);
            c1FlexGrid4.GetCellRange(2, 11).StyleNew.BackColor = Color.FromArgb(255, 255, 255);
            c1FlexGrid4.GetCellRange(2, 19).StyleNew.BackColor = Color.FromArgb(255, 255, 255);
            c1FlexGrid4.GetCellRange(2, 27).StyleNew.BackColor = Color.FromArgb(255, 255, 255);

            item3_2_1_1.BackColor = Color.FromArgb(255, 255, 255);
            item3_2_2_1.BackColor = Color.FromArgb(255, 255, 255);
            item3_2_3_1.BackColor = Color.FromArgb(255, 255, 255);
            item3_2_4_1.BackColor = Color.FromArgb(255, 255, 255);

            // 20210505 チェック処理の記述見直し
            // 更新不可の場合、ErrorFLGをtrueにして返却する
            //===================================================================================================
            // 新規登録時のチェック
            //===================================================================================================
            if (flg == 0)
            {
                // 引合タブの必須チェック
                if (hikiaiRequireCheck())
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "引合"));
                    ErrorFLG = true;
                }

                // 引合タブのデータチェック
                if (hikiaiDataCheck())
                {
                    ErrorFLG = true;
                }

                // 入札タブのデータチェック（業務配分と業務別配分のチェックのため追加）
                if (nyuusatsuDataCheck())
                {
                    ErrorFLG = true;
                }
            }

            //===================================================================================================
            // 更新、起案時のチェック
            //===================================================================================================
            if (flg == 1)
            {
                // 引合タブの必須チェック
                if (hikiaiRequireCheck())
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "引合"));
                    ErrorFLG = true;
                }

                // 契約タブの必須チェック（調査会様での入札が成立時）
                if (errorKeiyakuFlag)
                {
                    if (keiyakuRequireCheck())
                    {
                        // 起案済みの場合はエラーとする
                        if (item3_1_2.Checked == true)
                        {
                            // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                            set_error(GlobalMethod.GetMessage("E10010", "契約"));
                            ErrorFLG = true;
                        }
                    }
                }

                // 引合タブのデータチェック
                if (hikiaiDataCheck())
                {
                    ErrorFLG = true;
                }

                // 入札タブのデータチェック
                // 業務配分と業務別配分のチェックだけのため、条件判定しないようにコメント化
                //if (errorNyuusatsuFlag)
                //{
                if (nyuusatsuDataCheck())
                {
                    ErrorFLG = true;
                }
                //}

                // 契約タブのデータチェック（調査会様での入札が成立時）
                if (errorKeiyakuFlag)
                {
                    if (keiyakuDataCheck())
                    {
                        ErrorFLG = true;
                    }
                }

                // 技術者評価タブのデータチェック
                if (gijyutsushahyoukaDataCheck())
                {
                    ErrorFLG = true;
                }
            }

            //===================================================================================================
            // チェック用帳票出力時のチェック
            //===================================================================================================
            if (flg == 2)
            {
                // 引合タブの必須チェック
                if (hikiaiRequireCheck())
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "引合"));
                    ErrorFLG = true;
                }

                // 契約タブの必須チェック（調査会様での入札が成立時）
                if (errorKeiyakuFlag)
                {
                    if (keiyakuRequireCheck())
                    {
                        // 起案済みの場合はエラーとする
                        if (item3_1_2.Checked == true)
                        {
                            // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                            set_error(GlobalMethod.GetMessage("E10010", "契約"));
                            ErrorFLG = true;
                        }
                    }
                }

                // 引合タブのデータチェック
                if (hikiaiDataCheck())
                {
                    ErrorFLG = true;
                }

                // 入札タブのデータチェック
                // 業務配分と業務別配分のチェックだけのため、条件判定しないようにコメント化
                //if (errorNyuusatsuFlag)
                //{
                if (nyuusatsuDataCheck())
                {
                    ErrorFLG = true;
                }
                //}

                // 契約タブのデータチェック（調査会様での入札が成立時）
                if (errorKeiyakuFlag)
                {
                    if (keiyakuDataCheck())
                    {
                        ErrorFLG = true;
                    }
                }

                // 技術者評価タブのデータチェック
                if (gijyutsushahyoukaDataCheck())
                {
                    ErrorFLG = true;
                }

                //============================================================================
                // 起案用のチェック？（できれば外出ししてまとめたい）
                //============================================================================
                //Double totalZero = Convert.ToDouble(0);
                Double totalHundred = Convert.ToDouble(100);

                // 契約タブ
                // 売上年度 4桁じゃなかったらエラー
                if (4 != item3_1_5.SelectedValue.ToString().Length)
                {
                    set_error(GlobalMethod.GetMessage("E10011", ""));
                }

                // 入札タブ
                // 入札状況が入札成立でなければ起案エラー
                if (!GlobalMethod.GetCommonValue1("NYUUSATSU_SEIRITSU").Equals(item2_1_1.SelectedValue.ToString()))
                {
                    set_error(GlobalMethod.GetMessage("E10702", ""));
                }

                // 落札者が建設物価調査会でなければ起案エラー
                if (!GlobalMethod.GetCommonValue2("ENTORY_TOUKAI").Equals(item2_3_7.Text))
                {
                    set_error(GlobalMethod.GetMessage("E70048", ""));
                }

                // 契約タブ
                // 契約タブの1.契約情報の契約金額の税込が0円の場合
                if (errorKeiyakuFlag)
                {
                    long item13 = GetLong(item3_1_13.Text);
                    if (item13 == 0)
                    {
                        // 0円起案です。
                        set_error(GlobalMethod.GetMessage("W10701", ""));
                    }
                }

                // 税込
                // 契約タブの1.契約情報の消費税率が空ではない場合
                if (!String.IsNullOrEmpty(item3_1_10.Text))
                {
                    Double keiyakuAmount = GetDouble(item3_1_13.Text);  // 契約金額の税込
                    Decimal taxAmount = Decimal.Parse(item3_1_10.Text); // 消費税
                    Double inTaxAmount = GetDouble(item3_1_14.Text);    // 内消費税
                    Double taxPercent = Double.Parse(item3_1_10.Text);  // 消費税率

                    // 契約金額の税込 / (100 + 消費税率))* 消費税率, 0) の小数点切り捨て
                    Double amount = Math.Floor(keiyakuAmount / (totalHundred + taxPercent) * taxPercent);

                    // 内消費税がamountと一致しない
                    if (!Double.Equals(inTaxAmount, amount))
                    {
                        // 起案は出来ますが、契約金額(税込)と内消費税が一致しません。確認してください。
                        set_error(GlobalMethod.GetMessage("E10704", ""));

                    }
                }

                // 受託金額(税込)と配分額(税込)のチェック
                Double jutakuTax = GetDouble(item3_1_15.Text);      // 1.契約情報の受託金額(税込)
                Double totalAmount = GetDouble(item3_2_5_1.Text);   // 2.配分情報の配分情報の配分額(税込)の合計

                //受託金額(税込)と配分額(税込)の合計が一致しない
                if (!Double.Equals(jutakuTax, totalAmount))
                {
                    // 起案は出来ますが、受託契約金額と各配分額の合計が一致していません。確認して下さい。
                    set_error(GlobalMethod.GetMessage("E10705", ""));
                }

                if (errorKeiyakuFlag)
                {
                    // 契約タブ
                    // 契約工期至と売上年度のチェック
                    String format = "yyyy/MM/dd";   // 日付フォーマット

                    // 1.契約情報の契約工期至と売上年度が空でない場合
                    if (item3_1_7.CustomFormat == "" && !String.IsNullOrEmpty(item3_1_5.Text))
                    {
                        // 売上年度 +1年 の3月31日
                        int year = Int32.Parse(item3_1_5.SelectedValue.ToString()) + 1;
                        String date = year + "/03/31";

                        // 日付型
                        DateTime nextYear = DateTime.ParseExact(date, format, null);
                        DateTime keiyaku = DateTime.ParseExact(item3_1_7.Text, format, null);

                        // 売上年度+1/03/31よりも、契約工期の完了日が未来日付の場合エラー
                        if (nextYear.Date < keiyaku.Date)
                        {
                            // 工期完了日が売上年度を超えています。年度をまたぐ場合は、売上年度を工期完了日にあわせてください。
                            set_error(GlobalMethod.GetMessage("E10706", ""));
                        }
                    }

                    // 引合タブ
                    // 調査部　売上計上情報と業務別配分のチェック
                    // 7.業務内容の調査部 業務別配分の合計が100の場合
                    if (GetDouble(item1_7_2_13_1.Text).ToString("F2") == "100.00")
                    {
                        // 計上額の合計の取得
                        long keijoTotal = GetKeijogakuGoukei("調査部");

                        // 計上額の合計が0の場合、未入力ありと判断する
                        if (keijoTotal == 0)
                        {
                            // 売上計上情報の工期日付か売上計上額が未入力です。
                            set_error(GlobalMethod.GetMessage("E10715", "(調査部)"));
                        }

                        // 計上額の合計が0でない場合
                        if (keijoTotal != 0)
                        {
                            // 2.配分情報の配分額(税込)
                            long haibunTax = GetLong(item3_2_1_1.Text);

                            // 配分額(税込)と、計上額の合計のチェック
                            if (!long.Equals(haibunTax, keijoTotal))
                            {
                                set_error(GlobalMethod.GetMessage("E10717", "(調査部)"));
                            }

                        }
                    }

                    // 事業普及部　売上計上情報と業務別配分のチェック
                    // 7.業務内容の事業普及部 業務別配分の合計が100の場合
                    if (GetDouble(item1_7_1_2_1.Text).ToString("F2") == "100.00")
                    {
                        // 計上額の合計の取得
                        long keijoTotal = GetKeijogakuGoukei("事業普及部");

                        // 計上額の合計が0の場合、未入力ありと判断する
                        if (keijoTotal == 0)
                        {
                            // 売上計上情報の工期日付か売上計上額が未入力です。
                            set_error(GlobalMethod.GetMessage("E10715", "(事業普及部)"));
                        }

                        // 計上額の合計が0でない場合
                        if (keijoTotal != 0)
                        {
                            // 2.配分情報の配分額(税込)
                            long haibunTax = GetLong(item3_2_2_1.Text);

                            // 配分額(税込)と、計上額の合計のチェック
                            if (!long.Equals(haibunTax, keijoTotal))
                            {
                                set_error(GlobalMethod.GetMessage("E10717", "(事業普及部)"));
                            }

                        }
                    }

                    // 情報システム部　売上計上情報と業務別配分のチェック
                    // 7.業務内容の情報システム部 業務別配分の合計が100の場合
                    if (GetDouble(item1_7_1_3_1.Text).ToString("F2") == "100.00")
                    {
                        // 計上額の合計の取得
                        long keijoTotal = GetKeijogakuGoukei("情報システム部");

                        // 計上額の合計が0の場合、未入力ありと判断する
                        if (keijoTotal == 0)
                        {
                            // 売上計上情報の工期日付か売上計上額が未入力です。
                            set_error(GlobalMethod.GetMessage("E10715", "(情報システム部)"));
                        }

                        // 計上額の合計が0でない場合
                        if (keijoTotal != 0)
                        {
                            // 2.配分情報の配分額(税込)
                            long haibunTax = GetLong(item3_2_3_1.Text);

                            // 配分額(税込)と、計上額の合計のチェック
                            if (!long.Equals(haibunTax, keijoTotal))
                            {
                                set_error(GlobalMethod.GetMessage("E10717", "(情報システム部)"));
                            }

                        }
                    }

                    // 総合研究所　売上計上情報と業務別配分のチェック
                    // 7.業務内容の総合研究所 業務別配分の合計が100の場合
                    if (GetDouble(item1_7_1_4_1.Text).ToString("F2") == "100.00")
                    {
                        // 計上額の合計の取得
                        long keijoTotal = GetKeijogakuGoukei("総合研究所");

                        // 計上額の合計が0の場合、未入力ありと判断する
                        if (keijoTotal == 0)
                        {
                            // 売上計上情報の工期日付か売上計上額が未入力です。
                            set_error(GlobalMethod.GetMessage("E10715", "(総合研究所)"));
                        }

                        // 計上額の合計が0でない場合
                        if (keijoTotal != 0)
                        {
                            // 2.配分情報の配分額(税込)
                            long haibunTax = GetLong(item3_2_4_1.Text);

                            // 配分額(税込)と、計上額の合計のチェック
                            if (!long.Equals(haibunTax, keijoTotal))
                            {
                                set_error(GlobalMethod.GetMessage("E10717", "(総合研究所)"));
                            }

                        }
                    }

                }

            }

            //===================================================================================================
            // エントリーシート作成・出力時のチェック
            //===================================================================================================
            if (flg == 3)
            {
                // 引合タブの必須チェック
                if (hikiaiRequireCheck())
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "引合"));
                    ErrorFLG = true;
                }

                // 契約タブの必須チェック
                if (keiyakuRequireCheck())
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "契約"));
                    ErrorFLG = true;
                }

                // 引合タブのデータチェック
                if (hikiaiDataCheck())
                {
                    ErrorFLG = true;
                }

            }

            // チェック処理を実施した結果、更新可の場合はfalse：正常で返す
            return ErrorFLG;

            //引合タブのチェックは必ず行う
            //引合タブの必須チェック
            //必須チェックtrue:エラー / false:正常
            //requiredFlag = hikiaiRequireCheck();
            // 引合エラーフラグ true:エラー false:正常
            // 更新だけど、引合のエラーチェックはエラーにする為
            //if (!requiredFlag)
            //{
            //    //必須項目エラーの出力
            //    //if (requiredFlag)
            //    //{
            //        // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
            //        set_error(GlobalMethod.GetMessage("E10010", "引合"));
            //    //}
            //    hikiaiErrorFlg = true;
            //}

            // 1:更新(起案) 2:チェック用出力 3:エントリーチェックシート
            //if (flg >= 1)
            //{
            //    // 入札タブの3.入札結果の落札者の場合に契約タブをチェックする
            //    if (errorKeiyakuFlag)
            //    {
            //        // 契約タブの必須チェック true:エラー / false:正常
            //        if (keiyakuRequireCheck())
            //        {
            //            //必須チェック true：正常 / false：エラー
            //            requiredFlag = false;
            //            //必須項目エラーの出力
            //            if (!requiredFlag)
            //            {
            //                // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
            //                set_error(GlobalMethod.GetMessage("E10010", "契約"));
            //            }

            //            if (item3_1_2.Checked)
            //            {
            //                KeiyakuErrorFlg = true;
            //            }
            //        }
            //    }
            //}

            ////必須項目エラーの出力
            //if (!requiredFlag)
            //{
            //    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
            //    set_error(GlobalMethod.GetMessage("E10010", ""));
            //}

            //データチェック
            //引合タブ 2.基本情報 
            // 0:新規 1:更新 2:チェック用出力 3:エントリーチェックシート
            //if (flg <= 1)
            // 引合のデータチェックは全部で見る
            //if (flg <= 3)
            //    {
            //    // 引合データチェック true:正常 false：エラー
            //    if (hikiaiDataCheck())
            //    {
            //        // エラーフラグ true：エラー、false：正常
            //        ErrorFLG = true;
            //        hikiaiErrorFlg = true;
            //    }
            //}

            //契約タブ　1.契約情報 2:チェック用出力
            //if (flg == 2)
            //{
            //    //契約タブ　売上年度 4桁じゃなかったらエラー
            //    if (4 != item3_1_5.SelectedValue.ToString().Length)
            //    {
            //        //varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10011", ""));
            //    }
            //    //入札状況が入札成立でなければ起案エラー
            //    if (!GlobalMethod.GetCommonValue1("NYUUSATSU_SEIRITSU").Equals(item2_1_1.SelectedValue.ToString()))
            //    {
            //        //varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10702", ""));
            //    }
            //    //落札者が建設物価調査会でなければ起案エラー
            //    if (!GlobalMethod.GetCommonValue2("ENTORY_TOUKAI").Equals(item2_3_7.Text))
            //    {
            //        //varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E70048", ""));
            //    }
            //}

            //　1:更新
            //if (flg == 1)
            //{
            // 入札成立時、配分等をチェック
            //if (errorNyuusatsuFlag == true)
            //{
            //// エラーフラグtrue:正常 /false:エラー
            //if (!nyuusatsuDataCheck())
            //{
            //    // 配分チェックに引っかかった場合はエラーとする
            //    nyuusatsuErrorFlg = true;
            //    ErrorFLG = true;
            //}
            //}

            //事業部コード（案件番号の頭文字1つ）がAのとき
            // ⇒Tが正
            //String jigyoCd = Header1.Text.Substring(0, 1);
            //if ("T".Equals(jigyoCd))
            //{
            //    //契約図書が空
            //    if (String.IsNullOrEmpty(item3_1_26.Text))
            //    {
            //        set_error(GlobalMethod.GetMessage("W10601", ""));
            //    }
            //    //契約図書のフォーマット：^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$　が違う
            //    if (!System.Text.RegularExpressions.Regex.IsMatch(item3_1_26.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            //    {
            //        set_error(GlobalMethod.GetMessage("E10017", ""));
            //        ErrorFLG = true;
            //    }
            //}

            //　エラーとさせない為にコメントアウト
            ////調査部配分率が0ではない場合、調査部 業務別配分が100でないとエラー
            ////if (GetDouble(item3_7_1_1_1.Text) > 0 && GetDouble(item3_7_2_26_1.Text) != 100)
            //if (GetDouble(item3_7_1_1_1.Text) > 0 && item3_7_2_26_1.Text != "100.00%")
            //    {
            //    if (errorKeiyakuFlag == true)
            //    {
            //        set_error(GlobalMethod.GetMessage("E70045", "(契約タブ)"));
            //        //ErrorFLG = true;
            //    }
            //}
            // 調査部 業務別配分が100でないとエラー
            //if (item3_7_2_26_1.Text != "100.00%" && item3_7_2_26_1.Text != "0.00%")
            //{
            //    // 調査業務別　配分の合計が100になるように入力してください。
            //    set_error(GlobalMethod.GetMessage("E70045", "契約タブ"));
            //    ErrorFLG = true;
            //}
            //技術者評価タブ 評価・評点
            //    if (item4_1_1.Text != "" && (int.Parse(item4_1_1.Text) < 0 || int.Parse(item4_1_1.Text) > 100))
            //    {
            //        set_error(GlobalMethod.GetMessage("E10913", "業務評点"));
            //        ErrorFLG = true;
            //    }
            //    if (item4_1_3.Text != "" && (int.Parse(item4_1_3.Text) < 0 || int.Parse(item4_1_3.Text) > 100))
            //    {
            //        set_error(GlobalMethod.GetMessage("E10913", "管理技術者評点"));
            //        ErrorFLG = true;
            //    }
            //    if (item4_1_5.Text != "" && (int.Parse(item4_1_5.Text) < 0 || int.Parse(item4_1_5.Text) > 100))
            //    {
            //        set_error(GlobalMethod.GetMessage("E10913", "協力担当者評点"));
            //        ErrorFLG = true;
            //    }
            //    //請求書のパスのフォーマット：^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$　がちがう
            //    if (item4_1_8.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(item4_1_8.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            //    {
            //        set_error(GlobalMethod.GetMessage("E10017", ""));
            //        //ErrorFLG = true;
            //    }
            //}
            // 2:チェック用出力
            //if (flg == 2)
            //{
            ////契約タブ　売上年度 4桁じゃなかったらエラー
            //if (4 != item3_1_5.SelectedValue.ToString().Length)
            //{
            //    //varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E10011", ""));
            //}
            ////入札状況が入札成立でなければ起案エラー
            //if (!GlobalMethod.GetCommonValue1("NYUUSATSU_SEIRITSU").Equals(item2_1_1.SelectedValue.ToString()))
            //{
            //    //varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E10702", ""));
            //}
            ////落札者が建設物価調査会でなければ起案エラー
            //if (!GlobalMethod.GetCommonValue2("ENTORY_TOUKAI").Equals(item2_3_7.Text))
            //{
            //    //varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E70048", ""));
            //}

            //引合タブ
            //部門配分
            //Double totalZero = Convert.ToDouble(0);
            //Double totalHundred = Convert.ToDouble(100);

            //　エラーとさせない為にコメントアウト
            ////引合タブの7.業務内容の調査部の合計が0か100でない
            //Decimal chosaTotal = Decimal.Parse(item1_7_2_13_1.Text.Substring(0, item1_7_2_13_1.Text.Length - 1));
            //Double chosaTotal = GetDouble(item1_7_2_13_1.Text);
            //if (Decimal.Compare(chosaTotal, totalHundred) != 0 && Decimal.Compare(chosaTotal, totalZero) != 0)
            //{
            //    //varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E10703", "(調査部)"));

            //}

            ////引合タブの7.業務内容の総合研究所の合計が0か100でない　
            //Decimal sogoTotal = Decimal.Parse(item1_7_1_4_1.Text.Substring(0, item1_7_1_4_1.Text.Length - 1));
            //if (Decimal.Compare(sogoTotal, totalHundred) != 0 && Decimal.Compare(sogoTotal, totalZero) != 0)
            //{
            //    //varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E10703", "(総合研究所)"));

            //}


            //契約タブ
            //契約金額の税込
            //契約タブの1.契約情報の契約金額の税込が0円の場合
            //String item13 = item3_1_13.Text.Substring(1, item3_1_13.Text.Length - 1);
            //long item13 = GetLong(item3_1_13.Text);
            //if (item13 == 0)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        // 0円起案です。
            //        set_error(GlobalMethod.GetMessage("W10701", ""));
            //    }
            //}

            //税込
            //契約タブの1.契約情報の消費税率が空ではない場合
            //if (!String.IsNullOrEmpty(item3_1_10.Text))
            //{
            //    //契約金額の税込 
            //    //Decimal keiyakuAmount = Decimal.Parse(item3_1_13.Text.Substring(1, item3_1_13.Text.Length - 1));
            //    Double keiyakuAmount = GetDouble(item3_1_13.Text);

            //    //消費税 
            //    Decimal taxAmount = Decimal.Parse(item3_1_10.Text);
            //    //内消費税　
            //    //Decimal inTaxAmount = Decimal.Parse(item3_1_14.Text.Substring(1, item3_1_14.Text.Length - 1));
            //    Double inTaxAmount = GetDouble(item3_1_14.Text);
            //    //消費税率
            //    Double taxPercent = Double.Parse(item3_1_10.Text);


            //    //契約金額の税込 / (100 + 消費税率))* 消費税率, 0) の小数点切り捨て　amount
            //    Double amount = Math.Floor(keiyakuAmount / (totalHundred + taxPercent) * taxPercent);

            //    //内消費税がamountと一致しない
            //    if (!Double.Equals(inTaxAmount, amount))
            //    {

            //        //※このチェックに入ると、上でエラーがあってもエラーフラグが0：正常となる
            //        //GlobalMethod.outputMessage("E10704", "");
            //        // 起案は出来ますが、契約金額(税込)と内消費税が一致しません。確認してください。
            //        set_error(GlobalMethod.GetMessage("E10704", ""));
            //        // 正常
            //        varidateFlag = true;

            //    }
            //}

            //string jutakuTax_Str = item3_1_15.Text;
            //string totalAmount_Str = item3_2_5_1.Text;

            //jutakuTax_Str = jutakuTax_Str.Replace("¥", "");
            //jutakuTax_Str = jutakuTax_Str.Replace(",", "");
            //totalAmount_Str = totalAmount_Str.Replace("¥", "");
            //totalAmount_Str = totalAmount_Str.Replace(",", "");

            ////受託金額（税込）
            ////1.契約情報の受託金額(税込)
            ////Decimal jutakuTax = Decimal.Parse(jutakuTax_Str);
            //Double jutakuTax = GetDouble(item3_1_15.Text);
            ////2.配分情報の配分情報の配分額(税込)の合計
            ////Decimal totalAmount = Decimal.Parse(totalAmount_Str);
            //Double totalAmount = GetDouble(item3_2_5_1.Text);

            ////受託金額（税込）
            ////1.契約情報の受託金額(税込)
            ////Decimal jutakuTax = Decimal.Parse(item3_1_15.Text.Substring(1, item3_1_15.Text.Length - 1));
            //////2.配分情報の配分情報の配分額(税込)の合計
            ////Decimal totalAmount = Decimal.Parse(item3_2_5_1.Text.Substring(1, item3_2_5_1.Text.Length - 1));

            ////　エラーとさせない為にコメントアウト
            ////受託金額(税込)と配分額(税込)の合計が一致しない
            //if (!Double.Equals(jutakuTax, totalAmount))
            //{
            //    //※このチェックに入ると、上でエラーがあってもエラーフラグが0：正常となる
            //    // 起案は出来ますが、受託契約金額と各配分額の合計が一致していません。確認して下さい。
            //    set_error(GlobalMethod.GetMessage("E10705", ""));
            //    // 正常
            //    varidateFlag = true;
            //}


            //受託金額配分（調査部）                
            //2.配分情報の配分額(税込)
            //Decimal chosaTaxAllocation = Decimal.Parse(item3_2_1_1.Text.Substring(1, item3_2_1_1.Text.Length - 1));
            //Double chosaTaxAllocation = GetDouble(item3_2_1_1.Text);
            //2.配分情報の配分額(税抜)
            //Decimal chosaAllocation = Decimal.Parse(item3_2_1_2.Text.Substring(1, item3_2_1_2.Text.Length - 1));
            //Double chosaAllocation = GetDouble(item3_2_1_2.Text);

            //　エラーとさせない為にコメントアウト
            //配分額(税込)が0よりも上で、配分額(税抜)が0の場合
            //if (Decimal.Compare(chosaTaxAllocation, totalZero) > 0 && Decimal.Compare(chosaAllocation, totalZero) == 0)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        エラー
            //        varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10721", "(調査部)"));
            //    }
            //}


            //受託金額配分(事業普及部)                
            //2.配分情報の配分額(税込)
            //Decimal gyomuTaxAllocation = Decimal.Parse(item3_2_2_1.Text.Substring(1, item3_2_2_1.Text.Length - 1));
            //Double gyomuTaxAllocation = GetDouble(item3_2_2_1.Text);
            //2.配分情報の配分額(税抜)
            //Decimal gyomuAllocation = Decimal.Parse(item3_2_2_2.Text.Substring(1, item3_2_2_2.Text.Length - 1));
            //Double gyomuAllocation = GetDouble(item3_2_2_2.Text);

            //　エラーとさせない為にコメントアウト
            ////配分額(税込)が0よりも上で、配分額(税抜)が0の場合
            //if (Decimal.Compare(gyomuTaxAllocation, totalZero) > 0 && Decimal.Compare(gyomuAllocation, totalZero) == 0)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        //エラー
            //        //varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10721", "(事業普及部)"));
            //    }
            //}

            //受託金額配分(情報システム部)            
            //2.配分情報の配分額(税込)
            //Decimal johoTaxAllocation = Decimal.Parse(item3_2_3_1.Text.Substring(1, item3_2_3_1.Text.Length - 1));
            //Double johoTaxAllocation = GetDouble(item3_2_3_1.Text);
            //2.配分情報の配分額(税抜)
            //Decimal johoAllocation = Decimal.Parse(item3_2_3_2.Text.Substring(1, item3_2_3_2.Text.Length - 1));
            //Double johoAllocation = GetDouble(item3_2_3_2.Text);

            //　エラーとさせない為にコメントアウト
            ////配分額(税込)が0よりも上で、配分額(税抜)が0の場合
            //if (Decimal.Compare(johoTaxAllocation, totalZero) > 0 && Decimal.Compare(johoAllocation, totalZero) == 0)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        //エラー
            //        //varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10721", "(情報システム部)"));
            //    }
            //}

            //受託金額配分(総合研究所)          
            //2.配分情報の配分額(税込)
            //Decimal sogoTaxAllocation = Decimal.Parse(item3_2_4_1.Text.Substring(1, item3_2_4_1.Text.Length - 1));
            //Double sogoTaxAllocation = GetDouble(item3_2_4_1.Text);
            //2.配分情報の配分額(税抜)
            //Decimal sogoAllocation = Decimal.Parse(item3_2_4_2.Text.Substring(1, item3_2_4_2.Text.Length - 1));
            //Double sogoAllocation = GetDouble(item3_2_4_2.Text);

            //　エラーとさせない為にコメントアウト
            ////配分額(税込)が0よりも上で、配分額(税抜)が0の場合
            //if (Decimal.Compare(sogoTaxAllocation, totalZero) > 0 && Decimal.Compare(sogoAllocation, totalZero) == 0)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        //エラー
            //        //varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10721", "(総合研究所)"));
            //    }
            //}

            ////契約工期至
            ////日付フォーマット
            //String format = "yyyy/MM/dd";
            ////契約タブの1.契約情報の契約工期至と売上年度が空でない場合
            //if (item3_1_7.CustomFormat == "" && !String.IsNullOrEmpty(item3_1_5.Text))
            //{

            //    //売上年度 +1年 の3月31日
            //    int year = Int32.Parse(item3_1_5.SelectedValue.ToString()) + 1;
            //    String date = year + "/03/31";
            //    //日付型
            //    DateTime nextYear = DateTime.ParseExact(date, format, null);
            //    DateTime keiyaku = DateTime.ParseExact(item3_1_7.Text, format, null);
            //    //MessageBox.Show(date + ",  " + item3_1_7.Text, "確認", MessageBoxButtons.OKCancel);
            //    //売上年度+1/03/31よりも、契約工期の完了日が未来日付の場合エラー
            //    if (nextYear.Date < keiyaku.Date)
            //    {
            //        if (errorKeiyakuFlag == true)
            //        {
            //            //varidateFlag = false;
            //            // 工期完了日が売上年度を超えています。年度をまたぐ場合は、売上年度を工期完了日にあわせてください。
            //            set_error(GlobalMethod.GetMessage("E10706", ""));
            //        }
            //    }
            //}

            //請求書合計額
            //契約タブの1.契約情報の契約金額の税込
            //Decimal keiyakuTax = Decimal.Parse(item3_1_13.Text.Substring(1, item3_1_13.Text.Length - 1));
            //Double keiyakuTax = GetDouble(item3_1_13.Text);
            //契約タブの1.契約情報の契約金額の税込と、6.請求書情報の請求金額の請求合計額が一致していない
            //Decimal seikyuTotal = Decimal.Parse(item3_6_13.Text.Substring(1, item3_6_13.Text.Length - 1));
            //Double seikyuTotal = GetDouble(item3_6_13.Text);
            //　エラーとさせない為にコメントアウト
            //if (Decimal.Compare(keiyakuTax, seikyuTotal) != 0)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        //varidateFlag = true;
            //        set_error(GlobalMethod.GetMessage("E10707", ""));
            //    }
            //}

            //契約タブの1.契約情報の契約金額の税込と、2.配分情報の配分額(税込)の合計額一致していない
            //Decimal haibunTotal = Decimal.Parse(item3_2_5_1.Text.Substring(1, item3_2_5_1.Text.Length - 1));
            //Double haibunTotal = GetDouble(item3_2_5_1.Text);
            //　エラーとさせない為にコメントアウト
            //if (Decimal.Compare(keiyakuTax, seikyuTotal) != 0)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        //varidateFlag = true;
            //        set_error(GlobalMethod.GetMessage("E10720", ""));
            //    }
            //}

            ////契約工期至
            ////契約タブの1.契約情報の契約工期至と契約工期自が空でない
            //if (item3_1_6.CustomFormat != " " && item3_1_7.CustomFormat != " ")
            //{
            //    //日付型

            //    DateTime keiyakuFrom = DateTime.ParseExact(item3_1_6.Text, format, null);
            //    DateTime keiyakuEnd = DateTime.ParseExact(item3_1_7.Text, format, null);
            //    if (keiyakuFrom.Date > keiyakuEnd.Date)
            //    {
            //        if (errorKeiyakuFlag == true)
            //        {
            //            //varidateFlag = false;
            //            set_error(GlobalMethod.GetMessage("E10011", "(契約工期自・至)"));
            //        }
            //    }
            //}


            ////契約タブの6.売上計上情報の工期末日付が空でなく
            //for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //{
            //    if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "")
            //    {
            //        DateTime kokiDate;
            //        if (DateTime.TryParse(c1FlexGrid4[i, 1].ToString(), out kokiDate))
            //        {
            //            //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
            //            if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
            //            {
            //                if (errorKeiyakuFlag == true)
            //                {
            //                    //varidateFlag = false;
            //                    // 工期末日付は契約工期の期間内で設定して下さい。
            //                    set_error(GlobalMethod.GetMessage("E10708", ""));
            //                }
            //                break;
            //            }
            //        }
            //        else
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                // 工期末日付は契約工期の期間内で設定して下さい。
            //                set_error(GlobalMethod.GetMessage("E10708", ""));
            //                break;
            //            }
            //        }
            //    }
            //    if (c1FlexGrid4[i, 9] != null && c1FlexGrid4[i, 9].ToString() != "")
            //    {
            //        DateTime kokiDate;
            //        if (DateTime.TryParse(c1FlexGrid4[i, 9].ToString(), out kokiDate))
            //        {
            //            //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
            //            if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
            //            {
            //                if (errorKeiyakuFlag == true)
            //                {
            //                    //varidateFlag = false;
            //                    // 工期末日付は契約工期の期間内で設定して下さい。
            //                    set_error(GlobalMethod.GetMessage("E10708", ""));
            //                }
            //                break;
            //            }
            //        }
            //        else
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                // 工期末日付は契約工期の期間内で設定して下さい。
            //                set_error(GlobalMethod.GetMessage("E10708", ""));
            //                break;
            //            }
            //        }
            //    }
            //    if (c1FlexGrid4[i, 17] != null && c1FlexGrid4[i, 17].ToString() != "")
            //    {
            //        DateTime kokiDate;
            //        if (DateTime.TryParse(c1FlexGrid4[i, 17].ToString(), out kokiDate))
            //        {
            //            //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
            //            if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
            //            {
            //                if (errorKeiyakuFlag == true)
            //                {
            //                    //varidateFlag = false;
            //                    // 工期末日付は契約工期の期間内で設定して下さい。
            //                    set_error(GlobalMethod.GetMessage("E10708", ""));
            //                }
            //                break;
            //            }
            //        }
            //        else
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                // 工期末日付は契約工期の期間内で設定して下さい。
            //                set_error(GlobalMethod.GetMessage("E10708", ""));
            //                break;
            //            }
            //        }
            //    }
            //    if (c1FlexGrid4[i, 25] != null && c1FlexGrid4[i, 25].ToString() != "")
            //    {
            //        DateTime kokiDate;
            //        if (DateTime.TryParse(c1FlexGrid4[i, 25].ToString(), out kokiDate))
            //        {
            //            //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
            //            if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
            //            {
            //                if (errorKeiyakuFlag == true)
            //                {
            //                    //varidateFlag = false;
            //                    // 工期末日付は契約工期の期間内で設定して下さい。
            //                    set_error(GlobalMethod.GetMessage("E10708", ""));
            //                }
            //                break;
            //            }
            //        }
            //        else
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                // 工期末日付は契約工期の期間内で設定して下さい。
            //                set_error(GlobalMethod.GetMessage("E10708", ""));
            //                break;
            //            }
            //        }
            //    }
            //}

            //売上計上情報
            ////引合タブの7.業務内容の調査部 業務別配分の合計が100
            //chosaTotal = totalHundred;
            //if (Decimal.Compare(chosaTotal, totalHundred) == 0)
            //{
            //    Boolean nullFlag = false;
            //    //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない場合
            //    if ((c1FlexGrid4[2, 1] == null && c1FlexGrid4[2, 2] == null && c1FlexGrid4[2, 3] == null) || (c1FlexGrid4[2, 1] == null || c1FlexGrid4[2, 3] == null))
            //    {
            //        nullFlag = true;
            //        if (errorKeiyakuFlag == true)
            //        {
            //            //varidateFlag = false;
            //            set_error(GlobalMethod.GetMessage("E10715", "(調査部)"));
            //        }
            //    }

            //    //2.配分情報の配分額(税込)の調査部
            //    if (!nullFlag)
            //    {
            //        int haibunTax = GetInt(item3_2_1_1.Text);
            //        //売上計上情報の調査部の計上額
            //        int keijoTotal = 0;
            //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //        {
            //            if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "")
            //            {
            //                keijoTotal += GetInt(c1FlexGrid4[i, 3].ToString());
            //            }
            //        }
            //        //配分額(税込)の調査部と、調査部の計上額の合計が一致しない場合
            //        if (Decimal.Compare(haibunTax, keijoTotal) != 0)
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                //varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(調査部)"));
            //            }
            //        }
            //    }
            //}
            ////else 引合タブの7.業務内容の調査部 業務別配分の合計が100でない
            //else
            //{
            //    ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //    //if (c1FlexGrid4[2, 1] != null || c1FlexGrid4[2, 3] != null)
            //    //{
            //    //    if (errorKeiyakuFlag == true)
            //    //    {
            //    //        //varidateFlag = false;
            //    //        set_error(GlobalMethod.GetMessage("E10716", "(調査部)"));
            //    //    }
            //    //}
            //}


            ////引合タブの7.業務内容の事業普及部の配分率(%)が100
            //Decimal gyomuHaibun = Decimal.Parse(item1_7_1_2_1.Text.Substring(0, item1_7_1_2_1.Text.Length - 1));
            //if (Decimal.Compare(gyomuHaibun, totalHundred) == 0)
            //{
            //    Boolean nullFlag = false;
            //    //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
            //    if ((c1FlexGrid4[2, 4] == null && c1FlexGrid4[2, 5] == null && c1FlexGrid4[2, 6] == null) || (c1FlexGrid4[2, 4] == null || c1FlexGrid4[2, 6] == null))
            //    {
            //        //nullFlag = true;
            //        //if (errorKeiyakuFlag == true)
            //        //{
            //        //    //varidateFlag = false;
            //        //    set_error(GlobalMethod.GetMessage("E10715", "(事業普及部)"));
            //        //}
            //    }

            //    if (!nullFlag)
            //    {
            //        //2.配分情報の配分額(税込)の事業普及部

            //        int haibunTax = GetInt(item3_2_2_1.Text);
            //        //売上計上情報の事業普及部の計上額の合計
            //        int keijoTotal = 0;
            //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //        {
            //            if (c1FlexGrid4[i, 9] != null && c1FlexGrid4[i, 9].ToString() != "")
            //            {
            //                keijoTotal += GetInt(c1FlexGrid4[i, 11].ToString());
            //            }
            //        }
            //        //配分額(税込)の調査部と、事業普及部の計上額の合計が一致しない場合
            //        if (Decimal.Compare(haibunTax, keijoTotal) != 0)
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                //varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(事業普及部)"));
            //            }
            //        }
            //    }
            //}
            ////else 引合タブの7.業務内容の部門配分 事業普及部の配分率(％)が0でない
            //else if (Decimal.Compare(gyomuHaibun, totalZero) != 0)
            //{
            //    ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //    //if (c1FlexGrid4[2, 4] != null || c1FlexGrid4[2, 6] != null)
            //    //{
            //    //    if (errorKeiyakuFlag == true)
            //    //    {
            //    //        //varidateFlag = false;
            //    //        set_error(GlobalMethod.GetMessage("E10716", "(事業普及部)"));
            //    //    }
            //    //}
            //}


            ////引合タブの7.業務内容の情報システム部の配分率(%)が100
            //Decimal systemHaibun = Decimal.Parse(item1_7_1_3_1.Text.Substring(0, item1_7_1_3_1.Text.Length - 1));
            //if (Decimal.Compare(systemHaibun, totalHundred) == 0)
            //{
            //    Boolean nullFlag = false;
            //    //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
            //    if ((c1FlexGrid4[2, 7] == null && c1FlexGrid4[2, 8] == null && c1FlexGrid4[2, 9] == null) || (c1FlexGrid4[2, 7] == null || c1FlexGrid4[2, 9] == null))
            //    {
            //        nullFlag = true;
            //        if (errorKeiyakuFlag == true)
            //        {
            //            //varidateFlag = false;
            //            set_error(GlobalMethod.GetMessage("E10715", "(情報システム部)"));
            //        }
            //    }

            //    if (!nullFlag)
            //    {
            //        //2.配分情報の配分額(税込)の情報システム部
            //        int haibunTax = GetInt(item3_2_3_1.Text);
            //        //売上計上情報の情報システム部の計上額の合計
            //        int keijoTotal = 0;
            //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //        {
            //            if (c1FlexGrid4[i, 17] != null && c1FlexGrid4[i, 17].ToString() != "")
            //            {
            //                keijoTotal += GetInt(c1FlexGrid4[i, 19].ToString());
            //            }
            //        }
            //        //配分額(税込)の情報システム部と、情報システム部の計上額の合計が一致しない場合
            //        if (Decimal.Compare(haibunTax, keijoTotal) != 0)
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                //varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(情報システム部)"));
            //            }
            //        }
            //    }
            //}
            ////else 引合タブの7.業務内容の部門配分 情報システム部の配分率(％)が0でない
            //else if (Decimal.Compare(gyomuHaibun, totalZero) != 0)
            //{
            //    ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //    //if (c1FlexGrid4[2, 7] != null || c1FlexGrid4[2, 9] != null)
            //    //{
            //    //    if (errorKeiyakuFlag == true)
            //    //    {
            //    //        //varidateFlag = false;
            //    //        set_error(GlobalMethod.GetMessage("E10716", "(情報システム部)"));
            //    //    }
            //    //}
            //}

            ////引合タブの7.業務内容の総合研究所の配分率(%)が100
            //Decimal sogoHaibun = Decimal.Parse(item1_7_1_4_1.Text.Substring(0, item1_7_1_4_1.Text.Length - 1));
            //if (Decimal.Compare(sogoHaibun, totalHundred) == 0)
            //{
            //    Boolean nullFlag = false;
            //    //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
            //    if ((c1FlexGrid4[2, 10] == null && c1FlexGrid4[2, 11] == null && c1FlexGrid4[2, 12] == null) || (c1FlexGrid4[2, 10] == null || c1FlexGrid4[2, 12] == null))
            //    {
            //        //nullFlag = true;
            //        //if (errorKeiyakuFlag == true)
            //        //{
            //        //    //varidateFlag = false;
            //        //    set_error(GlobalMethod.GetMessage("E10715", "(総合研究所)"));
            //        //}
            //    }

            //    if (!nullFlag)
            //    {
            //        //2.配分情報の配分額(税込)の総合研究所
            //        int haibunTax = GetInt(item3_2_4_1.Text);
            //        //売上計上情報の総合研究所の計上額の合計
            //        int keijoTotal = 0;
            //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //        {
            //            if (c1FlexGrid4[i, 25] != null && c1FlexGrid4[i, 25].ToString() != "")
            //            {
            //                keijoTotal += GetInt(c1FlexGrid4[i, 27].ToString());
            //            }
            //        }
            //        //配分額(税込)の総合研究所と、総合研究所の計上額の合計が一致しない場合
            //        if (Decimal.Compare(haibunTax, keijoTotal) != 0)
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                //varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(総合研究所)"));
            //            }
            //        }
            //    }
            //}
            ////else 引合タブの7.業務内容の部門配分 情報システム部の配分率(％)が0でない
            //else if (Decimal.Compare(gyomuHaibun, totalZero) != 0)
            //{
            //    ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //    //if (c1FlexGrid4[2, 10] != null || c1FlexGrid4[2, 12] != null)
            //    //{
            //    //    if (errorKeiyakuFlag == true)
            //    //    {
            //    //        //varidateFlag = false;
            //    //        set_error(GlobalMethod.GetMessage("E10716", "(総合研究所)"));
            //    //    }
            //    //}
            //}

            //売上計上情報
            //引合タブの7.業務内容の調査部 業務別配分の合計が100
            //chosaTotal = totalHundred;
            //if (GetDouble(item1_7_2_13_1.Text).ToString("F2") == "100.00")
            //{
            //    // データ空フラグ true:未入力 false;データあり
            //    //Boolean nullFlag = false;
            //    Boolean nullFlag = true;
            //    //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない場合
            //    //if ((c1FlexGrid4[2, 1] == null && c1FlexGrid4[2, 2] == null && c1FlexGrid4[2, 3] == null) || (c1FlexGrid4[2, 1] == null || c1FlexGrid4[2, 3] == null))
            //    //{
            //    //    nullFlag = true;
            //    //    if (errorKeiyakuFlag == true)
            //    //    {
            //    //        //varidateFlag = false;
            //    //        // 売上計上情報の工期日付か売上計上額が未入力です。
            //    //        set_error(GlobalMethod.GetMessage("E10715", "(調査部)"));
            //    //    }
            //    //}

            //    //契約タブの6.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない場合
            //    // headerで2行使っている
            //    for (int i = 2;i < c1FlexGrid4.Rows.Count - 2; i++)
            //    {
            //        // 計上日、計上月、計上額が空の場合
            //        if ((c1FlexGrid4.Rows[i][1] != null && c1FlexGrid4.Rows[i][1] != "")
            //            //|| (c1FlexGrid4.Rows[i][2] != null && c1FlexGrid4.Rows[i][2] != "")
            //            || (c1FlexGrid4.Rows[i][3] != null && c1FlexGrid4.Rows[i][3].ToString() != "0"))
            //        {
            //            nullFlag = false;
            //            break;
            //        }
            //    }

            //    // 売上計上情報の調査部が全部未入力だった場合
            //    if(nullFlag == true)
            //    {
            //        if (errorKeiyakuFlag == true)
            //        {
            //            // 売上計上情報の工期日付か売上計上額が未入力です。
            //            set_error(GlobalMethod.GetMessage("E10715", "(調査部)"));
            //        }
            //    }

            //    //2.配分情報の配分額(税込)の調査部
            //    if (!nullFlag)
            //    {
            //        long haibunTax = GetLong(item3_2_1_1.Text);
            //        //売上計上情報の調査部の計上額
            //        long keijoTotal = 0;
            //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //        {
            //            // 金額はNULLがあり得るので除外
            //            if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "" && c1FlexGrid4.Rows[i][3] != null)
            //            {
            //                keijoTotal += GetLong(c1FlexGrid4[i, 3].ToString());
            //            }
            //        }
            //        //配分額(税込)の調査部と、調査部の計上額の合計が一致しない場合
            //        if (!long.Equals(haibunTax, keijoTotal))
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                //varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(調査部)"));
            //            }
            //        }
            //    }
            //}
            ////else 引合タブの7.業務内容の調査部 業務別配分の合計が100でない
            //else
            //{
            //    ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //    //if (c1FlexGrid4[2, 1] != null || c1FlexGrid4[2, 3] != null)
            //    //{
            //    //    if (errorKeiyakuFlag == true)
            //    //    {
            //    //        //varidateFlag = false;
            //    //        set_error(GlobalMethod.GetMessage("E10716", "(調査部)"));
            //    //    }
            //    //}
            //}


            ////引合タブの7.業務内容の事業普及部の配分率(%)が100
            //Double gyomuHaibun = Double.Parse(item1_7_1_2_1.Text.Substring(0, item1_7_1_2_1.Text.Length - 1));
            //if (Double.Equals(gyomuHaibun, totalHundred))
            //{
            //    //Boolean nullFlag = false;
            //    Boolean nullFlag = true;
            //    //契約タブの6.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない場合
            //    // headerで2行使っている
            //    for (int i = 2; i < c1FlexGrid4.Rows.Count - 2; i++)
            //    {
            //        // 計上日、計上月、計上額が空の場合
            //        if ((c1FlexGrid4.Rows[i][9] != null && c1FlexGrid4.Rows[i][9] != "")
            //            //|| (c1FlexGrid4.Rows[i][10] != null && c1FlexGrid4.Rows[i][10] != "")
            //            || (c1FlexGrid4.Rows[i][11] != null && c1FlexGrid4.Rows[i][11].ToString() != "0"))
            //        {
            //            nullFlag = false;
            //            break;
            //        }
            //    }

            //    // 売上計上情報の事業普及部が全部未入力だった場合
            //    if (nullFlag == true)
            //    {
            //        if (errorKeiyakuFlag == true)
            //        {
            //            // 売上計上情報の工期日付か売上計上額が未入力です。
            //            set_error(GlobalMethod.GetMessage("E10715", "(事業普及部)"));
            //        }
            //    }

            //    if (!nullFlag)
            //    {
            //        //2.配分情報の配分額(税込)の事業普及部

            //        long haibunTax = GetLong(item3_2_2_1.Text);
            //        //売上計上情報の事業普及部の計上額の合計
            //        long keijoTotal = 0;
            //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //        {
            //            if (c1FlexGrid4[i, 9] != null && c1FlexGrid4[i, 9].ToString() != "" && c1FlexGrid4.Rows[i][11] != null)
            //            {
            //                keijoTotal += GetLong(c1FlexGrid4[i, 11].ToString());
            //            }
            //        }
            //        //配分額(税込)の調査部と、事業普及部の計上額の合計が一致しない場合
            //        if (!long.Equals(haibunTax, keijoTotal))
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                //varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(事業普及部)"));
            //            }
            //        }
            //    }
            //}
            //else 引合タブの7.業務内容の部門配分 事業普及部の配分率(％)が0でない
            //else if (!Double.Equals(gyomuHaibun, totalZero))
            //{
            ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //if (c1FlexGrid4[2, 4] != null || c1FlexGrid4[2, 6] != null)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        //varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10716", "(事業普及部)"));
            //    }
            //}
            //}


            ////引合タブの7.業務内容の情報システム部の配分率(%)が100
            //Decimal systemHaibun = Decimal.Parse(item1_7_1_3_1.Text.Substring(0, item1_7_1_3_1.Text.Length - 1));
            //if (Double.Equals(systemHaibun, totalHundred))
            //{
            //    //Boolean nullFlag = false;
            //    Boolean nullFlag = true;
            //    //契約タブの6.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない場合
            //    // headerで2行使っている
            //    for (int i = 2; i < c1FlexGrid4.Rows.Count - 2; i++)
            //    {
            //        // 計上日、計上月、計上額が空の場合
            //        if ((c1FlexGrid4.Rows[i][17] != null && c1FlexGrid4.Rows[i][17] != "")
            //            //|| (c1FlexGrid4.Rows[i][18] != null && c1FlexGrid4.Rows[i][18] != "")
            //            || (c1FlexGrid4.Rows[i][19] != null && c1FlexGrid4.Rows[i][19].ToString() != "0"))
            //        {
            //            nullFlag = false;
            //            break;
            //        }
            //    }

            //    // 売上計上情報の情報システム部が全部未入力だった場合
            //    if (nullFlag == true)
            //    {
            //        if (errorKeiyakuFlag == true)
            //        {
            //            // 売上計上情報の工期日付か売上計上額が未入力です。
            //            set_error(GlobalMethod.GetMessage("E10715", "(情報システム部)"));
            //        }
            //    }

            //    if (!nullFlag)
            //    {
            //        //2.配分情報の配分額(税込)の情報システム部
            //        long haibunTax = GetLong(item3_2_3_1.Text);
            //        //売上計上情報の情報システム部の計上額の合計
            //        long keijoTotal = 0;
            //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //        {
            //            if (c1FlexGrid4[i, 17] != null && c1FlexGrid4[i, 17].ToString() != "" && c1FlexGrid4.Rows[i][19] != null)
            //            {
            //                keijoTotal += GetLong(c1FlexGrid4[i, 19].ToString());
            //            }
            //        }
            //        //配分額(税込)の情報システム部と、情報システム部の計上額の合計が一致しない場合
            //        if (!long.Equals(haibunTax, keijoTotal))
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                //varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(情報システム部)"));
            //            }
            //        }
            //    }
            //}
            //else 引合タブの7.業務内容の部門配分 情報システム部の配分率(％)が0でない
            //else if (!Double.Equals(gyomuHaibun, totalZero))
            //{
            ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //if (c1FlexGrid4[2, 7] != null || c1FlexGrid4[2, 9] != null)
            //{
            //    if (errorKeiyakuFlag == true)
            //    {
            //        //varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10716", "(情報システム部)"));
            //    }
            //}
            //}

            //引合タブの7.業務内容の総合研究所の配分率(%)が100
            //Double sogoHaibun = Double.Parse(item1_7_1_4_1.Text.Substring(0, item1_7_1_4_1.Text.Length - 1));
            //if (Double.Equals(sogoHaibun, totalHundred))
            //{
            //    //Boolean nullFlag = false;
            //    Boolean nullFlag = true;
            //    //契約タブの6.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない場合
            //    // headerで2行使っている
            //    for (int i = 2; i < c1FlexGrid4.Rows.Count - 2; i++)
            //    {
            //        // 計上日、計上月、計上額が空の場合
            //        if ((c1FlexGrid4.Rows[i][25] != null && c1FlexGrid4.Rows[i][25] != "")
            //            //|| (c1FlexGrid4.Rows[i][26] != null && c1FlexGrid4.Rows[i][26] != "")
            //            || (c1FlexGrid4.Rows[i][27] != null && c1FlexGrid4.Rows[i][27].ToString() != "0"))
            //        {
            //            nullFlag = false;
            //            break;
            //        }
            //    }

            //    // 売上計上情報の総合研究所が全部未入力だった場合
            //    if (nullFlag == true)
            //    {
            //        if (errorKeiyakuFlag == true)
            //        {
            //            // 売上計上情報の工期日付か売上計上額が未入力です。
            //            set_error(GlobalMethod.GetMessage("E10715", "(総合研究所)"));
            //        }
            //    }

            //    if (!nullFlag)
            //    {
            //        //2.配分情報の配分額(税込)の総合研究所
            //        long haibunTax = GetLong(item3_2_4_1.Text);
            //        //売上計上情報の総合研究所の計上額の合計
            //        long keijoTotal = 0;
            //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //        {
            //            if (c1FlexGrid4[i, 25] != null && c1FlexGrid4[i, 25].ToString() != "" && c1FlexGrid4.Rows[i][27] != null)
            //            {
            //                keijoTotal += GetLong(c1FlexGrid4[i, 27].ToString());
            //            }
            //        }
            //        //配分額(税込)の総合研究所と、総合研究所の計上額の合計が一致しない場合
            //        if (!long.Equals(haibunTax, keijoTotal))
            //        {
            //            if (errorKeiyakuFlag == true)
            //            {
            //                //varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(総合研究所)"));
            //            }
            //        }
            //    }
            //}
            ////else 引合タブの7.業務内容の部門配分 情報システム部の配分率(％)が0でない
            //else if (!Double.Equals(gyomuHaibun, totalZero))
            //{
            //    ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //    //if (c1FlexGrid4[2, 10] != null || c1FlexGrid4[2, 12] != null)
            //    //{
            //    //    if (errorKeiyakuFlag == true)
            //    //    {
            //    //        //varidateFlag = false;
            //    //        set_error(GlobalMethod.GetMessage("E10716", "(総合研究所)"));
            //    //    }
            //    //}
            //}
            //}

            // エラーフラグがtrue（エラー）、データチェックフラグがfalse（エラー）
            // 必須チェックフラグがfalse（エラー）
            // となっており、
            // 0:新規登録 1:更新の場合
            // エラーとする
            //if ((ErrorFLG || !varidateFlag || !requiredFlag) && flg != 1 && flg != 2　)
            //if ((ErrorFLG || !varidateFlag || !requiredFlag))
            //{
            //    // エラーでも1:更新、2:チェック用出力の場合、処理は通す
            //    // 0:新規登録 1:更新 2:チェック用出力(赤伝・黒伝) 3:エントリーチェックシート
            //    if (!hikiaiErrorFlg && !nyuusatsuErrorFlg && !KeiyakuErrorFlg && (flg == 1 || flg == 2))
            //    {
            //        // true:正常
            //        return true;
            //    }
            //    // エラー
            //    return false;
            //}
            //else
            //{
            //    // 正常
            //    return true;
            //}
        }

        //private Boolean KianError(int flag = 0)
        //{
        //    //0：チェック用帳票　1：起案処理
        //    Boolean requiredFlag = true;
        //    Boolean varidateFlag = true;
        //    item3_1_1.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_11.BackColor = Color.FromArgb(255, 255, 255);
        //    item3_1_8.BackColor = Color.FromArgb(255, 255, 255);
        //    label335.BackColor = Color.FromArgb(252, 228, 214);
        //    label333.BackColor = Color.FromArgb(252, 228, 214);
        //    label328.BackColor = Color.FromArgb(252, 228, 214);
        //    label325.BackColor = Color.FromArgb(252, 228, 214);
        //    item3_1_17.BackColor = SystemColors.Control;
        //    item1_7.BackColor = SystemColors.Control;

        //    //契約タブ
        //    //受託番号
        //    if (mode == "" && String.IsNullOrEmpty(item1_7.Text))
        //    {
        //        requiredFlag = false;
        //        item1_7.BackColor = Color.FromArgb(255, 204, 255);
        //    }


        //    //契約タブ
        //    //案件区分
        //    if (String.IsNullOrEmpty(item3_1_1.Text))
        //    {
        //        requiredFlag = false;
        //        item3_1_1.BackColor = Color.FromArgb(255, 204, 255);
        //    }
        //    //業務名称	
        //    if (String.IsNullOrEmpty(item3_1_11.Text))
        //    {
        //        requiredFlag = false;
        //        item3_1_11.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //契約区分
        //    if (String.IsNullOrEmpty(item3_1_8.Text))
        //    {
        //        requiredFlag = false;
        //        item3_1_8.BackColor = Color.FromArgb(255, 204, 255);
        //    }
        //    //契約締結(変更)日
        //    if (item3_1_3.CustomFormat == " ")
        //    {
        //        requiredFlag = false;
        //        item3_1_3.BackColor = Color.FromArgb(255, 204, 255);
        //        label335.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //起業日
        //    if (item3_1_4.CustomFormat == " ")
        //    {
        //        requiredFlag = false;
        //        item3_1_4.BackColor = Color.FromArgb(255, 204, 255);
        //        label333.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //契約工期自
        //    if (item3_1_6.CustomFormat == " ")
        //    {
        //        requiredFlag = false;
        //        item3_1_6.BackColor = Color.FromArgb(255, 204, 255);
        //        label328.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //契約工期至
        //    if (item3_1_7.CustomFormat == " ")
        //    {
        //        requiredFlag = false;
        //        item3_1_7.BackColor = Color.FromArgb(255, 204, 255);
        //        label325.BackColor = Color.FromArgb(255, 204, 255);
        //    }
        //    //変更・中止理由 変更伝票時のみチェック
        //    if (mode == "change" && String.IsNullOrEmpty(item3_1_17.Text))
        //    {
        //        requiredFlag = false;
        //        item3_1_17.BackColor = Color.FromArgb(255, 204, 255);
        //    }

        //    //必須項目エラーの出力
        //    if (!requiredFlag)
        //    {
        //        set_error(GlobalMethod.GetMessage("E10010", ""));
        //    }

        //    //契約タブ　売上年度 4桁じゃなかったらエラー
        //    if (4 != item3_1_5.SelectedValue.ToString().Length)
        //    {
        //        varidateFlag = false;
        //        set_error(GlobalMethod.GetMessage("E10011", ""));
        //    }
        //    //引合タブ
        //    if (mode != "change")
        //    {
        //        //入札状況が入札成立でなければ起案エラー
        //        if (!GlobalMethod.GetCommonValue1("NYUUSATSU_SEIRITSU").Equals(item2_1_1.SelectedValue.ToString()))
        //        {
        //            varidateFlag = false;
        //            set_error(GlobalMethod.GetMessage("E10702", ""));
        //        }
        //        //落札者が建設物価調査会でなければ起案エラー
        //        if (!GlobalMethod.GetCommonValue2("ENTORY_TOUKAI").Equals(item2_3_7.Text))
        //        {
        //            varidateFlag = false;
        //            set_error(GlobalMethod.GetMessage("E70048", ""));
        //        }
        //    }
        //    //部門配分
        //    //引合タブ
        //    if (mode != "change")
        //    {
        //        //引合タブの7.業務内容の調査部の合計が0か100でない
        //        if (GetDouble(item1_7_2_13_1.Text).ToString("F2") != "0.00" && GetDouble(item1_7_2_13_1.Text).ToString("F2") != "100.00")
        //        {
        //            varidateFlag = false;
        //            set_error(GlobalMethod.GetMessage("E10703", "(調査部)"));

        //        }

        //        //引合タブの7.業務内容の総合研究所の合計が0か100でない　
        //        if (GetDouble(item1_7_1_4_1.Text).ToString("F2") != "0.00" && GetDouble(item1_7_1_4_1.Text).ToString("F2") != "100.00")
        //        {
        //            varidateFlag = false;
        //            set_error(GlobalMethod.GetMessage("E10703", "(総合研究所)"));

        //        }
        //    }


        //    //契約タブ
        //    //契約金額の税込
        //    //契約タブの1.契約情報の契約金額の税込が0円の場合
        //    String item13 = item3_1_13.Text.Substring(1, item3_1_13.Text.Length - 1);
        //    if (GetInt(item3_1_13.Text) == 0)
        //    {
        //        set_error(GlobalMethod.GetMessage("W10701", ""));
        //    }

        //    //税込
        //    //契約タブの1.契約情報の消費税率が空ではない場合
        //    if (!String.IsNullOrEmpty(item3_1_10.Text))
        //    {
        //        //契約金額の税込 
        //        Decimal keiyakuAmount = GetInt(item3_1_13.Text);

        //        //消費税 
        //        Decimal taxAmount = GetInt(item3_1_10.Text);
        //        //内消費税　
        //        Decimal inTaxAmount = GetInt(item3_1_14.Text);
        //        //消費税率
        //        Decimal taxPercent = GetInt(item3_1_10.Text);


        //        //契約金額の税込 / (100 + 消費税率))* 消費税率, 0) の小数点切り捨て　amount
        //        Decimal amount = Math.Floor(keiyakuAmount * taxPercent / (100 + taxPercent));

        //        //内消費税がamountと一致しない
        //        if (Decimal.Compare(inTaxAmount, amount) != 0)
        //        {

        //            //※このチェックに入ると、上でエラーがあってもエラーフラグが0：正常となる
        //            GlobalMethod.outputMessage("E10704", "");
        //            // エラー
        //            varidateFlag = true;

        //        }
        //    }
        //    //受託金額（税込）
        //    //1.契約情報の受託金額(税込)
        //    Decimal jutakuTax = Decimal.Parse(item3_1_15.Text.Substring(1, item3_1_15.Text.Length - 1));
        //    //2.配分情報の配分情報の配分額(税込)の合計
        //    Decimal totalAmount = Decimal.Parse(item3_2_5_1.Text.Substring(1, item3_2_5_1.Text.Length - 1));

        //    //受託金額(税込)と配分額(税込)の合計が一致しない
        //    if (Decimal.Compare(jutakuTax, totalAmount) != 0)
        //    {
        //        //※このチェックに入ると、上でエラーがあってもエラーフラグが0：正常となる
        //        set_error(GlobalMethod.GetMessage("E10705", ""));
        //        // 正常にする
        //        varidateFlag = true;
        //    }
        //    //受託金額配分（調査部)
        //    //配分額(税込)が0よりも上で、配分額(税抜)が0の場合
        //    if (GetInt(item3_2_1_1.Text) > 0 && GetInt(item3_2_1_2.Text) == 0)
        //    {
        //        //エラー
        //        varidateFlag = false;
        //        set_error(GlobalMethod.GetMessage("E10721", "(調査部)"));
        //    }


        //    //受託金額配分(事業普及部) 
        //    //配分額(税込)が0よりも上で、配分額(税抜)が0の場合
        //    if (GetInt(item3_2_2_1.Text) > 0 && GetInt(item3_2_2_2.Text) == 0)
        //    {
        //        //エラー
        //        varidateFlag = false;
        //        set_error(GlobalMethod.GetMessage("E10721", "(事業普及部)"));
        //    }

        //    //受託金額配分(情報システム部) 

        //    //配分額(税込)が0よりも上で、配分額(税抜)が0の場合
        //    if (GetInt(item3_2_3_1.Text) > 0 && GetInt(item3_2_3_2.Text) == 0)
        //    {
        //        //エラー
        //        varidateFlag = false;
        //        set_error(GlobalMethod.GetMessage("E10721", "(情報システム部)"));
        //    }

        //    //受託金額配分(総合研究所) 
        //    //配分額(税込)が0よりも上で、配分額(税抜)が0の場合
        //    if (GetInt(item3_2_4_1.Text) > 0 && GetInt(item3_2_4_2.Text) == 0)
        //    {
        //        //エラー
        //        varidateFlag = false;
        //        set_error(GlobalMethod.GetMessage("E10721", "(総合研究所)"));
        //    }

        //    //契約工期至
        //    //日付フォーマット
        //    String format = "yyyy/MM/dd";
        //    //契約タブの1.契約情報の契約工期至と売上年度が空でない場合
        //    if (item3_1_7.CustomFormat == "" && !String.IsNullOrEmpty(item3_1_5.Text))
        //    {

        //        //売上年度 +1年 の3月31日
        //        int year = Int32.Parse(item3_1_5.SelectedValue.ToString()) + 1;
        //        String date = year + "/03/31";
        //        //日付型
        //        DateTime nextYear = DateTime.ParseExact(date, format, null);
        //        DateTime keiyaku = DateTime.ParseExact(item3_1_7.Text, format, null);
        //        //MessageBox.Show(date + ",  " + item3_1_7.Text, "確認", MessageBoxButtons.OKCancel);
        //        //売上年度+1/03/31よりも、契約工期の完了日が未来日付の場合エラー
        //        if (nextYear.Date < keiyaku.Date)
        //        {
        //            varidateFlag = false;
        //            set_error(GlobalMethod.GetMessage("E10706", ""));
        //        }
        //    }

        //    //請求書合計額
        //    //契約タブの1.契約情報の契約金額の税込と、6.請求書情報の請求金額の請求合計額が一致していない
        //    if (GetInt(item3_1_13.Text) != GetInt(item3_6_13.Text))
        //    {
        //        // 正常とする
        //        varidateFlag = true;
        //        set_error(GlobalMethod.GetMessage("E10707", ""));
        //    }

        //    //契約タブの1.契約情報の契約金額の税込と、2.配分情報の配分額(税込)の合計額一致していない
        //    if (GetInt(item3_1_13.Text) != GetInt(item3_2_5_1.Text))
        //    {
        //        // 正常とする
        //        varidateFlag = true;
        //        set_error(GlobalMethod.GetMessage("E10720", ""));
        //    }

        //    //契約工期至
        //    //契約タブの1.契約情報の契約工期至と契約工期自が空でない
        //    if (item3_1_6.CustomFormat != " " && item3_1_7.CustomFormat != " ")
        //    {
        //        if (item3_1_6.Value > item3_1_7.Value)
        //        {
        //            varidateFlag = false;
        //            set_error(GlobalMethod.GetMessage("E10011", "(契約工期自・至)"));
        //        }
        //    }

        //    //契約タブの6.売上計上情報の工期末日付が空でなく
        //    for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //    {
        //        if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "")
        //        {
        //            DateTime kokiDate;
        //            if (DateTime.TryParse(c1FlexGrid4[i, 1].ToString(), out kokiDate))
        //            {
        //                //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
        //                if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
        //                {
        //                    varidateFlag = false;
        //                    // 工期末日付は契約工期の期間内で設定して下さい。
        //                    set_error(GlobalMethod.GetMessage("E10708", ""));
        //                    break;
        //                }
        //            }
        //            else
        //            {
        //                // 工期末日付は契約工期の期間内で設定して下さい。
        //                set_error(GlobalMethod.GetMessage("E10708", ""));
        //                break;
        //            }
        //        }
        //        if (c1FlexGrid4[i, 9] != null && c1FlexGrid4[i, 9].ToString() != "")
        //        {
        //            DateTime kokiDate;
        //            if (DateTime.TryParse(c1FlexGrid4[i, 9].ToString(), out kokiDate))
        //            {
        //                //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
        //                if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
        //                {
        //                    varidateFlag = false;
        //                    // 工期末日付は契約工期の期間内で設定して下さい。
        //                    set_error(GlobalMethod.GetMessage("E10708", ""));
        //                    break;
        //                }
        //            }
        //            else
        //            {
        //                // 工期末日付は契約工期の期間内で設定して下さい。
        //                set_error(GlobalMethod.GetMessage("E10708", ""));
        //                break;
        //            }
        //        }
        //        if (c1FlexGrid4[i, 17] != null && c1FlexGrid4[i, 17].ToString() != "")
        //        {
        //            DateTime kokiDate;
        //            if (DateTime.TryParse(c1FlexGrid4[i, 17].ToString(), out kokiDate))
        //            {
        //                //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
        //                if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
        //                {
        //                    varidateFlag = false;
        //                    // 工期末日付は契約工期の期間内で設定して下さい。
        //                    set_error(GlobalMethod.GetMessage("E10708", ""));
        //                    break;
        //                }
        //            }
        //            else
        //            {
        //                // 工期末日付は契約工期の期間内で設定して下さい。
        //                set_error(GlobalMethod.GetMessage("E10708", ""));
        //                break;
        //            }
        //        }
        //        if (c1FlexGrid4[i, 25] != null && c1FlexGrid4[i, 25].ToString() != "")
        //        {
        //            DateTime kokiDate;
        //            if (DateTime.TryParse(c1FlexGrid4[i, 25].ToString(), out kokiDate))
        //            {
        //                //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
        //                if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
        //                {
        //                    varidateFlag = false;
        //                    // 工期末日付は契約工期の期間内で設定して下さい。
        //                    set_error(GlobalMethod.GetMessage("E10708", ""));
        //                    break;
        //                }
        //            }
        //            else
        //            {
        //                // 工期末日付は契約工期の期間内で設定して下さい。
        //                set_error(GlobalMethod.GetMessage("E10708", ""));
        //                break;
        //            }
        //        }
        //    }

        //    //売上計上情報
        //    if (mode != "change")
        //    {
        //        //引合タブの7.業務内容の調査部 業務別配分の合計が100
        //        if (GetDouble(item1_7_2_13_1.Text).ToString("F2") == "100.00")
        //        {
        //            Boolean nullFlag = false;
        //            //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない場合
        //            if ((c1FlexGrid4[2, 1] == null && c1FlexGrid4[2, 2] == null && c1FlexGrid4[2, 3] == null) || (c1FlexGrid4[2, 1] == null || c1FlexGrid4[2, 3] == null))
        //            {
        //                nullFlag = true;
        //                varidateFlag = false;
        //                set_error(GlobalMethod.GetMessage("E10715", "(調査部)"));
        //            }

        //            //2.配分情報の配分額(税込)の調査部
        //            if (!nullFlag)
        //            {
        //                long haibunTax = GetInt(item3_2_1_1.Text);
        //                //売上計上情報の調査部の計上額
        //                long keijoTotal = 0;
        //                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //                {
        //                    if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "")
        //                    {
        //                        keijoTotal += GetInt(c1FlexGrid4[i, 3].ToString());
        //                    }
        //                }
        //                //配分額(税込)の調査部と、調査部の計上額の合計が一致しない場合
        //                if (Decimal.Compare(haibunTax, keijoTotal) != 0)
        //                {
        //                    varidateFlag = false;
        //                    set_error(GlobalMethod.GetMessage("E10717", "(調査部)"));
        //                }
        //            }
        //        }
        //        //引合タブの7.業務内容の調査部 業務別配分の合計が100でない
        //        else
        //        {
        //            //契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
        //            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //            {
        //                if ((c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "") || (c1FlexGrid4[i, 3] != null && c1FlexGrid4[i, 3].ToString() != "" && GetInt(c1FlexGrid4[i, 3].ToString()) != 0))
        //                {
        //                    varidateFlag = false;
        //                    set_error(GlobalMethod.GetMessage("E10716", "(調査部)"));
        //                    break;
        //                }
        //            }
        //        }


        //        //引合タブの7.業務内容の事業普及部の配分率(%)が100
        //        if (GetDouble(item1_7_1_2_1.Text) > 0)
        //        {
        //            Boolean nullFlag = false;
        //            //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
        //            if ((c1FlexGrid4[2, 9] == null && c1FlexGrid4[2, 10] == null && c1FlexGrid4[2, 11] == null) || (c1FlexGrid4[2, 9] == null || c1FlexGrid4[2, 11] == null))
        //            {
        //                nullFlag = true;
        //                varidateFlag = false;
        //                set_error(GlobalMethod.GetMessage("E10715", "(事業普及部)"));
        //            }

        //            if (!nullFlag)
        //            {
        //                //2.配分情報の配分額(税込)の事業普及部
        //                long haibunTax = GetInt(item3_2_2_1.Text);
        //                //売上計上情報の事業普及部の計上額の合計
        //                long keijoTotal = 0;
        //                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //                {
        //                    if (c1FlexGrid4[i, 9] != null && c1FlexGrid4[i, 9].ToString() != "")
        //                    {
        //                        keijoTotal += GetInt(c1FlexGrid4[i, 11].ToString());
        //                    }
        //                }
        //                //配分額(税込)の調査部と、事業普及部の計上額の合計が一致しない場合
        //                if (Decimal.Compare(haibunTax, keijoTotal) != 0)
        //                {
        //                    varidateFlag = false;
        //                    set_error(GlobalMethod.GetMessage("E10717", "(事業普及部)"));
        //                }
        //            }
        //        }
        //        //else 引合タブの7.業務内容の部門配分 事業普及部の配分率(％)が0でない
        //        else
        //        {
        //            //契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
        //            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //            {
        //                if ((c1FlexGrid4[i, 9] != null && c1FlexGrid4[i, 9].ToString() != "") || (c1FlexGrid4[i, 11] != null && c1FlexGrid4[i, 11].ToString() != "" && GetInt(c1FlexGrid4[i, 11].ToString()) != 0))
        //                {
        //                    varidateFlag = false;
        //                    set_error(GlobalMethod.GetMessage("E10716", "(事業普及部)"));
        //                    break;
        //                }
        //            }
        //        }


        //        //引合タブの7.業務内容の情報システム部の配分率(%)が100
        //        if (GetDouble(item1_7_1_3_1.Text) > 0)
        //        {
        //            Boolean nullFlag = false;
        //            //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
        //            if ((c1FlexGrid4[2, 17] == null && c1FlexGrid4[2, 18] == null && c1FlexGrid4[2, 19] == null) || (c1FlexGrid4[2, 17] == null || c1FlexGrid4[2, 19] == null))
        //            {
        //                nullFlag = true;
        //                varidateFlag = false;
        //                set_error(GlobalMethod.GetMessage("E10715", "(情報システム部)"));
        //            }

        //            if (!nullFlag)
        //            {
        //                //2.配分情報の配分額(税込)の情報システム部
        //                long haibunTax = GetInt(item3_2_3_1.Text);
        //                //売上計上情報の事業普及部の計上額の合計
        //                long keijoTotal = 0;
        //                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //                {
        //                    if (c1FlexGrid4[i, 17] != null && c1FlexGrid4[i, 17].ToString() != "")
        //                    {
        //                        keijoTotal += GetInt(c1FlexGrid4[i, 19].ToString());
        //                    }
        //                }
        //                //配分額(税込)の情報システム部と、情報システム部の計上額の合計が一致しない場合
        //                if (Decimal.Compare(haibunTax, keijoTotal) != 0)
        //                {
        //                    varidateFlag = false;
        //                    set_error(GlobalMethod.GetMessage("E10717", "(情報システム部)"));
        //                }
        //            }
        //        }
        //        //else 引合タブの7.業務内容の部門配分 情報システム部の配分率(％)が0でない
        //        else
        //        {
        //            //契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
        //            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //            {
        //                if ((c1FlexGrid4[i, 17] != null && c1FlexGrid4[i, 17].ToString() != "") || (c1FlexGrid4[i, 19] != null && c1FlexGrid4[i, 19].ToString() != "" && GetInt(c1FlexGrid4[i, 19].ToString()) != 0))
        //                {
        //                    varidateFlag = false;
        //                    set_error(GlobalMethod.GetMessage("E10716", "(情報システム部)"));
        //                    break;
        //                }
        //            }
        //        }

        //        //引合タブの7.業務内容の総合研究所の配分率(%)が100
        //        if (GetDouble(item1_7_1_4_1.Text) > 0)
        //        {
        //            Boolean nullFlag = false;
        //            //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
        //            if ((c1FlexGrid4[2, 25] == null && c1FlexGrid4[2, 26] == null && c1FlexGrid4[2, 27] == null) || (c1FlexGrid4[2, 25] == null || c1FlexGrid4[2, 27] == null))
        //            {
        //                nullFlag = true;
        //                varidateFlag = false;
        //                set_error(GlobalMethod.GetMessage("E10715", "(総合研究所)"));
        //            }

        //            if (!nullFlag)
        //            {
        //                //2.配分情報の配分額(税込)の総合研究所
        //                long haibunTax = GetInt(item3_2_4_1.Text);
        //                //売上計上情報の事業普及部の計上額の合計
        //                long keijoTotal = 0;
        //                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //                {
        //                    if (c1FlexGrid4[i, 25] != null && c1FlexGrid4[i, 25].ToString() != "")
        //                    {
        //                        keijoTotal += GetInt(c1FlexGrid4[i, 27].ToString());
        //                    }
        //                }
        //                //配分額(税込)の総合研究所と、総合研究所の計上額の合計が一致しない場合
        //                if (Decimal.Compare(haibunTax, keijoTotal) != 0)
        //                {
        //                    varidateFlag = false;
        //                    set_error(GlobalMethod.GetMessage("E10717", "(総合研究所)"));
        //                }
        //            }
        //        }
        //        //else 引合タブの7.業務内容の部門配分 情報システム部の配分率(％)が0でない
        //        else
        //        {
        //            //契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
        //            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
        //            {
        //                if ((c1FlexGrid4[i, 25] != null && c1FlexGrid4[i, 25].ToString() != "") || (c1FlexGrid4[i, 27] != null && c1FlexGrid4[i, 27].ToString() != "" && GetInt(c1FlexGrid4[i, 27].ToString()) != 0))
        //                {
        //                    varidateFlag = false;
        //                    set_error(GlobalMethod.GetMessage("E10716", "(総合研究所)"));
        //                    break;
        //                }
        //            }
        //        }
        //    }


        //    //契約タブの調査部配分率が0ではない場合、調査部 業務別配分が100でないとエラー
        //    if (GetDouble(item3_7_1_1_1.Text) > 0 && item3_7_2_26_1.Text != "100.00%")
        //    {
        //        set_error(GlobalMethod.GetMessage("E70045", "契約タブ"));
        //        varidateFlag = false;
        //    }

        //    if (flag != 0 && (!requiredFlag || !varidateFlag))
        //    {
        //        return false;
        //    }
        //    return true;
        //}

        // 引合必須チェック
        private Boolean hikiaiRequireCheck()
        {
            // エラーフラグ true:エラー /false:正常
            Boolean errorFlg = false;
            //引合タブ　2.基本情報　
            //売上年度
            if (String.IsNullOrEmpty(item1_3.Text))
            {
                errorFlg = true;
                item1_3.BackColor = Color.FromArgb(255, 204, 255);
            }
            // 工期開始年度
            if (String.IsNullOrEmpty(item1_2_KoukiNendo.Text))
            {
                errorFlg = true;
                item1_2_KoukiNendo.BackColor = Color.FromArgb(255, 204, 255);
            }
            //登録日
            if (item1_9.CustomFormat != "")
            {
                errorFlg = true;
                item1_9.BackColor = Color.FromArgb(255, 204, 255);
                label77.BackColor = Color.FromArgb(255, 204, 255);
            }

            //受託課所支部
            if (String.IsNullOrEmpty(item1_10.Text))
            {
                errorFlg = true;
                item1_10.BackColor = Color.FromArgb(255, 204, 255);
            }

            //契約担当者
            if (String.IsNullOrEmpty(item1_11.Text))
            {
                errorFlg = true;
                item1_11.BackColor = Color.FromArgb(255, 204, 255);
            }

            //案件(受託)フォルダ
            if (String.IsNullOrEmpty(item1_12.Text))
            {
                errorFlg = true;
                item1_12.BackColor = Color.FromArgb(255, 204, 255);
            }

            //引合タブ　3.案件情報
            //業務名称	
            if (String.IsNullOrEmpty(item1_13.Text))
            {
                errorFlg = true;
                item1_13.BackColor = Color.FromArgb(255, 204, 255);
            }

            //契約区分
            if (String.IsNullOrEmpty(item1_14.Text))
            {
                errorFlg = true;
                item1_14.BackColor = Color.FromArgb(255, 204, 255);
            }

            //えんとり君修正STEP2　ご指摘：1392
            //入札方式
            if (String.IsNullOrEmpty(item1_15.Text))
            {
                errorFlg = true;
                item1_15.BackColor = Color.FromArgb(255, 204, 255);
            }
            // 538対応 入札状況は編集不可項目になったので、入札（予定）日のみチェック
            //入札状況が、1:入札前以外の場合、
            //if (item1_17.SelectedValue.ToString() != "1")
            //{
            //    //入札(予定)日
            //    if (item1_16.CustomFormat == " ")
            //    {
            //        errorFlg = false;
            //        item1_16.BackColor = Color.FromArgb(255, 204, 255);
            //        label22.BackColor = Color.FromArgb(255, 204, 255);
            //    }
            //}
            //入札(予定)日
            if (item1_16.CustomFormat == " ")
            {
                errorFlg = true;
                item1_16.BackColor = Color.FromArgb(255, 204, 255);
                label22.BackColor = Color.FromArgb(255, 204, 255);
            }

            //引合タブ　4.発注者情報
            //発注者コード
            if (String.IsNullOrEmpty(item1_19.Text))
            {
                errorFlg = true;
                item1_19.BackColor = Color.FromArgb(255, 204, 255);
            }

            //発注者区分1
            if (String.IsNullOrEmpty(item1_20.Text))
            {
                errorFlg = true;
                item1_20.BackColor = Color.FromArgb(255, 204, 255);
            }

            //発注者区分2
            if (String.IsNullOrEmpty(item1_21.Text))
            {
                errorFlg = true;
                item1_21.BackColor = Color.FromArgb(255, 204, 255);
            }

            //引合タブ　6.当会対応 
            //参考見積額(税抜)
            if (String.IsNullOrEmpty(item1_36.Text))
            {
                errorFlg = true;
                item1_36.BackColor = Color.FromArgb(255, 204, 255);
            }

            return errorFlg;
        }

        // 引合データチェック
        private Boolean hikiaiDataCheck()
        {
            // エラーフラグtrue:エラー /false:正常
            Boolean errorFlg = false;

            //契約担当者の部所CDと、受託課所支部（item1_10）が違う(未入力の場合チェックしない)
            if (item1_11.Text != "" && !item1_11_Busho.Text.Equals(item1_10.SelectedValue))
            {
                set_error(GlobalMethod.GetMessage("W10602", ""));
            }

            //案件受託フォルダのフォーマット: ^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$ が違う
            if (item1_12.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(item1_12.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                set_error(GlobalMethod.GetMessage("E10017", ""));
                errorFlg = true;
            }

            //引合タブ 5.発注担当者情報
            //メールアドレスのフォーマット：が違う
            if (item1_29.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(item1_29.Text, @"^((\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                set_error(GlobalMethod.GetMessage("E10605", ""));
                errorFlg = true;

            }

            // 入札データチェックに移動
            ////引合タブ 7.業務内容
            //Decimal totalZero = Convert.ToDecimal(0);
            //Decimal totalHundred = Convert.ToDecimal(100);

            ////調査部、事業普及部、情報システム部、総合研究所の合計が0か100でない場合
            ////Decimal bumonTotal = Decimal.Parse(item1_7_1_5_1.Text.Substring(0, item1_7_1_5_1.Text.Length - 1));
            //Decimal bumonTotal = GetDecimal(item1_7_1_5_1.Text);
            //if (Decimal.Compare(bumonTotal, totalHundred) != 0 && Decimal.Compare(bumonTotal, totalZero) != 0)
            //{
            //    set_error("部門配分の合計が不正です。");
            //    errorFlg = false;

            //}

            ////// %を除いた値を小数点付きの数値に変換する
            ////Decimal chousabu = Decimal.Parse(item1_7_1_1_1.Text.Substring(0, item1_7_1_1_1.Text.Length - 1));
            ////Decimal chosaTotal = Decimal.Parse(item1_7_2_13_1.Text.Substring(0, item1_7_2_13_1.Text.Length - 1));
            ////// 部門配分の調査部が0ではない場合、
            ////if (Decimal.Compare(chousabu, totalHundred) != 0) {
            ////    // 調査部 業務別配分の合計が0か100でない場合、
            ////    if (Decimal.Compare(chosaTotal, totalHundred) != 0 && Decimal.Compare(chosaTotal, totalZero) != 0)
            ////    {
            ////        // 業務配分の合計が不正です。
            ////        set_error(GlobalMethod.GetMessage("E10604", "(調査部)"));
            ////        ErrorFLG = true;
            ////    }
            ////}

            //// 調査部 配分率が0以上の場合
            //if (GetDouble(item1_7_1_1_1.Text) > 0)
            //{
            //    // 調査部 業務別配分が100でないとエラー
            //    if (item1_7_2_13_1.Text != "100.00%")
            //    {
            //        // 調査業務別　配分の合計が100になるように入力してください。
            //        set_error(GlobalMethod.GetMessage("E70045", ""));
            //        errorFlg = false;
            //    }

            //    Double zeroDouble = GetDouble("0");
            //    // 業務別配分の資材調査、営繕調査、機器類調査、工事費調査のいづれかが1以上でない
            //    if (GetDouble(item1_7_2_1_1.Text) == zeroDouble && GetDouble(item1_7_2_2_1.Text) == zeroDouble && GetDouble(item1_7_2_3_1.Text) == zeroDouble && GetDouble(item1_7_2_4_1.Text) == zeroDouble)
            //    {
            //        // 調査業務別 配分 資材調査、営繕調査、機器類調査、工事費調査のいずれかを入力してください。
            //        set_error(GlobalMethod.GetMessage("E70049", ""));
            //        errorFlg = false;
            //    }
            //}
            return errorFlg;
        }

        // 入札データチェック
        private Boolean nyuusatsuDataCheck()
        {
            // エラーフラグtrue:エラー正常 /false:正常
            Boolean errorFlg = false;

            //引合タブ 7.業務内容
            Decimal totalZero = Convert.ToDecimal(0);
            Decimal totalHundred = Convert.ToDecimal(100);

            //調査部、事業普及部、情報システム部、総合研究所の合計が0か100でない場合
            //Decimal bumonTotal = Decimal.Parse(item1_7_1_5_1.Text.Substring(0, item1_7_1_5_1.Text.Length - 1));
            Decimal bumonTotal = GetDecimal(item1_7_1_5_1.Text);
            if (Decimal.Compare(bumonTotal, totalHundred) != 0 && Decimal.Compare(bumonTotal, totalZero) != 0)
            {
                //set_error("部門配分の合計が不正です。");
                set_error(GlobalMethod.GetMessage("E10604", ""));
                errorFlg = true;
            }

            // 業務別配分の合計が0か100でないとエラー
            Decimal gyoumubetsuTotal = GetDecimal(item1_7_2_13_1.Text);
            if (Decimal.Compare(gyoumubetsuTotal, totalHundred) != 0 && Decimal.Compare(gyoumubetsuTotal, totalZero) != 0)
            {
                // 調査業務別　配分の合計が100になるように入力してください。
                set_error(GlobalMethod.GetMessage("E70045", ""));
                errorFlg = true;
            }

            //// %を除いた値を小数点付きの数値に変換する
            //Decimal chousabu = Decimal.Parse(item1_7_1_1_1.Text.Substring(0, item1_7_1_1_1.Text.Length - 1));
            //Decimal chosaTotal = Decimal.Parse(item1_7_2_13_1.Text.Substring(0, item1_7_2_13_1.Text.Length - 1));
            //// 部門配分の調査部が0ではない場合、
            //if (Decimal.Compare(chousabu, totalHundred) != 0) {
            //    // 調査部 業務別配分の合計が0か100でない場合、
            //    if (Decimal.Compare(chosaTotal, totalHundred) != 0 && Decimal.Compare(chosaTotal, totalZero) != 0)
            //    {
            //        // 業務配分の合計が不正です。
            //        set_error(GlobalMethod.GetMessage("E10604", "(調査部)"));
            //        ErrorFLG = true;
            //    }
            //}

            // 調査部 配分率が0以上の場合
            //if (GetDouble(item1_7_1_1_1.Text) > 0)
            //{
            // 調査部 業務別配分が100でないとエラー
            //if (item1_7_2_13_1.Text != "100.00%" && item1_7_2_13_1.Text != "0.00%")
            //    {
            //        // 調査業務別　配分の合計が100になるように入力してください。
            //        set_error(GlobalMethod.GetMessage("E70045", ""));
            //        errorFlg = false;
            //    }

            // 483 業務別配分の必須チェックを無しとし、横計の100%のチェックのみとする。
            //Double zeroDouble = GetDouble("0");
            //// 業務別配分の資材調査、営繕調査、機器類調査、工事費調査のいづれかが1以上でない
            //if (GetDouble(item1_7_2_1_1.Text) == zeroDouble && GetDouble(item1_7_2_2_1.Text) == zeroDouble && GetDouble(item1_7_2_3_1.Text) == zeroDouble && GetDouble(item1_7_2_4_1.Text) == zeroDouble)
            //{
            //    // 調査業務別 配分 資材調査、営繕調査、機器類調査、工事費調査のいずれかを入力してください。
            //    set_error(GlobalMethod.GetMessage("E70049", ""));
            //    errorFlg = false;
            //}
            //}
            return errorFlg;
        }

        // 契約タブ必須チェック
        private Boolean keiyakuRequireCheck()
        {
            // エラーフラグ true:エラー /false:正常
            Boolean errorFlg = false;
            //契約タブ 1.契約情報緒
            //契約締結(変更)日
            if (item3_1_3.CustomFormat == " ")
            {
                // 入札タブの3.入札結果の落札者の場合
                item3_1_3.BackColor = Color.FromArgb(255, 204, 255);
                label335.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            //起案日
            if (item3_1_4.CustomFormat == " ")
            {
                // 入札タブの3.入札結果の落札者の場合
                item3_1_4.BackColor = Color.FromArgb(255, 204, 255);
                label333.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            //契約工期自
            if (item3_1_6.CustomFormat == " ")
            {
                // 入札タブの3.入札結果の落札者の場合
                item3_1_6.BackColor = Color.FromArgb(255, 204, 255);
                label328.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            //契約工期至
            if (item3_1_7.CustomFormat == " ")
            {
                // 入札タブの3.入札結果の落札者の場合
                item3_1_7.BackColor = Color.FromArgb(255, 204, 255);
                label325.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            //契約金額の税込
            if (String.IsNullOrEmpty(item3_1_13.Text))
            {
                item3_1_13.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            //受託金額(税込)
            if (String.IsNullOrEmpty(item3_1_15.Text))
            {
                item3_1_15.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            //受託外金額(税込)
            if (String.IsNullOrEmpty(item3_1_16.Text))
            {
                item3_1_16.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            //契約タブ 4.管理者・担当者
            ////管理技術者 
            //if (String.IsNullOrEmpty(item3_4_1.Text))
            //{
            //    item3_4_1.BackColor = Color.FromArgb(255, 204, 255);
            //    errorFlg = false;
            //}

            //業務担当者
            if (String.IsNullOrEmpty(item3_4_4.Text))
            {
                item3_4_4.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            //窓口担当者
            if (String.IsNullOrEmpty(item3_4_5.Text))
            {
                item3_4_5.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            ////担当技術者 c1FlexGrid3が2行だったら 1行目ヘッダー
            //if (c1FlexGrid3.Rows.Count < 2)
            //{
            //    c1FlexGrid3.Rows.Add();
            //}
            ////2行目がnullでないことを確認する
            //if (c1FlexGrid3[1, 1] == null || c1FlexGrid3[1, 1].ToString() == "")
            //{
            //    c1FlexGrid3.GetCellRange(1, 1).StyleNew.BackColor = Color.FromArgb(255, 204, 255);
            //    c1FlexGrid3.GetCellRange(1, 2).StyleNew.BackColor = Color.FromArgb(255, 204, 255);
            //    errorFlg = false;
            //}

            //エントリ君修正STEP2
            if (item3_1_2.Checked == false)
            {
                //No.1440 ②調査部　業務配分がある場合にエラーを出す
                if (item3_7_1_1_1.Text != "0.00%")
                {
                    // 調査部 業務別配分が100でないとエラー
                    if (item3_7_2_26_1.Text != "100.00%")
                    {
                        // 調査業務別　配分の合計が100になるように入力してください。
                        item3_7_2_26_1.BackColor = Color.FromArgb(255, 204, 255);
                        label502.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }
                }
                Double total1 = 0;
                Double total2 = 0;
                Double total3 = 0;
                Double total4 = 0;
                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                {
                    if (c1FlexGrid4[i, 3] != null) total1 += GetDouble(c1FlexGrid4[i, 3].ToString());
                    if (c1FlexGrid4[i, 11] != null) total2 += GetDouble(c1FlexGrid4[i, 11].ToString());
                    if (c1FlexGrid4[i, 19] != null) total3 += GetDouble(c1FlexGrid4[i, 19].ToString());
                    if (c1FlexGrid4[i, 27] != null) total4 += GetDouble(c1FlexGrid4[i, 27].ToString());
                }

                //事業部配分の％と事業部の配分金額が異なっていても起案出来てしまう為、エラーとする。
                if (item3_7_1_1_1.Text != "0.00%")
                {
                    if (total1 == 0)
                    {
                        c1FlexGrid4.GetCellRange(2, 3).StyleNew.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }
                    if (GetLong(item3_7_1_6_1.Text) == 0)
                    {
                        item3_2_1_1.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }
                }
                if (item3_7_1_2_1.Text != "0.00%")
                {
                    if (total2 == 0)
                    {
                        c1FlexGrid4.GetCellRange(2, 11).StyleNew.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }

                    if (GetLong(item3_7_1_7_1.Text) == 0)
                    {
                        item3_2_2_1.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }
                }
                if (item3_7_1_3_1.Text != "0.00%")
                {
                    if (total3 == 0)
                    {
                        c1FlexGrid4.GetCellRange(2, 19).StyleNew.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }

                    if (GetLong(item3_7_1_8_1.Text) == 0)
                    {
                        item3_2_3_1.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }
                }
                if (item3_7_1_4_1.Text != "0.00%")
                {
                    if (total4 == 0)
                    {
                        c1FlexGrid4.GetCellRange(2, 27).StyleNew.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }
                    if (GetLong(item3_7_1_9_1.Text) == 0)
                    {
                        item3_2_4_1.BackColor = Color.FromArgb(255, 204, 255);
                        errorFlg = true;
                    }
                }
            }
            return errorFlg;
        }

        // 契約タブデータチェック
        private Boolean keiyakuDataCheck()
        {
            // エラーフラグ true:エラー /false:正常
            Boolean errorFlg = false;

            // 引合状況
            if (item1_1.Text != "発注確定")
            {
                set_error(GlobalMethod.GetMessage("E10723", ""));
                errorFlg = true;
            }

            // 入札状況
            if (item2_1_1.Text != GlobalMethod.GetCommonValue2("NYUUSATSU_SEIRITSU"))
            {
                set_error(GlobalMethod.GetMessage("E10724", ""));
                errorFlg = true;
            }

            // 事業部コード（案件番号の頭文字1つ）がTのとき
            String jigyoCd = Header1.Text.Substring(0, 1);
            if ("T".Equals(jigyoCd))
            {
                // 契約図書が空
                if (String.IsNullOrEmpty(item3_1_26.Text))
                {
                    set_error(GlobalMethod.GetMessage("W10601", ""));
                }

                // 契約図書のフォーマット：^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$　が違う
                if (!System.Text.RegularExpressions.Regex.IsMatch(item3_1_26.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    set_error(GlobalMethod.GetMessage("E10017", ""));
                    //errorFlg = true;
                }
            }

            // 調査部 業務別配分が100でないとエラー
            if (item3_7_2_26_1.Text != "100.00%" && item3_7_2_26_1.Text != "0.00%")
            {
                // 調査業務別　配分の合計が100になるように入力してください。
                set_error(GlobalMethod.GetMessage("E70045", "契約タブ"));
                errorFlg = true;
            }
            return errorFlg;
        }

        // 技術者評価タブデータチェック
        private Boolean gijyutsushahyoukaDataCheck()
        {
            // エラーフラグ true:エラー /false:正常
            Boolean errorFlg = false;

            // 業務評点
            if (item4_1_1.Text != "" && (int.Parse(item4_1_1.Text) < 0 || int.Parse(item4_1_1.Text) > 100))
            {
                set_error(GlobalMethod.GetMessage("E10913", "業務評点"));
                errorFlg = true;
            }

            // 管理技術者評点
            if (item4_1_3.Text != "" && (int.Parse(item4_1_3.Text) < 0 || int.Parse(item4_1_3.Text) > 100))
            {
                set_error(GlobalMethod.GetMessage("E10913", "管理技術者評点"));
                errorFlg = true;
            }

            // 協力担当者評点
            if (item4_1_5.Text != "" && (int.Parse(item4_1_5.Text) < 0 || int.Parse(item4_1_5.Text) > 100))
            {
                set_error(GlobalMethod.GetMessage("E10913", "協力担当者評点"));
                errorFlg = true;
            }

            //請求書のパスのフォーマット：^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$　がちがう
            if (item4_1_8.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(item4_1_8.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                set_error(GlobalMethod.GetMessage("E10017", ""));
                //errorFlg = true;
            }

            return errorFlg;

        }

        // えんとり君修正STEP2：ダミーデータ作成時
        //private Boolean KianError()
        private Boolean KianError(int dummyFlag = 0)
        {
            Boolean requiredFlag = true;
            Boolean varidateFlag = true;
            item3_1_1.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_11.BackColor = Color.FromArgb(255, 255, 255);
            item3_1_8.BackColor = Color.FromArgb(255, 255, 255);
            label335.BackColor = Color.FromArgb(252, 228, 214);
            label333.BackColor = Color.FromArgb(252, 228, 214);
            label328.BackColor = Color.FromArgb(252, 228, 214);
            label325.BackColor = Color.FromArgb(252, 228, 214);
            item3_1_17.BackColor = Color.FromArgb(255, 255, 255);
            //item1_7.BackColor = Color.FromArgb(255, 255, 255);

            set_error("", 0);

            //契約タブ
            // 必須項目エラーとエラーメッセージを別にするため、受託番号のチェック位置を変更する。
            ////受託番号
            ////if (mode == "" && String.IsNullOrEmpty(item1_7.Text))
            //if ((mode == "" || mode == "update") && String.IsNullOrEmpty(item1_7.Text))
            //{
            //    requiredFlag = false;
            //    item1_7.BackColor = Color.FromArgb(255, 204, 255);
            //}

            //set_error("", 0);

            //契約タブ
            //案件区分
            if (String.IsNullOrEmpty(item3_1_1.Text))
            {
                requiredFlag = false;
                item3_1_1.BackColor = Color.FromArgb(255, 204, 255);
            }
            //業務名称	
            if (String.IsNullOrEmpty(item3_1_11.Text))
            {
                requiredFlag = false;
                item3_1_11.BackColor = Color.FromArgb(255, 204, 255);
            }

            //契約区分
            if (String.IsNullOrEmpty(item3_1_8.Text))
            {
                requiredFlag = false;
                item3_1_8.BackColor = Color.FromArgb(255, 204, 255);
            }
            //契約締結(変更)日
            if (item3_1_3.CustomFormat == " ")
            {
                requiredFlag = false;
                item3_1_3.BackColor = Color.FromArgb(255, 204, 255);
                label335.BackColor = Color.FromArgb(255, 204, 255);
            }

            //起業日
            if (item3_1_4.CustomFormat == " ")
            {
                requiredFlag = false;
                item3_1_4.BackColor = Color.FromArgb(255, 204, 255);
                label333.BackColor = Color.FromArgb(255, 204, 255);
            }

            //契約工期自
            if (item3_1_6.CustomFormat == " ")
            {
                requiredFlag = false;
                item3_1_6.BackColor = Color.FromArgb(255, 204, 255);
                label328.BackColor = Color.FromArgb(255, 204, 255);
            }

            //契約工期至
            if (item3_1_7.CustomFormat == " ")
            {
                requiredFlag = false;
                item3_1_7.BackColor = Color.FromArgb(255, 204, 255);
                label325.BackColor = Color.FromArgb(255, 204, 255);
            }
            //変更・中止理由 変更伝票時のみチェック
            if (mode == "change" && String.IsNullOrEmpty(item3_1_17.Text))
            {
                requiredFlag = false;
                item3_1_17.BackColor = Color.FromArgb(255, 204, 255);
            }

            //業務担当者
            if (String.IsNullOrEmpty(item3_4_4.Text))
            {
                requiredFlag = false;
                item3_4_4.BackColor = Color.FromArgb(255, 204, 255);
            }

            //窓口担当者
            if (String.IsNullOrEmpty(item3_4_5.Text))
            {
                requiredFlag = false;
                item3_4_5.BackColor = Color.FromArgb(255, 204, 255);
            }

            //必須項目エラーの出力
            if (!requiredFlag)
            {
                set_error(GlobalMethod.GetMessage("E10010", ""));
            }

            //受託番号
            if ((mode == "" || mode == "update") && String.IsNullOrEmpty(item1_7.Text))
            {
                requiredFlag = false;
                //item1_7.BackColor = Color.FromArgb(255, 204, 255);
                set_error(GlobalMethod.GetMessage("E10722", ""));
            }

            //契約タブ　売上年度 4桁じゃなかったらエラー
            if (4 != item3_1_5.SelectedValue.ToString().Length)
            {
                varidateFlag = false;
                set_error(GlobalMethod.GetMessage("E10011", ""));
            }
            //引合タブ
            if (mode != "change")
            {
                //入札状況が入札成立でなければ起案エラー
                if (!GlobalMethod.GetCommonValue1("NYUUSATSU_SEIRITSU").Equals(item2_1_1.SelectedValue.ToString()))
                {
                    varidateFlag = false;
                    set_error(GlobalMethod.GetMessage("E10702", ""));
                }
                //落札者が建設物価調査会でなければ起案エラー
                if (!GlobalMethod.GetCommonValue2("ENTORY_TOUKAI").Equals(item2_3_7.Text))
                {
                    varidateFlag = false;
                    set_error(GlobalMethod.GetMessage("E70048", ""));
                }
            }
            //部門配分
            Double totalZero = Convert.ToDouble(0);
            Double totalHundred = Convert.ToDouble(100);
            //Double chosaTotal = 0;
            //引合タブ
            //if (mode != "change")
            //{
            // //引合タブの7.業務内容の調査部の合計が0か100でない
            //chosaTotal = Decimal.Parse(item1_7_2_13_1.Text.Substring(0, item1_7_2_13_1.Text.Length - 1));
            // if (Decimal.Compare(chosaTotal, totalHundred) != 0 && Decimal.Compare(chosaTotal, totalZero) != 0)
            // {
            //     varidateFlag = false;
            //     set_error(GlobalMethod.GetMessage("E10703", "(調査部)"));

            // }

            // //引合タブの7.業務内容の総合研究所の合計が0か100でない　
            // Decimal sogoTotal = Decimal.Parse(item1_7_1_4_1.Text.Substring(0, item1_7_1_4_1.Text.Length - 1));
            // if (Decimal.Compare(sogoTotal, totalHundred) != 0 && Decimal.Compare(sogoTotal, totalZero) != 0)
            // {
            //     varidateFlag = false;
            //     set_error(GlobalMethod.GetMessage("E10703", "(総合研究所)"));

            // }
            //}

            // 20210515 エラーのチェックを先に行い、ワーニングを後にする。（エラー時はワーニングのチェックをしないようにする）

            //契約金額の税込
            //契約タブの1.契約情報の契約金額の税込が0円の場合
            long item13 = GetLong(item3_1_13.Text);
            if (item13 == 0)
            {
                // 0円起案です。
                set_error(GlobalMethod.GetMessage("W10701", ""));
                varidateFlag = false;
            }

            //契約工期至
            String format = "yyyy/MM/dd";       //日付フォーマット
            //契約タブの1.契約情報の契約工期至と売上年度が空でない場合
            if (item3_1_7.CustomFormat == "" && !String.IsNullOrEmpty(item3_1_5.Text))
            {
                //売上年度 +1年 の3月31日
                int year = Int32.Parse(item3_1_5.SelectedValue.ToString()) + 1;
                String date = year + "/03/31";
                //日付型
                DateTime nextYear = DateTime.ParseExact(date, format, null);
                DateTime keiyaku = DateTime.ParseExact(item3_1_7.Text, format, null);
                //MessageBox.Show(date + ",  " + item3_1_7.Text, "確認", MessageBoxButtons.OKCancel);
                //売上年度+1/03/31よりも、契約工期の完了日が未来日付の場合エラー
                if (nextYear.Date < keiyaku.Date)
                {
                    varidateFlag = false;
                    set_error(GlobalMethod.GetMessage("E10706", ""));
                }
            }

            //契約タブの1.契約情報の契約工期至と契約工期自が空でない
            if (item3_1_6.CustomFormat != " " && item3_1_7.CustomFormat != " ")
            {
                //日付型

                DateTime keiyakuFrom = DateTime.ParseExact(item3_1_6.Text, format, null);
                DateTime keiyakuEnd = DateTime.ParseExact(item3_1_7.Text, format, null);
                if (keiyakuFrom.Date > keiyakuEnd.Date)
                {
                    varidateFlag = false;
                    set_error(GlobalMethod.GetMessage("E10011", "(契約工期自・至)"));
                }
            }

            //契約タブの6.売上計上情報の工期末日付が空でなく
            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            {
                if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "")
                {
                    DateTime kokiDate;
                    if (DateTime.TryParse(c1FlexGrid4[i, 1].ToString(), out kokiDate))
                    {
                        //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
                        if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
                        {
                            varidateFlag = false;
                            // 工期末日付は契約工期の期間内で設定して下さい。
                            set_error(GlobalMethod.GetMessage("E10708", ""));
                            break;
                        }
                    }
                    else
                    {
                        // 工期末日付は契約工期の期間内で設定して下さい。
                        set_error(GlobalMethod.GetMessage("E10708", ""));
                        break;
                    }
                }
                if (c1FlexGrid4[i, 9] != null && c1FlexGrid4[i, 9].ToString() != "")
                {
                    DateTime kokiDate;
                    if (DateTime.TryParse(c1FlexGrid4[i, 9].ToString(), out kokiDate))
                    {
                        //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
                        if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
                        {
                            varidateFlag = false;
                            // 工期末日付は契約工期の期間内で設定して下さい。
                            set_error(GlobalMethod.GetMessage("E10708", ""));
                            break;
                        }
                    }
                    else
                    {
                        // 工期末日付は契約工期の期間内で設定して下さい。
                        set_error(GlobalMethod.GetMessage("E10708", ""));
                        break;
                    }
                }
                if (c1FlexGrid4[i, 17] != null && c1FlexGrid4[i, 17].ToString() != "")
                {
                    DateTime kokiDate;
                    if (DateTime.TryParse(c1FlexGrid4[i, 17].ToString(), out kokiDate))
                    {
                        //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
                        if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
                        {
                            varidateFlag = false;
                            // 工期末日付は契約工期の期間内で設定して下さい。
                            set_error(GlobalMethod.GetMessage("E10708", ""));
                            break;
                        }
                    }
                    else
                    {
                        // 工期末日付は契約工期の期間内で設定して下さい。
                        set_error(GlobalMethod.GetMessage("E10708", ""));
                        break;
                    }
                }
                if (c1FlexGrid4[i, 25] != null && c1FlexGrid4[i, 25].ToString() != "")
                {
                    DateTime kokiDate;
                    if (DateTime.TryParse(c1FlexGrid4[i, 25].ToString(), out kokiDate))
                    {
                        //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
                        if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
                        {
                            varidateFlag = false;
                            // 工期末日付は契約工期の期間内で設定して下さい。
                            set_error(GlobalMethod.GetMessage("E10708", ""));
                            break;
                        }
                    }
                    else
                    {
                        // 工期末日付は契約工期の期間内で設定して下さい。
                        set_error(GlobalMethod.GetMessage("E10708", ""));
                        break;
                    }
                }
            }

            //契約タブの調査部配分率が0ではない場合
            // 調査部 業務別配分が100でないとエラー
            //if (item3_7_2_26_1.Text != "100.00%" && item3_7_2_26_1.Text != "0.00%")
            if (GetDouble(item3_7_1_1_1.Text) > 0)
            {
                if (item3_7_2_26_1.Text != "100.00%")
                {
                    // 調査業務別　配分の合計が100になるように入力してください。
                    set_error(GlobalMethod.GetMessage("E70045", "(契約タブ)"));
                    varidateFlag = false;
                }
            }
            // 契約タブの調査部配分が0の場合
            // 調査部 業務別配分の合計が0でないとエラー
            else
            {
                if (item3_7_2_26_1.Text != "0.00%")
                {
                    // 調査部　業務別配分の合計が不正です。
                    set_error(GlobalMethod.GetMessage("E10725", "(契約タブ)"));
                    varidateFlag = false;
                }
            }

            //エントリ君修正STEP2
            // 受託金額（税込）
            Double jutakuTax = GetDouble(item3_1_15.Text);      // 1.契約情報の受託金額(税込)
            Double totalAmount = GetDouble(item3_2_5_1.Text);   // 2.配分情報の配分情報の配分額(税込)の合計

            //受託金額(税込)と配分額(税込)の合計が一致しない
            if (!Double.Equals(jutakuTax, totalAmount))
            {
                set_error(GlobalMethod.GetMessage("E10705", ""));
                varidateFlag = false;
            }

            //請求書合計額
            Double keiyakuTax = GetDouble(item3_1_13.Text);     // 1.契約情報の契約金額の税込
            Double seikyuTotal = GetDouble(item3_6_13.Text);    // 6.請求書情報の請求金額の請求合計額

            //契約タブの1.契約情報の契約金額の税込と、6.請求書情報の請求金額の請求合計額が一致していない
            if (!Double.Equals(keiyakuTax, seikyuTotal))
            {
                set_error(GlobalMethod.GetMessage("E10707", ""));
                varidateFlag = false;
            }

            Double haibunTotal = GetDouble(item3_2_5_1.Text);   // 2.配分情報の配分額(税込)の合計額


            //No1440 修正
            ////契約タブの1.契約情報の契約金額の税込と、2.配分情報の配分額(税込)の合計額一致していない
            //if (!Double.Equals(keiyakuTax, haibunTotal))
            //{
            //    set_error(GlobalMethod.GetMessage("E10720", ""));
            //    varidateFlag = false;
            //}

            //No.1440 ・売上情報と契約金額（税込）の合計額が一致しているか？→NG・・・受託金額と売上情報の一致を確認
            // 売上情報と契約金額（税込）の合計額が一致しない
            Double uriageTotal = 0;
            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            {
                if (c1FlexGrid4[i, 3] != null) uriageTotal += GetDouble(c1FlexGrid4[i, 3].ToString());
                if (c1FlexGrid4[i, 11] != null) uriageTotal += GetDouble(c1FlexGrid4[i, 11].ToString());
                if (c1FlexGrid4[i, 19] != null) uriageTotal += GetDouble(c1FlexGrid4[i, 19].ToString());
                if (c1FlexGrid4[i, 27] != null) uriageTotal += GetDouble(c1FlexGrid4[i, 27].ToString());
            }
            ////No.1440 受託金額と売上情報の一致を確認に変更
            if (!Double.Equals(jutakuTax, uriageTotal))
            {
                set_error(GlobalMethod.GetMessage("E10731", ""));
                varidateFlag = false;
            }
            //if (!Double.Equals(keiyakuTax, uriageTotal))
            //{
            //    set_error(GlobalMethod.GetMessage("E10728", ""));
            //    varidateFlag = false;
            //}

            // No1209 要望ほか、６件 アラートからエラーメッセージへ変更する
            if (item2_2_3.SelectedValue != null && item2_2_3.SelectedValue.ToString().Equals("3") == false)
            {
                if ((item2_2_1.SelectedValue != null && (item2_2_1.SelectedValue.ToString().Equals("3") || item2_2_1.SelectedValue.ToString().Equals("4"))) ||
                    (item1_34.SelectedValue != null && item1_34.SelectedValue.ToString().Equals("4")))
                {
                    set_error(GlobalMethod.GetMessage("E10730", "(入札タブ)(引合タブ)"));
                    varidateFlag = false;
                }
            }

            if (requiredFlag && varidateFlag)
            {
                // 税込
                // 契約タブの1.契約情報の消費税率が空ではない場合
                if (!String.IsNullOrEmpty(item3_1_10.Text))
                {
                    Double keiyakuAmount = GetDouble(item3_1_13.Text);  // 契約金額の税込 
                    Double taxAmount = GetDouble(item3_1_10.Text);      // 消費税 
                    Double inTaxAmount = GetDouble(item3_1_14.Text);    // 内消費税
                    Double taxPercent = GetDouble(item3_1_10.Text);     // 消費税率

                    // 契約金額の税込 / (100 + 消費税率))* 消費税率, 0) の小数点切り捨て　amount
                    Double amount = Math.Floor(keiyakuAmount / (totalHundred + taxPercent) * taxPercent);

                    // 内消費税がamountと一致しない
                    if (!Double.Equals(inTaxAmount, amount))
                    {
                        GlobalMethod.outputMessage("E10704", "");
                    }
                }

                //えんとり君修正STEP2
                if(dummyFlag == 0) {
                    //①当会応札（入札タブ）が「対応前」のままで起案しようとしたらアラート表示する。
                    if(item2_2_1.SelectedValue != null && item2_2_1.SelectedValue.ToString().Equals("1"))
                    {
                        GlobalMethod.outputMessage("E10729", "(入札タブ)");
                    }

                    // No1209 要望ほか、６件 アラートからエラーメッセージへ変更する
                    ////②受注意欲（入札タブ）を「なし」以外にした状態で、当会応札（入札タブ）が「不参加」「辞退」、または参考見積（引合タブ）が「辞退」の場合にアラート表示してほしい
                    //if (item2_2_3.SelectedValue != null && item2_2_3.SelectedValue.ToString().Equals("3") == false)
                    //{
                    //    if((item2_2_1.SelectedValue != null && (item2_2_1.SelectedValue.ToString().Equals("3") || item2_2_1.SelectedValue.ToString().Equals("4"))) ||
                    //        (item1_34.SelectedValue != null && item1_34.SelectedValue.ToString().Equals("4")))
                    //    {
                    //        if (GlobalMethod.outputMessage("E10730", "(入札タブ)(引合タブ)",1) == DialogResult.Cancel)
                    //        {
                    //            varidateFlag = false;
                    //        }
                    //    }
                    //}
                }
            }
            //// エラーがなければワーニングのチェックを行う
            //if (requiredFlag && varidateFlag)
            //{
            //    // 税込
            //    // 契約タブの1.契約情報の消費税率が空ではない場合
            //    if (!String.IsNullOrEmpty(item3_1_10.Text))
            //    {
            //        Double keiyakuAmount = GetDouble(item3_1_13.Text);  // 契約金額の税込 
            //        Double taxAmount = GetDouble(item3_1_10.Text);      // 消費税 
            //        Double inTaxAmount = GetDouble(item3_1_14.Text);    // 内消費税
            //        Double taxPercent = GetDouble(item3_1_10.Text);     // 消費税率

            //        // 契約金額の税込 / (100 + 消費税率))* 消費税率, 0) の小数点切り捨て　amount
            //        Double amount = Math.Floor(keiyakuAmount / (totalHundred + taxPercent) * taxPercent);

            //        // 内消費税がamountと一致しない
            //        if (!Double.Equals(inTaxAmount, amount))
            //        {
            //            GlobalMethod.outputMessage("E10704", "");
            //        }
            //    }

            //    // 受託金額（税込）
            //    Double jutakuTax = GetDouble(item3_1_15.Text);      // 1.契約情報の受託金額(税込)
            //    Double totalAmount = GetDouble(item3_2_5_1.Text);   // 2.配分情報の配分情報の配分額(税込)の合計

            //    //受託金額(税込)と配分額(税込)の合計が一致しない
            //    if (!Double.Equals(jutakuTax, totalAmount))
            //    {
            //        set_error(GlobalMethod.GetMessage("E10705", ""));
            //    }

            //    //請求書合計額
            //    Double keiyakuTax = GetDouble(item3_1_13.Text);     // 1.契約情報の契約金額の税込
            //    Double seikyuTotal = GetDouble(item3_6_13.Text);    // 6.請求書情報の請求金額の請求合計額

            //    //契約タブの1.契約情報の契約金額の税込と、6.請求書情報の請求金額の請求合計額が一致していない
            //    if (!Double.Equals(keiyakuTax, seikyuTotal))
            //    {
            //        set_error(GlobalMethod.GetMessage("E10707", ""));
            //    }

            //    Double haibunTotal = GetDouble(item3_2_5_1.Text);   // 2.配分情報の配分額(税込)の合計額

            //    //契約タブの1.契約情報の契約金額の税込と、2.配分情報の配分額(税込)の合計額一致していない
            //    if (!Double.Equals(keiyakuTax, haibunTotal))
            //    {
            //        set_error(GlobalMethod.GetMessage("E10720", ""));
            //    }

            //}


            //契約タブ
            //契約金額の税込
            //契約タブの1.契約情報の契約金額の税込が0円の場合
            //String item13 = item3_1_13.Text.Substring(1, item3_1_13.Text.Length - 1);
            //long item13 = GetLong(item3_1_13.Text);
            //if (item13 == 0)
            //{
            //    // 0円起案です。
            //    set_error(GlobalMethod.GetMessage("W10701", ""));
            //    varidateFlag = false;
            //}

            //税込
            //契約タブの1.契約情報の消費税率が空ではない場合
            //if (!String.IsNullOrEmpty(item3_1_10.Text))
            //{
            //    //契約金額の税込 
            //    //Decimal keiyakuAmount = Decimal.Parse(item3_1_13.Text.Substring(1, item3_1_13.Text.Length - 1));
            //    Double keiyakuAmount = GetDouble(item3_1_13.Text);

            //    //消費税 
            //    //Decimal taxAmount = Decimal.Parse(item3_1_10.Text);
            //    Double taxAmount = GetDouble(item3_1_10.Text);
            //    //内消費税　
            //    //Decimal inTaxAmount = Decimal.Parse(item3_1_14.Text.Substring(1, item3_1_14.Text.Length - 1));
            //    Double inTaxAmount = GetDouble(item3_1_14.Text);
            //    //消費税率
            //    //Decimal taxPercent = Decimal.Parse(item3_1_10.Text);
            //    Double taxPercent = GetDouble(item3_1_10.Text);


            //    //契約金額の税込 / (100 + 消費税率))* 消費税率, 0) の小数点切り捨て　amount
            //    Double amount = Math.Floor(keiyakuAmount / (totalHundred + taxPercent) * taxPercent);

            //    //内消費税がamountと一致しない
            //    if (!Double.Equals(inTaxAmount, amount))
            //    {

            //        //※このチェックに入ると、上でエラーがあってもエラーフラグが0：正常となる
            //        GlobalMethod.outputMessage("E10704", "");
            //        // エラー
            //        //varidateFlag = true;

            //    }
            //}
            ////受託金額（税込）
            ////1.契約情報の受託金額(税込)
            ////Decimal jutakuTax = Decimal.Parse(item3_1_15.Text.Substring(1, item3_1_15.Text.Length - 1));
            //Double jutakuTax = GetDouble(item3_1_15.Text);
            ////2.配分情報の配分情報の配分額(税込)の合計
            ////Decimal totalAmount = Decimal.Parse(item3_2_5_1.Text.Substring(1, item3_2_5_1.Text.Length - 1));
            //Double totalAmount = GetDouble(item3_2_5_1.Text);

            ////受託金額(税込)と配分額(税込)の合計が一致しない
            //if (!Double.Equals(jutakuTax, totalAmount))
            //{
            //    //※このチェックに入ると、上でエラーがあってもエラーフラグが0：正常となる
            //    set_error(GlobalMethod.GetMessage("E10705", ""));
            //    // 正常にする
            //    //varidateFlag = true;
            //}
            //受託金額配分（調査部）                
            //2.配分情報の配分額(税込)
            //Decimal chosaTaxAllocation = Decimal.Parse(item3_2_1_1.Text.Substring(1, item3_2_1_1.Text.Length - 1));
            //Double chosaTaxAllocation = GetDouble(item3_2_1_1.Text);
            //2.配分情報の配分額(税抜)
            //Decimal chosaAllocation = Decimal.Parse(item3_2_1_2.Text.Substring(1, item3_2_1_2.Text.Length - 1));
            //Double chosaAllocation = GetDouble(item3_2_1_2.Text);
            ////配分額(税込)が0よりも上で、配分額(税抜)が0の場合
            //if (Decimal.Compare(chosaTaxAllocation, totalZero) > 0 && Decimal.Compare(chosaAllocation, totalZero) == 0)
            //{
            //    //エラー
            //    varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E10721", "(調査部)"));
            //}


            //受託金額配分(事業普及部)                
            //2.配分情報の配分額(税込)
            //Decimal gyomuTaxAllocation = Decimal.Parse(item3_2_2_1.Text.Substring(1, item3_2_2_1.Text.Length - 1));
            //Double gyomuTaxAllocation = GetDouble(item3_2_2_1.Text);
            //2.配分情報の配分額(税抜)
            //Decimal gyomuAllocation = Decimal.Parse(item3_2_2_2.Text.Substring(1, item3_2_2_2.Text.Length - 1));
            //Double gyomuAllocation = GetDouble(item3_2_2_2.Text);

            ////配分額(税込)が0よりも上で、配分額(税抜)が0の場合
            //if (Decimal.Compare(gyomuTaxAllocation, totalZero) > 0 && Decimal.Compare(gyomuAllocation, totalZero) == 0)
            //{
            //    //エラー
            //    varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E10721", "(事業普及部)"));
            //}

            //受託金額配分(情報システム部)            
            //2.配分情報の配分額(税込)
            //Decimal johoTaxAllocation = Decimal.Parse(item3_2_3_1.Text.Substring(1, item3_2_3_1.Text.Length - 1));
            //Double johoTaxAllocation = GetDouble(item3_2_3_1.Text);
            //2.配分情報の配分額(税抜)
            //Decimal johoAllocation = Decimal.Parse(item3_2_3_2.Text.Substring(1, item3_2_3_2.Text.Length - 1));
            //Double johoAllocation = GetDouble(item3_2_3_2.Text);

            ////配分額(税込)が0よりも上で、配分額(税抜)が0の場合
            //if (Decimal.Compare(johoTaxAllocation, totalZero) > 0 && Decimal.Compare(johoAllocation, totalZero) == 0)
            //{
            //    //エラー
            //    varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E10721", "(情報システム部)"));
            //}

            //受託金額配分(総合研究所)          
            //2.配分情報の配分額(税込)
            //Decimal sogoTaxAllocation = Decimal.Parse(item3_2_4_1.Text.Substring(1, item3_2_4_1.Text.Length - 1));
            //Double sogoTaxAllocation = GetDouble(item3_2_4_1.Text);
            //2.配分情報の配分額(税抜)
            //Decimal sogoAllocation = Decimal.Parse(item3_2_4_2.Text.Substring(1, item3_2_4_2.Text.Length - 1));
            //Double sogoAllocation = GetDouble(item3_2_4_2.Text);

            ////配分額(税込)が0よりも上で、配分額(税抜)が0の場合
            //if (Decimal.Compare(sogoTaxAllocation, totalZero) > 0 && Decimal.Compare(sogoAllocation, totalZero) == 0)
            //{
            //    //エラー
            //    varidateFlag = false;
            //    set_error(GlobalMethod.GetMessage("E10721", "(総合研究所)"));
            //}

            ////契約工期至
            ////日付フォーマット
            //String format = "yyyy/MM/dd";
            ////契約タブの1.契約情報の契約工期至と売上年度が空でない場合
            //if (item3_1_7.CustomFormat == "" && !String.IsNullOrEmpty(item3_1_5.Text))
            //{

            //    //売上年度 +1年 の3月31日
            //    int year = Int32.Parse(item3_1_5.SelectedValue.ToString()) + 1;
            //    String date = year + "/03/31";
            //    //日付型
            //    DateTime nextYear = DateTime.ParseExact(date, format, null);
            //    DateTime keiyaku = DateTime.ParseExact(item3_1_7.Text, format, null);
            //    //MessageBox.Show(date + ",  " + item3_1_7.Text, "確認", MessageBoxButtons.OKCancel);
            //    //売上年度+1/03/31よりも、契約工期の完了日が未来日付の場合エラー
            //    if (nextYear.Date < keiyaku.Date)
            //    {
            //        varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10706", ""));
            //    }
            //}

            ////請求書合計額
            ////契約タブの1.契約情報の契約金額の税込
            ////Decimal keiyakuTax = Decimal.Parse(item3_1_13.Text.Substring(1, item3_1_13.Text.Length - 1));
            //Double keiyakuTax = GetDouble(item3_1_13.Text);
            ////契約タブの1.契約情報の契約金額の税込と、6.請求書情報の請求金額の請求合計額が一致していない
            ////Decimal seikyuTotal = Decimal.Parse(item3_6_13.Text.Substring(1, item3_6_13.Text.Length - 1));
            //Double seikyuTotal = GetDouble(item3_6_13.Text);
            //if (!Double.Equals(keiyakuTax, seikyuTotal))
            //{
            //    // 正常とする
            //    //varidateFlag = true;
            //    set_error(GlobalMethod.GetMessage("E10707", ""));
            //}

            ////契約タブの1.契約情報の契約金額の税込と、2.配分情報の配分額(税込)の合計額一致していない
            ////Decimal haibunTotal = Decimal.Parse(item3_2_5_1.Text.Substring(1, item3_2_5_1.Text.Length - 1));
            //Double haibunTotal = GetDouble(item3_2_5_1.Text);
            //if (!Double.Equals(keiyakuTax, seikyuTotal))
            //{
            //    // 正常とする
            //    //varidateFlag = true;
            //    set_error(GlobalMethod.GetMessage("E10720", ""));
            //}

            ////契約工期至
            ////契約タブの1.契約情報の契約工期至と契約工期自が空でない
            //if (item3_1_6.CustomFormat != " " && item3_1_7.CustomFormat != " ")
            //{
            //    //日付型

            //    DateTime keiyakuFrom = DateTime.ParseExact(item3_1_6.Text, format, null);
            //    DateTime keiyakuEnd = DateTime.ParseExact(item3_1_7.Text, format, null);
            //    if (keiyakuFrom.Date > keiyakuEnd.Date)
            //    {
            //        varidateFlag = false;
            //        set_error(GlobalMethod.GetMessage("E10011", "(契約工期自・至)"));
            //    }
            //}

            ////契約タブの6.売上計上情報の工期末日付が空でなく
            //for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //{
            //    if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "")
            //    {
            //        DateTime kokiDate;
            //        if (DateTime.TryParse(c1FlexGrid4[i, 1].ToString(), out kokiDate))
            //        {
            //            //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
            //            if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
            //            {
            //                varidateFlag = false;
            //                // 工期末日付は契約工期の期間内で設定して下さい。
            //                set_error(GlobalMethod.GetMessage("E10708", ""));
            //                break;
            //            }
            //        }
            //        else
            //        {
            //            // 工期末日付は契約工期の期間内で設定して下さい。
            //            set_error(GlobalMethod.GetMessage("E10708", ""));
            //            break;
            //        }
            //    }
            //    if (c1FlexGrid4[i, 9] != null && c1FlexGrid4[i, 9].ToString() != "")
            //    {
            //        DateTime kokiDate;
            //        if (DateTime.TryParse(c1FlexGrid4[i, 9].ToString(), out kokiDate))
            //        {
            //            //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
            //            if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
            //            {
            //                varidateFlag = false;
            //                // 工期末日付は契約工期の期間内で設定して下さい。
            //                set_error(GlobalMethod.GetMessage("E10708", ""));
            //                break;
            //            }
            //        }
            //        else
            //        {
            //            // 工期末日付は契約工期の期間内で設定して下さい。
            //            set_error(GlobalMethod.GetMessage("E10708", ""));
            //            break;
            //        }
            //    }
            //    if (c1FlexGrid4[i, 17] != null && c1FlexGrid4[i, 17].ToString() != "")
            //    {
            //        DateTime kokiDate;
            //        if (DateTime.TryParse(c1FlexGrid4[i, 17].ToString(), out kokiDate))
            //        {
            //            //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
            //            if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
            //            {
            //                varidateFlag = false;
            //                // 工期末日付は契約工期の期間内で設定して下さい。
            //                set_error(GlobalMethod.GetMessage("E10708", ""));
            //                break;
            //            }
            //        }
            //        else
            //        {
            //            // 工期末日付は契約工期の期間内で設定して下さい。
            //            set_error(GlobalMethod.GetMessage("E10708", ""));
            //            break;
            //        }
            //    }
            //    if (c1FlexGrid4[i, 25] != null && c1FlexGrid4[i, 25].ToString() != "")
            //    {
            //        DateTime kokiDate;
            //        if (DateTime.TryParse(c1FlexGrid4[i, 25].ToString(), out kokiDate))
            //        {
            //            //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
            //            if (item3_1_6.Value > kokiDate || item3_1_7.Value < kokiDate)
            //            {
            //                varidateFlag = false;
            //                // 工期末日付は契約工期の期間内で設定して下さい。
            //                set_error(GlobalMethod.GetMessage("E10708", ""));
            //                break;
            //            }
            //        }
            //        else
            //        {
            //            // 工期末日付は契約工期の期間内で設定して下さい。

            //            set_error(GlobalMethod.GetMessage("E10708", ""));
            //            break;
            //        }
            //    }
            //}
            ////契約タブの調査部配分率が0ではない場合
            ////if (GetDouble(item3_7_1_1_1.Text) > 0)
            ////{
            //    // 調査部 業務別配分が100でないとエラー
            //    if (item3_7_2_26_1.Text != "100.00%" && item3_7_2_26_1.Text != "0.00%")
            //    {
            //        // 調査業務別　配分の合計が100になるように入力してください。
            //        set_error(GlobalMethod.GetMessage("E70045", "(契約タブ)"));
            //        varidateFlag = false;
            //    }
            // 調査部 業務別配分の資材調査、営繕調査、機器類調査、工事費調査のいづれかが1以上でない
            // item3_7_2_14_1 資材調査
            // item3_7_2_15_1 営繕調査
            // item3_7_2_16_1 機器類調査
            // item3_7_2_17_1 工事費調査

            // 483 業務別配分の必須チェックを無しとし、横計の100%のチェックのみとする。
            //Double zeroDouble = GetDouble("0");
            //// 業務別配分の資材調査、営繕調査、機器類調査、工事費調査のいづれかが1以上でない
            //if (GetDouble(item3_7_2_14_1.Text) == zeroDouble && GetDouble(item3_7_2_15_1.Text) == zeroDouble && GetDouble(item3_7_2_16_1.Text) == zeroDouble && GetDouble(item3_7_2_17_1.Text) == zeroDouble)
            //{
            //    // 調査業務別 配分 資材調査、営繕調査、機器類調査、工事費調査のいずれかを入力してください。
            //    set_error(GlobalMethod.GetMessage("E70049", "(契約タブ)"));
            //    varidateFlag = false;
            //}
            //}

            ////売上計上情報
            //if (mode != "change")
            //{
            //    //引合タブの7.業務内容の調査部 業務別配分の合計が100
            //    chosaTotal = totalHundred;
            //    if (Decimal.Compare(chosaTotal, totalHundred) == 0)
            //    {
            //        Boolean nullFlag = false;
            //        //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない場合
            //        if ((c1FlexGrid4[2, 1] == null && c1FlexGrid4[2, 2] == null && c1FlexGrid4[2, 3] == null) || (c1FlexGrid4[2, 1] == null || c1FlexGrid4[2, 3] == null))
            //        {
            //            nullFlag = true;
            //            varidateFlag = false;
            //            set_error(GlobalMethod.GetMessage("E10715", "(調査部)"));
            //        }

            //        //2.配分情報の配分額(税込)の調査部
            //        if (!nullFlag)
            //        {
            //            int haibunTax = GetInt(item3_2_1_1.Text);
            //            //売上計上情報の調査部の計上額
            //            int keijoTotal = 0;
            //            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            //            {
            //                if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "")
            //                {
            //                    keijoTotal += GetInt(c1FlexGrid4[i, 3].ToString());
            //                }
            //            }
            //            //配分額(税込)の調査部と、調査部の計上額の合計が一致しない場合
            //            if (Decimal.Compare(haibunTax, keijoTotal) != 0)
            //            {
            //                varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(調査部)"));
            //            }
            //        }
            //    }
            //    //else 引合タブの7.業務内容の調査部 業務別配分の合計が100でない
            //    else
            //    {
            //        ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //        //if (c1FlexGrid4[2, 1] != null || c1FlexGrid4[2, 3] != null)
            //        //{
            //        //    varidateFlag = false;
            //        //    set_error(GlobalMethod.GetMessage("E10716", "(調査部)"));
            //        //}
            //    }


            //    //引合タブの7.業務内容の事業普及部の配分率(%)が100
            //    Decimal gyomuHaibun = Decimal.Parse(item1_7_1_2_1.Text.Substring(0, item1_7_1_2_1.Text.Length - 1));
            //    if (Decimal.Compare(gyomuHaibun, totalHundred) == 0)
            //    {
            //        Boolean nullFlag = false;
            //        //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
            //        if ((c1FlexGrid4[2, 4] == null && c1FlexGrid4[2, 5] == null && c1FlexGrid4[2, 6] == null) || (c1FlexGrid4[2, 4] == null || c1FlexGrid4[2, 6] == null))
            //        {
            //            nullFlag = true;
            //            varidateFlag = false;
            //            set_error(GlobalMethod.GetMessage("E10715", "(事業普及部)"));
            //        }

            //        if (!nullFlag)
            //        {
            //            //2.配分情報の配分額(税込)の事業普及部
            //            Decimal haibunTax = Decimal.Parse(item3_2_2_1.Text.Substring(1, item3_2_2_1.Text.Length - 1));
            //            //売上計上情報の事業普及部の計上額の合計
            //            String totalStr = c1FlexGrid4[2, 6].ToString();
            //            Decimal keijoTotal = Decimal.Parse(totalStr.Substring(1, totalStr.Length - 1));
            //            //配分額(税込)の調査部と、事業普及部の計上額の合計が一致しない場合
            //            if (Decimal.Compare(haibunTax, keijoTotal) != 0)
            //            {
            //                varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(事業普及部)"));
            //            }
            //        }
            //    }
            //    //else 引合タブの7.業務内容の部門配分 事業普及部の配分率(％)が0でない
            //    else if (Decimal.Compare(gyomuHaibun, totalZero) != 0)
            //    {
            //        ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //        //if (c1FlexGrid4[2, 4] != null || c1FlexGrid4[2, 6] != null)
            //        //{
            //        //    varidateFlag = false;
            //        //    set_error(GlobalMethod.GetMessage("E10716", "(事業普及部)"));
            //        //}
            //    }


            //    //引合タブの7.業務内容の情報システム部の配分率(%)が100
            //    Decimal systemHaibun = Decimal.Parse(item1_7_1_3_1.Text.Substring(0, item1_7_1_3_1.Text.Length - 1));
            //    if (Decimal.Compare(systemHaibun, totalHundred) == 0)
            //    {
            //        Boolean nullFlag = false;
            //        //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
            //        if ((c1FlexGrid4[2, 7] == null && c1FlexGrid4[2, 8] == null && c1FlexGrid4[2, 9] == null) || (c1FlexGrid4[2, 7] == null || c1FlexGrid4[2, 9] == null))
            //        {
            //            nullFlag = true;
            //            varidateFlag = false;
            //            set_error(GlobalMethod.GetMessage("E10715", "(情報システム部)"));
            //        }

            //        if (!nullFlag)
            //        {
            //            //2.配分情報の配分額(税込)の情報システム部
            //            Decimal haibunTax = Decimal.Parse(item3_2_3_1.Text.Substring(1, item3_2_3_1.Text.Length - 1));
            //            //売上計上情報の情報システム部の計上額の合計
            //            String totalStr = c1FlexGrid4[2, 9].ToString();
            //            Decimal keijoTotal = Decimal.Parse(totalStr.Substring(1, totalStr.Length - 1));
            //            //配分額(税込)の情報システム部と、情報システム部の計上額の合計が一致しない場合
            //            if (Decimal.Compare(haibunTax, keijoTotal) != 0)
            //            {
            //                varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(情報システム部)"));
            //            }
            //        }
            //    }
            //    //else 引合タブの7.業務内容の部門配分 情報システム部の配分率(％)が0でない
            //    else if (Decimal.Compare(gyomuHaibun, totalZero) != 0)
            //    {
            //        ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //        //if (c1FlexGrid4[2, 7] != null || c1FlexGrid4[2, 9] != null)
            //        //{
            //        //    varidateFlag = false;
            //        //    set_error(GlobalMethod.GetMessage("E10716", "(情報システム部)"));
            //        //}
            //    }

            //    //引合タブの7.業務内容の総合研究所の配分率(%)が100
            //    Decimal sogoHaibun = Decimal.Parse(item1_7_1_4_1.Text.Substring(0, item1_7_1_4_1.Text.Length - 1));
            //    if (Decimal.Compare(sogoHaibun, totalHundred) == 0)
            //    {
            //        Boolean nullFlag = false;
            //        //契約タブの5.売上計上情報の工期末日付、計上月、計上額が全て入っていない、または工期末日付か計上額のどちらかしか入力されていない
            //        if ((c1FlexGrid4[2, 10] == null && c1FlexGrid4[2, 11] == null && c1FlexGrid4[2, 12] == null) || (c1FlexGrid4[2, 10] == null || c1FlexGrid4[2, 12] == null))
            //        {
            //            nullFlag = true;
            //            varidateFlag = false;
            //            set_error(GlobalMethod.GetMessage("E10715", "(総合研究所)"));
            //        }

            //        if (!nullFlag)
            //        {
            //            //2.配分情報の配分額(税込)の総合研究所
            //            Decimal haibunTax = Decimal.Parse(item3_2_4_1.Text.Substring(1, item3_2_4_1.Text.Length - 1));
            //            //売上計上情報の総合研究所の計上額の合計
            //            String totalStr = c1FlexGrid4[2, 12].ToString();
            //            Decimal keijoTotal = Decimal.Parse(totalStr.Substring(1, totalStr.Length - 1));
            //            //配分額(税込)の総合研究所と、総合研究所の計上額の合計が一致しない場合
            //            if (Decimal.Compare(haibunTax, keijoTotal) != 0)
            //            {
            //                varidateFlag = false;
            //                set_error(GlobalMethod.GetMessage("E10717", "(総合研究所)"));
            //            }
            //        }
            //    }
            //    //else 引合タブの7.業務内容の部門配分 情報システム部の配分率(％)が0でない
            //    else if (Decimal.Compare(gyomuHaibun, totalZero) != 0)
            //    {
            //        ////契約タブの5.売上計上情報の工期末日付か計上額がどれか１つでも入力されている
            //        //if (c1FlexGrid4[2, 10] != null || c1FlexGrid4[2, 12] != null)
            //        //{
            //        //    varidateFlag = false;
            //        //    set_error(GlobalMethod.GetMessage("E10716", "(総合研究所)"));
            //        //}
            //    }
            //}


            if (!requiredFlag || !varidateFlag)
            {
                return false;
            }
            return true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (mode == "insert" || mode == "keikaku")
            {
                // 売上年度、受託課所支部が正しいか確認してください
                //if (MessageBox.Show("新規登録を行いますがよろしいでしょうか？\r\n下記について確認して下さい。\r\n売上年度（工期開始年度）、受託課所支部が正しいか確認して下さい。\r\n※売上年度と工期開始年度が異なる場合は、工期開始年度で新規登録し、登録後売上年度に変更して下さい。", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                if (MessageBox.Show("新規登録を行いますがよろしいでしょうか？\r\n下記について確認して下さい。\r\n工期開始年度、受託課所支部が正しいか確認して下さい。", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    //// 新規登録を行いますが宜しいですか？
                    //if (MessageBox.Show(GlobalMethod.GetMessage("I10601", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                    //{
                    if (!ErrorFLG(0))
                    {
                        if (Execute_SQL(0))
                        {
                            //えんとり君修正STEP2
                            sItem1_10_ori = item1_10.SelectedValue.ToString(); //受託課所支部（契約部所）DB値
                            sItem1_2_KoukiNendo_ori = item1_2_KoukiNendo.SelectedValue.ToString(); //工期開始年度DB値
                            sJigyoubuHeadCD_ori = getJigyoubuHeadCD();
                        }
                    }
                    //}
                }
            }
            else
            {
                if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    //担当者項目の必須チェックを追加
                    if (!ErrorFLG(1))
                    {
                        if (Execute_SQL(1))
                        {
                            //えんとり君修正STEP2
                            sItem1_10_ori = item1_10.SelectedValue.ToString(); //受託課所支部（契約部所）DB値
                            sItem1_2_KoukiNendo_ori = item1_2_KoukiNendo.SelectedValue.ToString(); //工期開始年度DB値
                            sJigyoubuHeadCD_ori = getJigyoubuHeadCD();
                            //受託番号が採番された場合、「この案件番号の枝番で受託番号を作成する」ボタンを表示
                            if (item1_8.Text != "")
                            {
                                button10.Visible = true;
                            }
                            //「この案件番号の枝番で受託番号を作成する」ボタンは押下時に受託番号をチェックするように変更
                            ////受託番号が採番された場合、「この案件番号の枝番で受託番号を作成する」ボタンを有効化
                            //if (item1_8.Text != "")
                            //{
                            //    button10.Visible = true;
                            //    button10.BackColor = Color.FromArgb(42, 78, 122);
                            //    button10.Enabled = true;
                            //}
                        }
                    }
                }

            }

            //えんとり君修正STEP2 不具合1406
            FolderPathCheck();
        }

        private Boolean Execute_SQL(int mode)
        {
            string methodName = ".Execute_SQL";
            // ▼mode
            // 0:新規登録 
            // 1:更新 
            // 2:チェック用出力 
            // 3:起案
            // 4:変更伝票後の起案
            if (mode == 0)
            {
                int ankenID = GlobalMethod.getSaiban("AnkenJouhouID");
                string ankenNo = "";
                string ankenEda = "";
                string jutakubangou = "";
                string jigyoubuHeadCD = "";

                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

                //案件番号
                if (item1_6.Text == "")
                {
                    using (var conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        var cmd = conn.CreateCommand();

                        // 採番ロジックを以下に変更 20210222
                        // 契約区分（契約区分が調査部から始まるものはT等）、部所コードは受託部所（KashoShibuCD）、連番は、事業部ヘッド、年度、部所コードの連番

                        // 契約区分で業務分類CDを判定
                        // Mst_Jigyoubu に問い合わせる方法が無い為、
                        // 調査部が見つかった場合、T と判断
                        if (item1_14.Text.IndexOf("調査部") > -1)
                        {
                            jigyoubuHeadCD = "T";
                        }
                        else if (item1_14.Text.IndexOf("事業普及部") > -1)
                        {
                            jigyoubuHeadCD = "B";
                        }
                        else if (item1_14.Text.IndexOf("情シス部") > -1)
                        {
                            jigyoubuHeadCD = "J";
                        }
                        else if (item1_14.Text.IndexOf("総合研究所") > -1)
                        {
                            jigyoubuHeadCD = "K";
                        }

                        // 業務分類CD + 年度下2桁
                        //ankenNo = jigyoubuHeadCD + item1_3.SelectedValue.ToString().Substring(2, 2);
                        ankenNo = jigyoubuHeadCD + item1_2_KoukiNendo.SelectedValue.ToString().Substring(2, 2);

                        //cmd.CommandText = "SELECT  " +
                        //        "JigyoubuHeadCD " +

                        //        //参照テーブル
                        //        "FROM Mst_Busho " +
                        //        "WHERE GyoumuBushoCD = '" + item1_10.SelectedValue + "' ";
                        //var sda = new SqlDataAdapter(cmd);
                        //var dt = new DataTable();
                        //sda.Fill(dt);
                        //// JigyoubuHeadCD + 年度の下2桁 + 受託課所支部のCDの頭3桁
                        ////ankenNo = dt.Rows[0][0].ToString() + item1_3.SelectedValue.ToString().Substring(2, 2) + item1_10.SelectedValue.ToString().Substring(0, 3);
                        //ankenNo = dt.Rows[0][0].ToString() + item1_3.SelectedValue.ToString().Substring(2, 2);

                        // KashoShibuCD
                        cmd.CommandText = "SELECT  " +
                                "KashoShibuCD " +

                                //参照テーブル
                                "FROM Mst_Busho " +
                                "WHERE GyoumuBushoCD = '" + item1_10.SelectedValue.ToString() + "' ";
                        var sda = new SqlDataAdapter(cmd);
                        var dt = new DataTable();
                        sda.Fill(dt);
                        // KashoShibuCDが正しい
                        ankenNo = ankenNo + dt.Rows[0][0].ToString();

                        cmd.CommandText = "SELECT TOP 1 " +
                                " SUBSTRING(AnkenAnkenBangou,7,3) " +

                                //参照テーブル
                                "FROM AnkenJouhou " +
                                "WHERE AnkenAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + ankenNo + "%' and AnkenDeleteFlag != 1 ORDER BY AnkenAnkenBangou DESC";
                        sda = new SqlDataAdapter(cmd);
                        Console.WriteLine(cmd.CommandText);
                        dt = new DataTable();
                        sda.Fill(dt);

                        if (dt != null && dt.Rows.Count > 0)
                        {
                            int AnkenNoRenban;
                            if (int.TryParse(dt.Rows[0][0].ToString(), out AnkenNoRenban))
                            {
                                AnkenNoRenban++;
                            }
                            else
                            {
                                AnkenNoRenban = 1;
                            }
                            ankenNo += string.Format("{0:D3}", AnkenNoRenban);
                        }
                        else
                        {
                            ankenNo += "001";
                        }
                    }
                }
                else
                {
                    ankenNo = item1_6.Text;
                }
                // 受託番号
                if (item1_7.Text != "")
                {
                    jutakubangou = item1_7.Text;
                }
                //枝番
                if (item1_8.Text != "")
                {
                    ankenEda = item1_8.Text;
                }
                else
                {
                    //新規登録時は落札者は選択されない
                    /*
                    // item2_3_7：落札者（落札状況にチェックが無いと入らないので、落札状況はチェックなしとする）
                    // ENTORY_TOUKAI:建設物価調査会
                    if (item2_3_7.Text == GlobalMethod.GetCommonValue2("ENTORY_TOUKAI"))
                    {
                        // 受託番号の枝番を取得する
                        using (var conn = new SqlConnection(connStr))
                        {
                            try
                            {
                                conn.Open();
                                var cmd = conn.CreateCommand();

                                cmd.CommandText = "SELECT  " +
                                        "TOP 1 AnkenJutakuBangouEda " +
                                        "FROM AnkenJouhou " +
                                        "WHERE AnkenAnkenBangou = (SELECT AnkenAnkenBangou FROM AnkenJouhou WHERE AnkenJouhouID = " + AnkenID + ") and AnkenDeleteFlag != 1 order by AnkenJutakuBangouEda desc ";
                                var sda = new SqlDataAdapter(cmd);
                                var dt = new DataTable();
                                sda.Fill(dt);
                                if (dt != null && dt.Rows.Count > 0)
                                {
                                    // 枝番（-nn）を取得
                                    ankenEda = dt.Rows[0][0].ToString();
                                    //// 「-」落とし
                                    //ankenEda = ankenEda.Replace("-", "");
                                    int i = int.Parse(ankenEda);
                                    i += 1;
                                    ankenEda = "" + string.Format("{0:D2}", i);
                                }
                                else
                                {
                                    // 枝番（01）をセット
                                    ankenEda = "01";
                                }
                                dt.Clear();
                            }
                            catch (Exception)
                            {
                                // エラー
                                GlobalMethod.outputLogger("Execute_SQL", "枝番が取得できずにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                            }
                        }
                        // 受託番号の自動生成
                        jutakubangou = ankenNo + "-" + ankenEda;

                        // 引合タブの受託番号、枝番にセット
                        item1_7.Text = jutakubangou;
                        item1_8.Text = ankenEda;
                        Header2.Text = jutakubangou;
                    }
                    */
                }

                if (GlobalMethod.Check_Table(ankenID.ToString(), "AnkenJouhouID", "AnkenJouhou", ""))
                {
                    GlobalMethod.outputLogger("InsertAnken", "契約情報登録 ID重複エラー", ankenID.ToString(), UserInfos[1]);
                    set_error(GlobalMethod.GetMessage("E00091", "(契約情報ID重複)"));
                    return false;
                }

                // 案件（受託）フォルダー
                string ankenFolder = "";
                // 画面で選択している受託課所支部の部所フォルダ
                string replaceFolderName = "";
                // 置き換え対象のログインユーザーの部所フォルダ
                string replaceTargetFolderName = "";

                // 画面の受託課所支部の部所フォルダを取得する
                using (var conn = new SqlConnection(connStr))
                {
                    try
                    {
                        conn.Open();
                        var cmd = conn.CreateCommand();

                        cmd.CommandText = "SELECT  " +
                                "FolderPath " +
                                "FROM M_Folder " +
                                "WHERE MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + item1_10.SelectedValue.ToString() + "' ";
                        var sda = new SqlDataAdapter(cmd);
                        var dt = new DataTable();
                        sda.Fill(dt);
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            // フォルダパスを取得（例：$FOLDER_BASE$/111統括）
                            replaceFolderName = dt.Rows[0][0].ToString();
                            // 課所支部のフォルダ部分のみとする $FOLDER_BASE$/xxx 
                            replaceFolderName = replaceFolderName.Replace(@"$FOLDER_BASE$/", "");
                            // 課所支部のフォルダ部分のみとする $FOLDER_BASE$ しかない場合の対応
                            replaceFolderName = replaceFolderName.Replace(@"$FOLDER_BASE$", "");
                        }
                        dt.Clear();
                    }
                    catch (Exception)
                    {
                        // エラー
                        GlobalMethod.outputLogger("Execute_SQL", "M_Folderから画面の受託課所支部：" + item1_10.SelectedValue.ToString() + " のフォルダパスが取得できずにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                    }
                }
                // 置き換え対象のログインユーザーの部所フォルダを取得する
                using (var conn = new SqlConnection(connStr))
                {
                    try
                    {
                        conn.Open();
                        var cmd = conn.CreateCommand();

                        cmd.CommandText = "SELECT  " +
                                "FolderPath " +
                                "FROM M_Folder " +
                                "WHERE MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + UserInfos[2] + "' ";
                        var sda = new SqlDataAdapter(cmd);
                        var dt = new DataTable();
                        sda.Fill(dt);
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            // フォルダパスを取得（例：$FOLDER_BASE$/111統括）
                            replaceTargetFolderName = dt.Rows[0][0].ToString();
                            // 課所支部のフォルダ部分のみとする
                            replaceTargetFolderName = replaceTargetFolderName.Replace(@"$FOLDER_BASE$/", "");
                        }
                        dt.Clear();
                    }
                    catch (Exception)
                    {
                        // エラー
                        GlobalMethod.outputLogger("Execute_SQL", "M_Folderからログインユーザーの受託課所支部：" + UserInfos[2] + " のフォルダパスが取得できずにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                    }
                }

                // 案件（受託）フォルダ
                ankenFolder = item1_12.Text;
                if (replaceTargetFolderName != "")
                {
                    // 自分の部署フォルダを画面の選択している受託課所支部のフォルダに置き換える
                    ankenFolder = ankenFolder.Replace(replaceTargetFolderName, replaceFolderName);

                    // 867
                    // 工期開始年度　2021年度まで、　010北道
                    // 工期開始年度　2022年度から　　010北海
                    int koukinendo = 0;
                    if (int.TryParse(item1_2_KoukiNendo.SelectedValue.ToString(), out koukinendo))
                    {
                        if (koukinendo > 2021)
                        {
                            // 010北道
                            string str1 = GlobalMethod.GetCommonValue1("MADOGUCHI_HOKKAIDO_PATH");
                            // 010北海
                            string str2 = GlobalMethod.GetCommonValue2("MADOGUCHI_HOKKAIDO_PATH");

                            if (str1 != null && str2 != null)
                            {
                                ankenFolder = ankenFolder.Replace(str1, str2);
                            }
                        }
                    }
                }

                // 最終文字が\マークとなったら取り除く
                if (ankenFolder.Length > 0 && ankenFolder.EndsWith(@"\"))
                {
                    ankenFolder = ankenFolder.Substring(0, ankenFolder.Length - 1);
                }

                // 受託部所の所属長を取得する
                string BushoShozokuChou = "";
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    var dt = new DataTable();

                    cmd.CommandText = "SELECT BushoShozokuChou FROM Mst_Busho WHERE GyoumuBushoCD = '" + item1_10.SelectedValue + "'";

                    var sda = new SqlDataAdapter(cmd);
                    dt.Clear();
                    sda.Fill(dt);
                    BushoShozokuChou = dt.Rows[0][0].ToString();
                }

                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {
                        cmd.CommandText = "INSERT INTO AnkenJouhou ( " +
                                            "AnkenJouhouID " +
                                            ",AnkenHikiaijhokyo " +
                                            ",AnkenSakuseiKubun " +
                                            ",AnkenUriageNendo " +
                                            ",AnkenKeikakuBangou " +
                                            ",AnkenAnkenBangou " +
                                            ",AnkenJutakuBangou " +
                                            ",AnkenJutakuBangouEda " +
                                            ",AnkenTourokubi " +

                                            ",AnkenJutakushibu " +
                                            ",AnkenJutakubushoCD " +
                                            ",AnkenKeiyakusho " +
                                            ",AnkenTantoushaCD " +
                                            ",AnkenTantoushaMei " +

                                            ",AnkenGyoumuMei " +
                                            ",AnkenGyoumuKubun " +
                                            ",AnkenGyoumuKubunCD " +
                                            ",AnkenGyoumuKubunMei " +
                                            ",AnkenNyuusatsuHoushiki " +
                                            ",AnkenNyuusatsuYoteibi " +

                                            ",AnkenHachushaCD " +
                                            ",AnkenHachuushaMei " +
                                            ",AnkenHachushaKaMei " +
                                            ",AnkenHachuushaKaMei " +

                                            ",AnkenHachuushaKeiyakuBusho " +
                                            ",AnkenHachuushaKeiyakuTantou " +
                                            ",AnkenHachuushaKeiyakuTEL " +
                                            ",AnkenHachuushaKeiyakuFAX " +
                                            ",AnkenHachuushaKeiyakuMail " +
                                            ",AnkenHachuushaKeiyakuYuubin " +
                                            ",AnkenHachuushaKeiyakuJuusho " +
                                            ",AnkenHachuuDaihyouYakushoku " +
                                            ",AnkenHachuuDaihyousha " +

                                            ",AnkenToukaiSankouMitsumori " +
                                            ",AnkenToukaiJyutyuIyoku " +
                                            ",AnkenToukaiSankouMitsumoriGaku" +

                                            ",AnkenCreateProgram " +
                                            ",AnkenCreateDate " +
                                            ",AnkenCreateUser " +
                                            ",AnkenUpdateDate " +
                                            ",AnkenUpdateUser " +
                                            ",AnkenDeleteFlag " +
                                            ",AnkenSaishinFlg " +
                                            ",AnkenGyoumuKanrishaCD " +
                                            ",AnkenMadoguchiTantoushaCD " +
                                            ",GyoumuKanrishaCD " +
                                            ",AnkenKaisuu " +
                                            ",AnkenKoukiNendo " +

                                            ",AnkenKokyakuHyoukaComment " +
                                            ",AnkenToukaiHyoukaComment " +
                                            ",AnkenGyoumuKanrisha " +

                                            " ) VALUES ( " +
                                            ankenID +                                                                 // AnkenJouhouID
                                            ",  " + item1_1.SelectedValue + " " +                                     // AnkenHikiaijhokyo
                                            ", '" + item1_2.SelectedValue + "' " +                                    // AnkenSakuseiKubun
                                            ",  " + item1_3.SelectedValue + " " +                                     // AnkenUriageNendo
                                            ", '" + item1_4.Text + "' " +                                             // AnkenKeikakuBangou
                                            ",N'" + ankenNo + "' " +                                                   // AnkenAnkenBangou
                                            ",N'" + jutakubangou + "' " +                                              // AnkenJutakuBangou
                                            ",N'" + ankenEda + "' " +                                                  // AnkenJutakuBangouEda
                                            ",  " + Get_DateTimePicker("item1_9") + " " +                             // AnkenTourokubi

                                            ", N'" + item1_10.Text + "'" +                                             // AnkenJutakushibu
                                            ", N'" + item1_10.SelectedValue + "'" +                                    // AnkenJutakubushoCD
                                                                                                                       //", '" + GlobalMethod.ChangeSqlText(item1_12.Text, 0, 0) + "'" +           // AnkenKeiyakusho
                                            ", N'" + GlobalMethod.ChangeSqlText(ankenFolder, 0, 0) + "'" +           // AnkenKeiyakusho
                                            ", '" + item1_11_CD.Text + "'" +                                          // AnkenTantoushaCD
                                            ", N'" + item1_11.Text + "'" +                                             // AnkenTantoushaMei

                                            ", N'" + GlobalMethod.ChangeSqlText(item1_13.Text, 0, 0) + "'" +           // AnkenGyoumuMei
                                            ",  " + item1_14.SelectedValue + " " +                                    // AnkenGyoumuKubun
                                            ",  N'" + Get_GyoumuKubunCD(item1_14.SelectedValue.ToString()) + "' " +    // AnkenGyoumuKubunCD
                                            ",  N'" + item1_14.Text + "' " +                                           // AnkenGyoumuKubunMei
                                            ",  N'" + item1_15.SelectedValue + "' " +                                  // AnkenNyuusatsuHoushiki
                                            ",  " + Get_DateTimePicker("item1_16") + " " +                            // AnkenNyuusatsuYoteibi

                                            ", N'" + item1_19.Text + "'" +                                             // AnkenHachushaCD
                                            ", N'" + item1_23.Text + "'" +                                             // AnkenHachuushaMei
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_24.Text, 0, 0) + "'" +           // AnkenHachushaKaMei
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_23.Text, 0, 0) + "　" + GlobalMethod.ChangeSqlText(item1_24.Text, 0, 0) + "'" + // AnkenHachuushaKaMei

                                            ", N'" + GlobalMethod.ChangeSqlText(item1_25.Text, 0, 0) + "'" +           // AnkenHachuushaKeiyakuBusho
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_26.Text, 0, 0) + "'" +           // AnkenHachuushaKeiyakuTantou
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_27.Text, 0, 0) + "'" +           // AnkenHachuushaKeiyakuTEL
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_28.Text, 0, 0) + "'" +           // AnkenHachuushaKeiyakuFAX
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_29.Text, 0, 0) + "'" +           // AnkenHachuushaKeiyakuMail
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_30.Text, 0, 0) + "'" +           // AnkenHachuushaKeiyakuYuubin
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_31.Text, 0, 0) + "'" +           // AnkenHachuushaKeiyakuJuusho
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_32.Text, 0, 0) + "'" +           // AnkenHachuuDaihyouYakushoku
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_33.Text, 0, 0) + "'" +           // AnkenHachuuDaihyousha

                                            ", N'" + item1_34.SelectedValue + "'" +                                   // AnkenToukaiSankouMitsumori
                                            ", N'" + item1_35.SelectedValue + "'" +                                   // AnkenToukaiJyutyuIyoku
                                            ", " + item1_36.Text.Replace("¥", "").Replace(",", "") + " " +           // AnkenToukaiSankouMitsumoriGaku

                                            ", 'InsertEntry'" +                                                       // AnkenCreateProgram
                                            ",  GETDATE() " +                                                         // AnkenCreateDate
                                            ", N'" + UserInfos[0] + "'" +                                              // AnkenCreateUser
                                            ",  GETDATE() " +                                                         // AnkenUpdateDate
                                            ", N'" + UserInfos[0] + "'" +                                              // AnkenUpdateUser
                                            ", 0" +                                                                   // AnkenDeleteFlag
                                            ", 1" +                                                                   // AnkenSaishinFlg
                                            ", null" +                                                                // AnkenGyoumuKanrishaCD
                                            ", null" +                                                                // AnkenMadoguchiTantoushaCD
                                            ", null" +                                                                // GyoumuKanrishaCD
                                            ", 0" +                                                                   // AnkenKaisuu
                                            ", N'" + item1_2_KoukiNendo.SelectedValue.ToString() + "'" +               // AnkenKoukiNendo
                                            ", '特になし'" +                                                          // AnkenKokyakuHyoukaComment
                                            ", '特になし'" +                                                          // AnkenToukaiHyoukaComment
                                            ", N'" + BushoShozokuChou + "'" +                                          // AnkenGyoumuKanrisha
                                            ")";

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        // 今の計画番号の案件数
                        if (item1_4.Text != "")
                        {
                            cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_4.Text + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_4.Text + "' ";
                            Console.WriteLine(cmd.CommandText);
                            cmd.ExecuteNonQuery();
                        }
                        // 変更前の計画番号の案件数
                        if (beforeKeikakuBangou != "")
                        {
                            cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + beforeKeikakuBangou + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + beforeKeikakuBangou + "' ";
                            Console.WriteLine(cmd.CommandText);
                            cmd.ExecuteNonQuery();
                        }

                        for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                        {
                            if (c1FlexGrid1.Rows[i][2] != null && c1FlexGrid1.Rows[i][2].ToString() != "")
                            {
                                string AnkenZenkaiRakusatsushaID = "";
                                string AnkenZenkaiJutakuKingaku = "";
                                string KeiyakuZenkaiRakusatsushaID = "";
                                string AnkenZenkaiAnkenJouhouID = "";
                                string AnkenZenkaiJutakuZeinuki = "";

                                // 前回落札者ID
                                if (c1FlexGrid1.Rows[i][8].ToString() != "")
                                {
                                    AnkenZenkaiRakusatsushaID = c1FlexGrid1.Rows[i][8].ToString().Trim();
                                }
                                else
                                {
                                    AnkenZenkaiRakusatsushaID = "null";
                                }
                                // 前回受託金額
                                if (c1FlexGrid1.Rows[i][9].ToString() != "")
                                {
                                    AnkenZenkaiJutakuKingaku = c1FlexGrid1.Rows[i][9].ToString();
                                }
                                else
                                {
                                    AnkenZenkaiJutakuKingaku = "null";
                                }
                                // 前回契約落札者ID
                                if (c1FlexGrid1.Rows[i][14].ToString().Trim() != "")
                                {
                                    KeiyakuZenkaiRakusatsushaID = c1FlexGrid1.Rows[i][14].ToString().Trim();
                                }
                                else
                                {
                                    KeiyakuZenkaiRakusatsushaID = "null";
                                }
                                // 前回契約情報ID
                                if (c1FlexGrid1.Rows[i][2].ToString().Trim() != "")
                                {
                                    AnkenZenkaiAnkenJouhouID = c1FlexGrid1.Rows[i][2].ToString().Trim();
                                }
                                else
                                {
                                    AnkenZenkaiAnkenJouhouID = "null";
                                }
                                // 前回受託税抜
                                if (c1FlexGrid1.Rows[i][13].ToString() != "")
                                {
                                    AnkenZenkaiJutakuZeinuki = c1FlexGrid1.Rows[i][13].ToString();
                                }
                                else
                                {
                                    AnkenZenkaiJutakuZeinuki = "null";
                                }

                                if (c1FlexGrid1.Rows[i][7] == null)
                                {
                                    c1FlexGrid1.Rows[i][7] = c1FlexGrid1.Rows[i][7];
                                }

                                cmd.CommandText = "INSERT INTO AnkenJouhouZenkaiRakusatsu ( " +
                                                    "AnkenJouhouID " +
                                                    ",AnkenZenkaiRakusatsuID " +
                                                    ",AnkenZenkaiAnkenBangou " +
                                                    ",AnkenZenkaiJutakuBangou " +
                                                    ",AnkenZenkaiJutakuEdaban " +
                                                    ",AnkenZenkaiGyoumuMei " +

                                                    ",AnkenZenkaiRakusatsusha " +
                                                    ",AnkenZenkaiRakusatsushaID " +

                                                    ",AnkenZenkaiJutakuKingaku " +
                                                    ",KeiyakuZenkaiRakusatsushaID " +
                                                    ",AnkenZenkaiKyougouKigyouCD " +
                                                    ",AnkenZenkaiAnkenJouhouID " +
                                                    ",AnkenZenkaiJutakuZeinuki " +
                                                    " ) VALUES ( " +
                                                     ankenID +                                             // [AnkenJouhouID] [decimal](16, 0) NOT NULL,
                                                                                                           //", " + i +                                          // [AnkenZenkaiRakusatsuID] [int] NOT NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][16].ToString() + "'" +    // [AnkenZenkaiRakusatsuID] [int] NOT NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][3].ToString() + "'" +     // [AnkenZenkaiAnkenBangou] [nvarchar](40) NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][4].ToString() + "'" +     // [AnkenZenkaiJutakuBangou] [nvarchar](40) NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][5].ToString() + "'" +     // [AnkenZenkaiJutakuEdaban] [nvarchar](50) NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][6].ToString() + "'" +     // [AnkenZenkaiGyoumuMei] [nvarchar](150) NULL,

                                                     ", N'" + c1FlexGrid1.Rows[i][7].ToString() + "'" +     // [AnkenZenkaiRakusatsusha] [nvarchar](50) NULL,
                                                     ", " + AnkenZenkaiRakusatsushaID + "" +               // [AnkenZenkaiRakusatsushaID] [int] NULL,
                                                     ", " + AnkenZenkaiJutakuKingaku + "" +       // [AnkenZenkaiJutakuKingaku] [money] NULL,
                                                     ", " + KeiyakuZenkaiRakusatsushaID + "" +             // [KeiyakuZenkaiRakusatsushaID] [int] NULL,        
                                                     ", N'" + c1FlexGrid1.Rows[i][15].ToString() + "'" +    // [AnkenZenkaiKyougouKigyouCD] [nvarchar](24) NULL,
                                                     ", " + AnkenZenkaiAnkenJouhouID + "" +                // [AnkenZenkaiAnkenJouhouID] [decimal](16, 0) NULL,
                                                     ", " + c1FlexGrid1.Rows[i][13].ToString() + "" +      // [AnkenZenkaiJutakuZeinuki] [money] NULL,
                                                    ")";

                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();
                            }
                        }


                        if (GlobalMethod.Check_Table(ankenID.ToString(), "AnkenJouhouID", "KeiyakuJouhouEntory", ""))
                        {
                            GlobalMethod.outputLogger("InsertAnken", "契約情報(エントリー)登録 ID重複エラー", ankenID.ToString(), UserInfos[1]);
                            set_error(GlobalMethod.GetMessage("E00091", "(契約情報(エントリー)ID重複)"));
                            transaction.Rollback();
                            return false;
                        }

                        cmd.CommandText = $"INSERT INTO KeiyakuJouhouEntory ( KeiyakuJouhouEntoryID ,KeiyakuSakuseiKubunID ,KeiyakuSakuseiKubun ,KeiyakuHachuushaMei ,KeiyakuGyoumuKubun ,KeiyakuGyoumuMei ,JutakuBushoCD ,KeiyakuTantousha ,KeiyakuUriageHaibunCho1 ,KeiyakuUriageHaibunCho2 ,KeiyakuUriageHaibunJo1 ,KeiyakuUriageHaibunJo2 ,KeiyakuUriageHaibunJosys1 ,KeiyakuUriageHaibunJosys2 ,KeiyakuUriageHaibunKei1 ,KeiyakuUriageHaibunKei2 ,KeiyakuUriageHaibunChoGoukei ,KeiyakuUriageHaibunJoGoukei ,KeiyakuUriageHaibunJosysGoukei ,KeiyakuUriageHaibunKeiGoukei ,KeiyakuUriageHaibunGoukei ,KeiyakuKeiyakuTeiketsubi ,AnkenJouhouID ,KeiyakuDeleteFlag ,KeiyakuCreateDate ,KeiyakuCreateUser ,KeiyakuUpdateDate ,KeiyakuUpdateUser ,KeiyakuCreateProgram  ) VALUES ( {ankenID}, N'{item1_2.SelectedValue}', N'{item1_2.Text}', N'{item1_23.Text}', N'{item1_14.SelectedValue}', N'{item1_14.Text}',   null  ,   null  ,   0  ,   0  ,   0  ,   0  ,   0  ,   0  ,   0  ,   0  ,   0  ,   0  ,   0  ,   0  ,   0  ,   null  ,  {ankenID} ,   0  ,  GETDATE() , N'{UserInfos[0]}',  GETDATE() , N'{UserInfos[0]}', 'InsertEntry')";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();


                        GlobalMethod.outputLogger("InsertAnken", "契約情報更新時に案件情報を同時更新します AnkenjouhoID = " + ankenID, "", UserInfos[1]);

                        if (GlobalMethod.Check_Table(ankenID.ToString(), "GyoumuJouhouID", "GyoumuJouhou", ""))
                        {
                            set_error(GlobalMethod.GetMessage("E00091", "(業務情報ID重複)"));
                            transaction.Rollback();
                            return false;
                        }

                        cmd.CommandText = "INSERT INTO GyoumuJouhou ( " +
                                            "GyoumuJouhouID " +
                                            ",KanriGijutsushaCD " +
                                            ",ShousaTantoushaCD " +
                                            ",SinsaTantoushaCD " +
                                            ",GyoumuDeleteFlag " +
                                            ",GyoumuCreateDate " +
                                            ",GyoumuCreateUser " +
                                            ",GyoumuCreateProgram " +
                                            ",GyoumuUpdateDate " +
                                            ",GyoumuUpdateUser " +
                                            ",AnkenJouhouID " +
                                            " ) VALUES ( " +
                                            ankenID +
                                            ",   null  " +
                                            ",   null  " +
                                            ",   null  " +
                                            ",   0  " +
                                            ",  GETDATE() " +
                                            ", N'" + UserInfos[0] + "'" +
                                            ", 'InsertEntry'" +
                                            ",  GETDATE() " +
                                            ", N'" + UserInfos[0] + "'" +
                                            ",  " + ankenID + " " +
                                            ")";

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        //業務配分
                        int HaibunID = GlobalMethod.getSaiban("GyoumuHaibunID");
                        cmd.CommandText = "INSERT INTO GyoumuHaibun ( " +
                                            "GyoumuHaibunID " +
                                            ",GyoumuAnkenJouhouID " +
                                            ",GyoumuHibunKubun " +
                                            ",GyoumuChosaBuRitsu " +
                                            ",GyoumuJigyoFukyuBuRitsu " +
                                            ",GyoumuJyohouSystemBuRitsu " +
                                            ",GyoumuSougouKenkyuJoRitsu " +
                                            ",GyoumuShizaiChousaRitsu " +
                                            ",GyoumuEizenRitsu " +
                                            ",GyoumuKikiruiChousaRitsu " +
                                            ",GyoumuKoujiChousahiRitsu " +
                                            ",GyoumuSanpaiFukusanbutsuRitsu " +
                                            ",GyoumuHokakeChousaRitsu " +
                                            ",GyoumuShokeihiChousaRitsu " +
                                            ",GyoumuGenkaBunsekiRitsu " +
                                            ",GyoumuKijunsakuseiRitsu " +
                                            ",GyoumuKoukyouRoumuhiRitsu " +
                                            ",GyoumuRoumuhiKoukyouigaiRitsu " +
                                            ",GyoumuSonotaChousabuRitsu " +
                                            " ) VALUES ( " +
                                            HaibunID +
                                            ",  '" + ankenID + "'  " +
                                            ",   10  " +
                                            ", N'" + item1_7_1_1_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_1_2_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_1_3_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_1_4_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_1_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_2_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_3_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_4_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_5_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_6_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_7_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_8_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_9_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_10_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_11_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + "' " +
                                            ", N'" + item1_7_2_12_1.Text.Replace("%", "") + "' " +
                                            ")";

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        HaibunID = GlobalMethod.getSaiban("GyoumuHaibunID");
                        cmd.CommandText = "INSERT INTO GyoumuHaibun ( " +
                                            "GyoumuHaibunID " +
                                            ",GyoumuAnkenJouhouID " +
                                            ",GyoumuHibunKubun " +
                                            ",GyoumuChosaBuGaku " +
                                            ",GyoumuJigyoFukyuBuGaku " +
                                            ",GyoumuJyohouSystemBuGaku " +
                                            ",GyoumuSougouKenkyuJoGaku " +
                                            ",GyoumuShizaiChousaRitsu " +
                                            ",GyoumuEizenRitsu " +
                                            ",GyoumuKikiruiChousaRitsu " +
                                            ",GyoumuKoujiChousahiRitsu " +
                                            ",GyoumuSanpaiFukusanbutsuRitsu " +
                                            ",GyoumuHokakeChousaRitsu " +
                                            ",GyoumuShokeihiChousaRitsu " +
                                            ",GyoumuGenkaBunsekiRitsu " +
                                            ",GyoumuKijunsakuseiRitsu " +
                                            ",GyoumuKoukyouRoumuhiRitsu " +
                                            ",GyoumuRoumuhiKoukyouigaiRitsu " +
                                            ",GyoumuSonotaChousabuRitsu " +
                                            ",GyoumuShizaiChousaGaku " +
                                            ",GyoumuEizenGaku " +
                                            ",GyoumuKikiruiChousaGaku " +
                                            ",GyoumuKoujiChousahiGaku " +
                                            ",GyoumuSanpaiFukusanbutsuGaku " +
                                            ",GyoumuHokakeChousaGaku " +
                                            ",GyoumuShokeihiChousaGaku " +
                                            ",GyoumuGenkaBunsekiGaku " +
                                            ",GyoumuKijunsakuseiGaku " +
                                            ",GyoumuKoukyouRoumuhiGaku " +
                                            ",GyoumuRoumuhiKoukyouigaiGaku " +
                                            ",GyoumuSonotaChousabuGaku " +
                                            " ) VALUES ( " +
                                            HaibunID +
                                            ",  '" + ankenID + "'  " +
                                            ",   30  " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ", 0 " +
                                            ")";

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        if (GlobalMethod.Check_Table(ankenID.ToString(), "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("InsertAnken", "入札情報登録 ID重複エラー", ankenID.ToString(), UserInfos[1]);
                            set_error(GlobalMethod.GetMessage("E00091", "(入札情報ID重複)"));
                            transaction.Rollback();
                            return false;
                        }

                        cmd.CommandText = "INSERT INTO NyuusatsuJouhou ( " +
                                            "NyuusatsuJouhouID " +
                                            ",NyuusatsuHoushiki " +
                                            ",NyuusatsuRakusatuSougaku " +
                                            ",NyuusatsuRakusatsushaID " +        // 入札状況
                                            ",NyuusatsuRakusatsusha " +          // 落札者（建設物価調査会等が入る）
                                            ",NyuusatsuKyougouTashaID " +
                                            ",NyuusatsuGyoumuBikou " +
                                            ",NyuusatsuMitsumorigaku " +
                                            ",NyuusatsuCreateProgram " +
                                            ",NyuusatsuCreateDate " +
                                            ",NyuusatsuCreateUser " +
                                            ",NyuusatsuUpdateDate " +
                                            ",NyuusatsuUpdateUser " +
                                            ",NyuusatsuDeleteFlag " +
                                            ",AnkenJouhouID " +
                                            ") VALUES ( " +
                                            ankenID +
                                            ", N'" + item1_15.SelectedValue + "'" +
                                            ", 0" +
                                            ", N'" + item1_17.SelectedValue + "'" +
                                            ", ''" +
                                            ",   null  " +
                                            ", N'" + GlobalMethod.ChangeSqlText(item1_18.Text, 0, 0) + "'" +
                                            ", N'" + item1_36.Text + "'" +
                                            ", 'InsertEntry'" +
                                            ",  GETDATE() " +
                                            ", N'" + UserInfos[0] + "'" +
                                            ",  GETDATE() " +
                                            ", N'" + UserInfos[0] + "'" +
                                            ",  0 " +
                                            ",  " + ankenID + " " +
                                            ")";

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "INSERT INTO KokyakuKeiyakuJouhou ( " +
                                            "KokyakuKeiyakuID " +
                                            ",KokyakuDeleteFlag " +
                                            ",KokyakuCreateProgram " +
                                            ",KokyakuCreateUser " +
                                            ",KokyakuCreateDate " +
                                            ",KokyakuUpdateUser " +
                                            ",KokyakuUpdateDate " +
                                            ",AnkenJouhouID " +
                                            ") VALUES (" +
                                            ankenID +
                                            ", 0" +
                                            ", 'InsertEntry'" +
                                            ", N'" + UserInfos[0] + "'" +
                                            ", GETDATE()" +
                                            ", N'" + UserInfos[0] + "'" +
                                            ", GETDATE()" +
                                            ", " + ankenID +
                                            ")";

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        transaction.Commit();

                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        throw;
                        return false;
                    }
                    conn.Close();
                    //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報を登録しました ID:" + ankenID, "InsertEntry", "");
                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報を登録しました ID:" + ankenID, pgmName + methodName, "");


                    // 売上年度が2021年以上の場合にフォルダを作成しにいく & 受託番号がない場合（この案件番号の枝番で新規作成、でない場合）
                    //if (GetInt(item1_3.SelectedValue.ToString()) >= 2021 && AnkenbaBangou == "")
                    // フォルダ作成関連は工期開始年度で行う
                    if (GetInt(item1_2_KoukiNendo.SelectedValue.ToString()) >= 2021 && AnkenbaBangou == "")
                    {
                        GlobalMethod.CreateFolder(ankenID);
                    }
                    else
                    {
                        //GlobalMethod.outputLogger("CreateFolder", "事業分類CD:" + jigyoubuHeadCD + " 年度:" + item1_3.Text + " の為フォルダ自動生成なし", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                        GlobalMethod.outputLogger("CreateFolder", "事業分類CD:" + jigyoubuHeadCD + " 年度:" + item1_2_KoukiNendo.Text + " の為フォルダ自動生成なし", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                    }

                    // Roleと部所をみて、参照モードか更新モードを切り替える
                    // Role:1管理者 で、部所がログインユーザーの部所と異なる場合は、参照モード
                    // Role:2:システム管理者の場合は、無条件に更新モード
                    string formmode;
                    if (UserInfos[4].Equals("2"))
                    {
                        formmode = "update";
                    }
                    else
                    {
                        // ログインユーザーの部所と一致しているかどうか
                        if (UserInfos[2] == item1_10.SelectedValue.ToString())
                        {
                            formmode = "";
                        }
                        else
                        {
                            formmode = "view";
                        }
                    }
                    Entry_Input form = new Entry_Input();
                    form.mode = formmode;
                    form.AnkenID = ankenID.ToString();
                    form.UserInfos = this.UserInfos;
                    form.Show(this.Owner);
                    ownerflg = false;
                    this.Close();
                }

            }
            // 1:更新 2:チェック用出力 3:起案
            else if (mode >= 1 && mode != 4)
            {
                string ankenNo = "";
                string ankenEda = "";
                string jutakubangou = "";
                string ori_ankenNo = "";

                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                // 案件番号
                ankenNo = item1_6.Text;
                ori_ankenNo = item1_6.Text;
                //えんとり君修正STEP2
                //案件番号変更機能
                if (sItem1_10_ori.Equals(item1_10.SelectedValue.ToString()) == false || sItem1_2_KoukiNendo_ori.Equals(item1_2_KoukiNendo.SelectedValue.ToString()) == false) { 
                    string jigyoubuHeadCD = "";
                    // 契約区分で業務分類CDを判定
                    // Mst_Jigyoubu に問い合わせる方法が無い為、
                    // 調査部が見つかった場合、T と判断
                    if (item1_14.Text.IndexOf("調査部") > -1)
                    {
                        jigyoubuHeadCD = "T";
                    }
                    else if (item1_14.Text.IndexOf("事業普及部") > -1)
                    {
                        jigyoubuHeadCD = "B";
                    }
                    else if (item1_14.Text.IndexOf("情シス部") > -1)
                    {
                        jigyoubuHeadCD = "J";
                    }
                    else if (item1_14.Text.IndexOf("総合研究所") > -1)
                    {
                        jigyoubuHeadCD = "K";
                    }
                    using (var conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        var cmd = conn.CreateCommand();

                        // 業務分類CD + 年度下2桁
                        ankenNo = jigyoubuHeadCD + item1_2_KoukiNendo.SelectedValue.ToString().Substring(2, 2);

                        // KashoShibuCD
                        cmd.CommandText = "SELECT  " +
                                "KashoShibuCD " +

                                //参照テーブル
                                "FROM Mst_Busho " +
                                "WHERE GyoumuBushoCD = '" + item1_10.SelectedValue.ToString() + "' ";
                        var sda = new SqlDataAdapter(cmd);
                        var dt = new DataTable();
                        sda.Fill(dt);
                        // KashoShibuCDが正しい
                        ankenNo = ankenNo + dt.Rows[0][0].ToString();

                        cmd.CommandText = "SELECT TOP 1 " +
                                " SUBSTRING(AnkenAnkenBangou,7,3) " +

                                //参照テーブル
                                "FROM AnkenJouhou " +
                                "WHERE AnkenAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + ankenNo + "%' and AnkenDeleteFlag != 1 ORDER BY AnkenAnkenBangou DESC";
                        sda = new SqlDataAdapter(cmd);
                        Console.WriteLine(cmd.CommandText);
                        dt = new DataTable();
                        sda.Fill(dt);

                        if (dt != null && dt.Rows.Count > 0)
                        {
                            int AnkenNoRenban;
                            if (int.TryParse(dt.Rows[0][0].ToString(), out AnkenNoRenban))
                            {
                                AnkenNoRenban++;
                            }
                            else
                            {
                                AnkenNoRenban = 1;
                            }
                            ankenNo += string.Format("{0:D3}", AnkenNoRenban);
                        }
                        else
                        {
                            ankenNo += "001";
                        }
                        // 受託フォルダも変更する
                        item1_6.Text = ankenNo;
                        Header1.Text = ankenNo;
                        // No.1422 1196 案件番号の変更履歴を保存する
                        item1_37_kojinCD.Text = UserInfos[0];
                        item1_37.Text = UserInfos[1];
                        item1_38.Text = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                        item1_39.Text = ori_ankenNo;
                    }
                }
                // 受託番号
                if (item1_7.Text != "")
                {
                    jutakubangou = item1_7.Text;
                }
                // 受託番号枝番
                if (item1_8.Text != "")
                {
                    ankenEda = item1_8.Text;
                }
                // 受託番号が空の場合
                if (item1_7.Text == "")
                {
                    // item2_3_7：落札者（落札状況にチェックが無いと入らないので、落札状況はチェックなしとする）
                    // ENTORY_TOUKAI:建設物価調査会
                    if (item2_3_7.Text == GlobalMethod.GetCommonValue2("ENTORY_TOUKAI"))
                    {
                        // 受託番号の枝番を取得する
                        using (var conn = new SqlConnection(connStr))
                        {
                            try
                            {
                                conn.Open();
                                var cmd = conn.CreateCommand();
                                // true:作成する false：作成しない
                                Boolean JutakuEdaCreateFlag = true;
                                cmd.CommandText = "SELECT  " +
                                    "TOP 1 AnkenJutakuBangouEda " +
                                    "FROM AnkenJouhou " +
                                    "WHERE AnkenJouhouID = " + AnkenID + " AND AnkenJutakuBangouEda <> '' and AnkenJutakuBangouEda != '' and AnkenDeleteFlag != 1 ";
                                var sda = new SqlDataAdapter(cmd);
                                var dt = new DataTable();
                                sda.Fill(dt);
                                if (dt != null && dt.Rows.Count > 0)
                                {
                                    // 存在するので作成なし
                                    JutakuEdaCreateFlag = false;
                                    ankenEda = dt.Rows[0][0].ToString(); // 枝番を取得できた場合、変数に入れておく
                                }

                                if (JutakuEdaCreateFlag == true)
                                {
                                    cmd.CommandText = "SELECT  " +
                                        "TOP 1 AnkenJutakuBangouEda " +
                                        "FROM AnkenJouhou " +
                                        "WHERE AnkenAnkenBangou = (SELECT AnkenAnkenBangou FROM AnkenJouhou WHERE AnkenJouhouID = " + AnkenID + ") and AnkenJutakuBangouEda != '' and AnkenDeleteFlag != 1 order by AnkenJutakuBangouEda desc ";
                                    sda = new SqlDataAdapter(cmd);
                                    dt = new DataTable();
                                    sda.Fill(dt);
                                    if (dt != null && dt.Rows.Count > 0)
                                    {
                                        // 枝番（-nn）を取得
                                        ankenEda = dt.Rows[0][0].ToString();
                                        // 「-」落とし
                                        //ankenEda = ankenEda.Replace("-", "");
                                        int i = int.Parse(ankenEda);
                                        i += 1;
                                        ankenEda = "" + string.Format("{0:D2}", i);
                                    }
                                    else
                                    {
                                        // 枝番（-01）をセット
                                        ankenEda = "01";
                                    }
                                }
                                dt.Clear();
                            }
                            catch (Exception)
                            {
                                // エラー
                                GlobalMethod.outputLogger("Execute_SQL", "枝番が取得できずにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                            }
                        }
                        // 受託番号の自動生成
                        jutakubangou = ankenNo + "-" + ankenEda;

                        // 引合タブの受託番号、枝番にセット
                        item1_7.Text = jutakubangou;
                        item1_8.Text = ankenEda;
                        Header2.Text = jutakubangou;
                    }
                }
                else
                {
                    // エントリ君修正STEP2
                    //・受託番号を解除する為、受注チェックを外すことにより、受託番号を削除する。
                    //・受託番号削除は、起案後には行えない。起案後に行う場合は、起案解除後に行える。（システム管理者のみ）
                    if (UserInfos[4].Equals("2")) { 
                        if (item3_1_2.Checked == false && (item2_3_7.Text == GlobalMethod.GetCommonValue2("ENTORY_TOUKAI")) == false)
                        {
                            string checkBangou = Header2.Text;
                            bool isDel = false;
                            //・窓口ミハルが登録されている場合は、エラーメッセージを表示し、削除が出来ない。
                            using (var conn = new SqlConnection(connStr))
                            {
                                try
                                {
                                    conn.Open();
                                    var cmd = conn.CreateCommand();
                                    cmd.CommandText = "SELECT  " +
                                        "TOP 1 MadoguchiJutakuBangou " +
                                        "FROM MadoguchiJouhou " +
                                        "WHERE AnkenJouhouID = " + AnkenID + " AND MadoguchiJutakuBangou = '" + ankenNo + "' AND MadoguchiJutakuBangouEdaban = '" + item1_8.Text + "' AND MadoguchiDeleteFlag != 1 ";
                                    var sda = new SqlDataAdapter(cmd);
                                    var dt = new DataTable();
                                    sda.Fill(dt);
                                    if (dt != null && dt.Rows.Count > 0)
                                    {
                                        set_error(GlobalMethod.GetMessage("E10726", ""));
                                        isDel = false;
                                    }
                                    else
                                    {
                                        isDel = true;
                                    }
                                    dt.Clear();
                                }
                                catch (Exception)
                                {
                                    isDel = false;
                                    // エラー
                                    GlobalMethod.outputLogger("Execute_SQL", "窓口ミハルが登録されているかチェックにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                                }
                            }
                            //・単価契約が登録されている場合は、エラーメッセージを表示し、削除が出来ない。
                            if (isDel == true)
                            {
                                using (var conn = new SqlConnection(connStr))
                                {
                                    try
                                    {
                                        conn.Open();
                                        var cmd = conn.CreateCommand();
                                        cmd.CommandText = "SELECT  " +
                                            "TOP 1 TankakeiyakuJutakuBangou " +
                                            "FROM TankaKeiyaku " +
                                            "WHERE AnkenJouhouID = " + AnkenID + " AND TankakeiyakuJutakuBangou = '" + checkBangou + "' and TankakeiyakuDeleteFlag != 1 ";
                                        var sda = new SqlDataAdapter(cmd);
                                        var dt = new DataTable();
                                        sda.Fill(dt);
                                        if (dt != null && dt.Rows.Count > 0)
                                        {
                                            set_error(GlobalMethod.GetMessage("E10727", ""));
                                            isDel = false;
                                        }
                                        else
                                        {
                                            isDel = true;
                                        }
                                        dt.Clear();
                                    }
                                    catch (Exception)
                                    {
                                        isDel = false;
                                        // エラー
                                        GlobalMethod.outputLogger("Execute_SQL", "単価契約が登録されているかチェックにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                                    }
                                }
                            }
                            if (isDel == true)
                            {
                                jutakubangou = "";
                                ankenEda = "";
                                // 引合タブの受託番号、枝番にセット
                                item1_7.Text = jutakubangou;
                                item1_8.Text = ankenEda;
                                Header2.Text = jutakubangou;
                            }

                        }
                    }
                }

                // えんとり君修正STEP2 フォルダリムーブ処理
                if (RenameFolder(ori_ankenNo))
                {
                    // 移動履歴LOG残す
                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "フォルダ変更前：" + GlobalMethod.ChangeSqlText(sFolderRenameBef, 0, 0) + "→フォルダ変更後：" + GlobalMethod.ChangeSqlText(item1_12.Text, 0, 0), pgmName + methodName, "");
                }
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {
                        cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                     "AnkenHikiaijhokyo = N'" + item1_1.SelectedValue + "'" +
                                     ",AnkenUriageNendo = " + item3_1_5.SelectedValue +
                                     ",AnkenSakuseiKubun = N'" + item3_1_1.SelectedValue + "'" +
                                     ",AnkenKeikakuBangou = N'" + item1_4.Text + "'" +
                                     ",AnkenAnkenBangou = " + "N'" + item1_6.Text + "'" +
                                     ",AnkenJutakubangou = " + "N'" + jutakubangou + "'" +
                                     ",AnkenJutakuBangouEda = " + "N'" + ankenEda + "'" +
                                     ",AnkenTourokubi = " + Get_DateTimePicker("item1_9") +
                                     ",AnkenJutakushibu = " + "N'" + item1_10.Text + "'" +
                                     ",AnkenJutakubushoCD = " + "N'" + item1_10.SelectedValue + "'" +
                                     ",AnkenKeiyakusho = N'" + GlobalMethod.ChangeSqlText(item1_12.Text, 0, 0) + "'" +
                                     ",AnkenTantoushaCD = N'" + item1_11_CD.Text + "'" +
                                     ",AnkenTantoushaMei = N'" + item1_11.Text + "'" +
                                     ",AnkenGyoumuMei = " + "N'" + GlobalMethod.ChangeSqlText(item1_13.Text, 0, 0) + "'" +
                                     ",AnkenGyoumuKubun = N'" + item1_14.SelectedValue + "'" +
                                     ",AnkenGyoumuKubunCD = N'" + Get_GyoumuKubunCD(item1_14.SelectedValue.ToString()) + "'" +
                                     ",AnkenGyoumuKubunMei = " + "N'" + item1_14.Text + "'" +
                                     ",AnkenNyuusatsuHoushiki = " + "N'" + item1_15.SelectedValue + "'" +
                                     ",AnkenNyuusatsuYoteibi = " + Get_DateTimePicker("item1_16") +
                                     ",AnkenHachushaCD = " + "N'" + item1_19.Text + "'" +
                                     ",AnkenHachuushaMei = " + "N'" + item1_23.Text + "'" +
                                     ",AnkenHachushaKaMei = " + "N'" + GlobalMethod.ChangeSqlText(item1_24.Text, 0, 0) + "'" +
                                     ",AnkenHachuushaKaMei = " + "N'" + GlobalMethod.ChangeSqlText(item1_23.Text, 0, 0) + "　" + GlobalMethod.ChangeSqlText(item1_24.Text, 0, 0) + "'" +
                                     ",AnkenHachuushaKeiyakuBusho = " + "N'" + GlobalMethod.ChangeSqlText(item1_25.Text, 0, 0) + "'" +
                                    ",AnkenHachuushaKeiyakuTantou = " + "N'" + GlobalMethod.ChangeSqlText(item1_26.Text, 0, 0) + "'" +
                                    ",AnkenHachuushaKeiyakuTEL = " + "N'" + GlobalMethod.ChangeSqlText(item1_27.Text, 0, 0) + "'" +
                                    ",AnkenHachuushaKeiyakuFAX = " + "N'" + GlobalMethod.ChangeSqlText(item1_28.Text, 0, 0) + "'" +
                                    ",AnkenHachuushaKeiyakuMail = " + "N'" + GlobalMethod.ChangeSqlText(item1_29.Text, 0, 0) + "'" +
                                    ",AnkenHachuushaKeiyakuYuubin = " + "N'" + GlobalMethod.ChangeSqlText(item1_30.Text, 0, 0) + "'" +
                                    ",AnkenHachuushaKeiyakuJuusho = " + "N'" + GlobalMethod.ChangeSqlText(item1_31.Text, 0, 0) + "'" +
                                    ",AnkenHachuuDaihyouYakushoku = " + "N'" + GlobalMethod.ChangeSqlText(item1_32.Text, 0, 0) + "'" +
                                    ",AnkenHachuuDaihyousha = " + "N'" + GlobalMethod.ChangeSqlText(item1_33.Text, 0, 0) + "'" +
                                    ",AnkenToukaiSankouMitsumori = " + "N'" + item1_34.SelectedValue + "'" +
                                    ",AnkenToukaiJyutyuIyoku = " + "N'" + item1_35.SelectedValue + "'" +
                                    ",AnkenToukaiSankouMitsumoriGaku = " + item1_36.Text.Replace("¥", "").Replace(",", "") + " " +
                                    ",AnkenUpdateProgram = " + "'UpdateEntry'" +
                                    ",AnkenUpdateDate = " + "GETDATE()" +
                                    ",AnkenUpdateUser = " + "N'" + UserInfos[0] + "'" +
                                    ",AnkenToukaiOusatu = " + "N'" + item2_2_1.SelectedValue + "'" +
                                    ",AnkenKianzumi = " + (item3_1_2.Checked ? 1 : 0) +
                                    ",AnkenKeiyakuKoukiKaishibi = " + Get_DateTimePicker("item3_1_6") +
                                    ",AnkenKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("item3_1_7") +
                                    ",GyoumuKanrishaCD = " + "N'" + item3_4_4_CD.Text + "'" +
                                    ",GyoumuKanrishaMei = " + "N'" + item3_4_4.Text + "'" +
                                    ",AnkenKokyakuHyoukaComment = " + "N'" + GlobalMethod.ChangeSqlText(item4_1_9.Text, 0, 0) + "'" +
                                    ",AnkenToukaiHyoukaComment = " + "N'" + GlobalMethod.ChangeSqlText(item4_1_10.Text, 0, 0) + "'" +
                                    ",AnkenKoukiNendo = " + "'" + item1_2_KoukiNendo.SelectedValue.ToString() + "' " +
                                    //えんとり君修正STEP２
                                    ",AnkenFolderHenkouDatetime = " + (string.IsNullOrEmpty(item1_38.Text) ? "NULL " : "'" + item1_38.Text + "' ") +
                                    ",AnkenFolderHenkouTantoushaCD = '" + item1_37_kojinCD.Text + "' " +
                                    // No.1422 1196 案件番号の変更履歴を保存する
                                    ",AnkenHenkoumaeAnkenBangou = " + (string.IsNullOrEmpty(item1_39.Text) ? "NULL " : "'" + item1_39.Text + "' ") +
                                    " WHERE AnkenJouhouID = " + AnkenID;

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        // 計画情報の案件数を設定しなおし
                        cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_4.Text + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_4.Text + "' ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + beforeKeikakuBangou + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + beforeKeikakuBangou + "' ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();


                        cmd.CommandText = "DELETE FROM AnkenJouhouZenkaiRakusatsu " +
                                    " WHERE AnkenJouhouID = " + AnkenID;
                        cmd.ExecuteNonQuery();

                        for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                        {
                            if (c1FlexGrid1.Rows[i][3] != null && c1FlexGrid1.Rows[i][3].ToString() != "")
                            {
                                string AnkenZenkaiRakusatsushaID = "";
                                string AnkenZenkaiJutakuKingaku = "";
                                string KeiyakuZenkaiRakusatsushaID = "";
                                string AnkenZenkaiAnkenJouhouID = "";
                                string AnkenZenkaiJutakuZeinuki = "";

                                // 前回落札者ID
                                if (c1FlexGrid1.Rows[i][8].ToString() != "")
                                {
                                    AnkenZenkaiRakusatsushaID = c1FlexGrid1.Rows[i][8].ToString().Trim();
                                }
                                else
                                {
                                    AnkenZenkaiRakusatsushaID = "null";
                                }
                                // 前回受託金額
                                if (c1FlexGrid1.Rows[i][9].ToString() != "")
                                {
                                    AnkenZenkaiJutakuKingaku = c1FlexGrid1.Rows[i][9].ToString();
                                }
                                else
                                {
                                    AnkenZenkaiJutakuKingaku = "null";
                                }
                                // 前回契約落札者ID
                                if (c1FlexGrid1.Rows[i][14].ToString().Trim() != "")
                                {
                                    KeiyakuZenkaiRakusatsushaID = c1FlexGrid1.Rows[i][14].ToString().Trim();
                                }
                                else
                                {
                                    KeiyakuZenkaiRakusatsushaID = "null";
                                }
                                // 前回契約情報ID
                                if (c1FlexGrid1.Rows[i][2].ToString().Trim() != "")
                                {
                                    AnkenZenkaiAnkenJouhouID = c1FlexGrid1.Rows[i][2].ToString().Trim();
                                }
                                else
                                {
                                    AnkenZenkaiAnkenJouhouID = "null";
                                }
                                // 前回受託税抜
                                if (c1FlexGrid1.Rows[i][13].ToString() != "")
                                {
                                    AnkenZenkaiJutakuZeinuki = c1FlexGrid1.Rows[i][13].ToString();
                                }
                                else
                                {
                                    AnkenZenkaiJutakuZeinuki = "null";
                                }

                                cmd.CommandText = "INSERT INTO AnkenJouhouZenkaiRakusatsu ( " +
                                                    "AnkenJouhouID " +
                                                    ",AnkenZenkaiRakusatsuID " +
                                                    ",AnkenZenkaiAnkenBangou " +
                                                    ",AnkenZenkaiJutakuBangou " +
                                                    ",AnkenZenkaiJutakuEdaban " +
                                                    ",AnkenZenkaiGyoumuMei " +

                                                    ",AnkenZenkaiRakusatsusha " +
                                                    ",AnkenZenkaiRakusatsushaID " +

                                                    ",AnkenZenkaiJutakuKingaku " +
                                                    ",KeiyakuZenkaiRakusatsushaID " +
                                                    ",AnkenZenkaiKyougouKigyouCD " +
                                                    ",AnkenZenkaiAnkenJouhouID " +
                                                    ",AnkenZenkaiJutakuZeinuki " +
                                                    " ) VALUES ( " +
                                                     AnkenID +                                             // [AnkenJouhouID] [decimal](16, 0) NOT NULL,
                                                                                                           //", " + i +                                            // [AnkenZenkaiRakusatsuID] [int] NOT NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][16].ToString() + "'" +    // [AnkenZenkaiRakusatsuID] [int] NOT NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][3].ToString() + "'" +     // [AnkenZenkaiAnkenBangou] [nvarchar](40) NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][4].ToString() + "'" +     // [AnkenZenkaiJutakuBangou] [nvarchar](40) NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][5].ToString() + "'" +     // [AnkenZenkaiJutakuEdaban] [nvarchar](50) NULL,
                                                     ", N'" + c1FlexGrid1.Rows[i][6].ToString() + "'" +     // [AnkenZenkaiGyoumuMei] [nvarchar](150) NULL,

                                                     ", N'" + c1FlexGrid1.Rows[i][7].ToString() + "'" +     // [AnkenZenkaiRakusatsusha] [nvarchar](50) NULL,
                                                     ", " + AnkenZenkaiRakusatsushaID + "" +               // [AnkenZenkaiRakusatsushaID] [int] NULL,
                                                     ", " + AnkenZenkaiJutakuKingaku + "" +       // [AnkenZenkaiJutakuKingaku] [money] NULL,
                                                     ", " + KeiyakuZenkaiRakusatsushaID + "" +             // [KeiyakuZenkaiRakusatsushaID] [int] NULL,        
                                                     ", N'" + c1FlexGrid1.Rows[i][15].ToString() + "'" +    // [AnkenZenkaiKyougouKigyouCD] [nvarchar](24) NULL,
                                                     ", " + AnkenZenkaiAnkenJouhouID + "" +                // [AnkenZenkaiAnkenJouhouID] [decimal](16, 0) NULL,
                                                     ", " + AnkenZenkaiJutakuZeinuki + "" +      // [AnkenZenkaiJutakuZeinuki] [money] NULL,
                                                    ")";

                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();
                            }

                        }

                        if (!GlobalMethod.Check_Table(AnkenID, "AnkenJouhouID", "KeiyakuJouhouEntory", ""))
                        {
                            GlobalMethod.outputLogger("UpdateEntory", "契約情報(エントリー)更新 データなしエラー", AnkenID, UserInfos[1]);
                            set_error(GlobalMethod.GetMessage("E10009", "(契約情報(エントリー)データなし)"));
                            transaction.Rollback();
                            return false;
                        }

                        cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET " +
                                    "KeiyakuKeiyakuTeiketsubi = " + Get_DateTimePicker("item3_1_3") +
                                    ",KeiyakuSakuseibi = " + Get_DateTimePicker("item3_1_4") +
                                    ",KeiyakuKeiyakuKoukiKaishibi = " + Get_DateTimePicker("item3_1_6") +
                                    ",KeiyakuKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("item3_1_7") +
                                    ",KeiyakuKeiyakuKingaku = " + item3_1_12.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuZeikomiKingaku = " + item3_1_13.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuuchizeiKingaku = " + item3_1_14.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuShouhizeiritsu = N'" + item3_1_10.Text + "'" +
                                    ",KeiyakuHenkouChuushiRiyuu = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_17.Text, 0, 0) + "'" +
                                    ",KeiyakuBikou = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_19.Text, 0, 0) + "'" +
                                    ",KeiyakuShosha = " + (item3_1_20.Checked ? 1 : 0) +
                                    ",KeiyakuTokkiShiyousho = " + (item3_1_21.Checked ? 1 : 0) +
                                    ",KeiyakuMitsumorisho = " + (item3_1_22.Checked ? 1 : 0) +
                                    ",KeiyakuTanpinChousaMitsumorisho = " + (item3_1_23.Checked ? 1 : 0) +
                                    ",KeiyakuSonota = " + (item3_1_24.Checked ? 1 : 0) +
                                    ",KeiyakuSonotaNaiyou = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_25.Text, 0, 0) + "'" +
                                    ",KeiyakuZentokinUkewatashibi = " + Get_DateTimePicker("item3_6_11") +
                                    ",KeiyakuZentokin = " + item3_6_12.Text.Replace("¥", "").Replace(",", "") +
                                    ",Keiyakukeiyakukingakukei = " + item3_1_15.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuBetsuKeiyakuKingaku = " + item3_1_16.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi1 = " + " " + Get_DateTimePicker("item3_6_1") + "" +
                                    ",KeiyakuSeikyuuKingaku1 = " + item3_6_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi2 = " + " " + Get_DateTimePicker("item3_6_3") + "" +
                                    ",KeiyakuSeikyuuKingaku2 = " + item3_6_4.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi3 = " + " " + Get_DateTimePicker("item3_6_5") + "" +
                                    ",KeiyakuSeikyuuKingaku3 = " + item3_6_6.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi4 = " + " " + Get_DateTimePicker("item3_6_7") + "" +
                                    ",KeiyakuSeikyuuKingaku4 = " + item3_6_8.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi5 = " + " " + Get_DateTimePicker("item3_6_9") + "" +
                                    ",KeiyakuSeikyuuKingaku5 = " + item3_6_10.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSakuseiKubunID = " + "N'" + item3_1_1.SelectedValue + "'" +
                                    ",KeiyakuSakuseiKubun = " + "N'" + item3_1_1.Text + "'" +
                                    ",KeiyakuGyoumuKubun = " + "N'" + item1_14.SelectedValue + "'" +
                                    ",KeiyakuGyoumuMei = " + "N'" + item1_14.Text + "'" +
                                    ",KeiyakuJutakubangou = " + "N'" + item1_7.Text + "'" +
                                    ",KeiyakuEdaban = " + "N'" + item1_8.Text + "'" +
                                    ",KeiyakuKianzumi = " + (item3_1_2.Checked ? 1 : 0) +
                                    ",KeiyakuHachuushaMei = " + "N'" + item3_1_9.Text + "'" +
                                    ",KeiyakuHaibunChoZeinuki = " + item3_2_1_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuHaibunJoZeinuki = " + item3_2_2_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuHaibunJosysZeinuki = " + item3_2_3_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuHaibunKeiZeinuki = " + item3_2_4_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuHaibunZeinukiKei = " + item3_2_5_2.Text.Replace("¥", "").Replace(",", "").Replace("%", "") +
                                    ",KeiyakuUriageHaibunCho  = " + item3_2_1_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUriageHaibunJo   = " + item3_2_2_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUriageHaibunJosys  = " + item3_2_3_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUriageHaibunKei  = " + item3_2_4_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUriageHaibunGoukei = " + item3_2_5_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuTankeiMikomiCho  = " + item3_3_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuTankeiMikomiJo  = " + item3_3_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuTankeiMikomiJosys  = " + item3_3_3.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuTankeiMikomiKei  = " + item3_3_4.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuKurikoshiCho  = " + item3_7_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuKurikoshiJo  = " + item3_7_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuKurikoshiJosys  = " + item3_7_3.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuKurikoshiKei  = " + item3_7_4.Text.Replace("¥", "").Replace(",", "") +
                                    // えんとり君修正STEP2（RIBC項目追加）
                                    ",KeiyakuRIBCYouTankaDataMoushikomisho = " + (item3_ribc_price.Checked ? 1 : 0) +
                                    ",KeiyakuSashaKeiyu = " + (item3_sa_commpany.Checked ? 1 : 0) +
                                    ",KeiyakuRIBCYouTankaData = " + (item3_1_ribc.Checked ? 1 : 0) +
                                    ",KeiyakuUpdateProgram = " + "'UpdateEntry'" +
                                    ",KeiyakuUpdateDate = " + "GETDATE()" +
                                    ",KeiyakuUpdateUser = " + "N'" + UserInfos[0] + "'" +
                                    " WHERE AnkenJouhouID = " + AnkenID;

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        GlobalMethod.outputLogger("InsertAnken", "契約情報更新時に案件情報を同時更新します AnkenjouhoID = " + AnkenID, "", UserInfos[1]);

                        cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                    "AnkenKeiyakuKoukiKaishibi = " + Get_DateTimePicker("item3_1_6") +
                                    ",AnkenKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("item3_1_7") +
                                    ",AnkenKeiyakuTeiketsubi = " + Get_DateTimePicker("item3_1_3") +
                                    ",AnkenKeiyakuZeikomiKingaku = " + item3_1_13.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuC = " + item3_2_1_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuJ = " + item3_2_2_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuJs = " + item3_2_3_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuK = " + item3_2_4_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuSakuseibi = " + Get_DateTimePicker("item3_1_4") +
                                    " WHERE AnkenJouhouID = " + AnkenID;
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM RibcJouhou " +
                                    " WHERE RibcID = " + AnkenID;
                        cmd.ExecuteNonQuery();
                        int cnt = 0;
                        string RibcKoukiStart;
                        string RibcNouhinbi;
                        string RibcSeikyubi;
                        string RibcNyukinyoteibi;
                        string RibcKubun;
                        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                        {
                            // 新では計上額のみでも登録を可とする
                            //if (c1FlexGrid4.Rows[i][1] != null)
                            //{

                            // 計上日、計上月、計上額のどれかが入っていれば登録する
                            // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                            if ((c1FlexGrid4.Rows[i][1] != null && c1FlexGrid4.Rows[i][1] != "")
                                //|| (c1FlexGrid4.Rows[i][2] != null && c1FlexGrid4.Rows[i][2] != "" )
                                || (c1FlexGrid4.Rows[i][3] != null && c1FlexGrid4.Rows[i][3].ToString() != "0"))
                            {
                                RibcKoukiStart = "null";
                                if (c1FlexGrid4.Rows[i][4] != null)
                                {
                                    RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][4].ToString() + "'";
                                }
                                RibcNouhinbi = "null";
                                if (c1FlexGrid4.Rows[i][5] != null)
                                {
                                    RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][5].ToString() + "'";
                                }
                                RibcSeikyubi = "null";
                                if (c1FlexGrid4.Rows[i][6] != null)
                                {
                                    RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][6].ToString() + "'";
                                }
                                RibcNyukinyoteibi = "null";
                                if (c1FlexGrid4.Rows[i][7] != null)
                                {
                                    RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][7].ToString() + "'";
                                }
                                RibcKubun = "0";
                                if (c1FlexGrid4.Rows[i][8] != null)
                                {
                                    RibcKubun = "N'" + c1FlexGrid4.Rows[i][8].ToString() + "'";
                                }

                                cnt++;

                                //// 売上計上情報で、工期末日付、計上月、計上額、どれかが空だと落ちる
                                //if (c1FlexGrid4.Rows[i][1] != null && c1FlexGrid4.Rows[i][2] != null && c1FlexGrid4.Rows[i][3] != null)
                                //{

                                cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                            "RibcID " +
                                            ",RibcNo " +
                                            ",RibcKoukiEnd " +
                                            ",RibcUriageKeijyoTuki " +
                                            ",RibcSeikyuKingaku " +
                                            ",RibcKankeibusho " +
                                            ",RibcKoukiStart " +
                                            ",RibcNouhinbi " +
                                            ",RibcSeikyubi " +
                                            ",RibcNyukinyoteibi " +
                                            ",RibcKubun " +
                                            ") VALUES (" +
                                            AnkenID +
                                            "," + cnt + "";
                                // 工期末日付 RibcKoukiEnd
                                if (c1FlexGrid4.Rows[i][1] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][1].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",null";
                                }
                                // 計上月 RibcUriageKeijyoTuki
                                if (c1FlexGrid4.Rows[i][2] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][2].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'' " + "";
                                }
                                // 計上額 RibcSeikyuKingaku
                                if (c1FlexGrid4.Rows[i][3] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][3].ToString().Replace("¥", "").Replace(",", "") + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                }
                                cmd.CommandText = cmd.CommandText +
                                ",'127120' " +
                                "," + RibcKoukiStart +
                                "," + RibcNouhinbi +
                                "," + RibcSeikyubi +
                                "," + RibcNyukinyoteibi +
                                "," + RibcKubun +
                                ")";
                                //}
                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();
                            }
                            //}

                            // 新では計上額のみでも登録を可とする
                            //if (c1FlexGrid4.Rows[i][9] != null)
                            //{
                            // 計上日、計上月、計上額のどれかが入っていれば登録する
                            // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                            if ((c1FlexGrid4.Rows[i][9] != null && c1FlexGrid4.Rows[i][9] != "")
                                //|| (c1FlexGrid4.Rows[i][10] != null && c1FlexGrid4.Rows[i][10] != "")
                                || (c1FlexGrid4.Rows[i][11] != null && c1FlexGrid4.Rows[i][11].ToString() != "0"))
                            {
                                RibcKoukiStart = "null";
                                if (c1FlexGrid4.Rows[i][12] != null)
                                {
                                    RibcKoukiStart = "'" + c1FlexGrid4.Rows[i][12].ToString() + "'";
                                }
                                RibcNouhinbi = "null";
                                if (c1FlexGrid4.Rows[i][13] != null)
                                {
                                    RibcNouhinbi = "'" + c1FlexGrid4.Rows[i][13].ToString() + "'";
                                }
                                RibcSeikyubi = "null";
                                if (c1FlexGrid4.Rows[i][14] != null)
                                {
                                    RibcSeikyubi = "'" + c1FlexGrid4.Rows[i][14].ToString() + "'";
                                }
                                RibcNyukinyoteibi = "null";
                                if (c1FlexGrid4.Rows[i][15] != null)
                                {
                                    RibcNyukinyoteibi = "'" + c1FlexGrid4.Rows[i][15].ToString() + "'";
                                }
                                RibcKubun = "0";
                                if (c1FlexGrid4.Rows[i][16] != null)
                                {
                                    RibcKubun = "'" + c1FlexGrid4.Rows[i][16].ToString() + "'";
                                }
                                cnt++;
                                cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                            "RibcID " +
                                            ",RibcNo " +
                                            ",RibcKoukiEnd " +
                                            ",RibcUriageKeijyoTuki " +
                                            ",RibcSeikyuKingaku " +
                                            ",RibcKankeibusho " +
                                            ",RibcKoukiStart " +
                                            ",RibcNouhinbi " +
                                            ",RibcSeikyubi " +
                                            ",RibcNyukinyoteibi " +
                                            ",RibcKubun " +
                                            ") VALUES (" +
                                            AnkenID +
                                            "," + cnt +
                                             //",'" + c1FlexGrid4.Rows[i][9].ToString() + "'" +
                                             //",'" + c1FlexGrid4.Rows[i][10].ToString() + "'" +
                                             //",'" + c1FlexGrid4.Rows[i][11].ToString().Replace("¥", "").Replace(",", "") + "'" +
                                             "";
                                // 工期末日付 RibcKoukiEnd
                                if (c1FlexGrid4.Rows[i][9] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][9].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",null";
                                }
                                // 計上月 RibcUriageKeijyoTuki
                                if (c1FlexGrid4.Rows[i][10] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][10].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'' " + "";
                                }
                                // 計上額 RibcSeikyuKingaku
                                if (c1FlexGrid4.Rows[i][11] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][11].ToString().Replace("¥", "").Replace(",", "") + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                }
                                cmd.CommandText = cmd.CommandText + ",'129230' " +
                                            "," + RibcKoukiStart +
                                            "," + RibcNouhinbi +
                                            "," + RibcSeikyubi +
                                            "," + RibcNyukinyoteibi +
                                            "," + RibcKubun +
                                            ")";

                                cmd.ExecuteNonQuery();
                            }
                            //}

                            // 計上日、計上月、計上額のどれかが入っていれば登録する
                            // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                            if ((c1FlexGrid4.Rows[i][17] != null && c1FlexGrid4.Rows[i][17] != "")
                                //|| (c1FlexGrid4.Rows[i][18] != null && c1FlexGrid4.Rows[i][18] != "")
                                || (c1FlexGrid4.Rows[i][19] != null && c1FlexGrid4.Rows[i][19].ToString() != "0"))
                            {
                                // 新では計上額のみでも登録を可とする
                                //if (c1FlexGrid4.Rows[i][17] != null)
                                //{
                                RibcKoukiStart = "null";
                                if (c1FlexGrid4.Rows[i][20] != null)
                                {
                                    RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][20].ToString() + "'";
                                }
                                RibcNouhinbi = "null";
                                if (c1FlexGrid4.Rows[i][21] != null)
                                {
                                    RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][21].ToString() + "'";
                                }
                                RibcSeikyubi = "null";
                                if (c1FlexGrid4.Rows[i][22] != null)
                                {
                                    RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][22].ToString() + "'";
                                }
                                RibcNyukinyoteibi = "null";
                                if (c1FlexGrid4.Rows[i][23] != null)
                                {
                                    RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][23].ToString() + "'";
                                }
                                RibcKubun = "0";
                                if (c1FlexGrid4.Rows[i][24] != null)
                                {
                                    RibcKubun = "N'" + c1FlexGrid4.Rows[i][24].ToString() + "'";
                                }
                                cnt++;
                                cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                            "RibcID " +
                                            ",RibcNo " +
                                            ",RibcKoukiEnd " +
                                            ",RibcUriageKeijyoTuki " +
                                            ",RibcSeikyuKingaku " +
                                            ",RibcKankeibusho " +
                                            ",RibcKoukiStart " +
                                            ",RibcNouhinbi " +
                                            ",RibcSeikyubi " +
                                            ",RibcNyukinyoteibi " +
                                            ",RibcKubun " +
                                            ") VALUES (" +
                                            AnkenID +
                                            "," + cnt + "";
                                //",'" + c1FlexGrid4.Rows[i][17].ToString() + "'" +
                                //",'" + c1FlexGrid4.Rows[i][18].ToString() + "'" +
                                //",'" + c1FlexGrid4.Rows[i][19].ToString().Replace("¥", "").Replace(",", "") + "'" +
                                // 工期末日付 RibcKoukiEnd
                                if (c1FlexGrid4.Rows[i][17] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][17].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",null";
                                }
                                // 計上月 RibcUriageKeijyoTuki
                                if (c1FlexGrid4.Rows[i][18] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][18].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'' " + "";
                                }
                                // 計上額 RibcSeikyuKingaku
                                if (c1FlexGrid4.Rows[i][19] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][19].ToString().Replace("¥", "").Replace(",", "") + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                }
                                // 年度により、情報システム部の部コードを変更する
                                if (GetInt(item3_1_5.SelectedValue.ToString()) >= 2021)
                                {
                                    cmd.CommandText = cmd.CommandText + ",'128400' " +
                                                "," + RibcKoukiStart +
                                                "," + RibcNouhinbi +
                                                "," + RibcSeikyubi +
                                                "," + RibcNyukinyoteibi +
                                                "," + RibcKubun +
                                                ")";
                                }
                                else
                                {
                                    // 2021年度以前は127900
                                    cmd.CommandText = cmd.CommandText + ",'127900' " +
                                                "," + RibcKoukiStart +
                                                "," + RibcNouhinbi +
                                                "," + RibcSeikyubi +
                                                "," + RibcNyukinyoteibi +
                                                "," + RibcKubun +
                                                ")";
                                }
                                cmd.ExecuteNonQuery();
                            }
                            //}

                            // 新では計上額のみでも登録を可とする
                            //if (c1FlexGrid4.Rows[i][25] != null)
                            //{
                            // 計上日、計上月、計上額のどれかが入っていれば登録する
                            // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                            if ((c1FlexGrid4.Rows[i][25] != null && c1FlexGrid4.Rows[i][25] != "")
                                //|| (c1FlexGrid4.Rows[i][26] != null && c1FlexGrid4.Rows[i][26] != "")
                                || (c1FlexGrid4.Rows[i][27] != null && c1FlexGrid4.Rows[i][27].ToString() != "0"))
                            {
                                RibcKoukiStart = "null";
                                if (c1FlexGrid4.Rows[i][28] != null)
                                {
                                    RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][28].ToString() + "'";
                                }
                                RibcNouhinbi = "null";
                                if (c1FlexGrid4.Rows[i][29] != null)
                                {
                                    RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][29].ToString() + "'";
                                }
                                RibcSeikyubi = "null";
                                if (c1FlexGrid4.Rows[i][30] != null)
                                {
                                    RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][30].ToString() + "'";
                                }
                                RibcNyukinyoteibi = "null";
                                if (c1FlexGrid4.Rows[i][31] != null)
                                {
                                    RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][31].ToString() + "'";
                                }
                                RibcKubun = "0";
                                if (c1FlexGrid4.Rows[i][32] != null)
                                {
                                    RibcKubun = "N'" + c1FlexGrid4.Rows[i][32].ToString() + "'";
                                }
                                cnt++;
                                cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                            "RibcID " +
                                            ",RibcNo " +
                                            ",RibcKoukiEnd " +
                                            ",RibcUriageKeijyoTuki " +
                                            ",RibcSeikyuKingaku " +
                                            ",RibcKankeibusho " +
                                            ",RibcKoukiStart " +
                                            ",RibcNouhinbi " +
                                            ",RibcSeikyubi " +
                                            ",RibcNyukinyoteibi " +
                                            ",RibcKubun " +
                                            ") VALUES (" +
                                            AnkenID +
                                            "," + cnt + "";
                                //",'" + c1FlexGrid4.Rows[i][25].ToString() + "'" +
                                //",'" + c1FlexGrid4.Rows[i][26].ToString() + "'" +
                                //",'" + c1FlexGrid4.Rows[i][27].ToString().Replace("¥", "").Replace(",", "") + "'" +
                                if (c1FlexGrid4.Rows[i][25] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][25].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",null";
                                }
                                // 計上月 RibcUriageKeijyoTuki
                                if (c1FlexGrid4.Rows[i][26] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][26].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'' " + "";
                                }
                                // 計上額 RibcSeikyuKingaku
                                if (c1FlexGrid4.Rows[i][27] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][27].ToString().Replace("¥", "").Replace(",", "") + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                }

                                cmd.CommandText = cmd.CommandText + ",'150200' " +
                                            "," + RibcKoukiStart +
                                            "," + RibcNouhinbi +
                                            "," + RibcSeikyubi +
                                            "," + RibcNyukinyoteibi +
                                            "," + RibcKubun +
                                            ")";

                                cmd.ExecuteNonQuery();
                            }
                        }
                        //}

                        //業務情報
                        cmd.CommandText = "UPDATE GyoumuJouhou SET " +
                                    "GyoumuHyouten = " + "N'" + item4_1_1.Text + "'" +
                                    ",KanriGijutsushaCD = " + "N'" + item3_4_1_CD.Text + "'" +
                                    ",KanriGijutsushaNM = " + "N'" + item3_4_1.Text + "'" +
                                    ",GyoumuKanriHyouten = " + "N'" + item4_1_3.Text + "'" +
                                    ",ShousaTantoushaCD = " + "N'" + item3_4_2_CD.Text + "'" +
                                    ",ShousaTantoushaNM = " + "N'" + item3_4_2.Text + "'" +
                                    ",GyoumuShousaHyouten = " + "N'" + item4_1_5.Text + "'" +
                                    ",SinsaTantoushaCD = " + "N'" + item3_4_3_CD.Text + "'" +
                                    ",SinsaTantoushaNM = " + "N'" + item3_4_3.Text + "'" +
                                    ",GyoumuTECRISTourokuBangou = " + "N'" + item4_1_6.Text + "'" +
                                    ",GyoumuKeisaiTankaTeikyou = " + "''" +
                                    ",GyoumuChosakukenJouto = " + "''" +
                                    ",GyoumuSeikyuubi = " + Get_DateTimePicker("item4_1_7") +
                                    ",GyoumuSeikyuusho = " + "N'" + GlobalMethod.ChangeSqlText(item4_1_8.Text, 0, 0) + "'" +
                                    ",GyoumuHikiwatashiNaiyou = " + "''" +
                                    ",GyoumuUpdateDate = " + " GETDATE() " +
                                    ",GyoumuUpdateUser = " + "N'" + UserInfos[0] + "' " +
                                    ",GyoumuUpdateProgram = " + "'UpdateEntory' " +
                                    ",GyoumuDeleteFlag = " + "0 " +
                                    " WHERE AnkenJouhouID = " + AnkenID;
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        //業務配分
                        //業務配分
                        cmd.CommandText = "UPDATE GyoumuHaibun SET " +
                                            " GyoumuChosaBuRitsu = " + item1_7_1_1_1.Text.Replace("%", "") + " " +
                                            ",GyoumuJigyoFukyuBuRitsu = " + item1_7_1_2_1.Text.Replace("%", "") + " " +
                                            ",GyoumuJyohouSystemBuRitsu = " + item1_7_1_3_1.Text.Replace("%", "") + " " +
                                            ",GyoumuSougouKenkyuJoRitsu = " + item1_7_1_4_1.Text.Replace("%", "") + " " +
                                            ",GyoumuShizaiChousaRitsu = " + item1_7_2_1_1.Text.Replace("%", "") + " " +
                                            ",GyoumuEizenRitsu = " + item1_7_2_2_1.Text.Replace("%", "") + " " +
                                            ",GyoumuKikiruiChousaRitsu = " + item1_7_2_3_1.Text.Replace("%", "") + " " +
                                            ",GyoumuKoujiChousahiRitsu = " + item1_7_2_4_1.Text.Replace("%", "") + " " +
                                            ",GyoumuSanpaiFukusanbutsuRitsu = " + item1_7_2_5_1.Text.Replace("%", "") + " " +
                                            ",GyoumuHokakeChousaRitsu = " + item1_7_2_6_1.Text.Replace("%", "") + " " +
                                            ",GyoumuShokeihiChousaRitsu = " + item1_7_2_7_1.Text.Replace("%", "") + " " +
                                            ",GyoumuGenkaBunsekiRitsu = " + item1_7_2_8_1.Text.Replace("%", "") + " " +
                                            ",GyoumuKijunsakuseiRitsu = " + item1_7_2_9_1.Text.Replace("%", "") + " " +
                                            ",GyoumuKoukyouRoumuhiRitsu = " + item1_7_2_10_1.Text.Replace("%", "") + " " +
                                            ",GyoumuRoumuhiKoukyouigaiRitsu = " + item1_7_2_11_1.Text.Replace("%", "") + " " +
                                            ",GyoumuSonotaChousabuRitsu = " + item1_7_2_12_1.Text.Replace("%", "") + " " +
                                            " WHERE GyoumuAnkenJouhouID = " + AnkenID + " AND GyoumuHibunKubun = 10 ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE GyoumuHaibun SET " +
                                            " GyoumuChosaBuGaku " + " = " + item3_7_1_6_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuJigyoFukyuBuGaku " + " = " + item3_7_1_7_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuJyohouSystemBuGaku " + " = " + item3_7_1_8_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuSougouKenkyuJoGaku " + " = " + item3_7_1_9_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuShizaiChousaRitsu " + " = " + item3_7_2_14_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuEizenRitsu " + " = " + item3_7_2_15_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuKikiruiChousaRitsu " + " = " + item3_7_2_16_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuKoujiChousahiRitsu " + " = " + item3_7_2_17_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuSanpaiFukusanbutsuRitsu " + " = " + item3_7_2_18_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuHokakeChousaRitsu " + " = " + item3_7_2_19_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuShokeihiChousaRitsu " + " = " + item3_7_2_20_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuGenkaBunsekiRitsu " + " = " + item3_7_2_21_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuKijunsakuseiRitsu " + " = " + item3_7_2_22_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuKoukyouRoumuhiRitsu " + " = " + item3_7_2_23_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuRoumuhiKoukyouigaiRitsu " + " = " + item3_7_2_24_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuSonotaChousabuRitsu " + " = " + item3_7_2_25_1.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuShizaiChousaGaku " + " = " + item3_7_2_14_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuEizenGaku " + " = " + item3_7_2_15_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuKikiruiChousaGaku " + " = " + item3_7_2_16_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuKoujiChousahiGaku " + " = " + item3_7_2_17_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuSanpaiFukusanbutsuGaku " + " = " + item3_7_2_18_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuHokakeChousaGaku " + " = " + item3_7_2_19_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuShokeihiChousaGaku " + " = " + item3_7_2_20_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuGenkaBunsekiGaku " + " = " + item3_7_2_21_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuKijunsakuseiGaku " + " = " + item3_7_2_22_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuKoukyouRoumuhiGaku " + " = " + item3_7_2_23_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuRoumuhiKoukyouigaiGaku " + " = " + item3_7_2_24_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            ",GyoumuSonotaChousabuGaku " + " = " + item3_7_2_25_2.Text.Replace("%", "").Replace("¥", "").Replace(",", "") + " " +
                                            " WHERE GyoumuAnkenJouhouID = " + AnkenID + " AND GyoumuHibunKubun = 30 ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();


                        //業務情報技術担当者
                        cmd.CommandText = "DELETE GyoumuJouhouHyouronTantouL1 WHERE GyoumuJouhouID = '" + AnkenID + "' ";
                        cmd.ExecuteNonQuery();

                        for (int i = 1; i < c1FlexGrid5.Rows.Count; i++)
                        {
                            if (c1FlexGrid5.Rows[i][1] != null && c1FlexGrid5.Rows[i][1].ToString() != "")
                            {
                                string Hyouten = "";
                                if (c1FlexGrid5.Rows[i][3] != null && c1FlexGrid5.Rows[i][3].ToString() != "")
                                {
                                    Hyouten = c1FlexGrid5.Rows[i][3].ToString();
                                }
                                cmd.CommandText = "INSERT GyoumuJouhouHyouronTantouL1 ( " +
                                        "GyoumuJouhouID " +
                                        ", HyouronTantouID " +
                                        ", HyouronTantoushaCD " +
                                        ", HyouronTantoushaMei " +
                                        ", HyouronnTantoushaHyouten " +
                                        ") VALUES (" +
                                        "'" + AnkenID + "' " +
                                        "," + i +
                                        ",N'" + c1FlexGrid5.Rows[i][1].ToString() + "' " +
                                        ",N'" + c1FlexGrid5.Rows[i][2].ToString() + "' " +
                                        ",N'" + Hyouten + "' " +
                                        ") ";
                                cmd.ExecuteNonQuery();
                            }
                        }

                        // 名称の取得
                        string GyoumuJouhouMadoShibuMei = "";
                        string GyoumuJouhouMadoKamei = "";
                        DataTable dt2 = new DataTable();
                        cmd.CommandText = "SELECT ShibuMei, KaMei FROM Mst_Busho WHERE GyoumuBushoCD = '" + item3_4_5_Busho.Text + "'";

                        var sda2 = new SqlDataAdapter(cmd);
                        dt2.Clear();
                        sda2.Fill(dt2);
                        if (dt2 != null && dt2.Rows.Count > 0)
                        {
                            GyoumuJouhouMadoShibuMei = dt2.Rows[0][0].ToString();
                            GyoumuJouhouMadoKamei = dt2.Rows[0][1].ToString();
                        }

                        // 窓口担当者が複数いた場合の対応
                        DataTable gyoumuJouhouMadoguchiDT = new DataTable();
                        cmd.CommandText = "SELECT TOP 1 GyoumuJouhouMadoguchiID " +
                                        "FROM GyoumuJouhouMadoguchi " +
                                        "where GyoumuJouhouID = '" + AnkenID + "' " +
                                        "ORDER BY GyoumuJouhouMadoguchiID ";

                        string GyoumuJouhouMadoguchiID = "";

                        var gyoumuJouhouMadoguchiSda = new SqlDataAdapter(cmd);
                        gyoumuJouhouMadoguchiDT.Clear();
                        gyoumuJouhouMadoguchiSda.Fill(gyoumuJouhouMadoguchiDT);
                        if (gyoumuJouhouMadoguchiDT != null && gyoumuJouhouMadoguchiDT.Rows.Count > 0)
                        {
                            GyoumuJouhouMadoguchiID = gyoumuJouhouMadoguchiDT.Rows[0][0].ToString();
                        }

                        //業務情報窓口
                        //cmd.CommandText = "DELETE GyoumuJouhouMadoguchi WHERE GyoumuJouhouID = '" + AnkenID + "' ";
                        //cmd.ExecuteNonQuery();

                        // 窓口担当者が未設定の場合、登録しない。
                        //if (!(item3_4_5_CD.Text == "0") && !(item3_4_5_CD.Text == ""))
                        //{
                        //    int MadoguchiID = GlobalMethod.getSaiban("GyoumuJouhouMadoguchiID");

                        //    cmd.CommandText = "INSERT GyoumuJouhouMadoguchi ( " +
                        //            "GyoumuJouhouID " +
                        //            ", GyoumuJouhouMadoguchiID " +
                        //            ", GyoumuJouhouMadoGyoumuBushoCD " +
                        //            ", GyoumuJouhouMadoShibuMei " +
                        //            ", GyoumuJouhouMadoKamei " +
                        //            ", GyoumuJouhouMadoKojinCD " +
                        //            ", GyoumuJouhouMadoChousainMei " +
                        //            ") VALUES (" +
                        //            "" + AnkenID + " " +
                        //            "," + MadoguchiID +
                        //            ",'" + item3_4_5_Busho.Text + "' " +
                        //            //",'" + item3_4_5_Shibu.Text + "' " +
                        //            //",'" + item3_4_5_Ka.Text + "' " +
                        //            ",'" + GyoumuJouhouMadoShibuMei + "' " +
                        //            ",'" + GyoumuJouhouMadoKamei + "' " +
                        //            ",'" + item3_4_5_CD.Text + "' " +
                        //            ",'" + item3_4_5.Text + "' " +
                        //            ") ";
                        //    Console.WriteLine(cmd.CommandText);
                        //    cmd.ExecuteNonQuery();
                        //}

                        //窓口担当者の更新
                        if ((item3_4_5_CD.Text == "0") || (item3_4_5_CD.Text == ""))
                        {
                            cmd.CommandText = "DELETE GyoumuJouhouMadoguchi WHERE GyoumuJouhouID = '" + AnkenID + "' ";
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            // データが存在する場合
                            if (GyoumuJouhouMadoguchiID != "")
                            {
                                cmd.CommandText = "UPDATE GyoumuJouhouMadoguchi set " +
                                                "GyoumuJouhouMadoGyoumuBushoCD = N'" + item3_4_5_Busho.Text + "' " +
                                                ",GyoumuJouhouMadoShibuMei = N'" + GyoumuJouhouMadoShibuMei + "' " +
                                                ",GyoumuJouhouMadoKamei = N'" + GyoumuJouhouMadoKamei + "' " +
                                                ",GyoumuJouhouMadoKojinCD = N'" + item3_4_5_CD.Text + "' " +
                                                ",GyoumuJouhouMadoChousainMei = N'" + item3_4_5.Text + "' " +
                                                "WHERE GyoumuJouhouMadoguchiID = '" + GyoumuJouhouMadoguchiID + "' ";

                                cmd.ExecuteNonQuery();
                            }
                            // 存在していない場合
                            else
                            {
                                int MadoguchiID = GlobalMethod.getSaiban("GyoumuJouhouMadoguchiID");

                                cmd.CommandText = "INSERT GyoumuJouhouMadoguchi ( " +
                                        "GyoumuJouhouID " +
                                        ", GyoumuJouhouMadoguchiID " +
                                        ", GyoumuJouhouMadoGyoumuBushoCD " +
                                        ", GyoumuJouhouMadoShibuMei " +
                                        ", GyoumuJouhouMadoKamei " +
                                        ", GyoumuJouhouMadoKojinCD " +
                                        ", GyoumuJouhouMadoChousainMei " +
                                        ") VALUES (" +
                                        "" + AnkenID + " " +
                                        "," + MadoguchiID +
                                        ",N'" + item3_4_5_Busho.Text + "' " +
                                        //",'" + item3_4_5_Shibu.Text + "' " +
                                        //",'" + item3_4_5_Ka.Text + "' " +
                                        ",N'" + GyoumuJouhouMadoShibuMei + "' " +
                                        ",N'" + GyoumuJouhouMadoKamei + "' " +
                                        ",N'" + item3_4_5_CD.Text + "' " +
                                        ",N'" + item3_4_5.Text + "' " +
                                        ") ";
                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();
                            }
                        }

                        //排他制御エラー
                        if (!GlobalMethod.Check_Table(AnkenData_N.Rows[0][AnkenData_N.Columns.Count - 1].ToString(), "NyuusatsuUpdateDate", "NyuusatsuJouhou", "NyuusatsuJouhouID = '" + AnkenData_N.Rows[0][AnkenData_N.Columns.Count - 2].ToString() + "' "))
                        {
                            GlobalMethod.outputLogger("UpdateEntry", "入札情報更新 排他制御エラー", "ID:" + AnkenData_N.Rows[0][AnkenData_N.Columns.Count - 2].ToString(), UserInfos[1]);
                            GlobalMethod.GetMessage("E00091", "");
                        }

                        if (!GlobalMethod.Check_Table(AnkenData_N.Rows[0][AnkenData_N.Columns.Count - 2].ToString(), "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("UpdateEntry", "入札情報更新 データなしエラー", "ID:" + AnkenData_N.Rows[0][AnkenData_N.Columns.Count - 2].ToString(), UserInfos[1]);
                            GlobalMethod.GetMessage("E10009", "");
                        }

                        // 入札応札者新規登録フラグ true：初回登録（NyuusatsuRakusatsuShokaiDate を更新）
                        // 既にNyuusatsuRakusatsuShokaiDateが入っていない（NULL）で、NyuusatsuJouhouOusatsushaテーブルにデータを登録した場合、true
                        Boolean nyuusatsuOusatushaInsertFlg = false;
                        // 入札応札者更新フラグ true：更新（NyuusatsuRakusatsuSaisyuDate を更新）
                        // 応札者のGridを回して、NyuusatsuJouhouOusatsushaテーブルと一致しないデータが存在した場合、true
                        Boolean nyuusatsuOusatushaUpdateFlg = false;

                        // 登録日データフラグ false:データ登録なし true：データ登録あり
                        Boolean nyuusatsuOusatushaShokaiDateFlg = false;

                        // 登録日（NyuusatsuRakusatsuShokaiDate）が空でないデータを取得
                        cmd.CommandText = "SELECT  " +
                                            "NyuusatsuRakusatsuShokaiDate " +
                                            "FROM NyuusatsuJouhou " +
                                            "WHERE AnkenJouhouID = '" + AnkenID + "' AND NyuusatsuRakusatsuShokaiDate is not null";
                        var nyuusatsu_sda = new SqlDataAdapter(cmd);
                        var nyuusatsu_dt = new DataTable();
                        nyuusatsu_sda.Fill(nyuusatsu_dt);
                        if (nyuusatsu_dt != null && nyuusatsu_dt.Rows.Count > 0)
                        {
                            nyuusatsuOusatushaShokaiDateFlg = true;
                        }

                        //// 登録されているデータを取得する、IDはDelete、Insertの関係上使えないので除外
                        //cmd.CommandText = "SELECT  " +
                        //                "ISNULL(NyuusatsuRakusatsuJyuni,'') AS Jyuni" +  // 0:落札順位 NULLがあり得るので、空文字に置き換え
                        //                ", NyuusatsuRakusatsuJokyou " +                    // 1:落札状況
                        //                ", NyuusatsuOusatsushaID " +                       // 2:応札者ID
                        //                ", NyuusatsuOusatsusha " +                         // 3:応札者
                        //                ", NyuusatsuOusatsuKingaku " +                     // 4:応札金額
                        //                ", NyuusatsuOusatsuKyougouTashaID " +              // 5:競合他社ID
                        //                ", NyuusatsuRakusatsuComment " +                   // 6:コメント
                        //                ", NyuusatsuOusatsuKyougouKigyouCD " +             // 7:競合企業CD
                        //                "FROM NyuusatsuJouhouOusatsusha " +
                        //                "WHERE NyuusatsuJouhouID = '" + AnkenID + "' ";
                        //nyuusatsu_sda = new SqlDataAdapter(cmd);
                        //nyuusatsu_dt = new DataTable();
                        //nyuusatsu_sda.Fill(nyuusatsu_dt);
                        //if (nyuusatsu_dt != null && nyuusatsu_dt.Rows.Count > 0)
                        //{
                        //    // テーブルデータを回す
                        //    for (int i = 0; i < nyuusatsu_dt.Rows.Count; i++)
                        //    {
                        //        // 応札者Gridを回す
                        //        for (int j = 1; j < c1FlexGrid2.Rows.Count; j++)
                        //        {
                        //            // 4:NyuusatsuOusatsushaID が存在するかどうか
                        //            if (c1FlexGrid2.Rows[j][4] != null && c1FlexGrid2.Rows[j][4].ToString() != "")
                        //            {

                        //                // 落札順位
                        //                if (c1FlexGrid2.Rows[i][2] != null && c1FlexGrid2.Rows[i][2].ToString() != "")
                        //                {

                        //                }
                        //                // 落札状況
                        //                if (c1FlexGrid2.Rows[j][3] != null && c1FlexGrid2.Rows[j][3].ToString() == "True")
                        //                {

                        //                }
                        //                // 応札者ID
                        //                if (c1FlexGrid2.Rows[j][4] != null && c1FlexGrid2.Rows[j][4].ToString() == "")
                        //                {

                        //                }
                        //                // 応札者
                        //                if (c1FlexGrid2.Rows[j][5] != null && c1FlexGrid2.Rows[j][5].ToString() == "")
                        //                {

                        //                }
                        //                // 応札額（税抜）
                        //                if (c1FlexGrid2.Rows[j][6] != null && c1FlexGrid2.Rows[j][6].ToString() == "")
                        //                {

                        //                }
                        //                // コメント
                        //                if (c1FlexGrid2.Rows[j][7] != null && c1FlexGrid2.Rows[j][7].ToString() == "")
                        //                {

                        //                }


                        //            }
                        //        }
                        //    }
                        //}


                        cmd.CommandText = "DELETE NyuusatsuJouhouOusatsusha WHERE NyuusatsuJouhouID = '" + AnkenID + "' ";
                        cmd.ExecuteNonQuery();

                        //入札数の計算のため、入札情報前に入札応札者を登録
                        int nyusatsuCnt = 0;
                        for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
                        {
                            if (c1FlexGrid2.Rows[i][4] != null && c1FlexGrid2.Rows[i][4].ToString() != "")
                            {
                                int rakusatsuFLG = 0;
                                string juni = "null";
                                if (c1FlexGrid2.Rows[i][3] != null && c1FlexGrid2.Rows[i][3].ToString() == "True")
                                {
                                    rakusatsuFLG = 1;
                                }
                                if (c1FlexGrid2.Rows[i][2] != null && c1FlexGrid2.Rows[i][2].ToString() != "")
                                {
                                    juni = c1FlexGrid2.Rows[i][2].ToString();
                                }
                                cmd.CommandText = "INSERT NyuusatsuJouhouOusatsusha ( " +
                                        "NyuusatsuJouhouID " +
                                        ", NyuusatsuOusatsuID " +
                                        ", NyuusatsuRakusatsuJyuni " +//落札順位
                                        ", NyuusatsuRakusatsuJokyou " +//落札状況
                                        ", NyuusatsuOusatsushaID " +
                                        ", NyuusatsuOusatsusha " +
                                        ", NyuusatsuOusatsuKingaku " +
                                        ", NyuusatsuOusatsuKyougouTashaID " +
                                        ", NyuusatsuRakusatsuComment " +
                                        ", NyuusatsuOusatsuKyougouKigyouCD " +
                                        ") VALUES (" +
                                        "'" + AnkenID + "' " +
                                        "," + i +
                                        ", " + juni + " " +
                                        ",'" + rakusatsuFLG + "' " +
                                        ",N'" + c1FlexGrid2.Rows[i][4].ToString() + "' " +
                                        ",N'" + c1FlexGrid2.Rows[i][5].ToString() + "' ";
                                if (c1FlexGrid2.Rows[i][6] != null)
                                {
                                    cmd.CommandText += ",N'" + c1FlexGrid2.Rows[i][6].ToString().Replace("¥", "").Replace(",", "") + "' ";
                                }
                                else
                                {
                                    cmd.CommandText += ",'0' ";
                                }
                                cmd.CommandText += ",N'" + c1FlexGrid2.Rows[i][4].ToString() + "' " +
                                ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid2.Rows[i][7].ToString(), 0) + "' " +
                                ",N'" + c1FlexGrid2.Rows[i][8].ToString() + "' " +
                                ") ";
                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();
                                nyusatsuCnt++;

                                // 登録日がNULLでINSERTが行われたので、入札応札者新規登録フラグを立てる
                                if (nyuusatsuOusatushaShokaiDateFlg == false)
                                {
                                    nyuusatsuOusatushaInsertFlg = true;
                                    nyuusatsuOusatushaUpdateFlg = true;
                                }
                            }
                        }

                        cmd.CommandText = "SELECT  " +
                        //入札タブ
                        //入札参加者
                        "NyuusatsuRakusatsuJyuni " +//落札順位
                        ",NyuusatsuRakusatsuJokyou " +//落札状況
                        ",NyuusatsuOusatsushaID " +
                        ",NyuusatsuOusatsusha " +
                        ",NyuusatsuOusatsuKingaku " +
                        ",NyuusatsuRakusatsuComment " +//コメント
                        ",NyuusatsuOusatsuKyougouKigyouCD " +

                        //参照テーブル
                        "FROM AnkenJouhou " +
                        "LEFT JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                        "LEFT JOIN NyuusatsuJouhouOusatsusha ON NyuusatsuJouhou.NyuusatsuJouhouID =  NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID " +
                        "WHERE AnkenJouhou.AnkenJouhouID =  " + AnkenID.ToString();
                        var nyuusatuSda = new SqlDataAdapter(cmd);
                        AnkenData_Grid2.Clear();
                        nyuusatuSda.Fill(AnkenData_Grid2);

                        StringBuilder sb = new StringBuilder();
                        for (int i = 0; i < AnkenData_Grid2.Rows.Count; i++)
                        {
                            sb.Append(AnkenData_Grid2.Rows[i][0]);
                            sb.Append(",");
                            sb.Append(AnkenData_Grid2.Rows[i][1]);
                            sb.Append(",");
                            sb.Append(AnkenData_Grid2.Rows[i][2]);
                            sb.Append(",");
                            sb.Append(AnkenData_Grid2.Rows[i][3]);
                            sb.Append(",");
                            sb.Append(AnkenData_Grid2.Rows[i][4]);
                            sb.Append(",");
                            sb.Append(AnkenData_Grid2.Rows[i][5]);
                            sb.Append(",");
                            sb.Append(AnkenData_Grid2.Rows[i][6]);
                            sb.Append(",");
                        }
                        string updateC1FlexGrid2Data = sb.ToString();

                        // 入札者情報が同じかどうか
                        if (c1FlexGrid2Data != updateC1FlexGrid2Data)
                        {
                            nyuusatsuOusatushaUpdateFlg = true;
                        }
                        c1FlexGrid2Data = updateC1FlexGrid2Data;

                        // 新規
                        if (nyuusatsuOusatushaInsertFlg == true)
                        {
                            cmd.CommandText = "UPDATE NyuusatsuJouhou SET " +
                                "NyuusatsuRakusatsuShokaiDate = '" + DateTime.Today + "' " +
                                ",NyuusatsuRakusatsuSaisyuDate = '" + DateTime.Today + "' " +
                                "FROM NyuusatsuJouhou " +
                                "WHERE NyuusatsuJouhouID = '" + AnkenID + "' ";
                            cmd.ExecuteNonQuery();
                            // 表示も変更
                            item2_3_3.Text = DateTime.Today.ToString();
                            item2_3_4.Text = DateTime.Today.ToString();
                        }
                        // 更新
                        else if (nyuusatsuOusatushaUpdateFlg == true)
                        {
                            cmd.CommandText = "UPDATE NyuusatsuJouhou SET " +
                                "NyuusatsuRakusatsuSaisyuDate = '" + DateTime.Today + "' " +
                                "FROM NyuusatsuJouhou " +
                                "WHERE NyuusatsuJouhouID = '" + AnkenID + "' ";
                            cmd.ExecuteNonQuery();
                            // 表示も変更
                            item2_3_4.Text = DateTime.Today.ToString();
                        }

                        // No.278 競合他社IDが入っていない対応
                        string KyougouTashaID = "";
                        // 落札者がいれば、競合他社IDを取得する
                        if (item2_3_7.Text != null && item2_3_7.Text != "")
                        {
                            cmd.CommandText = "SELECT  " +
                            "KyougouTashaID " +
                            "FROM Mst_KyougouTasha " +
                            "WHERE KyougouMeishou = N'" + item2_3_7.Text + "' ";
                            var sda = new SqlDataAdapter(cmd);
                            var dt = new DataTable();
                            sda.Fill(dt);
                            KyougouTashaID = dt.Rows[0][0].ToString();
                        }

                        cmd.CommandText = "UPDATE NyuusatsuJouhou SET " +
                                     "NyuusatsuMitsumorigaku = " + "N'" + item1_36.Text.Replace("¥", "").Replace(",", "") + "'" +
                                     ",NyuusatsuRakusatsusha = " + " N'" + GlobalMethod.ChangeSqlText(item2_3_7.Text, 0, 0) + "' " +
                                     ",NyuusatsuRakusatsuKekkaDate = " + " " + Get_DateTimePicker("item2_1_2") + " " +
                                     ",NyuusatsuRakusatugaku = " + "N'" + item2_3_8.Text.Replace("¥", "").Replace(",", "") + "' " +
                                     ",NyuusatsuRakusatuSougaku = " + "N'" + item2_3_8.Text.Replace("¥", "").Replace(",", "") + "' " +
                                     ",NyuusatsuYoteiKakaku = " + "N'" + item2_3_5.Text.Replace("¥", "").Replace(",", "") + "' " +
                                     ",NyuusatsuTanpinMikomigaku = " + "N'" + item3_3_5.Text.Replace("¥", "").Replace(",", "") + "' " +
                                     ",NyuusatsuUpdateProgram = " + "'UpdateEntry' " +
                                     ",NyuusatsuUpdateDate = " + " GETDATE() " +
                                     ",NyuusatsuUpdateUser = " + "N'" + UserInfos[0] + "' " +
                                     ",NyuusatsuHoushiki = " + "N'" + item1_15.SelectedValue + "' " +
                                     ",NyuusatsuGyoumuBikou = " + "N'" + GlobalMethod.ChangeSqlText(item1_18.Text, 0, 0) + "' " +
                                     ",NyuusatsushaSuu = " + "N'" + nyusatsuCnt + "' " +
                                     ",NyuusatsuKekkaMemo = " + "N'" + GlobalMethod.ChangeSqlText(item2_3_12.Text, 0, 0) + "' " +
                                     //",NyuusatsuRakusatsuShokaiDate = " + Get_DateTimePicker("item2_3_3") +
                                     //",NyuusatsuRakusatsuSaisyuDate = " + Get_DateTimePicker("item2_3_4") +
                                     ",NyuusatsuRakusatsushaID = " + "N'" + item2_1_1.SelectedValue.ToString() + "' " +
                                     ",NyuusatsuRakusatsuShaJokyou = " + "" + item2_3_1.SelectedValue.ToString() + " " +
                                     ",NyuusatsuRakusatsuGakuJokyou = " + "" + item2_3_2.SelectedValue.ToString() + " " +
                                     ",NyuusatsuKyougouTasha = " + " N'" + GlobalMethod.ChangeSqlText(item2_3_7.Text, 0, 0) + "' ";

                        // 競合他社ID
                        if (KyougouTashaID != "")
                        {
                            cmd.CommandText += ",NyuusatsuKyougouTashaID = " + " N'" + KyougouTashaID + "' ";
                        }
                        else
                        {
                            // 空の場合はnull
                            cmd.CommandText += ",NyuusatsuKyougouTashaID = null ";
                        }

                        cmd.CommandText += " WHERE NyuusatsuJouhouID = " + AnkenID;
                        cmd.ExecuteNonQuery();

                        if (mode == 3)
                        {
                            cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                        "AnkenKianZumi = " + 1 +
                                        " WHERE AnkenJouhouID = " + AnkenID;
                            Console.WriteLine(cmd.CommandText);
                            cmd.ExecuteNonQuery();
                        }

                        transaction.Commit();

                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        conn.Close();
                        throw;
                    }

                    //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報 ID:" + AnkenID, "UpdateEntory", "");
                    //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報(エントリー)を更新しました ID:" + AnkenID, "UpdateEntory", "");
                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報(エントリー)を更新しました ID:" + AnkenID, pgmName + methodName, "");


                    if (mode == 1)
                    {
                        //更新時も窓口情報にデータを連携
                        try
                        {
                            cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                                        //"MadoguchiJutakuBangou = AnkenJutakuBangou " +
                                        "MadoguchiJutakuBangou = replace(AnkenJutakuBangou,'-' + AnkenJutakuBangouEda,'') " +
                                        ",MadoguchiJutakuBangouEdaban = AnkenJutakuBangouEda " +
                                        ",MadoguchiJutakuBushoCD = AnkenJutakubushoCD " +
                                        ",MadoguchiJutakubushoMeiOld = ShibuMei " +
                                        ",MadoguchiJutakuTantoushaID = AnkenTantoushaCD " +
                                        ",MadoguchiJutakuTantoushaOld = ChousainMei " +
                                        ",MadoguchiKanriGijutsusha = KanriGijutsushaCD " +
                                        ",MadoguchiGyoumuKanrishaCD = GyoumuKanrishaCD " +
                                        //",MadoguchiTantoushaBushoCD = GyoumuJouhouMadoGyoumuBushoCD " + //窓口部所
                                        //",MadoguchiBushoShozokuCD = GyoumuJouhouMadoGyoumuBushoCD " + //窓口部所所属
                                        // 1213 窓口ミハルの窓口担当者は、初期登録のみエントリくんとリンクさせる。
                                        //",MadoguchiTantoushaCD = GyoumuJouhouMadoKojinCD " + //窓口担当者
                                        " FROM AnkenJouhou " +
                                        " LEFT JOIN Mst_Busho ON GyoumuBushoCD = AnkenJutakubushoCD " +
                                        " LEFT JOIN Mst_Chousain ON KojinCD = AnkenTantoushaCD " +
                                        " LEFT JOIN GyoumuJouhou ON GyoumuJouhou.AnkenJouhouID = AnkenJouhou.AnkenJouhouID " +
                                        " LEFT JOIN GyoumuJouhouMadoguchi ON GyoumuJouhouMadoguchi.GyoumuJouhouID = AnkenJouhou.AnkenJouhouID " +
                                        //" WHERE MadoguchiJouhou.AnkenJouhouID = AnkenJouhou.AnkenJouhouID AND MadoguchiJouhou.AnkenJouhouID = " + AnkenID;
                                        " WHERE MadoguchiJouhou.AnkenJouhouID = AnkenJouhou.AnkenJouhouID AND MadoguchiJouhou.AnkenJouhouID = " + AnkenID + " " +
                                        "AND GyoumuJouhouMadoGyoumuBushoCD is not NULL "; // 既に窓口に連携済の場合、NULLが不可なので、更新をスルーさせる
                            cmd.ExecuteNonQuery();

                            // Garoon宛先追加
                            insertGaroonAtesakiTsuika(cmd);
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                        set_error(GlobalMethod.GetMessage("I00008", ""));
                        conn.Close();

                        // えんとり君修正STEP2
                        if (item1_7.Text == "")
                        {
                            label51.Visible = true;
                            txt_renamedfolder.Visible = true;
                            label115.Visible = true;
                        }
                        else
                        {
                            label51.Visible = false;
                            txt_renamedfolder.Visible = false;
                            label115.Visible = false;
                        }

                        return true;
                    }
                    if (mode == 2 || mode == 3)
                    {
                        GlobalMethod.outputLogger("KianEntry", AnkenID + ":" + Header2.Text + ":" + item1_10.SelectedValue.ToString() + ":" + item1_11_CD.Text, "", UserInfos[1]);
                        try
                        {
                            cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                                        //"MadoguchiJutakuBangou = AnkenJutakuBangou " +
                                        "MadoguchiJutakuBangou = replace(AnkenJutakuBangou,'-' + AnkenJutakuBangouEda,'') " +
                                        ",MadoguchiJutakuBangouEdaban = AnkenJutakuBangouEda " +
                                        ",MadoguchiJutakuBushoCD = AnkenJutakubushoCD " +
                                        ",MadoguchiJutakubushoMeiOld = ShibuMei " +
                                        ",MadoguchiJutakuTantoushaID = AnkenTantoushaCD " +
                                        ",MadoguchiJutakuTantoushaOld = ChousainMei " +
                                        ",MadoguchiKanriGijutsusha = KanriGijutsushaCD " +
                                        " FROM AnkenJouhou " +
                                        " LEFT JOIN Mst_Busho ON GyoumuBushoCD = AnkenJutakubushoCD " +
                                        " LEFT JOIN Mst_Chousain ON KojinCD = AnkenTantoushaCD " +
                                        " LEFT JOIN GyoumuJouhou ON GyoumuJouhou.AnkenJouhouID = AnkenJouhou.AnkenJouhouID " +
                                        " WHERE MadoguchiJouhou.AnkenJouhouID = AnkenJouhou.AnkenJouhouID AND MadoguchiJouhou.AnkenJouhouID = " + AnkenID;
                            cmd.ExecuteNonQuery();

                            //// チェック用帳票の出力時はPrintHistoryがあるので出力しない
                            //if (mode == 3)
                            //{
                            //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "起案しました ID:" + AnkenID, "UpdateEntory", "");
                            GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "起案しました ID:" + AnkenID, pgmName + methodName, "");
                            //}
                            // Garoon宛先追加
                            insertGaroonAtesakiTsuika(cmd);
                        }
                        catch (Exception)
                        {
                            conn.Close();
                            throw;
                        }

                    }
                    conn.Close();
                    // 計画番号を変更後のものを保持
                    beforeKeikakuBangou = item1_4.Text;
                }

            }
            // 変更伝票の起案
            else if (mode == 4)
            {
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {
                        cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                "AnkenUpdateDate = GETDATE() " +
                                ",AnkenUpdateUser = '" + UserInfos[0] + "' " +
                                ",AnkenSaishinFlg = 0 " +
                                ",AnkenUpdateProgram = 'ChangeKianEntry' " +
                                " WHERE AnkenJouhou.AnkenJouhouID = " + AnkenID;
                        var result = cmd.ExecuteNonQuery();

                        if (result == 0)
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }

                        string SakuseiKubun = item3_1_1.SelectedValue.ToString();
                        // 赤伝のAnkenJouhouID
                        int ankenNo = GlobalMethod.getSaiban("AnkenJouhouID");
                        item3_1_20_akaden.Text = ankenNo.ToString();

                        cmd.CommandText = "INSERT INTO AnkenJouhou ( " +
                                "AnkenJouhouID " +
                                ",AnkenSakuseiKubun " +
                                ",AnkenSaishinFlg " +
                                ",AnkenKishuKeikakugaku " +
                                ",AnkenKishuKeikakakugakuJf " +
                                ",AnkenKishuKeikakugakuJ " +
                                ",AnkenKeikakuZangaku " +
                                ",AnkenkeikakuZangakuJF " +
                                ",AnkenkeikakuZangakuJ " +
                                ",AnkenChokusetsuGenka " +
                                ",AnkenChokusetsuGenkaRitsu " +
                                ",AnkenGaichuuhi " +
                                ",AnkenJoukanDoboku " +
                                ",AnkenJoukanFukugou " +
                                ",AnkenJoukanGesuidou " +
                                ",AnkenJoukanHyoujun " +
                                ",AnkenJoukanIchiba " +
                                ",AnkenJoukanItiji " +
                                ",AnkenJoukanJutakuSonota " +
                                ",AnkenJoukanKentiku " +
                                ",AnkenJoukanKijunsho " +
                                ",AnkenJoukanKouwan " +
                                ",AnkenJoukanKuukou " +
                                ",AnkenJoukanSetsubi " +
                                ",AnkenJoukanSonota " +
                                ",AnkenJoukanSuidou " +
                                ",AnkenKeichoukaiKounyuuhi " +
                                ",AnkenKishuKeikakugakuK " +
                                ",AnkenKaisuu " +
                                ",AnkenCreateDate " +
                                ",AnkenCreateUser " +
                                ",AnkenCreateProgram " +
                                ",AnkenUpdateDate " +
                                ",AnkenUpdateUser " +
                                ",AnkenUpdateProgram " +
                                //",AnkenUriagebi " +
                                ",AnkenTourokubi " +
                                ",AnkenGyoumuMei " +
                                ",AnkenDeleteFlag " +

                                ",AnkenUriageNendo " +
                                ",AnkenHachushaKubunCD " +
                                ",AnkenHachushaKubunMei " +
                                ",AnkenHachuushaCodeID " +
                                ",AnkenHachuushaMei " +
                                ",AnkenGyoumuKubun " +
                                ",AnkenGyoumuKubunMei " +
                                ",AnkenNyuusatsuHoushiki " +
                                ",AnkenKyougouTasha " +
                                ",AnkenJutakubushoCD " +
                                ",AnkenJutakushibu " +
                                ",AnkenTantoushaCD " +
                                ",AnkenMadoguchiTantoushaCD " +
                                ",AnkenGyoumuKanrishaCD " +
                                ",AnkenGyoumuKanrisha " +
                                ",GyoumuKanrishaCD " +
                                ",AnkenHachuushaBusho " +
                                ",AnkenkeikakuZangakuK " +
                                ",AnkenJutakuBangou " +
                                ",AnkenJutakuBangouEda " +
                                ",AnkenNyuusatsuYoteibi " +
                                ",AnkenRakusatsusha " +
                                ",AnkenRakusatsuJouhou " +
                                ",AnkenKianZumi " +
                                ",AnkenKiangetsu " +
                                ",AnkenHanteiKubun " +
                                ",AnkenJoukanData " +
                                ",AnkenJoukanHachuuKikanCD " +
                                ",AnkenNyuukinKakuninbi " +
                                ",AnkenKanryouSakuseibi " +
                                ",AnkenHonbuKakuninbi " +
                                ",AnkenShizaiChousa " +
                                ",AnkenKoujiChousahi " +
                                ",AnkenKikiruiChousa " +
                                ",AnkenSanpaiFukusanbutsu " +
                                ",AnkenHokakeChousa " +
                                ",AnkenShokeihiChousa " +
                                ",AnkenGenkaBunseki " +
                                ",AnkenKijunsakusei " +
                                ",AnkenKoukyouRoumuhi " +
                                ",AnkenRoumuhiKoukyouigai " +
                                ",AnkenSonotaChousabu " +
                                ",AnkenOrdermadeJifubu " +
                                ",AnkenRIBCJifubu " +
                                ",AnkenSonotaJifubu " +
                                ",AnkenOrdermade " +
                                ",AnkenJouhouKaihatsu " +
                                ",AnkenRIBCJouhouKaihatsu " +
                                ",AnkenSoukenbu " +
                                ",AnkenSonotaJoujibu " +
                                ",AnkenTeikiTokuchou " +
                                ",AnkenTanpinTokuchou " +
                                ",AnkenKikiChousa " +
                                ",AnkenHachuushaIraibusho " +
                                ",AnkenHachuushaTantousha " +
                                ",AnkenHachuushaTEL " +
                                ",AnkenHachuushaFAX " +
                                ",AnkenHachuushaMail " +
                                ",AnkenHachuushaIraiYuubin " +
                                ",AnkenHachuushaIraiJuusho " +
                                ",AnkenHachuushaKeiyakuBusho " +
                                ",AnkenHachuushaKeiyakuTantou " +
                                ",AnkenHachuushaKeiyakuTEL " +
                                ",AnkenHachuushaKeiyakuFAX " +
                                ",AnkenHachuushaKeiyakuMail " +
                                ",AnkenHachuushaKeiyakuYuubin " +
                                ",AnkenHachuushaKeiyakuJuusho " +
                                ",AnkenHachuuDaihyouYakushoku " +
                                ",AnkenHachuuDaihyousha " +
                                ",AnkenRosenKawamei " +
                                ",AnkenGyoumuItakuKasho " +
                                ",AnkenJititaiKibunID " +
                                ",AnkenJititaiKubun " +
                                ",AnkenKeiyakuToshoNo " +
                                ",AnkenKirokuToshoNo " +
                                ",AnkenKirokuHokanNo " +
                                ",AnkenCDHokan " +
                                ",AnkenSeikaButsuHokanFile " +
                                ",AnkenSeikabutsuHokanbako " +
                                ",AnkenKokyakuHyoukaComment " +
                                ",AnkenToukaiHyoukaComment " +
                                ",AnkenKenCD " +
                                ",AnkenToshiCD " +
                                ",AnkenKeiyakusho " +
                                ",AnkenEizen " +
                                ",AnkenTantoushaMei " +
                                ",GyoumuKanrishaMei " +
                                ",AnkenGyoumuKubunCD " +
                                ",AnkenHachuushaKaMei " +
                                ",AnkenKeiyakuKoukiKaishibi " +
                                ",AnkenKeiyakuKoukiKanryoubi " +
                                ",AnkenKeiyakuTeiketsubi " +
                                ",AnkenKeiyakuZeikomiKingaku " +
                                ",AnkenKeiyakuUriageHaibunGakuC " +
                                ",AnkenKeiyakuUriageHaibunGakuJ " +
                                ",AnkenKeiyakuUriageHaibunGakuJs " +
                                ",AnkenKeiyakuUriageHaibunGakuK " +
                                ",AnkenKeiyakuUriageHaibunGakuR " +
                                ",AnkenKeiyakuSakuseibi " +
                                ",AnkenAnkenBangou " +
                                ",AnkenKeikakuBangou " +
                                ",AnkenHikiaijhokyo " +
                                ",AnkenKeikakuAnkenMei " +
                                ",AnkenToukaiSankouMitsumori " +
                                ",AnkenToukaiJyutyuIyoku " +
                                ",AnkenToukaiSankouMitsumoriGaku " +
                                ",AnkenHachushaKaMei " +
                                ",AnkenHachushaCD " +
                                ",AnkenToukaiOusatu " +
                                ",AnkenKoukiNendo " +
                                " ) SELECT " +
                                ankenNo;

                        if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                        {
                            cmd.CommandText += ",'02' " +
                                    ",0 " +
                                    ",- AnkenKishuKeikakugaku " +
                                    ",- AnkenKishuKeikakakugakuJf " +
                                    ",- AnkenKishuKeikakugakuJ ";
                        }
                        else
                        {
                            cmd.CommandText += ",'04' " +
                                    ",1 " +
                                    ",0 " +
                                    ",0 " +
                                    ",0 ";

                        }

                        cmd.CommandText += ",- AnkenKeikakuZangaku " +
                                ",- AnkenkeikakuZangakuJF " +
                                ",- AnkenkeikakuZangakuJ " +
                                ",- AnkenChokusetsuGenka " +
                                ",- AnkenChokusetsuGenkaRitsu " +
                                ",- AnkenGaichuuhi " +
                                ",- AnkenJoukanDoboku " +
                                ",- AnkenJoukanFukugou " +
                                ",- AnkenJoukanGesuidou " +
                                ",- AnkenJoukanHyoujun " +
                                ",- AnkenJoukanIchiba " +
                                ",- AnkenJoukanItiji " +
                                ",- AnkenJoukanJutakuSonota " +
                                ",- AnkenJoukanKentiku " +
                                ",- AnkenJoukanKijunsho " +
                                ",- AnkenJoukanKouwan " +
                                ",- AnkenJoukanKuukou " +
                                ",- AnkenJoukanSetsubi " +
                                ",- AnkenJoukanSonota " +
                                ",- AnkenJoukanSuidou " +
                                ",- AnkenKeichoukaiKounyuuhi " +
                                ",- AnkenKishuKeikakugakuK " +
                                ",AnkenKaisuu + 1 " +
                                ",GETDATE() " +
                                ",'" + UserInfos[0] + "' " +
                                ",'ChangeKianEntry' " +
                                ",GETDATE() " +
                                ",'" + UserInfos[0] + "' " +
                                ",'ChangeKianEntry' " +
                                //",null " +                  //AnkenUriagebi
                                ",AnkenTourokubi " +
                                ",AnkenGyoumuMei " +
                                ",0 " +

                                ",AnkenUriageNendo " +
                                ",AnkenHachushaKubunCD " +
                                ",AnkenHachushaKubunMei " +
                                ",AnkenHachuushaCodeID " +
                                ",AnkenHachuushaMei " +
                                ",AnkenGyoumuKubun " +
                                ",AnkenGyoumuKubunMei " +
                                ",AnkenNyuusatsuHoushiki " +
                                ",AnkenKyougouTasha " +
                                ",AnkenJutakubushoCD " +
                                ",AnkenJutakushibu " +
                                ",AnkenTantoushaCD " +
                                ",AnkenMadoguchiTantoushaCD " +
                                ",AnkenGyoumuKanrishaCD " +
                                ",AnkenGyoumuKanrisha " +
                                ",GyoumuKanrishaCD " +
                                ",AnkenHachuushaBusho " +
                                ",AnkenkeikakuZangakuK " +
                                ",AnkenJutakuBangou " +
                                ",AnkenJutakuBangouEda " +
                                ",AnkenNyuusatsuYoteibi " +
                                ",AnkenRakusatsusha " +
                                ",AnkenRakusatsuJouhou " +
                                ",AnkenKianZumi " +
                                ",AnkenKiangetsu " +
                                ",AnkenHanteiKubun " +
                                ",AnkenJoukanData " +
                                ",AnkenJoukanHachuuKikanCD " +
                                ",AnkenNyuukinKakuninbi " +
                                ",AnkenKanryouSakuseibi " +
                                ",AnkenHonbuKakuninbi " +
                                ",AnkenShizaiChousa " +
                                ",AnkenKoujiChousahi " +
                                ",AnkenKikiruiChousa " +
                                ",AnkenSanpaiFukusanbutsu " +
                                ",AnkenHokakeChousa " +
                                ",AnkenShokeihiChousa " +
                                ",AnkenGenkaBunseki " +
                                ",AnkenKijunsakusei " +
                                ",AnkenKoukyouRoumuhi " +
                                ",AnkenRoumuhiKoukyouigai " +
                                ",AnkenSonotaChousabu " +
                                ",AnkenOrdermadeJifubu " +
                                ",AnkenRIBCJifubu " +
                                ",AnkenSonotaJifubu " +
                                ",AnkenOrdermade " +
                                ",AnkenJouhouKaihatsu " +
                                ",AnkenRIBCJouhouKaihatsu " +
                                ",AnkenSoukenbu " +
                                ",AnkenSonotaJoujibu " +
                                ",AnkenTeikiTokuchou " +
                                ",AnkenTanpinTokuchou " +
                                ",AnkenKikiChousa " +
                                ",AnkenHachuushaIraibusho " +
                                ",AnkenHachuushaTantousha " +
                                ",AnkenHachuushaTEL " +
                                ",AnkenHachuushaFAX " +
                                ",AnkenHachuushaMail " +
                                ",AnkenHachuushaIraiYuubin " +
                                ",AnkenHachuushaIraiJuusho " +
                                ",AnkenHachuushaKeiyakuBusho " +
                                ",AnkenHachuushaKeiyakuTantou " +
                                ",AnkenHachuushaKeiyakuTEL " +
                                ",AnkenHachuushaKeiyakuFAX " +
                                ",AnkenHachuushaKeiyakuMail " +
                                ",AnkenHachuushaKeiyakuYuubin " +
                                ",AnkenHachuushaKeiyakuJuusho " +
                                ",AnkenHachuuDaihyouYakushoku " +
                                ",AnkenHachuuDaihyousha " +
                                ",AnkenRosenKawamei " +
                                ",AnkenGyoumuItakuKasho " +
                                ",AnkenJititaiKibunID " +
                                ",AnkenJititaiKubun " +
                                ",AnkenKeiyakuToshoNo " +
                                ",AnkenKirokuToshoNo " +
                                ",AnkenKirokuHokanNo " +
                                ",AnkenCDHokan " +
                                ",AnkenSeikaButsuHokanFile " +
                                ",AnkenSeikabutsuHokanbako " +
                                ",AnkenKokyakuHyoukaComment " +
                                ",AnkenToukaiHyoukaComment " +
                                ",AnkenKenCD " +
                                ",AnkenToshiCD " +
                                ",AnkenKeiyakusho " +
                                ",AnkenEizen " +
                                ",AnkenTantoushaMei " +
                                ",GyoumuKanrishaMei " +
                                ",AnkenGyoumuKubunCD " +
                                ",AnkenHachuushaKaMei " +
                                ",AnkenKeiyakuKoukiKaishibi " +
                                ",AnkenKeiyakuKoukiKanryoubi " +
                                ",AnkenKeiyakuTeiketsubi " +
                                //",AnkenKeiyakuZeikomiKingaku " +
                                //",AnkenKeiyakuUriageHaibunGakuC " +
                                //",AnkenKeiyakuUriageHaibunGakuJ " +
                                //",AnkenKeiyakuUriageHaibunGakuJs " +
                                //",AnkenKeiyakuUriageHaibunGakuK " +
                                //",AnkenKeiyakuUriageHaibunGakuR " +
                                ",AnkenKeiyakuZeikomiKingaku " +     // 契約タブの契約金額の税込
                                ",AnkenKeiyakuUriageHaibunGakuC " +  // 契約タブの受託金額配分の調査部、配分額（税込）
                                ",AnkenKeiyakuUriageHaibunGakuJ " +  // 契約タブの受託金額配分の事業普及部、配分額（税込）
                                ",AnkenKeiyakuUriageHaibunGakuJs " + // 契約タブの受託金額配分の情報システム部、配分額（税込）
                                ",AnkenKeiyakuUriageHaibunGakuK " +  // 契約タブの受託金額配分の総合研究所、配分額（税込）
                                ",AnkenKeiyakuUriageHaibunGakuR " +  // なし
                                ",AnkenKeiyakuSakuseibi " +
                                ",AnkenAnkenBangou " +
                                ",AnkenKeikakuBangou " +
                                ",AnkenHikiaijhokyo " +
                                ",AnkenKeikakuAnkenMei " +
                                ",AnkenToukaiSankouMitsumori " +
                                ",AnkenToukaiJyutyuIyoku " +
                                ",AnkenToukaiSankouMitsumoriGaku " +
                                ",AnkenHachushaKaMei " +
                                ",AnkenHachushaCD " +
                                ",AnkenToukaiOusatu " +
                                ",AnkenKoukiNendo " +
                                " FROM AnkenJouhou WHERE AnkenJouhou.AnkenJouhouID = " + AnkenID;
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();


                        if (!GlobalMethod.Check_Table(AnkenID, "KokyakuKeiyakuID", "KokyakuKeiyakuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }

                        cmd.CommandText = "INSERT INTO AnkenJouhouZenkaiRakusatsu ( " +
                                "AnkenJouhouID " +
                                ",AnkenZenkaiJutakuKingaku " +
                                ",AnkenZenkaiRakusatsuID " +

                                ",AnkenZenkaiJutakuBangou " +
                                ",AnkenZenkaiJutakuEdaban " +
                                ",AnkenZenkaiAnkenBangou " +
                                ",AnkenZenkaiRakusatsushaID " +
                                ",AnkenZenkaiRakusatsusha " +
                                ",AnkenZenkaiGyoumuMei " +
                                ",AnkenZenkaiKyougouKigyouCD " +
                                ",AnkenZenkaiJutakuZeinuki " +
                                ",KeiyakuZenkaiRakusatsushaID " +
                                " ) SELECT " +
                                ankenNo +
                                ",-AnkenZenkaiJutakuKingaku " +
                                ",AnkenZenkaiRakusatsuID " +
                                ",AnkenZenkaiJutakuBangou " +
                                ",AnkenZenkaiJutakuEdaban " +
                                ",AnkenZenkaiAnkenBangou " +
                                ",AnkenZenkaiRakusatsushaID " +
                                ",AnkenZenkaiRakusatsusha " +
                                ",AnkenZenkaiGyoumuMei " +
                                ",AnkenZenkaiKyougouKigyouCD " +
                                ",AnkenZenkaiJutakuZeinuki " +
                                ",KeiyakuZenkaiRakusatsushaID " +
                                " FROM AnkenJouhouZenkaiRakusatsu WHERE AnkenJouhouZenkaiRakusatsu.AnkenJouhouID = " + AnkenID;
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        if (!GlobalMethod.Check_Table(AnkenID, "KokyakuKeiyakuID", "KokyakuKeiyakuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "顧客契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }

                        cmd.CommandText = "INSERT INTO KokyakuKeiyakuJouhou ( " +
                                "AnkenJouhouID " +
                                ",KokyakuKeiyakuID " +
                                ",KokyakuCreateUser   " +
                                ",KokyakuCreateDate   " +
                                ",KokyakuCreateProgram" +
                                ",KokyakuUpdateUser   " +
                                ",KokyakuUpdateDate   " +
                                ",KokyakuDeleteFlag   " +
                                ",KokyakuUpdateProgram" +

                                ",KokyakuKeiyakuTanka " +
                                ",KokyakuKeiyakuChosakuken " +
                                ",KokyakuKeiyakuKeisai " +
                                ",KokyakuKeiyakuTokchouChosaku " +
                                ",KokyakuKeiyakuRiyuu " +
                                ",KokyakuMaebaraiJoukou " +
                                ",KokyakuMaebaraiSeikyuu " +
                                ",KokyakuSekkeiTanka " +
                                ",KokyakuSekisanKijun " +
                                ",KokyakuKaiteiGetsu " +
                                ",KokyakuShichouson " +
                                ",KokyakuGijutsuCenter " +
                                ",KokyakuSonota " +
                                ",KokyakuKeiyakuRiyuuTou " +
                                ",KokyakuDataTeikyou " +
                                ",KokyakuAlpha " +
                                ",KokyakuDataDoboku " +
                                ",KokyakuDataNourin " +
                                ",KokyakuDataEizen " +
                                ",KokyakuDataSonota " +
                                ",KokyakuDataSekouP " +
                                ",KokyakuDataDobokuKouji " +
                                ",KokyakuDataRIBC " +
                                ",KokyakuDataGoukei " +
                                ",KokyakuDataKeisaiTanka " +
                                ",KokyakuDataWebTeikyou " +
                                ",KokyakuDataKeiyaku " +
                                ",KokyakuDataTempFile " +
                                ",KokyakuDataTempFileData " +
                                ",KokyakuData05Comment " +
                                ",KokyakuData06Comment " +
                                ",KokyakuData07Comment " +
                                ",KokyakuDataMeiki " +
                                ",KokyakuDataTeikyouTensu " +
                                " ) SELECT " +
                                ankenNo +
                                "," + ankenNo +
                                ",'" + UserInfos[0] + "' " +
                                ",GETDATE() " +
                                ",'ChangeKianEntry' " +
                                ",'" + UserInfos[0] + "' " +
                                ",GETDATE() " +
                                ",0" +
                                ",'ChangeKianEntry' " +

                                ",KokyakuKeiyakuTanka " +
                                ",KokyakuKeiyakuChosakuken " +
                                ",KokyakuKeiyakuKeisai " +
                                ",KokyakuKeiyakuTokchouChosaku " +
                                ",KokyakuKeiyakuRiyuu " +
                                ",KokyakuMaebaraiJoukou " +
                                ",KokyakuMaebaraiSeikyuu " +
                                ",KokyakuSekkeiTanka " +
                                ",KokyakuSekisanKijun " +
                                ",KokyakuKaiteiGetsu " +
                                ",KokyakuShichouson " +
                                ",KokyakuGijutsuCenter " +
                                ",KokyakuSonota " +
                                ",KokyakuKeiyakuRiyuuTou " +
                                ",KokyakuDataTeikyou " +
                                ",KokyakuAlpha " +
                                ",KokyakuDataDoboku " +
                                ",KokyakuDataNourin " +
                                ",KokyakuDataEizen " +
                                ",KokyakuDataSonota " +
                                ",KokyakuDataSekouP " +
                                ",KokyakuDataDobokuKouji " +
                                ",KokyakuDataRIBC " +
                                ",KokyakuDataGoukei " +
                                ",KokyakuDataKeisaiTanka " +
                                ",KokyakuDataWebTeikyou " +
                                ",KokyakuDataKeiyaku " +
                                ",KokyakuDataTempFile " +
                                ",KokyakuDataTempFileData " +
                                ",KokyakuData05Comment " +
                                ",KokyakuData06Comment " +
                                ",KokyakuData07Comment " +
                                ",KokyakuDataMeiki " +
                                ",KokyakuDataTeikyouTensu " +
                                " FROM KokyakuKeiyakuJouhou WHERE KokyakuKeiyakuJouhou.AnkenJouhouID = " + AnkenID;
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        if (!GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "業務情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }

                        cmd.CommandText = "INSERT INTO GyoumuJouhou ( " +
                                "AnkenJouhouID " +
                                ",GyoumuJouhouID " +
                                ",GyoumuCreateDate " +
                                ",GyoumuCreateUser    " +
                                ",GyoumuCreateProgram " +
                                ",GyoumuUpdateDate    " +
                                ",GyoumuUpdateUser    " +
                                ",GyoumuUpdateProgram " +
                                ",GyoumuDeleteFlag " +

                                ",GyoumuHyouten " +
                                ",KanriGijutsushaCD " +
                                ",GyoumuKanriHyouten " +
                                ",ShousaTantoushaCD " +
                                ",SinsaTantoushaCD " +
                                ",GyoumuTECRISTourokuBangou " +
                                ",GyoumuKeisaiTankaTeikyou " +
                                ",GyoumuChosakukenJouto " +
                                ",GyoumuSeikyuubi " +
                                ",GyoumuSeikyuusho " +
                                ",GyoumuHikiwatashiNaiyou " +
                                ",KanriGijutsushaNM " +
                                ",ShousaTantoushaNM " +
                                ",SinsaTantoushaNM " +
                                ",GyoumuShousaHyouten " +
                                " ) SELECT " +
                                ankenNo +
                                "," + ankenNo +
                                ",GETDATE() " +
                                ",'" + UserInfos[0] + "' " +
                                ",'ChangeKianEntry' " +
                                ",GETDATE() " +
                                ",'" + UserInfos[0] + "' " +
                                ",'ChangeKianEntry' " +
                                ",0 " +

                                ",GyoumuHyouten " +
                                ",KanriGijutsushaCD " +
                                ",GyoumuKanriHyouten " +
                                ",ShousaTantoushaCD " +
                                ",SinsaTantoushaCD " +
                                ",GyoumuTECRISTourokuBangou " +
                                ",GyoumuKeisaiTankaTeikyou " +
                                ",GyoumuChosakukenJouto " +
                                ",GyoumuSeikyuubi " +
                                ",GyoumuSeikyuusho " +
                                ",GyoumuHikiwatashiNaiyou " +
                                ",KanriGijutsushaNM " +
                                ",ShousaTantoushaNM " +
                                ",SinsaTantoushaNM " +
                                ",GyoumuShousaHyouten " +
                                " FROM GyoumuJouhou WHERE GyoumuJouhou.AnkenJouhouID = " + AnkenID;
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyouronTantouL1", ""))
                        {
                            cmd.CommandText = "INSERT INTO GyoumuJouhouHyouronTantouL1 ( " +
                                    "GyoumuJouhouID " +
                                    ",HyouronTantouID " +

                                    ",HyouronTantoushaCD " +
                                    ",HyouronTantoushaMei " +
                                    ",HyouronnTantoushaHyouten " +
                                   " ) SELECT " +
                                    ankenNo +
                                    ",HyouronTantouID " +

                                    ",HyouronTantoushaCD " +
                                    ",HyouronTantoushaMei " +
                                    ",HyouronnTantoushaHyouten " +
                                    " FROM GyoumuJouhouHyouronTantouL1 WHERE GyoumuJouhouHyouronTantouL1.GyoumuJouhouID = " + AnkenID;
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouMadoguchi", ""))
                        {
                            //cmd.CommandText = "INSERT INTO GyoumuJouhouMadoguchi ( " +
                            //        "GyoumuJouhouID " +
                            //        ",GyoumuJouhouMadoguchiID " +
                            //        ",GyoumuJouhouMadoKojinCD " +
                            //        ",GyoumuJouhouMadoChousainMei " +

                            //        ",GyoumuJouhouMadoGyoumuBushoCD " +
                            //        ",GyoumuJouhouMadoShibuMei " +
                            //        ",GyoumuJouhouMadoKamei " +
                            //        " ) SELECT " +
                            //        ankenNo +
                            //        "," + GlobalMethod.getSaiban("GyoumuJouhouMadoguchiID") + " " +
                            //        ",GyoumuJouhouMadoKojinCD " +
                            //        ",GyoumuJouhouMadoChousainMei " +

                            //        ",GyoumuJouhouMadoGyoumuBushoCD " +
                            //        ",GyoumuJouhouMadoShibuMei " +
                            //        ",GyoumuJouhouMadoKamei " +
                            //        " FROM GyoumuJouhouMadoguchi WHERE GyoumuJouhouMadoguchi.GyoumuJouhouID = " + AnkenID;
                            //Console.WriteLine(cmd.CommandText);
                            //result = cmd.ExecuteNonQuery();

                            // 窓口担当者
                            // 新では1件しか入らないが、現行が複数件はいるので、複数件あった場合でも落ちないようにする
                            DataTable gyoumuJouhouDT = new DataTable();
                            cmd.CommandText = "SELECT " +
                                    " GyoumuJouhouMadoKojinCD " +
                                    ",GyoumuJouhouMadoChousainMei " +
                                    ",GyoumuJouhouMadoGyoumuBushoCD " +
                                    ",GyoumuJouhouMadoShibuMei " +
                                    ",GyoumuJouhouMadoKamei " +
                                    " FROM GyoumuJouhouMadoguchi WHERE GyoumuJouhouMadoguchi.GyoumuJouhouID = " + AnkenID;
                            var sda = new SqlDataAdapter(cmd);
                            sda.Fill(gyoumuJouhouDT);

                            if (gyoumuJouhouDT != null && gyoumuJouhouDT.Rows.Count > 0)
                            {
                                for (int i = 0; i < gyoumuJouhouDT.Rows.Count; i++)
                                {
                                    cmd.CommandText = "INSERT INTO GyoumuJouhouMadoguchi ( " +
                                            "GyoumuJouhouID " +
                                            ",GyoumuJouhouMadoguchiID " +
                                            ",GyoumuJouhouMadoKojinCD " +
                                            ",GyoumuJouhouMadoChousainMei " +
                                            ",GyoumuJouhouMadoGyoumuBushoCD " +
                                            ",GyoumuJouhouMadoShibuMei " +
                                            ",GyoumuJouhouMadoKamei " +
                                            " ) VALUES( " +
                                            ankenNo +                                                       // GyoumuJouhouID
                                            "," + GlobalMethod.getSaiban("GyoumuJouhouMadoguchiID") + " " + // GyoumuJouhouMadoguchiID
                                            "," + gyoumuJouhouDT.Rows[i][0].ToString() + " " +              // GyoumuJouhouMadoKojinCD
                                            ",N'" + gyoumuJouhouDT.Rows[i][1].ToString() + "' " +            // GyoumuJouhouMadoChousainMei
                                            ",N'" + gyoumuJouhouDT.Rows[i][2].ToString() + "' " +            // GyoumuJouhouMadoGyoumuBushoCD
                                            ",N'" + gyoumuJouhouDT.Rows[i][3].ToString() + "' " +            // GyoumuJouhouMadoShibuMei
                                            ",N'" + gyoumuJouhouDT.Rows[i][4].ToString() + "' " +            // GyoumuJouhouMadoKamei
                                            ") ";
                                    Console.WriteLine(cmd.CommandText);
                                    result = cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyoutenBusho", ""))
                        {
                            cmd.CommandText = "INSERT INTO GyoumuJouhouHyoutenBusho ( " +
                                    "GyoumuJouhouID " +
                                    ",HyoutenBushoID " +

                                    ",HyoutenKyouryokuBushoID " +
                                    ",HyoutenKyouryokuBushoMei " +
                                    " ) SELECT " +
                                    ankenNo +
                                    ",HyoutenBushoID " +

                                    ",HyoutenKyouryokuBushoID " +
                                    ",HyoutenKyouryokuBushoMei " +
                                    " FROM GyoumuJouhouHyoutenBusho WHERE GyoumuJouhouHyoutenBusho.GyoumuJouhouID = " + AnkenID;
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "KeiyakuJouhouEntoryID", "KeiyakuJouhouEntory", ""))
                        {
                            cmd.CommandText = "INSERT INTO KeiyakuJouhouEntory ( " +
                                    "AnkenJouhouID " +
                                    ",KeiyakuJouhouEntoryID " +
                                    ",KeiyakuKeiyakuKingaku " +
                                    ",KeiyakuZeikomiKingaku " +
                                    ",KeiyakuuchizeiKingaku " +
                                    ",KeiyakuUriageHaibunCho " +
                                    ",KeiyakuUriageHaibunGakuCho1 " +
                                    ",KeiyakuUriageHaibunGakuCho2 " +
                                    ",KeiyakuUriageHaibunJo " +
                                    ",KeiyakuUriageHaibunGakuJo1 " +
                                    ",KeiyakuUriageHaibunGakuJo2 " +
                                    ",KeiyakuUriageHaibunJosys " +
                                    ",KeiyakuUriageHaibunGakuJosys1 " +
                                    ",KeiyakuUriageHaibunGakuJosys2 " +
                                    ",KeiyakuUriageHaibunKei " +
                                    ",KeiyakuUriageHaibunGakuKei1 " +
                                    ",KeiyakuUriageHaibunGakuKei2 " +
                                    ",KeiyakuZentokin " +
                                    ",KeiyakuSeikyuuKingaku1 " +
                                    ",KeiyakuSeikyuuKingaku2 " +
                                    ",KeiyakuSeikyuuKingaku3 " +
                                    ",KeiyakuSeikyuuKingaku4 " +
                                    ",KeiyakuSeikyuuKingaku5 " +
                                    ",KeiyakuCreateDate " +
                                    ",KeiyakuCreateUser " +
                                    ",KeiyakuCreateProgram " +
                                    ",KeiyakuUpdateDate " +
                                    ",KeiyakuUpdateUser " +
                                    ",KeiyakuUpdateProgram " +
                                    ",KeiyakuBetsuKeiyakuKingaku " +
                                    ",KeiyakuKeiyakuKingakuKei " +
                                    ",KeiyakuUriageHaibunChoGoukei " +
                                    ",KeiyakuUriageHaibunJoGoukei " +
                                    ",KeiyakuUriageHaibunJosysGoukei " +
                                    ",KeiyakuUriageHaibunKeiGoukei " +
                                    ",KeiyakuUriageHaibunGoukei " +
                                    ",KeiyakuHaibunChoZeinuki " +
                                    ",KeiyakuHaibunJoZeinuki " +
                                    ",KeiyakuHaibunJosysZeinuki " +
                                    ",KeiyakuHaibunKeiZeinuki " +
                                    ",KeiyakuHaibunZeinukiKei " +
                                    ",KeiyakuDeleteFlag " +

                                    ",KeiyakuSakuseibi " +
                                    ",KeiyakuSakuseiKubunID " +
                                    ",KeiyakuSakuseiKubun " +
                                    ",KeiyakuHachuushaMei " +
                                    ",KeiyakuGyoumuKubun " +
                                    ",KeiyakuGyoumuMei " +
                                    ",JutakuBushoCD " +
                                    ",KeiyakuTantousha " +
                                    ",KeiyakuJutakubangou " +
                                    ",KeiyakuEdaban " +
                                    ",KeiyakuKianzumi " +
                                    ",KeiyakuNyuusatsuYoteibi " +
                                    ",KeiyakuKeiyakuTeiketsubi " +
                                    ",KeiyakuKeiyakuKoukiKaishibi " +
                                    ",KeiyakuKeiyakuKoukiKanryoubi " +
                                    ",KeiyakuShouhizeiritsu " +
                                    ",KeiyakuRIBCKeishiki " +
                                    ",KeiyakuUriageHaibunCho1 " +
                                    ",KeiyakuUriageHaibunCho2 " +
                                    ",KeiyakuUriageHaibunJo1 " +
                                    ",KeiyakuUriageHaibunJo2 " +
                                    ",KeiyakuUriageHaibunJosys1 " +
                                    ",KeiyakuUriageHaibunJosys2 " +
                                    ",KeiyakuUriageHaibunKei1 " +
                                    ",KeiyakuUriageHaibunKei2 " +
                                    ",KeiyakuHenkoukanryoubi " +
                                    ",KeiyakuHenkouChuushiRiyuu " +
                                    ",KeiyakuBikou " +
                                    ",KeiyakuShosha " +
                                    ",KeiyakuTokkiShiyousho " +
                                    ",KeiyakuMitsumorisho " +
                                    ",KeiyakuTanpinChousaMitsumorisho " +
                                    ",KeiyakuSonota " +
                                    ",KeiyakuSonotaNaiyou " +
                                    ",KeiyakuSeikyuubi " +
                                    ",KeiyakuKeiyakusho " +
                                    ",KeiyakuZentokinUkewatashibi " +
                                    ",KeiyakuSeikyuusaki " +
                                    ",KeiyakuSeikyuuTaishouKoukiS1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE1 " +
                                    ",KeiyakuSeikyuubi1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE2 " +
                                    ",KeiyakuSeikyuubi2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS3 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE3 " +
                                    ",KeiyakuSeikyuubi3 " +
                                    ",KeiyakuKankeibusho1 " +
                                    ",KeiyakuKankeibusho2 " +
                                    ",KeiyakuKankeibusho3 " +
                                    ",KeiyakuKankeibusho4 " +
                                    ",KeiyakuKankeibusho5 " +
                                    ",KeiyakuKankeibusho6 " +
                                    ",KeiyakuKankeibusho7 " +
                                    ",KeiyakuKankeibusho8 " +
                                    ",KeiyakuKankeibusho9 " +
                                    ",KeiyakuKankeibusho10 " +
                                    ",KeiyakuKankeibusho11 " +
                                    ",KeiyakuKankeibusho12 " +
                                    ",KeiyakuKankeibusho14 " +
                                    ",KeiyakuKankeibusho15 " +
                                    ",KeiyakuKankeibusho13 " +
                                    ",KeiyakuNyuukinYoteibi " +
                                    ",KeiyakuUriageHaibunCho1Mei " +
                                    ",KeiyakuUriageHaibunCho2Mei " +
                                    ",KeiyakuUriageHaibunJo1Mei " +
                                    ",KeiyakuUriageHaibunJo2Mei " +
                                    ",KeiyakuUriageHaibunJosys1Mei " +
                                    ",KeiyakuUriageHaibunJosys2Mei " +
                                    ",KeiyakuUriageHaibunKei1Mei " +
                                    ",KeiyakuUriageHaibunKei2Mei " +
                                    ",KeiyakuUriageHaibunRIBC " +
                                    ",KeiyakuUriageHaibunRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC1Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC2 " +
                                    ",KeiyakuUriageHaibunRIBC2Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC2 " +
                                    ",KeiyakuSeikyuubi4 " +
                                    ",KeiyakuSeikyuubi5 " +
                                    ",KeiyakuTankeiMikomiCho " +
                                    ",KeiyakuTankeiMikomiJo " +
                                    ",KeiyakuTankeiMikomiJosys " +
                                    ",KeiyakuTankeiMikomiKei " +
                                    ",KeiyakuKurikoshiCho " +
                                    ",KeiyakuKurikoshiJo " +
                                    ",KeiyakuKurikoshiJosys " +
                                    ",KeiyakuKurikoshiKei " +
                                    " ) SELECT " +
                                    ankenNo +
                                    "," + ankenNo +
                                    ",- KeiyakuKeiyakuKingaku " +
                                    ",- KeiyakuZeikomiKingaku " +
                                    ",- KeiyakuuchizeiKingaku " +
                                    ",- KeiyakuUriageHaibunCho " +
                                    ",- KeiyakuUriageHaibunGakuCho1 " +
                                    ",- KeiyakuUriageHaibunGakuCho2 " +
                                    ",- KeiyakuUriageHaibunJo " +
                                    ",- KeiyakuUriageHaibunGakuJo1 " +
                                    ",- KeiyakuUriageHaibunGakuJo2 " +
                                    ",- KeiyakuUriageHaibunJosys " +
                                    ",- KeiyakuUriageHaibunGakuJosys1 " +
                                    ",- KeiyakuUriageHaibunGakuJosys2 " +
                                    ",- KeiyakuUriageHaibunKei " +
                                    ",- KeiyakuUriageHaibunGakuKei1 " +
                                    ",- KeiyakuUriageHaibunGakuKei2 " +
                                    ",- KeiyakuZentokin " +
                                    ",- KeiyakuSeikyuuKingaku1 " +
                                    ",- KeiyakuSeikyuuKingaku2 " +
                                    ",- KeiyakuSeikyuuKingaku3 " +
                                    ",- KeiyakuSeikyuuKingaku4 " +
                                    ",- KeiyakuSeikyuuKingaku5 " +
                                    ",GETDATE() " +
                                    ",'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +
                                    ",GETDATE() " +
                                    ",'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +
                                    ",- KeiyakuBetsuKeiyakuKingaku " +
                                    ",- KeiyakuKeiyakuKingakuKei " +
                                    ",- KeiyakuUriageHaibunChoGoukei " +
                                    ",- KeiyakuUriageHaibunJoGoukei " +
                                    ",- KeiyakuUriageHaibunJosysGoukei " +
                                    ",- KeiyakuUriageHaibunKeiGoukei " +
                                    ",- KeiyakuUriageHaibunGoukei " +
                                    ",- KeiyakuHaibunChoZeinuki " +
                                    ",- KeiyakuHaibunJoZeinuki " +
                                    ",- KeiyakuHaibunJosysZeinuki " +
                                    ",- KeiyakuHaibunKeiZeinuki " +
                                    ",- KeiyakuHaibunZeinukiKei " +
                                    ",0 " +

                                    ",KeiyakuSakuseibi " +
                                    ",KeiyakuSakuseiKubunID " +
                                    ",KeiyakuSakuseiKubun " +
                                    ",KeiyakuHachuushaMei " +
                                    ",KeiyakuGyoumuKubun " +
                                    ",KeiyakuGyoumuMei " +
                                    ",JutakuBushoCD " +
                                    ",KeiyakuTantousha " +
                                    ",KeiyakuJutakubangou " +
                                    ",KeiyakuEdaban " +
                                    ",KeiyakuKianzumi " +
                                    ",KeiyakuNyuusatsuYoteibi " +
                                    ",KeiyakuKeiyakuTeiketsubi " +
                                    ",KeiyakuKeiyakuKoukiKaishibi " +
                                    ",KeiyakuKeiyakuKoukiKanryoubi " +
                                    ",KeiyakuShouhizeiritsu " +
                                    ",KeiyakuRIBCKeishiki " +
                                    ",KeiyakuUriageHaibunCho1 " +
                                    ",KeiyakuUriageHaibunCho2 " +
                                    ",KeiyakuUriageHaibunJo1 " +
                                    ",KeiyakuUriageHaibunJo2 " +
                                    ",KeiyakuUriageHaibunJosys1 " +
                                    ",KeiyakuUriageHaibunJosys2 " +
                                    ",KeiyakuUriageHaibunKei1 " +
                                    ",KeiyakuUriageHaibunKei2 " +
                                    ",KeiyakuHenkoukanryoubi " +
                                    ",KeiyakuHenkouChuushiRiyuu " +
                                    ",KeiyakuBikou " +
                                    ",KeiyakuShosha " +
                                    ",KeiyakuTokkiShiyousho " +
                                    ",KeiyakuMitsumorisho " +
                                    ",KeiyakuTanpinChousaMitsumorisho " +
                                    ",KeiyakuSonota " +
                                    ",KeiyakuSonotaNaiyou " +
                                    ",KeiyakuSeikyuubi " +
                                    ",KeiyakuKeiyakusho " +
                                    ",KeiyakuZentokinUkewatashibi " +
                                    ",KeiyakuSeikyuusaki " +
                                    ",KeiyakuSeikyuuTaishouKoukiS1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE1 " +
                                    ",KeiyakuSeikyuubi1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE2 " +
                                    ",KeiyakuSeikyuubi2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS3 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE3 " +
                                    ",KeiyakuSeikyuubi3 " +
                                    ",KeiyakuKankeibusho1 " +
                                    ",KeiyakuKankeibusho2 " +
                                    ",KeiyakuKankeibusho3 " +
                                    ",KeiyakuKankeibusho4 " +
                                    ",KeiyakuKankeibusho5 " +
                                    ",KeiyakuKankeibusho6 " +
                                    ",KeiyakuKankeibusho7 " +
                                    ",KeiyakuKankeibusho8 " +
                                    ",KeiyakuKankeibusho9 " +
                                    ",KeiyakuKankeibusho10 " +
                                    ",KeiyakuKankeibusho11 " +
                                    ",KeiyakuKankeibusho12 " +
                                    ",KeiyakuKankeibusho14 " +
                                    ",KeiyakuKankeibusho15 " +
                                    ",KeiyakuKankeibusho13 " +
                                    ",KeiyakuNyuukinYoteibi " +
                                    ",KeiyakuUriageHaibunCho1Mei " +
                                    ",KeiyakuUriageHaibunCho2Mei " +
                                    ",KeiyakuUriageHaibunJo1Mei " +
                                    ",KeiyakuUriageHaibunJo2Mei " +
                                    ",KeiyakuUriageHaibunJosys1Mei " +
                                    ",KeiyakuUriageHaibunJosys2Mei " +
                                    ",KeiyakuUriageHaibunKei1Mei " +
                                    ",KeiyakuUriageHaibunKei2Mei " +
                                    ",KeiyakuUriageHaibunRIBC " +
                                    ",KeiyakuUriageHaibunRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC1Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC2 " +
                                    ",KeiyakuUriageHaibunRIBC2Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC2 " +
                                    ",KeiyakuSeikyuubi4 " +
                                    ",KeiyakuSeikyuubi5 " +
                                    // VIPS 20220408 課題管理表No.1301(997) CHANGE 変更伝票時、下記項目を赤伝にマイナスの値で登録
                                    //",KeiyakuTankeiMikomiCho " +
                                    //",KeiyakuTankeiMikomiJo " +
                                    //",KeiyakuTankeiMikomiJosys " +
                                    //",KeiyakuTankeiMikomiKei " +
                                    //",KeiyakuKurikoshiCho " +
                                    //",KeiyakuKurikoshiJo " +
                                    //",KeiyakuKurikoshiJosys " +
                                    //",KeiyakuKurikoshiKei " +
                                    ",- KeiyakuTankeiMikomiCho " +
                                    ",- KeiyakuTankeiMikomiJo " +
                                    ",- KeiyakuTankeiMikomiJosys " +
                                    ",- KeiyakuTankeiMikomiKei " +
                                    ",- KeiyakuKurikoshiCho " +
                                    ",- KeiyakuKurikoshiJo " +
                                    ",- KeiyakuKurikoshiJosys " +
                                    ",- KeiyakuKurikoshiKei " +
                                    " FROM KeiyakuJouhouEntory WHERE KeiyakuJouhouEntory.AnkenJouhouID = " + AnkenID;
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "RibcID", "RibcJouhou", ""))
                        {
                            cmd.CommandText = "INSERT INTO RibcJouhou ( " +
                                    "RibcID " +
                                    ",RibcNo " +
                                    ",RibcSeikyuKingaku " +

                                    ",RibcKoukiStart " +
                                    ",RibcKoukiEnd " +
                                    ",RibcSeikyubi " +
                                    ",RibcNouhinbi " +
                                    ",RibcNyukinyoteibi " +
                                    ",RibcUriageKeijyoTuki " +
                                    ",RibcKankeibusho " +
                                    ",RibcKubun " +
                                    ",RibcKankeibushoMei " +
                                    " ) SELECT " +
                                    ankenNo +
                                    ",RibcNo " +
                                    ",- RibcSeikyuKingaku " +

                                    ",RibcKoukiStart " +
                                    ",RibcKoukiEnd " +
                                    ",RibcSeikyubi " +
                                    ",RibcNouhinbi " +
                                    ",RibcNyukinyoteibi " +
                                    ",RibcUriageKeijyoTuki " +
                                    ",RibcKankeibusho " +
                                    ",RibcKubun " +
                                    ",RibcKankeibushoMei " +
                                    " FROM RibcJouhou WHERE RibcJouhou.RibcID = " + AnkenID;
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                        {
                            cmd.CommandText = "INSERT INTO NyuusatsuJouhou ( " +
                                    "AnkenJouhouID " +
                                    ",NyuusatsuJouhouID " +
                                    ",NyuusatsuMitsumorigaku " +
                                    ",NyuusatsuOusatugaku " +
                                    ",NyuusatsuRakusatugaku " +
                                    ",NyuusatsuRakusatuSougaku " +
                                    ",NyuusatsuNendoKurikoshigaku " +
                                    ",NyuusatsuKyougouTashaID " +
                                    ",NyuusatsuKyougouTasha " +
                                    ",NyuusatsuRakusatsushaID " +
                                    ",NyuusatsuRakusatsusha " +
                                    ",NyuusatsuYoteiKakaku " +
                                    ",NyuusatsuHoushiki " +
                                    ",NyuusatsuKeiyakukeitaiCDSaishuu " +
                                    ",NyuusatsuDenshiNyuusatsu " +
                                    ",NyuusatsuTanpinMikomigaku " +
                                    ",NyuusatsushaSuu " +
                                    ",NyuusatsuGyoumuBikou " +
                                    ",NyuusatsuShoruiSoufu " +
                                    ",NyuusatsuDeleteFlag " +
                                    ",NyuusatsuRakusatsuKekkaDate " +
                                    ",NyuusatsuCreateDate " +
                                    ",NyuusatsuCreateUser " +
                                    ",NyuusatsuCreateProgram " +
                                    ",NyuusatsuUpdateDate " +
                                    ",NyuusatsuUpdateUser " +
                                    ",NyuusatsuUpdateProgram " +

                                    ",NyuusatsuRakusatsuShaJokyou " +
                                    ",NyuusatsuRakusatsuGakuJokyou " +
                                    ",NyuusatsuRakusatsuShokaiDate " +
                                    ",NyuusatsuRakusatsuSaisyuDate " +
                                    ",NyuusatsuKekkaMemo " +
                                    " ) SELECT " +
                                    "" + ankenNo +
                                    "," + ankenNo +
                                    ",- NyuusatsuMitsumorigaku " +
                                    ",- NyuusatsuOusatugaku " +
                                    ",- NyuusatsuRakusatugaku " +
                                    ",- NyuusatsuRakusatuSougaku " +
                                    ",- NyuusatsuNendoKurikoshigaku " +
                                    ",CASE WHEN NyuusatsuKyougouTashaID > 0 THEN NyuusatsuKyougouTashaID ELSE NULL END " +
                                    ",CASE WHEN NyuusatsuKyougouTashaID > 0 THEN NyuusatsuKyougouTasha ELSE NULL END " +
                                    ",CASE WHEN NyuusatsuRakusatsushaID > 0 THEN NyuusatsuRakusatsushaID ELSE NULL END " +
                                    ",CASE WHEN NyuusatsuRakusatsushaID > 0 THEN NyuusatsuRakusatsusha ELSE NULL END " +
                                    ",NyuusatsuYoteiKakaku " +
                                    ",NyuusatsuHoushiki " +
                                    ",NyuusatsuKeiyakukeitaiCDSaishuu " +
                                    ",NyuusatsuDenshiNyuusatsu " +
                                    ",NyuusatsuTanpinMikomigaku " +
                                    ",NyuusatsushaSuu " +
                                    ",NyuusatsuGyoumuBikou " +
                                    ",NyuusatsuShoruiSoufu " +
                                    ",NyuusatsuDeleteFlag " +
                                    ",NyuusatsuRakusatsuKekkaDate " +
                                    ",GETDATE() " +
                                    ",N'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +
                                    ",GETDATE() " +
                                    ",'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +

                                    ",NyuusatsuRakusatsuShaJokyou " +
                                    ",NyuusatsuRakusatsuGakuJokyou " +
                                    ",NyuusatsuRakusatsuShokaiDate " +
                                    ",NyuusatsuRakusatsuSaisyuDate " +
                                    ",NyuusatsuKekkaMemo " +
                                    " FROM NyuusatsuJouhou WHERE NyuusatsuJouhou.NyuusatsuJouhouID = " + AnkenID;
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "入札情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhouOusatsusha", ""))
                        {
                            cmd.CommandText = "INSERT INTO NyuusatsuJouhouOusatsusha ( " +
                                    "NyuusatsuJouhouID" +
                                    ",NyuusatsuOusatsuID" +
                                    ",NyuusatsuOusatsuKingaku" +

                                    ",NyuusatsuOusatsushaID" +
                                    ",NyuusatsuOusatsusha" +
                                    ",NyuusatsuOusatsuKyougouTashaID" +
                                    ",NyuusatsuOusatsuKyougouKigyouCD" +
                                    ",NyuusatsuRakusatsuJyuni" +
                                    ",NyuusatsuRakusatsuJokyou" +
                                    ",NyuusatsuRakusatsuComment" +
                                    " ) SELECT " +
                                    ankenNo +
                                    ",ROW_NUMBER() OVER(ORDER BY NyuusatsuJouhouID) " +    // ",NyuusatsuOusatsuID" +
                                    ",- NyuusatsuOusatsuKingaku" +

                                    ",NyuusatsuOusatsushaID" +
                                    ",NyuusatsuOusatsusha" +
                                    ",NyuusatsuOusatsuKyougouTashaID" +
                                    ",NyuusatsuOusatsuKyougouKigyouCD" +
                                    ",NyuusatsuRakusatsuJyuni" +
                                    ",NyuusatsuRakusatsuJokyou" +
                                    ",NyuusatsuRakusatsuComment" +
                                    " FROM NyuusatsuJouhouOusatsusha WHERE NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID = " + AnkenID;
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }
                        DataTable GH_dt = new DataTable();
                        GH_dt = GlobalMethod.getData("GyoumuHaibunID", "GyoumuAnkenJouhouID", "GyoumuHaibun", "GyoumuAnkenJouhouID = " + AnkenID);
                        if (GH_dt != null && GH_dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < GH_dt.Rows.Count; i++)
                            {
                                cmd.CommandText = "INSERT INTO GyoumuHaibun ( " +
                                        "GyoumuHaibunID " +
                                        ",GyoumuAnkenJouhouID " +
                                        ",GyoumuChosaBuRitsu " +
                                        ",GyoumuChosaBuGaku " +
                                        ",GyoumuJigyoFukyuBuRitsu " +
                                        ",GyoumuJigyoFukyuBuGaku " +
                                        ",GyoumuJyohouSystemBuRitsu " +
                                        ",GyoumuJyohouSystemBuGaku " +
                                        ",GyoumuSougouKenkyuJoRitsu " +
                                        ",GyoumuSougouKenkyuJoGaku " +
                                        ",GyoumuShizaiChousaRitsu " +
                                        ",GyoumuShizaiChousaGaku " +
                                        ",GyoumuEizenRitsu " +
                                        ",GyoumuEizenGaku " +
                                        ",GyoumuKikiruiChousaRitsu " +
                                        ",GyoumuKikiruiChousaGaku " +
                                        ",GyoumuKoujiChousahiRitsu " +
                                        ",GyoumuKoujiChousahiGaku " +
                                        ",GyoumuSanpaiFukusanbutsuRitsu " +
                                        ",GyoumuSanpaiFukusanbutsuGaku " +
                                        ",GyoumuHokakeChousaRitsu " +
                                        ",GyoumuHokakeChousaGaku " +
                                        ",GyoumuShokeihiChousaRitsu " +
                                        ",GyoumuShokeihiChousaGaku " +
                                        ",GyoumuGenkaBunsekiRitsu " +
                                        ",GyoumuGenkaBunsekiGaku " +
                                        ",GyoumuKijunsakuseiRitsu " +
                                        ",GyoumuKijunsakuseiGaku " +
                                        ",GyoumuKoukyouRoumuhiRitsu " +
                                        ",GyoumuKoukyouRoumuhiGaku " +
                                        ",GyoumuRoumuhiKoukyouigaiRitsu " +
                                        ",GyoumuRoumuhiKoukyouigaiGaku " +
                                        ",GyoumuSonotaChousabuRitsu " +
                                        ",GyoumuSonotaChousabuGaku " +
                                        ",GyoumuHibunKubun " +
                                        " ) SELECT " +
                                        GlobalMethod.getSaiban("GyoumuHaibunID") +
                                        "," + ankenNo +
                                        ",GyoumuChosaBuRitsu " +
                                        ",- GyoumuChosaBuGaku " +
                                        ",GyoumuJigyoFukyuBuRitsu " +
                                        ",- GyoumuJigyoFukyuBuGaku " +
                                        ",GyoumuJyohouSystemBuRitsu " +
                                        ",- GyoumuJyohouSystemBuGaku " +
                                        ",GyoumuSougouKenkyuJoRitsu " +
                                        ",- GyoumuSougouKenkyuJoGaku " +
                                        ",GyoumuShizaiChousaRitsu " +
                                        ",- GyoumuShizaiChousaGaku " +
                                        ",GyoumuEizenRitsu " +
                                        ",- GyoumuEizenGaku " +
                                        ",GyoumuKikiruiChousaRitsu " +
                                        ",- GyoumuKikiruiChousaGaku " +
                                        ",GyoumuKoujiChousahiRitsu " +
                                        ",- GyoumuKoujiChousahiGaku " +
                                        ",GyoumuSanpaiFukusanbutsuRitsu " +
                                        ",- GyoumuSanpaiFukusanbutsuGaku " +
                                        ",GyoumuHokakeChousaRitsu " +
                                        ",- GyoumuHokakeChousaGaku " +
                                        ",GyoumuShokeihiChousaRitsu " +
                                        ",- GyoumuShokeihiChousaGaku " +
                                        ",GyoumuGenkaBunsekiRitsu " +
                                        ",- GyoumuGenkaBunsekiGaku " +
                                        ",GyoumuKijunsakuseiRitsu " +
                                        ",- GyoumuKijunsakuseiGaku " +
                                        ",GyoumuKoukyouRoumuhiRitsu " +
                                        ",- GyoumuKoukyouRoumuhiGaku " +
                                        ",GyoumuRoumuhiKoukyouigaiRitsu " +
                                        ",- GyoumuRoumuhiKoukyouigaiGaku " +
                                        ",GyoumuSonotaChousabuRitsu " +
                                        ",- GyoumuSonotaChousabuGaku " +
                                        ",GyoumuHibunKubun " +
                                        " FROM GyoumuHaibun WHERE GyoumuHaibun.GyoumuHaibunID = " + GetInt(GH_dt.Rows[i][1].ToString());
                                Console.WriteLine(cmd.CommandText);
                                result = cmd.ExecuteNonQuery();
                            }
                        }
                        int ankenNo2 = 0;
                        //黒伝作成
                        if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                        {
                            // 黒伝のAnkenJouhouID
                            ankenNo2 = GlobalMethod.getSaiban("AnkenJouhouID");
                            item3_1_20_kuroden.Text = ankenNo2.ToString();

                            cmd.CommandText = "INSERT INTO AnkenJouhou ( " +
                                    "AnkenJouhouID " +
                                    ",AnkenSakuseiKubun " +
                                    ",AnkenKaisuu " +
                                    ",AnkenCreateDate " +
                                    ",AnkenCreateUser " +
                                    ",AnkenCreateProgram " +
                                    ",AnkenUpdateDate " +
                                    ",AnkenUpdateUser " +
                                    ",AnkenSaishinFlg " +
                                    ",AnkenUpdateProgram " +
                                    //",AnkenUriagebi " +
                                    ",AnkenTourokubi " +
                                    ",AnkenGyoumuMei " +
                                    ",AnkenDeleteFlag " +

                                ",AnkenKishuKeikakugaku " +
                                ",AnkenKishuKeikakakugakuJf " +
                                ",AnkenKishuKeikakugakuJ " +
                                ",AnkenKeikakuZangaku " +
                                ",AnkenkeikakuZangakuJF " +
                                ",AnkenkeikakuZangakuJ " +
                                ",AnkenChokusetsuGenka " +
                                ",AnkenChokusetsuGenkaRitsu " +
                                ",AnkenGaichuuhi " +
                                ",AnkenJoukanDoboku " +
                                ",AnkenJoukanFukugou " +
                                ",AnkenJoukanGesuidou " +
                                ",AnkenJoukanHyoujun " +
                                ",AnkenJoukanIchiba " +
                                ",AnkenJoukanItiji " +
                                ",AnkenJoukanJutakuSonota " +
                                ",AnkenJoukanKentiku " +
                                ",AnkenJoukanKijunsho " +
                                ",AnkenJoukanKouwan " +
                                ",AnkenJoukanKuukou " +
                                ",AnkenJoukanSetsubi " +
                                ",AnkenJoukanSonota " +
                                ",AnkenJoukanSuidou " +
                                ",AnkenKeichoukaiKounyuuhi " +
                                ",AnkenKishuKeikakugakuK " +
                                ",AnkenUriageNendo " +
                                ",AnkenHachushaKubunCD " +
                                ",AnkenHachushaKubunMei " +
                                ",AnkenHachuushaCodeID " +
                                ",AnkenHachuushaMei " +
                                ",AnkenGyoumuKubun " +
                                ",AnkenGyoumuKubunMei " +
                                ",AnkenNyuusatsuHoushiki " +
                                ",AnkenKyougouTasha " +
                                ",AnkenJutakubushoCD " +
                                ",AnkenJutakushibu " +
                                ",AnkenTantoushaCD " +
                                ",AnkenMadoguchiTantoushaCD " +
                                ",AnkenGyoumuKanrishaCD " +
                                ",AnkenGyoumuKanrisha " +
                                ",GyoumuKanrishaCD " +
                                ",AnkenHachuushaBusho " +
                                ",AnkenkeikakuZangakuK " +
                                ",AnkenJutakuBangou " +
                                ",AnkenJutakuBangouEda " +
                                ",AnkenNyuusatsuYoteibi " +
                                ",AnkenRakusatsusha " +
                                ",AnkenRakusatsuJouhou " +
                                ",AnkenKianZumi " +
                                ",AnkenKiangetsu " +
                                ",AnkenHanteiKubun " +
                                ",AnkenJoukanData " +
                                ",AnkenJoukanHachuuKikanCD " +
                                ",AnkenNyuukinKakuninbi " +
                                ",AnkenKanryouSakuseibi " +
                                ",AnkenHonbuKakuninbi " +
                                ",AnkenShizaiChousa " +
                                ",AnkenKoujiChousahi " +
                                ",AnkenKikiruiChousa " +
                                ",AnkenSanpaiFukusanbutsu " +
                                ",AnkenHokakeChousa " +
                                ",AnkenShokeihiChousa " +
                                ",AnkenGenkaBunseki " +
                                ",AnkenKijunsakusei " +
                                ",AnkenKoukyouRoumuhi " +
                                ",AnkenRoumuhiKoukyouigai " +
                                ",AnkenSonotaChousabu " +
                                ",AnkenOrdermadeJifubu " +
                                ",AnkenRIBCJifubu " +
                                ",AnkenSonotaJifubu " +
                                ",AnkenOrdermade " +
                                ",AnkenJouhouKaihatsu " +
                                ",AnkenRIBCJouhouKaihatsu " +
                                ",AnkenSoukenbu " +
                                ",AnkenSonotaJoujibu " +
                                ",AnkenTeikiTokuchou " +
                                ",AnkenTanpinTokuchou " +
                                ",AnkenKikiChousa " +
                                ",AnkenHachuushaIraibusho " +
                                ",AnkenHachuushaTantousha " +
                                ",AnkenHachuushaTEL " +
                                ",AnkenHachuushaFAX " +
                                ",AnkenHachuushaMail " +
                                ",AnkenHachuushaIraiYuubin " +
                                ",AnkenHachuushaIraiJuusho " +
                                ",AnkenHachuushaKeiyakuBusho " +
                                ",AnkenHachuushaKeiyakuTantou " +
                                ",AnkenHachuushaKeiyakuTEL " +
                                ",AnkenHachuushaKeiyakuFAX " +
                                ",AnkenHachuushaKeiyakuMail " +
                                ",AnkenHachuushaKeiyakuYuubin " +
                                ",AnkenHachuushaKeiyakuJuusho " +
                                ",AnkenHachuuDaihyouYakushoku " +
                                ",AnkenHachuuDaihyousha " +
                                ",AnkenRosenKawamei " +
                                ",AnkenGyoumuItakuKasho " +
                                ",AnkenJititaiKibunID " +
                                ",AnkenJititaiKubun " +
                                ",AnkenKeiyakuToshoNo " +
                                ",AnkenKirokuToshoNo " +
                                ",AnkenKirokuHokanNo " +
                                ",AnkenCDHokan " +
                                ",AnkenSeikaButsuHokanFile " +
                                ",AnkenSeikabutsuHokanbako " +
                                ",AnkenKokyakuHyoukaComment " +
                                ",AnkenToukaiHyoukaComment " +
                                ",AnkenKenCD " +
                                ",AnkenToshiCD " +
                                ",AnkenKeiyakusho " +
                                ",AnkenEizen " +
                                ",AnkenTantoushaMei " +
                                ",GyoumuKanrishaMei " +
                                ",AnkenGyoumuKubunCD " +
                                ",AnkenHachuushaKaMei " +
                                ",AnkenKeiyakuKoukiKaishibi " +
                                ",AnkenKeiyakuKoukiKanryoubi " +
                                ",AnkenKeiyakuTeiketsubi " +
                                ",AnkenKeiyakuZeikomiKingaku " +
                                ",AnkenKeiyakuUriageHaibunGakuC " +
                                ",AnkenKeiyakuUriageHaibunGakuJ " +
                                ",AnkenKeiyakuUriageHaibunGakuJs " +
                                ",AnkenKeiyakuUriageHaibunGakuK " +
                                ",AnkenKeiyakuUriageHaibunGakuR " +
                                ",AnkenKeiyakuSakuseibi " +
                                ",AnkenAnkenBangou " +
                                ",AnkenKeikakuBangou " +
                                ",AnkenHikiaijhokyo " +
                                ",AnkenKeikakuAnkenMei " +
                                ",AnkenToukaiSankouMitsumori " +
                                ",AnkenToukaiJyutyuIyoku " +
                                ",AnkenToukaiSankouMitsumoriGaku " +
                                ",AnkenHachushaKaMei " +
                                ",AnkenHachushaCD " +
                                ",AnkenToukaiOusatu " +
                                ",AnkenKoukiNendo " +
                                    " ) SELECT " +
                                    ankenNo2 +
                                    ",'03' " +
                                    ",AnkenKaisuu + 1 " +
                                    ",GETDATE() " +
                                    ",'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +
                                    ",GETDATE() " +
                                    ",'" + UserInfos[0] + "' " +
                                    ",1 " +
                                    ",'ChangeKianEntry' " +
                                    //",''" +           
                                    ",AnkenTourokubi " +
                                    ",AnkenGyoumuMei " +
                                    ",AnkenDeleteFlag " +

                                ",AnkenKishuKeikakugaku " +
                                ",AnkenKishuKeikakakugakuJf " +
                                ",AnkenKishuKeikakugakuJ " +
                                ",AnkenKeikakuZangaku " +
                                ",AnkenkeikakuZangakuJF " +
                                ",AnkenkeikakuZangakuJ " +
                                ",AnkenChokusetsuGenka " +
                                ",AnkenChokusetsuGenkaRitsu " +
                                ",AnkenGaichuuhi " +
                                ",AnkenJoukanDoboku " +
                                ",AnkenJoukanFukugou " +
                                ",AnkenJoukanGesuidou " +
                                ",AnkenJoukanHyoujun " +
                                ",AnkenJoukanIchiba " +
                                ",AnkenJoukanItiji " +
                                ",AnkenJoukanJutakuSonota " +
                                ",AnkenJoukanKentiku " +
                                ",AnkenJoukanKijunsho " +
                                ",AnkenJoukanKouwan " +
                                ",AnkenJoukanKuukou " +
                                ",AnkenJoukanSetsubi " +
                                ",AnkenJoukanSonota " +
                                ",AnkenJoukanSuidou " +
                                ",AnkenKeichoukaiKounyuuhi " +
                                ",AnkenKishuKeikakugakuK " +
                                ",AnkenUriageNendo " +
                                ",AnkenHachushaKubunCD " +
                                ",AnkenHachushaKubunMei " +
                                ",AnkenHachuushaCodeID " +
                                ",AnkenHachuushaMei " +
                                ",AnkenGyoumuKubun " +
                                ",AnkenGyoumuKubunMei " +
                                ",AnkenNyuusatsuHoushiki " +
                                ",AnkenKyougouTasha " +
                                ",AnkenJutakubushoCD " +
                                ",AnkenJutakushibu " +
                                ",AnkenTantoushaCD " +
                                ",AnkenMadoguchiTantoushaCD " +
                                ",AnkenGyoumuKanrishaCD " +
                                ",AnkenGyoumuKanrisha " +
                                ",GyoumuKanrishaCD " +
                                ",AnkenHachuushaBusho " +
                                ",AnkenkeikakuZangakuK " +
                                ",AnkenJutakuBangou " +
                                ",AnkenJutakuBangouEda " +
                                ",AnkenNyuusatsuYoteibi " +
                                ",AnkenRakusatsusha " +
                                ",AnkenRakusatsuJouhou " +
                                ",AnkenKianZumi " +
                                ",AnkenKiangetsu " +
                                ",AnkenHanteiKubun " +
                                ",AnkenJoukanData " +
                                ",AnkenJoukanHachuuKikanCD " +
                                ",AnkenNyuukinKakuninbi " +
                                ",AnkenKanryouSakuseibi " +
                                ",AnkenHonbuKakuninbi " +
                                ",AnkenShizaiChousa " +
                                ",AnkenKoujiChousahi " +
                                ",AnkenKikiruiChousa " +
                                ",AnkenSanpaiFukusanbutsu " +
                                ",AnkenHokakeChousa " +
                                ",AnkenShokeihiChousa " +
                                ",AnkenGenkaBunseki " +
                                ",AnkenKijunsakusei " +
                                ",AnkenKoukyouRoumuhi " +
                                ",AnkenRoumuhiKoukyouigai " +
                                ",AnkenSonotaChousabu " +
                                ",AnkenOrdermadeJifubu " +
                                ",AnkenRIBCJifubu " +
                                ",AnkenSonotaJifubu " +
                                ",AnkenOrdermade " +
                                ",AnkenJouhouKaihatsu " +
                                ",AnkenRIBCJouhouKaihatsu " +
                                ",AnkenSoukenbu " +
                                ",AnkenSonotaJoujibu " +
                                ",AnkenTeikiTokuchou " +
                                ",AnkenTanpinTokuchou " +
                                ",AnkenKikiChousa " +
                                ",AnkenHachuushaIraibusho " +
                                ",AnkenHachuushaTantousha " +
                                ",AnkenHachuushaTEL " +
                                ",AnkenHachuushaFAX " +
                                ",AnkenHachuushaMail " +
                                ",AnkenHachuushaIraiYuubin " +
                                ",AnkenHachuushaIraiJuusho " +
                                ",AnkenHachuushaKeiyakuBusho " +
                                ",AnkenHachuushaKeiyakuTantou " +
                                ",AnkenHachuushaKeiyakuTEL " +
                                ",AnkenHachuushaKeiyakuFAX " +
                                ",AnkenHachuushaKeiyakuMail " +
                                ",AnkenHachuushaKeiyakuYuubin " +
                                ",AnkenHachuushaKeiyakuJuusho " +
                                ",AnkenHachuuDaihyouYakushoku " +
                                ",AnkenHachuuDaihyousha " +
                                ",AnkenRosenKawamei " +
                                ",AnkenGyoumuItakuKasho " +
                                ",AnkenJititaiKibunID " +
                                ",AnkenJititaiKubun " +
                                ",AnkenKeiyakuToshoNo " +
                                ",AnkenKirokuToshoNo " +
                                ",AnkenKirokuHokanNo " +
                                ",AnkenCDHokan " +
                                ",AnkenSeikaButsuHokanFile " +
                                ",AnkenSeikabutsuHokanbako " +
                                ",AnkenKokyakuHyoukaComment " +
                                ",AnkenToukaiHyoukaComment " +
                                ",AnkenKenCD " +
                                ",AnkenToshiCD " +
                                ",AnkenKeiyakusho " +
                                ",AnkenEizen " +
                                ",AnkenTantoushaMei " +
                                ",GyoumuKanrishaMei " +
                                ",AnkenGyoumuKubunCD " +
                                ",AnkenHachuushaKaMei " +
                                ",AnkenKeiyakuKoukiKaishibi " +
                                ",AnkenKeiyakuKoukiKanryoubi " +
                                ",AnkenKeiyakuTeiketsubi " +
                                ",AnkenKeiyakuZeikomiKingaku " +
                                ",AnkenKeiyakuUriageHaibunGakuC " +
                                ",AnkenKeiyakuUriageHaibunGakuJ " +
                                ",AnkenKeiyakuUriageHaibunGakuJs " +
                                ",AnkenKeiyakuUriageHaibunGakuK " +
                                ",AnkenKeiyakuUriageHaibunGakuR " +
                                ",AnkenKeiyakuSakuseibi " +
                                ",AnkenAnkenBangou " +
                                ",AnkenKeikakuBangou " +
                                ",AnkenHikiaijhokyo " +
                                ",AnkenKeikakuAnkenMei " +
                                ",AnkenToukaiSankouMitsumori " +
                                ",AnkenToukaiJyutyuIyoku " +
                                ",AnkenToukaiSankouMitsumoriGaku " +
                                ",AnkenHachushaKaMei " +
                                ",AnkenHachushaCD " +
                                ",AnkenToukaiOusatu " +
                                ",AnkenKoukiNendo " +
                                    " FROM AnkenJouhou WHERE AnkenJouhou.AnkenJouhouID = " + AnkenID;
                            result = cmd.ExecuteNonQuery();


                            if (!GlobalMethod.Check_Table(AnkenID, "KokyakuKeiyakuID", "KokyakuKeiyakuJouhou", ""))
                            {
                                GlobalMethod.outputLogger("ChangeKianEntry", "基本情報が見つからない", "ID:" + AnkenID, "DEBUG");
                                transaction.Rollback();
                                conn.Close();
                                return false;
                            }

                            cmd.CommandText = "INSERT INTO AnkenJouhouZenkaiRakusatsu ( " +
                                    "AnkenJouhouID " +
                                    ",AnkenZenkaiRakusatsuID " +

                                ",AnkenZenkaiJutakuKingaku " +
                                ",AnkenZenkaiJutakuBangou " +
                                ",AnkenZenkaiJutakuEdaban " +
                                ",AnkenZenkaiAnkenBangou " +
                                ",AnkenZenkaiRakusatsushaID " +
                                ",AnkenZenkaiRakusatsusha " +
                                ",AnkenZenkaiGyoumuMei " +
                                ",AnkenZenkaiKyougouKigyouCD " +
                                ",AnkenZenkaiJutakuZeinuki " +
                                ",KeiyakuZenkaiRakusatsushaID " +
                                    " ) SELECT " +
                                    ankenNo2 +
                                    ",AnkenZenkaiRakusatsuID " +

                                ",AnkenZenkaiJutakuKingaku " +
                                ",AnkenZenkaiJutakuBangou " +
                                ",AnkenZenkaiJutakuEdaban " +
                                ",AnkenZenkaiAnkenBangou " +
                                ",AnkenZenkaiRakusatsushaID " +
                                ",AnkenZenkaiRakusatsusha " +
                                ",AnkenZenkaiGyoumuMei " +
                                ",AnkenZenkaiKyougouKigyouCD " +
                                ",AnkenZenkaiJutakuZeinuki " +
                                ",KeiyakuZenkaiRakusatsushaID " +
                                    " FROM AnkenJouhouZenkaiRakusatsu WHERE AnkenJouhouZenkaiRakusatsu.AnkenJouhouID = " + AnkenID;
                            result = cmd.ExecuteNonQuery();

                            if (!GlobalMethod.Check_Table(AnkenID, "KokyakuKeiyakuID", "KokyakuKeiyakuJouhou", ""))
                            {
                                GlobalMethod.outputLogger("ChangeKianEntry", "顧客契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                                transaction.Rollback();
                                conn.Close();
                                return false;
                            }

                            cmd.CommandText = "INSERT INTO KokyakuKeiyakuJouhou ( " +
                                    "AnkenJouhouID " +
                                    ",KokyakuKeiyakuID " +
                                    ",KokyakuCreateUser   " +
                                    ",KokyakuCreateDate   " +
                                    ",KokyakuCreateProgram" +
                                    ",KokyakuUpdateUser   " +
                                    ",KokyakuUpdateDate   " +
                                    ",KokyakuDeleteFlag   " +
                                    ",KokyakuUpdateProgram" +

                                ",KokyakuKeiyakuTanka " +
                                ",KokyakuKeiyakuChosakuken " +
                                ",KokyakuKeiyakuKeisai " +
                                ",KokyakuKeiyakuTokchouChosaku " +
                                ",KokyakuKeiyakuRiyuu " +
                                ",KokyakuMaebaraiJoukou " +
                                ",KokyakuMaebaraiSeikyuu " +
                                ",KokyakuSekkeiTanka " +
                                ",KokyakuSekisanKijun " +
                                ",KokyakuKaiteiGetsu " +
                                ",KokyakuShichouson " +
                                ",KokyakuGijutsuCenter " +
                                ",KokyakuSonota " +
                                ",KokyakuKeiyakuRiyuuTou " +
                                ",KokyakuDataTeikyou " +
                                ",KokyakuAlpha " +
                                ",KokyakuDataDoboku " +
                                ",KokyakuDataNourin " +
                                ",KokyakuDataEizen " +
                                ",KokyakuDataSonota " +
                                ",KokyakuDataSekouP " +
                                ",KokyakuDataDobokuKouji " +
                                ",KokyakuDataRIBC " +
                                ",KokyakuDataGoukei " +
                                ",KokyakuDataKeisaiTanka " +
                                ",KokyakuDataWebTeikyou " +
                                ",KokyakuDataKeiyaku " +
                                ",KokyakuDataTempFile " +
                                ",KokyakuDataTempFileData " +
                                ",KokyakuData05Comment " +
                                ",KokyakuData06Comment " +
                                ",KokyakuData07Comment " +
                                ",KokyakuDataMeiki " +
                                ",KokyakuDataTeikyouTensu " +
                                    " ) SELECT " +
                                    ankenNo2 +
                                    "," + ankenNo2 +
                                    ",N'" + UserInfos[0] + "' " +
                                    ",GETDATE() " +
                                    ",'ChangeKianEntry' " +
                                    ",N'" + UserInfos[0] + "' " +
                                    ",GETDATE() " +
                                    ",0" +
                                    ",'ChangeKianEntry' " +

                                ",KokyakuKeiyakuTanka " +
                                ",KokyakuKeiyakuChosakuken " +
                                ",KokyakuKeiyakuKeisai " +
                                ",KokyakuKeiyakuTokchouChosaku " +
                                ",KokyakuKeiyakuRiyuu " +
                                ",KokyakuMaebaraiJoukou " +
                                ",KokyakuMaebaraiSeikyuu " +
                                ",KokyakuSekkeiTanka " +
                                ",KokyakuSekisanKijun " +
                                ",KokyakuKaiteiGetsu " +
                                ",KokyakuShichouson " +
                                ",KokyakuGijutsuCenter " +
                                ",KokyakuSonota " +
                                ",KokyakuKeiyakuRiyuuTou " +
                                ",KokyakuDataTeikyou " +
                                ",KokyakuAlpha " +
                                ",KokyakuDataDoboku " +
                                ",KokyakuDataNourin " +
                                ",KokyakuDataEizen " +
                                ",KokyakuDataSonota " +
                                ",KokyakuDataSekouP " +
                                ",KokyakuDataDobokuKouji " +
                                ",KokyakuDataRIBC " +
                                ",KokyakuDataGoukei " +
                                ",KokyakuDataKeisaiTanka " +
                                ",KokyakuDataWebTeikyou " +
                                ",KokyakuDataKeiyaku " +
                                ",KokyakuDataTempFile " +
                                ",KokyakuDataTempFileData " +
                                ",KokyakuData05Comment " +
                                ",KokyakuData06Comment " +
                                ",KokyakuData07Comment " +
                                ",KokyakuDataMeiki " +
                                ",KokyakuDataTeikyouTensu " +
                                " FROM KokyakuKeiyakuJouhou WHERE KokyakuKeiyakuJouhou.AnkenJouhouID = " + AnkenID;
                            result = cmd.ExecuteNonQuery();

                            if (!GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhou", ""))
                            {
                                GlobalMethod.outputLogger("ChangeKianEntry", "業務情報が見つからない", "ID:" + AnkenID, "DEBUG");
                                transaction.Rollback();
                                conn.Close();
                                return false;
                            }

                            cmd.CommandText = "INSERT INTO GyoumuJouhou ( " +
                                    "AnkenJouhouID " +
                                    ",GyoumuJouhouID " +
                                    ",GyoumuCreateDate " +
                                    ",GyoumuCreateUser    " +
                                    ",GyoumuCreateProgram " +
                                    ",GyoumuUpdateDate    " +
                                    ",GyoumuUpdateUser    " +
                                    ",GyoumuUpdateProgram " +
                                    ",GyoumuDeleteFlag " +

                                ",GyoumuHyouten " +
                                ",KanriGijutsushaCD " +
                                ",GyoumuKanriHyouten " +
                                ",ShousaTantoushaCD " +
                                ",SinsaTantoushaCD " +
                                ",GyoumuTECRISTourokuBangou " +
                                ",GyoumuKeisaiTankaTeikyou " +
                                ",GyoumuChosakukenJouto " +
                                ",GyoumuSeikyuubi " +
                                ",GyoumuSeikyuusho " +
                                ",GyoumuHikiwatashiNaiyou " +
                                ",KanriGijutsushaNM " +
                                ",ShousaTantoushaNM " +
                                ",SinsaTantoushaNM " +
                                ",GyoumuShousaHyouten " +
                                    " ) SELECT " +
                                    ankenNo2 +
                                    "," + ankenNo2 +
                                    ",GETDATE() " +
                                    ",N'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +
                                    ",GETDATE() " +
                                    ",N'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +
                                    ",0 " +

                                ",GyoumuHyouten " +
                                ",KanriGijutsushaCD " +
                                ",GyoumuKanriHyouten " +
                                ",ShousaTantoushaCD " +
                                ",SinsaTantoushaCD " +
                                ",GyoumuTECRISTourokuBangou " +
                                ",GyoumuKeisaiTankaTeikyou " +
                                ",GyoumuChosakukenJouto " +
                                ",GyoumuSeikyuubi " +
                                ",GyoumuSeikyuusho " +
                                ",GyoumuHikiwatashiNaiyou " +
                                ",KanriGijutsushaNM " +
                                ",ShousaTantoushaNM " +
                                ",SinsaTantoushaNM " +
                                ",GyoumuShousaHyouten " +
                                " FROM GyoumuJouhou WHERE GyoumuJouhou.AnkenJouhouID = " + AnkenID;
                            result = cmd.ExecuteNonQuery();

                            if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyouronTantouL1", ""))
                            {
                                cmd.CommandText = "INSERT INTO GyoumuJouhouHyouronTantouL1 ( " +
                                        "GyoumuJouhouID " +
                                        ",HyouronTantouID " +
                                    ",HyouronTantoushaCD " +
                                    ",HyouronTantoushaMei " +
                                    ",HyouronnTantoushaHyouten " +
                                        " ) SELECT " +
                                        ankenNo2 +
                                        ",HyouronTantouID " +
                                    ",HyouronTantoushaCD " +
                                    ",HyouronTantoushaMei " +
                                    ",HyouronnTantoushaHyouten " +
                                        " FROM GyoumuJouhouHyouronTantouL1 WHERE GyoumuJouhouHyouronTantouL1.GyoumuJouhouID = " + AnkenID;
                                result = cmd.ExecuteNonQuery();
                            }

                            if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouMadoguchi", ""))
                            {
                                //cmd.CommandText = "INSERT INTO GyoumuJouhouMadoguchi ( " +
                                //        "GyoumuJouhouID " +
                                //        ",GyoumuJouhouMadoguchiID " +
                                //        ",GyoumuJouhouMadoKojinCD " +
                                //        ",GyoumuJouhouMadoChousainMei " +
                                //    ",GyoumuJouhouMadoGyoumuBushoCD " +
                                //    ",GyoumuJouhouMadoShibuMei " +
                                //        " ) SELECT " +
                                //        ankenNo2 +
                                //        "," + GlobalMethod.getSaiban("GyoumuJouhouMadoguchiID") + " " +
                                //        ",GyoumuJouhouMadoKojinCD " +
                                //        ",GyoumuJouhouMadoChousainMei " +
                                //    ",GyoumuJouhouMadoGyoumuBushoCD " +
                                //    ",GyoumuJouhouMadoShibuMei " +
                                //        " FROM GyoumuJouhouMadoguchi WHERE GyoumuJouhouMadoguchi.GyoumuJouhouID = " + AnkenID;
                                //result = cmd.ExecuteNonQuery();

                                // 窓口担当者
                                // 新では1件しか入らないが、現行が複数件はいるので、複数件あった場合でも落ちないようにする
                                DataTable gyoumuJouhouDT = new DataTable();
                                cmd.CommandText = "SELECT " +
                                        " GyoumuJouhouMadoKojinCD " +
                                        ",GyoumuJouhouMadoChousainMei " +
                                        ",GyoumuJouhouMadoGyoumuBushoCD " +
                                        ",GyoumuJouhouMadoShibuMei " +
                                        ",GyoumuJouhouMadoKamei " +
                                        " FROM GyoumuJouhouMadoguchi WHERE GyoumuJouhouMadoguchi.GyoumuJouhouID = " + AnkenID;
                                var sda = new SqlDataAdapter(cmd);
                                sda.Fill(gyoumuJouhouDT);

                                if (gyoumuJouhouDT != null && gyoumuJouhouDT.Rows.Count > 0)
                                {
                                    for (int i = 0; i < gyoumuJouhouDT.Rows.Count; i++)
                                    {
                                        cmd.CommandText = "INSERT INTO GyoumuJouhouMadoguchi ( " +
                                                "GyoumuJouhouID " +
                                                ",GyoumuJouhouMadoguchiID " +
                                                ",GyoumuJouhouMadoKojinCD " +
                                                ",GyoumuJouhouMadoChousainMei " +
                                                ",GyoumuJouhouMadoGyoumuBushoCD " +
                                                ",GyoumuJouhouMadoShibuMei " +
                                                ",GyoumuJouhouMadoKamei " +
                                                " ) VALUES( " +
                                                ankenNo2 +                                                       // GyoumuJouhouID
                                                "," + GlobalMethod.getSaiban("GyoumuJouhouMadoguchiID") + " " + // GyoumuJouhouMadoguchiID
                                                "," + gyoumuJouhouDT.Rows[i][0].ToString() + " " +              // GyoumuJouhouMadoKojinCD
                                                ",N'" + gyoumuJouhouDT.Rows[i][1].ToString() + "' " +            // GyoumuJouhouMadoChousainMei
                                                ",N'" + gyoumuJouhouDT.Rows[i][2].ToString() + "' " +            // GyoumuJouhouMadoGyoumuBushoCD
                                                ",N'" + gyoumuJouhouDT.Rows[i][3].ToString() + "' " +            // GyoumuJouhouMadoShibuMei
                                                ",N'" + gyoumuJouhouDT.Rows[i][4].ToString() + "' " +            // GyoumuJouhouMadoKamei
                                                ") ";
                                        Console.WriteLine(cmd.CommandText);
                                        result = cmd.ExecuteNonQuery();
                                    }
                                }

                            }

                            if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyoutenBusho", ""))
                            {
                                cmd.CommandText = "INSERT INTO GyoumuJouhouHyoutenBusho ( " +
                                        "GyoumuJouhouID " +
                                        ",HyoutenBushoID " +
                                    ",HyoutenKyouryokuBushoID " +
                                    ",HyoutenKyouryokuBushoMei " +
                                        " ) SELECT " +
                                        ankenNo2 +
                                        ",HyoutenBushoID " +
                                    ",HyoutenKyouryokuBushoID " +
                                    ",HyoutenKyouryokuBushoMei " +
                                        " FROM GyoumuJouhouHyoutenBusho WHERE GyoumuJouhouHyoutenBusho.GyoumuJouhouID = " + AnkenID;
                                result = cmd.ExecuteNonQuery();
                            }

                            if (GlobalMethod.Check_Table(AnkenID, "KeiyakuJouhouEntoryID", "KeiyakuJouhouEntory", ""))
                            {
                                cmd.CommandText = "INSERT INTO KeiyakuJouhouEntory ( " +
                                        "AnkenJouhouID " +
                                        ",KeiyakuJouhouEntoryID " +
                                        ",KeiyakuCreateDate " +
                                        ",KeiyakuCreateUser " +
                                        ",KeiyakuCreateProgram " +
                                        ",KeiyakuUpdateDate " +
                                        ",KeiyakuUpdateUser " +
                                        ",KeiyakuUpdateProgram " +
                                        ",KeiyakuDeleteFlag " +

                                    ",KeiyakuKeiyakuKingaku " +
                                    ",KeiyakuZeikomiKingaku " +
                                    ",KeiyakuuchizeiKingaku " +
                                    ",KeiyakuUriageHaibunCho " +
                                    ",KeiyakuUriageHaibunGakuCho1 " +
                                    ",KeiyakuUriageHaibunGakuCho2 " +
                                    ",KeiyakuUriageHaibunJo " +
                                    ",KeiyakuUriageHaibunGakuJo1 " +
                                    ",KeiyakuUriageHaibunGakuJo2 " +
                                    ",KeiyakuUriageHaibunJosys " +
                                    ",KeiyakuUriageHaibunGakuJosys1 " +
                                    ",KeiyakuUriageHaibunGakuJosys2 " +
                                    ",KeiyakuUriageHaibunKei " +
                                    ",KeiyakuUriageHaibunGakuKei1 " +
                                    ",KeiyakuUriageHaibunGakuKei2 " +
                                    ",KeiyakuZentokin " +
                                    ",KeiyakuSeikyuuKingaku1 " +
                                    ",KeiyakuSeikyuuKingaku2 " +
                                    ",KeiyakuSeikyuuKingaku3 " +
                                    ",KeiyakuSeikyuuKingaku4 " +
                                    ",KeiyakuSeikyuuKingaku5 " +
                                    ",KeiyakuBetsuKeiyakuKingaku " +
                                    ",KeiyakuKeiyakuKingakuKei " +
                                    ",KeiyakuUriageHaibunChoGoukei " +
                                    ",KeiyakuUriageHaibunJoGoukei " +
                                    ",KeiyakuUriageHaibunJosysGoukei " +
                                    ",KeiyakuUriageHaibunKeiGoukei " +
                                    ",KeiyakuUriageHaibunGoukei " +
                                    ",KeiyakuHaibunChoZeinuki " +
                                    ",KeiyakuHaibunJoZeinuki " +
                                    ",KeiyakuHaibunJosysZeinuki " +
                                    ",KeiyakuHaibunKeiZeinuki " +
                                    ",KeiyakuHaibunZeinukiKei " +
                                    ",KeiyakuSakuseibi " +
                                    ",KeiyakuSakuseiKubunID " +
                                    ",KeiyakuSakuseiKubun " +
                                    ",KeiyakuHachuushaMei " +
                                    ",KeiyakuGyoumuKubun " +
                                    ",KeiyakuGyoumuMei " +
                                    ",JutakuBushoCD " +
                                    ",KeiyakuTantousha " +
                                    ",KeiyakuJutakubangou " +
                                    ",KeiyakuEdaban " +
                                    ",KeiyakuKianzumi " +
                                    ",KeiyakuNyuusatsuYoteibi " +
                                    ",KeiyakuKeiyakuTeiketsubi " +
                                    ",KeiyakuKeiyakuKoukiKaishibi " +
                                    ",KeiyakuKeiyakuKoukiKanryoubi " +
                                    ",KeiyakuShouhizeiritsu " +
                                    ",KeiyakuRIBCKeishiki " +
                                    ",KeiyakuUriageHaibunCho1 " +
                                    ",KeiyakuUriageHaibunCho2 " +
                                    ",KeiyakuUriageHaibunJo1 " +
                                    ",KeiyakuUriageHaibunJo2 " +
                                    ",KeiyakuUriageHaibunJosys1 " +
                                    ",KeiyakuUriageHaibunJosys2 " +
                                    ",KeiyakuUriageHaibunKei1 " +
                                    ",KeiyakuUriageHaibunKei2 " +
                                    ",KeiyakuHenkoukanryoubi " +
                                    ",KeiyakuHenkouChuushiRiyuu " +
                                    ",KeiyakuBikou " +
                                    ",KeiyakuShosha " +
                                    ",KeiyakuTokkiShiyousho " +
                                    ",KeiyakuMitsumorisho " +
                                    ",KeiyakuTanpinChousaMitsumorisho " +
                                    ",KeiyakuSonota " +
                                    ",KeiyakuSonotaNaiyou " +
                                    ",KeiyakuSeikyuubi " +
                                    ",KeiyakuKeiyakusho " +
                                    ",KeiyakuZentokinUkewatashibi " +
                                    ",KeiyakuSeikyuusaki " +
                                    ",KeiyakuSeikyuuTaishouKoukiS1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE1 " +
                                    ",KeiyakuSeikyuubi1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE2 " +
                                    ",KeiyakuSeikyuubi2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS3 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE3 " +
                                    ",KeiyakuSeikyuubi3 " +
                                    ",KeiyakuKankeibusho1 " +
                                    ",KeiyakuKankeibusho2 " +
                                    ",KeiyakuKankeibusho3 " +
                                    ",KeiyakuKankeibusho4 " +
                                    ",KeiyakuKankeibusho5 " +
                                    ",KeiyakuKankeibusho6 " +
                                    ",KeiyakuKankeibusho7 " +
                                    ",KeiyakuKankeibusho8 " +
                                    ",KeiyakuKankeibusho9 " +
                                    ",KeiyakuKankeibusho10 " +
                                    ",KeiyakuKankeibusho11 " +
                                    ",KeiyakuKankeibusho12 " +
                                    ",KeiyakuKankeibusho14 " +
                                    ",KeiyakuKankeibusho15 " +
                                    ",KeiyakuKankeibusho13 " +
                                    ",KeiyakuNyuukinYoteibi " +
                                    ",KeiyakuUriageHaibunCho1Mei " +
                                    ",KeiyakuUriageHaibunCho2Mei " +
                                    ",KeiyakuUriageHaibunJo1Mei " +
                                    ",KeiyakuUriageHaibunJo2Mei " +
                                    ",KeiyakuUriageHaibunJosys1Mei " +
                                    ",KeiyakuUriageHaibunJosys2Mei " +
                                    ",KeiyakuUriageHaibunKei1Mei " +
                                    ",KeiyakuUriageHaibunKei2Mei " +
                                    ",KeiyakuUriageHaibunRIBC " +
                                    ",KeiyakuUriageHaibunRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC1Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC2 " +
                                    ",KeiyakuUriageHaibunRIBC2Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC2 " +
                                    ",KeiyakuSeikyuubi4 " +
                                    ",KeiyakuSeikyuubi5 " +
                                    ",KeiyakuTankeiMikomiCho " +
                                    ",KeiyakuTankeiMikomiJo " +
                                    ",KeiyakuTankeiMikomiJosys " +
                                    ",KeiyakuTankeiMikomiKei " +
                                    ",KeiyakuKurikoshiCho " +
                                    ",KeiyakuKurikoshiJo " +
                                    ",KeiyakuKurikoshiJosys " +
                                    ",KeiyakuKurikoshiKei " +
                                        " ) SELECT " +
                                        ankenNo2 +
                                        "," + ankenNo2 +
                                        ",GETDATE() " +
                                        ",N'" + UserInfos[0] + "' " +
                                        ",'ChangeKianEntry' " +
                                        ",GETDATE() " +
                                        ",N'" + UserInfos[0] + "' " +
                                        ",'ChangeKianEntry' " +
                                        ",KeiyakuDeleteFlag " +

                                    ",KeiyakuKeiyakuKingaku " +
                                    ",KeiyakuZeikomiKingaku " +
                                    ",KeiyakuuchizeiKingaku " +
                                    ",KeiyakuUriageHaibunCho " +
                                    ",KeiyakuUriageHaibunGakuCho1 " +
                                    ",KeiyakuUriageHaibunGakuCho2 " +
                                    ",KeiyakuUriageHaibunJo " +
                                    ",KeiyakuUriageHaibunGakuJo1 " +
                                    ",KeiyakuUriageHaibunGakuJo2 " +
                                    ",KeiyakuUriageHaibunJosys " +
                                    ",KeiyakuUriageHaibunGakuJosys1 " +
                                    ",KeiyakuUriageHaibunGakuJosys2 " +
                                    ",KeiyakuUriageHaibunKei " +
                                    ",KeiyakuUriageHaibunGakuKei1 " +
                                    ",KeiyakuUriageHaibunGakuKei2 " +
                                    ",KeiyakuZentokin " +
                                    ",KeiyakuSeikyuuKingaku1 " +
                                    ",KeiyakuSeikyuuKingaku2 " +
                                    ",KeiyakuSeikyuuKingaku3 " +
                                    ",KeiyakuSeikyuuKingaku4 " +
                                    ",KeiyakuSeikyuuKingaku5 " +
                                    ",KeiyakuBetsuKeiyakuKingaku " +
                                    ",KeiyakuKeiyakuKingakuKei " +
                                    ",KeiyakuUriageHaibunChoGoukei " +
                                    ",KeiyakuUriageHaibunJoGoukei " +
                                    ",KeiyakuUriageHaibunJosysGoukei " +
                                    ",KeiyakuUriageHaibunKeiGoukei " +
                                    ",KeiyakuUriageHaibunGoukei " +
                                    ",KeiyakuHaibunChoZeinuki " +
                                    ",KeiyakuHaibunJoZeinuki " +
                                    ",KeiyakuHaibunJosysZeinuki " +
                                    ",KeiyakuHaibunKeiZeinuki " +
                                    ",KeiyakuHaibunZeinukiKei " +
                                    ",KeiyakuSakuseibi " +
                                    ",KeiyakuSakuseiKubunID " +
                                    ",KeiyakuSakuseiKubun " +
                                    ",KeiyakuHachuushaMei " +
                                    ",KeiyakuGyoumuKubun " +
                                    ",KeiyakuGyoumuMei " +
                                    ",JutakuBushoCD " +
                                    ",KeiyakuTantousha " +
                                    ",KeiyakuJutakubangou " +
                                    ",KeiyakuEdaban " +
                                    ",KeiyakuKianzumi " +
                                    ",KeiyakuNyuusatsuYoteibi " +
                                    ",KeiyakuKeiyakuTeiketsubi " +
                                    ",KeiyakuKeiyakuKoukiKaishibi " +
                                    ",KeiyakuKeiyakuKoukiKanryoubi " +
                                    ",KeiyakuShouhizeiritsu " +
                                    ",KeiyakuRIBCKeishiki " +
                                    ",KeiyakuUriageHaibunCho1 " +
                                    ",KeiyakuUriageHaibunCho2 " +
                                    ",KeiyakuUriageHaibunJo1 " +
                                    ",KeiyakuUriageHaibunJo2 " +
                                    ",KeiyakuUriageHaibunJosys1 " +
                                    ",KeiyakuUriageHaibunJosys2 " +
                                    ",KeiyakuUriageHaibunKei1 " +
                                    ",KeiyakuUriageHaibunKei2 " +
                                    ",KeiyakuHenkoukanryoubi " +
                                    ",KeiyakuHenkouChuushiRiyuu " +
                                    ",KeiyakuBikou " +
                                    ",KeiyakuShosha " +
                                    ",KeiyakuTokkiShiyousho " +
                                    ",KeiyakuMitsumorisho " +
                                    ",KeiyakuTanpinChousaMitsumorisho " +
                                    ",KeiyakuSonota " +
                                    ",KeiyakuSonotaNaiyou " +
                                    ",KeiyakuSeikyuubi " +
                                    ",KeiyakuKeiyakusho " +
                                    ",KeiyakuZentokinUkewatashibi " +
                                    ",KeiyakuSeikyuusaki " +
                                    ",KeiyakuSeikyuuTaishouKoukiS1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE1 " +
                                    ",KeiyakuSeikyuubi1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE2 " +
                                    ",KeiyakuSeikyuubi2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS3 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE3 " +
                                    ",KeiyakuSeikyuubi3 " +
                                    ",KeiyakuKankeibusho1 " +
                                    ",KeiyakuKankeibusho2 " +
                                    ",KeiyakuKankeibusho3 " +
                                    ",KeiyakuKankeibusho4 " +
                                    ",KeiyakuKankeibusho5 " +
                                    ",KeiyakuKankeibusho6 " +
                                    ",KeiyakuKankeibusho7 " +
                                    ",KeiyakuKankeibusho8 " +
                                    ",KeiyakuKankeibusho9 " +
                                    ",KeiyakuKankeibusho10 " +
                                    ",KeiyakuKankeibusho11 " +
                                    ",KeiyakuKankeibusho12 " +
                                    ",KeiyakuKankeibusho14 " +
                                    ",KeiyakuKankeibusho15 " +
                                    ",KeiyakuKankeibusho13 " +
                                    ",KeiyakuNyuukinYoteibi " +
                                    ",KeiyakuUriageHaibunCho1Mei " +
                                    ",KeiyakuUriageHaibunCho2Mei " +
                                    ",KeiyakuUriageHaibunJo1Mei " +
                                    ",KeiyakuUriageHaibunJo2Mei " +
                                    ",KeiyakuUriageHaibunJosys1Mei " +
                                    ",KeiyakuUriageHaibunJosys2Mei " +
                                    ",KeiyakuUriageHaibunKei1Mei " +
                                    ",KeiyakuUriageHaibunKei2Mei " +
                                    ",KeiyakuUriageHaibunRIBC " +
                                    ",KeiyakuUriageHaibunRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC1Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC2 " +
                                    ",KeiyakuUriageHaibunRIBC2Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC2 " +
                                    ",KeiyakuSeikyuubi4 " +
                                    ",KeiyakuSeikyuubi5 " +
                                    ",KeiyakuTankeiMikomiCho " +
                                    ",KeiyakuTankeiMikomiJo " +
                                    ",KeiyakuTankeiMikomiJosys " +
                                    ",KeiyakuTankeiMikomiKei " +
                                    ",KeiyakuKurikoshiCho " +
                                    ",KeiyakuKurikoshiJo " +
                                    ",KeiyakuKurikoshiJosys " +
                                    ",KeiyakuKurikoshiKei " +
                                        " FROM KeiyakuJouhouEntory WHERE KeiyakuJouhouEntory.AnkenJouhouID = " + AnkenID;
                                result = cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                GlobalMethod.outputLogger("ChangeKianEntry", "契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                                transaction.Rollback();
                                conn.Close();
                                return false;
                            }

                            if (GlobalMethod.Check_Table(AnkenID, "RibcID", "RibcJouhou", ""))
                            {
                                cmd.CommandText = "INSERT INTO RibcJouhou ( " +
                                        "RibcID " +
                                        ",RibcNo " +
                                        ",RibcSeikyuKingaku " +

                                    ",RibcKoukiStart " +
                                    ",RibcKoukiEnd " +
                                    ",RibcSeikyubi " +
                                    ",RibcNouhinbi " +
                                    ",RibcNyukinyoteibi " +
                                    ",RibcUriageKeijyoTuki " +
                                    ",RibcKankeibusho " +
                                    ",RibcKubun " +
                                    ",RibcKankeibushoMei " +
                                        " ) SELECT " +
                                        ankenNo2 +
                                        ",RibcNo " +
                                        ",RibcSeikyuKingaku " +

                                    ",RibcKoukiStart " +
                                    ",RibcKoukiEnd " +
                                    ",RibcSeikyubi " +
                                    ",RibcNouhinbi " +
                                    ",RibcNyukinyoteibi " +
                                    ",RibcUriageKeijyoTuki " +
                                    ",RibcKankeibusho " +
                                    ",RibcKubun " +
                                    ",RibcKankeibushoMei " +
                                        " FROM RibcJouhou WHERE RibcJouhou.RibcID = " + AnkenID;
                                result = cmd.ExecuteNonQuery();
                            }

                            if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                            {
                                cmd.CommandText = "INSERT INTO NyuusatsuJouhou ( " +
                                        "AnkenJouhouID " +
                                        ",NyuusatsuJouhouID " +
                                        ",NyuusatsuCreateDate " +
                                        ",NyuusatsuCreateUser " +
                                        ",NyuusatsuCreateProgram " +
                                        ",NyuusatsuUpdateDate " +
                                        ",NyuusatsuUpdateUser " +
                                        ",NyuusatsuUpdateProgram " +
                                        ",NyuusatsuKyougouTashaID " +
                                        ",NyuusatsuKyougouTasha " +
                                        ",NyuusatsuRakusatsushaID " +
                                        ",NyuusatsuRakusatsusha " +
                                        ",NyuusatsuRakusatugaku " +
                                        ",NyuusatsuOusatugaku " +
                                        ",NyuusatsuYoteiKakaku " +
                                        ",NyuusatsuHoushiki " +
                                        ",NyuusatsuKeiyakukeitaiCDSaishuu " +
                                        ",NyuusatsuDenshiNyuusatsu " +
                                        ",NyuusatsuTanpinMikomigaku " +
                                        ",NyuusatsushaSuu " +
                                        ",NyuusatsuGyoumuBikou " +
                                        ",NyuusatsuShoruiSoufu " +
                                        ",NyuusatsuDeleteFlag " +
                                        ",NyuusatsuRakusatuSougaku " +
                                        ",NyuusatsuRakusatsuKekkaDate " +
                                        ",NyuusatsuNendoKurikoshigaku " +

                                    ",NyuusatsuMitsumorigaku " +
                                    ",NyuusatsuRakusatsuShaJokyou " +
                                    ",NyuusatsuRakusatsuGakuJokyou " +
                                    ",NyuusatsuRakusatsuShokaiDate " +
                                    ",NyuusatsuRakusatsuSaisyuDate " +
                                    ",NyuusatsuKekkaMemo " +
                                        " ) SELECT " +
                                        ankenNo2 +
                                        "," + ankenNo2 +
                                        ",GETDATE() " +
                                        ",N'" + UserInfos[0] + "' " +
                                        ",'ChangeKianEntry' " +
                                        ",GETDATE() " +
                                        ",N'" + UserInfos[0] + "' " +
                                        ",'ChangeKianEntry' " +
                                        ",CASE WHEN NyuusatsuKyougouTashaID > 0 THEN NyuusatsuKyougouTashaID ELSE NULL END " +
                                        ",CASE WHEN NyuusatsuKyougouTashaID > 0 THEN NyuusatsuKyougouTasha ELSE NULL END " +
                                        ",CASE WHEN NyuusatsuRakusatsushaID > 0 THEN NyuusatsuRakusatsushaID ELSE NULL END " +
                                        ",CASE WHEN NyuusatsuRakusatsushaID > 0 THEN NyuusatsuRakusatsusha ELSE NULL END " +
                                        ",NyuusatsuRakusatugaku " +
                                        ",NyuusatsuOusatugaku " +
                                        ",NyuusatsuYoteiKakaku " +
                                        ",NyuusatsuHoushiki " +
                                        ",NyuusatsuKeiyakukeitaiCDSaishuu " +
                                        ",NyuusatsuDenshiNyuusatsu " +
                                        ",NyuusatsuTanpinMikomigaku " +
                                        ",NyuusatsushaSuu " +
                                        ",NyuusatsuGyoumuBikou " +
                                        ",NyuusatsuShoruiSoufu " +
                                        ",NyuusatsuDeleteFlag " +
                                        ",NyuusatsuRakusatuSougaku " +
                                        ",NyuusatsuRakusatsuKekkaDate " +
                                        ",NyuusatsuNendoKurikoshigaku " +

                                    ",NyuusatsuMitsumorigaku " +
                                    ",NyuusatsuRakusatsuShaJokyou " +
                                    ",NyuusatsuRakusatsuGakuJokyou " +
                                    ",NyuusatsuRakusatsuShokaiDate " +
                                    ",NyuusatsuRakusatsuSaisyuDate " +
                                    ",NyuusatsuKekkaMemo " +
                                        " FROM NyuusatsuJouhou WHERE NyuusatsuJouhou.NyuusatsuJouhouID = " + AnkenID;
                                result = cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                GlobalMethod.outputLogger("ChangeKianEntry", "入札情報が見つからない", "ID:" + AnkenID, "DEBUG");
                                transaction.Rollback();
                                conn.Close();
                                return false;
                            }

                            if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhouOusatsusha", ""))
                            {
                                cmd.CommandText = "INSERT INTO NyuusatsuJouhouOusatsusha ( " +
                                        "NyuusatsuJouhouID " +
                                        ",NyuusatsuOusatsuID" +
                                        ",NyuusatsuOusatsuKingaku" +

                                        ",NyuusatsuOusatsushaID" +
                                        ",NyuusatsuOusatsusha" +
                                        ",NyuusatsuOusatsuKyougouTashaID" +
                                        ",NyuusatsuOusatsuKyougouKigyouCD" +
                                        ",NyuusatsuRakusatsuJyuni" +
                                        ",NyuusatsuRakusatsuJokyou" +
                                        ",NyuusatsuRakusatsuComment" +
                                        " ) SELECT " +
                                        ankenNo2 +
                                        ",ROW_NUMBER() OVER(ORDER BY NyuusatsuJouhouID) " +
                                        ",NyuusatsuOusatsuKingaku" +

                                        ",NyuusatsuOusatsushaID" +
                                        ",NyuusatsuOusatsusha" +
                                        ",NyuusatsuOusatsuKyougouTashaID" +
                                        ",NyuusatsuOusatsuKyougouKigyouCD" +
                                        ",NyuusatsuRakusatsuJyuni" +
                                        ",NyuusatsuRakusatsuJokyou" +
                                        ",NyuusatsuRakusatsuComment" +
                                        " FROM NyuusatsuJouhouOusatsusha WHERE NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID = " + AnkenID;
                                result = cmd.ExecuteNonQuery();
                            }
                            if (GH_dt != null && GH_dt.Rows.Count > 0)
                            {
                                for (int i = 0; i < GH_dt.Rows.Count; i++)
                                {
                                    cmd.CommandText = "INSERT INTO GyoumuHaibun ( " +
                                            "GyoumuHaibunID " +
                                            ",GyoumuAnkenJouhouID " +
                                            ",GyoumuChosaBuRitsu " +
                                            ",GyoumuChosaBuGaku " +
                                            ",GyoumuJigyoFukyuBuRitsu " +
                                            ",GyoumuJigyoFukyuBuGaku " +
                                            ",GyoumuJyohouSystemBuRitsu " +
                                            ",GyoumuJyohouSystemBuGaku " +
                                            ",GyoumuSougouKenkyuJoRitsu " +
                                            ",GyoumuSougouKenkyuJoGaku " +
                                            ",GyoumuShizaiChousaRitsu " +
                                            ",GyoumuShizaiChousaGaku " +
                                            ",GyoumuEizenRitsu " +
                                            ",GyoumuEizenGaku " +
                                            ",GyoumuKikiruiChousaRitsu " +
                                            ",GyoumuKikiruiChousaGaku " +
                                            ",GyoumuKoujiChousahiRitsu " +
                                            ",GyoumuKoujiChousahiGaku " +
                                            ",GyoumuSanpaiFukusanbutsuRitsu " +
                                            ",GyoumuSanpaiFukusanbutsuGaku " +
                                            ",GyoumuHokakeChousaRitsu " +
                                            ",GyoumuHokakeChousaGaku " +
                                            ",GyoumuShokeihiChousaRitsu " +
                                            ",GyoumuShokeihiChousaGaku " +
                                            ",GyoumuGenkaBunsekiRitsu " +
                                            ",GyoumuGenkaBunsekiGaku " +
                                            ",GyoumuKijunsakuseiRitsu " +
                                            ",GyoumuKijunsakuseiGaku " +
                                            ",GyoumuKoukyouRoumuhiRitsu " +
                                            ",GyoumuKoukyouRoumuhiGaku " +
                                            ",GyoumuRoumuhiKoukyouigaiRitsu " +
                                            ",GyoumuRoumuhiKoukyouigaiGaku " +
                                            ",GyoumuSonotaChousabuRitsu " +
                                            ",GyoumuSonotaChousabuGaku " +
                                            ",GyoumuHibunKubun " +
                                            " ) SELECT " +
                                            GlobalMethod.getSaiban("GyoumuHaibunID") +
                                            "," + ankenNo2 +
                                            ",GyoumuChosaBuRitsu " +
                                            ",GyoumuChosaBuGaku " +
                                            ",GyoumuJigyoFukyuBuRitsu " +
                                            ",GyoumuJigyoFukyuBuGaku " +
                                            ",GyoumuJyohouSystemBuRitsu " +
                                            ",GyoumuJyohouSystemBuGaku " +
                                            ",GyoumuSougouKenkyuJoRitsu " +
                                            ",GyoumuSougouKenkyuJoGaku " +
                                            ",GyoumuShizaiChousaRitsu " +
                                            ",GyoumuShizaiChousaGaku " +
                                            ",GyoumuEizenRitsu " +
                                            ",GyoumuEizenGaku " +
                                            ",GyoumuKikiruiChousaRitsu " +
                                            ",GyoumuKikiruiChousaGaku " +
                                            ",GyoumuKoujiChousahiRitsu " +
                                            ",GyoumuKoujiChousahiGaku " +
                                            ",GyoumuSanpaiFukusanbutsuRitsu " +
                                            ",GyoumuSanpaiFukusanbutsuGaku " +
                                            ",GyoumuHokakeChousaRitsu " +
                                            ",GyoumuHokakeChousaGaku " +
                                            ",GyoumuShokeihiChousaRitsu " +
                                            ",GyoumuShokeihiChousaGaku " +
                                            ",GyoumuGenkaBunsekiRitsu " +
                                            ",GyoumuGenkaBunsekiGaku " +
                                            ",GyoumuKijunsakuseiRitsu " +
                                            ",GyoumuKijunsakuseiGaku " +
                                            ",GyoumuKoukyouRoumuhiRitsu " +
                                            ",GyoumuKoukyouRoumuhiGaku " +
                                            ",GyoumuRoumuhiKoukyouigaiRitsu " +
                                            ",GyoumuRoumuhiKoukyouigaiGaku " +
                                            ",GyoumuSonotaChousabuRitsu " +
                                            ",GyoumuSonotaChousabuGaku " +
                                            ",GyoumuHibunKubun " +
                                            " FROM GyoumuHaibun WHERE GyoumuHaibun.GyoumuHaibunID = " + GetInt(GH_dt.Rows[i][1].ToString());
                                    Console.WriteLine(cmd.CommandText);
                                    result = cmd.ExecuteNonQuery();
                                }
                            }

                        }
                        transaction.Commit();

                        transaction = conn.BeginTransaction();
                        cmd.Transaction = transaction;


                        if (GlobalMethod.Check_Table(AnkenID, "KeiyakuJouhouEntoryID", "KeiyakuJouhouEntory", ""))
                        {
                            cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET " +
                                     "KeiyakuHenkouChuushiRiyuu = N'" + item3_1_17.Text + "' " +
                                    ",KeiyakuSakuseibi = " + Get_DateTimePicker("item3_1_4");
                            if (SakuseiKubun == "02")
                            {
                                cmd.CommandText += ",KeiyakuKeiyakuTeiketsubi = " + Get_DateTimePicker("item3_1_3");
                            }
                            cmd.CommandText += " WHERE KeiyakuJouhouEntoryID = " + ankenNo;
                            result = cmd.ExecuteNonQuery();
                        }

                        if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                        {
                            cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                    "AnkenSakuseiKubun = N'" + item3_1_1.SelectedValue.ToString() + "' " +
                                    ",AnkenGyoumuKubun = N'" + item3_1_8.SelectedValue.ToString() + "' " +
                                    ",AnkenGyoumuKubunMei = N'" + item3_1_8.Text + "' " +
                                    ",AnkenGyoumuMei = N'" + item3_1_11.Text + "' " +
                                    ",AnkenKianzumi = '1' " +
                                    ",AnkenUriageNendo = N'" + item3_1_5.SelectedValue.ToString() + "' " +
                                    ",AnkenKeiyakuTeiketsubi = " + Get_DateTimePicker("item3_1_3") +
                                    ",AnkenKeiyakuKoukiKaishibi = " + Get_DateTimePicker("item3_1_6") +
                                    ",AnkenKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("item3_1_7") +
                                    ",AnkenKeiyakuZeikomiKingaku = " + item3_1_13.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuC = " + item3_2_1_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuJ = " + item3_2_2_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuJs = " + item3_2_3_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuK = " + item3_2_4_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuSakuseibi = " + Get_DateTimePicker("item3_1_4") +
                                    ",GyoumuKanrishaCD = " + "N'" + item3_4_4_CD.Text + "'" +
                                    ",GyoumuKanrishaMei = " + "N'" + item3_4_4.Text + "'" +
                                    ",AnkenUpdateUser = N'" + UserInfos[0] + "' " +
                                    //",AnkenKoukiNendo = '" + item1_2_KoukiNendo.SelectedValue.ToString() + "' " +
                                    " WHERE AnkenJouhou.AnkenJouhouID = " + ankenNo2;
                            result = cmd.ExecuteNonQuery();


                            //業務情報
                            cmd.CommandText = "UPDATE GyoumuJouhou SET " +
                                        "GyoumuHyouten = " + "N'" + item4_1_1.Text + "'" +
                                        ",KanriGijutsushaCD = " + "N'" + item3_4_1_CD.Text + "'" +
                                        ",KanriGijutsushaNM = " + "N'" + item3_4_1.Text + "'" +
                                        ",GyoumuKanriHyouten = " + "N'" + item3_4_1_Hyoten.Text + "'" +
                                        ",ShousaTantoushaCD = " + "N'" + item3_4_2_CD.Text + "'" +
                                        ",ShousaTantoushaNM = " + "N'" + item3_4_2.Text + "'" +
                                        ",GyoumuShousaHyouten = " + "N'" + item3_4_2_Hyoten.Text + "'" +
                                        ",SinsaTantoushaCD = " + "N'" + item3_4_3_CD.Text + "'" +
                                        ",SinsaTantoushaNM = " + "N'" + item3_4_3.Text + "'" +
                                        //",GyoumuTECRISTourokuBangou = " + "'" + item4_1_6.Text + "'" +
                                        //",GyoumuKeisaiTankaTeikyou = " + "''" +
                                        //",GyoumuChosakukenJouto = " + "''" +
                                        //",GyoumuSeikyuubi = " + Get_DateTimePicker("item4_1_7") +
                                        //",GyoumuSeikyuusho = " + "'" + GlobalMethod.ChangeSqlText(item4_1_8.Text, 0, 0) + "'" +
                                        //",GyoumuHikiwatashiNaiyou = " + "''" +
                                        ",GyoumuUpdateDate = " + " GETDATE() " +
                                        ",GyoumuUpdateUser = " + "N'" + UserInfos[0] + "' " +
                                        ",GyoumuUpdateProgram = " + "'UpdateEntory' " +
                                        ",GyoumuDeleteFlag = " + "0 " +
                                        " WHERE AnkenJouhouID = " + ankenNo2;
                            Console.WriteLine(cmd.CommandText);
                            cmd.ExecuteNonQuery();

                            //業務情報技術担当者
                            cmd.CommandText = "DELETE GyoumuJouhouHyouronTantouL1 WHERE GyoumuJouhouID = '" + ankenNo2 + "' ";
                            cmd.ExecuteNonQuery();

                            for (int i = 1; i < c1FlexGrid5.Rows.Count; i++)
                            {
                                if (c1FlexGrid3.Rows[i][1] != null && c1FlexGrid3.Rows[i][1].ToString() != "")
                                {
                                    string Hyouten = "";
                                    if (c1FlexGrid3.Rows[i][3] != null && c1FlexGrid3.Rows[i][3].ToString() != "")
                                    {
                                        Hyouten = c1FlexGrid3.Rows[i][3].ToString();
                                    }
                                    cmd.CommandText = "INSERT GyoumuJouhouHyouronTantouL1 ( " +
                                            "GyoumuJouhouID " +
                                            ", HyouronTantouID " +
                                            ", HyouronTantoushaCD " +
                                            ", HyouronTantoushaMei " +
                                            ", HyouronnTantoushaHyouten " +
                                            ") VALUES (" +
                                            "'" + ankenNo2 + "' " +
                                            "," + i +
                                            ",N'" + c1FlexGrid5.Rows[i][1].ToString() + "' " +
                                            ",N'" + c1FlexGrid5.Rows[i][2].ToString() + "' " +
                                            ",N'" + Hyouten + "' " +
                                            ") ";
                                    Console.WriteLine(cmd.CommandText);
                                    cmd.ExecuteNonQuery();
                                }
                            }

                            // 名称の取得
                            string GyoumuJouhouMadoShibuMei = "";
                            string GyoumuJouhouMadoKamei = "";
                            DataTable dt2 = new DataTable();
                            cmd.CommandText = "SELECT ShibuMei, KaMei FROM Mst_Busho WHERE GyoumuBushoCD = '" + item3_4_5_Busho.Text + "'";
                            var sda2 = new SqlDataAdapter(cmd);
                            dt2.Clear();
                            sda2.Fill(dt2);
                            if (dt2 != null && dt2.Rows.Count > 0)
                            {
                                GyoumuJouhouMadoShibuMei = dt2.Rows[0][0].ToString();
                                GyoumuJouhouMadoKamei = dt2.Rows[0][1].ToString();
                            }

                            //窓口担当者の更新
                            if ((item3_4_5_CD.Text == "0") || (item3_4_5_CD.Text == ""))
                            {
                                cmd.CommandText = "DELETE GyoumuJouhouMadoguchi WHERE GyoumuJouhouID = '" + ankenNo2 + "' ";
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                //cmd.CommandText = "UPDATE GyoumuJouhouMadoguchi  SET " +
                                //            " GyoumuJouhouMadoGyoumuBushoCD = " + "'" + item3_4_5_Busho.Text + "' " +
                                //            //", GyoumuJouhouMadoShibuMei = " + "'" + item3_4_5_Shibu.Text + "' " +
                                //            //", GyoumuJouhouMadoKamei = " + "'" + item3_4_5_Ka.Text + "' " +
                                //            ", GyoumuJouhouMadoShibuMei = " + "'" + GyoumuJouhouMadoShibuMei + "' " +
                                //            ", GyoumuJouhouMadoKamei = " + "'" + GyoumuJouhouMadoKamei + "' " +
                                //            ", GyoumuJouhouMadoKojinCD = " + "'" + item3_4_5_CD.Text + "' " +
                                //            ", GyoumuJouhouMadoChousainMei = " + "'" + item3_4_5.Text + "' " +
                                //            " WHERE GyoumuJouhouID =  " + ankenNo2;
                                //Console.WriteLine(cmd.CommandText);
                                //cmd.ExecuteNonQuery();

                                // 窓口担当者が複数いた場合の対応
                                DataTable gyoumuJouhouMadoguchiDT = new DataTable();
                                cmd.CommandText = "SELECT TOP 1 GyoumuJouhouMadoguchiID " +
                                                "FROM GyoumuJouhouMadoguchi " +
                                                "where GyoumuJouhouID = '" + ankenNo2 + "' " +
                                                "ORDER BY GyoumuJouhouMadoguchiID ";

                                string GyoumuJouhouMadoguchiID = "";

                                var gyoumuJouhouMadoguchiSda = new SqlDataAdapter(cmd);
                                gyoumuJouhouMadoguchiDT.Clear();
                                gyoumuJouhouMadoguchiSda.Fill(gyoumuJouhouMadoguchiDT);
                                if (gyoumuJouhouMadoguchiDT != null && gyoumuJouhouMadoguchiDT.Rows.Count > 0)
                                {
                                    GyoumuJouhouMadoguchiID = gyoumuJouhouMadoguchiDT.Rows[0][0].ToString();
                                }
                                // データが存在する場合
                                if (GyoumuJouhouMadoguchiID != "")
                                {
                                    cmd.CommandText = "UPDATE GyoumuJouhouMadoguchi set " +
                                                    "GyoumuJouhouMadoGyoumuBushoCD = N'" + item3_4_5_Busho.Text + "' " +
                                                    ",GyoumuJouhouMadoShibuMei = N'" + GyoumuJouhouMadoShibuMei + "' " +
                                                    ",GyoumuJouhouMadoKamei = N'" + GyoumuJouhouMadoKamei + "' " +
                                                    ",GyoumuJouhouMadoKojinCD = N'" + item3_4_5_CD.Text + "' " +
                                                    ",GyoumuJouhouMadoChousainMei = N'" + item3_4_5.Text + "' " +
                                                    "WHERE GyoumuJouhouMadoguchiID = N'" + GyoumuJouhouMadoguchiID + "' ";

                                    cmd.ExecuteNonQuery();
                                }

                            }

                            cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET " +
                                        "KeiyakuKeiyakuTeiketsubi = " + Get_DateTimePicker("item3_1_3") +
                                        ",KeiyakuSakuseibi = " + Get_DateTimePicker("item3_1_4") +
                                        ",KeiyakuKeiyakuKoukiKaishibi = " + Get_DateTimePicker("item3_1_6") +
                                        ",KeiyakuKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("item3_1_7") +
                                        ",KeiyakuKeiyakuKingaku = " + item3_1_12.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuZeikomiKingaku = " + item3_1_13.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuuchizeiKingaku = " + item3_1_14.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuShouhizeiritsu = N'" + item3_1_10.Text + "'" +
                                        ",KeiyakuHenkouChuushiRiyuu = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_17.Text, 0, 0) + "'" +
                                        ",KeiyakuBikou = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_19.Text, 0, 0) + "'" +
                                        ",KeiyakuShosha = " + (item3_1_20.Checked ? 1 : 0) +
                                        ",KeiyakuTokkiShiyousho = " + (item3_1_21.Checked ? 1 : 0) +
                                        ",KeiyakuMitsumorisho = " + (item3_1_22.Checked ? 1 : 0) +
                                        ",KeiyakuTanpinChousaMitsumorisho = " + (item3_1_23.Checked ? 1 : 0) +
                                        ",KeiyakuSonota = " + (item3_1_24.Checked ? 1 : 0) +
                                        ",KeiyakuSonotaNaiyou = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_25.Text, 0, 0) + "'" +
                                        ",KeiyakuZentokinUkewatashibi = " + Get_DateTimePicker("item3_6_11") +
                                        ",KeiyakuZentokin = " + item3_6_12.Text.Replace("¥", "").Replace(",", "") +
                                        ",Keiyakukeiyakukingakukei = " + item3_1_15.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuBetsuKeiyakuKingaku = " + item3_1_16.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuSeikyuubi1 = " + " " + Get_DateTimePicker("item3_6_1") + "" +
                                        ",KeiyakuSeikyuuKingaku1 = " + item3_6_2.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuSeikyuubi2 = " + " " + Get_DateTimePicker("item3_6_3") + "" +
                                        ",KeiyakuSeikyuuKingaku2 = " + item3_6_4.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuSeikyuubi3 = " + " " + Get_DateTimePicker("item3_6_5") + "" +
                                        ",KeiyakuSeikyuuKingaku3 = " + item3_6_6.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuSeikyuubi4 = " + " " + Get_DateTimePicker("item3_6_7") + "" +
                                        ",KeiyakuSeikyuuKingaku4 = " + item3_6_8.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuSeikyuubi5 = " + " " + Get_DateTimePicker("item3_6_9") + "" +
                                        ",KeiyakuSeikyuuKingaku5 = " + item3_6_10.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuSakuseiKubunID = " + "N'" + item3_1_1.SelectedValue + "'" +
                                        ",KeiyakuSakuseiKubun = " + "N'" + item3_1_1.Text + "'" +
                                        ",KeiyakuGyoumuKubun = " + "N'" + item3_1_8.SelectedValue + "'" +
                                        ",KeiyakuGyoumuMei = " + "N'" + item3_1_8.Text + "'" +
                                        //",KeiyakuJutakubangou = " + "'" + item1_7.Text + "'" +
                                        //",KeiyakuEdaban = " + "'" + item1_8.Text + "'" +
                                        ",KeiyakuKianzumi = " + (item3_1_2.Checked ? 1 : 0) +
                                        ",KeiyakuHachuushaMei = " + "N'" + item3_1_9.Text + "'" +
                                        ",KeiyakuHaibunChoZeinuki = " + item3_2_1_2.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuHaibunJoZeinuki = " + item3_2_2_2.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuHaibunJosysZeinuki = " + item3_2_3_2.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuHaibunKeiZeinuki = " + item3_2_4_2.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuHaibunZeinukiKei = " + item3_2_5_2.Text.Replace("¥", "").Replace(",", "").Replace("%", "") +
                                        ",KeiyakuUriageHaibunCho  = " + item3_2_1_1.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuUriageHaibunJo   = " + item3_2_2_1.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuUriageHaibunJosys  = " + item3_2_3_1.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuUriageHaibunKei  = " + item3_2_4_1.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuUriageHaibunGoukei = " + item3_2_5_1.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuTankeiMikomiCho  = " + item3_3_1.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuTankeiMikomiJo  = " + item3_3_2.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuTankeiMikomiJosys  = " + item3_3_3.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuTankeiMikomiKei  = " + item3_3_4.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuKurikoshiCho  = " + item3_7_1.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuKurikoshiJo  = " + item3_7_2.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuKurikoshiJosys  = " + item3_7_3.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuKurikoshiKei  = " + item3_7_4.Text.Replace("¥", "").Replace(",", "") +
                                        ",KeiyakuUpdateProgram = " + "'ChangeKianEntry'" +
                                        ",KeiyakuUpdateDate = " + "GETDATE()" +
                                        ",KeiyakuUpdateUser = " + "N'" + UserInfos[0] + "'" +
                                        // えんとり君修正STEP2（RIBC項目追加）
                                        ",KeiyakuRIBCYouTankaDataMoushikomisho = " + (item3_ribc_price.Checked ? 1 : 0) +
                                        ",KeiyakuSashaKeiyu = " + (item3_sa_commpany.Checked ? 1 : 0) +
                                        ",KeiyakuRIBCYouTankaData = " + (item3_1_ribc.Checked ? 1 : 0) +
                                " WHERE AnkenJouhouID = " + ankenNo2;
                            result = cmd.ExecuteNonQuery();

                            cmd.CommandText = "DELETE FROM RibcJouhou " +
                                    " WHERE RibcID = " + ankenNo2;
                            cmd.ExecuteNonQuery();
                            int cnt = 0;
                            string RibcKoukiStart;
                            string RibcNouhinbi;
                            string RibcSeikyubi;
                            string RibcNyukinyoteibi;
                            string RibcKubun;
                            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                            {

                                // 新では計上額のみでも登録を可とする
                                //if (c1FlexGrid4.Rows[i][1] != null)
                                //{
                                // 計上日、計上月、計上額のどれかが入っていれば登録する
                                // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                                if ((c1FlexGrid4.Rows[i][1] != null && c1FlexGrid4.Rows[i][1] != "")
                                    || (c1FlexGrid4.Rows[i][2] != null && c1FlexGrid4.Rows[i][2] != "")
                                    || (c1FlexGrid4.Rows[i][3] != null && c1FlexGrid4.Rows[i][3] != "0"))
                                {
                                    RibcKoukiStart = "null";
                                    if (c1FlexGrid4.Rows[i][4] != null)
                                    {
                                        RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][4].ToString() + "'";
                                    }
                                    RibcNouhinbi = "null";
                                    if (c1FlexGrid4.Rows[i][5] != null)
                                    {
                                        RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][5].ToString() + "'";
                                    }
                                    RibcSeikyubi = "null";
                                    if (c1FlexGrid4.Rows[i][6] != null)
                                    {
                                        RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][6].ToString() + "'";
                                    }
                                    RibcNyukinyoteibi = "null";
                                    if (c1FlexGrid4.Rows[i][7] != null)
                                    {
                                        RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][7].ToString() + "'";
                                    }
                                    RibcKubun = "0";
                                    if (c1FlexGrid4.Rows[i][8] != null)
                                    {
                                        RibcKubun = "N'" + c1FlexGrid4.Rows[i][8].ToString() + "'";
                                    }

                                    cnt++;
                                    cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                                "RibcID " +
                                                ",RibcNo " +
                                                ",RibcKoukiEnd " +
                                                ",RibcUriageKeijyoTuki " +
                                                ",RibcSeikyuKingaku " +
                                                ",RibcKankeibusho " +
                                                ",RibcKoukiStart " +
                                                ",RibcNouhinbi " +
                                                ",RibcSeikyubi " +
                                                ",RibcNyukinyoteibi " +
                                                ",RibcKubun " +
                                                ") VALUES (" +
                                               ankenNo2 +
                                                "," + cnt + "";
                                    //",'" + c1FlexGrid4.Rows[i][1].ToString() + "'" +
                                    //",'" + c1FlexGrid4.Rows[i][2].ToString() + "'" +
                                    //",'" + c1FlexGrid4.Rows[i][3].ToString().Replace("¥", "").Replace(",", "") + "'" +
                                    // 工期末日付 RibcKoukiEnd
                                    if (c1FlexGrid4.Rows[i][1] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][1].ToString() + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",null";
                                    }
                                    // 計上月 RibcUriageKeijyoTuki
                                    if (c1FlexGrid4.Rows[i][2] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][2].ToString() + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'' " + "";
                                    }
                                    // 計上額 RibcSeikyuKingaku
                                    if (c1FlexGrid4.Rows[i][3] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][3].ToString().Replace("¥", "").Replace(",", "") + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                    }
                                    cmd.CommandText = cmd.CommandText +
                                    ",'127120' " +
                                                "," + RibcKoukiStart +
                                                "," + RibcNouhinbi +
                                                "," + RibcSeikyubi +
                                                "," + RibcNyukinyoteibi +
                                                "," + RibcKubun +
                                                ")";

                                    Console.WriteLine(cmd.CommandText);
                                    cmd.ExecuteNonQuery();
                                }
                                //}

                                // 新では計上額のみでも登録を可とする
                                //if (c1FlexGrid4.Rows[i][9] != null)
                                //{
                                // 計上日、計上月、計上額のどれかが入っていれば登録する
                                // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                                if ((c1FlexGrid4.Rows[i][9] != null && c1FlexGrid4.Rows[i][9] != "")
                                    || (c1FlexGrid4.Rows[i][10] != null && c1FlexGrid4.Rows[i][10] != "")
                                    || (c1FlexGrid4.Rows[i][11] != null && c1FlexGrid4.Rows[i][11] != "0"))
                                {
                                    RibcKoukiStart = "null";
                                    if (c1FlexGrid4.Rows[i][12] != null)
                                    {
                                        RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][12].ToString() + "'";
                                    }
                                    RibcNouhinbi = "null";
                                    if (c1FlexGrid4.Rows[i][13] != null)
                                    {
                                        RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][13].ToString() + "'";
                                    }
                                    RibcSeikyubi = "null";
                                    if (c1FlexGrid4.Rows[i][14] != null)
                                    {
                                        RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][14].ToString() + "'";
                                    }
                                    RibcNyukinyoteibi = "null";
                                    if (c1FlexGrid4.Rows[i][15] != null)
                                    {
                                        RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][15].ToString() + "'";
                                    }
                                    RibcKubun = "0";
                                    if (c1FlexGrid4.Rows[i][16] != null)
                                    {
                                        RibcKubun = "N'" + c1FlexGrid4.Rows[i][16].ToString() + "'";
                                    }
                                    cnt++;
                                    cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                                "RibcID " +
                                                ",RibcNo " +
                                                ",RibcKoukiEnd " +
                                                ",RibcUriageKeijyoTuki " +
                                                ",RibcSeikyuKingaku " +
                                                ",RibcKankeibusho " +
                                                ",RibcKoukiStart " +
                                                ",RibcNouhinbi " +
                                                ",RibcSeikyubi " +
                                                ",RibcNyukinyoteibi " +
                                                ",RibcKubun " +
                                                ") VALUES (" +
                                               ankenNo2 +
                                                "," + cnt + "";
                                    //",'" + c1FlexGrid4.Rows[i][9].ToString() + "'" +
                                    //",'" + c1FlexGrid4.Rows[i][10].ToString() + "'" +
                                    //",'" + c1FlexGrid4.Rows[i][11].ToString().Replace("¥", "").Replace(",", "") + "'" +

                                    if (c1FlexGrid4.Rows[i][9] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][9].ToString() + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",null";
                                    }
                                    // 計上月 RibcUriageKeijyoTuki
                                    if (c1FlexGrid4.Rows[i][10] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][10].ToString() + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'' " + "";
                                    }
                                    // 計上額 RibcSeikyuKingaku
                                    if (c1FlexGrid4.Rows[i][11] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][11].ToString().Replace("¥", "").Replace(",", "") + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                    }
                                    cmd.CommandText = cmd.CommandText + ",'129230' " +
                                                "," + RibcKoukiStart +
                                                "," + RibcNouhinbi +
                                                "," + RibcSeikyubi +
                                                "," + RibcNyukinyoteibi +
                                                "," + RibcKubun +
                                                ")";

                                    cmd.ExecuteNonQuery();
                                }
                                //}

                                // 新では計上額のみでも登録を可とする
                                //if (c1FlexGrid4.Rows[i][17] != null)
                                //{
                                // 計上日、計上月、計上額のどれかが入っていれば登録する
                                // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                                if ((c1FlexGrid4.Rows[i][17] != null && c1FlexGrid4.Rows[i][17] != "")
                                    || (c1FlexGrid4.Rows[i][18] != null && c1FlexGrid4.Rows[i][18] != "")
                                    || (c1FlexGrid4.Rows[i][19] != null && c1FlexGrid4.Rows[i][19] != "0"))
                                {
                                    RibcKoukiStart = "null";
                                    if (c1FlexGrid4.Rows[i][20] != null)
                                    {
                                        RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][20].ToString() + "'";
                                    }
                                    RibcNouhinbi = "null";
                                    if (c1FlexGrid4.Rows[i][21] != null)
                                    {
                                        RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][21].ToString() + "'";
                                    }
                                    RibcSeikyubi = "null";
                                    if (c1FlexGrid4.Rows[i][22] != null)
                                    {
                                        RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][22].ToString() + "'";
                                    }
                                    RibcNyukinyoteibi = "null";
                                    if (c1FlexGrid4.Rows[i][23] != null)
                                    {
                                        RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][23].ToString() + "'";
                                    }
                                    RibcKubun = "0";
                                    if (c1FlexGrid4.Rows[i][24] != null)
                                    {
                                        RibcKubun = "N'" + c1FlexGrid4.Rows[i][24].ToString() + "'";
                                    }
                                    cnt++;
                                    cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                                "RibcID " +
                                                ",RibcNo " +
                                                ",RibcKoukiEnd " +
                                                ",RibcUriageKeijyoTuki " +
                                                ",RibcSeikyuKingaku " +
                                                ",RibcKankeibusho " +
                                                ",RibcKoukiStart " +
                                                ",RibcNouhinbi " +
                                                ",RibcSeikyubi " +
                                                ",RibcNyukinyoteibi " +
                                                ",RibcKubun " +
                                                ") VALUES (" +
                                               ankenNo2 +
                                                "," + cnt + "";
                                    //",'" + c1FlexGrid4.Rows[i][17].ToString() + "'" +
                                    //",'" + c1FlexGrid4.Rows[i][18].ToString() + "'" +
                                    //",'" + c1FlexGrid4.Rows[i][19].ToString().Replace("¥", "").Replace(",", "") + "'" +
                                    if (c1FlexGrid4.Rows[i][17] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][17].ToString() + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",null";
                                    }
                                    // 計上月 RibcUriageKeijyoTuki
                                    if (c1FlexGrid4.Rows[i][18] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][18].ToString() + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'' " + "";
                                    }
                                    // 計上額 RibcSeikyuKingaku
                                    if (c1FlexGrid4.Rows[i][19] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][19].ToString().Replace("¥", "").Replace(",", "") + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                    }
                                    // 年度により、情報システム部の部コードを変更する
                                    if (GetInt(item3_1_5.SelectedValue.ToString()) >= 2021)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'128400' " +
                                                    "," + RibcKoukiStart +
                                                    "," + RibcNouhinbi +
                                                    "," + RibcSeikyubi +
                                                    "," + RibcNyukinyoteibi +
                                                    "," + RibcKubun +
                                                    ")";
                                    }
                                    else
                                    {
                                        // 2021年度以前は127900
                                        cmd.CommandText = cmd.CommandText + ",'127900' " +
                                                    "," + RibcKoukiStart +
                                                    "," + RibcNouhinbi +
                                                    "," + RibcSeikyubi +
                                                    "," + RibcNyukinyoteibi +
                                                    "," + RibcKubun +
                                                    ")";
                                    }

                                    cmd.ExecuteNonQuery();
                                }
                                //}

                                // 新では計上額のみでも登録を可とする
                                //if (c1FlexGrid4.Rows[i][25] != null)
                                //{
                                // 計上日、計上月、計上額のどれかが入っていれば登録する
                                // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                                if ((c1FlexGrid4.Rows[i][25] != null && c1FlexGrid4.Rows[i][25] != "")
                                    || (c1FlexGrid4.Rows[i][26] != null && c1FlexGrid4.Rows[i][26] != "")
                                    || (c1FlexGrid4.Rows[i][27] != null && c1FlexGrid4.Rows[i][27] != "0"))
                                {
                                    RibcKoukiStart = "null";
                                    if (c1FlexGrid4.Rows[i][28] != null)
                                    {
                                        RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][28].ToString() + "'";
                                    }
                                    RibcNouhinbi = "null";
                                    if (c1FlexGrid4.Rows[i][29] != null)
                                    {
                                        RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][29].ToString() + "'";
                                    }
                                    RibcSeikyubi = "null";
                                    if (c1FlexGrid4.Rows[i][30] != null)
                                    {
                                        RibcSeikyubi = "'" + c1FlexGrid4.Rows[i][30].ToString() + "'";
                                    }
                                    RibcNyukinyoteibi = "null";
                                    if (c1FlexGrid4.Rows[i][31] != null)
                                    {
                                        RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][31].ToString() + "'";
                                    }
                                    RibcKubun = "0";
                                    if (c1FlexGrid4.Rows[i][32] != null)
                                    {
                                        RibcKubun = "N'" + c1FlexGrid4.Rows[i][32].ToString() + "'";
                                    }
                                    cnt++;
                                    cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                                "RibcID " +
                                                ",RibcNo " +
                                                ",RibcKoukiEnd " +
                                                ",RibcUriageKeijyoTuki " +
                                                ",RibcSeikyuKingaku " +
                                                ",RibcKankeibusho " +
                                                ",RibcKoukiStart " +
                                                ",RibcNouhinbi " +
                                                ",RibcSeikyubi " +
                                                ",RibcNyukinyoteibi " +
                                                ",RibcKubun " +
                                                ") VALUES (" +
                                               ankenNo2 +
                                                "," + cnt + "";
                                    //",'" + c1FlexGrid4.Rows[i][25].ToString() + "'" +
                                    //",'" + c1FlexGrid4.Rows[i][26].ToString() + "'" +
                                    //",'" + c1FlexGrid4.Rows[i][27].ToString().Replace("¥", "").Replace(",", "") + "'" +

                                    if (c1FlexGrid4.Rows[i][25] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][25].ToString() + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",null";
                                    }
                                    // 計上月 RibcUriageKeijyoTuki
                                    if (c1FlexGrid4.Rows[i][26] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][26].ToString() + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'' " + "";
                                    }
                                    // 計上額 RibcSeikyuKingaku
                                    if (c1FlexGrid4.Rows[i][27] != null)
                                    {
                                        cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][27].ToString().Replace("¥", "").Replace(",", "") + "'";
                                    }
                                    else
                                    {
                                        cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                    }
                                    cmd.CommandText = cmd.CommandText + ",'150200' " +
                                                "," + RibcKoukiStart +
                                                "," + RibcNouhinbi +
                                                "," + RibcSeikyubi +
                                                "," + RibcNyukinyoteibi +
                                                "," + RibcKubun +
                                                ")";

                                    cmd.ExecuteNonQuery();
                                }
                            }

                            //業務配分登録
                            cmd.CommandText = "UPDATE GyoumuHaibun SET " +
                                                " GyoumuChosaBuGaku " + " = N'" + item3_7_1_6_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuJigyoFukyuBuGaku " + " = N'" + item3_7_1_7_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuJyohouSystemBuGaku " + " = N'" + item3_7_1_8_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuSougouKenkyuJoGaku " + " = N'" + item3_7_1_9_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuShizaiChousaRitsu " + " = N'" + item3_7_2_14_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuEizenRitsu " + " = N'" + item3_7_2_15_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuKikiruiChousaRitsu " + " = N'" + item3_7_2_16_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuKoujiChousahiRitsu " + " = N'" + item3_7_2_17_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuSanpaiFukusanbutsuRitsu " + " = N'" + item3_7_2_18_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuHokakeChousaRitsu " + " = N'" + item3_7_2_19_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuShokeihiChousaRitsu " + " = N'" + item3_7_2_20_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuGenkaBunsekiRitsu " + " = N'" + item3_7_2_21_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuKijunsakuseiRitsu " + " = N'" + item3_7_2_22_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuKoukyouRoumuhiRitsu " + " = N'" + item3_7_2_23_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuRoumuhiKoukyouigaiRitsu " + " = N'" + item3_7_2_24_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuSonotaChousabuRitsu " + " = N'" + item3_7_2_25_1.Text.Replace("%", "") + "' " +
                                                ",GyoumuShizaiChousaGaku " + " = N'" + item3_7_2_14_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuEizenGaku " + " = N'" + item3_7_2_15_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuKikiruiChousaGaku " + " = N'" + item3_7_2_16_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuKoujiChousahiGaku " + " = N'" + item3_7_2_17_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuSanpaiFukusanbutsuGaku " + " = N'" + item3_7_2_18_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuHokakeChousaGaku " + " = N'" + item3_7_2_19_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuShokeihiChousaGaku " + " = N'" + item3_7_2_20_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuGenkaBunsekiGaku " + " = N'" + item3_7_2_21_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuKijunsakuseiGaku " + " = N'" + item3_7_2_22_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuKoukyouRoumuhiGaku " + " = N'" + item3_7_2_23_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuRoumuhiKoukyouigaiGaku " + " = N'" + item3_7_2_24_2.Text.Replace("%", "") + "' " +
                                                ",GyoumuSonotaChousabuGaku " + " = N'" + item3_7_2_25_2.Text.Replace("%", "") + "' " +
                                                " WHERE GyoumuAnkenJouhouID = " + ankenNo2 + " AND GyoumuHibunKubun = 30 ";
                            Console.WriteLine(cmd.CommandText);
                            cmd.ExecuteNonQuery();

                            // 窓口の案件情報IDを最新に置き換える
                            //if (GlobalMethod.Check_Table(AnkenID, "AnkenJouhouID", "MadoguchiJouhou", ""))
                            //{
                            cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                                        "AnkenJouhouID = '" + ankenNo2 + "' " +
                                        ",MadoguchiAnkenJouhouID = '" + ankenNo2 + "' " +
                                        " WHERE MadoguchiJouhou.AnkenJouhouID = " + AnkenID;
                            result = cmd.ExecuteNonQuery();

                            // 単価契約の案件情報IDを最新に置き換える
                            cmd.CommandText = "UPDATE TankaKeiyaku SET " +
                                    "AnkenJouhouID = '" + ankenNo2 + "' " +
                                    " WHERE TankaKeiyaku.AnkenJouhouID = " + AnkenID;
                            cmd.ExecuteNonQuery();
                            //}

                            // 変更伝票の起案で変更した管理技術者を更新する
                            cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                            // 1213 窓口ミハルの窓口担当者は、初期登録のみエントリくんとリンクさせる。
                            //"MadoguchiTantoushaCD = " + "'" + item3_4_5_CD.Text + "' " +
                            //",MadoguchiKanriGijutsusha = " + "'" + item3_4_1_CD.Text + "'" +
                            "MadoguchiKanriGijutsusha = " + "N'" + item3_4_1_CD.Text + "' " +
                            " WHERE AnkenJouhouID = " + ankenNo2;
                            cmd.ExecuteNonQuery();

                            //transaction.Commit();

                            //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "赤伝を作成しました ID:" + ankenNo, "ChangeKianEntry", "");
                            //if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                            //{
                            //    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "黒伝を作成しました ID:" + ankenNo2, "ChangeKianEntry", "");
                            //    set_error(GlobalMethod.GetMessage("I10710", ""));
                            //}
                            //else
                            //{
                            //    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "中止伝票を追加しました。 ID:" + AnkenID, "ChangeKianEntry", "");
                            //    set_error(GlobalMethod.GetMessage("I10711", ""));
                            //}
                        }

                        transaction.Commit();

                        //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "赤伝を作成しました ID:" + ankenNo, "ChangeKianEntry", "");
                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "赤伝を作成しました ID:" + ankenNo, pgmName + methodName, "");
                        if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                        {
                            //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "黒伝を作成しました ID:" + ankenNo2, "ChangeKianEntry", "");
                            GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "黒伝を作成しました ID:" + ankenNo2, pgmName + methodName, "");
                            set_error(GlobalMethod.GetMessage("I10710", ""));
                        }
                        else
                        {
                            //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "中止伝票を追加しました。 ID:" + AnkenID, "ChangeKianEntry", "");
                            GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "中止伝票を追加しました。 ID:" + AnkenID, pgmName + methodName, "");
                            set_error(GlobalMethod.GetMessage("I10711", ""));
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
            }

            return true;
        }

        // えんとり君修正STEP2 フォルダリムーブ処理
        private bool RenameFolder(string ori_ankenNo)
        {
            // えんとり君修正STEP2
            bool isMoveOk = false;
            string folderTo = GlobalMethod.ChangeSqlText(txt_renamedfolder.Text, 0, 0);
            string sJigyoubuHeadCD = getJigyoubuHeadCD();
            int isRename = 0; // 0:何もしない、1:リネーム、2:削除のみ、3：新規作成
            //if (sFolderRenameBef.Equals(folderTo) == false)
            if (string.IsNullOrEmpty(folderTo) == false && sFolderRenameBef.Equals(folderTo) == false)
            {
                // リネームボタン押下時
                if (string.IsNullOrEmpty(item3_1_20_reset_ankenno.Text) == false)
                {
                    folderTo = folderTo.Replace(item3_1_20_reset_ankenno.Text, item1_6.Text);
                }
                if (sJigyoubuHeadCD_ori.Equals("T") && sJigyoubuHeadCD.Equals("T"))
                {
                    //もっとファイル
                    if (item1_12.Text.Contains(ori_ankenNo))
                    {
                        // リネーム前後、すべて調査部の場合、リネームを実施する
                        isRename = 1;
                    }
                    else
                    {
                        isRename = 3;
                    }
                }
                else if (sJigyoubuHeadCD_ori.Equals("T"))
                {
                    // リネーム前のみ調査部なら、リネームで、もとフォルダを削除する
                    isRename = 2;
                }
                else if (sJigyoubuHeadCD.Equals("T"))
                {
                    // リネーム後のみ調査部なら、新規作成する
                    isRename = 3;
                }

            }
            else if (ori_ankenNo.Equals(item1_6.Text) == false)
            {
                folderTo = GlobalMethod.ChangeSqlText(item1_12.Text, 0, 0);
                // リネームボタン押下しない　AND　案件番号自動変更
                if (sJigyoubuHeadCD_ori.Equals("T") && sJigyoubuHeadCD.Equals("T"))
                {
                    // リネーム前後、すべて調査部の場合、リネームを実施する
                    isRename = 1;
                }
                else if (sJigyoubuHeadCD.Equals("T"))
                {
                    // リネーム後のみ調査部なら、新規作成する
                    isRename = 3;
                }
                else if (sJigyoubuHeadCD_ori.Equals("T"))
                {
                    // リネームボタン押下しない、契約部署のみ変更する場合
                    if (folderTo.Equals(sFolderRenameBef) == false)
                    {
                        isRename = 4;
                    }
                    else
                    {
                        txt_renamedfolder.Text = "";
                        item3_1_20_reset_ankenno.Text = "";
                        isRename = 5;
                    }
                }
                else
                {
                    if (folderTo.Contains(ori_ankenNo))
                    {
                        isRename = 4;
                    }
                    else
                    {
                        isRename = 5;
                    }
                }
            }
            else
            {
                isRename = 5;
            }

            //案件番号も変更する場合
            if (ori_ankenNo.Equals(item1_6.Text) == false && (isRename == 1 || isRename == 3 || isRename == 4))
            {
                folderTo = GlobalMethod.ChangeSqlText(folderTo.Replace("\\" + ori_ankenNo, "\\" + item1_6.Text), 0, 0);
            }

            // リネームを実行する
            if (isRename == 1 || isRename == 4)
            {
                bool isError = false;
                //調査部から調査部----------------------------------------------------
                //E10018	元フォルダが見つかりませんでした。確認して下さい。
                if (Directory.Exists(sFolderRenameBef) == false)
                {
                    isError = true;
                    set_error(GlobalMethod.GetMessage("E10018", "(引合)"));
                }

                //E10019 リネームするフォルダが既に存在します。確認して下さい。
                if (Directory.Exists(folderTo) == true)
                {
                    isError = true;
                    set_error(GlobalMethod.GetMessage("E10019", "(引合)"));
                }

                //E10020 リネームするフォルダ（支部のフォルダ）が見つかりませんでした。確認して下さい。
                int lstI = folderTo.LastIndexOf("\\");
                string sTo = folderTo;
                if (lstI >= 0)
                {
                    sTo = folderTo.Substring(0, lstI);
                }
                if (Directory.Exists(sTo) == false)
                {
                    isError = true;
                    set_error(GlobalMethod.GetMessage("E10020", "(引合)"));
                }
                // ファイルを移動処理
                if (isError == false)
                {
                    try
                    {
                        // DirectoryInfoのインスタンスを生成する
                        DirectoryInfo di = new DirectoryInfo(sFolderRenameBef);

                        // ディレクトリを移動する
                        di.MoveTo(folderTo);
                        isMoveOk = true;
                    }
                    catch (Exception ex)
                    {
                        // 移動失敗
                        set_error(GlobalMethod.GetMessage("E70065", "(引合)"));
                        GlobalMethod.outputLogger("UpdateEntory->FileMove", ex.Message, AnkenID, UserInfos[1]);
                    }
                }
            }
            else if (isRename == 2)
            {
                //調査部からほかの部門----------------------------------------------------
                isMoveOk = true;
            }
            else if (isRename == 3)
            {
                //ほかの部門から調査部----------------------------------------------------
                bool isCreateOk = false;
                // フォルダが存在しない場合、作成する
                if (!File.Exists(folderTo))
                {
                    try
                    {
                        DirectoryInfo di = new DirectoryInfo(folderTo);
                        di.Create();
                        isCreateOk = true;
                    }
                    catch (Exception)
                    {
                        // フォルダを作成する権限がありません。
                        set_error(GlobalMethod.GetMessage("E70046", "(引合)"));
                    }
                }
                else
                {
                    // 案件テーブルにすでに登録されているか
                    string where = "";
                    where = "AnkenKeiyakusho = N'" + GlobalMethod.ChangeSqlText(folderTo, 0, 0) + "'" +
                        " AND AnkenDeleteFlag = 0 AND AnkenAnkenBangou <> '" + ori_ankenNo + "' ";
                    DataTable dtFolder = GlobalMethod.getData("AnkenKeiyakusho", "AnkenKeiyakusho", "AnkenJouhou", where);
                    if (dtFolder is null || dtFolder.Rows.Count <= 0)
                    {
                        isCreateOk = true;
                    }
                }
                if (isCreateOk)
                {
                    DataTable CreateList = GlobalMethod.getData("CommonMasterID", "CommonValue1", "M_CommonMaster", "CommonMasterKye = 'ANKEN_BANGOU_FOLDER' ORDER BY CommonMasterID ");
                    if (CreateList != null && CreateList.Rows.Count > 0)
                    {
                        for (int i = 0; i < CreateList.Rows.Count; i++)
                        {

                            DirectoryInfo di = new DirectoryInfo(folderTo + "\\" + CreateList.Rows[i][0].ToString());
                            if (!Directory.Exists(folderTo + "\\" + CreateList.Rows[i][0].ToString()))
                            {
                                try
                                {
                                    di.Create();
                                }
                                catch (Exception)
                                {
                                    // フォルダを作成する権限がありません。
                                    set_error(GlobalMethod.GetMessage("E70046", "(引合)"));
                                    break;
                                }
                            }
                        }
                        // ここまできたらフォルダ作成が成功なので、
                        isMoveOk = true;
                    }
                }
            }
            else if (isRename == 5)
            {
                //何もしない
            }
            else
            {
                // フォルダ関連は工期開始年度で作成する
                string FolderBase = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", item1_2_KoukiNendo.SelectedValue.ToString());
                string FolderPath = FolderBase.Replace("/", @"\");

                // 変更後のパスは基本パスの場合、元パスを削除するか
                if (FolderPath.Length == sFolderRenameBef.Length)
                {
                    //元パスも基本バスなら、なにもしない
                }
                else
                {
                    string sBef = sFolderRenameBef;
                    int iBef = sBef.Length - sBef.Replace(@"\", "").Length;
                    sBef = FolderBase;
                    int iBase = sBef.Length - sBef.Replace(@"\", "").Length;
                    if (iBef <= iBase)
                    {
                        // 何もしない
                    }
                    else
                    {
                        isMoveOk = true;
                    }
                }
            }

            if (isMoveOk)
            {
                // 元フォルダがあるだったら、削除をする
                if (isRename != 3 && isRename != 4)
                {
                    try
                    {
                        // 元フォルダを削除する
                        if (Directory.Exists(sFolderRenameBef))
                        {
                            Directory.Delete(sFolderRenameBef, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        GlobalMethod.outputLogger("UpdateEntory->FileDelete", ex.Message, AnkenID, UserInfos[1]);
                    }
                }
                // No.1422 1196 案件番号の変更履歴を保存する
                //item1_37_kojinCD.Text = UserInfos[0];
                //item1_37.Text = UserInfos[1];
                //item1_38.Text = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                item1_12.Text = folderTo;
                txt_renamedfolder.Text = "";
                item3_1_20_reset_ankenno.Text = "";

                // 現在フォルダを更新
                sFolderRenameBef = item1_12.Text;
            }
            else
            {
                // No1420（差戻） もっとの更新機能が残す
                ////元フォルダへ戻す
                //item1_12.Text = sFolderRenameBef;
                txt_renamedfolder.Text = "";
                item3_1_20_reset_ankenno.Text = "";
            }
            return isMoveOk;
        }

        // えんとり君修正STEP2 確認シートのダミーデータ作成
        private int Create_DummyData()
        {
            int rtnAnkenNo = 0;
            string methodName = ".Create_DummyData";
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();

                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;

                try
                {
                    string SakuseiKubun = item3_1_1.SelectedValue.ToString();
                    // ダミーの赤伝のAnkenJouhouID取得
                    int ankenNo = GlobalMethod.getSaiban("AnkenJouhouID");

                    // 案件情報　INSERT　カラム共通
                    string sAkSqlCom = getAnkenJouhouInsertSQL();
                    // 案件情報の赤伝のダミーデータ作成 ----------------------------------------------------------------------------------
                    cmd.CommandText = sAkSqlCom + " SELECT " + ankenNo + getAnkenJouhouInsertVal(SakuseiKubun);
                    Console.WriteLine(cmd.CommandText);
                    var result = cmd.ExecuteNonQuery();

                    // 案件情報前回落札情報作成
                    cmd.CommandText = getAnkenJouhouZenkaiRakusatsuInsertSQL(ankenNo);
                    Console.WriteLine(cmd.CommandText);
                    result = cmd.ExecuteNonQuery();

                    //顧客契約情報存在チェック処理
                    if (!GlobalMethod.Check_Table(AnkenID, "KokyakuKeiyakuID", "KokyakuKeiyakuJouhou", ""))
                    {
                        GlobalMethod.outputLogger("ChangeKianEntry", "顧客契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                        transaction.Rollback();
                        conn.Close();
                        return rtnAnkenNo;
                    }
                    cmd.CommandText = getKokyakuKeiyakuJouhouInsertSQL(ankenNo);
                    Console.WriteLine(cmd.CommandText);
                    result = cmd.ExecuteNonQuery();

                    //業務情報存在チェック処理
                    if (!GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhou", ""))
                    {
                        GlobalMethod.outputLogger("ChangeKianEntry", "業務情報が見つからない", "ID:" + AnkenID, "DEBUG");
                        transaction.Rollback();
                        conn.Close();
                        return rtnAnkenNo;
                    }
                    cmd.CommandText = getGyoumuJouhouInsertSql(ankenNo);
                    Console.WriteLine(cmd.CommandText);
                    result = cmd.ExecuteNonQuery();

                    if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyouronTantouL1", ""))
                    {
                        cmd.CommandText = getGyoumuJouhouHyouronTantouL1InsertSQL(ankenNo);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();
                    }

                    if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouMadoguchi", ""))
                    {
                        // 窓口担当者
                        // 新では1件しか入らないが、現行が複数件はいるので、複数件あった場合でも落ちないようにする
                        exeGyoumuJouhouMadoguchi(ankenNo, cmd);
                    }

                    if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyoutenBusho", ""))
                    {
                        cmd.CommandText = getGyoumuJouhouHyoutenBushoInsertSql(ankenNo);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();
                    }

                    // 契約情報存在チェック
                    if (!GlobalMethod.Check_Table(AnkenID, "KeiyakuJouhouEntoryID", "KeiyakuJouhouEntory", ""))
                    {
                        GlobalMethod.outputLogger("ChangeKianEntry", "契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                        transaction.Rollback();
                        conn.Close();
                        return rtnAnkenNo;
                    }

                    cmd.CommandText = getKeiyakuJouhouEntoryInsertSql(ankenNo);
                    Console.WriteLine(cmd.CommandText);
                    result = cmd.ExecuteNonQuery();

                    if (GlobalMethod.Check_Table(AnkenID, "RibcID", "RibcJouhou", ""))
                    {
                        cmd.CommandText = getRibcJouhouInsertSql(ankenNo);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();
                    }

                    if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                    {
                        cmd.CommandText = getNyuusatsuJouhouInsertSql(ankenNo);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        GlobalMethod.outputLogger("ChangeKianEntry", "入札情報が見つからない", "ID:" + AnkenID, "DEBUG");
                        transaction.Rollback();
                        conn.Close();
                        return rtnAnkenNo;
                    }

                    if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhouOusatsusha", ""))
                    {
                        cmd.CommandText = getNyuusatsuJouhouOusatsushaInsertSql(ankenNo);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();
                    }

                    DataTable GH_dt = new DataTable();
                    GH_dt = GlobalMethod.getData("GyoumuHaibunID", "GyoumuAnkenJouhouID", "GyoumuHaibun", "GyoumuAnkenJouhouID = " + AnkenID);
                    if (GH_dt != null && GH_dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < GH_dt.Rows.Count; i++)
                        {
                            cmd.CommandText = getGyoumuHaibunInsertSql(ankenNo, GetInt(GH_dt.Rows[i][1].ToString()));
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }
                    }

                    // 案件情報の黒伝のダミーデータ作成 ----------------------------------------------------------------------------------
                    int ankenNo2 = 0;
                    if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                    {
                        // 黒伝のAnkenJouhouID
                        ankenNo2 = GlobalMethod.getSaiban("AnkenJouhouID");
                        //黒伝のダミーデータ作成
                        cmd.CommandText = sAkSqlCom + " SELECT " + ankenNo2 + getAnkenJouhouInsertVal(SakuseiKubun, 1);
                        result = cmd.ExecuteNonQuery();

                        // 案件情報前回落札情報作成
                        cmd.CommandText = getAnkenJouhouZenkaiRakusatsuInsertSQL(ankenNo2, 1);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        //顧客契約情報存在チェック処理
                        if (!GlobalMethod.Check_Table(AnkenID, "KokyakuKeiyakuID", "KokyakuKeiyakuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "顧客契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return rtnAnkenNo;
                        }
                        cmd.CommandText = getKokyakuKeiyakuJouhouInsertSQL(ankenNo2);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        //業務情報存在チェック処理
                        if (!GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "業務情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return rtnAnkenNo;
                        }
                        cmd.CommandText = getGyoumuJouhouInsertSql(ankenNo2);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyouronTantouL1", ""))
                        {
                            cmd.CommandText = getGyoumuJouhouHyouronTantouL1InsertSQL(ankenNo2);
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouMadoguchi", ""))
                        {
                            // 窓口担当者
                            // 新では1件しか入らないが、現行が複数件はいるので、複数件あった場合でも落ちないようにする
                            exeGyoumuJouhouMadoguchi(ankenNo2, cmd);
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyoutenBusho", ""))
                        {
                            cmd.CommandText = getGyoumuJouhouHyoutenBushoInsertSql(ankenNo2);
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }

                        // 契約情報存在チェック
                        if (!GlobalMethod.Check_Table(AnkenID, "KeiyakuJouhouEntoryID", "KeiyakuJouhouEntory", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return rtnAnkenNo;
                        }
                        cmd.CommandText = getKeiyakuJouhouEntoryInsertSql(ankenNo2, 1);
                        Console.WriteLine(cmd.CommandText);
                        result = cmd.ExecuteNonQuery();

                        if (GlobalMethod.Check_Table(AnkenID, "RibcID", "RibcJouhou", ""))
                        {
                            cmd.CommandText = getRibcJouhouInsertSql(ankenNo2, 1);
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                        {
                            cmd.CommandText = getNyuusatsuJouhouInsertSql(ankenNo2, 1);
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "入札情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return rtnAnkenNo;
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhouOusatsusha", ""))
                        {
                            cmd.CommandText = getNyuusatsuJouhouOusatsushaInsertSql(ankenNo2, 1);
                            Console.WriteLine(cmd.CommandText);
                            result = cmd.ExecuteNonQuery();
                        }

                        GH_dt = new DataTable();
                        GH_dt = GlobalMethod.getData("GyoumuHaibunID", "GyoumuAnkenJouhouID", "GyoumuHaibun", "GyoumuAnkenJouhouID = " + AnkenID);
                        if (GH_dt != null && GH_dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < GH_dt.Rows.Count; i++)
                            {
                                cmd.CommandText = getGyoumuHaibunInsertSql(ankenNo2, GetInt(GH_dt.Rows[i][1].ToString()), 1);
                                Console.WriteLine(cmd.CommandText);
                                result = cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    transaction.Commit();

                    transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    if (GlobalMethod.Check_Table(AnkenID, "KeiyakuJouhouEntoryID", "KeiyakuJouhouEntory", ""))
                    {
                        cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET " +
                                 "KeiyakuHenkouChuushiRiyuu = N'" + item3_1_17.Text + "' " +
                                ",KeiyakuSakuseibi = " + Get_DateTimePicker("item3_1_4");
                        if (SakuseiKubun == "02")
                        {
                            cmd.CommandText += ",KeiyakuKeiyakuTeiketsubi = " + Get_DateTimePicker("item3_1_3");
                        }
                        cmd.CommandText += " WHERE KeiyakuJouhouEntoryID = " + ankenNo;
                        result = cmd.ExecuteNonQuery();
                    }

                    if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                    {
                        cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                    "AnkenSakuseiKubun = N'" + item3_1_1.SelectedValue.ToString() + "' " +
                                    ",AnkenGyoumuKubun = N'" + item3_1_8.SelectedValue.ToString() + "' " +
                                    ",AnkenGyoumuKubunMei = N'" + item3_1_8.Text + "' " +
                                    ",AnkenGyoumuMei = N'" + item3_1_11.Text + "' " +
                                    ",AnkenKianzumi = '1' " +
                                    ",AnkenUriageNendo = N'" + item3_1_5.SelectedValue.ToString() + "' " +
                                    ",AnkenKeiyakuTeiketsubi = " + Get_DateTimePicker("item3_1_3") +
                                    ",AnkenKeiyakuKoukiKaishibi = " + Get_DateTimePicker("item3_1_6") +
                                    ",AnkenKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("item3_1_7") +
                                    ",AnkenKeiyakuZeikomiKingaku = " + item3_1_13.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuC = " + item3_2_1_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuJ = " + item3_2_2_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuJs = " + item3_2_3_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuUriageHaibunGakuK = " + item3_2_4_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",AnkenKeiyakuSakuseibi = " + Get_DateTimePicker("item3_1_4") +
                                    ",GyoumuKanrishaCD = " + "N'" + item3_4_4_CD.Text + "'" +
                                    ",GyoumuKanrishaMei = " + "N'" + item3_4_4.Text + "'" +
                                    ",AnkenUpdateUser = N'" + UserInfos[0] + "' " +
                                    " WHERE AnkenJouhou.AnkenJouhouID = " + ankenNo2;
                        //業務情報
                        cmd.CommandText = "UPDATE GyoumuJouhou SET " +
                                    "GyoumuHyouten = " + "N'" + item4_1_1.Text + "'" +
                                    ",KanriGijutsushaCD = " + "N'" + item3_4_1_CD.Text + "'" +
                                    ",KanriGijutsushaNM = " + "N'" + item3_4_1.Text + "'" +
                                    ",GyoumuKanriHyouten = " + "N'" + item3_4_1_Hyoten.Text + "'" +
                                    ",ShousaTantoushaCD = " + "N'" + item3_4_2_CD.Text + "'" +
                                    ",ShousaTantoushaNM = " + "N'" + item3_4_2.Text + "'" +
                                    ",GyoumuShousaHyouten = " + "N'" + item3_4_2_Hyoten.Text + "'" +
                                    ",SinsaTantoushaCD = " + "N'" + item3_4_3_CD.Text + "'" +
                                    ",SinsaTantoushaNM = " + "N'" + item3_4_3.Text + "'" +
                                    ",GyoumuUpdateDate = " + " GETDATE() " +
                                    ",GyoumuUpdateUser = " + "N'" + UserInfos[0] + "' " +
                                    ",GyoumuUpdateProgram = " + "'UpdateEntory' " +
                                    ",GyoumuDeleteFlag = " + "0 " +
                                    " WHERE AnkenJouhouID = " + ankenNo2;
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                        result = cmd.ExecuteNonQuery();

                        //業務情報技術担当者
                        cmd.CommandText = "DELETE GyoumuJouhouHyouronTantouL1 WHERE GyoumuJouhouID = '" + ankenNo2 + "' ";
                        cmd.ExecuteNonQuery();

                        for (int i = 1; i < c1FlexGrid5.Rows.Count; i++)
                        {
                            if (c1FlexGrid3.Rows[i][1] != null && c1FlexGrid3.Rows[i][1].ToString() != "")
                            {
                                string Hyouten = "";
                                if (c1FlexGrid3.Rows[i][3] != null && c1FlexGrid3.Rows[i][3].ToString() != "")
                                {
                                    Hyouten = c1FlexGrid3.Rows[i][3].ToString();
                                }
                                cmd.CommandText = "INSERT GyoumuJouhouHyouronTantouL1 ( " +
                                        "GyoumuJouhouID " +
                                        ", HyouronTantouID " +
                                        ", HyouronTantoushaCD " +
                                        ", HyouronTantoushaMei " +
                                        ", HyouronnTantoushaHyouten " +
                                        ") VALUES (" +
                                        "'" + ankenNo2 + "' " +
                                        "," + i +
                                        ",N'" + c1FlexGrid5.Rows[i][1].ToString() + "' " +
                                        ",N'" + c1FlexGrid5.Rows[i][2].ToString() + "' " +
                                        ",N'" + Hyouten + "' " +
                                        ") ";
                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        // 名称の取得
                        string GyoumuJouhouMadoShibuMei = "";
                        string GyoumuJouhouMadoKamei = "";
                        DataTable dt2 = new DataTable();
                        cmd.CommandText = "SELECT ShibuMei, KaMei FROM Mst_Busho WHERE GyoumuBushoCD = '" + item3_4_5_Busho.Text + "'";
                        var sda2 = new SqlDataAdapter(cmd);
                        dt2.Clear();
                        sda2.Fill(dt2);
                        if (dt2 != null && dt2.Rows.Count > 0)
                        {
                            GyoumuJouhouMadoShibuMei = dt2.Rows[0][0].ToString();
                            GyoumuJouhouMadoKamei = dt2.Rows[0][1].ToString();
                        }

                        //窓口担当者の更新
                        if ((item3_4_5_CD.Text == "0") || (item3_4_5_CD.Text == ""))
                        {
                            cmd.CommandText = "DELETE GyoumuJouhouMadoguchi WHERE GyoumuJouhouID = '" + ankenNo2 + "' ";
                            cmd.ExecuteNonQuery();
                        }
                        else
                        {
                            // 窓口担当者が複数いた場合の対応
                            DataTable gyoumuJouhouMadoguchiDT = new DataTable();
                            cmd.CommandText = "SELECT TOP 1 GyoumuJouhouMadoguchiID " +
                                            "FROM GyoumuJouhouMadoguchi " +
                                            "where GyoumuJouhouID = '" + ankenNo2 + "' " +
                                            "ORDER BY GyoumuJouhouMadoguchiID ";

                            string GyoumuJouhouMadoguchiID = "";

                            var gyoumuJouhouMadoguchiSda = new SqlDataAdapter(cmd);
                            gyoumuJouhouMadoguchiDT.Clear();
                            gyoumuJouhouMadoguchiSda.Fill(gyoumuJouhouMadoguchiDT);
                            if (gyoumuJouhouMadoguchiDT != null && gyoumuJouhouMadoguchiDT.Rows.Count > 0)
                            {
                                GyoumuJouhouMadoguchiID = gyoumuJouhouMadoguchiDT.Rows[0][0].ToString();
                            }
                            // データが存在する場合
                            if (GyoumuJouhouMadoguchiID != "")
                            {
                                cmd.CommandText = "UPDATE GyoumuJouhouMadoguchi set " +
                                                "GyoumuJouhouMadoGyoumuBushoCD = N'" + item3_4_5_Busho.Text + "' " +
                                                ",GyoumuJouhouMadoShibuMei = N'" + GyoumuJouhouMadoShibuMei + "' " +
                                                ",GyoumuJouhouMadoKamei = N'" + GyoumuJouhouMadoKamei + "' " +
                                                ",GyoumuJouhouMadoKojinCD = N'" + item3_4_5_CD.Text + "' " +
                                                ",GyoumuJouhouMadoChousainMei = N'" + item3_4_5.Text + "' " +
                                                "WHERE GyoumuJouhouMadoguchiID = N'" + GyoumuJouhouMadoguchiID + "' ";

                                cmd.ExecuteNonQuery();
                            }

                        }

                        cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET " +
                                    "KeiyakuKeiyakuTeiketsubi = " + Get_DateTimePicker("item3_1_3") +
                                    ",KeiyakuSakuseibi = " + Get_DateTimePicker("item3_1_4") +
                                    ",KeiyakuKeiyakuKoukiKaishibi = " + Get_DateTimePicker("item3_1_6") +
                                    ",KeiyakuKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("item3_1_7") +
                                    ",KeiyakuKeiyakuKingaku = " + item3_1_12.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuZeikomiKingaku = " + item3_1_13.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuuchizeiKingaku = " + item3_1_14.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuShouhizeiritsu = N'" + item3_1_10.Text + "'" +
                                    ",KeiyakuHenkouChuushiRiyuu = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_17.Text, 0, 0) + "'" +
                                    ",KeiyakuBikou = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_19.Text, 0, 0) + "'" +
                                    ",KeiyakuShosha = " + (item3_1_20.Checked ? 1 : 0) +
                                    ",KeiyakuTokkiShiyousho = " + (item3_1_21.Checked ? 1 : 0) +
                                    ",KeiyakuMitsumorisho = " + (item3_1_22.Checked ? 1 : 0) +
                                    ",KeiyakuTanpinChousaMitsumorisho = " + (item3_1_23.Checked ? 1 : 0) +
                                    ",KeiyakuSonota = " + (item3_1_24.Checked ? 1 : 0) +
                                    ",KeiyakuSonotaNaiyou = " + "N'" + GlobalMethod.ChangeSqlText(item3_1_25.Text, 0, 0) + "'" +
                                    ",KeiyakuZentokinUkewatashibi = " + Get_DateTimePicker("item3_6_11") +
                                    ",KeiyakuZentokin = " + item3_6_12.Text.Replace("¥", "").Replace(",", "") +
                                    ",Keiyakukeiyakukingakukei = " + item3_1_15.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuBetsuKeiyakuKingaku = " + item3_1_16.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi1 = " + " " + Get_DateTimePicker("item3_6_1") + "" +
                                    ",KeiyakuSeikyuuKingaku1 = " + item3_6_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi2 = " + " " + Get_DateTimePicker("item3_6_3") + "" +
                                    ",KeiyakuSeikyuuKingaku2 = " + item3_6_4.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi3 = " + " " + Get_DateTimePicker("item3_6_5") + "" +
                                    ",KeiyakuSeikyuuKingaku3 = " + item3_6_6.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi4 = " + " " + Get_DateTimePicker("item3_6_7") + "" +
                                    ",KeiyakuSeikyuuKingaku4 = " + item3_6_8.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSeikyuubi5 = " + " " + Get_DateTimePicker("item3_6_9") + "" +
                                    ",KeiyakuSeikyuuKingaku5 = " + item3_6_10.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuSakuseiKubunID = " + "N'" + item3_1_1.SelectedValue + "'" +
                                    ",KeiyakuSakuseiKubun = " + "N'" + item3_1_1.Text + "'" +
                                    ",KeiyakuGyoumuKubun = " + "N'" + item3_1_8.SelectedValue + "'" +
                                    ",KeiyakuGyoumuMei = " + "N'" + item3_1_8.Text + "'" +
                                    ",KeiyakuKianzumi = " + (item3_1_2.Checked ? 1 : 0) +
                                    ",KeiyakuHachuushaMei = " + "N'" + item3_1_9.Text + "'" +
                                    ",KeiyakuHaibunChoZeinuki = " + item3_2_1_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuHaibunJoZeinuki = " + item3_2_2_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuHaibunJosysZeinuki = " + item3_2_3_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuHaibunKeiZeinuki = " + item3_2_4_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuHaibunZeinukiKei = " + item3_2_5_2.Text.Replace("¥", "").Replace(",", "").Replace("%", "") +
                                    ",KeiyakuUriageHaibunCho  = " + item3_2_1_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUriageHaibunJo   = " + item3_2_2_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUriageHaibunJosys  = " + item3_2_3_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUriageHaibunKei  = " + item3_2_4_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUriageHaibunGoukei = " + item3_2_5_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuTankeiMikomiCho  = " + item3_3_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuTankeiMikomiJo  = " + item3_3_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuTankeiMikomiJosys  = " + item3_3_3.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuTankeiMikomiKei  = " + item3_3_4.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuKurikoshiCho  = " + item3_7_1.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuKurikoshiJo  = " + item3_7_2.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuKurikoshiJosys  = " + item3_7_3.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuKurikoshiKei  = " + item3_7_4.Text.Replace("¥", "").Replace(",", "") +
                                    ",KeiyakuUpdateProgram = " + "'ChangeKianEntry'" +
                                    ",KeiyakuUpdateDate = " + "GETDATE()" +
                                    ",KeiyakuUpdateUser = " + "N'" + UserInfos[0] + "'" +
                                    // えんとり君修正STEP2（RIBC項目追加）
                                    ",KeiyakuRIBCYouTankaDataMoushikomisho = " + (item3_ribc_price.Checked ? 1 : 0) +
                                    ",KeiyakuSashaKeiyu = " + (item3_sa_commpany.Checked ? 1 : 0) +
                                    ",KeiyakuRIBCYouTankaData = " + (item3_1_ribc.Checked ? 1 : 0) +
                                    " WHERE AnkenJouhouID = " + ankenNo2;
                        result = cmd.ExecuteNonQuery();

                        cmd.CommandText = "DELETE FROM RibcJouhou " +
                                " WHERE RibcID = " + ankenNo2;
                        cmd.ExecuteNonQuery();
                        int cnt = 0;
                        string RibcKoukiStart;
                        string RibcNouhinbi;
                        string RibcSeikyubi;
                        string RibcNyukinyoteibi;
                        string RibcKubun;
                        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                        {

                            // 新では計上額のみでも登録を可とする
                            // 計上日、計上月、計上額のどれかが入っていれば登録する
                            // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                            if ((c1FlexGrid4.Rows[i][1] != null && c1FlexGrid4.Rows[i][1].ToString() != "")
                                || (c1FlexGrid4.Rows[i][2] != null && c1FlexGrid4.Rows[i][2].ToString() != "")
                                || (c1FlexGrid4.Rows[i][3] != null && c1FlexGrid4.Rows[i][3].ToString() != "0"))
                            {
                                RibcKoukiStart = "null";
                                if (c1FlexGrid4.Rows[i][4] != null)
                                {
                                    RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][4].ToString() + "'";
                                }
                                RibcNouhinbi = "null";
                                if (c1FlexGrid4.Rows[i][5] != null)
                                {
                                    RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][5].ToString() + "'";
                                }
                                RibcSeikyubi = "null";
                                if (c1FlexGrid4.Rows[i][6] != null)
                                {
                                    RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][6].ToString() + "'";
                                }
                                RibcNyukinyoteibi = "null";
                                if (c1FlexGrid4.Rows[i][7] != null)
                                {
                                    RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][7].ToString() + "'";
                                }
                                RibcKubun = "0";
                                if (c1FlexGrid4.Rows[i][8] != null)
                                {
                                    RibcKubun = "N'" + c1FlexGrid4.Rows[i][8].ToString() + "'";
                                }

                                cnt++;
                                cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                            "RibcID " +
                                            ",RibcNo " +
                                            ",RibcKoukiEnd " +
                                            ",RibcUriageKeijyoTuki " +
                                            ",RibcSeikyuKingaku " +
                                            ",RibcKankeibusho " +
                                            ",RibcKoukiStart " +
                                            ",RibcNouhinbi " +
                                            ",RibcSeikyubi " +
                                            ",RibcNyukinyoteibi " +
                                            ",RibcKubun " +
                                            ") VALUES (" +
                                           ankenNo2 +
                                            "," + cnt + "";
                                if (c1FlexGrid4.Rows[i][1] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][1].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",null";
                                }
                                // 計上月 RibcUriageKeijyoTuki
                                if (c1FlexGrid4.Rows[i][2] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][2].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'' " + "";
                                }
                                // 計上額 RibcSeikyuKingaku
                                if (c1FlexGrid4.Rows[i][3] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][3].ToString().Replace("¥", "").Replace(",", "") + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                }
                                cmd.CommandText = cmd.CommandText +
                                ",'127120' " +
                                            "," + RibcKoukiStart +
                                            "," + RibcNouhinbi +
                                            "," + RibcSeikyubi +
                                            "," + RibcNyukinyoteibi +
                                            "," + RibcKubun +
                                            ")";

                                Console.WriteLine(cmd.CommandText);
                                cmd.ExecuteNonQuery();
                            }

                            // 新では計上額のみでも登録を可とする
                            // 計上日、計上月、計上額のどれかが入っていれば登録する
                            // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                            if ((c1FlexGrid4.Rows[i][9] != null && c1FlexGrid4.Rows[i][9].ToString() != "")
                                || (c1FlexGrid4.Rows[i][10] != null && c1FlexGrid4.Rows[i][10].ToString() != "")
                                || (c1FlexGrid4.Rows[i][11] != null && c1FlexGrid4.Rows[i][11].ToString() != "0"))
                            {
                                RibcKoukiStart = "null";
                                if (c1FlexGrid4.Rows[i][12] != null)
                                {
                                    RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][12].ToString() + "'";
                                }
                                RibcNouhinbi = "null";
                                if (c1FlexGrid4.Rows[i][13] != null)
                                {
                                    RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][13].ToString() + "'";
                                }
                                RibcSeikyubi = "null";
                                if (c1FlexGrid4.Rows[i][14] != null)
                                {
                                    RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][14].ToString() + "'";
                                }
                                RibcNyukinyoteibi = "null";
                                if (c1FlexGrid4.Rows[i][15] != null)
                                {
                                    RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][15].ToString() + "'";
                                }
                                RibcKubun = "0";
                                if (c1FlexGrid4.Rows[i][16] != null)
                                {
                                    RibcKubun = "N'" + c1FlexGrid4.Rows[i][16].ToString() + "'";
                                }
                                cnt++;
                                cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                            "RibcID " +
                                            ",RibcNo " +
                                            ",RibcKoukiEnd " +
                                            ",RibcUriageKeijyoTuki " +
                                            ",RibcSeikyuKingaku " +
                                            ",RibcKankeibusho " +
                                            ",RibcKoukiStart " +
                                            ",RibcNouhinbi " +
                                            ",RibcSeikyubi " +
                                            ",RibcNyukinyoteibi " +
                                            ",RibcKubun " +
                                            ") VALUES (" +
                                           ankenNo2 +
                                            "," + cnt + "";

                                if (c1FlexGrid4.Rows[i][9] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][9].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",null";
                                }
                                // 計上月 RibcUriageKeijyoTuki
                                if (c1FlexGrid4.Rows[i][10] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][10].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'' " + "";
                                }
                                // 計上額 RibcSeikyuKingaku
                                if (c1FlexGrid4.Rows[i][11] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][11].ToString().Replace("¥", "").Replace(",", "") + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                }
                                cmd.CommandText = cmd.CommandText + ",'129230' " +
                                            "," + RibcKoukiStart +
                                            "," + RibcNouhinbi +
                                            "," + RibcSeikyubi +
                                            "," + RibcNyukinyoteibi +
                                            "," + RibcKubun +
                                            ")";

                                cmd.ExecuteNonQuery();
                            }
                            //}

                            // 新では計上額のみでも登録を可とする
                            // 計上日、計上月、計上額のどれかが入っていれば登録する
                            // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                            if ((c1FlexGrid4.Rows[i][17] != null && c1FlexGrid4.Rows[i][17].ToString() != "")
                                || (c1FlexGrid4.Rows[i][18] != null && c1FlexGrid4.Rows[i][18].ToString() != "")
                                || (c1FlexGrid4.Rows[i][19] != null && c1FlexGrid4.Rows[i][19].ToString() != "0"))
                            {
                                RibcKoukiStart = "null";
                                if (c1FlexGrid4.Rows[i][20] != null)
                                {
                                    RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][20].ToString() + "'";
                                }
                                RibcNouhinbi = "null";
                                if (c1FlexGrid4.Rows[i][21] != null)
                                {
                                    RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][21].ToString() + "'";
                                }
                                RibcSeikyubi = "null";
                                if (c1FlexGrid4.Rows[i][22] != null)
                                {
                                    RibcSeikyubi = "N'" + c1FlexGrid4.Rows[i][22].ToString() + "'";
                                }
                                RibcNyukinyoteibi = "null";
                                if (c1FlexGrid4.Rows[i][23] != null)
                                {
                                    RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][23].ToString() + "'";
                                }
                                RibcKubun = "0";
                                if (c1FlexGrid4.Rows[i][24] != null)
                                {
                                    RibcKubun = "N'" + c1FlexGrid4.Rows[i][24].ToString() + "'";
                                }
                                cnt++;
                                cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                            "RibcID " +
                                            ",RibcNo " +
                                            ",RibcKoukiEnd " +
                                            ",RibcUriageKeijyoTuki " +
                                            ",RibcSeikyuKingaku " +
                                            ",RibcKankeibusho " +
                                            ",RibcKoukiStart " +
                                            ",RibcNouhinbi " +
                                            ",RibcSeikyubi " +
                                            ",RibcNyukinyoteibi " +
                                            ",RibcKubun " +
                                            ") VALUES (" +
                                           ankenNo2 +
                                            "," + cnt + "";

                                if (c1FlexGrid4.Rows[i][17] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][17].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",null";
                                }
                                // 計上月 RibcUriageKeijyoTuki
                                if (c1FlexGrid4.Rows[i][18] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][18].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'' " + "";
                                }
                                // 計上額 RibcSeikyuKingaku
                                if (c1FlexGrid4.Rows[i][19] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][19].ToString().Replace("¥", "").Replace(",", "") + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                }
                                // 年度により、情報システム部の部コードを変更する
                                if (GetInt(item3_1_5.SelectedValue.ToString()) >= 2021)
                                {
                                    cmd.CommandText = cmd.CommandText + ",'128400' " +
                                                "," + RibcKoukiStart +
                                                "," + RibcNouhinbi +
                                                "," + RibcSeikyubi +
                                                "," + RibcNyukinyoteibi +
                                                "," + RibcKubun +
                                                ")";
                                }
                                else
                                {
                                    // 2021年度以前は127900
                                    cmd.CommandText = cmd.CommandText + ",'127900' " +
                                                "," + RibcKoukiStart +
                                                "," + RibcNouhinbi +
                                                "," + RibcSeikyubi +
                                                "," + RibcNyukinyoteibi +
                                                "," + RibcKubun +
                                                ")";
                                }

                                cmd.ExecuteNonQuery();
                            }

                            // 新では計上額のみでも登録を可とする
                            // 計上日、計上月、計上額のどれかが入っていれば登録する
                            // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                            if ((c1FlexGrid4.Rows[i][25] != null && c1FlexGrid4.Rows[i][25].ToString() != "")
                                || (c1FlexGrid4.Rows[i][26] != null && c1FlexGrid4.Rows[i][26].ToString() != "")
                                || (c1FlexGrid4.Rows[i][27] != null && c1FlexGrid4.Rows[i][27].ToString() != "0"))
                            {
                                RibcKoukiStart = "null";
                                if (c1FlexGrid4.Rows[i][28] != null)
                                {
                                    RibcKoukiStart = "N'" + c1FlexGrid4.Rows[i][28].ToString() + "'";
                                }
                                RibcNouhinbi = "null";
                                if (c1FlexGrid4.Rows[i][29] != null)
                                {
                                    RibcNouhinbi = "N'" + c1FlexGrid4.Rows[i][29].ToString() + "'";
                                }
                                RibcSeikyubi = "null";
                                if (c1FlexGrid4.Rows[i][30] != null)
                                {
                                    RibcSeikyubi = "'" + c1FlexGrid4.Rows[i][30].ToString() + "'";
                                }
                                RibcNyukinyoteibi = "null";
                                if (c1FlexGrid4.Rows[i][31] != null)
                                {
                                    RibcNyukinyoteibi = "N'" + c1FlexGrid4.Rows[i][31].ToString() + "'";
                                }
                                RibcKubun = "0";
                                if (c1FlexGrid4.Rows[i][32] != null)
                                {
                                    RibcKubun = "N'" + c1FlexGrid4.Rows[i][32].ToString() + "'";
                                }
                                cnt++;
                                cmd.CommandText = "INSERT INTO RibcJouhou (" +
                                            "RibcID " +
                                            ",RibcNo " +
                                            ",RibcKoukiEnd " +
                                            ",RibcUriageKeijyoTuki " +
                                            ",RibcSeikyuKingaku " +
                                            ",RibcKankeibusho " +
                                            ",RibcKoukiStart " +
                                            ",RibcNouhinbi " +
                                            ",RibcSeikyubi " +
                                            ",RibcNyukinyoteibi " +
                                            ",RibcKubun " +
                                            ") VALUES (" +
                                           ankenNo2 +
                                            "," + cnt + "";
                                if (c1FlexGrid4.Rows[i][25] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][25].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",null";
                                }
                                // 計上月 RibcUriageKeijyoTuki
                                if (c1FlexGrid4.Rows[i][26] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][26].ToString() + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'' " + "";
                                }
                                // 計上額 RibcSeikyuKingaku
                                if (c1FlexGrid4.Rows[i][27] != null)
                                {
                                    cmd.CommandText = cmd.CommandText + ",N'" + c1FlexGrid4.Rows[i][27].ToString().Replace("¥", "").Replace(",", "") + "'";
                                }
                                else
                                {
                                    cmd.CommandText = cmd.CommandText + ",'0' " + "";
                                }
                                cmd.CommandText = cmd.CommandText + ",'150200' " +
                                            "," + RibcKoukiStart +
                                            "," + RibcNouhinbi +
                                            "," + RibcSeikyubi +
                                            "," + RibcNyukinyoteibi +
                                            "," + RibcKubun +
                                            ")";

                                cmd.ExecuteNonQuery();
                            }
                        }

                        //業務配分登録
                        cmd.CommandText = "UPDATE GyoumuHaibun SET " +
                                            " GyoumuChosaBuGaku " + " = N'" + item3_7_1_6_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuJigyoFukyuBuGaku " + " = N'" + item3_7_1_7_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuJyohouSystemBuGaku " + " = N'" + item3_7_1_8_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuSougouKenkyuJoGaku " + " = N'" + item3_7_1_9_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuShizaiChousaRitsu " + " = N'" + item3_7_2_14_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuEizenRitsu " + " = N'" + item3_7_2_15_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuKikiruiChousaRitsu " + " = N'" + item3_7_2_16_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuKoujiChousahiRitsu " + " = N'" + item3_7_2_17_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuSanpaiFukusanbutsuRitsu " + " = N'" + item3_7_2_18_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuHokakeChousaRitsu " + " = N'" + item3_7_2_19_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuShokeihiChousaRitsu " + " = N'" + item3_7_2_20_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuGenkaBunsekiRitsu " + " = N'" + item3_7_2_21_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuKijunsakuseiRitsu " + " = N'" + item3_7_2_22_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuKoukyouRoumuhiRitsu " + " = N'" + item3_7_2_23_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuRoumuhiKoukyouigaiRitsu " + " = N'" + item3_7_2_24_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuSonotaChousabuRitsu " + " = N'" + item3_7_2_25_1.Text.Replace("%", "") + "' " +
                                            ",GyoumuShizaiChousaGaku " + " = N'" + item3_7_2_14_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuEizenGaku " + " = N'" + item3_7_2_15_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuKikiruiChousaGaku " + " = N'" + item3_7_2_16_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuKoujiChousahiGaku " + " = N'" + item3_7_2_17_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuSanpaiFukusanbutsuGaku " + " = N'" + item3_7_2_18_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuHokakeChousaGaku " + " = N'" + item3_7_2_19_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuShokeihiChousaGaku " + " = N'" + item3_7_2_20_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuGenkaBunsekiGaku " + " = N'" + item3_7_2_21_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuKijunsakuseiGaku " + " = N'" + item3_7_2_22_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuKoukyouRoumuhiGaku " + " = N'" + item3_7_2_23_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuRoumuhiKoukyouigaiGaku " + " = N'" + item3_7_2_24_2.Text.Replace("%", "") + "' " +
                                            ",GyoumuSonotaChousabuGaku " + " = N'" + item3_7_2_25_2.Text.Replace("%", "") + "' " +
                                            " WHERE GyoumuAnkenJouhouID = " + ankenNo2 + " AND GyoumuHibunKubun = 30 ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }

                    transaction.Commit();
                    
                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "確認シートのダミー赤伝を作成しました ID:" + ankenNo, pgmName + methodName, "");
                    if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                    {
                        rtnAnkenNo = ankenNo2;
                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "確認シートのダミー黒伝を作成しました ID:" + ankenNo2, pgmName + methodName, "");
                    }
                    else
                    {
                        rtnAnkenNo = ankenNo;
                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "確認シートのダミー中止伝票を追加しました。 ID:" + AnkenID, pgmName + methodName, "");
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
            return rtnAnkenNo;
        }

        /// <summary>
        /// えんとり君修正STEP2 確認シート：ダミーデータ
        /// </summary>
        /// <returns></returns>
        private string getAnkenJouhouInsertSQL()
        {
            return "INSERT INTO AnkenJouhou ( " +
                            "AnkenJouhouID " +
                            ",AnkenSakuseiKubun " +
                            ",AnkenSaishinFlg " +
                            ",AnkenKishuKeikakugaku " +
                            ",AnkenKishuKeikakakugakuJf " +
                            ",AnkenKishuKeikakugakuJ " +
                            ",AnkenKeikakuZangaku " +
                            ",AnkenkeikakuZangakuJF " +
                            ",AnkenkeikakuZangakuJ " +
                            ",AnkenChokusetsuGenka " +
                            ",AnkenChokusetsuGenkaRitsu " +
                            ",AnkenGaichuuhi " +
                            ",AnkenJoukanDoboku " +
                            ",AnkenJoukanFukugou " +
                            ",AnkenJoukanGesuidou " +
                            ",AnkenJoukanHyoujun " +
                            ",AnkenJoukanIchiba " +
                            ",AnkenJoukanItiji " +
                            ",AnkenJoukanJutakuSonota " +
                            ",AnkenJoukanKentiku " +
                            ",AnkenJoukanKijunsho " +
                            ",AnkenJoukanKouwan " +
                            ",AnkenJoukanKuukou " +
                            ",AnkenJoukanSetsubi " +
                            ",AnkenJoukanSonota " +
                            ",AnkenJoukanSuidou " +
                            ",AnkenKeichoukaiKounyuuhi " +
                            ",AnkenKishuKeikakugakuK " +
                            ",AnkenKaisuu " +
                            ",AnkenCreateDate " +
                            ",AnkenCreateUser " +
                            ",AnkenCreateProgram " +
                            ",AnkenUpdateDate " +
                            ",AnkenUpdateUser " +
                            ",AnkenUpdateProgram " +
                            ",AnkenTourokubi " +
                            ",AnkenGyoumuMei " +
                            ",AnkenDeleteFlag " +
                            ",AnkenUriageNendo " +
                            ",AnkenHachushaKubunCD " +
                            ",AnkenHachushaKubunMei " +
                            ",AnkenHachuushaCodeID " +
                            ",AnkenHachuushaMei " +
                            ",AnkenGyoumuKubun " +
                            ",AnkenGyoumuKubunMei " +
                            ",AnkenNyuusatsuHoushiki " +
                            ",AnkenKyougouTasha " +
                            ",AnkenJutakubushoCD " +
                            ",AnkenJutakushibu " +
                            ",AnkenTantoushaCD " +
                            ",AnkenMadoguchiTantoushaCD " +
                            ",AnkenGyoumuKanrishaCD " +
                            ",AnkenGyoumuKanrisha " +
                            ",GyoumuKanrishaCD " +
                            ",AnkenHachuushaBusho " +
                            ",AnkenkeikakuZangakuK " +
                            ",AnkenJutakuBangou " +
                            ",AnkenJutakuBangouEda " +
                            ",AnkenNyuusatsuYoteibi " +
                            ",AnkenRakusatsusha " +
                            ",AnkenRakusatsuJouhou " +
                            ",AnkenKianZumi " +
                            ",AnkenKiangetsu " +
                            ",AnkenHanteiKubun " +
                            ",AnkenJoukanData " +
                            ",AnkenJoukanHachuuKikanCD " +
                            ",AnkenNyuukinKakuninbi " +
                            ",AnkenKanryouSakuseibi " +
                            ",AnkenHonbuKakuninbi " +
                            ",AnkenShizaiChousa " +
                            ",AnkenKoujiChousahi " +
                            ",AnkenKikiruiChousa " +
                            ",AnkenSanpaiFukusanbutsu " +
                            ",AnkenHokakeChousa " +
                            ",AnkenShokeihiChousa " +
                            ",AnkenGenkaBunseki " +
                            ",AnkenKijunsakusei " +
                            ",AnkenKoukyouRoumuhi " +
                            ",AnkenRoumuhiKoukyouigai " +
                            ",AnkenSonotaChousabu " +
                            ",AnkenOrdermadeJifubu " +
                            ",AnkenRIBCJifubu " +
                            ",AnkenSonotaJifubu " +
                            ",AnkenOrdermade " +
                            ",AnkenJouhouKaihatsu " +
                            ",AnkenRIBCJouhouKaihatsu " +
                            ",AnkenSoukenbu " +
                            ",AnkenSonotaJoujibu " +
                            ",AnkenTeikiTokuchou " +
                            ",AnkenTanpinTokuchou " +
                            ",AnkenKikiChousa " +
                            ",AnkenHachuushaIraibusho " +
                            ",AnkenHachuushaTantousha " +
                            ",AnkenHachuushaTEL " +
                            ",AnkenHachuushaFAX " +
                            ",AnkenHachuushaMail " +
                            ",AnkenHachuushaIraiYuubin " +
                            ",AnkenHachuushaIraiJuusho " +
                            ",AnkenHachuushaKeiyakuBusho " +
                            ",AnkenHachuushaKeiyakuTantou " +
                            ",AnkenHachuushaKeiyakuTEL " +
                            ",AnkenHachuushaKeiyakuFAX " +
                            ",AnkenHachuushaKeiyakuMail " +
                            ",AnkenHachuushaKeiyakuYuubin " +
                            ",AnkenHachuushaKeiyakuJuusho " +
                            ",AnkenHachuuDaihyouYakushoku " +
                            ",AnkenHachuuDaihyousha " +
                            ",AnkenRosenKawamei " +
                            ",AnkenGyoumuItakuKasho " +
                            ",AnkenJititaiKibunID " +
                            ",AnkenJititaiKubun " +
                            ",AnkenKeiyakuToshoNo " +
                            ",AnkenKirokuToshoNo " +
                            ",AnkenKirokuHokanNo " +
                            ",AnkenCDHokan " +
                            ",AnkenSeikaButsuHokanFile " +
                            ",AnkenSeikabutsuHokanbako " +
                            ",AnkenKokyakuHyoukaComment " +
                            ",AnkenToukaiHyoukaComment " +
                            ",AnkenKenCD " +
                            ",AnkenToshiCD " +
                            ",AnkenKeiyakusho " +
                            ",AnkenEizen " +
                            ",AnkenTantoushaMei " +
                            ",GyoumuKanrishaMei " +
                            ",AnkenGyoumuKubunCD " +
                            ",AnkenHachuushaKaMei " +
                            ",AnkenKeiyakuKoukiKaishibi " +
                            ",AnkenKeiyakuKoukiKanryoubi " +
                            ",AnkenKeiyakuTeiketsubi " +
                            ",AnkenKeiyakuZeikomiKingaku " +
                            ",AnkenKeiyakuUriageHaibunGakuC " +
                            ",AnkenKeiyakuUriageHaibunGakuJ " +
                            ",AnkenKeiyakuUriageHaibunGakuJs " +
                            ",AnkenKeiyakuUriageHaibunGakuK " +
                            ",AnkenKeiyakuUriageHaibunGakuR " +
                            ",AnkenKeiyakuSakuseibi " +
                            ",AnkenAnkenBangou " +
                            ",AnkenKeikakuBangou " +
                            ",AnkenHikiaijhokyo " +
                            ",AnkenKeikakuAnkenMei " +
                            ",AnkenToukaiSankouMitsumori " +
                            ",AnkenToukaiJyutyuIyoku " +
                            ",AnkenToukaiSankouMitsumoriGaku " +
                            ",AnkenHachushaKaMei " +
                            ",AnkenHachushaCD " +
                            ",AnkenToukaiOusatu " +
                            ",AnkenKoukiNendo) ";
        }

        /// <summary>
        /// えんとり君修正STEP2 確認シート：ダミーデータ
        /// </summary>
        /// <param name="SakuseiKubun">案件区分</param>
        /// <param name="flag">0:赤伝、1:黒伝</param>
        /// <returns></returns>
        private string getAnkenJouhouInsertVal(string SakuseiKubun, int flag = 0)
        {
            StringBuilder sb = new StringBuilder();

            if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
            {
                if (flag == 0)
                {
                    sb.Append(",'02' ");    //",AnkenSakuseiKubun "
                    sb.Append(",0 ");       //",AnkenSaishinFlg "
                    sb.Append(",- AnkenKishuKeikakugaku ");     //",AnkenKishuKeikakugaku "
                    sb.Append(",- AnkenKishuKeikakakugakuJf "); //",AnkenKishuKeikakakugakuJf "
                    sb.Append(",- AnkenKishuKeikakugakuJ ");    //",AnkenKishuKeikakugakuJ "
                }
                else
                {
                    sb.Append(",'").Append(SakuseiKubun).Append("' ");  //",AnkenSakuseiKubun "
                    sb.Append(",0 ");       //",AnkenSaishinFlg "
                    sb.Append(",AnkenKishuKeikakugaku ");     //",AnkenKishuKeikakugaku "
                    sb.Append(",AnkenKishuKeikakakugakuJf "); //",AnkenKishuKeikakakugakuJf "
                    sb.Append(",AnkenKishuKeikakugakuJ ");    //",AnkenKishuKeikakugakuJ "
                }
            }
            else
            {
                sb.Append(",'04' ");    //",AnkenSakuseiKubun "
                sb.Append(",0 ");       //",AnkenSaishinFlg "
                sb.Append(",0 ");       //",AnkenKishuKeikakugaku "
                sb.Append(",0 ");       //",AnkenKishuKeikakakugakuJf
                sb.Append(",0 ");       //",AnkenKishuKeikakugakuJ "
            }
            if (flag == 0) { 
                sb.Append(",- AnkenKeikakuZangaku ");
                sb.Append(",- AnkenkeikakuZangakuJF ");
                sb.Append(",- AnkenkeikakuZangakuJ ");
                sb.Append(",- AnkenChokusetsuGenka ");
                sb.Append(",- AnkenChokusetsuGenkaRitsu ");
                sb.Append(",- AnkenGaichuuhi ");
                sb.Append(",- AnkenJoukanDoboku ");
                sb.Append(",- AnkenJoukanFukugou ");
                sb.Append(",- AnkenJoukanGesuidou ");
                sb.Append(",- AnkenJoukanHyoujun ");
                sb.Append(",- AnkenJoukanIchiba ");
                sb.Append(",- AnkenJoukanItiji ");
                sb.Append(",- AnkenJoukanJutakuSonota ");
                sb.Append(",- AnkenJoukanKentiku ");
                sb.Append(",- AnkenJoukanKijunsho ");
                sb.Append(",- AnkenJoukanKouwan ");
                sb.Append(",- AnkenJoukanKuukou ");
                sb.Append(",- AnkenJoukanSetsubi ");
                sb.Append(",- AnkenJoukanSonota ");
                sb.Append(",- AnkenJoukanSuidou ");
                sb.Append(",- AnkenKeichoukaiKounyuuhi ");
                sb.Append(",- AnkenKishuKeikakugakuK ");
            }
            else
            {
                sb.Append(",AnkenKeikakuZangaku ");
                sb.Append(",AnkenkeikakuZangakuJF ");
                sb.Append(",AnkenkeikakuZangakuJ ");
                sb.Append(",AnkenChokusetsuGenka ");
                sb.Append(",AnkenChokusetsuGenkaRitsu ");
                sb.Append(",AnkenGaichuuhi ");
                sb.Append(",AnkenJoukanDoboku ");
                sb.Append(",AnkenJoukanFukugou ");
                sb.Append(",AnkenJoukanGesuidou ");
                sb.Append(",AnkenJoukanHyoujun ");
                sb.Append(",AnkenJoukanIchiba ");
                sb.Append(",AnkenJoukanItiji ");
                sb.Append(",AnkenJoukanJutakuSonota ");
                sb.Append(",AnkenJoukanKentiku ");
                sb.Append(",AnkenJoukanKijunsho ");
                sb.Append(",AnkenJoukanKouwan ");
                sb.Append(",AnkenJoukanKuukou ");
                sb.Append(",AnkenJoukanSetsubi ");
                sb.Append(",AnkenJoukanSonota ");
                sb.Append(",AnkenJoukanSuidou ");
                sb.Append(",AnkenKeichoukaiKounyuuhi ");
                sb.Append(",AnkenKishuKeikakugakuK ");
            }
            sb.Append(",AnkenKaisuu + 1 ");
            sb.Append(",GETDATE() ");
            sb.Append(",'").Append(UserInfos[0]).Append("' ");
            sb.Append(",'ChangeKianEntry' ");
            sb.Append(",GETDATE() ");
            sb.Append(",'").Append(UserInfos[0]).Append("' ");
            sb.Append(",'ChangeKianEntry' ");
            sb.Append(",AnkenTourokubi ");
            sb.Append(",AnkenGyoumuMei ");
            sb.Append(",1 ");
            sb.Append(",AnkenUriageNendo ");
            sb.Append(",AnkenHachushaKubunCD ");
            sb.Append(",AnkenHachushaKubunMei ");
            sb.Append(",AnkenHachuushaCodeID ");
            sb.Append(",AnkenHachuushaMei ");
            sb.Append(",AnkenGyoumuKubun ");
            sb.Append(",AnkenGyoumuKubunMei ");
            sb.Append(",AnkenNyuusatsuHoushiki ");
            sb.Append(",AnkenKyougouTasha ");
            sb.Append(",AnkenJutakubushoCD ");
            sb.Append(",AnkenJutakushibu ");
            sb.Append(",AnkenTantoushaCD ");
            sb.Append(",AnkenMadoguchiTantoushaCD ");
            sb.Append(",AnkenGyoumuKanrishaCD ");
            sb.Append(",AnkenGyoumuKanrisha ");
            sb.Append(",GyoumuKanrishaCD ");
            sb.Append(",AnkenHachuushaBusho ");
            sb.Append(",AnkenkeikakuZangakuK ");
            sb.Append(",AnkenJutakuBangou ");
            sb.Append(",AnkenJutakuBangouEda ");
            sb.Append(",AnkenNyuusatsuYoteibi ");
            sb.Append(",AnkenRakusatsusha ");
            sb.Append(",AnkenRakusatsuJouhou ");
            sb.Append(",AnkenKianZumi ");
            sb.Append(",AnkenKiangetsu ");
            sb.Append(",AnkenHanteiKubun ");
            sb.Append(",AnkenJoukanData ");
            sb.Append(",AnkenJoukanHachuuKikanCD ");
            sb.Append(",AnkenNyuukinKakuninbi ");
            sb.Append(",AnkenKanryouSakuseibi ");
            sb.Append(",AnkenHonbuKakuninbi ");
            sb.Append(",AnkenShizaiChousa ");
            sb.Append(",AnkenKoujiChousahi ");
            sb.Append(",AnkenKikiruiChousa ");
            sb.Append(",AnkenSanpaiFukusanbutsu ");
            sb.Append(",AnkenHokakeChousa ");
            sb.Append(",AnkenShokeihiChousa ");
            sb.Append(",AnkenGenkaBunseki ");
            sb.Append(",AnkenKijunsakusei ");
            sb.Append(",AnkenKoukyouRoumuhi ");
            sb.Append(",AnkenRoumuhiKoukyouigai ");
            sb.Append(",AnkenSonotaChousabu ");
            sb.Append(",AnkenOrdermadeJifubu ");
            sb.Append(",AnkenRIBCJifubu ");
            sb.Append(",AnkenSonotaJifubu ");
            sb.Append(",AnkenOrdermade ");
            sb.Append(",AnkenJouhouKaihatsu ");
            sb.Append(",AnkenRIBCJouhouKaihatsu ");
            sb.Append(",AnkenSoukenbu ");
            sb.Append(",AnkenSonotaJoujibu ");
            sb.Append(",AnkenTeikiTokuchou ");
            sb.Append(",AnkenTanpinTokuchou ");
            sb.Append(",AnkenKikiChousa ");
            sb.Append(",AnkenHachuushaIraibusho ");
            sb.Append(",AnkenHachuushaTantousha ");
            sb.Append(",AnkenHachuushaTEL ");
            sb.Append(",AnkenHachuushaFAX ");
            sb.Append(",AnkenHachuushaMail ");
            sb.Append(",AnkenHachuushaIraiYuubin ");
            sb.Append(",AnkenHachuushaIraiJuusho ");
            sb.Append(",AnkenHachuushaKeiyakuBusho ");
            sb.Append(",AnkenHachuushaKeiyakuTantou ");
            sb.Append(",AnkenHachuushaKeiyakuTEL ");
            sb.Append(",AnkenHachuushaKeiyakuFAX ");
            sb.Append(",AnkenHachuushaKeiyakuMail ");
            sb.Append(",AnkenHachuushaKeiyakuYuubin ");
            sb.Append(",AnkenHachuushaKeiyakuJuusho ");
            sb.Append(",AnkenHachuuDaihyouYakushoku ");
            sb.Append(",AnkenHachuuDaihyousha ");
            sb.Append(",AnkenRosenKawamei ");
            sb.Append(",AnkenGyoumuItakuKasho ");
            sb.Append(",AnkenJititaiKibunID ");
            sb.Append(",AnkenJititaiKubun ");
            sb.Append(",AnkenKeiyakuToshoNo ");
            sb.Append(",AnkenKirokuToshoNo ");
            sb.Append(",AnkenKirokuHokanNo ");
            sb.Append(",AnkenCDHokan ");
            sb.Append(",AnkenSeikaButsuHokanFile ");
            sb.Append(",AnkenSeikabutsuHokanbako ");
            sb.Append(",AnkenKokyakuHyoukaComment ");
            sb.Append(",AnkenToukaiHyoukaComment ");
            sb.Append(",AnkenKenCD ");
            sb.Append(",AnkenToshiCD ");
            sb.Append(",AnkenKeiyakusho ");
            sb.Append(",AnkenEizen ");
            sb.Append(",AnkenTantoushaMei ");
            sb.Append(",GyoumuKanrishaMei ");
            sb.Append(",AnkenGyoumuKubunCD ");
            sb.Append(",AnkenHachuushaKaMei ");
            sb.Append(",AnkenKeiyakuKoukiKaishibi ");
            sb.Append(",AnkenKeiyakuKoukiKanryoubi ");
            sb.Append(",AnkenKeiyakuTeiketsubi ");
            sb.Append(",AnkenKeiyakuZeikomiKingaku ");     // 契約タブの契約金額の税込
            sb.Append(",AnkenKeiyakuUriageHaibunGakuC ");  // 契約タブの受託金額配分の調査部、配分額（税込）
            sb.Append(",AnkenKeiyakuUriageHaibunGakuJ ");  // 契約タブの受託金額配分の事業普及部、配分額（税込）
            sb.Append(",AnkenKeiyakuUriageHaibunGakuJs "); // 契約タブの受託金額配分の情報システム部、配分額（税込）
            sb.Append(",AnkenKeiyakuUriageHaibunGakuK ");  // 契約タブの受託金額配分の総合研究所、配分額（税込）
            sb.Append(",AnkenKeiyakuUriageHaibunGakuR ");  // なし
            sb.Append(",AnkenKeiyakuSakuseibi ");
            sb.Append(",AnkenAnkenBangou ");
            sb.Append(",AnkenKeikakuBangou ");
            sb.Append(",AnkenHikiaijhokyo ");
            sb.Append(",AnkenKeikakuAnkenMei ");
            sb.Append(",AnkenToukaiSankouMitsumori ");
            sb.Append(",AnkenToukaiJyutyuIyoku ");
            sb.Append(",AnkenToukaiSankouMitsumoriGaku ");
            sb.Append(",AnkenHachushaKaMei ");
            sb.Append(",AnkenHachushaCD ");
            sb.Append(",AnkenToukaiOusatu ");
            sb.Append(",AnkenKoukiNendo ");
            sb.Append(" FROM AnkenJouhou WHERE AnkenJouhou.AnkenJouhouID = ");
            sb.Append(AnkenID);
            return sb.ToString();
        }

        /// <summary>
        /// SQL文作成
        /// </summary>
        /// <param name="ankenNo">作成案件番号</param>
        /// <param name="flag">0:赤伝、1:黒伝</param>
        /// <returns></returns>
        private string getAnkenJouhouZenkaiRakusatsuInsertSQL(int ankenNo, int flag = 0)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("INSERT INTO AnkenJouhouZenkaiRakusatsu ( ");
            sb.Append("AnkenJouhouID ");
            sb.Append(",AnkenZenkaiJutakuKingaku ");
            sb.Append(",AnkenZenkaiRakusatsuID ");
            sb.Append(",AnkenZenkaiJutakuBangou ");
            sb.Append(",AnkenZenkaiJutakuEdaban ");
            sb.Append(",AnkenZenkaiAnkenBangou ");
            sb.Append(",AnkenZenkaiRakusatsushaID ");
            sb.Append(",AnkenZenkaiRakusatsusha ");
            sb.Append(",AnkenZenkaiGyoumuMei ");
            sb.Append(",AnkenZenkaiKyougouKigyouCD ");
            sb.Append(",AnkenZenkaiJutakuZeinuki ");
            sb.Append(",KeiyakuZenkaiRakusatsushaID ");
            sb.Append(" ) SELECT ");
            sb.Append(ankenNo);
            if(flag == 0)
            {
                sb.Append(",-AnkenZenkaiJutakuKingaku ");
            }
            else
            {
                sb.Append(",AnkenZenkaiJutakuKingaku ");
            }
            sb.Append(",AnkenZenkaiRakusatsuID ");
            sb.Append(",AnkenZenkaiJutakuBangou ");
            sb.Append(",AnkenZenkaiJutakuEdaban ");
            sb.Append(",AnkenZenkaiAnkenBangou ");
            sb.Append(",AnkenZenkaiRakusatsushaID ");
            sb.Append(",AnkenZenkaiRakusatsusha ");
            sb.Append(",AnkenZenkaiGyoumuMei ");
            sb.Append(",AnkenZenkaiKyougouKigyouCD ");
            sb.Append(",AnkenZenkaiJutakuZeinuki ");
            sb.Append(",KeiyakuZenkaiRakusatsushaID ");
            sb.Append(" FROM AnkenJouhouZenkaiRakusatsu WHERE AnkenJouhouZenkaiRakusatsu.AnkenJouhouID = ");
            sb.Append(AnkenID);
            return sb.ToString();
        }

        /// <summary>
        /// SQL文作成
        /// </summary>
        /// <param name="ankenNo">作成案件番号</param>
        /// <returns></returns>
        private string getKokyakuKeiyakuJouhouInsertSQL(int ankenNo)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("INSERT INTO KokyakuKeiyakuJouhou ( ");
            sb.Append("AnkenJouhouID ");
            sb.Append(",KokyakuKeiyakuID ");
            sb.Append(",KokyakuCreateUser   ");
            sb.Append(",KokyakuCreateDate   ");
            sb.Append(",KokyakuCreateProgram");
            sb.Append(",KokyakuUpdateUser   ");
            sb.Append(",KokyakuUpdateDate   ");
            sb.Append(",KokyakuDeleteFlag   ");
            sb.Append(",KokyakuUpdateProgram");
            sb.Append(",KokyakuKeiyakuTanka ");
            sb.Append(",KokyakuKeiyakuChosakuken ");
            sb.Append(",KokyakuKeiyakuKeisai ");
            sb.Append(",KokyakuKeiyakuTokchouChosaku ");
            sb.Append(",KokyakuKeiyakuRiyuu ");
            sb.Append(",KokyakuMaebaraiJoukou ");
            sb.Append(",KokyakuMaebaraiSeikyuu ");
            sb.Append(",KokyakuSekkeiTanka ");
            sb.Append(",KokyakuSekisanKijun ");
            sb.Append(",KokyakuKaiteiGetsu ");
            sb.Append(",KokyakuShichouson ");
            sb.Append(",KokyakuGijutsuCenter ");
            sb.Append(",KokyakuSonota ");
            sb.Append(",KokyakuKeiyakuRiyuuTou ");
            sb.Append(",KokyakuDataTeikyou ");
            sb.Append(",KokyakuAlpha ");
            sb.Append(",KokyakuDataDoboku ");
            sb.Append(",KokyakuDataNourin ");
            sb.Append(",KokyakuDataEizen ");
            sb.Append(",KokyakuDataSonota ");
            sb.Append(",KokyakuDataSekouP ");
            sb.Append(",KokyakuDataDobokuKouji ");
            sb.Append(",KokyakuDataRIBC ");
            sb.Append(",KokyakuDataGoukei ");
            sb.Append(",KokyakuDataKeisaiTanka ");
            sb.Append(",KokyakuDataWebTeikyou ");
            sb.Append(",KokyakuDataKeiyaku ");
            sb.Append(",KokyakuDataTempFile ");
            sb.Append(",KokyakuDataTempFileData ");
            sb.Append(",KokyakuData05Comment ");
            sb.Append(",KokyakuData06Comment ");
            sb.Append(",KokyakuData07Comment ");
            sb.Append(",KokyakuDataMeiki ");
            sb.Append(",KokyakuDataTeikyouTensu ");
            sb.Append(" ) SELECT ");
            sb.Append(ankenNo);
            sb.Append(",");
            sb.Append(ankenNo);
            sb.Append(",'");
            sb.Append(UserInfos[0]);
            sb.Append("' ");
            sb.Append(",GETDATE() ");
            sb.Append(",'ChangeKianEntry' ");
            sb.Append(",'");
            sb.Append(UserInfos[0]);
            sb.Append("' ");
            sb.Append(",GETDATE() ");
            sb.Append(",1");
            sb.Append(",'ChangeKianEntry' ");
            sb.Append(",KokyakuKeiyakuTanka ");
            sb.Append(",KokyakuKeiyakuChosakuken ");
            sb.Append(",KokyakuKeiyakuKeisai ");
            sb.Append(",KokyakuKeiyakuTokchouChosaku ");
            sb.Append(",KokyakuKeiyakuRiyuu ");
            sb.Append(",KokyakuMaebaraiJoukou ");
            sb.Append(",KokyakuMaebaraiSeikyuu ");
            sb.Append(",KokyakuSekkeiTanka ");
            sb.Append(",KokyakuSekisanKijun ");
            sb.Append(",KokyakuKaiteiGetsu ");
            sb.Append(",KokyakuShichouson ");
            sb.Append(",KokyakuGijutsuCenter ");
            sb.Append(",KokyakuSonota ");
            sb.Append(",KokyakuKeiyakuRiyuuTou ");
            sb.Append(",KokyakuDataTeikyou ");
            sb.Append(",KokyakuAlpha ");
            sb.Append(",KokyakuDataDoboku ");
            sb.Append(",KokyakuDataNourin ");
            sb.Append(",KokyakuDataEizen ");
            sb.Append(",KokyakuDataSonota ");
            sb.Append(",KokyakuDataSekouP ");
            sb.Append(",KokyakuDataDobokuKouji ");
            sb.Append(",KokyakuDataRIBC ");
            sb.Append(",KokyakuDataGoukei ");
            sb.Append(",KokyakuDataKeisaiTanka ");
            sb.Append(",KokyakuDataWebTeikyou ");
            sb.Append(",KokyakuDataKeiyaku ");
            sb.Append(",KokyakuDataTempFile ");
            sb.Append(",KokyakuDataTempFileData ");
            sb.Append(",KokyakuData05Comment ");
            sb.Append(",KokyakuData06Comment ");
            sb.Append(",KokyakuData07Comment ");
            sb.Append(",KokyakuDataMeiki ");
            sb.Append(",KokyakuDataTeikyouTensu ");
            sb.Append(" FROM KokyakuKeiyakuJouhou WHERE KokyakuKeiyakuJouhou.AnkenJouhouID = ");
            sb.Append(AnkenID);
            return sb.ToString();
        }

        private string getGyoumuJouhouInsertSql(int ankenNo)
        {
            return "INSERT INTO GyoumuJouhou ( " +
                    "AnkenJouhouID " +
                    ",GyoumuJouhouID " +
                    ",GyoumuCreateDate " +
                    ",GyoumuCreateUser    " +
                    ",GyoumuCreateProgram " +
                    ",GyoumuUpdateDate    " +
                    ",GyoumuUpdateUser    " +
                    ",GyoumuUpdateProgram " +
                    ",GyoumuDeleteFlag " +

                    ",GyoumuHyouten " +
                    ",KanriGijutsushaCD " +
                    ",GyoumuKanriHyouten " +
                    ",ShousaTantoushaCD " +
                    ",SinsaTantoushaCD " +
                    ",GyoumuTECRISTourokuBangou " +
                    ",GyoumuKeisaiTankaTeikyou " +
                    ",GyoumuChosakukenJouto " +
                    ",GyoumuSeikyuubi " +
                    ",GyoumuSeikyuusho " +
                    ",GyoumuHikiwatashiNaiyou " +
                    ",KanriGijutsushaNM " +
                    ",ShousaTantoushaNM " +
                    ",SinsaTantoushaNM " +
                    ",GyoumuShousaHyouten " +
                    " ) SELECT " +
                    ankenNo +
                    "," + ankenNo +
                    ",GETDATE() " +
                    ",'" + UserInfos[0] + "' " +
                    ",'ChangeKianEntry' " +
                    ",GETDATE() " +
                    ",'" + UserInfos[0] + "' " +
                    ",'ChangeKianEntry' " +
                    ",1 " +

                    ",GyoumuHyouten " +
                    ",KanriGijutsushaCD " +
                    ",GyoumuKanriHyouten " +
                    ",ShousaTantoushaCD " +
                    ",SinsaTantoushaCD " +
                    ",GyoumuTECRISTourokuBangou " +
                    ",GyoumuKeisaiTankaTeikyou " +
                    ",GyoumuChosakukenJouto " +
                    ",GyoumuSeikyuubi " +
                    ",GyoumuSeikyuusho " +
                    ",GyoumuHikiwatashiNaiyou " +
                    ",KanriGijutsushaNM " +
                    ",ShousaTantoushaNM " +
                    ",SinsaTantoushaNM " +
                    ",GyoumuShousaHyouten " +
                    " FROM GyoumuJouhou WHERE GyoumuJouhou.AnkenJouhouID = " + AnkenID;
        }

        private string getGyoumuJouhouHyouronTantouL1InsertSQL(int ankenNo)
        {
            return "INSERT INTO GyoumuJouhouHyouronTantouL1 ( " +
                                "GyoumuJouhouID " +
                                ",HyouronTantouID " +

                                ",HyouronTantoushaCD " +
                                ",HyouronTantoushaMei " +
                                ",HyouronnTantoushaHyouten " +
                               " ) SELECT " +
                                ankenNo +
                                ",HyouronTantouID " +

                                ",HyouronTantoushaCD " +
                                ",HyouronTantoushaMei " +
                                ",HyouronnTantoushaHyouten " +
                                " FROM GyoumuJouhouHyouronTantouL1 WHERE GyoumuJouhouHyouronTantouL1.GyoumuJouhouID = " + AnkenID;
        }

        private void exeGyoumuJouhouMadoguchi(int ankenNo, SqlCommand cmd)
        {
            DataTable gyoumuJouhouDT = new DataTable();
            cmd.CommandText = "SELECT " +
                    " GyoumuJouhouMadoKojinCD " +
                    ",GyoumuJouhouMadoChousainMei " +
                    ",GyoumuJouhouMadoGyoumuBushoCD " +
                    ",GyoumuJouhouMadoShibuMei " +
                    ",GyoumuJouhouMadoKamei " +
                    " FROM GyoumuJouhouMadoguchi WHERE GyoumuJouhouMadoguchi.GyoumuJouhouID = " + AnkenID;
            var sda = new SqlDataAdapter(cmd);
            sda.Fill(gyoumuJouhouDT);

            if (gyoumuJouhouDT != null && gyoumuJouhouDT.Rows.Count > 0)
            {
                for (int i = 0; i < gyoumuJouhouDT.Rows.Count; i++)
                {
                    cmd.CommandText = "INSERT INTO GyoumuJouhouMadoguchi ( " +
                            "GyoumuJouhouID " +
                            ",GyoumuJouhouMadoguchiID " +
                            ",GyoumuJouhouMadoKojinCD " +
                            ",GyoumuJouhouMadoChousainMei " +
                            ",GyoumuJouhouMadoGyoumuBushoCD " +
                            ",GyoumuJouhouMadoShibuMei " +
                            ",GyoumuJouhouMadoKamei " +
                            " ) VALUES( " +
                            ankenNo +                                                       // GyoumuJouhouID
                            "," + GlobalMethod.getSaiban("GyoumuJouhouMadoguchiID") + " " + // GyoumuJouhouMadoguchiID
                            "," + gyoumuJouhouDT.Rows[i][0].ToString() + " " +              // GyoumuJouhouMadoKojinCD
                            ",N'" + gyoumuJouhouDT.Rows[i][1].ToString() + "' " +            // GyoumuJouhouMadoChousainMei
                            ",N'" + gyoumuJouhouDT.Rows[i][2].ToString() + "' " +            // GyoumuJouhouMadoGyoumuBushoCD
                            ",N'" + gyoumuJouhouDT.Rows[i][3].ToString() + "' " +            // GyoumuJouhouMadoShibuMei
                            ",N'" + gyoumuJouhouDT.Rows[i][4].ToString() + "' " +            // GyoumuJouhouMadoKamei
                            ") ";
                    Console.WriteLine(cmd.CommandText);
                    var result = cmd.ExecuteNonQuery();
                }
            }
        }

        private string getGyoumuJouhouHyoutenBushoInsertSql(int ankenNo)
        {
            return "INSERT INTO GyoumuJouhouHyoutenBusho ( " +
                                "GyoumuJouhouID " +
                                ",HyoutenBushoID " +

                                ",HyoutenKyouryokuBushoID " +
                                ",HyoutenKyouryokuBushoMei " +
                                " ) SELECT " +
                                ankenNo +
                                ",HyoutenBushoID " +

                                ",HyoutenKyouryokuBushoID " +
                                ",HyoutenKyouryokuBushoMei " +
                                " FROM GyoumuJouhouHyoutenBusho WHERE GyoumuJouhouHyoutenBusho.GyoumuJouhouID = " + AnkenID;
        }

        /// <summary>
        /// SQL文作成
        /// </summary>
        /// <param name="ankenNo">作成案件番号</param>
        /// <param name="flag">0:赤伝、1:黒伝</param>
        /// <returns></returns>
        private string getKeiyakuJouhouEntoryInsertSql(int ankenNo, int flag = 0)
        {
            string sFix = "- ";
            if(flag == 1)
            {
                sFix = "";
            }
            return "INSERT INTO KeiyakuJouhouEntory ( " +
                                    "AnkenJouhouID " +
                                    ",KeiyakuJouhouEntoryID " +
                                    ",KeiyakuKeiyakuKingaku " +
                                    ",KeiyakuZeikomiKingaku " +
                                    ",KeiyakuuchizeiKingaku " +
                                    ",KeiyakuUriageHaibunCho " +
                                    ",KeiyakuUriageHaibunGakuCho1 " +
                                    ",KeiyakuUriageHaibunGakuCho2 " +
                                    ",KeiyakuUriageHaibunJo " +
                                    ",KeiyakuUriageHaibunGakuJo1 " +
                                    ",KeiyakuUriageHaibunGakuJo2 " +
                                    ",KeiyakuUriageHaibunJosys " +
                                    ",KeiyakuUriageHaibunGakuJosys1 " +
                                    ",KeiyakuUriageHaibunGakuJosys2 " +
                                    ",KeiyakuUriageHaibunKei " +
                                    ",KeiyakuUriageHaibunGakuKei1 " +
                                    ",KeiyakuUriageHaibunGakuKei2 " +
                                    ",KeiyakuZentokin " +
                                    ",KeiyakuSeikyuuKingaku1 " +
                                    ",KeiyakuSeikyuuKingaku2 " +
                                    ",KeiyakuSeikyuuKingaku3 " +
                                    ",KeiyakuSeikyuuKingaku4 " +
                                    ",KeiyakuSeikyuuKingaku5 " +
                                    ",KeiyakuCreateDate " +
                                    ",KeiyakuCreateUser " +
                                    ",KeiyakuCreateProgram " +
                                    ",KeiyakuUpdateDate " +
                                    ",KeiyakuUpdateUser " +
                                    ",KeiyakuUpdateProgram " +
                                    ",KeiyakuBetsuKeiyakuKingaku " +
                                    ",KeiyakuKeiyakuKingakuKei " +
                                    ",KeiyakuUriageHaibunChoGoukei " +
                                    ",KeiyakuUriageHaibunJoGoukei " +
                                    ",KeiyakuUriageHaibunJosysGoukei " +
                                    ",KeiyakuUriageHaibunKeiGoukei " +
                                    ",KeiyakuUriageHaibunGoukei " +
                                    ",KeiyakuHaibunChoZeinuki " +
                                    ",KeiyakuHaibunJoZeinuki " +
                                    ",KeiyakuHaibunJosysZeinuki " +
                                    ",KeiyakuHaibunKeiZeinuki " +
                                    ",KeiyakuHaibunZeinukiKei " +
                                    ",KeiyakuDeleteFlag " +

                                    ",KeiyakuSakuseibi " +
                                    ",KeiyakuSakuseiKubunID " +
                                    ",KeiyakuSakuseiKubun " +
                                    ",KeiyakuHachuushaMei " +
                                    ",KeiyakuGyoumuKubun " +
                                    ",KeiyakuGyoumuMei " +
                                    ",JutakuBushoCD " +
                                    ",KeiyakuTantousha " +
                                    ",KeiyakuJutakubangou " +
                                    ",KeiyakuEdaban " +
                                    ",KeiyakuKianzumi " +
                                    ",KeiyakuNyuusatsuYoteibi " +
                                    ",KeiyakuKeiyakuTeiketsubi " +
                                    ",KeiyakuKeiyakuKoukiKaishibi " +
                                    ",KeiyakuKeiyakuKoukiKanryoubi " +
                                    ",KeiyakuShouhizeiritsu " +
                                    ",KeiyakuRIBCKeishiki " +
                                    ",KeiyakuUriageHaibunCho1 " +
                                    ",KeiyakuUriageHaibunCho2 " +
                                    ",KeiyakuUriageHaibunJo1 " +
                                    ",KeiyakuUriageHaibunJo2 " +
                                    ",KeiyakuUriageHaibunJosys1 " +
                                    ",KeiyakuUriageHaibunJosys2 " +
                                    ",KeiyakuUriageHaibunKei1 " +
                                    ",KeiyakuUriageHaibunKei2 " +
                                    ",KeiyakuHenkoukanryoubi " +
                                    ",KeiyakuHenkouChuushiRiyuu " +
                                    ",KeiyakuBikou " +
                                    ",KeiyakuShosha " +
                                    ",KeiyakuTokkiShiyousho " +
                                    ",KeiyakuMitsumorisho " +
                                    ",KeiyakuTanpinChousaMitsumorisho " +
                                    ",KeiyakuSonota " +
                                    ",KeiyakuSonotaNaiyou " +
                                    ",KeiyakuSeikyuubi " +
                                    ",KeiyakuKeiyakusho " +
                                    ",KeiyakuZentokinUkewatashibi " +
                                    ",KeiyakuSeikyuusaki " +
                                    ",KeiyakuSeikyuuTaishouKoukiS1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE1 " +
                                    ",KeiyakuSeikyuubi1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE2 " +
                                    ",KeiyakuSeikyuubi2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS3 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE3 " +
                                    ",KeiyakuSeikyuubi3 " +
                                    ",KeiyakuKankeibusho1 " +
                                    ",KeiyakuKankeibusho2 " +
                                    ",KeiyakuKankeibusho3 " +
                                    ",KeiyakuKankeibusho4 " +
                                    ",KeiyakuKankeibusho5 " +
                                    ",KeiyakuKankeibusho6 " +
                                    ",KeiyakuKankeibusho7 " +
                                    ",KeiyakuKankeibusho8 " +
                                    ",KeiyakuKankeibusho9 " +
                                    ",KeiyakuKankeibusho10 " +
                                    ",KeiyakuKankeibusho11 " +
                                    ",KeiyakuKankeibusho12 " +
                                    ",KeiyakuKankeibusho14 " +
                                    ",KeiyakuKankeibusho15 " +
                                    ",KeiyakuKankeibusho13 " +
                                    ",KeiyakuNyuukinYoteibi " +
                                    ",KeiyakuUriageHaibunCho1Mei " +
                                    ",KeiyakuUriageHaibunCho2Mei " +
                                    ",KeiyakuUriageHaibunJo1Mei " +
                                    ",KeiyakuUriageHaibunJo2Mei " +
                                    ",KeiyakuUriageHaibunJosys1Mei " +
                                    ",KeiyakuUriageHaibunJosys2Mei " +
                                    ",KeiyakuUriageHaibunKei1Mei " +
                                    ",KeiyakuUriageHaibunKei2Mei " +
                                    ",KeiyakuUriageHaibunRIBC " +
                                    ",KeiyakuUriageHaibunRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC1Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC2 " +
                                    ",KeiyakuUriageHaibunRIBC2Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC2 " +
                                    ",KeiyakuSeikyuubi4 " +
                                    ",KeiyakuSeikyuubi5 " +
                                    ",KeiyakuTankeiMikomiCho " +
                                    ",KeiyakuTankeiMikomiJo " +
                                    ",KeiyakuTankeiMikomiJosys " +
                                    ",KeiyakuTankeiMikomiKei " +
                                    ",KeiyakuKurikoshiCho " +
                                    ",KeiyakuKurikoshiJo " +
                                    ",KeiyakuKurikoshiJosys " +
                                    ",KeiyakuKurikoshiKei " +
                                    " ) SELECT " +
                                    ankenNo +
                                    "," + ankenNo +
                                    "," + sFix + "KeiyakuKeiyakuKingaku " +
                                    "," + sFix + "KeiyakuZeikomiKingaku " +
                                    "," + sFix + "KeiyakuuchizeiKingaku " +
                                    "," + sFix + "KeiyakuUriageHaibunCho " +
                                    "," + sFix + "KeiyakuUriageHaibunGakuCho1 " +
                                    "," + sFix + "KeiyakuUriageHaibunGakuCho2 " +
                                    "," + sFix + "KeiyakuUriageHaibunJo " +
                                    "," + sFix + "KeiyakuUriageHaibunGakuJo1 " +
                                    "," + sFix + "KeiyakuUriageHaibunGakuJo2 " +
                                    "," + sFix + "KeiyakuUriageHaibunJosys " +
                                    "," + sFix + "KeiyakuUriageHaibunGakuJosys1 " +
                                    "," + sFix + "KeiyakuUriageHaibunGakuJosys2 " +
                                    "," + sFix + "KeiyakuUriageHaibunKei " +
                                    "," + sFix + "KeiyakuUriageHaibunGakuKei1 " +
                                    "," + sFix + "KeiyakuUriageHaibunGakuKei2 " +
                                    "," + sFix + "KeiyakuZentokin " +
                                    "," + sFix + "KeiyakuSeikyuuKingaku1 " +
                                    "," + sFix + "KeiyakuSeikyuuKingaku2 " +
                                    "," + sFix + "KeiyakuSeikyuuKingaku3 " +
                                    "," + sFix + "KeiyakuSeikyuuKingaku4 " +
                                    "," + sFix + "KeiyakuSeikyuuKingaku5 " +
                                    ",GETDATE() " +
                                    ",'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +
                                    ",GETDATE() " +
                                    ",'" + UserInfos[0] + "' " +
                                    ",'ChangeKianEntry' " +
                                    "," + sFix + "KeiyakuBetsuKeiyakuKingaku " +
                                    "," + sFix + "KeiyakuKeiyakuKingakuKei " +
                                    "," + sFix + "KeiyakuUriageHaibunChoGoukei " +
                                    "," + sFix + "KeiyakuUriageHaibunJoGoukei " +
                                    "," + sFix + "KeiyakuUriageHaibunJosysGoukei " +
                                    "," + sFix + "KeiyakuUriageHaibunKeiGoukei " +
                                    "," + sFix + "KeiyakuUriageHaibunGoukei " +
                                    "," + sFix + "KeiyakuHaibunChoZeinuki " +
                                    "," + sFix + "KeiyakuHaibunJoZeinuki " +
                                    "," + sFix + "KeiyakuHaibunJosysZeinuki " +
                                    "," + sFix + "KeiyakuHaibunKeiZeinuki " +
                                    "," + sFix + "KeiyakuHaibunZeinukiKei " +
                                    ",1 " +

                                    ",KeiyakuSakuseibi " +
                                    ",KeiyakuSakuseiKubunID " +
                                    ",KeiyakuSakuseiKubun " +
                                    ",KeiyakuHachuushaMei " +
                                    ",KeiyakuGyoumuKubun " +
                                    ",KeiyakuGyoumuMei " +
                                    ",JutakuBushoCD " +
                                    ",KeiyakuTantousha " +
                                    ",KeiyakuJutakubangou " +
                                    ",KeiyakuEdaban " +
                                    ",KeiyakuKianzumi " +
                                    ",KeiyakuNyuusatsuYoteibi " +
                                    ",KeiyakuKeiyakuTeiketsubi " +
                                    ",KeiyakuKeiyakuKoukiKaishibi " +
                                    ",KeiyakuKeiyakuKoukiKanryoubi " +
                                    ",KeiyakuShouhizeiritsu " +
                                    ",KeiyakuRIBCKeishiki " +
                                    ",KeiyakuUriageHaibunCho1 " +
                                    ",KeiyakuUriageHaibunCho2 " +
                                    ",KeiyakuUriageHaibunJo1 " +
                                    ",KeiyakuUriageHaibunJo2 " +
                                    ",KeiyakuUriageHaibunJosys1 " +
                                    ",KeiyakuUriageHaibunJosys2 " +
                                    ",KeiyakuUriageHaibunKei1 " +
                                    ",KeiyakuUriageHaibunKei2 " +
                                    ",KeiyakuHenkoukanryoubi " +
                                    ",KeiyakuHenkouChuushiRiyuu " +
                                    ",KeiyakuBikou " +
                                    ",KeiyakuShosha " +
                                    ",KeiyakuTokkiShiyousho " +
                                    ",KeiyakuMitsumorisho " +
                                    ",KeiyakuTanpinChousaMitsumorisho " +
                                    ",KeiyakuSonota " +
                                    ",KeiyakuSonotaNaiyou " +
                                    ",KeiyakuSeikyuubi " +
                                    ",KeiyakuKeiyakusho " +
                                    ",KeiyakuZentokinUkewatashibi " +
                                    ",KeiyakuSeikyuusaki " +
                                    ",KeiyakuSeikyuuTaishouKoukiS1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE1 " +
                                    ",KeiyakuSeikyuubi1 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE2 " +
                                    ",KeiyakuSeikyuubi2 " +
                                    ",KeiyakuSeikyuuTaishouKoukiS3 " +
                                    ",KeiyakuSeikyuuTaishouKoukiE3 " +
                                    ",KeiyakuSeikyuubi3 " +
                                    ",KeiyakuKankeibusho1 " +
                                    ",KeiyakuKankeibusho2 " +
                                    ",KeiyakuKankeibusho3 " +
                                    ",KeiyakuKankeibusho4 " +
                                    ",KeiyakuKankeibusho5 " +
                                    ",KeiyakuKankeibusho6 " +
                                    ",KeiyakuKankeibusho7 " +
                                    ",KeiyakuKankeibusho8 " +
                                    ",KeiyakuKankeibusho9 " +
                                    ",KeiyakuKankeibusho10 " +
                                    ",KeiyakuKankeibusho11 " +
                                    ",KeiyakuKankeibusho12 " +
                                    ",KeiyakuKankeibusho14 " +
                                    ",KeiyakuKankeibusho15 " +
                                    ",KeiyakuKankeibusho13 " +
                                    ",KeiyakuNyuukinYoteibi " +
                                    ",KeiyakuUriageHaibunCho1Mei " +
                                    ",KeiyakuUriageHaibunCho2Mei " +
                                    ",KeiyakuUriageHaibunJo1Mei " +
                                    ",KeiyakuUriageHaibunJo2Mei " +
                                    ",KeiyakuUriageHaibunJosys1Mei " +
                                    ",KeiyakuUriageHaibunJosys2Mei " +
                                    ",KeiyakuUriageHaibunKei1Mei " +
                                    ",KeiyakuUriageHaibunKei2Mei " +
                                    ",KeiyakuUriageHaibunRIBC " +
                                    ",KeiyakuUriageHaibunRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC1Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC1 " +
                                    ",KeiyakuUriageHaibunRIBC2 " +
                                    ",KeiyakuUriageHaibunRIBC2Mei " +
                                    ",KeiyakuUriageHaibunGakuRIBC2 " +
                                    ",KeiyakuSeikyuubi4 " +
                                    ",KeiyakuSeikyuubi5 " +
                                    "," + sFix + "KeiyakuTankeiMikomiCho " +
                                    "," + sFix + "KeiyakuTankeiMikomiJo " +
                                    "," + sFix + "KeiyakuTankeiMikomiJosys " +
                                    "," + sFix + "KeiyakuTankeiMikomiKei " +
                                    "," + sFix + "KeiyakuKurikoshiCho " +
                                    "," + sFix + "KeiyakuKurikoshiJo " +
                                    "," + sFix + "KeiyakuKurikoshiJosys " +
                                    "," + sFix + "KeiyakuKurikoshiKei " +
                                    " FROM KeiyakuJouhouEntory WHERE KeiyakuJouhouEntory.AnkenJouhouID = " + AnkenID;
        }

        /// <summary>
        /// SQL文作成
        /// </summary>
        /// <param name="ankenNo">作成案件番号</param>
        /// <param name="flag">0:赤伝、1:黒伝</param>
        private string getRibcJouhouInsertSql(int ankenNo, int flag = 0)
        {
            string sFix = "- ";
            if (flag == 1)
            {
                sFix = "";
            }
            return "INSERT INTO RibcJouhou ( " +
                    "RibcID " +
                    ",RibcNo " +
                    ",RibcSeikyuKingaku " +

                    ",RibcKoukiStart " +
                    ",RibcKoukiEnd " +
                    ",RibcSeikyubi " +
                    ",RibcNouhinbi " +
                    ",RibcNyukinyoteibi " +
                    ",RibcUriageKeijyoTuki " +
                    ",RibcKankeibusho " +
                    ",RibcKubun " +
                    ",RibcKankeibushoMei " +
                    " ) SELECT " +
                    ankenNo +
                    ",RibcNo " +
                    "," + sFix + "RibcSeikyuKingaku " +

                    ",RibcKoukiStart " +
                    ",RibcKoukiEnd " +
                    ",RibcSeikyubi " +
                    ",RibcNouhinbi " +
                    ",RibcNyukinyoteibi " +
                    ",RibcUriageKeijyoTuki " +
                    ",RibcKankeibusho " +
                    ",RibcKubun " +
                    ",RibcKankeibushoMei " +
                    " FROM RibcJouhou WHERE RibcJouhou.RibcID = " + AnkenID;
        }

        /// <summary>
        /// SQL文作成
        /// </summary>
        /// <param name="ankenNo">作成案件番号</param>
        /// <param name="flag">0:赤伝、1:黒伝</param>
        private string getNyuusatsuJouhouInsertSql(int ankenNo, int flag = 0)
        {
            string sFix = "- ";
            if (flag == 1)
            {
                sFix = "";
            }
            return "INSERT INTO NyuusatsuJouhou ( " +
                    "AnkenJouhouID " +
                    ",NyuusatsuJouhouID " +
                    ",NyuusatsuMitsumorigaku " +
                    ",NyuusatsuOusatugaku " +
                    ",NyuusatsuRakusatugaku " +
                    ",NyuusatsuRakusatuSougaku " +
                    ",NyuusatsuNendoKurikoshigaku " +
                    ",NyuusatsuKyougouTashaID " +
                    ",NyuusatsuKyougouTasha " +
                    ",NyuusatsuRakusatsushaID " +
                    ",NyuusatsuRakusatsusha " +
                    ",NyuusatsuYoteiKakaku " +
                    ",NyuusatsuHoushiki " +
                    ",NyuusatsuKeiyakukeitaiCDSaishuu " +
                    ",NyuusatsuDenshiNyuusatsu " +
                    ",NyuusatsuTanpinMikomigaku " +
                    ",NyuusatsushaSuu " +
                    ",NyuusatsuGyoumuBikou " +
                    ",NyuusatsuShoruiSoufu " +
                    ",NyuusatsuDeleteFlag " +
                    ",NyuusatsuRakusatsuKekkaDate " +
                    ",NyuusatsuCreateDate " +
                    ",NyuusatsuCreateUser " +
                    ",NyuusatsuCreateProgram " +
                    ",NyuusatsuUpdateDate " +
                    ",NyuusatsuUpdateUser " +
                    ",NyuusatsuUpdateProgram " +

                    ",NyuusatsuRakusatsuShaJokyou " +
                    ",NyuusatsuRakusatsuGakuJokyou " +
                    ",NyuusatsuRakusatsuShokaiDate " +
                    ",NyuusatsuRakusatsuSaisyuDate " +
                    ",NyuusatsuKekkaMemo " +
                    " ) SELECT " +
                    "" + ankenNo +
                    "," + ankenNo +
                    "," + sFix + "NyuusatsuMitsumorigaku " +
                    "," + sFix + "NyuusatsuOusatugaku " +
                    "," + sFix + "NyuusatsuRakusatugaku " +
                    "," + sFix + "NyuusatsuRakusatuSougaku " +
                    "," + sFix + "NyuusatsuNendoKurikoshigaku " +
                    ",CASE WHEN NyuusatsuKyougouTashaID > 0 THEN NyuusatsuKyougouTashaID ELSE NULL END " +
                    ",CASE WHEN NyuusatsuKyougouTashaID > 0 THEN NyuusatsuKyougouTasha ELSE NULL END " +
                    ",CASE WHEN NyuusatsuRakusatsushaID > 0 THEN NyuusatsuRakusatsushaID ELSE NULL END " +
                    ",CASE WHEN NyuusatsuRakusatsushaID > 0 THEN NyuusatsuRakusatsusha ELSE NULL END " +
                    ",NyuusatsuYoteiKakaku " +
                    ",NyuusatsuHoushiki " +
                    ",NyuusatsuKeiyakukeitaiCDSaishuu " +
                    ",NyuusatsuDenshiNyuusatsu " +
                    ",NyuusatsuTanpinMikomigaku " +
                    ",NyuusatsushaSuu " +
                    ",NyuusatsuGyoumuBikou " +
                    ",NyuusatsuShoruiSoufu " +
                    ",1 " +
                    ",NyuusatsuRakusatsuKekkaDate " +
                    ",GETDATE() " +
                    ",N'" + UserInfos[0] + "' " +
                    ",'ChangeKianEntry' " +
                    ",GETDATE() " +
                    ",'" + UserInfos[0] + "' " +
                    ",'ChangeKianEntry' " +

                    ",NyuusatsuRakusatsuShaJokyou " +
                    ",NyuusatsuRakusatsuGakuJokyou " +
                    ",NyuusatsuRakusatsuShokaiDate " +
                    ",NyuusatsuRakusatsuSaisyuDate " +
                    ",NyuusatsuKekkaMemo " +
                    " FROM NyuusatsuJouhou WHERE NyuusatsuJouhou.NyuusatsuJouhouID = " + AnkenID;
        }

        /// <summary>
        /// SQL文作成
        /// </summary>
        /// <param name="ankenNo">作成案件番号</param>
        /// <param name="flag">0:赤伝、1:黒伝</param>
        private string getNyuusatsuJouhouOusatsushaInsertSql(int ankenNo, int flag = 0)
        {
            string sFix = "- ";
            if (flag == 1)
            {
                sFix = "";
            }
            return "INSERT INTO NyuusatsuJouhouOusatsusha ( " +
                                "NyuusatsuJouhouID" +
                                ",NyuusatsuOusatsuID" +
                                ",NyuusatsuOusatsuKingaku" +

                                ",NyuusatsuOusatsushaID" +
                                ",NyuusatsuOusatsusha" +
                                ",NyuusatsuOusatsuKyougouTashaID" +
                                ",NyuusatsuOusatsuKyougouKigyouCD" +
                                ",NyuusatsuRakusatsuJyuni" +
                                ",NyuusatsuRakusatsuJokyou" +
                                ",NyuusatsuRakusatsuComment" +
                                " ) SELECT " +
                                ankenNo +
                                ",ROW_NUMBER() OVER(ORDER BY NyuusatsuJouhouID) " +    // ",NyuusatsuOusatsuID" +
                                "," + sFix + "NyuusatsuOusatsuKingaku" +

                                ",NyuusatsuOusatsushaID" +
                                ",NyuusatsuOusatsusha" +
                                ",NyuusatsuOusatsuKyougouTashaID" +
                                ",NyuusatsuOusatsuKyougouKigyouCD" +
                                ",NyuusatsuRakusatsuJyuni" +
                                ",NyuusatsuRakusatsuJokyou" +
                                ",NyuusatsuRakusatsuComment" +
                                " FROM NyuusatsuJouhouOusatsusha WHERE NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID = " + AnkenID;
        }

        /// <summary>
        /// SQL文作成
        /// </summary>
        /// <param name="ankenNo">作成する案件番号</param>
        /// <param name="gyoumuHaibunID">業務配分ID</param>
        /// <param name="flag">0:赤伝、1:黒伝</param>
        /// <returns></returns>
        private string getGyoumuHaibunInsertSql(int ankenNo, int gyoumuHaibunID, int flag = 0)
        {
            string sFix = "- ";
            if (flag == 1)
            {
                sFix = "";
            }

            return "INSERT INTO GyoumuHaibun ( " +
                                    "GyoumuHaibunID " +
                                    ",GyoumuAnkenJouhouID " +
                                    ",GyoumuChosaBuRitsu " +
                                    ",GyoumuChosaBuGaku " +
                                    ",GyoumuJigyoFukyuBuRitsu " +
                                    ",GyoumuJigyoFukyuBuGaku " +
                                    ",GyoumuJyohouSystemBuRitsu " +
                                    ",GyoumuJyohouSystemBuGaku " +
                                    ",GyoumuSougouKenkyuJoRitsu " +
                                    ",GyoumuSougouKenkyuJoGaku " +
                                    ",GyoumuShizaiChousaRitsu " +
                                    ",GyoumuShizaiChousaGaku " +
                                    ",GyoumuEizenRitsu " +
                                    ",GyoumuEizenGaku " +
                                    ",GyoumuKikiruiChousaRitsu " +
                                    ",GyoumuKikiruiChousaGaku " +
                                    ",GyoumuKoujiChousahiRitsu " +
                                    ",GyoumuKoujiChousahiGaku " +
                                    ",GyoumuSanpaiFukusanbutsuRitsu " +
                                    ",GyoumuSanpaiFukusanbutsuGaku " +
                                    ",GyoumuHokakeChousaRitsu " +
                                    ",GyoumuHokakeChousaGaku " +
                                    ",GyoumuShokeihiChousaRitsu " +
                                    ",GyoumuShokeihiChousaGaku " +
                                    ",GyoumuGenkaBunsekiRitsu " +
                                    ",GyoumuGenkaBunsekiGaku " +
                                    ",GyoumuKijunsakuseiRitsu " +
                                    ",GyoumuKijunsakuseiGaku " +
                                    ",GyoumuKoukyouRoumuhiRitsu " +
                                    ",GyoumuKoukyouRoumuhiGaku " +
                                    ",GyoumuRoumuhiKoukyouigaiRitsu " +
                                    ",GyoumuRoumuhiKoukyouigaiGaku " +
                                    ",GyoumuSonotaChousabuRitsu " +
                                    ",GyoumuSonotaChousabuGaku " +
                                    ",GyoumuHibunKubun " +
                                    " ) SELECT " +
                                    GlobalMethod.getSaiban("GyoumuHaibunID") +
                                    "," + ankenNo +
                                    ",GyoumuChosaBuRitsu " +
                                    "," + sFix + "GyoumuChosaBuGaku " +
                                    ",GyoumuJigyoFukyuBuRitsu " +
                                    "," + sFix + "GyoumuJigyoFukyuBuGaku " +
                                    ",GyoumuJyohouSystemBuRitsu " +
                                    "," + sFix + "GyoumuJyohouSystemBuGaku " +
                                    ",GyoumuSougouKenkyuJoRitsu " +
                                    "," + sFix + "GyoumuSougouKenkyuJoGaku " +
                                    ",GyoumuShizaiChousaRitsu " +
                                    "," + sFix + "GyoumuShizaiChousaGaku " +
                                    ",GyoumuEizenRitsu " +
                                    "," + sFix + "GyoumuEizenGaku " +
                                    ",GyoumuKikiruiChousaRitsu " +
                                    "," + sFix + "GyoumuKikiruiChousaGaku " +
                                    ",GyoumuKoujiChousahiRitsu " +
                                    "," + sFix + "GyoumuKoujiChousahiGaku " +
                                    ",GyoumuSanpaiFukusanbutsuRitsu " +
                                    "," + sFix + "GyoumuSanpaiFukusanbutsuGaku " +
                                    ",GyoumuHokakeChousaRitsu " +
                                    "," + sFix + "GyoumuHokakeChousaGaku " +
                                    ",GyoumuShokeihiChousaRitsu " +
                                    "," + sFix + "GyoumuShokeihiChousaGaku " +
                                    ",GyoumuGenkaBunsekiRitsu " +
                                    "," + sFix + "GyoumuGenkaBunsekiGaku " +
                                    ",GyoumuKijunsakuseiRitsu " +
                                    "," + sFix + "GyoumuKijunsakuseiGaku " +
                                    ",GyoumuKoukyouRoumuhiRitsu " +
                                    "," + sFix + "GyoumuKoukyouRoumuhiGaku " +
                                    ",GyoumuRoumuhiKoukyouigaiRitsu " +
                                    "," + sFix + "GyoumuRoumuhiKoukyouigaiGaku " +
                                    ",GyoumuSonotaChousabuRitsu " +
                                    "," + sFix + "GyoumuSonotaChousabuGaku " +
                                    ",GyoumuHibunKubun " +
                                    " FROM GyoumuHaibun WHERE GyoumuHaibun.GyoumuHaibunID = " + gyoumuHaibunID;
        }

        // Garoon宛先追加登録
        private void insertGaroonAtesakiTsuika(SqlCommand cmd)
        {
            // 管理技術者が空でない場合、Garoon連携宛先追加テーブルに追加する
            if (item3_4_1.Text != null && item3_4_1.Text != "")
            {
                // Garoon連携宛先追加テーブルに存在するか確認
                cmd.CommandText = "SELECT  " +
                    " mj.MadoguchiID " +
                    ",gta.GaroonTsuikaAtesakiMadoguchiID " +
                    "FROM AnkenJouhou aj " +
                    "INNER JOIN MadoguchiJouhou mj ON mj.AnkenJouhouID = aj.AnkenJouhouID " +
                    "LEFT  JOIN GaroonTsuikaAtesaki gta ON gta.GaroonTsuikaAtesakiMadoguchiID = mj.MadoguchiID AND GaroonTsuikaAtesakiTantoushaCD = '" + item3_4_1_CD.Text + "'" +
                    "WHERE aj.AnkenJouhouID = '" + AnkenID + "' ";
                var sda = new SqlDataAdapter(cmd);
                var dt = new DataTable();
                sda.Fill(dt);

                String MadoguchiID = "";
                String GaroonTsuikaAtesakiMadoguchiID = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    MadoguchiID = dt.Rows[0][0].ToString();
                    GaroonTsuikaAtesakiMadoguchiID = dt.Rows[0][1].ToString();
                }
                // MadoguchiJouhouがあり、GaroonTsuikaAtesakiに存在しない場合、Garoon連携宛先追加テーブルに管理技術者を追加する
                if (MadoguchiID != "" && GaroonTsuikaAtesakiMadoguchiID == "")
                {
                    var dt9 = new DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      "KojinCD " +
                      ",ChousainMei " +
                      ",mc.GyoumuBushoCD " +
                      //",mb.ShibuMei + ' ' + IsNull(mb.KaMei,'') AS BushoMei " +
                      ",BushokanriboKamei " +
                      "FROM Mst_Chousain mc " +
                      "LEFT JOIN Mst_Busho mb ON mb.GyoumuBushoCD = mc.GyoumuBushoCD " +
                      "WHERE mc.KojinCD = '" + item3_4_1_CD.Text + "' ";

                    //データ取得
                    var sda9 = new SqlDataAdapter(cmd);
                    sda9.Fill(dt9);

                    string KojinCD = "";
                    string ChousainMei = "";
                    string GyoumuBushoCD = "";
                    string BushoMei = "";

                    if (dt9 != null && dt9.Rows.Count > 0)
                    {
                        KojinCD = dt9.Rows[0][0].ToString();
                        ChousainMei = dt9.Rows[0][1].ToString();
                        GyoumuBushoCD = dt9.Rows[0][2].ToString();
                        BushoMei = dt9.Rows[0][3].ToString();

                        int saibanGaroonTsuikaAtesakiID = GlobalMethod.getSaiban("GaroonTsuikaAtesakiID");

                        // GaroonTsuikaAtesakiに登録
                        cmd.CommandText = "INSERT INTO GaroonTsuikaAtesaki ( " +
                        " GaroonTsuikaAtesakiID " +
                        ",GaroonTsuikaAtesakiMadoguchiID " +
                        ",GaroonTsuikaAtesakiBushoCD " +
                        ",GaroonTsuikaAtesakiBusho " +
                        ",GaroonTsuikaAtesakiTantoushaCD " +
                        ",GaroonTsuikaAtesakiTantousha " +
                        ",GaroonTsuikaAtesakiCreateDate " +
                        ",GaroonTsuikaAtesakiCreateUser " +
                        ",GaroonTsuikaAtesakiCreateProgram " +
                        ",GaroonTsuikaAtesakiUpdateDate " +
                        ",GaroonTsuikaAtesakiUpdateUser " +
                        ",GaroonTsuikaAtesakiUpdateProgram " +
                        ",GaroonTsuikaAtesakiDeleteFlag " +
                        ") VALUES (" +
                        "'" + saibanGaroonTsuikaAtesakiID + "' " + // GaroonTsuikaAtesakiID
                        ",'" + MadoguchiID + "' " +                // GaroonTsuikaAtesakiMadoguchiID
                        ",'" + GyoumuBushoCD + "' " +              // GaroonTsuikaAtesakiBushoCD
                        ",N'" + BushoMei + "' " +                   // GaroonTsuikaAtesakiBusho
                        ",'" + KojinCD + "' " +                    // GaroonTsuikaAtesakiTantoushaCD
                        ",N'" + ChousainMei + "' " +                // GaroonTsuikaAtesakiTantousha
                        ",SYSDATETIME() " +                        // GaroonTsuikaAtesakiCreateDate
                        ",N'" + UserInfos[0] + "'" +                // GaroonTsuikaAtesakiCreateUser
                        ",'Entory_insertGaroonAtesakiTsuika'" +    // GaroonTsuikaAtesakiCreateProgram
                        ",SYSDATETIME() " +                        // GaroonTsuikaAtesakiUpdateDate
                        ",N'" + UserInfos[0] + "'" +                // GaroonTsuikaAtesakiUpdateUser
                        ",'Entory_insertGaroonAtesakiTsuika'" +    // GaroonTsuikaAtesakiUpdateProgram
                        ",0 " +                                    // GaroonTsuikaAtesakiDeleteFlag
                        ") ";

                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        private string Get_GyoumuKubunCD(string ID)
        {
            string cd = "";
            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                var dt = new System.Data.DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                  "GyoumuKubunCD " +
                  "FROM " + "Mst_GyoumuKubun " +
                  "WHERE GyoumuNarabijunCD = '" + ID + "'";

                //データ取得
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    cd = dt.Rows[0][0].ToString();
                }
            }

            return cd;
        }

        // 削除ボタン
        private void button12_Click(object sender, EventArgs e)
        {
            string methodName = ".btnDelete_Click";
            // エントリ君修正STEP2
            using (Popup_MessageBox dlg = new Popup_MessageBox("確認", GlobalMethod.GetMessage("I10605", ""), "案件番号"))
            {
                // if (MessageBox.Show(GlobalMethod.GetMessage("I10605", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    if (dlg.GetInputText().Equals(Header1.Text))
                    {
                        ErrorMessage.Text = "";
                        Boolean ErrorFLG = false;
                        if (item3_1_2.Checked == true)
                        {
                            set_error(GlobalMethod.GetMessage("E10601", ""));
                            ErrorFLG = true;
                        }
                        if (item3_1_1.SelectedValue != null && item3_1_1.SelectedValue.ToString() == "05")
                        {
                            // E10602:計画業務は削除できません。
                            set_error(GlobalMethod.GetMessage("E10602", ""));
                            ErrorFLG = true;
                        }

                        var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                        var dt = new System.Data.DataTable();
                        var dt2 = new System.Data.DataTable();
                        using (var conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            var cmd = conn.CreateCommand();
                            cmd.CommandText = "SELECT  " +
                                    "AnkenSaishinFlg " +
                                    ",AnkenSakuseiKubun " +
                                    ",AnkenJutakuBangou " +
                                    ",AnkenJutakuBangouEda " +

                                    //参照テーブル
                                    "FROM AnkenJouhou " +
                                    "WHERE AnkenJouhou.AnkenJouhouID =  " + AnkenID.ToString();
                            var sda = new SqlDataAdapter(cmd);
                            dt.Clear();
                            sda.Fill(dt);
                            if (dt.Rows.Count == 0)
                            {
                                // E10009:対象データは存在しません。
                                set_error(GlobalMethod.GetMessage("E10009", ""));
                                ErrorFLG = true;
                            }
                            conn.Close();
                        }
                        //えんとり君修正STEP2　単価契約　OR　窓口ミハルが存在するなら削除しない
                        if (!ErrorFLG) {
                            // 単価契約
                            DataTable tankaData = GlobalMethod.getData("AnkenJouhouID", "TankaKeiyakuID", "TankaKeiyaku", "AnkenJouhouID = " + AnkenID.ToString());
                            if(tankaData!= null && tankaData.Rows.Count > 0)
                            {
                                set_error(GlobalMethod.GetMessage("E10608", ""));
                                ErrorFLG = true;
                            }
                            else
                            {
                                // 窓口ミハル
                                DataTable mdData = GlobalMethod.getData("AnkenJouhouID", "MadoguchiID", "MadoguchiJouhou", "AnkenJouhouID = " + AnkenID.ToString());
                                if (mdData != null && mdData.Rows.Count > 0)
                                {
                                    set_error(GlobalMethod.GetMessage("E10609", ""));
                                    ErrorFLG = true;
                                }
                            }
                        }
                        if (!ErrorFLG)
                        {
                            using (var conn = new SqlConnection(connStr))
                            {
                                conn.Open();
                                SqlTransaction transaction = conn.BeginTransaction();
                                var cmd = conn.CreateCommand();
                                cmd.Transaction = transaction;

                                try
                                {
                                    cmd.CommandText = "UPDATE KokyakuKeiyakuJouhou SET KokyakuDeleteFlag = 1 " +
                                        "WHERE AnkenJouhouID =  " + AnkenID.ToString();
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "UPDATE GyoumuJouhou SET GyoumuDeleteFlag = 1 " +
                                        "WHERE AnkenJouhouID =  " + AnkenID.ToString();
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET KeiyakuDeleteFlag = 1 " +
                                        "WHERE AnkenJouhouID =  " + AnkenID.ToString();
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "UPDATE NyuusatsuJouhou SET NyuusatsuDeleteFlag = 1 " +
                                        "WHERE AnkenJouhouID =  " + AnkenID.ToString();
                                    cmd.ExecuteNonQuery();

                                    cmd.CommandText = "UPDATE AnkenJouhou SET AnkenDeleteFlag = 1 " +
                                        "WHERE AnkenJouhouID =  " + AnkenID.ToString();
                                    cmd.ExecuteNonQuery();

                                    // 最新フラグが0:最新ではない
                                    // 契約区分
                                    // 03:契約変更（黒伝）
                                    // 06:契約変更（黒伝・金額変更）
                                    // 07:契約変更（黒伝・工期変更）
                                    // 08:契約変更（黒伝・金額工期変更）
                                    // 09:契約変更（黒伝・その他）
                                    if (dt.Rows[0][0].ToString() == "1" && (dt.Rows[0][1].ToString() == "03"
                                        || dt.Rows[0][1].ToString() == "06" || dt.Rows[0][1].ToString() == "07"
                                        || dt.Rows[0][1].ToString() == "08" || dt.Rows[0][1].ToString() == "09"
                                        ))
                                    {
                                        cmd.CommandText = "SELECT " +
                                                " AnkenJouhouID " +
                                                ",AnkenSaishinFlg " +
                                                ",AnkenSakuseiKubun " +

                                                //参照テーブル
                                                "FROM AnkenJouhou " +
                                                //"WHERE AnkenJouhou.AnkenJutakuBangou = " + AnkenID.ToString() +
                                                //" AND AnkenJouhou.AnkenJutakuBangouEda =  " + AnkenID.ToString() +
                                                "WHERE AnkenJouhou.AnkenJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_7.Text + "' " +
                                                " AND AnkenJouhou.AnkenJutakuBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC =  N'" + item1_8.Text + "' " +
                                                " AND AnkenJouhou.AnkenSakuseiKubun IN ('01','03','06','07','08','09') " +
                                                " AND AnkenJouhou.AnkenJouhouID != '" + AnkenID + "' " + // 自分を除く
                                                " AND AnkenJouhou.AnkenDeleteFlag != 1 " + // 削除された案件を除く
                                                " ORDER BY AnkenJutakuBangou DESC, AnkenJouhouID DESC ";
                                        var sda = new SqlDataAdapter(cmd);
                                        dt2.Clear();
                                        sda.Fill(dt2);

                                        for (int i = 0; i < dt2.Rows.Count; i++)
                                        {
                                            if (i == 0)
                                            {
                                                cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                                    " AnkenSaishinFlg = 1 " +
                                                    "WHERE AnkenJouhouID =  " + dt2.Rows[i][0].ToString();
                                                cmd.ExecuteNonQuery();

                                                // 窓口のAnkenJouhouIDも同様に更新する
                                                cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                                                    "AnkenJouhouID = '" + dt2.Rows[i][0].ToString() + "' " +
                                                    ",MadoguchiAnkenJouhouID = '" + dt2.Rows[i][0].ToString() + "' " +
                                                    " WHERE MadoguchiJouhou.AnkenJouhouID = " + AnkenID;
                                                cmd.ExecuteNonQuery();

                                                // 単価契約のAnkenJouhouIDも同様に更新する
                                                cmd.CommandText = "UPDATE TankaKeiyaku SET " +
                                                    "AnkenJouhouID = '" + dt2.Rows[i][0].ToString() + "' " +
                                                    " WHERE TankaKeiyaku.AnkenJouhouID = " + AnkenID;
                                                cmd.ExecuteNonQuery();
                                            }
                                            else
                                            {
                                                cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                                    " AnkenSaishinFlg = 0 " +
                                                    "WHERE AnkenJouhouID =  " + dt2.Rows[i][0].ToString();
                                                cmd.ExecuteNonQuery();
                                            }
                                        }
                                        // 最新フラグを落とす
                                        cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                        " AnkenSaishinFlg = 0 " +
                                        "WHERE AnkenJouhouID = '" + AnkenID + "'";
                                        cmd.ExecuteNonQuery();


                                        DataTable dt3 = new DataTable();
                                        // 赤伝も削除する
                                        cmd.CommandText = "SELECT " +
                                                " AnkenJouhouID " +
                                                //参照テーブル
                                                "FROM AnkenJouhou " +
                                                "WHERE AnkenJouhou.AnkenJutakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_7.Text + "' " +
                                                " AND AnkenJouhou.AnkenJutakuBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC =  N'" + item1_8.Text + "' " +
                                                " AND AnkenJouhou.AnkenSakuseiKubun = '02' " +
                                                " AND AnkenJouhou.AnkenDeleteFlag != 1 " + // 削除された案件を除く
                                                " ORDER BY AnkenJutakuBangou DESC, AnkenJouhouID DESC ";
                                        sda = new SqlDataAdapter(cmd);
                                        dt3.Clear();
                                        sda.Fill(dt3);

                                        string akadenAnkenJouhouID = "";
                                        // 直近の02:契約変更（赤伝）を削除する
                                        //不具合No1331（1083）
                                        //赤伝を先に削除するとエラーになる。Count間違い
                                        //if (dt3 != null && dt3.Rows.Count >= 0)
                                        if (dt3 != null && dt3.Rows.Count > 0)
                                        {
                                            akadenAnkenJouhouID = dt3.Rows[0][0].ToString();

                                            cmd.CommandText = "UPDATE KokyakuKeiyakuJouhou SET KokyakuDeleteFlag = 1 " +
                                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                                            cmd.ExecuteNonQuery();

                                            cmd.CommandText = "UPDATE GyoumuJouhou SET GyoumuDeleteFlag = 1 " +
                                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                                            cmd.ExecuteNonQuery();

                                            cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET KeiyakuDeleteFlag = 1 " +
                                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                                            cmd.ExecuteNonQuery();

                                            cmd.CommandText = "UPDATE NyuusatsuJouhou SET NyuusatsuDeleteFlag = 1 " +
                                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                                            cmd.ExecuteNonQuery();

                                            cmd.CommandText = "UPDATE AnkenJouhou SET AnkenDeleteFlag = 1 " +
                                                "WHERE AnkenJouhouID =  " + akadenAnkenJouhouID.ToString();
                                            cmd.ExecuteNonQuery();
                                        }
                                    }
                                    // 計画の案件数更新
                                    if (item1_4.Text != "")
                                    {
                                        cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_4.Text + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_4.Text + "' ";
                                        Console.WriteLine(cmd.CommandText);
                                        cmd.ExecuteNonQuery();
                                    }
                                    if (beforeKeikakuBangou != "")
                                    {
                                        cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + beforeKeikakuBangou + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + beforeKeikakuBangou + "' ";
                                        Console.WriteLine(cmd.CommandText);
                                        cmd.ExecuteNonQuery();
                                    }

                                    transaction.Commit();

                                    // 234 削除から戻ってきたら、検索条件が初期化されている対応
                                    //Entry_Search form = new Entry_Search();
                                    //form.UserInfos = this.UserInfos;
                                    //form.Show();
                                    //this.Close();

                                    // 更新履歴の登録
                                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報を削除しました ID:" + AnkenID, pgmName + methodName, "");

                                    this.Owner.Show();
                                    this.Close();

                                }
                                catch (Exception)
                                {
                                    transaction.Rollback();
                                    throw;
                                }
                            }
                        }
                    }
                    else
                    {
                        set_error(GlobalMethod.GetMessage("E10009", ""));
                    }
                }
            }
        }

        private void c1FlexGrid3_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = ((C1FlexGrid)sender).HitTest(new System.Drawing.Point(e.X, e.Y));

            if (hti.Column == 2 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                Popup_ChousainList form = new Popup_ChousainList();
                //form.nendo = item3_1_5.SelectedValue.ToString();
                form.nendo = DateTime.Today.Year.ToString();
                form.Busho = BushoCD;
                form.ShowDialog();

                if (form.ReturnValue != null && form.ReturnValue[0] != null)
                {
                    c1FlexGrid3.Rows[_row][_col - 1] = form.ReturnValue[0];
                    c1FlexGrid3.Rows[_row][_col] = form.ReturnValue[1];
                    c1FlexGrid5.Rows[_row][_col - 1] = form.ReturnValue[0];
                    c1FlexGrid5.Rows[_row][_col] = form.ReturnValue[1];
                }
            }
            if (hti.Column == 0 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                if (MessageBox.Show("行を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    c1FlexGrid3.RemoveItem(_row);
                    c1FlexGrid5.RemoveItem(_row);
                    Resize_Grid("c1FlexGrid3");
                    Resize_Grid("c1FlexGrid5");
                }
            }
        }

        private void c1FlexGrid5_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {

        }

        private void label585_Click(object sender, EventArgs e)
        {
            Popup_Anken form = new Popup_Anken();
            form.mode = "kurikoshi";
            form.ShowDialog();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(System.Environment.CurrentDirectory + "/Resource/PDF/入札方式の説明.pdf");
        }

        // 戻るボタン
        private void button18_Click(object sender, EventArgs e)
        {
            //if (GlobalMethod.outputMessage("I00013", "") == DialogResult.OK)
            if (MessageBox.Show(GlobalMethod.GetMessage("I00013", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                if (mode == "change" && ChangeFlag == 1)
                {
                    Entry_Input form = new Entry_Input();
                    form.UserInfos = UserInfos;
                    form.KianFLG = true;
                    form.mode = "";
                    if (item3_1_20_kuroden.Text != "")
                    {
                        form.AnkenID = item3_1_20_kuroden.Text;
                    }
                    else if (item3_1_20_akaden.Text != "")
                    {
                        form.AnkenID = item3_1_20_akaden.Text;
                    }
                    else
                    {
                        form.AnkenID = AnkenID;
                    }
                    form.Show(this.Owner);
                    ownerflg = false;
                    this.Close();
                }
                else
                {
                    this.Owner.Show();
                    this.Close();
                }
            }
        }

        private void Entry_Input_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Owner.Visible == false && ownerflg)
            {
                this.Owner.Show();
                this.Owner.Close();
            }
        }

        private void label548_Click(object sender, EventArgs e)
        {
            ErrorMessage.Text = "";

            if (c1FlexGrid3.Rows.Count < 11)
            {
                c1FlexGrid3.AllowAddNew = true;
                Resize_Grid("c1FlexGrid3");
                c1FlexGrid3.Rows.Add();
                c1FlexGrid3.AllowAddNew = false;
                c1FlexGrid5.AllowAddNew = true;
                Resize_Grid("c1FlexGrid5");
                c1FlexGrid5.Rows.Add();
                c1FlexGrid5.AllowAddNew = false;
            }
            else
            {
                set_error(GlobalMethod.GetMessage("E10914", ""));
            }
        }

        private void label111_Click(object sender, EventArgs e)
        {
            if (c1FlexGrid4.Rows.Count < 14)
            {
                c1FlexGrid4.AllowAddNew = true;
                Resize_Grid("c1FlexGrid4");
                c1FlexGrid4.Rows.Add();
                // 追加した位置の行数取得
                int num = (c1FlexGrid4.Rows.Count - 2);
                // n回目、の n を取得
                int row_num = (c1FlexGrid4.Rows.Count - 3);
                c1FlexGrid4.Rows[num][0] = row_num + "回目";
                c1FlexGrid4.AllowAddNew = false;
            }
        }

        private void c1FlexGrid_AfterResizeRow(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            Resize_Grid(((C1FlexGrid)sender).Name);
        }

        public void Resize_Grid(string name)
        {
            Control[] cs;
            cs = this.Controls.Find(name, true);
            if (cs.Length > 0)
            {
                var fx = (C1FlexGrid)cs[0];
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

        // 契約タブの管理技術者のプロンプトボタン
        private void pictureBox6_Click_1(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //form.nendo = item3_1_5.SelectedValue.ToString();
            form.nendo = DateTime.Today.Year.ToString();
            form.Busho = BushoCD;
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item3_4_1.Text = form.ReturnValue[1];
                item3_4_1_CD.Text = form.ReturnValue[0];
                item4_1_2.Text = form.ReturnValue[1];
                //item3_4_5_Busho.Text = form.ReturnValue[2];
                //item3_4_5_Shibu.Text = form.ReturnValue[3];
                //item3_4_5_Ka.Text = form.ReturnValue[4];
            }
            item3_4_1.Focus();
        }
        //照査担当者プロンプト
        private void pictureBox8_Click(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //form.nendo = item3_1_5.SelectedValue.ToString();
            form.nendo = DateTime.Today.Year.ToString();
            form.Busho = BushoCD;
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item3_4_2.Text = form.ReturnValue[1];
                item3_4_2_CD.Text = form.ReturnValue[0];
                item4_1_4.Text = form.ReturnValue[1];
            }
            item3_4_2.Focus();
        }
        //審査担当者プロンプト
        private void pictureBox9_Click(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //form.nendo = item3_1_5.SelectedValue.ToString();
            form.nendo = DateTime.Today.Year.ToString();
            form.Busho = BushoCD;
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item3_4_3.Text = form.ReturnValue[1];
                item3_4_3_CD.Text = form.ReturnValue[0];
            }
            item3_4_3.Focus();
        }
        //業務管理者プロンプト
        private void pictureBox10_Click(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //form.nendo = item3_1_5.SelectedValue.ToString();
            form.nendo = DateTime.Today.Year.ToString();
            form.Busho = BushoCD;
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item3_4_4.Text = form.ReturnValue[1];
                item3_4_4_CD.Text = form.ReturnValue[0];
            }
            item3_4_4.Focus();
        }
        //契約タブ窓口担当者プロンプト
        private void pictureBox11_Click(object sender, EventArgs e)
        {
            Popup_ChousainList form = new Popup_ChousainList();
            //form.nendo = item3_1_5.SelectedValue.ToString();
            form.nendo = DateTime.Today.Year.ToString();
            form.Busho = BushoCD;
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item3_4_5.Text = form.ReturnValue[1];
                item3_4_5_CD.Text = form.ReturnValue[0];
                item3_4_5_Busho.Text = form.ReturnValue[2];
                item3_4_5_Shibu.Text = form.ReturnValue[3];
                item3_4_5_Ka.Text = form.ReturnValue[4];
            }
            item3_4_5.Focus();
        }

        private void ChousainList_Click2(object sender, EventArgs e)
        {

        }

        private void pictureBox14_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("氏名を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                item4_1_2.Text = "";
                item3_4_1.Text = "";
                item3_4_1_CD.Text = "";
            }
        }

        private void pictureBox15_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("氏名を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                item4_1_4.Text = "";
                item3_4_2.Text = "";
                item3_4_2_CD.Text = "";
            }
        }

        private void get_date()
        {
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT TOP 1 " +
                            //引合タブ
                            //引合状況
                            "AnkenHikiaijhokyo " +
                            //基本情報
                            ",AnkenSakuseiKubun " +
                            ",AnkenUriageNendo " +
                            ",AnkenKeikakuBangou " +
                            ",KeikakuAnkenMei " +
                            ",AnkenAnkenBangou " +
                            ",AnkenJutakuBangou " +
                            ",AnkenJutakuBangouEda " +
                            ",CASE AnkenTourokubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenTourokubi,'yyyy/MM/dd') END " +
                            ",AnkenJutakubushoCD " +
                            ",AnkenTantoushaMei " +
                            ",AnkenKeiyakusho " +
                            //案件情報
                            ",AnkenGyoumuMei " +
                            ",AnkenGyoumuKubun " +
                            ",AnkenNyuusatsuHoushiki " +
                            ",CASE AnkenNyuusatsuYoteibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenNyuusatsuYoteibi,'yyyy/MM/dd') END " +
                            ",NyuusatsuRakusatsushaID " +
                            ",NyuusatsuGyoumuBikou " +
                            //発注者情報
                            ",AnkenHachushaCD " +
                            ",HachushaKubun1Mei " +
                            ",HachushaKubun2Mei " +
                            ",TodouhukenMei " +
                            ",HachushaMei " +
                            ",AnkenHachushaKaMei " +
                            //発注担当者情報
                            ",AnkenHachuushaKeiyakuBusho " +
                            ",AnkenHachuushaKeiyakuTantou " +
                            ",AnkenHachuushaKeiyakuTEL " +
                            ",AnkenHachuushaKeiyakuFAX " +
                            ",AnkenHachuushaKeiyakuMail " +
                            ",AnkenHachuushaKeiyakuYuubin " +
                            ",AnkenHachuushaKeiyakuJuusho " +
                            ",AnkenHachuuDaihyouYakushoku " +
                            ",AnkenHachuuDaihyousha " +
                            //当会対応
                            ",AnkenToukaiSankouMitsumori " +
                            ",AnkenToukaiJyutyuIyoku " +
                            ",ISNULL(NyuusatsuMitsumorigaku,0) " +
                            //業務内容
                            ",ISNULL(GyoumuChosaBuRitsu,0) " +
                            ",ISNULL(GyoumuJigyoFukyuBuRitsu,0) " +
                            ",ISNULL(GyoumuJyohouSystemBuRitsu,0) " +
                            ",ISNULL(GyoumuSougouKenkyuJoRitsu,0) " +
                            ",ISNULL(GyoumuShizaiChousaRitsu,0) " +
                            ",ISNULL(GyoumuEizenRitsu,0) " +
                            ",ISNULL(GyoumuKikiruiChousaRitsu,0) " +
                            ",ISNULL(GyoumuKoujiChousahiRitsu,0) " +
                            ",ISNULL(GyoumuSanpaiFukusanbutsuRitsu,0) " +
                            ",ISNULL(GyoumuHokakeChousaRitsu,0) " +
                            ",ISNULL(GyoumuShokeihiChousaRitsu,0) " +
                            ",ISNULL(GyoumuGenkaBunsekiRitsu,0) " +
                            ",ISNULL(GyoumuKijunsakuseiRitsu,0) " +
                            ",ISNULL(GyoumuKoukyouRoumuhiRitsu,0) " +
                            ",ISNULL(GyoumuRoumuhiKoukyouigaiRitsu,0) " +
                            ",ISNULL(GyoumuSonotaChousabuRitsu,0) " +
                            //処理用データ
                            ",AnkenKaisuu" +
                            ",AnkenSaishinFlg " +

                            // 不足データ取得
                            ",AnkenTantoushaCD " +
                            ",Mst_Chousain.GyoumuBushoCD " +

                            // 398 工期開始年度対応
                            ",AnkenKoukiNendo " + // 56:工期開始年度

                            // えんとり君修正STEP2(INDEX:57～) 
                            ",ISNULL(mc.ChousainMei,'') " +
                            ",ISNULL(AnkenFolderHenkouTantoushaCD,'') " +
                            ",CASE WHEN AnkenFolderHenkouDatetime IS NULL THEN '' ELSE FORMAT(AnkenFolderHenkouDatetime,'yyyy/MM/dd HH:mm:ss') END " +
                            // えんとり君修正STEP2(INDEX:60～) 
                            ",ISNULL(mb.JigyoubuHeadCD,'') " +
                            // No.1422 1196 案件番号の変更履歴を保存する
                            ",ISNULL(AnkenHenkoumaeAnkenBangou,'') " + 

                    //参照テーブル
                    "FROM AnkenJouhou " +
                            "LEFT JOIN Mst_SakuseiKubun ON AnkenSakuseiKubun = SakuseiKubunID " +
                            "LEFT JOIN Mst_Busho mb ON AnkenJutakubushoCD = GyoumuBushoCD " +
                            "LEFT JOIN Mst_KeiyakuKeitai ON AnkenNyuusatsuHoushiki = KeiyakuKeitaiCD " +
                            "LEFT JOIN Mst_Hachusha ON AnkenHachushaCD = HachushaCD " +
                            "LEFT JOIN Mst_HachushaKubun1 ON Mst_HachushaKubun1.HachushaKubun1CD = Mst_Hachusha.HachushaKubun1CD " +
                            "LEFT JOIN Mst_HachushaKubun2 ON Mst_HachushaKubun2.HachushaKubun2CD = Mst_Hachusha.HachushaKubun2CD " +
                            "LEFT JOIN Mst_Todouhuken ON Mst_Todouhuken.TodouhukenCD = Mst_Hachusha.TodouhukenCD " +
                            "LEFT JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                            "LEFT JOIN KeiyakuJouhouEntory ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID " +
                            "LEFT JOIN KeikakuJouhou ON AnkenJouhou.AnkenKeikakuBangou = KeikakuJouhou.KeikakuBangou " +
                            "LEFT JOIN GyoumuHaibun ON AnkenJouhou.AnkenJouhouID = GyoumuHaibun.GyoumuAnkenJouhouID AND GyoumuHibunKubun = '10' " +
                            "LEFT JOIN Mst_Chousain ON AnkenJouhou.AnkenTantoushaCD = Mst_Chousain.KojinCD " +
                            "LEFT JOIN Mst_Chousain mc ON AnkenJouhou.AnkenFolderHenkouTantoushaCD = mc.KojinCD " +
                            "WHERE AnkenJouhou.AnkenJouhouID =  " + AnkenID.ToString();
                    var sda = new SqlDataAdapter(cmd);
                    AnkenData_H.Clear();
                    sda.Fill(AnkenData_H);
                }

                AnkenData_Grid1.Clear();
                // この案件を元にコピー
                // エントリ君修正STEP1
                //if (mode == "insert")
                if (mode == "insert" || isKeikakuAnkenNew == true)
                {
                    using (var conn = new SqlConnection(connStr))
                    {
                        string toukai = GlobalMethod.GetCommonValue2("ENTORY_TOUKAI");

                        var cmd = conn.CreateCommand();
                        cmd.CommandText = "SELECT  " +
                                //引合タブ
                                //過去案件
                                "AnkenZenkaiAnkenJouhouID " +
                                ",AnkenZenkaiAnkenBangou " +
                                ",AnkenZenkaiJutakuBangou " +
                                ",AnkenZenkaiJutakuEdaban " +
                                ",AnkenZenkaiGyoumuMei " +
                                ",AnkenZenkaiRakusatsusha " +
                                ",AnkenZenkaiRakusatsushaID " +

                                ",ISNULL(AnkenZenkaiJutakuKingaku,0) AS 'AnkenZenkaiJutakuKingaku' " +
                                ",ISNULL(NyuusatsuOusatsuKingaku,0) AS 'NyuusatsuOusatugaku' " + // 入札応札額
                                ",ISNULL(NyuusatsuMitsumorigaku,0) AS 'NyuusatsuMitsumorigaku' " +
                                //",ISNULL(KeiyakuZeikomiKingaku,0) AS 'KeiyakuZeikomiKingaku' " +
                                ",ISNULL(KeiyakuKeiyakuKingaku,0) AS 'KeiyakuKeiyakuKingaku' " + // 契約金額 税抜
                                                                                                 //",ISNULL(AnkenZenkaiJutakuZeinuki,0) AS 'AnkenZenkaiJutakuZeinuki' " +
                                ",ISNULL(KeiyakuHaibunZeinukiKei,0) AS 'KeiyakuHaibunZeinukiKei' " + // 前回受託金額（税抜）
                                ",KeiyakuZenkaiRakusatsushaID " +
                                ",AnkenZenkaiKyougouKigyouCD " +

                                ",AnkenZenkaiRakusatsuID " + // 14:前回落札ID データ追い出しで使用
                                ",AnkenUriageNendo + '_' + CONVERT(NVARCHAR, AnkenNyuusatsuYoteibi,111) + '_' + CONVERT(NVARCHAR, AnkenJouhou.AnkenJouhouID) AS sortKey " + // 15:SotKey
                                                                                                                                                                            //参照テーブル
                                "FROM AnkenJouhouZenkaiRakusatsu " +
                                "LEFT JOIN NyuusatsuJouhou ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                                "LEFT JOIN KeiyakuJouhouEntory ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID " +
                                "LEFT JOIN AnkenJouhou ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = AnkenJouhou.AnkenJouhouID " +
                                "LEFT JOIN (select NyuusatsuJouhouID,min(NyuusatsuOusatsuKingaku) AS NyuusatsuOusatsuKingaku FROM NyuusatsuJouhouOusatsusha where NyuusatsuOusatsusha COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + toukai + "' group by NyuusatsuJouhouID) T1 " +
                                "  ON T1.NyuusatsuJouhouID = AnkenJouhou.AnkenJouhouID " +
                                "WHERE AnkenJouhouZenkaiRakusatsu.AnkenJouhouID =  " + AnkenID.ToString() + " " +
                                // 降順
                                "ORDER BY AnkenJouhou.AnkenUriageNendo desc, AnkenNyuusatsuYoteibi desc,AnkenJouhouZenkaiRakusatsu.AnkenJouhouID desc";
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(AnkenData_Grid1);

                        // 前回受託番号ID
                        string AnkenZenkaiRakusatsuID = "1";
                        int num = 0;
                        int maxNum = 1;
                        if (AnkenData_Grid1.Rows.Count > 0)
                        {
                            // ヘッダーを除いて回し、前回受託番号IDの最大値を取得する
                            for (int i = 0; i < AnkenData_Grid1.Rows.Count; i++)
                            {
                                AnkenZenkaiRakusatsuID = AnkenData_Grid1.Rows[i][14].ToString();
                                if (int.TryParse(AnkenZenkaiRakusatsuID, out num))
                                {
                                    if (maxNum < num)
                                    {
                                        maxNum = num;
                                    }
                                }
                            }
                            // 最大値 + 1
                            maxNum = maxNum + 1;
                        }

                        cmd.CommandText = "SELECT  " +
                                //引合タブ
                                //過去案件
                                "AnkenJouhou.AnkenJouhouID AS 'AnkenZenkaiAnkenJouhouID' " +
                                ",AnkenAnkenBangou AS 'AnkenZenkaiAnkenBangou' " +
                                ",AnkenJutakuBangou AS 'AnkenZenkaiJutakuBangou' " +
                                ",AnkenJutakuBangouEda AS 'AnkenZenkaiJutakuEdaban' " +
                                ",AnkenGyoumuMei AS 'AnkenZenkaiGyoumuMei' " +
                                ",NyuusatsuRakusatsusha AS 'AnkenZenkaiRakusatsusha' " +
                                ",NyuusatsuRakusatsushaID AS 'AnkenZenkaiRakusatsushaID' " +

                                ",ISNULL(NyuusatsuRakusatugaku,0) AS 'AnkenZenkaiJutakuKingaku' " +
                                //",NyuusatsuOusatugaku AS 'NyuusatsuOusatugaku' " + // 応札額
                                ",ISNULL(NyuusatsuOusatsuKingaku,0) AS 'NyuusatsuOusatugaku' " + // 入札応札額
                                ",NyuusatsuMitsumorigaku AS 'NyuusatsuMitsumorigaku' " +
                                //",ISNULL(KeiyakuZeikomiKingaku,0) AS 'KeiyakuZeikomiKingaku' " +
                                ",ISNULL(KeiyakuKeiyakuKingaku,0) AS 'KeiyakuKeiyakuKingaku' " + // 契約金額 税抜
                                                                                                 //",ISNULL(Keiyakukeiyakukingakukei,0) AS 'AnkenZenkaiJutakuZeinuki' " +
                                                                                                 // 前回受託金額は、受託金額（税込）から消費税分を引いた値
                                ",ISNULL(KeiyakuHaibunZeinukiKei,0) AS 'KeiyakuHaibunZeinukiKei' " + // 前回受託金額（税抜）
                                ",NyuusatsuKyougouTashaID AS 'KeiyakuZenkaiRakusatsushaID' " +
                                ",KyougouKigyouCD AS 'AnkenZenkaiKyougouKigyouCD' " +
                                //",0 AS Dummy " +
                                "," + maxNum + " AS AnkenZenkaiRakusatsuID " + // 前回受託番号ID
                                ",AnkenUriageNendo + '_' + CONVERT(NVARCHAR, AnkenNyuusatsuYoteibi,111) + '_' + CONVERT(NVARCHAR, AnkenJouhou.AnkenJouhouID) AS sortKey " + // 15:SotKey

                                //参照テーブル
                                "FROM AnkenJouhou " +
                                "LEFT JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                                "LEFT JOIN KeiyakuJouhouEntory ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID " +
                                "LEFT JOIN Mst_KyougouTasha ON Mst_KyougouTasha.KyougouTashaID = NyuusatsuJouhou.NyuusatsuKyougouTashaID " +
                                "LEFT JOIN (select NyuusatsuJouhouID,min(NyuusatsuOusatsuKingaku) AS NyuusatsuOusatsuKingaku FROM NyuusatsuJouhouOusatsusha where NyuusatsuOusatsusha COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + toukai + "' group by NyuusatsuJouhouID) T1 " +
                                "  ON T1.NyuusatsuJouhouID = AnkenJouhou.AnkenJouhouID " +

                                "WHERE AnkenJouhou.AnkenJouhouID =  " + AnkenID.ToString();
                        sda = new SqlDataAdapter(cmd);
                        //AnkenData_Grid1.Clear();
                        sda.Fill(AnkenData_Grid1);

                        // 取得した過去案件が5件を超えている場合、追い出しを行う
                        if (AnkenData_Grid1.Rows.Count > 5)
                        {
                            // 案件前回ID（AnkenZenkaiRakusatsuID）が一番古い（低い）データが追い出し対象
                            cmd.CommandText = "SELECT TOP 1 " +
                                "AnkenZenkaiRakusatsuID " +
                                //参照テーブル
                                "FROM AnkenJouhouZenkaiRakusatsu " +
                                "WHERE AnkenJouhouZenkaiRakusatsu.AnkenJouhouID =  " + AnkenID.ToString() + " " +
                                // 降順
                                "ORDER BY AnkenZenkaiRakusatsuID";
                            var zenkaiSda = new SqlDataAdapter(cmd);
                            DataTable zenkaiRakusatsuDT = new DataTable();
                            zenkaiSda.Fill(zenkaiRakusatsuDT);

                            string zenkaiRakusatsuID = "";
                            if (zenkaiRakusatsuDT != null && zenkaiRakusatsuDT.Rows.Count > 0)
                            {
                                zenkaiRakusatsuID = zenkaiRakusatsuDT.Rows[0][0].ToString();
                            }

                            // AnkenData_Grid1の最終行はコピー元データなので除外（-1）して回す
                            for (int i = 0; i < AnkenData_Grid1.Rows.Count - 1; i++)
                            {
                                if (zenkaiRakusatsuID.Equals(AnkenData_Grid1.Rows[i][14].ToString()))
                                {
                                    // 追い出し対象をDataTableから除外
                                    AnkenData_Grid1.Rows.RemoveAt(i);
                                    break;
                                }
                            }
                            // sortKeyでソート
                            DataRow[] selectedRows = AnkenData_Grid1.Select("", "sortKey");

                            AnkenData_Grid1 = new DataTable();

                            // DataRowからDataTableに変換
                            AnkenData_Grid1 = selectedRows.CopyToDataTable();
                        }

                    }
                }
                // 新規以外
                else
                {
                    using (var conn = new SqlConnection(connStr))
                    {
                        string toukai = GlobalMethod.GetCommonValue2("ENTORY_TOUKAI");

                        var cmd = conn.CreateCommand();
                        cmd.CommandText = "SELECT  " +
                                //引合タブ
                                //過去案件
                                "AnkenZenkaiAnkenJouhouID " +
                                ",AnkenZenkaiAnkenBangou " +
                                ",AnkenZenkaiJutakuBangou " +
                                ",AnkenZenkaiJutakuEdaban " +
                                ",AnkenZenkaiGyoumuMei " +
                                ",AnkenZenkaiRakusatsusha " +
                                ",AnkenZenkaiRakusatsushaID " +

                                ",ISNULL(AnkenZenkaiJutakuKingaku,0) AS 'AnkenZenkaiJutakuKingaku' " +
                                //",ISNULL(NyuusatsuOusatugaku,0) AS 'NyuusatsuOusatugaku' " +
                                ",ISNULL(NyuusatsuOusatsuKingaku,0) AS 'NyuusatsuOusatugaku' " + // 入札応札額
                                ",ISNULL(NyuusatsuMitsumorigaku,0) AS 'NyuusatsuMitsumorigaku' " +
                                //",ISNULL(KeiyakuZeikomiKingaku,0) AS 'KeiyakuZeikomiKingaku' " +
                                ",ISNULL(KeiyakuKeiyakuKingaku,0) AS 'KeiyakuKeiyakuKingaku' " + // 契約金額 税抜
                                ",ISNULL(AnkenZenkaiJutakuZeinuki,0) AS 'AnkenZenkaiJutakuZeinuki' " +
                                ",KeiyakuZenkaiRakusatsushaID " +
                                ",AnkenZenkaiKyougouKigyouCD " +
                                //",0 AS Dummy " +
                                ",AnkenZenkaiRakusatsuID " +  // 前回受託番号ID
                                ",AnkenUriageNendo + '_' + CONVERT(NVARCHAR, AnkenNyuusatsuYoteibi,111) + '_' + CONVERT(NVARCHAR, AnkenJouhou.AnkenJouhouID) AS sortKey " + // 15:SotKey

                                //参照テーブル
                                "FROM AnkenJouhouZenkaiRakusatsu " +
                                "LEFT JOIN NyuusatsuJouhou ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                                "LEFT JOIN KeiyakuJouhouEntory ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID " +
                                "LEFT JOIN AnkenJouhou ON AnkenJouhouZenkaiRakusatsu.AnkenZenkaiAnkenJouhouID = AnkenJouhou.AnkenJouhouID " +
                                "LEFT JOIN (select NyuusatsuJouhouID,min(NyuusatsuOusatsuKingaku) AS NyuusatsuOusatsuKingaku FROM NyuusatsuJouhouOusatsusha where NyuusatsuOusatsusha = N'" + toukai + "' group by NyuusatsuJouhouID) T1 " +
                                "  ON T1.NyuusatsuJouhouID = AnkenJouhou.AnkenJouhouID " +
                                "WHERE AnkenJouhouZenkaiRakusatsu.AnkenJouhouID =  " + AnkenID.ToString() + " " +
                                "ORDER BY AnkenJouhou.AnkenUriageNendo, AnkenNyuusatsuYoteibi,AnkenJouhouZenkaiRakusatsu.AnkenJouhouID ";
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(AnkenData_Grid1);
                    }
                }

                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT TOP 1 " +
                            //入札タブ
                            "NyuusatsuRakusatsushaID " +
                            ",CASE NyuusatsuRakusatsuKekkaDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuRakusatsuKekkaDate,'yyyy/MM/dd') END AS kekkaDate " +
                            ",CASE AnkenNyuusatsuYoteibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenNyuusatsuYoteibi,'yyyy/MM/dd') END AS yoteibi " +//入札予定日予定
                            ",AnkenHikiaijhokyo " +
                            ",AnkenSakuseiKubun " +
                            ",NyuusatsuGyoumuBikou " +
                            ",AnkenToukaiOusatu " +
                            ",AnkenToukaiSankouMitsumori " +
                            ",AnkenToukaiJyutyuIyoku " +
                            ",ISNULL(NyuusatsuMitsumorigaku,0) " +
                            ",NyuusatsuRakusatsuShaJokyou " +
                            ",NyuusatsuRakusatsuGakuJokyou " +
                            //",CASE NyuusatsuCreateDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuCreateDate,'yyyy/MM/dd') END " +
                            //",CASE NyuusatsuUpdateDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuUpdateDate,'yyyy/MM/dd') END " +
                            ",CASE NyuusatsuRakusatsuShokaiDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuRakusatsuShokaiDate,'yyyy/MM/dd') END AS syokaiDate " +
                            ",CASE NyuusatsuRakusatsuSaisyuDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(NyuusatsuRakusatsuSaisyuDate,'yyyy/MM/dd') END AS saisyuDate" +
                            ",ISNULL(NyuusatsuYoteiKakaku,0) " +
                            ",ISNULL(NyuusatsushaSuu,0) " +
                            ",NyuusatsuRakusatsusha " +
                            ",ISNULL(NyuusatsuRakusatugaku,0) " +
                            ",NyuusatsuOusatsusha " +
                            ",NyuusatsuOusatsuKingaku " +
                            ",NyuusatsuKekkaMemo " +
                            //業務内容
                            ",ISNULL(GyoumuChosaBuRitsu,0) " +
                            ",ISNULL(GyoumuJigyoFukyuBuRitsu,0) " +
                            ",ISNULL(GyoumuJyohouSystemBuRitsu,0) " +
                            ",ISNULL(GyoumuSougouKenkyuJoRitsu,0) " +
                            ",ISNULL(GyoumuShizaiChousaRitsu,0) " +
                            ",ISNULL(GyoumuEizenRitsu,0) " +
                            ",ISNULL(GyoumuKikiruiChousaRitsu,0) " +
                            ",ISNULL(GyoumuKoujiChousahiRitsu,0) " +
                            ",ISNULL(GyoumuSanpaiFukusanbutsuRitsu,0) " +
                            ",ISNULL(GyoumuHokakeChousaRitsu,0) " +
                            ",ISNULL(GyoumuShokeihiChousaRitsu,0) " +
                            ",ISNULL(GyoumuGenkaBunsekiRitsu,0) " +
                            ",ISNULL(GyoumuKijunsakuseiRitsu,0) " +
                            ",ISNULL(GyoumuKoukyouRoumuhiRitsu,0) " +
                            ",ISNULL(GyoumuRoumuhiKoukyouigaiRitsu,0) " +
                            ",ISNULL(GyoumuSonotaChousabuRitsu,0) " +
                            ",NyuusatsuJouhou.NyuusatsuJouhouID " +
                            ",NyuusatsuUpdateDate " +


                            //参照テーブル
                            "FROM AnkenJouhou " +
                            "LEFT JOIN Mst_SakuseiKubun ON AnkenSakuseiKubun = SakuseiKubunID " +
                            "LEFT JOIN Mst_Busho ON AnkenJutakubushoCD = GyoumuBushoCD " +
                            "LEFT JOIN Mst_KeiyakuKeitai ON AnkenNyuusatsuHoushiki = KeiyakuKeitaiCD " +
                            "LEFT JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                            "LEFT JOIN NyuusatsuJouhouOusatsusha ON NyuusatsuJouhou.NyuusatsuJouhouID =  NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID " +
                            "LEFT JOIN KeiyakuJouhouEntory ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID " +
                            "LEFT JOIN GyoumuHaibun ON AnkenJouhou.AnkenJouhouID = GyoumuHaibun.GyoumuAnkenJouhouID AND GyoumuHibunKubun = '20'" +
                            "WHERE AnkenJouhou.AnkenJouhouID =  " + AnkenID.ToString();
                    var sda = new SqlDataAdapter(cmd);
                    AnkenData_N.Clear();
                    sda.Fill(AnkenData_N);
                }

                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT  " +
                            //入札タブ
                            //入札参加者
                            "NyuusatsuRakusatsuJyuni " +//落札順位
                            ",NyuusatsuRakusatsuJokyou " +//落札状況
                            ",NyuusatsuOusatsushaID " +
                            ",NyuusatsuOusatsusha " +
                            ",NyuusatsuOusatsuKingaku " +
                            ",NyuusatsuRakusatsuComment " +//コメント
                            ",NyuusatsuOusatsuKyougouKigyouCD " +

                            //参照テーブル
                            "FROM AnkenJouhou " +
                            "LEFT JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                            "LEFT JOIN NyuusatsuJouhouOusatsusha ON NyuusatsuJouhou.NyuusatsuJouhouID =  NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID " +
                            "WHERE AnkenJouhou.AnkenJouhouID =  " + AnkenID.ToString();
                    var sda = new SqlDataAdapter(cmd);
                    AnkenData_Grid2.Clear();
                    sda.Fill(AnkenData_Grid2);

                    StringBuilder sb = new StringBuilder();
                    for (int i = 0; i < AnkenData_Grid2.Rows.Count; i++)
                    {
                        sb.Append(AnkenData_Grid2.Rows[i][0]);
                        sb.Append(",");
                        sb.Append(AnkenData_Grid2.Rows[i][1]);
                        sb.Append(",");
                        sb.Append(AnkenData_Grid2.Rows[i][2]);
                        sb.Append(",");
                        sb.Append(AnkenData_Grid2.Rows[i][3]);
                        sb.Append(",");
                        sb.Append(AnkenData_Grid2.Rows[i][4]);
                        sb.Append(",");
                        sb.Append(AnkenData_Grid2.Rows[i][5]);
                        sb.Append(",");
                        sb.Append(AnkenData_Grid2.Rows[i][6]);
                        sb.Append(",");
                    }
                    // 入札タブの入札者情報が変更されたかの判断で使用する
                    c1FlexGrid2Data = sb.ToString();

                }
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT TOP 1 " +
                            //契約タブ
                            //契約情報
                            "AnkenSakuseiKubun " +
                            ",AnkenKianZumi " +
                            ",CASE KeiyakuKeiyakuTeiketsubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuKeiyakuTeiketsubi,'yyyy/MM/dd') END " +
                            ",CASE KeiyakuSakuseibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSakuseibi,'yyyy/MM/dd') END " +
                            ",AnkenUriageNendo " +
                            ",CASE AnkenKeiyakuKoukiKaishibi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenKeiyakuKoukiKaishibi,'yyyy/MM/dd') END " +//契約工期自
                            ",CASE AnkenKeiyakuKoukiKanryoubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(AnkenKeiyakuKoukiKanryoubi,'yyyy/MM/dd') END " +//契約工期至
                            ",KeiyakuGyoumuKubun " +
                            ",AnkenHachuushaKaMei " +
                            ",KeiyakuShouhizeiritsu " +
                            ",KeiyakuGyoumuMei " +
                            ",ISNULL(KeiyakuKeiyakuKingaku,0) " +
                            ",ISNULL(KeiyakuZeikomiKingaku,0) " +
                            ",ISNULL(KeiyakuuchizeiKingaku,0) " +
                            ",ISNULL(Keiyakukeiyakukingakukei,0) " +
                            ",ISNULL(KeiyakuBetsuKeiyakuKingaku,0) " +
                            ",KeiyakuHenkouChuushiRiyuu " +
                            ",NyuusatsuGyoumuBikou " +
                            ",KeiyakuBikou " +
                            ",KeiyakuShosha " +
                            ",KeiyakuTokkiShiyousho " +
                            ",KeiyakuMitsumorisho " +
                            ",KeiyakuTanpinChousaMitsumorisho " +
                            ",KeiyakuSonota " +
                            ",KeiyakuSonotaNaiyou " +
                            ",AnkenKeiyakusho " +
                            //配分情報
                            ",ISNULL(KeiyakuUriageHaibunCho,0) " +
                            ",ISNULL(KeiyakuUriageHaibunJo,0) " +
                            ",ISNULL(KeiyakuUriageHaibunJosys,0) " +
                            ",ISNULL(KeiyakuUriageHaibunKei,0) " +
                            ",ISNULL(KeiyakuHaibunChoZeinuki,0) " +
                            ",ISNULL(KeiyakuHaibunJoZeinuki,0) " +
                            ",ISNULL(KeiyakuHaibunJosysZeinuki,0) " +
                            ",ISNULL(KeiyakuHaibunKeiZeinuki,0) " +
                            //単契等の見込み補正額
                            ",ISNULL(KeiyakuTankeiMikomiCho,0) " +
                            ",ISNULL(KeiyakuTankeiMikomiJo,0) " +
                            ",ISNULL(KeiyakuTankeiMikomiJosys,0) " +
                            ",ISNULL(KeiyakuTankeiMikomiKei,0) " +
                            //年度繰越額
                            ",ISNULL(KeiyakuKurikoshiCho,0) " +
                            ",ISNULL(KeiyakuKurikoshiJo,0) " +
                            ",ISNULL(KeiyakuKurikoshiJosys,0) " +
                            ",ISNULL(KeiyakuKurikoshiKei,0) " +
                            //管理者・担当者
                            ",KanriGijutsushaCD " +
                            ",KanriGijutsushaNM " +
                            ",ShousaTantoushaCD " +
                            ",ShousaTantoushaNM " +
                            ",SinsaTantoushaCD " +
                            ",SinsaTantoushaNM " +
                            ",GyoumuKanrishaCD " +
                            ",GyoumuKanrishaMei " +
                            ",GyoumuJouhouMadoKojinCD " +
                            ",GyoumuJouhouMadoChousainMei " +
                            ",GyoumuJouhouMadoGyoumuBushoCD " +
                            ",GyoumuJouhouMadoShibuMei " +
                            ",GyoumuJouhouMadoKamei " +
                            //請求書情報
                            ",CASE KeiyakuSeikyuubi1 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi1,'yyyy/MM/dd') END " +
                            ",ISNULL(KeiyakuSeikyuuKingaku1,0) " +
                            ",CASE KeiyakuSeikyuubi2 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi2,'yyyy/MM/dd') END " +
                            ",ISNULL(KeiyakuSeikyuuKingaku2,0) " +
                            ",CASE KeiyakuSeikyuubi3 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi3,'yyyy/MM/dd') END " +
                            ",ISNULL(KeiyakuSeikyuuKingaku3,0) " +
                            ",CASE KeiyakuSeikyuubi4 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi4,'yyyy/MM/dd') END " +
                            ",ISNULL(KeiyakuSeikyuuKingaku4,0) " +
                            ",CASE KeiyakuSeikyuubi5 WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KeiyakuSeikyuubi5,'yyyy/MM/dd') END " +
                            ",ISNULL(KeiyakuSeikyuuKingaku5,0) " +
                            ",CASE KeiyakuZentokinUkewatashibi WHEN '1753/01/01' THEN null ELSE FORMAT(KeiyakuZentokinUkewatashibi,'yyyy/MM/dd') END " +
                            ",ISNULL(KeiyakuZentokin,0) " +
                            //業務内容
                            ",ISNULL(KeiyakuHaibunChoZeinuki,0) " +
                            ",ISNULL(KeiyakuHaibunJoZeinuki,0) " +
                            ",ISNULL(KeiyakuHaibunJosysZeinuki,0) " +
                            ",ISNULL(KeiyakuHaibunKeiZeinuki,0) " +
                            ",ISNULL(GyoumuShizaiChousaRitsu,0) " +
                            ",ISNULL(GyoumuEizenRitsu,0) " +
                            ",ISNULL(GyoumuKikiruiChousaRitsu,0) " +
                            ",ISNULL(GyoumuKoujiChousahiRitsu,0) " +
                            ",ISNULL(GyoumuSanpaiFukusanbutsuRitsu ,0)" +
                            ",ISNULL(GyoumuHokakeChousaRitsu,0) " +
                            ",ISNULL(GyoumuShokeihiChousaRitsu,0) " +
                            ",ISNULL(GyoumuGenkaBunsekiRitsu,0) " +
                            ",ISNULL(GyoumuKijunsakuseiRitsu,0) " +
                            ",ISNULL(GyoumuKoukyouRoumuhiRitsu,0) " +
                            ",ISNULL(GyoumuRoumuhiKoukyouigaiRitsu,0) " +
                            ",ISNULL(GyoumuSonotaChousabuRitsu,0) " +
                            ",ISNULL(GyoumuShizaiChousaGaku,0) " +
                            ",ISNULL(GyoumuEizenGaku,0) " +
                            ",ISNULL(GyoumuKikiruiChousaGaku,0) " +
                            ",ISNULL(GyoumuKoujiChousahiGaku,0) " +
                            ",ISNULL(GyoumuSanpaiFukusanbutsuGaku,0) " +
                            ",ISNULL(GyoumuHokakeChousaGaku,0) " +
                            ",ISNULL(GyoumuShokeihiChousaGaku,0) " +
                            ",ISNULL(GyoumuGenkaBunsekiGaku,0) " +
                            ",ISNULL(GyoumuKijunsakuseiGaku,0) " +
                            ",ISNULL(GyoumuKoukyouRoumuhiGaku,0) " +
                            ",ISNULL(GyoumuRoumuhiKoukyouigaiGaku,0) " +
                            ",ISNULL(GyoumuSonotaChousabuGaku,0) " +

                            // 不足分取得
                            ",AnkenGyoumuMei " +

                            //えんとり君修正STEP2（RIBC項目追加）
                            ",KeiyakuRIBCYouTankaDataMoushikomisho " +
                            ",KeiyakuSashaKeiyu " +
                            ",KeiyakuRIBCYouTankaData " +

                    //参照テーブル
                    "FROM AnkenJouhou " +
                            "LEFT JOIN Mst_SakuseiKubun ON AnkenSakuseiKubun = SakuseiKubunID " +
                            "LEFT JOIN Mst_Busho ON AnkenJutakubushoCD = GyoumuBushoCD " +
                            "LEFT JOIN Mst_KeiyakuKeitai ON AnkenNyuusatsuHoushiki = KeiyakuKeitaiCD " +
                            "LEFT JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                            "LEFT JOIN KeiyakuJouhouEntory ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID " +
                            "LEFT JOIN GyoumuHaibun ON AnkenJouhou.AnkenJouhouID = GyoumuHaibun.GyoumuAnkenJouhouID AND GyoumuHibunKubun = '30'" +
                            "LEFT JOIN GyoumuJouhou ON AnkenJouhou.AnkenJouhouID = GyoumuJouhou.AnkenJouhouID " +
                            "LEFT JOIN GyoumuJouhouMadoguchi ON GyoumuJouhouMadoguchi.GyoumuJouhouID = GyoumuJouhou.GyoumuJouhouID " +
                            "WHERE AnkenJouhou.AnkenJouhouID =  " + AnkenID.ToString();
                    var sda = new SqlDataAdapter(cmd);
                    Console.WriteLine(cmd.CommandText);
                    AnkenData_K.Clear();
                    sda.Fill(AnkenData_K);
                }
                //売上請求情報
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT  " +
                            "RibcNo " +
                            ",RibcKoukiEnd " +
                            ",RibcUriageKeijyoTuki " +
                            ",ISNULL(RibcSeikyuKingaku,0) " +
                            ",JigyoubuHeadCD " +
                            ",RibcKoukiStart " +
                            ",RibcNouhinbi " +
                            ",RibcSeikyubi " +
                            ",RibcNyukinyoteibi " +
                            ",RibcKubun " +

                            //参照テーブル
                            "FROM RibcJouhou " +
                            "LEFT JOIN Mst_Busho ON RibcKankeibusho = GyoumuBushoCD " +
                            " WHERE RibcID = " + AnkenID +
                            " ORDER BY RibcKankeibusho, RibcNo";
                    var sda = new SqlDataAdapter(cmd);
                    AnkenData_Grid4.Clear();
                    sda.Fill(AnkenData_Grid4);
                }
                //技術担当者
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT  " +
                            "HyouronTantoushaCD " +
                            ",HyouronTantoushaMei " +
                            ",HyouronnTantoushaHyouten " +

                            //参照テーブル
                            "FROM GyoumuJouhouHyouronTantouL1 " +
                            " WHERE GyoumuJouhouID = " + AnkenID +
                            " ORDER BY HyouronTantouID ";
                    var sda = new SqlDataAdapter(cmd);
                    AnkenData_Grid5.Clear();
                    sda.Fill(AnkenData_Grid5);
                }

                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT  " +
                            //技術者評価タブ
                            "GyoumuHyouten " +
                            ",KanriGijutsushaNM " +
                            ",GyoumuKanriHyouten " +
                            ",ShousaTantoushaNM " +
                            ",GyoumuShousaHyouten " +
                            ",GyoumuTECRISTourokuBangou " +
                            ",CASE GyoumuSeikyuubi WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(GyoumuSeikyuubi,'yyyy/MM/dd') END " +
                            //",GyoumuSeikyuusho " +
                            ",AnkenKeiyakusho " +
                            ",AnkenKokyakuHyoukaComment " +
                            ",AnkenToukaiHyoukaComment " +

                            //参照テーブル
                            "FROM AnkenJouhou " +
                            "LEFT JOIN GyoumuJouhou ON AnkenJouhou.AnkenJouhouID = GyoumuJouhou.AnkenJouhouID " +
                            "LEFT JOIN NyuusatsuJouhou ON AnkenJouhou.AnkenJouhouID = NyuusatsuJouhou.AnkenJouhouID " +
                            "LEFT JOIN KeiyakuJouhouEntory ON AnkenJouhou.AnkenJouhouID = KeiyakuJouhouEntory.AnkenJouhouID " +
                            "WHERE AnkenJouhou.AnkenJouhouID =  " + AnkenID.ToString();
                    var sda = new SqlDataAdapter(cmd);
                    AnkenData_G.Clear();
                    sda.Fill(AnkenData_G);
                }

            }
            catch (Exception)
            {
                throw;
            }
            set_data(0);
        }

        private void set_data(int pagenum)
        {
            DateTime tmpDate;
            if (AnkenData_H.Rows.Count >= 1)
            {
                //ヘッダー情報
                Header1.Text = AnkenData_H.Rows[0][5].ToString();
                Header2.Text = AnkenData_H.Rows[0][6].ToString();
                Header3.Text = AnkenData_H.Rows[0][22].ToString() + "　" + AnkenData_H.Rows[0][23].ToString();
                Header4.Text = AnkenData_H.Rows[0][12].ToString();

                //【引合】
                //引合状況
                item1_1.SelectedValue = AnkenData_H.Rows[0][0].ToString();

                // 引合　引合状況　→発注確定　　→入札タブが表示（or操作可能になる）
                // 2:発注確定以外の場合、引合タブ非表示
                //if (item1_1.SelectedValue != null || item1_1.SelectedValue.ToString() == "2")
                //{
                //    // 表示
                //    this.tab.TabPages.Insert(0, this.tabPage3); // 入札
                //    this.tab.TabPages.Insert(0, this.tabPage4); // 契約
                //}
                //else
                //{
                //    // 非表示
                //    this.tab.TabPages.Remove(this.tabPage3); // 入札
                //    this.tab.TabPages.Remove(this.tabPage4); // 契約
                //}

                //基本情報
                item1_2.SelectedValue = AnkenData_H.Rows[0][1].ToString();
                // 工期開始年度
                item1_2_KoukiNendo.SelectedValue = AnkenData_H.Rows[0][56].ToString();
                item1_3.SelectedValue = AnkenData_H.Rows[0][2].ToString();
                item1_4.Text = AnkenData_H.Rows[0][3].ToString();
                beforeKeikakuBangou = AnkenData_H.Rows[0][3].ToString(); // 計画番号

                // VIPS 20220221 課題管理表No.1273(967) ADD 計画番号コピー制御用
                // この業務を元に新規登録ボタンから遷移してきたときは空にする
                if (CopyMode == "1")
                {
                    item1_4.Text = "";
                    beforeKeikakuBangou = "";
                }

                item1_5.Text = AnkenData_H.Rows[0][4].ToString();
                item1_6.Text = AnkenData_H.Rows[0][5].ToString();
                item1_7.Text = AnkenData_H.Rows[0][6].ToString();
                item1_8.Text = AnkenData_H.Rows[0][7].ToString();
                if (AnkenData_H.Rows[0][8].ToString() != "")
                {
                    item1_9.Text = AnkenData_H.Rows[0][8].ToString();
                }
                else
                {
                    item1_9.Text = "";
                    item1_9.CustomFormat = " ";
                }
                //item1_10.SelectedValue = AnkenData_H.Rows[0][9].ToString();
                item1_10.SelectedValue = AnkenData_H.Rows[0][9].ToString();
                BushoCD = AnkenData_H.Rows[0][9].ToString();
                item1_11.Text = AnkenData_H.Rows[0][10].ToString();
                item1_11_CD.Text = AnkenData_H.Rows[0][54].ToString();
                item1_11_Busho.Text = AnkenData_H.Rows[0][55].ToString();
                item1_12.Text = AnkenData_H.Rows[0][11].ToString();

                // えんとり君修正STEP2
                item1_37.Text = AnkenData_H.Rows[0][57].ToString();
                item1_37_kojinCD.Text = AnkenData_H.Rows[0][58].ToString();
                item1_38.Text = AnkenData_H.Rows[0][59].ToString();
                // No.1422 1196 案件番号の変更履歴を保存する
                item1_39.Text = AnkenData_H.Rows[0][61].ToString();
                sJigyoubuHeadCD_ori = AnkenData_H.Rows[0][60].ToString();
                //案件情報
                item1_13.Text = AnkenData_H.Rows[0][12].ToString();
                item1_14.SelectedValue = AnkenData_H.Rows[0][13].ToString();
                item1_15.SelectedValue = AnkenData_H.Rows[0][14].ToString();
                if (AnkenData_H.Rows[0][15].ToString() != "")
                {
                    item1_16.Text = AnkenData_H.Rows[0][15].ToString();
                }
                else
                {
                    item1_16.Text = "";
                    item1_16.CustomFormat = " ";
                }
                item1_17.SelectedValue = AnkenData_H.Rows[0][16].ToString();
                //えんとり君修正STEP2
                //案件メモ・参考見積額をコピーしない。
                // No1209 要望ほか、６件 案件メモは残す
                item1_18.Text = AnkenData_H.Rows[0][17].ToString();
                //if (CopyMode != "1")
                //{
                //    item1_18.Text = AnkenData_H.Rows[0][17].ToString();
                //}
                //発注者情報
                item1_19.Text = AnkenData_H.Rows[0][18].ToString();
                item1_20.Text = AnkenData_H.Rows[0][19].ToString();
                item1_21.Text = AnkenData_H.Rows[0][20].ToString();
                item1_22.Text = AnkenData_H.Rows[0][21].ToString();
                item1_23.Text = AnkenData_H.Rows[0][22].ToString();
                item1_24.Text = AnkenData_H.Rows[0][23].ToString();
                //発注担当者情報
                item1_25.Text = AnkenData_H.Rows[0][24].ToString();
                item1_26.Text = AnkenData_H.Rows[0][25].ToString();
                item1_27.Text = AnkenData_H.Rows[0][26].ToString();
                item1_28.Text = AnkenData_H.Rows[0][27].ToString();
                item1_29.Text = AnkenData_H.Rows[0][28].ToString();
                item1_30.Text = AnkenData_H.Rows[0][29].ToString();
                item1_31.Text = AnkenData_H.Rows[0][30].ToString();
                item1_32.Text = AnkenData_H.Rows[0][31].ToString();
                item1_33.Text = AnkenData_H.Rows[0][32].ToString();
                //当会応札
                item1_34.SelectedValue = AnkenData_H.Rows[0][33].ToString();
                item1_35.SelectedValue = AnkenData_H.Rows[0][34].ToString();
                //えんとり君修正STEP2
                //案件メモ・参考見積額をコピーしない。
                //item1_36.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_H.Rows[0][35]));
                if (CopyMode != "1")
                {
                    item1_36.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_H.Rows[0][35]));
                }
                //業務情報（引合）
                item1_7_1_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][36]));
                item1_7_1_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][37]));
                item1_7_1_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][38]));
                item1_7_1_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][39]));
                item1_7_2_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][40]));
                item1_7_2_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][41]));
                item1_7_2_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][42]));
                item1_7_2_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][43]));
                item1_7_2_5_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][44]));
                item1_7_2_6_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][45]));
                item1_7_2_7_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][46]));
                item1_7_2_8_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][47]));
                item1_7_2_9_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][48]));
                item1_7_2_10_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][49]));
                item1_7_2_11_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][50]));
                item1_7_2_12_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][51]));

                TotalPercent("item1_7_1_", "_1", 5);
                TotalPercent("item1_7_2_", "_1", 13);

                //業務情報（入札）
                //エントリ君修正STEP1
                //計画詳細の案件番号から新規登録ボタンで遷移した場合は、受託した場合、それ以外でデータの貼り付けが異なる
                //AnkenData_Grid2 ←落札者情報が格納されているDataTable
                //　このテーブルを操作し、1000かつ落札Flagが立っている場合は契約情報から自動セットする
                //AnkenData_N
                //　入札データ格納されているDatatable
                //AnkenData_K
                //　契約データ格納されているDatatable
                if (isKeikakuAnkenNew == true)
                {
                    //受託か否か
                    bool isJutaku = false;
                    //企業コード　：　1001 建設物価調査会
                    //受託フラグ　：　NyuusatsuRakusatsuJokyou = 1
                    //↑上記が見つかれば受託案件
                    for (int i = 0; i < AnkenData_Grid2.Rows.Count; i++)
                    {
                        if (AnkenData_Grid2.Rows[i][2].ToString() == "1001" && AnkenData_Grid2.Rows[i][1].ToString() == "1")
                        {
                            isJutaku = true;
                            break;
                        }
                    }
                    //受託だった場合は契約情報からコピー
                    if (isJutaku == true)
                    {
                        //契約でも部署ごとの配分はAnkenData_Hのデータ[GyoumuHaibun]テーブルの[GyoumuHibunKubun]が10で良いようだ。
                        item2_4_1_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][36]));
                        item2_4_1_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][37]));
                        item2_4_1_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][38]));
                        item2_4_1_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][39]));
                        item2_4_2_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][71]));
                        item2_4_2_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][72]));
                        item2_4_2_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][73]));
                        item2_4_2_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][74]));
                        item2_4_2_5_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][75]));
                        item2_4_2_6_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][76]));
                        item2_4_2_7_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][77]));
                        item2_4_2_8_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][78]));
                        item2_4_2_9_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][79]));
                        item2_4_2_10_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][80]));
                        item2_4_2_11_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][81]));
                        item2_4_2_12_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][82]));
                    }
                    else
                    {
                        item2_4_1_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][21]));
                        item2_4_1_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][22]));
                        item2_4_1_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][23]));
                        item2_4_1_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][24]));
                        item2_4_2_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][25]));
                        item2_4_2_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][26]));
                        item2_4_2_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][27]));
                        item2_4_2_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][28]));
                        item2_4_2_5_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][29]));
                        item2_4_2_6_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][30]));
                        item2_4_2_7_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][31]));
                        item2_4_2_8_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][32]));
                        item2_4_2_9_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][33]));
                        item2_4_2_10_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][34]));
                        item2_4_2_11_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][35]));
                        item2_4_2_12_1.Text = GetPercentText(Convert.ToDouble(AnkenData_N.Rows[0][36]));
                    }
                }
                else
                {
                    item2_4_1_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][36]));
                    item2_4_1_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][37]));
                    item2_4_1_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][38]));
                    item2_4_1_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][39]));
                    item2_4_2_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][40]));
                    item2_4_2_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][41]));
                    item2_4_2_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][42]));
                    item2_4_2_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][43]));
                    item2_4_2_5_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][44]));
                    item2_4_2_6_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][45]));
                    item2_4_2_7_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][46]));
                    item2_4_2_8_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][47]));
                    item2_4_2_9_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][48]));
                    item2_4_2_10_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][49]));
                    item2_4_2_11_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][50]));
                    item2_4_2_12_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][51]));
                }
                //item2_4_1_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][36]));
                //item2_4_1_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][37]));
                //item2_4_1_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][38]));
                //item2_4_1_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][39]));
                //item2_4_2_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][40]));
                //item2_4_2_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][41]));
                //item2_4_2_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][42]));
                //item2_4_2_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][43]));
                //item2_4_2_5_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][44]));
                //item2_4_2_6_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][45]));
                //item2_4_2_7_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][46]));
                //item2_4_2_8_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][47]));
                //item2_4_2_9_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][48]));
                //item2_4_2_10_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][49]));
                //item2_4_2_11_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][50]));
                //item2_4_2_12_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][51]));

                TotalPercent("item2_4_1_", "_1", 5);
                TotalPercent("item2_4_2_", "_1", 13);

                //業務情報（契約）
                item3_7_1_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][36]));
                item3_7_1_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][37]));
                item3_7_1_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][38]));
                item3_7_1_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][39]));
                item3_7_2_1_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][40]));
                item3_7_2_2_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][41]));
                item3_7_2_3_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][42]));
                item3_7_2_4_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][43]));
                item3_7_2_5_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][44]));
                item3_7_2_6_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][45]));
                item3_7_2_7_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][46]));
                item3_7_2_8_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][47]));
                item3_7_2_9_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][48]));
                item3_7_2_10_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][49]));
                item3_7_2_11_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][50]));
                item3_7_2_12_1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0][51]));

                TotalPercent("item3_7_1_", "_1", 5);
                TotalPercent("item3_7_2_", "_1", 13);

                saishinFLG = Convert.ToInt32(AnkenData_H.Rows[0][53]);
                FolderPathCheck();
            }
            // 新規じゃない、またはこの業務を元に新規登録の場合
            if (mode != "insert" || (mode == "insert" && AnkenbaBangou == ""))
            {
                for (int k = 0; k < AnkenData_Grid1.Rows.Count; k++)
                {
                    if (k > 0)
                    {
                        c1FlexGrid1.Rows.Add();
                    }
                    if (k >= 5)
                    {
                        break;
                    }
                    //【過去案件】
                    // SortKey 以外をc1FlexGridにセットする
                    for (int i = 0; i < AnkenData_Grid1.Columns.Count - 1; i++)
                    {
                        c1FlexGrid1.Rows[k + 1][i + 2] = AnkenData_Grid1.Rows[k][i];
                    }
                }
            }
            // 過去案件情報がない場合、
            if (AnkenData_Grid1.Rows.Count == 0)
            {
                // 過去案件情報の前回受託番号IDに1を入れておく
                c1FlexGrid1.Rows[1][16] = 1;
            }
            if (AnkenData_N.Rows.Count >= 1)
            {

                //【入札】
                //入札状況
                item2_1_1.SelectedValue = AnkenData_N.Rows[0][0].ToString();
                // 不具合397 入札　入札状況　→入札成立　　→契約タブが表示（or操作可能になる）
                //if (item2_1_1.SelectedValue != null && "入札成立".Equals(item2_1_1.Text))
                //{
                //    this.tab.TabPages.Remove(this.tabPage4); // 契約
                //}
                //else
                //{
                //    //this.tab.TabPages.Insert(0, this.tabPage4);
                //}


                item2_1_2.Text = AnkenData_N.Rows[0][1].ToString();
                if (AnkenData_N.Rows[0][1].ToString() != "")
                {
                    item2_1_2.Text = AnkenData_N.Rows[0][1].ToString();
                }
                else
                {
                    item2_1_2.Text = "";
                    item2_1_2.CustomFormat = " ";
                }
                if (AnkenData_N.Rows[0][2] != null && AnkenData_N.Rows[0][2].ToString() != "")
                {
                    item2_1_3.Text = AnkenData_N.Rows[0][2].ToString();
                }
                else
                {
                    item2_1_3.Text = "";
                    item2_1_3.CustomFormat = " ";
                }
                item2_1_4.SelectedValue = AnkenData_N.Rows[0][3].ToString();
                item2_1_5.SelectedValue = AnkenData_N.Rows[0][4].ToString();
                item2_1_6.Text = AnkenData_N.Rows[0][5].ToString();
                //当会応札
                if (AnkenData_N.Rows[0][6].ToString() != "")
                {
                    item2_2_1.SelectedValue = AnkenData_N.Rows[0][6].ToString();
                }
                item2_2_2.SelectedValue = AnkenData_N.Rows[0][7].ToString();
                item2_2_3.SelectedValue = AnkenData_N.Rows[0][8].ToString();
                item2_2_4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_N.Rows[0][9]));
                //入札結果
                if (AnkenData_N.Rows[0][10].ToString() != "")
                {
                    item2_3_1.SelectedValue = AnkenData_N.Rows[0][10].ToString();
                }
                if (AnkenData_N.Rows[0][11].ToString() != "")
                {
                    item2_3_2.SelectedValue = AnkenData_N.Rows[0][11].ToString();
                }
                if (AnkenData_N.Rows[0][12].ToString() != null && AnkenData_N.Rows[0][12].ToString() != "")
                {
                    item2_3_3.Text = AnkenData_N.Rows[0][12].ToString();
                    item2_3_3.CustomFormat = "";
                }
                else
                {
                    item2_3_3.CustomFormat = " ";
                }
                if (AnkenData_N.Rows[0][13].ToString() != "")
                {
                    item2_3_4.Text = AnkenData_N.Rows[0][13].ToString();
                    item2_3_4.CustomFormat = "";
                }
                else
                {
                    item2_3_4.CustomFormat = " ";
                }
                item2_3_5.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_N.Rows[0][14]));
                item2_3_6.Text = Convert.ToInt32(AnkenData_N.Rows[0][15]).ToString();
                item2_3_7.Text = AnkenData_N.Rows[0][16].ToString();
                item2_3_8.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_N.Rows[0][17]));
                //item2_3_9.Text = AnkenData_N.Rows[0][18].ToString();
                //item2_3_10.Text = AnkenData_N.Rows[0][19].ToString();
                //item2_3_12.Text = AnkenData_N.Rows[0][20].ToString();
                //item2_3_9.Text = AnkenData_N.Rows[0][18].ToString();
                //item2_3_10.Text = AnkenData_N.Rows[0][19].ToString();
                item2_3_12.Text = AnkenData_N.Rows[0][20].ToString();
                get_guidance();
            }
            for (int k = 0; k < AnkenData_Grid2.Rows.Count; k++)
            {
                if (k > 0)
                {
                    c1FlexGrid2.Rows.Add();
                }
                //【入札参加者】
                for (int i = 0; i < AnkenData_Grid2.Columns.Count; i++)
                {
                    c1FlexGrid2.Rows[k + 1][i + 2] = AnkenData_Grid2.Rows[k][i].ToString();
                }
            }


            if (AnkenData_K.Rows.Count >= 1)
            {

                //【契約】
                //契約情報
                item3_1_1.SelectedValue = AnkenData_K.Rows[0][0].ToString();
                // えんとり君修正STEP2 伝票変更前も確認シートを出力する時、案件区分
                sAnkenSakuseiKubun_ori = AnkenData_K.Rows[0][0].ToString();
                if (AnkenData_K.Rows[0][1].ToString() == "1")
                {
                    item3_1_2.Checked = true;
                }
                else
                {
                    item3_1_2.Checked = false;
                }

                //VIPS 20220427 課題管理表No.1277(971) CHANGE 「契約締結（変更）」日を空欄で表示
                //if (AnkenData_K.Rows[0][2].ToString() != null && AnkenData_K.Rows[0][2].ToString() != "")
                if (AnkenData_K.Rows[0][2].ToString() != null && AnkenData_K.Rows[0][2].ToString() != "" && mode != "change")
                {
                    Console.WriteLine("'" + AnkenData_K.Rows[0][2].ToString() + "'");
                    item3_1_3.Text = AnkenData_K.Rows[0][2].ToString();
                }
                else
                {
                    Console.WriteLine("'" + AnkenData_K.Rows[0][2].ToString() + "'");
                    item3_1_3.CustomFormat = " ";
                }

                if (AnkenData_K.Rows[0][3].ToString() != null && AnkenData_K.Rows[0][3].ToString() != "")
                {
                    item3_1_4.Text = AnkenData_K.Rows[0][3].ToString();
                }
                else
                {
                    item3_1_4.CustomFormat = " ";
                }
                item3_1_5.SelectedValue = AnkenData_K.Rows[0][4].ToString(); if (AnkenData_K.Rows[0][5].ToString() != "")
                {
                    item3_1_6.Text = AnkenData_K.Rows[0][5].ToString();
                }
                else
                {
                    item3_1_6.CustomFormat = " ";
                }
                if (AnkenData_K.Rows[0][6].ToString() != "")
                {
                    item3_1_7.Text = AnkenData_K.Rows[0][6].ToString();
                }
                else
                {
                    item3_1_7.CustomFormat = " ";
                }
                item3_1_8.SelectedValue = AnkenData_K.Rows[0][7].ToString();
                item3_1_9.Text = AnkenData_K.Rows[0][8].ToString();
                item3_1_10.Text = AnkenData_K.Rows[0][9].ToString();
                item3_1_11.Text = AnkenData_K.Rows[0][95].ToString();
                item3_1_12.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][11]));
                item3_1_13.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][12]));
                item3_1_14.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][13]));
                item3_1_15.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][14]));
                item3_1_16.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][15]));
                item3_1_17.Text = AnkenData_K.Rows[0][16].ToString();
                item3_1_18.Text = AnkenData_K.Rows[0][17].ToString();
                item3_1_19.Text = AnkenData_K.Rows[0][18].ToString();
                if (AnkenData_K.Rows[0][19].ToString() == "1")
                {
                    item3_1_20.Checked = true;
                }
                if (AnkenData_K.Rows[0][20].ToString() == "1")
                {
                    item3_1_21.Checked = true;
                }
                if (AnkenData_K.Rows[0][21].ToString() == "1")
                {
                    item3_1_22.Checked = true;
                }
                if (AnkenData_K.Rows[0][22].ToString() == "1")
                {
                    item3_1_23.Checked = true;
                }
                if (AnkenData_K.Rows[0][23].ToString() == "1")
                {
                    item3_1_24.Checked = true;
                }
                item3_1_25.Text = AnkenData_K.Rows[0][24].ToString();
                item3_1_26.Text = AnkenData_K.Rows[0][25].ToString();
                //配分情報
                item3_2_1_1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][26]));
                item3_2_2_1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][27]));
                item3_2_3_1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][28]));
                item3_2_4_1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][29]));
                TotalMoney("item3_2_", "_1", 5);
                item3_2_1_2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][30]));
                item3_2_2_2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][31]));
                item3_2_3_2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][32]));
                item3_2_4_2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][33]));
                TotalMoney("item3_2_", "_2", 5);
                //単契
                item3_3_1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][34]));
                item3_3_2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][35]));
                item3_3_3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][36]));
                item3_3_4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][37]));
                TotalMoney("item3_3_", "", 5);
                //年度繰越
                item3_7_1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][38]));
                item3_7_2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][39]));
                item3_7_3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][40]));
                item3_7_4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][41]));
                TotalMoney("item3_7_", "", 5);
                //管理者
                item3_4_1_CD.Text = AnkenData_K.Rows[0][42].ToString();
                item3_4_1.Text = AnkenData_K.Rows[0][43].ToString();
                item3_4_2_CD.Text = AnkenData_K.Rows[0][44].ToString();
                item3_4_2.Text = AnkenData_K.Rows[0][45].ToString();
                item3_4_3_CD.Text = AnkenData_K.Rows[0][46].ToString();
                item3_4_3.Text = AnkenData_K.Rows[0][47].ToString();
                item3_4_4_CD.Text = AnkenData_K.Rows[0][48].ToString();
                item3_4_4.Text = AnkenData_K.Rows[0][49].ToString();
                item3_4_5_CD.Text = AnkenData_K.Rows[0][50].ToString();
                item3_4_5.Text = AnkenData_K.Rows[0][51].ToString();
                item3_4_5_Busho.Text = AnkenData_K.Rows[0][52].ToString();
                item3_4_5_Shibu.Text = AnkenData_K.Rows[0][53].ToString();
                item3_4_5_Ka.Text = AnkenData_K.Rows[0][54].ToString();
                //請求書
                //Set_datetime(AnkenData_K.Rows[0][43].ToString(), "item3_6_1");
                item3_6_1.Text = AnkenData_K.Rows[0][55].ToString();
                if (AnkenData_K.Rows[0][55].ToString() != "")
                {
                    item3_6_1.Text = AnkenData_K.Rows[0][55].ToString();
                }
                else
                {
                    item3_6_1.CustomFormat = " ";
                }
                item3_6_2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][56]));
                if (AnkenData_K.Rows[0][57].ToString() != "")
                {
                    item3_6_3.Text = AnkenData_K.Rows[0][57].ToString();
                }
                else
                {
                    item3_6_3.CustomFormat = " ";
                }
                item3_6_4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][58]));
                if (AnkenData_K.Rows[0][59].ToString() != "")
                {
                    item3_6_5.Text = AnkenData_K.Rows[0][59].ToString();
                }
                else
                {
                    item3_6_5.CustomFormat = " ";
                }
                item3_6_6.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][60]));
                if (AnkenData_K.Rows[0][61].ToString() != "")
                {
                    item3_6_7.Text = AnkenData_K.Rows[0][61].ToString();
                }
                else
                {
                    item3_6_7.CustomFormat = " ";
                }
                item3_6_8.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][62]));
                if (AnkenData_K.Rows[0][63].ToString() != "")
                {
                    item3_6_9.Text = AnkenData_K.Rows[0][63].ToString();
                }
                else
                {
                    item3_6_9.CustomFormat = " ";
                }
                item3_6_10.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][64]));
                if (AnkenData_K.Rows[0][65].ToString() != "")
                {
                    item3_6_11.Text = AnkenData_K.Rows[0][65].ToString();
                }
                else
                {
                    item3_6_11.CustomFormat = " ";
                }
                item3_6_12.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0][66]));
                Total3_6();

                //業務内容
                item3_7_1_6_1.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][67]));
                item3_7_1_7_1.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][68]));
                item3_7_1_8_1.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][69]));
                item3_7_1_9_1.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][70]));
                item3_7_2_14_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][71]));
                item3_7_2_15_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][72]));
                item3_7_2_16_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][73]));
                item3_7_2_17_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][74]));
                item3_7_2_18_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][75]));
                item3_7_2_19_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][76]));
                item3_7_2_20_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][77]));
                item3_7_2_21_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][78]));
                item3_7_2_22_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][79]));
                item3_7_2_23_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][80]));
                item3_7_2_24_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][81]));
                item3_7_2_25_1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0][82]));
                item3_7_2_14_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][83]));
                item3_7_2_15_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][84]));
                item3_7_2_16_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][85]));
                item3_7_2_17_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][86]));
                item3_7_2_18_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][87]));
                item3_7_2_19_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][88]));
                item3_7_2_20_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][89]));
                item3_7_2_21_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][90]));
                item3_7_2_22_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][91]));
                item3_7_2_24_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][93]));
                item3_7_2_25_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][94]));
                item3_7_2_23_2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0][92]));

                TotalMoney("item3_7_1_", "_1", 5, 6);
                TotalPercent("item3_7_2_", "_1", 13, 14);
                TotalMoney("item3_7_2_", "_2", 13, 14);
            }

            //売上請求情報
            int rowCntT = 0;
            int rowCntB = 0;
            int rowCntJ = 0;
            int rowCntK = 0;
            int row = 1;
            for (int k = 0; k < AnkenData_Grid4.Rows.Count; k++)
            {
                if ("T".Equals(AnkenData_Grid4.Rows[k][4].ToString()))
                {
                    rowCntT++;
                    if (row < rowCntT)
                    {
                        //c1FlexGrid4.Rows.Add();
                        row += 1;
                    }
                    c1FlexGrid4.Rows[rowCntT + 1][1] = AnkenData_Grid4.Rows[k][1].ToString();
                    c1FlexGrid4.Rows[rowCntT + 1][2] = AnkenData_Grid4.Rows[k][2].ToString();
                    c1FlexGrid4.Rows[rowCntT + 1][3] = AnkenData_Grid4.Rows[k][3].ToString();
                    c1FlexGrid4.Rows[rowCntT + 1][4] = AnkenData_Grid4.Rows[k][5].ToString();
                    c1FlexGrid4.Rows[rowCntT + 1][5] = AnkenData_Grid4.Rows[k][6].ToString();
                    c1FlexGrid4.Rows[rowCntT + 1][6] = AnkenData_Grid4.Rows[k][7].ToString();
                    c1FlexGrid4.Rows[rowCntT + 1][7] = AnkenData_Grid4.Rows[k][8].ToString();
                    c1FlexGrid4.Rows[rowCntT + 1][8] = AnkenData_Grid4.Rows[k][9].ToString();
                }
                if ("B".Equals(AnkenData_Grid4.Rows[k][4].ToString()))
                {
                    rowCntB++;
                    if (row < rowCntB)
                    {
                        //c1FlexGrid4.Rows.Add();
                        row += 1;
                    }
                    c1FlexGrid4.Rows[rowCntB + 1][9] = AnkenData_Grid4.Rows[k][1].ToString();
                    c1FlexGrid4.Rows[rowCntB + 1][10] = AnkenData_Grid4.Rows[k][2].ToString();
                    c1FlexGrid4.Rows[rowCntB + 1][11] = AnkenData_Grid4.Rows[k][3].ToString();
                    c1FlexGrid4.Rows[rowCntB + 1][12] = AnkenData_Grid4.Rows[k][5].ToString();
                    c1FlexGrid4.Rows[rowCntB + 1][13] = AnkenData_Grid4.Rows[k][6].ToString();
                    c1FlexGrid4.Rows[rowCntB + 1][14] = AnkenData_Grid4.Rows[k][7].ToString();
                    c1FlexGrid4.Rows[rowCntB + 1][15] = AnkenData_Grid4.Rows[k][8].ToString();
                    c1FlexGrid4.Rows[rowCntB + 1][16] = AnkenData_Grid4.Rows[k][9].ToString();
                }
                if ("J".Equals(AnkenData_Grid4.Rows[k][4].ToString()))
                {
                    rowCntJ++;
                    if (row < rowCntJ)
                    {
                        //c1FlexGrid4.Rows.Add();
                        row += 1;
                    }
                    c1FlexGrid4.Rows[rowCntJ + 1][17] = AnkenData_Grid4.Rows[k][1].ToString();
                    c1FlexGrid4.Rows[rowCntJ + 1][18] = AnkenData_Grid4.Rows[k][2].ToString();
                    c1FlexGrid4.Rows[rowCntJ + 1][19] = AnkenData_Grid4.Rows[k][3].ToString();
                    c1FlexGrid4.Rows[rowCntJ + 1][20] = AnkenData_Grid4.Rows[k][5].ToString();
                    c1FlexGrid4.Rows[rowCntJ + 1][21] = AnkenData_Grid4.Rows[k][6].ToString();
                    c1FlexGrid4.Rows[rowCntJ + 1][22] = AnkenData_Grid4.Rows[k][7].ToString();
                    c1FlexGrid4.Rows[rowCntJ + 1][23] = AnkenData_Grid4.Rows[k][8].ToString();
                    c1FlexGrid4.Rows[rowCntJ + 1][24] = AnkenData_Grid4.Rows[k][9].ToString();
                }
                if ("K".Equals(AnkenData_Grid4.Rows[k][4].ToString()))
                {
                    rowCntK++;
                    if (row < rowCntK)
                    {
                        //c1FlexGrid4.Rows.Add();
                        row += 1;
                    }
                    c1FlexGrid4.Rows[rowCntK + 1][25] = AnkenData_Grid4.Rows[k][1].ToString();
                    c1FlexGrid4.Rows[rowCntK + 1][26] = AnkenData_Grid4.Rows[k][2].ToString();
                    c1FlexGrid4.Rows[rowCntK + 1][27] = AnkenData_Grid4.Rows[k][3].ToString();
                    c1FlexGrid4.Rows[rowCntK + 1][28] = AnkenData_Grid4.Rows[k][5].ToString();
                    c1FlexGrid4.Rows[rowCntK + 1][29] = AnkenData_Grid4.Rows[k][6].ToString();
                    c1FlexGrid4.Rows[rowCntK + 1][30] = AnkenData_Grid4.Rows[k][7].ToString();
                    c1FlexGrid4.Rows[rowCntK + 1][31] = AnkenData_Grid4.Rows[k][8].ToString();
                    c1FlexGrid4.Rows[rowCntK + 1][32] = AnkenData_Grid4.Rows[k][9].ToString();
                }
            }
            for (int k = 0; k < AnkenData_Grid5.Rows.Count; k++)
            {
                if (k > 0)
                {
                    c1FlexGrid3.Rows.Add();
                    c1FlexGrid5.Rows.Add();
                }
                ////【技術担当者】
                //for (int i = 0; i < 2; i++)
                //{
                //    c1FlexGrid3.Rows[k + 1][i + 1] = AnkenData_Grid5.Rows[k][i].ToString();
                //}
                //【技術担当者】
                for (int i = 0; i < AnkenData_Grid5.Columns.Count; i++)
                {
                    c1FlexGrid3.Rows[k + 1][i + 1] = AnkenData_Grid5.Rows[k][i].ToString();
                    c1FlexGrid5.Rows[k + 1][i + 1] = AnkenData_Grid5.Rows[k][i].ToString();
                }
            }

            if (AnkenData_G.Rows.Count >= 1)
            {
                item4_1_1.Text = AnkenData_G.Rows[0][0].ToString();
                item4_1_2.Text = AnkenData_G.Rows[0][1].ToString();
                item4_1_3.Text = AnkenData_G.Rows[0][2].ToString();
                item3_4_1_Hyoten.Text = AnkenData_G.Rows[0][2].ToString();
                item4_1_4.Text = AnkenData_G.Rows[0][3].ToString();
                item4_1_5.Text = AnkenData_G.Rows[0][4].ToString();
                item3_4_2_Hyoten.Text = AnkenData_G.Rows[0][4].ToString();
                item4_1_6.Text = AnkenData_G.Rows[0][5].ToString();
                if (AnkenData_G.Rows[0][6] != null && AnkenData_G.Rows[0][6].ToString() != "")
                {
                    item4_1_7.Text = AnkenData_G.Rows[0][6].ToString();
                }
                else
                {
                    item4_1_7.CustomFormat = " ";
                }
                // 技術評価者タブの請求書は 02契約関係図書 を付ける
                item4_1_8.Text = AnkenData_G.Rows[0][7].ToString() + @"\02契約関係図書";
                item4_1_9.Text = AnkenData_G.Rows[0][8].ToString();
                item4_1_10.Text = AnkenData_G.Rows[0][9].ToString();

                // フォルダパス振り直し
                set_folder();
            }
            // 引合タブの配分の合計計算
            TotalPercent("item1_7_1_", "_1", 5);
            copy_haibun("item1_7_1_", "item2_4_1_", 5);
            copy_haibun("item1_7_1_", "item3_7_1_", 5);
            TotalPercent("item1_7_2_", "_1", 13);
            copy_haibun("item1_7_2_", "item2_4_2_", 13);
            copy_haibun("item1_7_2_", "item3_7_2_", 13);

            // えんとり君修正STEP2（RIBC項目追加）
            if (AnkenData_K.Rows[0][96].ToString() == "1")
            {
                item3_ribc_price.Checked = true;
            }
            if (AnkenData_K.Rows[0][97].ToString() == "1")
            {
                item3_sa_commpany.Checked = true;
            }
            if (AnkenData_K.Rows[0][98].ToString() == "1")
            {
                item3_1_ribc.Checked = true;
            }
        }

        private void c1FlexGrid4_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            if (c1FlexGrid4.Col == 1 || c1FlexGrid4.Col == 9 || c1FlexGrid4.Col == 17 || c1FlexGrid4.Col == 25)
            {
                // 工期末日付を空にした場合
                if (c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col] == null)
                {
                    c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col + 1] = null;
                }
                else
                {
                    c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col + 1] = DateTime.Parse(c1FlexGrid4.Rows[c1FlexGrid4.Row][c1FlexGrid4.Col].ToString()).ToString("yyyy/MM");
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (item3_1_7.CustomFormat != "")
            {
                set_error("", 0);
                set_error(GlobalMethod.GetMessage("E10011", ""));
            }
            else
            {
                string GyoumuCD = item3_1_8.SelectedValue.ToString();
                if (GyoumuCD == "1" || GyoumuCD == "2" || GyoumuCD == "3" || GyoumuCD == "4")
                {
                    c1FlexGrid4.Rows[2][1] = item3_1_7.Text;
                    c1FlexGrid4.Rows[2][2] = DateTime.Parse(item3_1_7.Text).ToString("yyyy/MM");
                }
                else if (GyoumuCD == "5" || GyoumuCD == "6")
                {
                    c1FlexGrid4.Rows[2][9] = item3_1_7.Text;
                    c1FlexGrid4.Rows[2][10] = DateTime.Parse(item3_1_7.Text).ToString("yyyy/MM");
                }
                else if (GyoumuCD == "7")
                {
                    c1FlexGrid4.Rows[2][17] = item3_1_7.Text;
                    c1FlexGrid4.Rows[2][18] = DateTime.Parse(item3_1_7.Text).ToString("yyyy/MM");
                }
                else if (GyoumuCD == "8")
                {
                    c1FlexGrid4.Rows[2][25] = item3_1_7.Text;
                    c1FlexGrid4.Rows[2][26] = DateTime.Parse(item3_1_7.Text).ToString("yyyy/MM");
                }
                item3_6_1.Text = item3_1_7.Text;
            }
        }

        private void label586_Click(object sender, EventArgs e)
        {
            if (item3_1_7.CustomFormat != "")
            {
                item3_6_1.CustomFormat = " ";
            }
            else
            {
                item3_6_1.Text = item3_1_7.Text;
            }
            item3_6_2.Text = item3_1_13.Text;
        }

        // 起案ボタン
        private void button14_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show(GlobalMethod.GetMessage("I10704", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {

                if (!ErrorFLG(1))
                {

                    if (KianError())
                    {
                        //if (ErrorFLG(3))
                        //{
                        //    if (KianError(1))
                        //    {
                        if (Execute_SQL(3))
                        {
                            Entry_Input form = new Entry_Input();
                            form.UserInfos = UserInfos;
                            //form.Message = GlobalMethod.GetMessage("I10708", "");
                            form.Message = ErrorMessage.Text + Environment.NewLine + GlobalMethod.GetMessage("I10708", "");
                            form.KianFLG = true;
                            form.mode = "";
                            form.AnkenID = AnkenID;
                            form.Show(this.Owner);
                            ownerflg = false;
                            this.Close();
                        }
                    }
                    // 起案時に帳票出力はない
                    //string[] result = GlobalMethod.InsertReportWork(1, UserInfos[0], new string[] { AnkenID, Header1.Text, "0" });

                    //if(result != null)
                    //{
                    //    if (result[0] == "1")
                    //    {
                    //        set_error(result[1]);
                    //    }
                    //    else
                    //    {
                    //        Popup_Download form = new Popup_Download();
                    //        form.TopLevel = false;
                    //        this.Controls.Add(form);
                    //        form.ExcelPath = result[2];
                    //        form.Show();
                    //        form.BringToFront();
                    //    }
                    //}
                }
            }
        }

        // 変更伝票時の起案ボタン
        private void button13_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10704", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                //set_error("", 0);
                //if (KianError(1))

                if (KianError())
                {

                    if (Execute_SQL(4))
                    {
                        // 変更後の起案ボタン
                        button13.Enabled = false;
                        if (item3_1_1.Text != null && item3_1_1.Text != "" && (item3_1_1.SelectedValue.ToString() == "03" || int.Parse(item3_1_1.SelectedValue.ToString()) > 5))
                        {
                            // 案件区分が以下の場合のみ赤伝ボタン有効化
                            // 03 契約変更(黒伝)
                            // 06 契約変更(黒伝・金額変更)
                            // 07 契約変更(黒伝・工期変更)
                            // 08 契約変更(黒伝・金額工期変更)
                            // 09 契約変更(黒伝・その他)
                            // 赤伝作成・出力ボタン
                            // えんとり君修正STEP2 変更伝票画面で「チェック用帳票出力・内容確認」追加　赤伝作成・出力ボタン流用する
                            //button20.Enabled = true;
                            //button20.BackColor = Color.FromArgb(42, 78, 122);
                        }
                        //// 赤伝作成・出力ボタン
                        //button20.Enabled = true;
                        // 黒伝・中止伝票作成・出力
                        button21.Enabled = true;

                        button13.BackColor = Color.DarkGray;
                        //button20.BackColor = Color.FromArgb(42, 78, 122);
                        button21.BackColor = Color.FromArgb(42, 78, 122);

                        // 起案したので、案件区分を編集不可に
                        item3_1_1.Enabled = false;
                    }
                    else
                    {
                        set_error(GlobalMethod.GetMessage("E10009", ""));
                    }
                    // 起案時に帳票出力はない
                    //string[] result = GlobalMethod.InsertReportWork(1, UserInfos[0], new string[] { AnkenID, Header1.Text, "0" });

                    //if (result != null)
                    //{
                    //    if (result[0] == "1")
                    //    {
                    //        set_error(result[1]);
                    //    }
                    //    else
                    //    {
                    //        Popup_Download form = new Popup_Download();
                    //        form.TopLevel = false;
                    //        this.Controls.Add(form);
                    //        form.ExcelPath = result[2];
                    //        form.Show();
                    //        form.BringToFront();
                    //    }
                    //}

                }
            }
        }

        // 起案解除ボタン
        private void button17_Click(object sender, EventArgs e)
        {
            string methodName = ".btnKianKaijo_Click";

            if (MessageBox.Show(GlobalMethod.GetMessage("I10707", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {
                        cmd.CommandText = "UPDATE AnkenJouhou SET " +
                                    "AnkenKianzumi = 0" +
                                    " WHERE AnkenJouhouID = " + AnkenID;
                        //var sda = new SqlDataAdapter(cmd);
                        cmd.ExecuteNonQuery();

                        transaction.Commit();
                        //GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "起案解除しました ID:" + AnkenID, "InsertEntry", "");
                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "起案解除しました ID:" + AnkenID, pgmName + methodName, "");
                        //set_error(GlobalMethod.GetMessage("I10709",""));
                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        throw;
                    }
                    conn.Close();
                }
                Entry_Input form = new Entry_Input();
                form.UserInfos = UserInfos;
                form.Message = GlobalMethod.GetMessage("I10709", "");
                form.KianKaijoFLG = true;
                form.mode = "update";
                form.AnkenID = AnkenID;
                form.Show(this.Owner);
                ownerflg = false;
                this.Close();
            }
        }

        // 変更伝票ボタン
        private void button16_Click(object sender, EventArgs e)
        {
            Boolean ErrorFLG = true;
            set_error("", 0);
            if (int.Parse(UserInfos[4]) != 2 && !UserInfos[2].Equals(item1_10.SelectedValue.ToString()))
            {
                set_error(GlobalMethod.GetMessage("E10003", ""));
                ErrorFLG = false;
            }

            if (saishinFLG != 1)
            {
                set_error(GlobalMethod.GetMessage("E10006", ""));
                ErrorFLG = false;
            }
            if (String.IsNullOrEmpty(item3_1_1.Text) || item3_1_1.SelectedValue.ToString() == "02" || item3_1_1.SelectedValue.ToString() == "04")
            {
                set_error(GlobalMethod.GetMessage("E10007", ""));
                ErrorFLG = false;
            }
            if (!item3_1_2.Checked)
            {
                set_error(GlobalMethod.GetMessage("E10008", ""));
                ErrorFLG = false;
            }

            if (ErrorFLG)
            {
                Entry_Input form = new Entry_Input();
                form.mode = "change";
                form.AnkenID = AnkenID;
                form.UserInfos = UserInfos;
                form.ChangeFlag = 1;
                form.Show(this.Owner);
                ownerflg = false;
                this.Close();
            }
        }

        // エントリーシート作成・出力
        private void button15_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                if (!ErrorFLG(3))
                {
                    Execute_SQL(2);
                    // えんとり君修正STEP2
                    //string[] result = GlobalMethod.InsertReportWork(1, UserInfos[0], new string[] { AnkenID, Header1.Text, "0", "0" });
                    int ListID = 1;
                    if (item3_1_1.Text != null && item3_1_1.Text != "" && (item3_1_1.SelectedValue.ToString() == "03" || int.Parse(item3_1_1.SelectedValue.ToString()) > 5))
                    {
                        // 01:新規
                        // 02:契約変更(赤伝)
                        // 03:契約変更(黒伝)
                        // 04:中止
                        // 05:計画
                        // 06:契約変更(黒伝・金額変更)
                        // 07:契約変更(黒伝・工期変更)
                        // 08:契約変更(黒伝・金額工期変更)
                        // 09:契約変更(黒伝・その他)
                        ListID = 352;
                    }
                    string[] result = GlobalMethod.InsertReportWork(ListID, UserInfos[0], new string[] { AnkenID, Header1.Text, "0", "0" });
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
                            form.ExcelName = Path.GetFileName(result[3]);
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
        }


        private void TotalMoney(string name, string lastname, int num, int start = 1)
        {
            long total = 0;
            num += start - 1;
            for (int i = start; i < num; i++)
            {
                total += GetLong(this.Controls.Find(name + i + lastname, true)[0].Text);
            }
            Controls.Find(name + num + lastname, true)[0].Text = GetMoneyTextLong(total);
        }
        private void TotalPercent(string name, string lastname, int num, int start = 1)
        {
            double total = 0.00;
            num += start - 1;
            for (int i = start; i < num; i++)
            {
                total += GetDouble(this.Controls.Find(name + i + lastname, true)[0].Text);
            }
            Controls.Find(name + num + lastname, true)[0].Text = string.Format("{0:F2}", total) + "%";
        }

        //受託金額配分（税込）
        private void Total3_2_1_Leave(object sender, EventArgs e)
        {
            TotalMoney("item3_2_", "_1", 5);
            //税抜算出
            System.Windows.Forms.TextBox text = (System.Windows.Forms.TextBox)sender;
            Control[] cs = this.Controls.Find(text.Name.TrimEnd('1') + "2", true);
            if (cs.Length > 0)
            {
                long zei = GetInt(item3_1_10.Text) + 100;
                ((System.Windows.Forms.TextBox)cs[0]).Text = GetMoneyTextLong(Get_Zeinuki(GetLong(text.Text)));
            }
            TotalMoney("item3_2_", "_2", 5);
            set_keiyaku_haibun();

            //// 全部計算指せる
            //TotalMoney("item3_2_", "_1", 5);
            //TotalMoney("item3_2_", "_2", 5);
            //TotalMoney("item3_7_", "", 5);
            //TotalMoney("item3_3_", "", 5);
            //Total3_6();
            //set_keiyaku_haibun();

            SetKeiyakuHaibunKingaku();

        }
        private void Total3_2_2_Leave(object sender, EventArgs e)
        {
            TotalMoney("item3_2_", "_2", 5);
            set_keiyaku_haibun();

            // 全部計算指せる
            TotalMoney("item3_2_", "_1", 5);
            TotalMoney("item3_2_", "_2", 5);
            TotalMoney("item3_7_", "", 5);
            TotalMoney("item3_3_", "", 5);
            Total3_6();
            set_keiyaku_haibun();
        }

        private void set_keiyaku_haibun()
        {
            //業務配分にコピー
            item3_7_1_6_1.Text = item3_2_1_2.Text;
            item3_7_1_7_1.Text = item3_2_2_2.Text;
            item3_7_1_8_1.Text = item3_2_3_2.Text;
            item3_7_1_9_1.Text = item3_2_4_2.Text;
            item3_7_1_10_1.Text = item3_2_5_2.Text;
        }
        private void Total3_7_Leave(object sender, EventArgs e)
        {
            TotalMoney("item3_7_", "", 5);

            //// 全部計算指せる
            //TotalMoney("item3_2_", "_1", 5);
            //TotalMoney("item3_2_", "_2", 5);
            //TotalMoney("item3_7_", "", 5);
            //TotalMoney("item3_3_", "", 5);
            //Total3_6();
            //set_keiyaku_haibun();
        }

        private void Total3_3_Leave(object sender, EventArgs e)
        {
            TotalMoney("item3_3_", "", 5);

            //// 全部計算指せる
            //TotalMoney("item3_2_", "_1", 5);
            //TotalMoney("item3_2_", "_2", 5);
            //TotalMoney("item3_7_", "", 5);
            //TotalMoney("item3_3_", "", 5);
            //Total3_6();
            //set_keiyaku_haibun();
        }

        private void Total3_6_TextChanged(object sender, EventArgs e)
        {
            Total3_6();

            //// 全部計算指せる
            //TotalMoney("item3_2_", "_1", 5);
            //TotalMoney("item3_2_", "_2", 5);
            //TotalMoney("item3_7_", "", 5);
            //TotalMoney("item3_3_", "", 5);
            //Total3_6();
            //set_keiyaku_haibun();
        }

        private void Total3_6_Leave(object sender, EventArgs e)
        {
            Total3_6();

            //// 全部計算指せる
            //TotalMoney("item3_2_", "_1", 5);
            //TotalMoney("item3_2_", "_2", 5);
            //TotalMoney("item3_7_", "", 5);
            //TotalMoney("item3_3_", "", 5);
            //Total3_6();
            //set_keiyaku_haibun();
        }

        private void Total3_6()
        {
            long total = 0;
            total += GetLong(item3_6_2.Text);
            total += GetLong(item3_6_4.Text);
            total += GetLong(item3_6_6.Text);
            total += GetLong(item3_6_8.Text);
            total += GetLong(item3_6_10.Text);
            total += GetLong(item3_6_12.Text);
            item3_6_13.Text = GetMoneyTextLong(total);
        }

        private void Total2_4_1_Leave(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox text = (System.Windows.Forms.TextBox)sender;
            text.Text = GetPercentText(GetDouble(text.Text));

            TotalPercent("item2_4_1_", "_1", 5);
            copy_haibun("item2_4_1_", "item1_7_1_", 5);
            copy_haibun("item2_4_1_", "item3_7_1_", 5);
        }

        private void Total2_4_2_Leave(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox text = (System.Windows.Forms.TextBox)sender;
            text.Text = GetPercentText(GetDouble(text.Text));

            TotalPercent("item2_4_2_", "_1", 13);
            copy_haibun("item2_4_2_", "item1_7_2_", 13);
            copy_haibun("item2_4_2_", "item3_7_2_", 13);
        }


        private void Tota1_7_1_Leave(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox text = (System.Windows.Forms.TextBox)sender;
            text.Text = GetPercentText(GetDouble(text.Text));

            TotalPercent("item1_7_1_", "_1", 5);
            copy_haibun("item1_7_1_", "item2_4_1_", 5);
            copy_haibun("item1_7_1_", "item3_7_1_", 5);
        }

        private void Tota1_7_2_Leave(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox text = (System.Windows.Forms.TextBox)sender;
            text.Text = GetPercentText(GetDouble(text.Text));

            TotalPercent("item1_7_2_", "_1", 13);
            copy_haibun("item1_7_2_", "item2_4_2_", 13);
            copy_haibun("item1_7_2_", "item3_7_2_", 13);
        }
        private void Tota3_7_2_Leave(object sender, EventArgs e)
        {
            TotalPercent("item3_7_2_", "_1", 13, 14);
            SetKeiyakuHaibunKingaku();
        }

        private void item3_7_2_TextChanged(object sender, EventArgs e)
        {
            TotalPercent("item3_7_2_", "_1", 13, 14);
            SetKeiyakuHaibunKingaku();
        }

        private long Get_Zeinuki(long num)
        {
            long zei = GetInt(item3_1_10.Text) + 100;
            long tmp = num * 100;
            long zeinuki = tmp / zei;
            return zeinuki;

            //decimal zei = GetInt(item3_1_10.Text) + 100;
            //decimal tmp = num * 100;
            //decimal zeinuki = tmp / zei;
            //return (long)Math.Round(zeinuki);
        }

        private void SetKeiyakuHaibunKingaku()
        {
            // 受託金額(税込)
            //int total = GetInt(item3_1_15.Text);
            long total = 0;
            //double total = 0;
            // 受託金額（税込）を基に配分額(税抜)に入れる為、消費税率を考慮する
            double zei = GetInt(item3_1_10.Text);
            // 消費税率が入力されていれば
            if (zei > 0)
            {
                //double double_zei = (100 + zei) / 100;
                //// 税抜 = 税込 / (1 + (消費税 / 100)) 
                //total = GetInt(item3_1_15.Text) / double_zei;

                // 受託金額配分（税抜）に計算を合わせる
                string kingaku = GetMoneyTextLong(Get_Zeinuki(GetLong(item3_1_15.Text)));
                total = GetLong(kingaku);
            }
            else
            {
                total = GetLong(item3_1_15.Text);
            }

            // 契約 配分額(税抜)
            total = GetLong(item3_7_1_6_1.Text);

            Control[] cs1;
            Control[] cs2;
            for (int i = 14; i <= 25; i++)
            {
                // 調査部 業務配分別配分 契約 配分額(税抜)
                cs1 = this.Controls.Find("item3_7_2_" + i + "_1", true);
                cs2 = this.Controls.Find("item3_7_2_" + i + "_2", true);
                if (cs1.Length > 0 && cs2.Length > 0)
                {
                    double percent = GetDouble(cs1[0].Text);
                    long haibun = 0;
                    if (total * percent != 0)
                    {
                        //haibun = (int)Math.Ceiling(total * percent / 100);
                        haibun = (long)Math.Round(total * percent / 100);
                    }
                    cs2[0].Text = GetMoneyTextLong(haibun);
                }
            }
            TotalMoney("item3_7_2_", "_2", 13, 14);
        }

        private void copy_haibun(string tab1, string tab2, int num)
        {
            Control[] cs1;
            Control[] cs2;
            for (int i = 1; i <= num; i++)
            {
                cs1 = this.Controls.Find(tab1 + i + "_1", true);
                cs2 = this.Controls.Find(tab2 + i + "_1", true);
                if (cs1.Length > 0 && cs2.Length > 0)
                {
                    if (cs2[0].GetType().Equals(typeof(System.Windows.Forms.Label)) && cs1[0].GetType().Equals(typeof(System.Windows.Forms.Label)))
                    {
                        ((System.Windows.Forms.Label)cs2[0]).Text = ((System.Windows.Forms.Label)cs1[0]).Text;
                    }
                    else if (cs2[0].GetType().Equals(typeof(System.Windows.Forms.Label)))
                    {
                        ((System.Windows.Forms.Label)cs2[0]).Text = ((System.Windows.Forms.TextBox)cs1[0]).Text;
                    }
                    else if (cs1[0].GetType().Equals(typeof(System.Windows.Forms.Label)))
                    {
                        ((System.Windows.Forms.TextBox)cs2[0]).Text = ((System.Windows.Forms.Label)cs1[0]).Text;
                    }
                    else
                    {
                        ((System.Windows.Forms.TextBox)cs2[0]).Text = ((System.Windows.Forms.TextBox)cs1[0]).Text;
                    }
                }
            }
        }

        private int GetInt(string str)
        {
            int num = 0;
            int.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out num);
            return num;
        }
        private Decimal GetDecimal(string str)
        {
            Decimal num = 0;
            Decimal.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out num);
            return num;
        }
        private double GetDouble(string str)
        {
            double num = 0.00;
            double.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out num);
            return num;
        }

        private string GetMoneyText(int num)
        {
            string str = string.Format("{0:C}", num);
            return str;
        }
        private string GetPercentText(double num)
        {
            if (num > 100)
            {
                num = 100;
            }
            string str = string.Format("{0:F2}", num) + "%";
            return str;
        }
        private long GetLong(string str)
        {
            long num = 0;
            long.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out num);
            return num;
        }
        private string GetMoneyTextLong(long num)
        {
            string str = string.Format("{0:C}", num);
            return str;
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
            e.DrawFocusRectangle();
        }

        private string Get_DateTimePicker(string name)
        {
            Control[] cs;
            cs = this.Controls.Find(name, true);

            if (cs.Length == 0)
            {
                return "null";
            }

            if (((DateTimePicker)cs[0]).CustomFormat != "")
            {
                return "null";
            }

            return ("'" + ((DateTimePicker)cs[0]).Value.ToString() + "'");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(System.Environment.CurrentDirectory + "/Resource/PDF/発注者区分の説明.pdf");
        }

        // この業務を元に新規登録ボタン
        private void button1_Click(object sender, EventArgs e)
        {
            Entry_Input form = new Entry_Input();
            form.mode = "insert";
            form.AnkenID = AnkenID;
            form.UserInfos = this.UserInfos;
            form.CopyMode = "1";
            ownerflg = false;
            form.Show(this.Owner);
            this.Close();
            //Show(this.Owner);
            //this.Close();
        }

        // この案件番号の枝番で受託番号を作成するボタン
        private void button10_Click(object sender, EventArgs e)
        {
            ////受託番号が採番されていない場合は、処理を終了
            //if (item1_8.Text == "")
            //{
            //    set_error("", 0);
            //    set_error("受託番号が採番されていません。落札者を建設物価調査会に設定して更新してください。");
            //    return;
            //}
            Entry_Input form = new Entry_Input();
            form.mode = "insert";
            form.AnkenID = AnkenID;
            form.AnkenbaBangou = item1_8.Text;
            form.UserInfos = this.UserInfos;
            form.CopyMode = "2";
            ownerflg = false;
            //form.Show(this);
            //this.Hide();
            form.Show(this.Owner);
            this.Close();
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

        // 引合の売上年度 変更
        private void item1_3_TextChanged(object sender, EventArgs e)
        {
            item3_1_5.SelectedValue = item1_3.SelectedValue.ToString();
            //set_combo_shibu(item1_3.SelectedValue.ToString());
            //if (mode == "insert" || mode == "keikaku")
            //{
            //    setFolderPath();
            //    FolderPathCheck();
            //}
        }
        // 契約タブの売上年度
        private void item3_1_5_TextChanged(object sender, EventArgs e)
        {
            if (mode != "insert" && mode != "keikaku")
            {
                item1_3.SelectedValue = item3_1_5.SelectedValue.ToString();
                //set_combo_shibu(item1_3.SelectedValue.ToString());
            }
        }

        private void item1_2_KoukiNendo_TextChanged(object sender, EventArgs e)
        {
            set_combo_shibu(item1_2_KoukiNendo.SelectedValue.ToString());
            if (mode == "insert" || mode == "keikaku")
            {
                setFolderPath();
                FolderPathCheck();

                // 工期開始年度に合わせて売上年度を変更する
                // DataSourceにセットした時など、想定外のとこでもTextChangedが動いていたため、値のチェックを入れる
                if (int.TryParse(item1_2_KoukiNendo.SelectedValue.ToString(), out int num))
                {
                    item1_3.SelectedValue = item1_2_KoukiNendo.SelectedValue.ToString();
                }
            }
        }

        // 案件（受託）フォルダの値をセット
        private void setFolderPath()
        {
            // 新規登録時のみ案件（受託）フォルダの値を動的に変更する
            if (mode == "insert" || mode == "keikaku")
            {
                // 案件（受託）フォルダを取得
                string folderPath = item1_12.Text;
                // 年度のパスを調べる 売上年度 の部分
                string keyWord = @"\\2[0-9]{3}\\";
                // 売上年度 の開始位置を取得
                //int cnt = folderPath.IndexOf(keyWord);
                int cnt = 0;
                string targetWord = "";

                System.Text.RegularExpressions.MatchCollection matche = System.Text.RegularExpressions.Regex.Matches(folderPath, keyWord);
                foreach (System.Text.RegularExpressions.Match m in matche)
                {
                    targetWord = m.Value;
                }

                if (targetWord != "")
                {
                    // \2020\ を\売上年度\で置換する
                    //folderPath = folderPath.Replace(targetWord.ToString(), "\\" + item1_3.SelectedValue.ToString() + "\\");
                    //  \2020\ を\工期開始年度\で置換する
                    folderPath = folderPath.Replace(targetWord.ToString(), "\\" + item1_2_KoukiNendo.SelectedValue.ToString() + "\\");

                    // 867
                    // 工期開始年度　2021年度まで、　010北道
                    // 工期開始年度　2022年度から　　010北海
                    int koukinendo = 0;
                    if (int.TryParse(item1_2_KoukiNendo.SelectedValue.ToString(), out koukinendo))
                    {
                        if (koukinendo > 2021)
                        {
                            // 010北道
                            string str1 = GlobalMethod.GetCommonValue1("MADOGUCHI_HOKKAIDO_PATH");
                            // 010北海
                            string str2 = GlobalMethod.GetCommonValue2("MADOGUCHI_HOKKAIDO_PATH");

                            if (str1 != null && str2 != null)
                            {
                                folderPath = folderPath.Replace(str1, str2);
                            }
                        }
                    }



                    item1_12.Text = folderPath;
                }
            }
        }

        private void pictureBox17_Click(object sender, EventArgs e)
        {
            // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
            if (System.Text.RegularExpressions.Regex.IsMatch(item4_1_8.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                // 指定されたフォルダパスが存在するなら開く
                if (item4_1_8.Text != "" && item4_1_8.Text != null && Directory.Exists(item4_1_8.Text))
                {
                    System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item4_1_8.Text));
                }
            }
        }
        // 契約タブ 契約図書
        private void pictureBox7_Click(object sender, EventArgs e)
        {
            //System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item3_1_26.Text));
            // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
            if (System.Text.RegularExpressions.Regex.IsMatch(item3_1_26.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                // 指定されたフォルダパスが存在するなら開く
                if (item3_1_26.Text != "" && item3_1_26.Text != null && Directory.Exists(item3_1_26.Text))
                {
                    System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(item3_1_26.Text));
                }
            }
        }

        private void item3_1_3_Leave(object sender, EventArgs e)
        {
            if (item3_1_3.CustomFormat != "")
            {
                // 契約締結変更日を入力しないと消費税率を取得できません。
                set_error(GlobalMethod.GetMessage("I10713", ""));
            }
            else
            {
                string where = "(TaxStartDay <= '" + item3_1_3.Text + "' ) AND (ISNULL(TaxEndDay,'9999/12/31') >= '" + item3_1_3.Text + "' ) " +
                                " AND TaxKuni = 'JPN' AND ISNULL(TaxDeleteFlag,0) = 0 ";
                DataTable dt = GlobalMethod.getData("TaxZeiritsu", "TaxZeiritsu", "M_Tax", where);

                if (dt != null && dt.Rows.Count > 0)
                {
                    item3_1_10.Text = dt.Rows[0][0].ToString();
                    // 小数点以下を削り取る
                    int comma = item3_1_10.Text.IndexOf(".");
                    item3_1_10.Text = item3_1_10.Text.Substring(0, comma);
                }
                else
                {
                    where = "(TaxStartDay IS null OR TaxStartDay <= '" + item3_1_3.Text + "' ) AND (TaxEndDay IS null OR TaxEndDay >= '" + item3_1_3.Text + "' ) " +
                                " AND TaxKuni = 'JPN' AND ISNULL(TaxDeleteFlag,0) = 0 ";
                    dt.Clear();
                    dt = GlobalMethod.getData("TaxZeiritsu", "TaxZeiritsu", "M_Tax", where);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        item3_1_10.Text = dt.Rows[0][0].ToString();
                    }
                    else
                    {
                        item3_1_10.Text = "0";
                    }
                }
            }
        }

        private void item3_1_6_Leave(object sender, EventArgs e)
        {
            //if (item3_1_6.CustomFormat == "" && item3_1_7.CustomFormat == "")
            //{
            //    set_error("", 0);
            //    if (item3_1_6.Value > item3_1_7.Value)
            //    {
            //        set_error(GlobalMethod.GetMessage("E10011", "(契約工期 開始・終了)"));
            //        item3_1_6.CustomFormat = " ";
            //    }
            //}
        }

        // 契約工期至
        private void item3_1_7_Leave(object sender, EventArgs e)
        {
            if (item3_1_6.CustomFormat == "" && item3_1_7.CustomFormat == "")
            {
                // No.204 工期末日付のコピーボタンが反応しない
                // エラーメッセージが消えることでボタンが押せていないので、エラーでない場合はメッセージを消さないように修正
                //set_error("", 0);
                if (item3_1_6.Value > item3_1_7.Value)
                {
                    set_error("", 0);
                    set_error(GlobalMethod.GetMessage("E10011", "(契約工期 開始・終了)"));
                    item3_1_7.CustomFormat = " ";
                }
            }

            if (item3_1_7.CustomFormat == "")
            {
                DataTable dt = GlobalMethod.getData("NendoID", "NendoID", "Mst_Nendo", "Nendo_Sdate <= '" + item3_1_7.Text + "' AND Nendo_EDate >= '" + item3_1_7.Text + "' ");
                if (dt != null && dt.Rows.Count > 0)
                {
                    // 売上年度
                    item3_1_5.SelectedValue = dt.Rows[0][0].ToString();
                    // 変更伝票以外の場合、引合タブの売上年度も更新する
                    if (mode != "change")
                    {
                        item1_3.SelectedValue = dt.Rows[0][0].ToString();
                    }
                }
            }
        }

        // 消費税率
        private void item3_1_10_TextChanged(object sender, EventArgs e)
        {
            if (item3_1_10.Text == "0")
            {
                item3_1_15.ReadOnly = false;
            }
            else
            {
                item3_1_15.ReadOnly = true;
            }
            calc_kingaku();
            //// 全部計算する
            //TotalMoney("item3_2_", "_1", 5);
            //TotalMoney("item3_2_", "_2", 5);
            //TotalMoney("item3_7_", "", 5);
            //TotalMoney("item3_3_", "", 5);
            //Total3_6();
            //set_keiyaku_haibun();
            //int kingaku = GetInt(item3_1_13.Text) - GetInt(item3_1_16.Text);
            //item3_1_15.Text = GetMoneyText(kingaku);

        }

        //private void item3_1_12_Leave(object sender, EventArgs e)
        //{
        //    if (item3_1_10.Text != "" && item3_1_10.Text != "0")
        //    {
        //        //int zeinuki = int.Parse(item3_1_12.Text.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty));
        //        //int syouhizeiritu = int.Parse(item3_1_10.Text);
        //        //int syouhizei = syouhizeiritu * zeinuki / 100;
        //        //int zeikomi = zeinuki + syouhizei;

        //        //item3_1_13.Text = string.Format("{0:C}", zeikomi);
        //        //item3_1_14.Text = string.Format("{0:C}", syouhizei);
        //        calc_kingaku();
        //    }
        //}


        // 消費税率が入力されていれば、
        // 税抜（自動計算用）を基に
        // 税込と内消費税を計算して表示する
        private void calc_kingaku()
        {
            //int zeinuki = int.Parse(item3_1_12.Text.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty));
            //int syouhizeiritu = 0;
            //if (item3_1_10.Text != null && item3_1_10.Text != "" && item3_1_10.SelectedText != "") { 
            //    syouhizeiritu = int.Parse(item3_1_10.SelectedText);
            //}
            //int syouhizei = syouhizeiritu * zeinuki / 100;
            //int zeikomi = zeinuki + syouhizei;

            //item3_1_13.Text = string.Format("{0:C}", zeikomi);
            //item3_1_14.Text = string.Format("{0:C}", syouhizei);
            if (item3_1_10.Text != "" && item3_1_10.Text != "0")
            {
                long zeinuki = GetLong(item3_1_12.Text);

                int i = 0;
                // 数値に変換できるか確認
                if (Int32.TryParse(item3_1_10.Text, out i))
                {
                    long syouhizeiritu = long.Parse(item3_1_10.Text);
                    long syouhizei = syouhizeiritu * zeinuki / 100;
                    long zeikomi = zeinuki + syouhizei;

                    item3_1_13.Text = string.Format("{0:C}", zeikomi);
                    item3_1_14.Text = string.Format("{0:C}", syouhizei);
                }

            }
        }

        private void item3_1_15_Leave(object sender, EventArgs e)
        {
            SetKeiyakuHaibunKingaku();
        }

        // 計画番号クリアボタン
        private void pictureBox3_Click_1(object sender, EventArgs e)
        {
            if (MessageBox.Show("計画情報を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                // 計画番号、計画案件名のクリア
                item1_4.Text = "";
                item1_5.Text = "";
                item1_4.Focus();
            }
        }

        // 契約タブの5.管理者・担当者の×ボタン
        private void pictureBox18_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("管理技術者を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                item3_4_1.Text = "";
                item3_4_1_CD.Text = "";
                item4_1_2.Text = "";
                item3_4_1.Focus();
            }
        }

        // 契約タブの照査技術者の×ボタン
        private void pictureBox19_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("照査技術者を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                item3_4_2.Text = "";
                item3_4_2_CD.Text = "";
                item4_1_4.Text = "";
                item3_4_2.Focus();
            }
        }

        // 契約タブの審査技術者の×ボタン
        private void pictureBox20_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("審査技術者を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                item3_4_3.Text = "";
                item3_4_3_CD.Text = "";
                item3_4_3.Focus();
            }
        }

        // 契約タブの業務担当者の×ボタン
        private void pictureBox21_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("業務担当者を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                item3_4_4.Text = "";
                item3_4_4_CD.Text = "";
                item3_4_4.Focus();
            }
        }

        // 契約タブの窓口担当者の×ボタン
        private void pictureBox22_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("窓口担当者を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                item3_4_5.Text = "";
                item3_4_5_CD.Text = "";
                item3_4_5.Focus();
            }
        }

        // 受託金額（税込）自動計算
        private void JutakuKingaku_TextChanged(object sender, EventArgs e)
        {
            long kingaku = GetLong(item3_1_13.Text) - GetLong(item3_1_16.Text);
            item3_1_15.Text = GetMoneyTextLong(kingaku);
        }
        //案件フォルダパスの存在確認とアイコン変更
        private void item1_12_Leave(object sender, EventArgs e)
        {
            FolderPathCheck();
        }

        private void FolderPathCheck()
        {
            // 案件（受託）フォルダ
            if (Directory.Exists(item1_12.Text))
            {
                item1_12_icon.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
            }
            else
            {
                item1_12_icon.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
            }
            // 契約図書
            if (Directory.Exists(item3_1_26.Text))
            {
                pictureBox7.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
            }
            else
            {
                pictureBox7.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
            }
            // 請求書
            if (Directory.Exists(item4_1_8.Text))
            {
                pictureBox17.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
            }
            else
            {
                pictureBox17.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
            }

        }

        private void c1FlexGrid2_AfterEdit(object sender, RowColEventArgs e)
        {
            if (c1FlexGrid2.GetCellCheck(e.Row, 3) == C1.Win.C1FlexGrid.CheckEnum.Checked)
            {
                // c1FlexGrid2.Rows[e.Row][5]がNullの場合に、エラーになるので回避
                if (c1FlexGrid2.Rows[e.Row][5] != null)
                {
                    item2_3_7.Text = c1FlexGrid2.Rows[e.Row][5].ToString();
                    if (c1FlexGrid2.Rows[e.Row][6] == null || c1FlexGrid2.Rows[e.Row][6].ToString() == "")
                    {
                        c1FlexGrid2.Rows[e.Row][6] = 0;
                    }
                    item2_3_8.Text = GetMoneyTextLong(GetLong(c1FlexGrid2.Rows[e.Row][6].ToString()));
                }
            }
        }
        private void item3_1_10_Leave(object sender, EventArgs e)
        {
            long kingaku = GetLong(item3_1_13.Text) - GetLong(item3_1_16.Text);
            // 受託金額（税込）
            item3_1_15.Text = GetMoneyTextLong(kingaku);

            //消費税変更時
            //受託配分(税抜)を設定
            item3_2_1_2.Text = GetMoneyTextLong(Get_Zeinuki(GetLong(item3_2_1_1.Text)));
            item3_2_2_2.Text = GetMoneyTextLong(Get_Zeinuki(GetLong(item3_2_2_1.Text)));
            item3_2_3_2.Text = GetMoneyTextLong(Get_Zeinuki(GetLong(item3_2_3_1.Text)));
            item3_2_4_2.Text = GetMoneyTextLong(Get_Zeinuki(GetLong(item3_2_4_1.Text)));
            TotalMoney("item3_2_", "_2", 5);
            set_keiyaku_haibun();

        }

        private void item3_1_15_TextChanged(object sender, EventArgs e)
        {
            SetKeiyakuHaibunKingaku();
        }
        private void c1FlexGrid2_CellChecked(object sender, RowColEventArgs e)
        {

            if (e.Col == 3 & e.Row > 0)
            {
                var _row = e.Row;
                var _col = e.Col;
                for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
                {
                    if (_row != i)
                    {
                        c1FlexGrid2.SetCellCheck(i, 3, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                    }
                    else
                    {
                        if (c1FlexGrid2.GetCellCheck(i, 3) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                        {
                            if (c1FlexGrid2.Rows[i][5] != null && c1FlexGrid2.Rows[i][5].ToString() != "")
                            {
                                item2_3_7.Text = c1FlexGrid2.Rows[i][5].ToString();
                            }
                            else
                            {
                                item2_3_7.Text = "";
                            }
                            if (c1FlexGrid2.Rows[i][6] != null)
                            {
                                item2_3_8.Text = GetMoneyTextLong(GetLong(c1FlexGrid2.Rows[i][6].ToString()));
                            }
                            else
                            {
                                item2_3_8.Text = GetMoneyTextLong(0);
                            }
                        }
                        else
                        {
                            item2_3_7.Text = "";
                            item2_3_8.Text = GetMoneyTextLong(0);
                        }
                    }
                }
            }
        }
        private void AnkenMemo_TextChanged(object sender, EventArgs e)
        {
            string str = ((System.Windows.Forms.TextBox)sender).Text;
            if (mode != "insert" && mode != "change")
            {
                item1_18.Text = str;
                item2_1_6.Text = str;
                item3_1_18.Text = str;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Entry_keikaku_Search form = new Entry_keikaku_Search();
            form.UserInfos = UserInfos;
            form.Show();
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Entry_Search form = new Entry_Search();
            form.UserInfos = UserInfos;
            ownerflg = false;
            form.Show();
            this.Close();
        }
        private void item1_24_TextChanged(object sender, EventArgs e)
        {
            item3_1_9.Text = item1_23.Text + " " + item1_24.Text;
        }

        private void item3_1_12_TextChanged(object sender, EventArgs e)
        {
            if (item3_1_10.Text != "" && item3_1_10.Text != "0")
            {
                calc_kingaku();
            }
        }

        // 引合タブの参考見積額
        private void SankouMitsumoriGaku_TextChanged(object sender, EventArgs e)
        {
            item2_2_4.Text = GetMoneyTextLong(GetLong(item1_36.Text));
        }
        // 案件受託フォルダ
        private void folder_TextChanged(object sender, EventArgs e)
        {
            set_folder();
            //// 案件（受託）フォルダをコピー
            //// 契約タブ 契約図書
            //item3_1_26.Text = item1_12.Text;

            //string JigyoubuHeadCD = "";
            //if (item1_10.Text != null && item1_10.Text != null) { 
            //    string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

            //    using (var conn = new SqlConnection(connStr))
            //    {
            //        conn.Open();
            //        var cmd = conn.CreateCommand();
            //        var dt = new System.Data.DataTable();
            //        //SQL生成
            //        cmd.CommandText = "SELECT " +
            //          "JigyoubuHeadCD " +
            //          "FROM " + "Mst_Busho " +
            //          "WHERE GyoumuBushoCD = '" + item1_10.SelectedValue + "' ";

            //        //データ取得
            //        var sda = new SqlDataAdapter(cmd);
            //        sda.Fill(dt);

            //        if (dt.Rows.Count > 0)
            //        {
            //            JigyoubuHeadCD = dt.Rows[0][0].ToString();
            //        }
            //    }
            //}
            //// 他部所の場合は、\02契約関係図書を付けない
            //if (!"T".Equals(JigyoubuHeadCD) && !"".Equals(JigyoubuHeadCD))
            //{
            //    // 他部署の場合、請求書は案件（受託）フォルダと同じ
            //    item4_1_8.Text = item1_12.Text;
            //}
            //else
            //{
            //    // 技術担当者 請求書
            //    item4_1_8.Text = item1_12.Text + @"\02契約関係図書";
            //}
        }

        // フォルダセット
        private void set_folder()
        {
            // 案件（受託）フォルダをコピー
            // 契約タブ 契約図書
            item3_1_26.Text = item1_12.Text;

            string JigyoubuHeadCD = "";
            if (item1_10.Text != null && item1_10.Text != null)
            {
                string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    var dt = new System.Data.DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      "JigyoubuHeadCD " +
                      "FROM " + "Mst_Busho " +
                      "WHERE GyoumuBushoCD = '" + item1_10.SelectedValue + "' ";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        JigyoubuHeadCD = dt.Rows[0][0].ToString();
                    }
                }
            }
            // 他部所の場合は、\02契約関係図書を付けない
            if (!"T".Equals(JigyoubuHeadCD) && !"".Equals(JigyoubuHeadCD))
            {
                // 他部署の場合、請求書は案件（受託）フォルダと同じ
                item4_1_8.Text = item1_12.Text;
            }
            else
            {
                // 技術担当者 請求書
                item4_1_8.Text = item1_12.Text + @"\02契約関係図書";
            }
        }

        // 売上年度のマウスホイールイベント
        private void item1_3_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }
        // マウスホイールイベントでコンボ値が変わらないように
        private void item_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }
        // 数値項目のコピーとペースト
        private void item_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == (Keys.Control | Keys.C))
            {
                // Ctrl + C
                ((System.Windows.Forms.TextBox)sender).Copy();
            }
            else if (e.KeyData == (Keys.Control | Keys.V))
            {
                // Ctrl + V
                ((System.Windows.Forms.TextBox)sender).Paste();
            }
        }

        // 引合タブ 入札（予定）日
        private void item1_16_CloseUp(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";
            item2_1_3.Text = item1_16.Text;

        }
        // 引合タブ 入札（予定）日
        private void item2_1_3_CloseUp(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";
            item1_16.Text = item2_1_3.Text;
        }

        // 業務名称
        private void GyoumuMei_TextChanged(object sender, EventArgs e)
        {
            string str = ((System.Windows.Forms.TextBox)sender).Text;
            if (mode != "insert" && mode != "change")
            {
                item1_13.Text = str;
                item3_1_11.Text = str;
            }
        }
        //ヘッダー「計画」ボタン押下処理
        private void button5_Click_1(object sender, EventArgs e)
        {
            Entry_keikaku_Search form = new Entry_keikaku_Search();
            form.UserInfos = UserInfos;
            form.Show();
            this.Close();
        }

        // ヘッダーの案件ボタン
        private void button4_Click_1(object sender, EventArgs e)
        {
            Entry_Search form = new Entry_Search();
            form.UserInfos = this.UserInfos;
            form.Show();
            this.Close();
        }
        private void item1_10_TextChanged(object sender, EventArgs e)
        {
            BushoCD = item1_10.SelectedValue.ToString();
        }

        // 受託課所支部
        private void Jutakukashosibu_TextChanged(object sender, EventArgs e)
        {
            // フォルダ設定しなおし
            set_folder();
        }

        private void item1_4_TextChanged(object sender, EventArgs e)
        {
            //計画番号選択時、計画情報の契約区分をセット
            if (item1_4.Text != "")
            {
                DataTable dt = GlobalMethod.getData("KeikakuGyoumuKubunMei", "KeikakuGyoumuKubun", "KeikakuJouhou", "KeikakuDeleteFlag <> 1 AND KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_4.Text + "'");
                if (dt != null && dt.Rows.Count > 0 && dt.Rows[0][1] != null)
                {
                    item1_14.SelectedValue = dt.Rows[0][0].ToString();
                }
            }
        }

        // KeyUp / KeyDown / ↑ / ↓
        // Form の KeyPreview を true に設定すること
        private void Entry_InputKeyDown(object sender, KeyEventArgs e)
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
                    if ("引合".Equals(tabName))
                    {
                        this.tabPage1.AutoScrollPosition = new System.Drawing.Point(-this.tabPage1.AutoScrollPosition.X, -this.tabPage1.AutoScrollPosition.Y + 600);
                    }
                    if ("入札".Equals(tabName))
                    {
                        this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y + 600);
                    }
                    if ("契約".Equals(tabName))
                    {
                        this.tabPage4.AutoScrollPosition = new System.Drawing.Point(-this.tabPage4.AutoScrollPosition.X, -this.tabPage4.AutoScrollPosition.Y + 600);
                    }
                    if ("技術者評価".Equals(tabName))
                    {
                        this.tabPage6.AutoScrollPosition = new System.Drawing.Point(-this.tabPage6.AutoScrollPosition.X, -this.tabPage6.AutoScrollPosition.Y + 600);
                    }
                }
                if (e.KeyCode == Keys.PageUp)
                {
                    if ("引合".Equals(tabName))
                    {
                        this.tabPage1.AutoScrollPosition = new System.Drawing.Point(-this.tabPage1.AutoScrollPosition.X, -this.tabPage1.AutoScrollPosition.Y - 600);
                    }
                    if ("入札".Equals(tabName))
                    {
                        this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y - 600);
                    }
                    if ("契約".Equals(tabName))
                    {
                        this.tabPage4.AutoScrollPosition = new System.Drawing.Point(-this.tabPage4.AutoScrollPosition.X, -this.tabPage4.AutoScrollPosition.Y - 600);
                    }
                    if ("技術者評価".Equals(tabName))
                    {
                        this.tabPage6.AutoScrollPosition = new System.Drawing.Point(-this.tabPage6.AutoScrollPosition.X, -this.tabPage6.AutoScrollPosition.Y - 600);
                    }
                }
            }
            else
            {
                // タブのタイトルを取得 引合、入札、契約、技術者評価
                string tabName = this.tab.SelectedTab.Text;
                if (e.KeyCode == Keys.PageDown || e.KeyCode == Keys.Down)
                {
                    if ("引合".Equals(tabName))
                    {
                        this.tabPage1.AutoScrollPosition = new System.Drawing.Point(-this.tabPage1.AutoScrollPosition.X, -this.tabPage1.AutoScrollPosition.Y + 600);
                    }
                    if ("入札".Equals(tabName))
                    {
                        this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y + 600);
                    }
                    if ("契約".Equals(tabName))
                    {
                        this.tabPage4.AutoScrollPosition = new System.Drawing.Point(-this.tabPage4.AutoScrollPosition.X, -this.tabPage4.AutoScrollPosition.Y + 600);
                    }
                    if ("技術者評価".Equals(tabName))
                    {
                        this.tabPage6.AutoScrollPosition = new System.Drawing.Point(-this.tabPage6.AutoScrollPosition.X, -this.tabPage6.AutoScrollPosition.Y + 600);
                    }
                }
                if (e.KeyCode == Keys.PageUp || e.KeyCode == Keys.Up)
                {
                    if ("引合".Equals(tabName))
                    {
                        this.tabPage1.AutoScrollPosition = new System.Drawing.Point(-this.tabPage1.AutoScrollPosition.X, -this.tabPage1.AutoScrollPosition.Y - 600);
                    }
                    if ("入札".Equals(tabName))
                    {
                        this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y - 600);
                    }
                    if ("契約".Equals(tabName))
                    {
                        this.tabPage4.AutoScrollPosition = new System.Drawing.Point(-this.tabPage4.AutoScrollPosition.X, -this.tabPage4.AutoScrollPosition.Y - 600);
                    }
                    if ("技術者評価".Equals(tabName))
                    {
                        this.tabPage6.AutoScrollPosition = new System.Drawing.Point(-this.tabPage6.AutoScrollPosition.X, -this.tabPage6.AutoScrollPosition.Y - 600);
                    }
                }
            }
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        // 入札タブの入札参加者Grid
        private void c1FlexGrid2_BeforeEdit(object sender, RowColEventArgs e)
        {
            switch (e.Col)
            {
                // 応札額（税抜）
                case 6:
                    c1FlexGrid2.ImeMode = ImeMode.Disable;
                    break;
                default:
                    c1FlexGrid2.ImeMode = ImeMode.Off;
                    break;
            }
        }

        // 契約タブの売上計上情報Grid
        private void c1FlexGrid4_BeforeEdit(object sender, RowColEventArgs e)
        {
            switch (e.Col)
            {
                case 1:
                case 2:
                case 3:
                case 9:
                case 10:
                case 11:
                case 17:
                case 18:
                case 19:
                case 25:
                case 26:
                case 27:
                    c1FlexGrid4.ImeMode = ImeMode.Disable;
                    break;
                default:
                    c1FlexGrid4.ImeMode = ImeMode.Off;
                    break;
            }
        }

        // 技術者評価タブの担当技術者
        private void c1FlexGrid5_BeforeEdit(object sender, RowColEventArgs e)
        {
            switch (e.Col)
            {
                // 評点
                case 3:
                    c1FlexGrid5.ImeMode = ImeMode.Disable;
                    break;
                default:
                    c1FlexGrid5.ImeMode = ImeMode.Off;
                    break;
            }
        }

        private void item1_2_KoukiNendo_SelectedIndexChanged(object sender, EventArgs e)
        {
            set_combo_shibu(item1_2_KoukiNendo.SelectedValue.ToString());
            if (mode == "insert" || mode == "keikaku")
            {
                setFolderPath();
                FolderPathCheck();

                // 工期開始年度に合わせて売上年度を変更する
                // DataSourceにセットした時など、想定外のとこでもTextChangedが動いていたため、値のチェックを入れる
                if (int.TryParse(item1_2_KoukiNendo.SelectedValue.ToString(), out int num))
                {
                    item1_3.SelectedValue = item1_2_KoukiNendo.SelectedValue.ToString();
                }
            }
        }

        // 売上計上情報から指定された部所の計上額の合計を取得する
        private long GetKeijogakuGoukei(string targetBusho)
        {
            long keijogakuKei = 0;
            Boolean nullFlag = true;    // データ空フラグ true:未入力 false;データあり

            // 取得対象の部所の列を設定する
            int j = 0;
            switch (targetBusho)
            {
                case ("調査部"):
                    j = 1;
                    break;
                case ("事業普及部"):
                    j = 9;
                    break;
                case ("情報システム部"):
                    j = 17;
                    break;
                case ("総合研究所"):
                    j = 25;
                    break;
                default:
                    j = 1;  // 0だとエラーになったら困るので、念のため1をセットしておく
                    break;
            }

            // 工期末日付、計上月、計上額の入力チェック
            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)    // ヘッダー分の2行を除く
            {
                // 工期末日付、計上額が空の場合（計上月は工期末日付から自動編集されるため除く）
                if ((c1FlexGrid4.Rows[i][j] != null && c1FlexGrid4.Rows[i][j].ToString() != "")
                    || (c1FlexGrid4.Rows[i][j + 2] != null && c1FlexGrid4.Rows[i][j + 2].ToString() != "0"))
                {
                    nullFlag = false;
                    break;
                }
            }

            // 未入力だった場合、計上額の合計をゼロで返す
            if (nullFlag)
            {
                keijogakuKei = 0;
            }

            // 入力があった場合、計上額の合計を求める
            else
            {
                for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                {
                    // 金額はNULLがあり得るので除外
                    if (c1FlexGrid4[i, j] != null && c1FlexGrid4[i, j].ToString() != "" && c1FlexGrid4.Rows[i][j + 2] != null)
                    {
                        keijogakuKei += GetLong(c1FlexGrid4[i, j + 2].ToString());
                    }
                }
            }

            return keijogakuKei;
        }

        // VIPS 20220415 コンポーネント最新化にあたり修正
        private void c1FlexGrid_OwnerDrawCell(object sender, OwnerDrawCellEventArgs e)
        {
            if (e.Row >= 1 && e.Col == 1)
            {
                e.Image = Img_DeleteRowNonactive;
            }
        }
        private void c1FlexGrid5_OwnerDrawCell(object sender, OwnerDrawCellEventArgs e)
        {
            if (e.Row >= 1 && e.Col == 0)
            {
                e.Image = Img_DeleteRowNonactive;
            }
        }

        //不具合管理表　1310(1028)　↓↓↓
        enum excelIndex : int
        {
            busho_shozoku = 7
            , tantosha = 8
            , mail = 9
            , post_address = 10
            , tel = 11
            , fax = 12
            , irai_gyoumu = 13
            , irai_naiyou = 14
            , rikou_kikan = 22  //2022/06/14 エクセルのフォーマットが変更となり5列追加されたため以降＋５した
            , mitsumori_mokuteki = 25
            , chosa_yotei = 26
            , mitumori_jisseki = 27
            , jouki_shitsumon = 28
        }
        private void btnHanei_Click(object sender, EventArgs e)
        {
            //テキストボックスに、エクセルの行をコピペしたものが入っていることが前提。タブ区切りでセットされる
            //文字がない場合は何もしない
            //タブ数によりどうするか？　要検討

            //配列インデックス


            string textboxBuffer = txtTayoriData.Text;
            string[] words = textboxBuffer.Split('\t');

            //txtTayoriDataに何も入力がない場合は何もしない
            if (textboxBuffer.Trim().Equals(""))
            {
                return;
            }

            //for debug
            //foreach(string word in words)
            //{
            //    Console.WriteLine(word);
            //}

            //前後のダブルクオーテーションを消す。面倒なので配列全体に先にやってしまう。
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = deleteDoubleQuotation(words[i]);
            }

            //部署・所属名
            if (words.Length > (int)excelIndex.busho_shozoku)
            {
                item1_25.Text = words[(int)excelIndex.busho_shozoku];
            }
            //ご担当者名
            if (words.Length > (int)excelIndex.tantosha)
            {
                item1_26.Text = words[(int)excelIndex.tantosha];
            }
            //メールアドレス
            if (words.Length > (int)excelIndex.mail)
            {
                item1_29.Text = words[(int)excelIndex.mail];
            }
            //郵便番号
            if (words.Length > (int)excelIndex.post_address)
            {
                item1_30.Text = getPostAddress(words[(int)excelIndex.post_address], true);
            }
            //住所
            if (words.Length > (int)excelIndex.post_address)
            {
                item1_31.Text = getPostAddress(words[(int)excelIndex.post_address], false);
            }
            //電話番号
            if (words.Length > (int)excelIndex.tel)
            {
                item1_27.Text = getTelNumber(words[(int)excelIndex.tel]);
            }
            //FAX番号
            if (words.Length > (int)excelIndex.fax)
            {
                item1_28.Text = getTelNumber(words[(int)excelIndex.fax]);
            }
            //ご依頼業務名称
            if (words.Length > (int)excelIndex.irai_gyoumu)
            {
                item1_13.Text = words[(int)excelIndex.irai_gyoumu];
            }

            //以降は一つのテキストエリアに:デリミタを付加し、改行なしで結合
            string tmpBuff = "";
            //ご依頼内容の概要
            if (words.Length > (int)excelIndex.irai_naiyou)
            {
                if (words[(int)excelIndex.irai_naiyou].Trim().Equals("") == false)
                {
                    tmpBuff += words[(int)excelIndex.irai_naiyou] + ":";
                }
            }
            //履行期間
            if (words.Length > (int)excelIndex.rikou_kikan)
            {
                if (words[(int)excelIndex.rikou_kikan].Trim().Equals("") == false)
                {
                    tmpBuff += words[(int)excelIndex.rikou_kikan] + ":";
                }
            }
            //見積り依頼の目的
            if (words.Length > (int)excelIndex.mitsumori_mokuteki)
            {
                if (words[(int)excelIndex.mitsumori_mokuteki].Trim().Equals("") == false)
                {
                    tmpBuff += words[(int)excelIndex.mitsumori_mokuteki] + ":";
                }
            }
            //調査の予定時期
            if (words.Length > (int)excelIndex.chosa_yotei)
            {
                if (words[(int)excelIndex.chosa_yotei].Trim().Equals("") == false)
                {
                    tmpBuff += words[(int)excelIndex.chosa_yotei] + ":";
                }
            }
            //見積依頼の実績
            if (words.Length > (int)excelIndex.mitumori_jisseki)
            {
                if (words[(int)excelIndex.mitumori_jisseki].Trim().Equals("") == false)
                {
                    tmpBuff += words[(int)excelIndex.mitumori_jisseki] + ":";
                }
            }
            //上記質問で、「見積依頼の実績あり」と回答された方は、依頼時期、業務概要などについて、記入してください。
            if (words.Length > (int)excelIndex.jouki_shitsumon)
            {
                if (words[(int)excelIndex.jouki_shitsumon].Trim().Equals("") == false)
                {
                    tmpBuff += words[(int)excelIndex.jouki_shitsumon] + ":";
                }
            }
            item1_18.Text = tmpBuff;
        }

        //エクセルのセルに改行コードが入ってるとダブルクォーテーションが付加されてしまうので、消す
        private string deleteDoubleQuotation(string orgBuff)
        {
            //前後のダブルクオーテーションを消す
            return orgBuff.Trim(new char[] { '"' });
        }

        //電話番号データから電話番号と思われるところを返却する。例）0980-53-1212（内線285）→0980-53-1212
        private string getTelNumber(string orgBuff)
        {
            //半角変換
            orgBuff = Microsoft.VisualBasic.Strings.StrConv(orgBuff, Microsoft.VisualBasic.VbStrConv.Narrow, 0x411);
            //「(」が見つかれば、そこで分割し、トリム+半角スペースを削除して返却する
            string[] strArr = orgBuff.Split('(');
            return strArr[0].Trim().Replace(" ", "");
        }
        //郵便番号＋住所のデータから郵便番号と住所のどちらかを返却する
        private string getPostAddress(string orgBuff, bool isPost)
        {
            //ここで半角変換しちゃダメ。住所に全角数字が使われてるので。
            //半角スペースで分割。配列インデックス0に郵便番号、1以降に住所の想定
            string[] strArr = orgBuff.Split(' ');
            if (isPost)
            {
                //半角変換
                string tmpBuff = Microsoft.VisualBasic.Strings.StrConv(strArr[0], Microsoft.VisualBasic.VbStrConv.Narrow, 0x411);
                //郵便番号
                return tmpBuff.Trim().Replace(" ", "").Replace("-", "");
            }
            else
            {
                //配列が半角スペースで分割されていれば再度半角スペースで結合する
                if (strArr.Length > 1)
                {
                    string tmpBuff = "";
                    for (int i = 1; i < strArr.Length; i++)
                    {
                        if (tmpBuff != "")
                        {
                            tmpBuff += " ";
                        }
                        tmpBuff += strArr[i];
                    }
                    return tmpBuff;
                }
                else
                {
                    return "";
                }

            }

        }
        //不具合管理表　1310(1028)　↑↑↑

        private void tab_DrawItem(object sender, DrawItemEventArgs e)
        {
            GlobalMethod.tabDisplaySet(tab, sender, e);
        }

        //エントリ君修正STEP2
        private void label43_Click(object sender, EventArgs e)
        {
            //契約金額（税込）　item3_1_13
            //契約金額（税抜）　item3_1_12
            string sDt = item3_1_7.Text.Trim();
            string sYm = "";
            try
            {
                sYm = DateTime.Parse(sDt).ToString("yyyy/MM");
            }
            catch (Exception)
            {
                // 何もしない
            }
            string GyoumuCD = item3_1_8.SelectedValue.ToString();
            if (GyoumuCD == "1" || GyoumuCD == "2" || GyoumuCD == "3" || GyoumuCD == "4")
            {
                if (sDt != "") c1FlexGrid4.Rows[2][1] = sDt;
                if (sYm != "") c1FlexGrid4.Rows[2][2] = sYm;
                c1FlexGrid4.Rows[2][3] = item3_1_13.Text;
            }
            else if (GyoumuCD == "5" || GyoumuCD == "6")
            {
                if (sDt != "") c1FlexGrid4.Rows[2][9] = sDt;
                if (sYm != "") c1FlexGrid4.Rows[2][10] = sYm;
                c1FlexGrid4.Rows[2][11] = item3_1_13.Text;
            }
            else if (GyoumuCD == "7")
            {
                if (sDt != "") c1FlexGrid4.Rows[2][17] = sDt;
                if (sYm != "") c1FlexGrid4.Rows[2][18] = sYm;
                c1FlexGrid4.Rows[2][19] = item3_1_13.Text;
            }
            else if (GyoumuCD == "8")
            {
                if (sDt != "") c1FlexGrid4.Rows[2][25] = sDt;
                if (sYm != "") c1FlexGrid4.Rows[2][26] = sYm;
                c1FlexGrid4.Rows[2][27] = item3_1_13.Text;
            }
        }

        //エントリ君修正STEP2
        private void label326_Click(object sender, EventArgs e)
        {
            //契約金額（税込）　item3_1_13
            //契約金額（税抜）　item3_1_12
            //配分額（税込）
            item3_2_1_1.Text = item3_1_13.Text;
            TotalMoney("item3_2_", "_1", 5);

            //配分額（税抜）
            item3_2_1_2.Text = item3_1_12.Text;
            TotalMoney("item3_2_", "_2", 5);

            set_keiyaku_haibun();

            // 全部計算指せる
            SetKeiyakuHaibunKingaku();
        }

        //エントリ君修正STEP2
        private void item3_1_27_Leave(object sender, EventArgs e)
        {
            long zeinuki = GetLong(item3_1_12.Text) + GetLong(item3_1_27.Text);
            item3_1_12.Text = string.Format("{0:C}", zeinuki);
        }

        //エントリ君修正STEP2
        private void label51_Click(object sender, EventArgs e)
        {
            // フォルダリネーム========================================================
            string sBushoCd = item1_10.SelectedValue.ToString();//受託課所支部（契約部所）
            string sYear = item1_2_KoukiNendo.SelectedValue.ToString();   // 工期開始年度
            string sGyomu = item1_13.Text;   // 業務名称
            string sGOrder = item1_23.Text;//発注者名
            //if(sBushoCd.Equals(sFolderBushoCDRenameBef) && sYear.Equals(sFolderYearRenameBef) && sGyomu.Equals(sFolderGyomuRenameBef) && sGOrder.Equals(sFolderOrderRenameBef))
            //{
            //    return;
            //}
            sFolderBushoCDRenameBef = sBushoCd;    //受託課所支部（契約部所）
            sFolderYearRenameBef = sYear;   // 工期開始年度
            sFolderGyomuRenameBef = sGyomu;   // 業務名称
            sFolderOrderRenameBef = sGOrder;//発注者名

            // 案件（受託）フォルダ初期値設定 取得
            String discript = "FolderPath";
            String value = "FolderPath ";
            String table = "M_Folder";
            String where = "MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + sBushoCd + "' ";

            // //xxxx/00Cyousa/00調査情報部門共有/$NENDO$/200受託調査関連
            // フォルダ関連は工期開始年度で作成する
            string FolderBase = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", sYear);
            string FolderPath = "";

            DataTable dt = new System.Data.DataTable();
            dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null && dt.Rows.Count > 0)
            {
                // $FOLDER_BASE$/004 本部
                FolderPath = dt.Rows[0][0].ToString();
            }
            if (FolderBase != "" && FolderPath != "")
            {
                FolderPath = FolderPath.Replace(@"$FOLDER_BASE$", FolderBase);
                FolderPath = FolderPath.Replace("/", @"\");

                string jCd = getJigyoubuHeadCD();

                if (jCd.Equals("T"))
                {
                    // 空白はトリム
                    sGOrder = System.Text.RegularExpressions.Regex.Replace(sGOrder, @"\s", "");
                    if (sGOrder.Length > 10)
                    {
                        sGOrder = sGOrder.Substring(0, 10);
                    }
                    sGyomu = System.Text.RegularExpressions.Regex.Replace(sGyomu, @"\s", "");
                    if (sGyomu.Length > 20)
                    {
                        sGyomu = sGyomu.Substring(0, 20);
                    }
                    string ankenNo = item1_6.Text;
                    if (sItem1_10_ori.Equals(item1_10.SelectedValue.ToString()) == false || sItem1_2_KoukiNendo_ori.Equals(item1_2_KoukiNendo.SelectedValue.ToString()) == false)
                    {
                        string jigyoubuHeadCD = "";
                        // 契約区分で業務分類CDを判定
                        // Mst_Jigyoubu に問い合わせる方法が無い為、
                        // 調査部が見つかった場合、T と判断
                        if (item1_14.Text.IndexOf("調査部") > -1)
                        {
                            jigyoubuHeadCD = "T";
                        }
                        else if (item1_14.Text.IndexOf("事業普及部") > -1)
                        {
                            jigyoubuHeadCD = "B";
                        }
                        else if (item1_14.Text.IndexOf("情シス部") > -1)
                        {
                            jigyoubuHeadCD = "J";
                        }
                        else if (item1_14.Text.IndexOf("総合研究所") > -1)
                        {
                            jigyoubuHeadCD = "K";
                        }
                        var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                        using (var conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            var cmd = conn.CreateCommand();

                            // 業務分類CD + 年度下2桁
                            ankenNo = jigyoubuHeadCD + item1_2_KoukiNendo.SelectedValue.ToString().Substring(2, 2);

                            // KashoShibuCD
                            cmd.CommandText = "SELECT  " +
                                    "KashoShibuCD " +

                                    //参照テーブル
                                    "FROM Mst_Busho " +
                                    "WHERE GyoumuBushoCD = '" + item1_10.SelectedValue.ToString() + "' ";
                            var sda = new SqlDataAdapter(cmd);
                            var dtB = new DataTable();
                            sda.Fill(dtB);
                            // KashoShibuCDが正しい
                            ankenNo = ankenNo + dtB.Rows[0][0].ToString();
                        }
                        ankenNo = ankenNo + "●●●";
                    }
                    FolderPath = FolderPath + "\\" + ankenNo + "_" + sGOrder + "_" + sGyomu;
                    item3_1_20_reset_ankenno.Text = ankenNo;
                }
            }
            else
            {
                FolderPath = FolderBase;
                FolderPath = FolderPath.Replace("/", @"\");
            }
            // 案件（受託）フォルダ
            //item1_12.Text = FolderPath;
            txt_renamedfolder.Text = FolderPath;
        }


        //エントリ君修正STEP2:部署のヘッダーマックを取得する
        private string getJigyoubuHeadCD()
        {
            //SQL変数
            string discript = "GyoumuBushoCD";
            string value = "JigyoubuHeadCD";
            string table = "Mst_Busho";
            string where = "GyoumuBushoCD = '" + item1_10.SelectedValue.ToString() + "'";
            //データ取得
            DataTable combodt1 = new System.Data.DataTable();
            combodt1 = GlobalMethod.getData(discript, value, table, where);
            if (combodt1 != null && combodt1.Rows.Count > 0)
            {
                if (combodt1.Rows[0][0] != null)
                    return combodt1.Rows[0][0].ToString();
                else
                    return "";
            }
            else
            {
                return "";
            }
        }
    }
}
