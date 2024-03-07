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
    public partial class Entry_Input_New : Form
    {
        #region 画面遷移変数など -----------------------------------------------------
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

        /// <summary>
        /// 画面遷移モード
        /// </summary>
        public enum MODE
        {
            /// <summary>
            /// 「計画」から「新規登録」
            /// 「計画」から「前回案件番号を元に新規登録」
            /// </summary>
            PLAN = 0,

            /// <summary>
            /// この業務を元に新規登録
            /// この発注者を元に新規登録
            /// </summary>
            INSERT,

            /// <summary>
            /// 検索明細の編集ボタンからの編集
            /// </summary>
            UPDATE,

            /// <summary>
            /// 閲覧
            /// 2:システム管理者以外のユーザ新規登録後、
            /// 　所属部署と受託課所支部が一致しない時
            /// </summary>
            VIEW,

            /// <summary>
            /// 伝票変更
            /// </summary>
            CHANGE,

            /// <summary>
            /// 空
            /// </summary>
            SPACE
        }

        /// <summary>
        /// 過去案件からコピーで新規登録の区分
        /// </summary>
        public enum COPY
        {
            /// <summary>
            /// この業務を元に新規登録
            /// </summary>
            GM = 0,

            /// <summary>
            /// この発注者を元に新規登録
            /// </summary>
            HC,

            /// <summary>
            /// コピーなし
            /// </summary>
            NO,

            /// <summary>
            /// この案件番号の枝番で受託番号を作成する
            /// </summary>
            ED
        }

        /// <summary>
        /// 画面遷移モード
        /// </summary>
        public MODE mode = MODE.SPACE;
        /// <summary>
        /// 新規登録時：コピーモード
        /// </summary>
        public COPY copy = COPY.NO;
        /// <summary>
        /// 案件ID
        /// </summary>
        public string AnkenID = "";
        /// <summary>
        /// 案件枝番
        /// </summary>
        public string AnkenbaBangou = "";
        /// <summary>
        /// ログインユーザ情報
        /// </summary>
        public string[] UserInfos;

        /// <summary>
        /// 変更伝票がどのボタンから遷移したかのフラグ
        /// </summary>
        public int ChangeFlag = 0;
        /// <summary>
        /// 起案フラグ
        /// </summary>
        private bool KianFLG = false;
        /// <summary>
        /// 計画画面から新規登録
        /// </summary>
        public string KeikakuID = "";
        public string Hatyusya = "";

        /// <summary>
        /// メッセージ内容
        /// </summary>
        private string Message = "";

        /// <summary>
        /// 親か遷移するフラグ
        /// </summary>
        private bool ownerflg = true;

        #endregion

        #region えんとり君画面専用変数など -------------------------------------------
        /// <summary>
        /// クラス名
        /// </summary>
        private string pgmName = "Entry_Input_New";

        /// <summary>
        /// DB接続情報
        /// </summary>
        private string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();

        /// <summary>
        /// DBから検索、更新、削除、追加などの操作オブジェクト
        /// </summary>
        private Entry_Input_New_DbClass EntryInputDbClass = new Entry_Input_New_DbClass();

        /// <summary>
        /// 案件ヘッダー検索結果
        /// </summary>
        private System.Data.DataTable AnkenData_H = new System.Data.DataTable();

        private System.Data.DataTable AnkenData_N = new System.Data.DataTable();
        /// <summary>
        /// 契約情報
        /// </summary>
        private System.Data.DataTable AnkenData_K = new System.Data.DataTable();
        private System.Data.DataTable AnkenData_G = new System.Data.DataTable();
        /// <summary>
        /// 基本情報：過去案件リスト
        /// </summary>
        private System.Data.DataTable AnkenData_Grid1 = new System.Data.DataTable();
        /// <summary>
        /// 入札：入札参加者リスト
        /// </summary>
        private System.Data.DataTable AnkenData_Grid2 = new System.Data.DataTable();
        /// <summary>
        /// 契約：担当者リスト
        /// </summary>
        private System.Data.DataTable AnkenData_Grid3 = new System.Data.DataTable();
        /// <summary>
        /// 契約：売上計上情報リスト
        /// </summary>
        private System.Data.DataTable AnkenData_Grid4 = new System.Data.DataTable();
        /// <summary>
        /// 契約（技術者評価）：担当技術者リスト
        /// </summary>
        private System.Data.DataTable AnkenData_Grid5 = new System.Data.DataTable();

        /// <summary>
        /// コンポーネント最新化に
        /// </summary>
        private Image Img_DeleteRowNonactive;


        /// <summary>
        /// 共通処理クラスオブジェクト
        /// </summary>
        private GlobalMethod GlobalMethod = new GlobalMethod();

        /// <summary>
        /// 初期値：受託課所支部（契約部所）DB値
        /// </summary>
        private string sJyutakuKasyoSibuCdOri = "";
        /// <summary>
        /// 初期値：工期開始年度DB値
        /// </summary>
        private string sKokiStartYearOri = "";

        #endregion

        #region えんとり君画面専用変数など -------------------------------------------
        private int saishinFLG;
        private bool KianKaijoFLG = false;

        private string BushoCD = "";
        private string c1FlexGrid2Data = "";
        private string beforeKeikakuBangou = "";
        private string sFolderRenameBef = "";    //ファイルを移動するため、変更前のフォルダを保存する
        private string sFolderYearRenameBef = "";   // 工期開始年度
        private string sAnkenSakuseiKubun_ori = ""; // 案件区分変更前の値
        private string sJigyoubuHeadCD_ori = "";    // 事業部ヘッダーコード
        //計画詳細画面の「前回案件番号を元に新規登録」ボタンから遷移してきたときTrue
        public bool isKeikakuAnkenNew = false;

        /// <summary>
        /// 基本情報等一覧タブ　エラーリスト
        /// </summary>
        private bool[] baseErrors = { true,true,true};
        /// <summary>
        /// フォルダ変更ボタンをクリックしたかどうかフラグ
        /// </summary>
		private bool isClickedRenameFolderButton = false;
		#endregion

		#region フォーム　イベント ---------------------------------------------------
		public Entry_Input_New()
        {
            InitializeComponent();

            #region イベント設定 ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
            // 基本情報等一覧 ================================================================================
            // No.1533
            //base_tbl07_3_cmbOen.MouseWheel += cmb_MouseWheel;
            base_tbl03_cmbKokiSalesYear.MouseWheel += cmb_MouseWheel;
            base_tbl03_cmbKokiStartYear.MouseWheel += cmb_MouseWheel;
            base_tbl03_cmbKeiyakuKubun.MouseWheel += cmb_MouseWheel;
            base_tbl02_cmbJyutakuKasyoSibu.MouseWheel += cmb_MouseWheel;
            base_tbl02_cmbAnkenKubun.MouseWheel += cmb_MouseWheel;
            base_tbl09_cmbNotOrderReason.MouseWheel += cmb_MouseWheel;
            base_tbl09_cmbNotOrderStats.MouseWheel += cmb_MouseWheel;
            base_tbl09_cmbOrderIyoku.MouseWheel += cmb_MouseWheel;
            base_tbl09_cmbSankomitumori.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbRakusatuAmtStats.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbRakusatuStats.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbNyusatuStats.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbKinsiNaiyo.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbTokaiOsatu.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbKinsiUmu.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbOrderKubun.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbOrderIyoku.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbSankoMitumori.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbLowestUmu.MouseWheel += cmb_MouseWheel;
            base_tbl10_cmbNyusatuHosiki.MouseWheel += cmb_MouseWheel;
            // 事前打診 ======================================================================================
            // 入札 ==========================================================================================
            // 契約 ==========================================================================================
            // 技術者評価 ====================================================================================
            #endregion 非表示設定 ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

            // メッセージフォントサイズ設定
            this.ErrorMessage.Font = new System.Drawing.Font(this.ErrorMessage.Font.Name, float.Parse(GlobalMethod.GetCommonValue1("DSP_ERROR_FONTSIZE")));
        }

        private void Entry_Input_New_Shown(object sender, EventArgs e)
        {
            this.TopMost = true;
            this.Refresh();
            this.TopMost = false;
        }

        private void Entry_Input_New_Load(object sender, EventArgs e)
        {
            Console.WriteLine("====Entry_Input_New_Load ST================================================");
            //レイアウトロジックを停止する
            this.SuspendLayout();
            #region フォーム情報設定 ------------------------------------------------------
            //タブの文字装飾変更対応
            //文字表示を大きくする場合は、デザイナでTabのItemSize.widthを変更する。窓口、特命課長、自分大臣は、125で設定すると、14ポイントぐらいのサイズでいける
            tab.DrawMode = TabDrawMode.OwnerDrawFixed;

            //バージョン設定
            lblVersion.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                lblBootMode.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            //ログインユーザ情報
            lblLoginInfo.Text = UserInfos[3] + "：" + UserInfos[1];

            // コンポーネント最新化
            Img_DeleteRowNonactive = Image.FromFile("Resource/Image/DeleteRow.gif");
            #endregion

            #region c1FlexGrid　各種設定 --------------------------------------------------
            // 基本情報等一覧 ================================================================================
            // -- １．進捗段階　の　契約の　表示／非表示　設定
            int iHeight = (mode == MODE.INSERT || mode == MODE.PLAN) ? 0 : 32;
            base_tbl01_input.RowStyles[2].Height = iHeight;
            if(iHeight > 0)
            {
                base_tbl01_chkKeiyaku.TabStop = true;
            }
            // -- ２．基本情報 のリネーム情報　表示／非表示　設定
            bool iVisible = (mode == MODE.INSERT || mode == MODE.PLAN) ? false : true;
            setVisibleToRenameFolder(iVisible);

            if (mode != MODE.INSERT && mode != MODE.PLAN)
            {
                // 新規作成以外のみ設定する
                if (mode != MODE.CHANGE) {
                    // 事前打診 
                    // 入札 
                    // -- 昇順降順アイコン設定
                    bid_tbl03_4_c1FlexGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
                    bid_tbl03_4_c1FlexGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

                    // 技術者評価
                    // -- 昇順降順アイコン設定
                    te_c1FlexGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
                    te_c1FlexGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
                }

                // 契約
                // -- ３．案件情報 ---------------------------------
                // -- ５．管理者・担当者 ---------------------------
                // 昇順降順アイコン設定
                ca_tbl05_txtTanto_c1FlexGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
                ca_tbl05_txtTanto_c1FlexGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

                // -- ６．売上計上情報 -----------------------------
                // 昇順降順アイコン設定
                ca_tbl06_c1FlexGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
                ca_tbl06_c1FlexGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

                // タイトル設定
                ca_tbl06_c1FlexGrid.Rows[0].AllowMerging = true;
                ca_tbl06_c1FlexGrid.Rows[2][0] = "1回目";
                ca_tbl06_c1FlexGrid.Rows[1][1] = "工期末日付";
                ca_tbl06_c1FlexGrid.Rows[1][2] = "計上月";
                ca_tbl06_c1FlexGrid.Rows[1][3] = "計上額";
                ca_tbl06_c1FlexGrid.Rows[1][9] = "工期末日付";
                ca_tbl06_c1FlexGrid.Rows[1][10] = "計上月";
                ca_tbl06_c1FlexGrid.Rows[1][11] = "計上額";
                ca_tbl06_c1FlexGrid.Rows[1][17] = "工期末日付";
                ca_tbl06_c1FlexGrid.Rows[1][18] = "計上月";
                ca_tbl06_c1FlexGrid.Rows[1][19] = "計上額";
                ca_tbl06_c1FlexGrid.Rows[1][25] = "工期末日付";
                ca_tbl06_c1FlexGrid.Rows[1][26] = "計上月";
                ca_tbl06_c1FlexGrid.Rows[1][27] = "計上額";

                // 最初に12月分表示する
                int num = 2;

                for (int i = 0; i < 11; i++)
                {
                    num = i + 2;
                    ca_tbl06_c1FlexGrid.Rows.Add();
                    ca_tbl06_c1FlexGrid.Rows[num + 1][0] = num + "回目";
                    Resize_Grid("ca_tbl06_c1FlexGrid");
                }

                C1.Win.C1FlexGrid.CellRange rng = ca_tbl06_c1FlexGrid.GetCellRange(0, 0, 0, 27);
                rng.Style = ca_tbl06_c1FlexGrid.Styles["FixedBumon"];
            }
            #endregion c1FlexGrid　各種設定 ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓


            #region コントロール　各種設定 ------------------------------------------------
            // コンボ内容の設定
            set_combo();

            // Role:2システム管理者 以外の場合、起案解除は非表示
            if (!UserInfos[4].Equals("2")) ca_btnKianKaijyo.Visible = false;
            // 案件IDがある場合、検索を実行する（編集とコピーで新規作成）
            if (AnkenID != "")
            {
                get_data();
            }
            // タグ上部の共通部表示設定　と　タブの表示／非表示設定
            setVisibleHeaders();
            // 各項目の初期設定
            if (mode == MODE.INSERT || mode == MODE.PLAN)
            {
                //base_tbl01_dtpDtPrior.CustomFormat = "";
                //base_tbl01_dtpDtPrior.Value = DateTime.Now;
                //base_tbl01_chkJizendasin.Checked = true;

                //base_tbl01_dtpDtBid.CustomFormat = "";
                //base_tbl01_dtpDtBid.Value = DateTime.Now;
                //base_tbl01_chkNyusatu.Checked = true;

                // 案件区分
                base_tbl02_cmbAnkenKubun.SelectedValue = "01";

                base_tbl03_cmbKokiSalesYear.SelectedValue = GlobalMethod.GetTodayNendo();
                base_tbl03_cmbKokiStartYear.SelectedValue = GlobalMethod.GetTodayNendo();

                // 新規時 受託課所支部を自部所にセット
                if (AnkenID != null && AnkenID != "")
                {
                    DataTable ankenDt = GlobalMethod.getData("AnkenJutakubushoCD", "AnkenJutakubushoCD", "AnkenJouhou", "AnkenJouhouID = " + AnkenID);
                    if (ankenDt != null && ankenDt.Rows.Count > 0)
                    {
                        base_tbl02_cmbJyutakuKasyoSibu.SelectedValue = ankenDt.Rows[0][0].ToString();
                    }
                }
                else
                {
                    base_tbl02_cmbJyutakuKasyoSibu.SelectedValue = UserInfos[2];
                }

                // 案件（受託）フォルダ初期値設定
                if (AnkenbaBangou == "")
                {
                    // 案件（受託）フォルダ
                    base_tbl02_txtAnkenFolder.Text = getBaseFolderPath(UserInfos[2]);
                }

                //元にする案件あり
                if (AnkenID != "")
                {
                    if (AnkenbaBangou == "")
                    {
                        //「この案件を元に新規登録」
                        base_tbl02_txtAnkenNo.Text = "";
                        base_tbl02_txtJyutakuNo.Text = "";
                        base_tbl02_txtJyutakuEdNo.Text = "";
                        base_tbl10_dtpNyusatuDt.Text = "";
                        base_tbl10_dtpNyusatuDt.CustomFormat = " ";
                    }
                    else
                    {
                        //「この案件番号の枝番で受託番号を作成する」
                        base_tbl09_numSankomitumoriAmt.Text = GetMoneyTextLong(0);

                        base_tbl08_c1FlexGrid.Rows.Count = 2;
                        Resize_Grid("base_tbl08_c1FlexGrid");

                        // 過去案件情報のRakusatushaIDに1をセット
                        base_tbl08_c1FlexGrid.Rows[1][16] = 1;
                    }
                    ////「この案件を元に新規登録」「この案件番号の枝番で受託番号を作成する」共通初期化
                    //base_tbl09_cmbSankomitumori.SelectedValue = "1"; //参考見積
                }
                else
                {
                    // 新規時
                    // 過去案件情報のRakusatushaIDに1をセット
                    base_tbl08_c1FlexGrid.Rows[1][16] = 1;
                }
                //計画詳細の「新規案件」ボタン押下時
                if (mode == MODE.PLAN)
                {
                    DataTable Keikakudt = EntryInputDbClass.KeikakuData(KeikakuID);
                    if (Keikakudt != null && Keikakudt.Rows.Count > 0)
                    {
                        DataRow dr = Keikakudt.Rows[0];
                        base_tbl02_txtKeikakuNo.Text = dr["KeikakuBangou"].ToString();//計画番号
                        base_tbl02_txtKeikakuAKName.Text = dr["KeikakuAnkenMei"].ToString();//計画案件名
                        base_tbl03_txtGyomuName.Text = dr["KeikakuAnkenMei"].ToString();//業務名称
                        //契約区分
                        if (string.IsNullOrEmpty(dr["KeikakuGyoumuKubun"].ToString()) == false)
                        {
                            base_tbl03_cmbKeiyakuKubun.SelectedValue = dr["KeikakuGyoumuKubun"].ToString();
                        }
                        //工期自
                        object obj = dr["KeikakuKoukiKaishibi"];
                        if (obj != null && obj.ToString() != "")
                        {
                            base_tbl03_dtpKokiFrom.Text = obj.ToString();
                        }
                        else
                        {
                            base_tbl03_dtpKokiFrom.Text = "";
                            base_tbl03_dtpKokiFrom.CustomFormat = " ";
                        }
                        //工期至
                        obj = dr["KeikakuKoukiShuryoubi"];
                        if (obj != null && obj.ToString() != "")
                        {
                            base_tbl03_dtpKokiTo.Text = obj.ToString();
                        }
                        else
                        {
                            base_tbl03_dtpKokiTo.Text = "";
                            base_tbl03_dtpKokiTo.CustomFormat = " ";
                        }
                        //工期開始年度
                        if (string.IsNullOrEmpty(dr["KeikakuKaishiNendo"].ToString()) == false)
                        {
                            base_tbl03_cmbKokiStartYear.SelectedValue = dr["KeikakuKaishiNendo"].ToString();
                        }
                        //売上年度
                        if (string.IsNullOrEmpty(dr["KeikakuUriageNendo"].ToString()) == false)
                        {
                            base_tbl03_cmbKokiSalesYear.SelectedValue = dr["KeikakuUriageNendo"].ToString();
                        }

                        // ７．業務配分
                        base_tbl07_1_numPercent1.Text = GetPercentText(GetDouble(dr["bmPercent1"].ToString()));
                        base_tbl07_1_numPercent2.Text = GetPercentText(GetDouble(dr["bmPercent2"].ToString()));
                        base_tbl07_1_numPercent3.Text = GetPercentText(GetDouble(dr["bmPercent3"].ToString()));
                        base_tbl07_1_numPercent4.Text = GetPercentText(GetDouble(dr["bmPercent4"].ToString()));
                        GetTotalPercent("base_tbl07_1_numPercent", 5);
                        base_tbl07_2_numPercent1.Text = GetPercentText(GetDouble(dr["percent1"].ToString()));
                        base_tbl07_2_numPercent2.Text = GetPercentText(GetDouble(dr["percent2"].ToString()));
                        base_tbl07_2_numPercent3.Text = GetPercentText(GetDouble(dr["percent3"].ToString()));
                        base_tbl07_2_numPercent4.Text = GetPercentText(GetDouble(dr["percent4"].ToString()));
                        base_tbl07_2_numPercent5.Text = GetPercentText(GetDouble(dr["percent5"].ToString()));
                        base_tbl07_2_numPercent6.Text = GetPercentText(GetDouble(dr["percent6"].ToString()));
                        base_tbl07_2_numPercent7.Text = GetPercentText(GetDouble(dr["percent7"].ToString()));
                        base_tbl07_2_numPercent8.Text = GetPercentText(GetDouble(dr["percent8"].ToString()));
                        base_tbl07_2_numPercent9.Text = GetPercentText(GetDouble(dr["percent9"].ToString()));
                        base_tbl07_2_numPercent10.Text = GetPercentText(GetDouble(dr["percent10"].ToString()));
                        base_tbl07_2_numPercent11.Text = GetPercentText(GetDouble(dr["percent11"].ToString()));
                        base_tbl07_2_numPercent12.Text = GetPercentText(GetDouble(dr["percent12"].ToString()));
                        base_tbl07_2_numPercentAll.Text = GetPercentText(GetDouble(dr["percentAll"].ToString()));
                    }
                }

                //「この案件番号の枝番で受託番号を作成」ボタン押下時
                if (mode == MODE.INSERT && AnkenbaBangou != null && AnkenbaBangou != "")
                {
                    int Eda = 0;
                    string EdaStr = "";

                    // 案件番号の中で、枝番が数値で構成されていて、最大の物を取得する
                    DataTable AnkenEdadt = GlobalMethod.getData("' '", "max(AnkenJutakuBangouEda)", "AnkenJouhou", "AnkenAnkenBangou = '" + base_tbl02_txtAnkenNo.Text + "' AND AnkenJutakuBangouEda LIKE '%[0-9]%' AND AnkenDeleteFlag = 0");
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

                    base_tbl02_txtJyutakuEdNo.Text = string.Format("{0:D2}", Eda);
                    base_tbl02_txtJyutakuNo.Text = base_tbl02_txtAnkenNo.Text + "-" + base_tbl02_txtJyutakuEdNo.Text;
                }
                // フォルダチェック
                FolderPathCheck();
            }
            else
            {
                if (mode != MODE.CHANGE)
                {
                    // 基本情報
                    if (base_tbl02_txtJyutakuNo.Text == "")
                    {
                        if (ca_tbl01_chkKian.Checked)
                        {
                            setVisibleToRenameFolder(false);
                        }
                        else
                        {
                            setVisibleToRenameFolder(true);
                        }
                    }
                    else
                    {
                        setVisibleToRenameFolder(false);
                    }
                }
            }

            // 各タブのコントローラ　表示可否と編集可否設定処理
            setVisibleDetails();

            // IMEモード変更（コントローラのプロパティで設定）
            #endregion


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
                btnNewByCopy.Visible = false;
            }
            if (GlobalMethod.GetCommonValue1("COPYBUTTON_FLAG", "2") == "0")
            {
                btnNewByBranchNo.Visible = false;
            }
            //「この案件の発注者に新規登録」
            if (GlobalMethod.GetCommonValue1("COPYBUTTON_FLAG", "3") == "0")
            {
                btnNewByOrder.Visible = false;
            }

            sFolderRenameBef = base_tbl02_txtAnkenFolder.Text;
            if (base_tbl03_cmbKokiStartYear.SelectedValue == null)
            {
                sFolderYearRenameBef = base_tbl03_cmbKokiStartYear.Text.Substring(0, 4);
            }
            else
            {
                sFolderYearRenameBef = base_tbl03_cmbKokiStartYear.SelectedValue.ToString();   // 工期開始年度
            }

            if(AnkenData_H.Rows.Count > 0)
            {
                base_tbl02_cmbJyutakuKasyoSibu.SelectedValue = AnkenData_H.Rows[0]["AnkenJutakubushoCD"].ToString();
            }
            sJyutakuKasyoSibuCdOri = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue == null ? "" : base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString(); //受託課所支部（契約部所）DB値
            sKokiStartYearOri = sFolderYearRenameBef; //工期開始年度DB値

            bool bVisible = UserInfos[4].Equals("2");
            ca_tbl01_lblKanrisya.Visible = bVisible;
            ca_tbl01_lblSasya.Visible = bVisible;
            ca_tbl01_chkSasya.Visible = bVisible;

            // 各項目の要らないイベント
            setCmbEvent();
            base_tbl01_chkJizendasin.CheckedChanged += CheckBox_CheckedChanged;
            base_tbl01_chkNyusatu.CheckedChanged += CheckBox_CheckedChanged;
            base_tbl01_chkKeiyaku.CheckedChanged += CheckBox_CheckedChanged;

            // 呼び出し親が非表示設定
            this.Owner.Hide();
            Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
            //レイアウトロジックを再開する
            this.ResumeLayout();
            Console.WriteLine(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"));
            Console.WriteLine("====Entry_Input_New_Load ED================================================");
        }

        #region 画面項目へデータ設定処理
        /// <summary>
        /// データ取得処理
        /// </summary>
        private void get_data()
        {
            // 基本情報
            AnkenData_H = EntryInputDbClass.AnkenData_H(AnkenID);

            // 入札（応札情報）
            AnkenData_N = EntryInputDbClass.AnkenData_N(AnkenID);

            // 契約情報取得
            AnkenData_K = EntryInputDbClass.AnkenData_K(AnkenID);

            // 業務情報
            AnkenData_G = EntryInputDbClass.AnkenData_G(AnkenID);

            // 過去案件リスト
            string tokai = GlobalMethod.GetCommonValue2("ENTORY_TOUKAI");
            bool isInsert = (mode == MODE.INSERT || this.isKeikakuAnkenNew) ? true : false;
            AnkenData_Grid1 = EntryInputDbClass.AnkenData_Grid1(tokai, AnkenID, isInsert);

            //入札タブの入札者情報
            AnkenData_Grid2 = EntryInputDbClass.AnkenData_Grid2(AnkenID);

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

            //AnkenData_Grid3 = EntryInputDbClass.AnkenData_Grid3(AnkenID);
            // 売上請求情報
            AnkenData_Grid4 = EntryInputDbClass.AnkenData_Grid4(AnkenID);
            // 担当者リスト
            AnkenData_Grid5 = EntryInputDbClass.AnkenData_Grid5(AnkenID);

            // 取得データを画面へ設定処理
            set_data();

            //No1668 画面表示時に仮作成した案件フォルダフルパスと、案件（受託）フォルダが異なる場合、更新時にフォルダ強制変更になるようにする
            string FolderPath = "";
            string ankenNo = "";
            MakeFolderFullPath(ref FolderPath, ref ankenNo);


			//初期表示時のフォルダ再作成時、"●●●"は現行案件Noとする
			string ankenNoRight3 = GlobalMethod.Right(base_tbl02_txtAnkenNo.Text, 3);

			FolderPath = FolderPath.Replace("●●●", ankenNoRight3);
			ankenNo = ankenNo.Replace("●●●", ankenNoRight3);

			// 案件（受託）フォルダと仮作成したFolderPathが空でない、かつ異なる場合
			if (FolderPath.Length != 0 && base_tbl02_txtAnkenFolder.Text.Length != 0 && FolderPath != base_tbl02_txtAnkenFolder.Text)
            {
                // 変更後の案件（受託）フォルダに仮作成したFolderPathをセット
                base_tbl02_txtRenameFolder.Text = FolderPath;
            }
            //案件（受託）フォルダのフルパス作成時に案件番号が作成されていれば非表示案件Noにセット
            if (ankenNo.Length != 0)
            {
                ca_tbl01_hidResetAnkenno.Text = ankenNo;
            }
        }

        /// <summary>
        /// 画面へデータ設定処理
        /// </summary>
        private void set_data()
        {
            //ヘッダー情報
            if (AnkenData_H != null || AnkenData_H.Rows.Count >=1 )
            {
                tblAKInfo_lblAnkenNo.Text = AnkenData_H.Rows[0]["AnkenAnkenBangou"].ToString();
                tblAKInfo_lblJyutakuNo.Text = AnkenData_H.Rows[0]["AnkenJutakuBangou"].ToString();
                tblAKInfo_lblOrderInfo.Text = AnkenData_H.Rows[0]["HachushaMei"].ToString() + "　" + AnkenData_H.Rows[0]["AnkenHachushaKaMei"].ToString();
                tblAKInfo_lblAnkenName.Text = AnkenData_H.Rows[0]["AnkenGyoumuMei"].ToString();
            }

            // 基本情報等一覧タブ　----
            set_data_base();

            if (mode != MODE.INSERT && mode != MODE.PLAN)
            {
                // 事前打診タブ　----
                set_data_prior();
                // 入札タブ　----
                set_data_bid();
                // 契約　----
                // 技術者評価　----
                set_data_ca();
                FolderPathCheck();
            }
        }

        /// <summary>
        /// 基本情報等一覧タブ へデータ設定
        /// </summary>
        private void set_data_base()
        {
            // 基本情報等一覧タブ　----
            sJigyoubuHeadCD_ori = AnkenData_H.Rows[0]["JigyoubuHdCD"].ToString();

            string sDt = "";
            if (AnkenData_H == null || AnkenData_H.Rows.Count <= 0) return;
            beforeKeikakuBangou = AnkenData_H.Rows[0]["AnkenKeikakuBangou"].ToString(); // 計画番号
            if (copy == COPY.GM || copy == COPY.HC)
            {
                beforeKeikakuBangou = "";
            }
            object obj = null;
            //１．進捗段階　----
            if (mode != MODE.INSERT && mode != MODE.PLAN)
            {
                //■事前打診 登録日
                obj = AnkenData_H.Rows[0]["AnkenJizenDashinDate"];
                if (obj != null && obj.ToString() != "")
                {
                    base_tbl01_dtpDtPrior.Text = obj.ToString();
                }
                else
                {
                    base_tbl01_dtpDtPrior.Text = "";
                    base_tbl01_dtpDtPrior.CustomFormat = " ";
                }
                base_tbl01_chkJizendasin.Checked = (AnkenData_H.Rows[0]["AnkenJizenDashinCheck"].ToString() == "1");
                //■入札 登録日（入札情報）
                obj = AnkenData_H.Rows[0]["AnkenNyuusatuDate"];
                if (obj != null && obj.ToString() != "")
                {
                    base_tbl01_dtpDtBid.Text = obj.ToString();
                }
                else
                {
                    base_tbl01_dtpDtBid.Text = "";
                    base_tbl01_dtpDtBid.CustomFormat = " ";
                }
                base_tbl01_chkNyusatu.Checked = (AnkenData_H.Rows[0]["AnkenNyuusatuCheck"].ToString() == "1");
                //■契約 登録日（契約情報）
                obj = AnkenData_H.Rows[0]["AnkenKeiyakuDate"];
                if (obj != null && obj.ToString() != "")
                {
                    base_tbl01_dtpDtCa.Text = obj.ToString();
                }
                else
                {
                    base_tbl01_dtpDtCa.Text = "";
                    base_tbl01_dtpDtCa.CustomFormat = " ";
                }
                base_tbl01_chkKeiyaku.Checked = (AnkenData_H.Rows[0]["AnkenKeiyakuCheck"].ToString() == "1");
            }
            //２．基本情報　----
            if (copy != COPY.GM && copy != COPY.HC)
            {
                // 案件区分
                base_tbl02_cmbAnkenKubun.SelectedValue = AnkenData_H.Rows[0]["AnkenSakuseiKubun"].ToString();
                // 計画番号
                base_tbl02_txtKeikakuNo.Text = beforeKeikakuBangou;
                // 計画案件名
                base_tbl02_txtKeikakuAKName.Text = AnkenData_H.Rows[0]["KeikakuAnkenMei"].ToString();
                // 案件情報ID
                if (mode != MODE.INSERT && mode != MODE.PLAN)
                {
                    base_tbl02_txtAnkenID.Text = AnkenData_H.Rows[0]["AnkenJouhouID"].ToString();
                }
                // 案件番号
                base_tbl02_txtAnkenNo.Text = AnkenData_H.Rows[0]["AnkenAnkenBangou"].ToString();
                // 受託番号
                base_tbl02_txtJyutakuNo.Text = AnkenData_H.Rows[0]["AnkenJutakuBangou"].ToString();
                base_tbl02_txtJyutakuEdNo.Text = AnkenData_H.Rows[0]["AnkenJutakuBangouEda"].ToString();
            }
            // 受託課所支部 工期開始年度の後ろに設定する（連動するため）
            // 契約担当者
            BushoCD = AnkenData_H.Rows[0]["AnkenJutakubushoCD"].ToString();
            base_tbl02_txtKeiyakuTanto.Text = AnkenData_H.Rows[0]["AnkenTantoushaMei"].ToString();
            base_tbl02_txtKeiyakuTantoCD.Text = AnkenData_H.Rows[0]["AnkenTantoushaCD"].ToString();
            base_tbl02_txtKeiyakuTantoBusho.Text = AnkenData_H.Rows[0]["AnkenTantoushaBushoCD"].ToString();
            if (copy != COPY.GM && copy != COPY.HC)
            {
                // 案件(受託)フォルダ
                base_tbl02_txtAnkenFolder.Text = AnkenData_H.Rows[0]["AnkenKeiyakusho"].ToString();
            }
            if (mode != MODE.INSERT && mode != MODE.PLAN)
            {
                base_tbl02_txtAnkenChanger.Text = AnkenData_H.Rows[0]["ChousainName"].ToString();
                base_tbl02_txtAnkenChangerCD.Text = AnkenData_H.Rows[0]["AnkenFolderRenameTantouCd"].ToString();
                base_tbl02_txtAnkenChangDt.Text = AnkenData_H.Rows[0]["AnkenFolderHenkouDt"].ToString();
                base_tbl02_txtAnkenChangHistory.Text = AnkenData_H.Rows[0]["BefChangeAnkenNo"].ToString();
            }

            //３．案件情報　----
            //業務名称
            base_tbl03_txtGyomuName.Text = AnkenData_H.Rows[0]["AnkenGyoumuMei"].ToString();
            // No1593 「この発注者を元にコピーをする」際に、部門配分と契約区分がコピーされない。
            //契約区分
            base_tbl03_cmbKeiyakuKubun.SelectedValue = AnkenData_H.Rows[0]["AnkenGyoumuKubun"].ToString();
            if (copy != COPY.HC)
            {
                // No1593 「この発注者を元にコピーをする」際に、部門配分と契約区分がコピーされない。
                ////契約区分
                //base_tbl03_cmbKeiyakuKubun.SelectedValue = AnkenData_H.Rows[0]["AnkenGyoumuKubun"].ToString();
                if (copy != COPY.GM)
                {
                    //工期自
                    sDt = AnkenData_H.Rows[0]["KeiyakuKoukiKaishibi"].ToString();
                    if (sDt != "")
                    {
                        base_tbl03_dtpKokiFrom.Text = sDt;
                    }
                    else
                    {
                        base_tbl03_dtpKokiFrom.CustomFormat = " ";
                    }
                    //工期至
                    sDt = AnkenData_H.Rows[0]["KeiyakuKoukiKanryoubi"].ToString();
                    if (sDt != "")
                    {
                        base_tbl03_dtpKokiTo.Text = sDt;
                    }
                    else
                    {
                        base_tbl03_dtpKokiTo.CustomFormat = " ";
                    }
                    //工期開始年度
                    base_tbl03_cmbKokiStartYear.SelectedValue = AnkenData_H.Rows[0]["AnkenKoukiNendo"].ToString();
                    //売上年度
                    base_tbl03_cmbKokiSalesYear.SelectedValue = AnkenData_H.Rows[0]["AnkenUriageNendo"].ToString();
                }
                //案件メモ(基本情報) 
                base_tbl03_txtAnkenMemo.Text = AnkenData_H.Rows[0]["AnkenAnkenMemoKihon"].ToString();
            }
            // 受託課所支部(工期開始年度より、リストが設定するため、工期開始年度の後ろに設定する)
            base_tbl02_cmbJyutakuKasyoSibu.SelectedValue = AnkenData_H.Rows[0]["AnkenJutakubushoCD"].ToString();
            // ４．発注者情報
            // 発注者コード
            base_tbl04_txtOrderCd.Text = AnkenData_H.Rows[0]["AnkenHachushaCD"].ToString();
            // 発注者区分1
            base_tbl04_txtOrderKubun1.Text = AnkenData_H.Rows[0]["HachushaKubun1Mei"].ToString();
            // 発注者区分2
            base_tbl04_txtOrderKubun2.Text = AnkenData_H.Rows[0]["HachushaKubun2Mei"].ToString();
            // 都道府県名
            base_tbl04_txtTodofuken.Text = AnkenData_H.Rows[0]["TodouhukenMei"].ToString();
            // 発注者名
            base_tbl04_txtOrderName.Text = AnkenData_H.Rows[0]["HachushaMei"].ToString();
            // 発注者課名
            base_tbl04_txtOrderKamei.Text = AnkenData_H.Rows[0]["AnkenHachushaKaMei"].ToString();

            // ５．発注担当者情報（調査窓口）
            //部署
            base_tbl05_txtBusho.Text = AnkenData_H.Rows[0]["AnkenHachuushaIraibusho"].ToString();
            //担当者名
            base_tbl05_txtTanto.Text = AnkenData_H.Rows[0]["AnkenHachuushaTantousha"].ToString();
            //電話
            base_tbl05_txtTel.Text = AnkenData_H.Rows[0]["AnkenHachuushaTEL"].ToString();
            //FAX
            base_tbl05_txtFax.Text = AnkenData_H.Rows[0]["AnkenHachuushaFAX"].ToString();
            //E - Mail
            base_tbl05_txtEmail.Text = AnkenData_H.Rows[0]["AnkenHachuushaMail"].ToString();
            //郵便番号
            base_tbl05_txtZip.Text = AnkenData_H.Rows[0]["AnkenHachuushaIraiYuubin"].ToString();
            //住所
            base_tbl05_txtAddress.Text = AnkenData_H.Rows[0]["AnkenHachuushaIraiJuusho"].ToString();

            // ６．発注担当者情報（契約窓口）
            //部署
            base_tbl06_txtBusho.Text = AnkenData_H.Rows[0]["AnkenHachuushaKeiyakuBusho"].ToString();
            //担当者名
            base_tbl06_txtTanto.Text = AnkenData_H.Rows[0]["AnkenHachuushaKeiyakuTantou"].ToString();
            //電話
            base_tbl06_txtTel.Text = AnkenData_H.Rows[0]["AnkenHachuushaKeiyakuTEL"].ToString();
            //FAX
            base_tbl06_txtFax.Text = AnkenData_H.Rows[0]["AnkenHachuushaKeiyakuFAX"].ToString();
            //E - Mail
            base_tbl06_txtEmail.Text = AnkenData_H.Rows[0]["AnkenHachuushaKeiyakuMail"].ToString();
            //郵便番号
            base_tbl06_txtZip.Text = AnkenData_H.Rows[0]["AnkenHachuushaKeiyakuYuubin"].ToString();
            //住所
            base_tbl06_txtAddress.Text = AnkenData_H.Rows[0]["AnkenHachuushaKeiyakuJuusho"].ToString();
            // 発注元代表者_役職
            base_tbl06_txtOrderYakusyoku.Text = AnkenData_H.Rows[0]["AnkenHachuuDaihyouYakushoku"].ToString();
            // 発注元代表者_氏名
            base_tbl06_txtOrderSimei.Text = AnkenData_H.Rows[0]["AnkenHachuuDaihyousha"].ToString();

            // No1593 「この発注者を元にコピーをする」際に、部門配分と契約区分がコピーされない。
            //// ７．業務配分
            //if (copy != COPY.HC)
            //{
            // 部門配分 【事前打診・入札】配分率(%)
            //調査部
            base_tbl07_1_numPercent1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["BuRitsu1"]));
            //事業普及部
            base_tbl07_1_numPercent2.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["BuRitsu2"]));
            //情報システム部
            base_tbl07_1_numPercent3.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["BuRitsu3"]));
            //総合研究所
            base_tbl07_1_numPercent4.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["BuRitsu4"]));
            //合計
            GetTotalPercent("base_tbl07_1_numPercent", 5);
            if (copy != COPY.HC)
            {
                // 調査部 業務別配分 【事前打診・入札】配分率(%)
                base_tbl07_2_numPercent1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu1"]));//資材調査
                base_tbl07_2_numPercent2.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu2"]));//営繕調査
                base_tbl07_2_numPercent3.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu3"]));//機器類調査
                base_tbl07_2_numPercent4.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu4"]));//工事費調査
                base_tbl07_2_numPercent5.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu5"]));//産廃調査
                base_tbl07_2_numPercent6.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu6"]));//歩掛調査
                base_tbl07_2_numPercent7.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu7"]));//諸経費調査
                base_tbl07_2_numPercent8.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu8"]));//原価分析調査
                base_tbl07_2_numPercent9.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu9"]));//準作成改訂
                base_tbl07_2_numPercent10.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu10"]));//公共労務費調査
                base_tbl07_2_numPercent11.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu11"]));//労務費公共以外
                base_tbl07_2_numPercent12.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu12"]));//その他調査部
                GetTotalPercent("base_tbl07_2_numPercent", 13);//合計

                // No.1533
                //// 応援依頼の有無
                //obj = AnkenData_H.Rows[0]["AnkenOueniraiUmu"];
                //if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                //{
                //    base_tbl07_3_cmbOen.SelectedValue = obj.ToString();
                //}
                // 応援依頼メモ
                base_tbl07_3_txtOenMemo.Text = AnkenData_H.Rows[0]["AnkenOuenIraiMemo"].ToString();
                // 応援依頼先(リストは工期開始年度より変更ので、工期開始年度設定後、設定する)
                // 連動箇所でチェック設定
            }
            //８．過去案件情報
            if (copy != COPY.HC)
            {
                // 新規じゃない、またはこの業務を元に新規登録の場合
                if (mode != MODE.INSERT || (mode == MODE.INSERT && AnkenbaBangou == ""))
                {
                    for (int k = 0; k < AnkenData_Grid1.Rows.Count; k++)
                    {
                        if (k > 0)
                        {
                            base_tbl08_c1FlexGrid.Rows.Add();
                        }
                        if (k >= 5)
                        {
                            break;
                        }
                        //【過去案件】
                        // SortKey 以外をc1FlexGridにセットする
                        for (int i = 0; i < AnkenData_Grid1.Columns.Count - 1; i++)
                        {
                            base_tbl08_c1FlexGrid.Rows[k + 1][i + 2] = AnkenData_Grid1.Rows[k][i];
                        }
                    }
                }
                // 過去案件情報がない場合、
                if (AnkenData_Grid1.Rows.Count == 0)
                {
                    // 過去案件情報の前回受託番号IDに1を入れておく
                    base_tbl08_c1FlexGrid.Rows[1][16] = 1;
                }
                Resize_Grid("base_tbl08_c1FlexGrid");
            }
            // ９．事前打診・参考見積
            if (copy != COPY.GM && copy != COPY.HC)
            {
                //事前打診依頼日 AnkenJizenDashinIraibi
                obj = AnkenData_H.Rows[0]["AnkenJizenDashinIraibi"];
                if (obj != null && obj.ToString() != "")
                {
                    base_tbl09_dtpJizenDasinIraiDt.Text = obj.ToString();
                }
                else
                {
                    base_tbl09_dtpJizenDasinIraiDt.Text = "";
                    base_tbl09_dtpJizenDasinIraiDt.CustomFormat = " ";
                }

                //参考見積対応
                obj = AnkenData_H.Rows[0]["AnkenToukaiSankouMitsumori"];
                if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                {
                    base_tbl09_cmbSankomitumori.SelectedValue = obj.ToString();
                }
                //参考見積額(税抜) ISNULL(AnkenToukaiSankouMitsumoriGaku, 0)         AS NyuusatsuMitsumoriAmt
                if (copy != COPY.GM)
                {
                    base_tbl09_numSankomitumoriAmt.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_H.Rows[0]["NyuusatsuMitsumoriAmt"]));
                }
                //受注意欲
                obj = AnkenData_H.Rows[0]["AnkenToukaiJyutyuIyoku"];
                if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                {
                    base_tbl09_cmbOrderIyoku.SelectedValue = obj.ToString();
                }
                //発注予定・見込日
                obj = AnkenData_H.Rows[0]["AnkenHachuuYoteiMikomibi"];
                if (obj != null && obj.ToString() != "")
                {
                    base_tbl09_dtpOrderYoteiDt.Text = obj.ToString();
                }
                else
                {
                    base_tbl09_dtpOrderYoteiDt.Text = "";
                    base_tbl09_dtpOrderYoteiDt.CustomFormat = " ";
                }
                //未発注状況
                obj = AnkenData_H.Rows[0]["AnkenMihachuuJoukyou"];
                if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                {
                    base_tbl09_cmbNotOrderStats.SelectedValue = obj.ToString();
                }
                //「発注無し」の理由
                obj = AnkenData_H.Rows[0]["AnkenHachuunashiRiyuu"];
                if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                {
                    base_tbl09_cmbNotOrderReason.SelectedValue = obj.ToString();
                }
                //「その他」の内容
                base_tbl09_txtOthenComment.Text = AnkenData_H.Rows[0]["AnkenSonotaNaiyou"].ToString();
            }
            //１０．入札情報・入札結果
            if (copy != COPY.HC)
            {
                if (AnkenData_N != null && AnkenData_N.Rows.Count > 0)
                {
                    //業務発注区分
                    obj = AnkenData_N.Rows[0]["NyuusatsuGyoumuHachuukubun"];
                    if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                    {
                        base_tbl10_cmbOrderKubun.SelectedValue = obj.ToString();
                    }
                    //入札方式
                    obj = AnkenData_N.Rows[0]["NyuusatsuHoushiki"];
                    if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                    {
                        base_tbl10_cmbNyusatuHosiki.SelectedValue = obj.ToString();
                    }
                    //最低制限価格有無
                    obj = AnkenData_N.Rows[0]["NyuusatsuSaiteiKakakuUmu"];
                    if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                    {
                        base_tbl10_cmbLowestUmu.SelectedValue = obj.ToString();
                    }
                    if (copy != COPY.GM)
                    {
                        //入札(予定)日
                        sDt = AnkenData_N.Rows[0]["yoteibi"].ToString();
                        if (sDt != "")
                        {
                            base_tbl10_dtpNyusatuDt.Text = sDt;
                        }
                        else
                        {
                            base_tbl10_dtpNyusatuDt.Text = "";
                            base_tbl10_dtpNyusatuDt.CustomFormat = " ";
                        }
                        //参考見積対応
                        obj = AnkenData_N.Rows[0]["NyuusatsuSankoumitsumoriTaiou"];
                        if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                        {
                            base_tbl10_cmbSankoMitumori.SelectedValue = obj.ToString();
                        }
                        //参考見積額(税抜)
                        base_tbl10_numSankoMitumoriAmt.Text = string.Format("{0:C}", GetLong(AnkenData_N.Rows[0]["NyuusatsuSankoumitsumoriKingaku"].ToString()));
                        //受注意欲
                        obj = AnkenData_N.Rows[0]["NyuusatsuJuchuuIyoku"];
                        if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                        {
                            base_tbl10_cmbOrderIyoku.SelectedValue = obj.ToString();
                        }
                        //当会応札 ▼▼▼
                        sDt = AnkenData_N.Rows[0]["AnkenToukaiOusatu"].ToString();
                        if (sDt != "")
                        {
                            base_tbl10_cmbTokaiOsatu.SelectedValue = sDt;
                        }
                    }
                    //再委託禁止条項の記載有無
                    obj = AnkenData_N.Rows[0]["NyuusatsuSaiitakuKinshiUmu"];
                    if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false && obj.ToString().Equals("0") == false)
                    {
                        base_tbl10_cmbKinsiUmu.SelectedValue = obj.ToString();
                    }
                    //再委託禁止条項の内容
                    obj = AnkenData_N.Rows[0]["NyuusatsuSaiitakuKinshiNaiyou"];
                    if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false && obj.ToString().Equals("0") == false)
                    {
                        base_tbl10_cmbKinsiNaiyo.SelectedValue = obj.ToString();
                    }
                    //その他の内容
                    base_tbl10_txtOtherNaiyo.Text = AnkenData_N.Rows[0]["NyuusatsuSaiitakuSonotaNaiyou"].ToString();
                    if (copy != COPY.GM)
                    {
                        //入札状況
                        obj = AnkenData_N.Rows[0]["NyuusatsuRakusatsushaID"];
                        if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
                        {
                            base_tbl10_cmbNyusatuStats.SelectedValue = obj.ToString();
                        }
                        //予定価格(税抜) ▼▼▼
                        base_tbl10_txtYoteiAmt.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_N.Rows[0]["NyuusatsuYoteiKakaku"]));
                        //応札数▼▼▼
                        base_tbl10_txtOsatuNum.Text = Convert.ToInt32(AnkenData_N.Rows[0]["NyuusatsushaSuu"]).ToString();
                        //落札者状況 ▼▼▼
                        string sObj = AnkenData_N.Rows[0]["NyuusatsuRakusatsuShaJokyou"].ToString();
                        if (sObj != "")
                        {
                            base_tbl10_cmbRakusatuStats.SelectedValue = sObj;
                        }
                        //落札者 ▼▼▼
                        base_tbl10_txtRakusatuSya.Text = AnkenData_N.Rows[0]["NyuusatsuRakusatsusha"].ToString();
                        //落札額状況 ▼▼▼
                        sObj = AnkenData_N.Rows[0]["NyuusatsuRakusatsuGakuJokyou"].ToString();
                        if (sObj != "")
                        {
                            base_tbl10_cmbRakusatuAmtStats.SelectedValue = sObj;
                        }
                        // No1594 「この案件番号の枝番で受託番号を作成する」際に落札額がコピー不要
                        if (copy != COPY.ED)
                        {
                            //落札額(税抜) ▼▼▼
                            base_tbl10_txtRakusatuAmt.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_N.Rows[0]["NyuusatsuRakusatugaku"]));
                        }
                    }
                }
            }
            // １１．契約情報
            if (mode != MODE.INSERT && mode != MODE.PLAN)
            {
                if (AnkenData_K != null || AnkenData_K.Rows.Count > 0)
                {
                    //契約締結(変更)日 KeiyakuKeiyakuTeiketsubiD
                    obj = AnkenData_K.Rows[0]["KeiyakuKeiyakuTeiketsubiD"];
                    if (obj != null && obj.ToString() != "")
                    {
                        base_tbl11_1_dtpKeiyakuChangeDt.Text = obj.ToString();
                    }
                    else
                    {
                        base_tbl11_1_dtpKeiyakuChangeDt.Text = "";
                        base_tbl11_1_dtpKeiyakuChangeDt.CustomFormat = " ";
                    }
                    //起案済
                    base_tbl11_1_chkKianzumi.Checked = (AnkenData_K.Rows[0]["AnkenKianZumi"].ToString() == "1"); ;
                    //起案日
                    obj = AnkenData_K.Rows[0]["KeiyakuSakuseibiD"];
                    if (obj != null && obj.ToString() != "")
                    {
                        base_tbl11_1_dtpKianDt.Text = obj.ToString();
                    }
                    else
                    {
                        base_tbl11_1_dtpKianDt.Text = "";
                        base_tbl11_1_dtpKianDt.CustomFormat = " ";
                    }
                    
                    //単契等の見込補正額(年度内)
                    base_tbl11_2_numAmt1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Mikomi1"]));
                    base_tbl11_2_numAmt2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Mikomi2"]));
                    base_tbl11_2_numAmt3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Mikomi3"]));
                    base_tbl11_2_numAmt4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Mikomi4"]));
                    GetTotalMoney("base_tbl11_2_numAmt", 5);
                    //年度繰越額(年度跨ぎ)
                    base_tbl11_3_numAmt1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Kurikoshi1"]));
                    base_tbl11_3_numAmt2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Kurikoshi2"]));
                    base_tbl11_3_numAmt3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Kurikoshi3"]));
                    base_tbl11_3_numAmt4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Kurikoshi4"]));
                    GetTotalMoney("base_tbl11_3_numAmt", 5);

                    //基本情報一覧での表示は、契約があれば、契約を表示。ない場合は、入札を表示する。
                    //再委託禁止条項の記載有無
                    obj = AnkenData_K.Rows[0]["KeiyakuSaiitakuKinshiUmu"];
                    if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false && obj.ToString().Equals("0") == false)
                    {
                        base_tbl10_cmbKinsiUmu.SelectedValue = obj.ToString();

                        //再委託禁止条項の内容
                        obj = AnkenData_K.Rows[0]["KeiyakuSaiitakuKinshiNaiyou"];
                        if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
                        {
                            base_tbl10_cmbKinsiNaiyo.SelectedValue = obj.ToString();
                        }
                        //その他の内容
                        base_tbl10_txtOtherNaiyo.Text = AnkenData_K.Rows[0]["KeiyakuSaiitakuSonotaNaiyou"].ToString();
                    }
                }
            }
            saishinFLG = Convert.ToInt32(AnkenData_H.Rows[0]["AnkenSaishinFlg"]);
            FolderPathCheck();

        }

        /// <summary>
        /// 事前打診タブ へデータ設定
        /// </summary>
        private void set_data_prior()
        {
            if (AnkenData_H == null || AnkenData_H.Rows.Count <= 0) return;
            DataRow drH = AnkenData_H.Rows[0];

            //１．事前打診状況
            //事前打診依頼日
            object obj = drH["AnkenJizenDashinIraibi"];
            if (obj != null && obj.ToString() != "")
            {
                prior_tbl01_dtpDasinIraiDt.Text = obj.ToString();
            }
            else
            {
                prior_tbl01_dtpDasinIraiDt.Text = "";
                prior_tbl01_dtpDasinIraiDt.CustomFormat = " ";
            }
            //参考見積対応
            obj = drH["AnkenToukaiSankouMitsumori"];
            if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
            {
                prior_tbl01_cmbMitumori.SelectedValue = obj.ToString();
            }
            //参考見積額(税抜)
            prior_tbl01_txtMitumoriAmt.Text = string.Format("{0:C}", Convert.ToInt64(drH["NyuusatsuMitsumoriAmt"]));

            //受注意欲
            obj = drH["AnkenToukaiJyutyuIyoku"];
            if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
            {
                prior_tbl01_cmbOrderIyoku.SelectedValue = obj.ToString();
            }
            //発注予定(見込)日
            obj = drH["AnkenHachuuYoteiMikomibi"];
            if (obj != null && obj.ToString() != "")
            {
                prior_tbl01_dtpOrderYoteiDt.Text = obj.ToString();
            }
            else
            {
                prior_tbl01_dtpOrderYoteiDt.Text = "";
                prior_tbl01_dtpOrderYoteiDt.CustomFormat = " ";
            }
            //案件メモ(事前打診)
            prior_tbl01_txtAnkenMemo.Text = drH["AnkenAnkenMemoJizendashin"].ToString();

            //２．未発注 ※事前打診はあったが、その後入札公告・見積合せ等の案内がない場合
            //未発注の登録日
            obj = drH["AnkenMihachuuTourokubi"];
            if (obj != null && obj.ToString() != "")
            {
                prior_tbl02_dtpNotOrderDt.Text = obj.ToString();
            }
            else
            {
                prior_tbl02_dtpNotOrderDt.Text = "";
                prior_tbl02_dtpNotOrderDt.CustomFormat = " ";
            }
            //未発注状況AnkenMihachuuJoukyou
            obj = drH["AnkenMihachuuJoukyou"];
            if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
            {
                prior_tbl02_cmbNotOrderStats.SelectedValue = obj.ToString();
            }
            //「発注なし」の理由AnkenHachuunashiRiyuu
            obj = drH["AnkenHachuunashiRiyuu"];
            if (obj != null && string.IsNullOrEmpty(obj.ToString()) == false)
            {
                prior_tbl02_cmbNotOrderReason.SelectedValue = obj.ToString();
            }
            //「その他」の内容AnkenSonotaNaiyou
            prior_tbl02_txtOtherNaiyo.Text = drH["AnkenSonotaNaiyou"].ToString();
            //案件メモ(未発注)AnkenAnkenMemoMihachuu
            prior_tbl02_txtAnkenMemo.Text = drH["AnkenAnkenMemoMihachuu"].ToString();
        }

        /// <summary>
        /// 入札タブ へデータ設定
        /// </summary>
        private void set_data_bid()
        {
            //１．入札情報
            // 入札情報登録日
            string sDt = AnkenData_N.Rows[0]["NyuusatsuJouhouTourokubi"].ToString();
            if (sDt != "")
            {
                bid_tbl01_dtpBidInfoDt.Text = sDt;
            }
            else
            {
                bid_tbl01_dtpBidInfoDt.Text = "";
                bid_tbl01_dtpBidInfoDt.CustomFormat = " ";
            }
            //業務発注区分
            object obj = AnkenData_N.Rows[0]["NyuusatsuGyoumuHachuukubun"];
            if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
            {
                bid_tbl01_cmbOrderKubun.SelectedValue = obj.ToString();
            }
            //入札方式
            obj = AnkenData_N.Rows[0]["NyuusatsuHoushiki"];
            if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
            {
                bid_tbl01_cmbBidhosiki.SelectedValue = obj.ToString();
            }
            //最低制限価格有無
            obj = AnkenData_N.Rows[0]["NyuusatsuSaiteiKakakuUmu"];
            if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
            {
                bid_tbl01_cmbLowestUmu.SelectedValue = obj.ToString();
            }
            //入札(予定)日
            sDt = AnkenData_N.Rows[0]["yoteibi"].ToString();
            if (sDt != "")
            {
                bid_tbl01_dtpBidYoteiDt.Text = sDt;
            }
            else
            {
                bid_tbl01_dtpBidYoteiDt.Text = "";
                bid_tbl01_dtpBidYoteiDt.CustomFormat = " ";
            }
            //参考見積対応
            obj = AnkenData_N.Rows[0]["NyuusatsuSankoumitsumoriTaiou"];
            if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
            {
                bid_tbl01_cmbMitumori.SelectedValue = obj.ToString();
            }
            //参考見積額(税抜)
            bid_tbl01_txtMitumoriAmt.Text = string.Format("{0:C}", GetLong(AnkenData_N.Rows[0]["NyuusatsuSankoumitsumoriKingaku"].ToString()));
            //受注意欲
            obj = AnkenData_N.Rows[0]["NyuusatsuJuchuuIyoku"];
            if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
            {
                bid_tbl01_cmbOrderIyoku.SelectedValue = obj.ToString();
            }
            //当会応札 ▼▼▼
            sDt = AnkenData_N.Rows[0]["AnkenToukaiOusatu"].ToString();
            if (sDt != "")
            {
                bid_tbl01_cmbTokaiOsatu.SelectedValue = sDt;
            }
            //２．再委託禁止条項
            //再委託禁止条項の記載有無
            obj = AnkenData_N.Rows[0]["NyuusatsuSaiitakuKinshiUmu"];
            if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
            {
                bid_tbl02_cmbKinsiUmu.SelectedValue = obj.ToString();
            }
            //再委託禁止条項の内容
            obj = AnkenData_N.Rows[0]["NyuusatsuSaiitakuKinshiNaiyou"];
            if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
            {
                bid_tbl02_cmbKinsiNaiyo.SelectedValue = obj.ToString();
            }
            //その他の内容
            bid_tbl02_txtOtherNaiyo.Text = AnkenData_N.Rows[0]["NyuusatsuSaiitakuSonotaNaiyou"].ToString();
            //３．入札結果
            //入札結果登録日
            sDt = AnkenData_N.Rows[0]["kekkaDate"].ToString();
            if (sDt != "")
            {
                bid_tbl03_1_dtpBidResultDt.Text = sDt;
            }
            else
            {
                bid_tbl03_1_dtpBidResultDt.Text = "";
                bid_tbl03_1_dtpBidResultDt.CustomFormat = " ";
            }
            //入札状況
            obj = AnkenData_N.Rows[0]["NyuusatsuRakusatsushaID"];
            if (!(obj == null || string.IsNullOrEmpty(obj.ToString())))
            {
                bid_tbl03_1_cmbBidStatus.SelectedValue = obj.ToString();
            }

            //予定価格(税抜) ▼▼▼
            bid_tbl03_1_txtYoteiPrice.Text = string.Format("{0:C}", GetLong(AnkenData_N.Rows[0]["NyuusatsuYoteiKakaku"].ToString()));
            //応札数▼▼▼
            bid_tbl03_1_txtOsatuNum.Text = GetInt(AnkenData_N.Rows[0]["NyuusatsushaSuu"].ToString()).ToString();
            //落札者状況 ▼▼▼
            sDt = AnkenData_N.Rows[0]["NyuusatsuRakusatsuShaJokyou"].ToString();
            if (sDt != "")
            {
                bid_tbl03_1_cmbRakusatuStatus.SelectedValue = sDt;
            }
            //落札者 ▼▼▼
            bid_tbl03_1_txtRakusatuSya.Text = AnkenData_N.Rows[0]["NyuusatsuRakusatsusha"].ToString();
            //落札額状況 ▼▼▼
            sDt = AnkenData_N.Rows[0]["NyuusatsuRakusatsuGakuJokyou"].ToString();
            if (sDt != "")
            {
                bid_tbl03_1_cmbRakusatuAmtStatus.SelectedValue = sDt;
            }
            //落札額(税抜) ▼▼▼
            bid_tbl03_1_numRakusatuAmt.Text = string.Format("{0:C}", GetLong(AnkenData_N.Rows[0]["NyuusatsuRakusatugaku"].ToString()));
            //案件メモ(入札)
            bid_tbl03_1_txtBidMemo.Text = AnkenData_N.Rows[0]["NyuusatsuAnkenMemoNuusatsu"].ToString();

            //応札者 登録日
            sDt = AnkenData_N.Rows[0]["syokaiDate"].ToString();
            if (sDt != "")
            {
                bid_tbl03_4_dtpInsDate.Text = sDt;
                bid_tbl03_4_dtpInsDate.CustomFormat = "";
            }
            else
            {
                bid_tbl03_4_dtpInsDate.CustomFormat = " ";
            }
            //応札者 更新日
            sDt = AnkenData_N.Rows[0]["saisyuDate"].ToString();
            if (sDt != "")
            {
                bid_tbl03_4_dtpInsDate.Text = sDt;
                bid_tbl03_4_dtpInsDate.CustomFormat = "";
            }
            else
            {
                bid_tbl03_4_dtpInsDate.CustomFormat = " ";
            }
            // 入札参加者リスト
            for (int k = 0; k < AnkenData_Grid2.Rows.Count; k++)
            {
                if (k > 0)
                {
                    bid_tbl03_4_c1FlexGrid.Rows.Add();
                }
                //【入札参加者】
                for (int i = 0; i < AnkenData_Grid2.Columns.Count - 1; i++)
                {
                    bid_tbl03_4_c1FlexGrid.Rows[k + 1][i + 2] = AnkenData_Grid2.Rows[k][i].ToString();
                }
            }
        }

        /// <summary>
        /// 契約、技術者評価タブ へデータ設定
        /// </summary>
        private void set_data_ca()
        {
            //２．配分情報・業務内容
            //部門配分 【事前打診・入札】		配分率(%)
            if (AnkenData_H != null && AnkenData_H.Rows.Count > 0)
            {
                ca_tbl02_1_numPercent1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["BuRitsu1"]));
                ca_tbl02_1_numPercent2.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["BuRitsu2"]));
                ca_tbl02_1_numPercent3.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["BuRitsu3"]));
                ca_tbl02_1_numPercent4.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["BuRitsu4"]));
                GetTotalPercent("ca_tbl02_1_numPercent", 5);

                //調査部 業務別配分 【事前打診・入札】		配分率(%)
                ca_tbl02_2_numPercent1.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu1"]));
                ca_tbl02_2_numPercent2.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu2"]));
                ca_tbl02_2_numPercent3.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu3"]));
                ca_tbl02_2_numPercent4.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu4"]));
                ca_tbl02_2_numPercent5.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu5"]));
                ca_tbl02_2_numPercent6.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu6"]));
                ca_tbl02_2_numPercent7.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu7"]));
                ca_tbl02_2_numPercent8.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu8"]));
                ca_tbl02_2_numPercent9.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu9"]));
                ca_tbl02_2_numPercent10.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu10"]));
                ca_tbl02_2_numPercent11.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu11"]));
                ca_tbl02_2_numPercent12.Text = GetPercentText(Convert.ToDouble(AnkenData_H.Rows[0]["ChousaRitsu12"]));
                GetTotalPercent("ca_tbl02_2_numPercent", 13);
            }

            if (AnkenData_K == null || AnkenData_K.Rows.Count <= 0) return;
            //１．契約情報
            #region 
            //契約締結(変更)日　
            object obj = AnkenData_K.Rows[0]["KeiyakuKeiyakuTeiketsubiD"];
            if (obj != null && obj.ToString() != "" && mode != MODE.CHANGE)
            {
                ca_tbl01_dtpChangeDt.Text = obj.ToString();
            }
            else
            {
                ca_tbl01_dtpChangeDt.CustomFormat = " ";
            }
            //案件区分
            sAnkenSakuseiKubun_ori = AnkenData_K.Rows[0]["AnkenSakuseiKubun"].ToString();
            ca_tbl01_cmbAnkenKubun.SelectedValue = sAnkenSakuseiKubun_ori;
            //起案済
            ca_tbl01_chkKian.Checked = (AnkenData_K.Rows[0]["AnkenKianZumi"].ToString() == "1");
            //起案日
            obj = AnkenData_K.Rows[0]["KeiyakuSakuseibiD"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl01_dtpKianDt.Text = obj.ToString();
            }
            else
            {
                ca_tbl01_dtpKianDt.CustomFormat = " ";
            }
            //発注者名・課名
            ca_tbl01_txtOrderKamei.Text = AnkenData_K.Rows[0]["AnkenHachuushaKaMei"].ToString();
            //業務名称(契約)
            ca_tbl01_txtAnkenName.Text = AnkenData_K.Rows[0]["AnkenGyoumuMei"].ToString();
            //契約区分
            ca_tbl01_cmbCaKubun.SelectedValue = AnkenData_K.Rows[0]["KeiyakuGyoumuKubun"].ToString();
            //工期自
            obj = AnkenData_K.Rows[0]["KeiyakuKoukiKaishibi"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl01_dtpKokiFrom.Text = obj.ToString();
                ca_tbl01_cmbStartYear.SelectedValue = AnkenData_K.Rows[0]["AnkenKoukiNendo"].ToString();//工期開始年度
            }
            else
            {
                ca_tbl01_dtpKokiFrom.CustomFormat = " ";
            }
            //工期至
            obj = AnkenData_K.Rows[0]["KeiyakuKoukiKanryoubi"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl01_dtpKokiTo.Text = obj.ToString();
                ca_tbl01_cmbSalesYear.SelectedValue = AnkenData_K.Rows[0]["AnkenUriageNendo"].ToString();//売上年度
            }
            else
            {
                ca_tbl01_dtpKokiTo.CustomFormat = " ";
            }

            // サ社経由
            if (AnkenData_K.Rows[0]["KeiyakuSashaKeiyu"].ToString() == "1")
            {
                ca_tbl01_chkSasya.Checked = true; 
            }
            //Ribc有
            if (AnkenData_K.Rows[0]["KeiyakuRIBCYouTankaData"].ToString() == "1")
            {
                ca_tbl01_chkRibcAri.Checked = true;
            }
            //契約金額 税抜(自動計算用)
            ca_tbl01_txtZeinukiAmt.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakuKeiyakuAmt"]));
            //契約金額 税込
            ca_tbl01_txtZeikomiAmt.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakuZeikomiAmt"]));
            //契約金額 内消費税
            ca_tbl01_txtSyohizeiAmt.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakuuchizeiAmt"]));
            //消費税率
            ca_tbl01_txtTax.Text = AnkenData_K.Rows[0]["KeiyakuShouhizeiritsu"].ToString();
            //受託金額(税込)
            ca_tbl01_txtJyutakuAmt.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakukeiyakuAmtkukei"]));
            //受託外金額(税込)
            ca_tbl01_txtJyutakuGaiAmt.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakuBetsuKeiyakuAmt"]));
            //基本情報等一覧へ連動：契約金額(税抜)※受託外をのぞく
            base_tbl11_1_txtKeiyakuAmt.Text = string.Format("{0:C}", Get_Zeinuki(Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakukeiyakuAmtkukei"])));
            //変更・中止理由
            ca_tbl01_txtRiyu.Text = AnkenData_K.Rows[0]["KeiyakuHenkouChuushiRiyuu"].ToString();
            //案件メモ(契約)
            ca_tbl01_txtAnkenMemo.Text = AnkenData_K.Rows[0]["KeiyakuAnkenMemoKeiyaku"].ToString();
            //備考
            ca_tbl01_txtBiko.Text = AnkenData_K.Rows[0]["KeiyakuBikou"].ToString();
            //別途資料
            if (AnkenData_K.Rows[0]["KeiyakuShosha"].ToString() == "1")
            {
                ca_tbl01_chkCaSyosya.Checked = true; //契約書写
            }
            if (AnkenData_K.Rows[0]["KeiyakuTokkiShiyousho"].ToString() == "1")
            {
                ca_tbl01_chkSiyosyo.Checked = true;//特記仕様書
            }
            if (AnkenData_K.Rows[0]["KeiyakuMitsumorisho"].ToString() == "1")
            {
                ca_tbl01_chkMitumorisyo.Checked = true;//見積書
            }
            if (AnkenData_K.Rows[0]["KeiyakuTanpinChousaMitsumorisho"].ToString() == "1")
            {
                ca_tbl01_chkTanpinTyosa.Checked = true;//単品調査内訳書
            }
            if (AnkenData_K.Rows[0]["KeiyakuSonota"].ToString() == "1")
            {
                ca_tbl01_chkOther.Checked = true;//その他
            }
            if (AnkenData_K.Rows[0]["KeiyakuRIBCYouTankaDataMoushikomisho"].ToString() == "1")
            {
                ca_tbl01_chkRibcSyo.Checked = true;//RIBC用単価データ申込書
            }
            //その他備考
            ca_tbl01_txtOtherBiko.Text = AnkenData_K.Rows[0]["KeiyakuSonotaNaiyou"].ToString();
            //契約図書
            ca_tbl01_txtTosyo.Text = AnkenData_K.Rows[0]["AnkenKeiyakusho"].ToString();

            obj = AnkenData_K.Rows[0]["KeiyakuSaiitakuKinshiUmu"];
            if(obj == null || string.IsNullOrEmpty(obj.ToString()) || obj.ToString().Equals("0"))
            {
                if (AnkenData_N != null || AnkenData_N.Rows.Count >0)
                {
                    obj = AnkenData_N.Rows[0]["NyuusatsuSaiitakuKinshiUmu"];
                    if(obj != null && obj.ToString().Equals("1"))
                    {
                        // 入札段階で再委託条項がありの場合、表示し入力出来るようにする。無し、不明は空欄で入力できるようにする。空欄の場合はエラーとする。
                        //再委託禁止条項の記載有無
                        ca_tbl01_cmbKinsiUmu.SelectedValue = obj.ToString();
                        //再委託禁止条項の内容
                        obj = AnkenData_N.Rows[0]["NyuusatsuSaiitakuKinshiNaiyou"];
                        ca_tbl01_cmbKinsiNaiyo.SelectedValue = obj.ToString();
                        //その他の内容 
                        ca_tbl01_txtOtherNaiyo.Text = AnkenData_N.Rows[0]["NyuusatsuSaiitakuSonotaNaiyou"].ToString();
                    }
                }
            }
            else
            {
                //再委託禁止条項の記載有無
                ca_tbl01_cmbKinsiUmu.SelectedValue = obj.ToString();
                //再委託禁止条項の内容
                obj = AnkenData_K.Rows[0]["KeiyakuSaiitakuKinshiNaiyou"];
                ca_tbl01_cmbKinsiNaiyo.SelectedValue = obj.ToString();
                //その他の内容 
                ca_tbl01_txtOtherNaiyo.Text = AnkenData_K.Rows[0]["KeiyakuSaiitakuSonotaNaiyou"].ToString();
            }
            #endregion

            //２．配分情報・業務内容
            //部門配分 【事前打診・入札】		配分率(%)
            // ↑で設定する           
            //部門配分 【契約後】		配分率(%)
            ca_tbl02_AftCaBm_numPercent1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["Haibun1"]));
            ca_tbl02_AftCaBm_numPercent2.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["Haibun2"]));
            ca_tbl02_AftCaBm_numPercent3.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["Haibun3"]));
            ca_tbl02_AftCaBm_numPercent4.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["Haibun4"]));
            GetTotalPercent("ca_tbl02_AftCaBm_numPercent", 5);

            // 基本情報等一覧：７．配分情報・業務内容　部門配分　【契約後】配分率(%)
            base_tbl07_4_lblRate1.Text = ca_tbl02_AftCaBm_numPercent1.Text;
            base_tbl07_4_lblRate2.Text = ca_tbl02_AftCaBm_numPercent2.Text;
            base_tbl07_4_lblRate3.Text = ca_tbl02_AftCaBm_numPercent3.Text;
            base_tbl07_4_lblRate4.Text = ca_tbl02_AftCaBm_numPercent4.Text;
            base_tbl07_4_lblRateAll.Text = ca_tbl02_AftCaBm_numPercentAll.Text;

            //部門配分 【契約後】		配分額(税込)
            ca_tbl02_AftCaBmZeikomi_numAmt1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Uriage1"]));
            ca_tbl02_AftCaBmZeikomi_numAmt2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Uriage2"]));
            ca_tbl02_AftCaBmZeikomi_numAmt3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Uriage3"]));
            ca_tbl02_AftCaBmZeikomi_numAmt4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Uriage4"]));
            GetTotalMoney("ca_tbl02_AftCaBmZeikomi_numAmt", 5);
            // 基本情報等一覧：７．配分情報・業務内容　部門配分　【契約後】配分率(%)

            //部門配分 【契約後】		配分額(税抜)
            ca_tbl02_AftCaBm_numAmt1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakuHaibunZeinuki1"]));
            ca_tbl02_AftCaBm_numAmt2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakuHaibunZeinuki2"]));
            ca_tbl02_AftCaBm_numAmt3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakuHaibunZeinuki3"]));
            ca_tbl02_AftCaBm_numAmt4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["KeiyakuHaibunZeinuki4"]));
            GetTotalMoney("ca_tbl02_AftCaBm_numAmt", 5);

            // 基本情報等一覧：７．配分情報・業務内容　部門配分　【契約後】配分額(税抜)
            base_tbl07_4_lblAmt1.Text = ca_tbl02_AftCaBm_numAmt1.Text;
            base_tbl07_4_lblAmt2.Text = ca_tbl02_AftCaBm_numAmt2.Text;
            base_tbl07_4_lblAmt3.Text = ca_tbl02_AftCaBm_numAmt3.Text;
            base_tbl07_4_lblAmt4.Text = ca_tbl02_AftCaBm_numAmt4.Text;
            base_tbl07_4_lblAmtAll.Text = ca_tbl02_AftCaBm_numAmtAll.Text;


            //調査部 業務別配分 【事前打診・入札】		配分率(%)
            //↑で設定する

            //調査部 業務別配分 【契約後】		配分率(%)
            ca_tbl02_AftCaTs_numPercent1.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu1"]));
            ca_tbl02_AftCaTs_numPercent2.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu2"]));
            ca_tbl02_AftCaTs_numPercent3.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu3"]));
            ca_tbl02_AftCaTs_numPercent4.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu4"]));
            ca_tbl02_AftCaTs_numPercent5.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu5"]));
            ca_tbl02_AftCaTs_numPercent6.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu6"]));
            ca_tbl02_AftCaTs_numPercent7.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu7"]));
            ca_tbl02_AftCaTs_numPercent8.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu8"]));
            ca_tbl02_AftCaTs_numPercent9.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu9"]));
            ca_tbl02_AftCaTs_numPercent10.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu10"]));
            ca_tbl02_AftCaTs_numPercent11.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu11"]));
            ca_tbl02_AftCaTs_numPercent12.Text = GetPercentText(Convert.ToDouble(AnkenData_K.Rows[0]["GyoumuRitsu12"]));
            GetTotalPercent("ca_tbl02_AftCaTs_numPercent", 13);

            // 基本情報等一覧：７．配分情報・業務内容　調査部 業務別配分　【契約後】配分率(%)
            base_tbl07_5_lblRate1.Text = ca_tbl02_AftCaTs_numPercent1.Text;
            base_tbl07_5_lblRate2.Text = ca_tbl02_AftCaTs_numPercent2.Text;
            base_tbl07_5_lblRate3.Text = ca_tbl02_AftCaTs_numPercent3.Text;
            base_tbl07_5_lblRate4.Text = ca_tbl02_AftCaTs_numPercent4.Text;
            base_tbl07_5_lblRate5.Text = ca_tbl02_AftCaTs_numPercent5.Text;
            base_tbl07_5_lblRate6.Text = ca_tbl02_AftCaTs_numPercent6.Text;
            base_tbl07_5_lblRate7.Text = ca_tbl02_AftCaTs_numPercent7.Text;
            base_tbl07_5_lblRate8.Text = ca_tbl02_AftCaTs_numPercent8.Text;
            base_tbl07_5_lblRate9.Text = ca_tbl02_AftCaTs_numPercent9.Text;
            base_tbl07_5_lblRate10.Text = ca_tbl02_AftCaTs_numPercent10.Text;
            base_tbl07_5_lblRate11.Text = ca_tbl02_AftCaTs_numPercent11.Text;
            base_tbl07_5_lblRate12.Text = ca_tbl02_AftCaTs_numPercent12.Text;
            base_tbl07_5_lblRateAll.Text = ca_tbl02_AftCaTs_numPercentAll.Text;

            //調査部 業務別配分 【契約後】		配分額(税抜)
            ca_tbl02_AftCaTs_numAmt1.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku1"]));
            ca_tbl02_AftCaTs_numAmt2.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku2"]));
            ca_tbl02_AftCaTs_numAmt3.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku3"]));
            ca_tbl02_AftCaTs_numAmt4.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku4"]));
            ca_tbl02_AftCaTs_numAmt5.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku5"]));
            ca_tbl02_AftCaTs_numAmt6.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku6"]));
            ca_tbl02_AftCaTs_numAmt7.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku7"]));
            ca_tbl02_AftCaTs_numAmt8.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku8"]));
            ca_tbl02_AftCaTs_numAmt9.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku9"]));
            ca_tbl02_AftCaTs_numAmt10.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku10"]));
            ca_tbl02_AftCaTs_numAmt11.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku11"]));
            ca_tbl02_AftCaTs_numAmt12.Text = GetMoneyTextLong(Convert.ToInt64(AnkenData_K.Rows[0]["GyoumuGaku12"]));
            GetTotalMoney("ca_tbl02_AftCaTs_numAmt", 13);

            // 基本情報等一覧：７．配分情報・業務内容　調査部 業務別配分　【契約後】配分額(税抜)
            base_tbl07_5_lblAmt1.Text = ca_tbl02_AftCaTs_numAmt1.Text;
            base_tbl07_5_lblAmt2.Text = ca_tbl02_AftCaTs_numAmt2.Text;
            base_tbl07_5_lblAmt3.Text = ca_tbl02_AftCaTs_numAmt3.Text;
            base_tbl07_5_lblAmt4.Text = ca_tbl02_AftCaTs_numAmt4.Text;
            base_tbl07_5_lblAmt5.Text = ca_tbl02_AftCaTs_numAmt5.Text;
            base_tbl07_5_lblAmt6.Text = ca_tbl02_AftCaTs_numAmt6.Text;
            base_tbl07_5_lblAmt7.Text = ca_tbl02_AftCaTs_numAmt7.Text;
            base_tbl07_5_lblAmt8.Text = ca_tbl02_AftCaTs_numAmt8.Text;
            base_tbl07_5_lblAmt9.Text = ca_tbl02_AftCaTs_numAmt9.Text;
            base_tbl07_5_lblAmt10.Text = ca_tbl02_AftCaTs_numAmt10.Text;
            base_tbl07_5_lblAmt11.Text = ca_tbl02_AftCaTs_numAmt11.Text;
            base_tbl07_5_lblAmt12.Text = ca_tbl02_AftCaTs_numAmt12.Text;
            base_tbl07_5_lblAmtAll.Text = ca_tbl02_AftCaTs_numAmtAll.Text;

            //３．単契等の見込補正額(年度内)
            //調査部部門配分額(税抜)
            //事業普及部部門配分額(税抜)
            //情報ｼｽﾃﾑ部部門配分額(税抜)
            //総合研究部部門配分額(税抜)
            //単契等の見込補正額(税抜)
            ca_tbl03_numAmt1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Mikomi1"]));
            ca_tbl03_numAmt2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Mikomi2"]));
            ca_tbl03_numAmt3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Mikomi3"]));
            ca_tbl03_numAmt4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Mikomi4"]));
            GetTotalMoney("ca_tbl03_numAmt", 5);
            //４．年度繰越額(年度跨ぎ)
            //調査部部門配分額(税抜)
            //事業普及部部門配分額(税抜)
            //情報ｼｽﾃﾑ部部門配分額(税抜)
            //総合研究部部門配分額(税抜)
            //年度繰越額合計(税抜)
            ca_tbl04_numKurikosiAmt1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Kurikoshi1"]));
            ca_tbl04_numKurikosiAmt2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Kurikoshi2"]));
            ca_tbl04_numKurikosiAmt3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Kurikoshi3"]));
            ca_tbl04_numKurikosiAmt4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["Kurikoshi4"]));
            GetTotalMoney("ca_tbl04_numKurikosiAmt",5);
            #region ５．管理者・担当者
            //管理技術者
            ca_tbl05_txtKanriCD.Text = AnkenData_K.Rows[0]["KanriGijutsushaCD"].ToString();
            ca_tbl05_txtKanri.Text = AnkenData_K.Rows[0]["KanriGijutsushaNM"].ToString();
            //照査技術者
            ca_tbl05_txtSyosaCD.Text = AnkenData_K.Rows[0]["ShousaTantoushaCD"].ToString();
            ca_tbl05_txtSyosa.Text = AnkenData_K.Rows[0]["ShousaTantoushaNM"].ToString();
            //担当技術者 技術者評価：担当技術者も一緒に設定する
            for (int k = 0; k < AnkenData_Grid5.Rows.Count; k++)
            {
                if (k > 0)
                {
                    ca_tbl05_txtTanto_c1FlexGrid.Rows.Add();
                    te_c1FlexGrid.Rows.Add();
                }
                //【技術担当者】
                for (int i = 0; i < AnkenData_Grid5.Columns.Count; i++)
                {
                    ca_tbl05_txtTanto_c1FlexGrid.Rows[k + 1][i + 1] = AnkenData_Grid5.Rows[k][i].ToString();
                    te_c1FlexGrid.Rows[k + 1][i + 1] = AnkenData_Grid5.Rows[k][i].ToString();
                }
            }
            //審査担当者
            ca_tbl05_txtSinsaCD.Text = AnkenData_K.Rows[0]["SinsaTantoushaCD"].ToString();
            ca_tbl05_txtSinsa.Text = AnkenData_K.Rows[0]["SinsaTantoushaNM"].ToString();
            //業務管理者
            ca_tbl05_txtGyomuCD.Text = AnkenData_K.Rows[0]["GyoumuKanrishaCD"].ToString();
            ca_tbl05_txtGyomu.Text = AnkenData_K.Rows[0]["GyoumuKanrishaMei"].ToString();
            //窓口担当者
            ca_tbl05_txtMadoguchiCD.Text = AnkenData_K.Rows[0]["GyoumuJouhouMadoKojinCD"].ToString();
            ca_tbl05_txtMadoguchi.Text = AnkenData_K.Rows[0]["GyoumuJouhouMadoChousainMei"].ToString();
            ca_tbl05_txtMadoguchiBusho.Text = AnkenData_K.Rows[0]["GyoumuJouhouMadoGyoumuBushoCD"].ToString();
            ca_tbl05_txtMadoguchiShibu.Text = AnkenData_K.Rows[0]["GyoumuJouhouMadoShibuMei"].ToString();
            ca_tbl05_txtMadoguchiKa.Text = AnkenData_K.Rows[0]["GyoumuJouhouMadoKamei"].ToString();
            #endregion

            #region ６．売上計上情報
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
                        row += 1;
                    }
                    ca_tbl06_c1FlexGrid.Rows[rowCntT + 1][1] = AnkenData_Grid4.Rows[k][1].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntT + 1][2] = AnkenData_Grid4.Rows[k][2].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntT + 1][3] = AnkenData_Grid4.Rows[k][3].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntT + 1][4] = AnkenData_Grid4.Rows[k][5].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntT + 1][5] = AnkenData_Grid4.Rows[k][6].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntT + 1][6] = AnkenData_Grid4.Rows[k][7].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntT + 1][7] = AnkenData_Grid4.Rows[k][8].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntT + 1][8] = AnkenData_Grid4.Rows[k][9].ToString();
                }
                if ("B".Equals(AnkenData_Grid4.Rows[k][4].ToString()))
                {
                    rowCntB++;
                    if (row < rowCntB)
                    {
                        row += 1;
                    }
                    ca_tbl06_c1FlexGrid.Rows[rowCntB + 1][9] = AnkenData_Grid4.Rows[k][1].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntB + 1][10] = AnkenData_Grid4.Rows[k][2].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntB + 1][11] = AnkenData_Grid4.Rows[k][3].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntB + 1][12] = AnkenData_Grid4.Rows[k][5].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntB + 1][13] = AnkenData_Grid4.Rows[k][6].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntB + 1][14] = AnkenData_Grid4.Rows[k][7].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntB + 1][15] = AnkenData_Grid4.Rows[k][8].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntB + 1][16] = AnkenData_Grid4.Rows[k][9].ToString();
                }
                if ("J".Equals(AnkenData_Grid4.Rows[k][4].ToString()))
                {
                    rowCntJ++;
                    if (row < rowCntJ)
                    {
                        row += 1;
                    }
                    ca_tbl06_c1FlexGrid.Rows[rowCntJ + 1][17] = AnkenData_Grid4.Rows[k][1].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntJ + 1][18] = AnkenData_Grid4.Rows[k][2].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntJ + 1][19] = AnkenData_Grid4.Rows[k][3].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntJ + 1][20] = AnkenData_Grid4.Rows[k][5].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntJ + 1][21] = AnkenData_Grid4.Rows[k][6].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntJ + 1][22] = AnkenData_Grid4.Rows[k][7].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntJ + 1][23] = AnkenData_Grid4.Rows[k][8].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntJ + 1][24] = AnkenData_Grid4.Rows[k][9].ToString();
                }
                if ("K".Equals(AnkenData_Grid4.Rows[k][4].ToString()))
                {
                    rowCntK++;
                    if (row < rowCntK)
                    {
                        row += 1;
                    }
                    ca_tbl06_c1FlexGrid.Rows[rowCntK + 1][25] = AnkenData_Grid4.Rows[k][1].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntK + 1][26] = AnkenData_Grid4.Rows[k][2].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntK + 1][27] = AnkenData_Grid4.Rows[k][3].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntK + 1][28] = AnkenData_Grid4.Rows[k][5].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntK + 1][29] = AnkenData_Grid4.Rows[k][6].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntK + 1][30] = AnkenData_Grid4.Rows[k][7].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntK + 1][31] = AnkenData_Grid4.Rows[k][8].ToString();
                    ca_tbl06_c1FlexGrid.Rows[rowCntK + 1][32] = AnkenData_Grid4.Rows[k][9].ToString();
                }
            }
            #endregion

            //７．請求書情報
            // 請求日付１
            obj = AnkenData_K.Rows[0]["Seikyuubi1"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl07_dtpRequst1.Text = obj.ToString();
            }
            else
            {
                ca_tbl07_dtpRequst1.CustomFormat = " ";
            }
            // 請求日付2
            obj = AnkenData_K.Rows[0]["Seikyuubi2"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl07_dtpRequst2.Text = obj.ToString();
            }
            else
            {
                ca_tbl07_dtpRequst2.CustomFormat = " ";
            }
            // 請求日付3
            obj = AnkenData_K.Rows[0]["Seikyuubi3"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl07_dtpRequst3.Text = obj.ToString();
            }
            else
            {
                ca_tbl07_dtpRequst3.CustomFormat = " ";
            }
            // 請求日付4
            obj = AnkenData_K.Rows[0]["Seikyuubi4"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl07_dtpRequst4.Text = obj.ToString();
            }
            else
            {
                ca_tbl07_dtpRequst4.CustomFormat = " ";
            }
            // 請求日付5
            obj = AnkenData_K.Rows[0]["Seikyuubi5"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl07_dtpRequst5.Text = obj.ToString();
            }
            else
            {
                ca_tbl07_dtpRequst5.CustomFormat = " ";
            }
            // 前払金日付
            obj = AnkenData_K.Rows[0]["Seikyuubi6"];
            if (obj != null && obj.ToString() != "")
            {
                ca_tbl07_dtpRequst6.Text = obj.ToString();
            }
            else
            {
                ca_tbl07_dtpRequst6.CustomFormat = " ";
            }
            // 請求金額１
            ca_tbl07_txtRequst1.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["SeikyuuAmt1"]));
            // 請求金額２
            ca_tbl07_txtRequst2.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["SeikyuuAmt2"]));
            // 請求金額３
            ca_tbl07_txtRequst3.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["SeikyuuAmt3"]));
            // 請求金額４
            ca_tbl07_txtRequst4.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["SeikyuuAmt4"]));
            // 請求金額５
            ca_tbl07_txtRequst5.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["SeikyuuAmt5"]));
            // 前払金
            ca_tbl07_txtRequst6.Text = string.Format("{0:C}", Convert.ToInt64(AnkenData_K.Rows[0]["SeikyuuAmt6"]));
            GetTotalMoney("ca_tbl07_txtRequst", 7);

            // 技術者評価
            if (AnkenData_G != null && AnkenData_G.Rows.Count > 0)
            {
                te_txtPoint.Text = AnkenData_G.Rows[0]["GyoumuHyouten"].ToString();
                te_lblKanri.Text = AnkenData_G.Rows[0]["KanriGijutsushaNM"].ToString();
                te_txtKanriPoint.Text = AnkenData_G.Rows[0]["GyoumuKanriHyouten"].ToString();
                ca_tbl05_txtKanriHyoten.Text = AnkenData_G.Rows[0]["GyoumuKanriHyouten"].ToString();
                te_lblSyosa.Text = AnkenData_G.Rows[0]["ShousaTantoushaNM"].ToString();
                te_txtSyosaPoint.Text = AnkenData_G.Rows[0]["GyoumuShousaHyouten"].ToString();
                ca_tbl05_txtSyosaHyoten.Text = AnkenData_G.Rows[0]["GyoumuShousaHyouten"].ToString();
                te_txtTecris.Text = AnkenData_G.Rows[0]["GyoumuTECRISTourokuBangou"].ToString();
                obj = AnkenData_G.Rows[0]["GyoumuSeikyuubi"];
                if (obj != null && obj.ToString() != "")
                {
                    te_dtpSeikyusyaDt.Text = obj.ToString();
                }
                else
                {
                    te_dtpSeikyusyaDt.CustomFormat = " ";
                }
                // 技術評価者タブの請求書は 02契約関係図書 を付ける
                te_txtSeikyusyo.Text = AnkenData_G.Rows[0]["AnkenKeiyakusho"].ToString() + @"\02契約関係図書";
                te_txtCustomComment.Text = AnkenData_G.Rows[0]["AnkenKokyakuHyoukaComment"].ToString();
                te_txtTokaiComment.Text = AnkenData_G.Rows[0]["AnkenToukaiHyoukaComment"].ToString();
            }
            // フォルダパス振り直し
            set_folder();
        }
        #endregion


        /// <summary>
        /// 各タブのコンボボックスリストの設定処理
        /// </summary>
        private void set_combo()
        {
            GlobalMethod GlobalMethod = new GlobalMethod();
            string discript = "";
            string value = "";
            string table = "";
            string where = "";

            // -- 受託課所支部　→年度の変更よりリスト設定するので、ここは設定しない

            #region 案件区分 -----------------------------------
            // -- 基本情報等一覧：２．基本情報
            discript = "SakuseiKubun";
            value = "SakuseiKubunID";
            table = "Mst_SakuseiKubun";
            where = " SakuseiKubunID = '01' ";
            DataTable cmbAnk1 = GlobalMethod.getData(discript, value, table, where);
            this.base_tbl02_cmbAnkenKubun.DataSource = cmbAnk1;
            this.base_tbl02_cmbAnkenKubun.DisplayMember = "Discript";
            this.base_tbl02_cmbAnkenKubun.ValueMember = "Value";


            // -- 契約：１．契約情報
            if (mode == MODE.CHANGE)
            {
                where = " SakuseiKubunID >= '04' AND SakuseiKubunID <> '05' ";
            }
            else
            {
                // 案件情報IDが存在する（更新等）
                if (AnkenID != "" && mode != MODE.INSERT && mode != MODE.PLAN)
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
                        ca_tbl01_cmbAnkenKubun.Enabled = false;
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
            DataTable cmbAnk2 = GlobalMethod.getData(discript, value, table, where);
            this.ca_tbl01_cmbAnkenKubun.DataSource = cmbAnk2;
            this.ca_tbl01_cmbAnkenKubun.DisplayMember = "Discript";
            this.ca_tbl01_cmbAnkenKubun.ValueMember = "Value";

            #endregion

            #region 契約区分 -----------------------------------
            // -- 基本情報等一覧：３．案件情報
            discript = "GyoumuKubunHyouji";
            value = "GyoumuNarabijunCD";
            table = "Mst_GyoumuKubun";
            where = "GyoumuNarabijunCD < 100 ";
            //コンボボックスデータ取得
            DataTable cmbKydiv = GlobalMethod.getData(discript, value, table, where);
            if (cmbKydiv != null)
            {
                DataRow dr = cmbKydiv.NewRow();
                cmbKydiv.Rows.InsertAt(dr, 0);
            }
            this.base_tbl03_cmbKeiyakuKubun.DataSource = cmbKydiv;
            this.base_tbl03_cmbKeiyakuKubun.DisplayMember = "Discript";
            this.base_tbl03_cmbKeiyakuKubun.ValueMember = "Value";

            // -- 契約：１．契約情報
            this.ca_tbl01_cmbCaKubun.DataSource = cmbKydiv;
            this.ca_tbl01_cmbCaKubun.DisplayMember = "Discript";
            this.ca_tbl01_cmbCaKubun.ValueMember = "Value";

            #endregion

            #region 工期開始年度、売上年度 ---------------------
            discript = "NendoSeireki";
            value = "NendoID";
            table = "Mst_Nendo";
            where = "";
            //コンボボックスデータ取得
            DataTable cmbYear = GlobalMethod.getData(discript, value, table, where);
            // -- 基本情報等一覧：３．案件情報
            this.base_tbl03_cmbKokiStartYear.DataSource = cmbYear;  // -- 工期開始年度
            this.base_tbl03_cmbKokiStartYear.DisplayMember = "Discript";
            this.base_tbl03_cmbKokiStartYear.ValueMember = "Value";
            DataTable cmbSalesYear = cmbYear.AsEnumerable().CopyToDataTable();
            this.base_tbl03_cmbKokiSalesYear.DataSource = cmbSalesYear;  // -- 売上年度
            this.base_tbl03_cmbKokiSalesYear.DisplayMember = "Discript";
            this.base_tbl03_cmbKokiSalesYear.ValueMember = "Value";

            // -- 契約：１．契約情報
            this.ca_tbl01_cmbStartYear.DataSource = cmbYear;  // -- 工期開始年度
            this.ca_tbl01_cmbStartYear.DisplayMember = "Discript";
            this.ca_tbl01_cmbStartYear.ValueMember = "Value";

            this.ca_tbl01_cmbSalesYear.DataSource = cmbSalesYear;  // -- 売上年度
            this.ca_tbl01_cmbSalesYear.DisplayMember = "Discript";
            this.ca_tbl01_cmbSalesYear.ValueMember = "Value";

            #endregion

            // No.1533 削除
            #region 応援依頼の有無 -----------------------------
            //DataTable cmbOeUm = new System.Data.DataTable();
            //cmbOeUm.Columns.Add("Value", typeof(int));
            //cmbOeUm.Columns.Add("Discript", typeof(string));
            //cmbOeUm.Rows.Add(0, "");
            //cmbOeUm.Rows.Add(1, "あり");
            //cmbOeUm.Rows.Add(2, "なし");
            //// -- 基本情報等一覧：７．業務配分
            //this.base_tbl07_3_cmbOen.DataSource = cmbOeUm;
            //this.base_tbl07_3_cmbOen.DisplayMember = "Discript";
            //this.base_tbl07_3_cmbOen.ValueMember = "Value";
            #endregion

            #region 「発注無し」の理由 -------------------------
            // -- 基本情報等一覧：９．事前打診・参考見積
            DataTable cmbNoOrderReason = new System.Data.DataTable();
            cmbNoOrderReason.Columns.Add("Value", typeof(int));
            cmbNoOrderReason.Columns.Add("Discript", typeof(string));
            cmbNoOrderReason.Rows.Add(0, "");
            cmbNoOrderReason.Rows.Add(1, "工事計画の中止・延期");
            cmbNoOrderReason.Rows.Add(2, "全社辞退");
            cmbNoOrderReason.Rows.Add(3, "調査費用の不足");
            cmbNoOrderReason.Rows.Add(4, "その他");
            this.base_tbl09_cmbNotOrderReason.DataSource = cmbNoOrderReason;
            this.base_tbl09_cmbNotOrderReason.DisplayMember = "Discript";
            this.base_tbl09_cmbNotOrderReason.ValueMember = "Value";
            // -- 事前打診：２．未発注
            this.prior_tbl02_cmbNotOrderReason.DataSource = cmbNoOrderReason;
            this.prior_tbl02_cmbNotOrderReason.DisplayMember = "Discript";
            this.prior_tbl02_cmbNotOrderReason.ValueMember = "Value";
            #endregion

            #region 未発注状況 ---------------------------------
            DataTable cmbNoOrderSt = new System.Data.DataTable();
            cmbNoOrderSt.Columns.Add("Value", typeof(int));
            cmbNoOrderSt.Columns.Add("Discript", typeof(string));
            cmbNoOrderSt.Rows.Add(0, "");
            cmbNoOrderSt.Rows.Add(1, "発注無し");
            cmbNoOrderSt.Rows.Add(2, "不明");
            // -- 基本情報等一覧：９．事前打診・参考見積
            this.base_tbl09_cmbNotOrderStats.DataSource = cmbNoOrderSt;
            this.base_tbl09_cmbNotOrderStats.DisplayMember = "Discript";
            this.base_tbl09_cmbNotOrderStats.ValueMember = "Value";

            // -- 事前打診：２．未発注
            this.prior_tbl02_cmbNotOrderStats.DataSource = cmbNoOrderSt;
            this.prior_tbl02_cmbNotOrderStats.DisplayMember = "Discript";
            this.prior_tbl02_cmbNotOrderStats.ValueMember = "Value";
            #endregion

            #region 受注意欲  ----------------------------------
            DataTable cmbOrderIyk = new System.Data.DataTable();
            cmbOrderIyk.Columns.Add("Value", typeof(int));
            cmbOrderIyk.Columns.Add("Discript", typeof(string));
            cmbOrderIyk.Rows.Add(0, "");
            cmbOrderIyk.Rows.Add(1, "フラット");
            cmbOrderIyk.Rows.Add(2, "あり");
            cmbOrderIyk.Rows.Add(3, "なし");
            DataTable cmbPOrderIyk = cmbOrderIyk.AsEnumerable().CopyToDataTable();
            // -- 基本情報等一覧：９．事前打診・参考見積
            this.base_tbl09_cmbOrderIyoku.DataSource = cmbPOrderIyk;
            this.base_tbl09_cmbOrderIyoku.DisplayMember = "Discript";
            this.base_tbl09_cmbOrderIyoku.ValueMember = "Value";

            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbOrderIyoku.DataSource = cmbOrderIyk;
            this.base_tbl10_cmbOrderIyoku.DisplayMember = "Discript";
            this.base_tbl10_cmbOrderIyoku.ValueMember = "Value";

            // -- 事前打診：１．事前打診状況
            this.prior_tbl01_cmbOrderIyoku.DataSource = cmbPOrderIyk;
            this.prior_tbl01_cmbOrderIyoku.DisplayMember = "Discript";
            this.prior_tbl01_cmbOrderIyoku.ValueMember = "Value";

            // -- 入札：１．入札情報
            this.bid_tbl01_cmbOrderIyoku.DataSource = cmbOrderIyk;
            this.bid_tbl01_cmbOrderIyoku.DisplayMember = "Discript";
            this.bid_tbl01_cmbOrderIyoku.ValueMember = "Value";
            #endregion

            #region 参考見積対応 -------------------------------
            DataTable cmbSkMtmr = new System.Data.DataTable();
            cmbSkMtmr.Columns.Add("Value", typeof(int));
            cmbSkMtmr.Columns.Add("Discript", typeof(string));
            cmbSkMtmr.Rows.Add(0, "");
            cmbSkMtmr.Rows.Add(1, "検討中");
            cmbSkMtmr.Rows.Add(2, "提出");
            cmbSkMtmr.Rows.Add(4, "辞退");
            cmbSkMtmr.Rows.Add(3, "依頼無し");
            DataTable cmbPSkMtmr = cmbSkMtmr.AsEnumerable().CopyToDataTable();
            // -- 基本情報等一覧：９．事前打診・参考見積
            this.base_tbl09_cmbSankomitumori.DataSource = cmbPSkMtmr;
            this.base_tbl09_cmbSankomitumori.DisplayMember = "Discript";
            this.base_tbl09_cmbSankomitumori.ValueMember = "Value";

            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbSankoMitumori.DataSource = cmbSkMtmr;
            this.base_tbl10_cmbSankoMitumori.DisplayMember = "Discript";
            this.base_tbl10_cmbSankoMitumori.ValueMember = "Value";

            // -- 事前打診：１．事前打診状況
            this.prior_tbl01_cmbMitumori.DataSource = cmbPSkMtmr;
            this.prior_tbl01_cmbMitumori.DisplayMember = "Discript";
            this.prior_tbl01_cmbMitumori.ValueMember = "Value";

            // -- 入札：１．入札情報
            this.bid_tbl01_cmbMitumori.DataSource = cmbSkMtmr;
            this.bid_tbl01_cmbMitumori.DisplayMember = "Discript";
            this.bid_tbl01_cmbMitumori.ValueMember = "Value";
            #endregion

            #region 入札方式 -----------------------------------
            discript = "KeiyakuKeitai";
            value = "KeiyakuKeitaiCD";
            table = "Mst_KeiyakuKeitai";
            where = "KeiyakuKeitaiNarabijun < 20";
            DataTable cmbBidHsk = GlobalMethod.getData(discript, value, table, where);
            if (cmbBidHsk != null)
            {
                DataRow dr = cmbBidHsk.NewRow();
                cmbBidHsk.Rows.InsertAt(dr, 0);
            }

            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbNyusatuHosiki.DataSource = cmbBidHsk;
            this.base_tbl10_cmbNyusatuHosiki.DisplayMember = "Discript";
            this.base_tbl10_cmbNyusatuHosiki.ValueMember = "Value";

            // -- 入札：１．入札情報
            this.bid_tbl01_cmbBidhosiki.DataSource = cmbBidHsk;
            this.bid_tbl01_cmbBidhosiki.DisplayMember = "Discript";
            this.bid_tbl01_cmbBidhosiki.ValueMember = "Value";
            #endregion

            #region 入札状況 -----------------------------------
            discript = "RakusatsuShaMei";
            value = "RakusatsuShaID";
            table = "Mst_RakusatsuSha";
            where = "RakusatsuShaNarabijun > 0";
            DataTable cmbBidSts = GlobalMethod.getData(discript, value, table, where);
            if (cmbBidSts != null)
            {
                DataRow dr = cmbBidSts.NewRow();
                cmbBidSts.Rows.InsertAt(dr, 0);
            }
            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbNyusatuStats.DataSource = cmbBidSts;
            this.base_tbl10_cmbNyusatuStats.DisplayMember = "Discript";
            this.base_tbl10_cmbNyusatuStats.ValueMember = "Value";

            // -- 入札：３．入札結果
            this.bid_tbl03_1_cmbBidStatus.DataSource = cmbBidSts;
            this.bid_tbl03_1_cmbBidStatus.DisplayMember = "Discript";
            this.bid_tbl03_1_cmbBidStatus.ValueMember = "Value";
            #endregion

            #region 再委託禁止条項の内容 -----------------------
            DataTable cmbKsNaiyo = new System.Data.DataTable();
            cmbKsNaiyo.Columns.Add("Value", typeof(int));
            cmbKsNaiyo.Columns.Add("Discript", typeof(string));
            cmbKsNaiyo.Rows.Add(0, "");
            cmbKsNaiyo.Rows.Add(1, "禁止(具体的な禁止範囲の記載なし)");
            cmbKsNaiyo.Rows.Add(2, "書面承諾あれば可(具体的な禁止範囲の記載なし)");
            cmbKsNaiyo.Rows.Add(3, "禁止(具体的な禁止範囲の記載あり)");
            cmbKsNaiyo.Rows.Add(4, "書面承諾あれば可(具体的な禁止範囲の記載あり)");
            cmbKsNaiyo.Rows.Add(5, "その他");
            DataTable cmbCaKsNaiyo = cmbKsNaiyo.AsEnumerable().CopyToDataTable();
            DataTable cmbBidKsNaiyo = cmbKsNaiyo.AsEnumerable().CopyToDataTable();
            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbKinsiNaiyo.DataSource = cmbKsNaiyo;
            this.base_tbl10_cmbKinsiNaiyo.DisplayMember = "Discript";
            this.base_tbl10_cmbKinsiNaiyo.ValueMember = "Value";

            // -- 入札：２．再委託禁止条項
            this.bid_tbl02_cmbKinsiNaiyo.DataSource = cmbBidKsNaiyo;
            this.bid_tbl02_cmbKinsiNaiyo.DisplayMember = "Discript";
            this.bid_tbl02_cmbKinsiNaiyo.ValueMember = "Value";

            // -- 契約：１．契約情報
            this.ca_tbl01_cmbKinsiNaiyo.DataSource = cmbCaKsNaiyo;
            this.ca_tbl01_cmbKinsiNaiyo.DisplayMember = "Discript";
            this.ca_tbl01_cmbKinsiNaiyo.ValueMember = "Value";

            #endregion

            #region 再委託禁止条項の記載有無--------------------
            DataTable cmbKsUm = new System.Data.DataTable();
            cmbKsUm.Columns.Add("Value", typeof(int));
            cmbKsUm.Columns.Add("Discript", typeof(string));
            cmbKsUm.Rows.Add(0, "");
            cmbKsUm.Rows.Add(1, "あり");
            cmbKsUm.Rows.Add(2, "なし");
            cmbKsUm.Rows.Add(3, "不明");
            DataTable cmbCaKsUm = cmbKsUm.AsEnumerable().CopyToDataTable();
            DataTable cmbBidKsUm = cmbKsUm.AsEnumerable().CopyToDataTable();
            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbKinsiUmu.DataSource = cmbKsUm;
            this.base_tbl10_cmbKinsiUmu.DisplayMember = "Discript";
            this.base_tbl10_cmbKinsiUmu.ValueMember = "Value";

            // -- 入札：２．再委託禁止条項
            this.bid_tbl02_cmbKinsiUmu.DataSource = cmbBidKsUm;
            this.bid_tbl02_cmbKinsiUmu.DisplayMember = "Discript";
            this.bid_tbl02_cmbKinsiUmu.ValueMember = "Value";

            // -- 契約：１．契約情報
            this.ca_tbl01_cmbKinsiUmu.DataSource = cmbCaKsUm;
            this.ca_tbl01_cmbKinsiUmu.DisplayMember = "Discript";
            this.ca_tbl01_cmbKinsiUmu.ValueMember = "Value";

            #endregion

            #region 最低制限価格有無----------------------------
            DataTable cmbLowestUm = new System.Data.DataTable();
            cmbLowestUm.Columns.Add("Value", typeof(int));
            cmbLowestUm.Columns.Add("Discript", typeof(string));
            cmbLowestUm.Rows.Add(0, "");
            cmbLowestUm.Rows.Add(1, "あり（調査基準価格）");
            cmbLowestUm.Rows.Add(2, "あり（失格）");
            cmbLowestUm.Rows.Add(3, "なし");

            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbLowestUmu.DataSource = cmbLowestUm;
            this.base_tbl10_cmbLowestUmu.DisplayMember = "Discript";
            this.base_tbl10_cmbLowestUmu.ValueMember = "Value";

            // -- 入札：１．入札情報
            this.bid_tbl01_cmbLowestUmu.DataSource = cmbLowestUm;
            this.bid_tbl01_cmbLowestUmu.DisplayMember = "Discript";
            this.bid_tbl01_cmbLowestUmu.ValueMember = "Value";

            #endregion

            #region 業務発注区分--------------------------------
            DataTable cmbGmOrderKb = new System.Data.DataTable();
            cmbGmOrderKb.Columns.Add("Value", typeof(int));
            cmbGmOrderKb.Columns.Add("Discript", typeof(string));
            cmbGmOrderKb.Rows.Add(0, "");
            cmbGmOrderKb.Rows.Add(1, "物品・役務");
            cmbGmOrderKb.Rows.Add(2, "測量・コンサル");
            cmbGmOrderKb.Rows.Add(3, "不明");

            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbOrderKubun.DataSource = cmbGmOrderKb;
            this.base_tbl10_cmbOrderKubun.DisplayMember = "Discript";
            this.base_tbl10_cmbOrderKubun.ValueMember = "Value";

            // -- 入札：１．入札情報
            this.bid_tbl01_cmbOrderKubun.DataSource = cmbGmOrderKb;
            this.bid_tbl01_cmbOrderKubun.DisplayMember = "Discript";
            this.bid_tbl01_cmbOrderKubun.ValueMember = "Value";

            #endregion

            #region 落札者状況----------------------------------
            DataTable cmbRkstSt = new System.Data.DataTable();
            cmbRkstSt.Columns.Add("Value", typeof(int));
            cmbRkstSt.Columns.Add("Discript", typeof(string));
            cmbRkstSt.Rows.Add(0, "");
            cmbRkstSt.Rows.Add(1, "判明");
            cmbRkstSt.Rows.Add(2, "不明");
            cmbRkstSt.Rows.Add(3, "推定");

            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbRakusatuStats.DataSource = cmbRkstSt;
            this.base_tbl10_cmbRakusatuStats.DisplayMember = "Discript";
            this.base_tbl10_cmbRakusatuStats.ValueMember = "Value";

            // -- 入札：３．入札結果
            this.bid_tbl03_1_cmbRakusatuStatus.DataSource = cmbRkstSt;
            this.bid_tbl03_1_cmbRakusatuStatus.DisplayMember = "Discript";
            this.bid_tbl03_1_cmbRakusatuStatus.ValueMember = "Value";

            #endregion

            #region 落札額状況----------------------------------
            DataTable cmbRkstAmt = new System.Data.DataTable();
            cmbRkstAmt.Columns.Add("Value", typeof(int));
            cmbRkstAmt.Columns.Add("Discript", typeof(string));
            cmbRkstAmt.Rows.Add(0, "");
            cmbRkstAmt.Rows.Add(1, "判明");
            cmbRkstAmt.Rows.Add(2, "不明");
            cmbRkstAmt.Rows.Add(3, "推定");

            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbRakusatuAmtStats.DataSource = cmbRkstAmt;
            this.base_tbl10_cmbRakusatuAmtStats.DisplayMember = "Discript";
            this.base_tbl10_cmbRakusatuAmtStats.ValueMember = "Value";

            // -- 入札：３．入札結果
            this.bid_tbl03_1_cmbRakusatuAmtStatus.DataSource = cmbRkstAmt;
            this.bid_tbl03_1_cmbRakusatuAmtStatus.DisplayMember = "Discript";
            this.bid_tbl03_1_cmbRakusatuAmtStatus.ValueMember = "Value";

            #endregion

            #region 当会応札------------------------------------
            DataTable cmbTokai = new System.Data.DataTable();
            cmbTokai.Columns.Add("Value", typeof(int));
            cmbTokai.Columns.Add("Discript", typeof(string));
            cmbTokai.Rows.Add(0, "");
            cmbTokai.Rows.Add(1, "検討中");
            cmbTokai.Rows.Add(2, "応札");
            cmbTokai.Rows.Add(3, "不参加");
            cmbTokai.Rows.Add(4, "辞退");

            // -- 基本情報等一覧：１０．入札情報・入札結果
            this.base_tbl10_cmbTokaiOsatu.DataSource = cmbTokai;
            this.base_tbl10_cmbTokaiOsatu.DisplayMember = "Discript";
            this.base_tbl10_cmbTokaiOsatu.ValueMember = "Value";

            // -- 入札：１．入札情報
            this.bid_tbl01_cmbTokaiOsatu.DataSource = cmbTokai;
            this.bid_tbl01_cmbTokaiOsatu.DisplayMember = "Discript";
            this.bid_tbl01_cmbTokaiOsatu.ValueMember = "Value";

            #endregion

        }

        /// <summary>
        /// 受託課所支部コンボボックスリストの設定処理
        /// </summary>
        /// <param name="nendo"></param>
        private void set_combo_shibu(string nendo)
        {
            //受託課所支部
            string SelectedValue = "";
            if (base_tbl02_cmbJyutakuKasyoSibu.Text != "")
            {
                SelectedValue = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString();
            }
            //SQL変数
            string discript = "Mst_Busho.ShibuMei + ' ' + IsNull(Mst_Busho.KaMei,'') ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            StringBuilder where = new StringBuilder();
            where.Append("GyoumuBushoCD < '999990'");
            where.Append(" AND BushoNewOld <= 1");
            where.Append(" AND BushoEntryHyoujiFlg = 1");
            where.Append(" AND ISNULL(BushoDeleteFlag,0) = 0");
            where.Append(" AND NOT GyoumuBushoCD LIKE '121%'");
            where.Append(" AND ISNULL(KashoShibuCD,'') <> ''");
            // 工期開始年度よりの条件
            int FromNendo;
            if (int.TryParse(nendo, out FromNendo))
            {
                int ToNendo = int.Parse(nendo) + 1;

                where.Append(" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '").Append(ToNendo).Append("/3/31' )");
                where.Append(" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '").Append(FromNendo).Append("/4/1' )");
                
                where.Append("");
                // 工期開始年度が2021年度未満の場合、旧積シス（127910）をコンボに追加する
                if (FromNendo < 2021)
                {
                    // イレギュラー対応の為、以下の条件を付与
                    where.Append(" OR (GyoumuBushoCD = '127910') ");
                }
            }

            Console.WriteLine(where);
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where.ToString());
            base_tbl02_cmbJyutakuKasyoSibu.DataSource = combodt;
            base_tbl02_cmbJyutakuKasyoSibu.DisplayMember = "Discript";
            base_tbl02_cmbJyutakuKasyoSibu.ValueMember = "Value";
            if (SelectedValue != "")
            {
                base_tbl02_cmbJyutakuKasyoSibu.SelectedValue = SelectedValue;
            }
        }

        /// <summary>
        /// 応援依頼先コントロール設定
        /// </summary>
        /// <param name="nendo"></param>
        private void set_oensaki_list(string nendo)
        {
            //SQL変数
            string discript = "CASE WHEN ISNULL(Mst_Busho.KaMei, '') = '' THEN Mst_Busho.BushokanriboKameiRaku ELSE Mst_Busho.KaMei END ";
            string value = "Mst_Busho.GyoumuBushoCD ";
            string table = "Mst_Busho";
            StringBuilder where = new StringBuilder();
            where.Append("GyoumuBushoCD < '999990'");
            where.Append(" AND BushoNewOld <= 1");
            where.Append(" AND BushoEntryHyoujiFlg = 1");
            where.Append(" AND ISNULL(BushoDeleteFlag,0) = 0");
            where.Append(" AND NOT GyoumuBushoCD LIKE '121%'");
            where.Append(" AND ISNULL(KashoShibuCD,'') <> ''");

            // STEP3 No1460
            where.Append(" AND BushoEntoriNarabijun > 50");
            where.Append(" AND (GyoumuBushoCD >= '160000'");
            where.Append(" OR (GyoumuBushoCD > '127100' AND GyoumuBushoCD < '127800'))");
            // 工期開始年度よりの条件
            int FromNendo;
            if (int.TryParse(nendo, out FromNendo))
            {
                int ToNendo = int.Parse(nendo) + 1;

                where.Append(" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '").Append(ToNendo).Append("/3/31' )");
                where.Append(" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '").Append(FromNendo).Append("/4/1' )");

                where.Append("");
                // 工期開始年度が2021年度未満の場合、旧積シス（127910）をコンボに追加する
                if (FromNendo < 2021)
                {
                    // イレギュラー対応の為、以下の条件を付与
                    where.Append(" OR (GyoumuBushoCD = '127910') ");
                }
            }
            where.Append(" ORDER BY BushoEntoriNarabijun");
            Console.WriteLine(where);

            base_tbl07_3_tblOenIrai1.Controls.Clear();
            base_tbl07_3_tblOenIrai2.Controls.Clear();
            base_tbl07_3_tblOenIrai3.Controls.Clear();
            base_tbl07_3_tblOenIrai2.Height = 0;
            base_tbl07_3_tblOenIrai3.Height = 0;
            //応援依頼先リスト取得
            DataTable dt = GlobalMethod.getData(discript, value, table, where.ToString());
            if(dt != null && dt.Rows.Count > 0)
            {
                // ２行目
                if (dt.Rows.Count > 10)
                {
                    base_tbl07_3_tblOenIrai2.Height = 27;
                }
                // ３行目
                if (dt.Rows.Count > 20)
                {
                    base_tbl07_3_tblOenIrai3.Height = 27;
                }
                int iLocation = 0;
                int iWidth = 100;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];
                    System.Windows.Forms.CheckBox chk = new System.Windows.Forms.CheckBox();
                    chk.Anchor = System.Windows.Forms.AnchorStyles.Left;
                    chk.AutoSize = true;
                    chk.Name = "base_tbl07_3_chk" + (i + 1).ToString();
                    chk.Size = new System.Drawing.Size(iWidth, 21);
                    chk.Text = dr[1].ToString();
                    chk.Tag = dr[0].ToString();
                    chk.UseVisualStyleBackColor = true;

                    if (i < 10)
                    {
                        //base_tbl07_3_tblOenIrai1.Controls.Add(chk, i, 0);
                        iLocation = i * iWidth + 3;
                        chk.Location = new System.Drawing.Point(iLocation, 3);
                        base_tbl07_3_tblOenIrai1.Controls.Add(chk);
                    }
                    else if(i >= 10 && i < 20)
                    {
                        iLocation = (i -10) * iWidth + 3;
                        //base_tbl07_3_tblOenIrai2.Controls.Add(chk, i-10, 0);
                        chk.Location = new System.Drawing.Point(iLocation, 3);
                        base_tbl07_3_tblOenIrai2.Controls.Add(chk);
                    }
                    else if (i >= 20 && i < 30)
                    {
                        //base_tbl07_3_tblOenIrai3.Controls.Add(chk, i - 20, 0);
                        iLocation = (i - 20) * iWidth + 3;
                        chk.Location = new System.Drawing.Point(iLocation, 3);
                        base_tbl07_3_tblOenIrai3.Controls.Add(chk);
                    }
                }
            }
            else
            {
                base_tbl07_3_tblOenIrai1.Height = 0;
            }

            // チェック項目設定処理
            if (copy != COPY.HC)
            {
                List<string> lst = EntryInputDbClass.getAnkenOuenIraisaki(AnkenID);
                if (lst.Count > 0)
                {
                    if (base_tbl07_3_tblOenIrai1.Height > 0)
                    {
                        foreach (Control child in base_tbl07_3_tblOenIrai1.Controls)
                        {
                            if (child is System.Windows.Forms.CheckBox)
                            {
                                System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)child;
                                if (lst.Contains(chk.Tag.ToString()))
                                {
                                    chk.Checked = true;
                                }
                            }
                        }
                    }
                    if (base_tbl07_3_tblOenIrai2.Height > 0)
                    {
                        foreach (Control child in base_tbl07_3_tblOenIrai2.Controls)
                        {
                            if (child is System.Windows.Forms.CheckBox)
                            {
                                System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)child;
                                if (lst.Contains(chk.Tag.ToString()))
                                {
                                    chk.Checked = true;
                                }
                            }
                        }
                    }
                    if (base_tbl07_3_tblOenIrai3.Height > 0)
                    {
                        foreach (Control child in base_tbl07_3_tblOenIrai3.Controls)
                        {
                            if (child is System.Windows.Forms.CheckBox)
                            {
                                System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)child;
                                if (lst.Contains(chk.Tag.ToString()))
                                {
                                    chk.Checked = true;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 案件フォルダパスを取得する
        /// </summary>
        /// <returns></returns>
        private string getBaseFolderPath(string sFolderBushoCD)
        {
            String discript = "FolderPath";
            String value = "FolderPath ";
            String table = "M_Folder";
            String where = "MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + sFolderBushoCD + "' ";

            // //xxxx/00Cyousa/00調査情報部門共有/$NENDO$/200受託調査関連
            // フォルダ関連は工期開始年度で作成する
            string FolderBase = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", base_tbl03_cmbKokiStartYear.SelectedValue.ToString());
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

            //No1563 1314　北海道のフォルダ名が間違ってる。　×　010北道　○　010北海
            // 工期開始年度　2021年度まで、　010北道
            // 工期開始年度　2022年度から　　010北海
            int koukinendo = 0;
            if (int.TryParse(base_tbl03_cmbKokiStartYear.SelectedValue.ToString(), out koukinendo))
            {
                FolderPath = change_hokaido_path(FolderPath, koukinendo);
            }
            return FolderPath;
        }

        /// <summary>
        /// 案件フォルダパスを取得する
        /// </summary>
        /// <returns></returns>
        private string getReplaceFolderPath(string sFolderBushoCD, string sLogMsg)
        {
            String discript = "FolderPath";
            String value = "FolderPath ";
            String table = "M_Folder";
            String where = "MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + sFolderBushoCD + "' ";

            string FolderPath = "";

            DataTable dt = new System.Data.DataTable();
            dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null && dt.Rows.Count > 0)
            {
                // $FOLDER_BASE$/004 本部
                FolderPath = dt.Rows[0][0].ToString();
                // 課所支部のフォルダ部分のみとする
                FolderPath = FolderPath.Replace(@"$FOLDER_BASE$/", "");
                // 課所支部のフォルダ部分のみとする $FOLDER_BASE$ しかない場合の対応
                FolderPath = FolderPath.Replace(@"$FOLDER_BASE$", "");
            }
            else
            {
                // エラー
                GlobalMethod.outputLogger("Execute_SQL", sLogMsg + sFolderBushoCD + " のフォルダパスが取得できずにエラー", "ID:" + AnkenID + " mode:0", "DEBUG");
            }
            
            return FolderPath;
        }

        /// <summary>
        /// フォルダ変更コントローラセット
        /// 　表示／非表示設定
        /// </summary>
        /// <param name="isVisible"></param>
        private void setVisibleToRenameFolder(bool isVisible)
        {
            base_tbl02_btnRenameFolder.Visible = isVisible;
            base_tbl02_txtRenameFolder.Visible = isVisible;
            base_tbl02_lblRenameFolder.Visible = isVisible;
        }

        /// <summary>
        /// 各モードでヘッダー部が表示／非表示設定
        /// 　ボタン、伝票出力回数など
        /// </summary>
        private void setVisibleHeaders()
        {
            if(mode == MODE.CHANGE)
            {
                // 伝票変更なら：契約情報表示のみ
                this.tab.TabPages.Remove(this.tabBase);
                this.tab.TabPages.Remove(this.tabPrior);
                this.tab.TabPages.Remove(this.tabBid);
                this.tab.TabPages.Remove(this.tabTE);

                ca_lblChangeCnt.Text = (Convert.ToInt32(AnkenData_H.Rows[0]["AnkenKaisuu"]) + 1).ToString();// 伝票変更回数
                lblTitle.Text = "■エントリくん 変更伝票"; // 画面タイトル変更
                                                //ボタン
                btnNewByCopy.Visible = false;   //この業務を元に新規登録
                btnUpdate.Visible = false;      // 更新
                btnNewByBranchNo.Visible = false;   //この案件番号の枝番で受託番号を作成する
                btnNewByOrder.Visible = false;
                btnDelete.Visible = false; // 削除ボタン
                //契約タブ
                ca_tblComment.Visible = true;
                ca_tblCs.Visible = true;
                ca_tblCaBtn.Visible = false;
            }
            else if (mode == MODE.INSERT || mode == MODE.PLAN)
            {
                // 新規登録なら：基本情報のみ表示する
                this.tab.TabPages.Remove(this.tabPrior);
                this.tab.TabPages.Remove(this.tabBid);
                this.tab.TabPages.Remove(this.tabCA);
                this.tab.TabPages.Remove(this.tabTE);

                //一部名称の変更
                btnUpdate.Text = "新規登録";
                lblTitle.Text = "■エントリくん 新規追加";
                //一部ボタンの非表示化
                tblAKInfo.Visible = false;
                btnNewByCopy.Visible = false;
                btnNewByBranchNo.Visible = false;
                btnNewByOrder.Visible = false;
                btnDelete.Visible = false;
                //コピペテキストと反映するボタンが配置されたテールブルレイアウトパネルを表示有効化する
                tblTayori.Visible = true;
                //案件番号などの表が乗ってる列をサイズゼロにする
                headerTbl.ColumnStyles[5] = new ColumnStyle(SizeType.Absolute, 0.0F);
                //削除ボタン乗ってる列のサイズをゼロパーセントにする。
                headerTbl.ColumnStyles[6] = new ColumnStyle(SizeType.Percent, 0);
                //else追加　新規・計画以外は反映用のテキストとボタンを消す
            }else if(mode == MODE.VIEW)
            {
                //一部名称の変更
                lblTitle.Text = "■エントリくん 参照";
                //一部ボタンの非表示化
                btnNewByCopy.Visible = false;
                btnUpdate.Visible = false;
                btnNewByBranchNo.Visible = false;
                btnNewByOrder.Visible = false;
                btnDelete.Visible = false;
                ca_tblComment.Visible = false;
                ca_tblCaBtn.Visible = false;
                ca_tblCs.Visible = false;
                ca_tblButton.Visible = false;
            }
        }

        /// <summary>
        /// 各タグページ項目
        ///     表示／非表示設定
        ///     編集可否設定
        /// </summary>
        private void setVisibleDetails()
        {
            base_tbl08_c1FlexGrid.Height = 4 + 22 * base_tbl08_c1FlexGrid.Rows.Count;
            // 新規登録
            if (mode == MODE.INSERT || mode == MODE.PLAN)
            {
                // 基本情報 ----------------------------
                // フォルダ変更
                setVisibleToRenameFolder(false);

                // 案件変更履歴
                base_tbl02_txtAnkenChanger.Visible = false;
                base_tbl02_txtAnkenChangDt.Visible = false;
                base_tbl02_txtAnkenChangHistory.Visible = false;
                base_tbl02_lblAnkenChanger.Visible = false;
                base_tbl02_lblAnkenChangDt.Visible = false;
                base_tbl02_lblAnkenChangHistory.Visible = false;

                //base_tbl07_input3.Visible = false;
                base_tbl07_3_lblComent.Visible = false;
                base_tbl07_input4.Visible = false;
                base_tbl07_input5.Visible = false;
                base_tblMemo_lblTitle.Text = "※各タブ（「事前打診」「入札」）への引用について";
                base_tblMemo_lblRow1Note.Visible = false;
                base_tblMemo_lblRow2Note.Text = "※下記「9.事前打診・参考見積」以降の引用データは、「事前打診」「入札」タブに反映されます。";

                //base_tblMemo_parent.Visible = false;
                //// ９．事前打診・参考見積
                //base_tbl09_parent.Visible = false;
                //// １０．入札状況・入札結果
                //base_tbl10_parent.Visible = false;
                // １１．契約状況
                base_tbl11_parent.Visible = false;
            }
            else
            {
                //新規・計画以外は反映用のテキストとボタンを消す
                txtTayoriData.Visible = false;
                btnHanei.Visible = false;
                tblTayori.Visible = false;
                //コピペテキストと反映するボタンの配置された列幅をゼロにする
                headerTbl.ColumnStyles[7] = new ColumnStyle(SizeType.Absolute, 0.0F);
                // 基本情報-------
                // -- １．進捗段階
                base_tbl01_dtpDtPrior.Enabled = false;  // ■事前打診 登録日
                base_tbl01_dtpDtBid.Enabled = false;    // ■入札 登録日
                base_tbl01_dtpDtCa.Enabled = false;     // ■契約 登録日

                // -- ２．基本情報 設定なし
                // -- ３．案件情報
                base_tbl03_cmbKokiStartYear.Enabled = false;    // 工期開始年度
                base_tbl03_cmbKokiSalesYear.Enabled = false;    // 売上年度
                // -- ４～６ 設定なし
                // -- ７．配分情報・業務内容
                base_tbl07_3_lblComent.Visible = true;
                //base_tbl07_input3.Visible = true;
                base_tbl07_input4.Visible = true;
                base_tbl07_input5.Visible = true;

                //base_tblMemo_lblTitle.Text = "";
                //base_tblMemo_lblRow1Note.Visible = true;
                //base_tblMemo_lblRow2Note.Text = "※下記「9.事前打診・参考見積」以降は、閲覧フォームです（入力不可）。「事前打診」「入札」「契約」タブに登録されたデータが表示されます。";
                // -- ９．事前打診・参考見積
                //base_tbl09_parent.Visible = true;
                base_tbl09_numSankomitumoriAmt.ReadOnly = true; // 参考見積額(税抜)
                base_tbl09_cmbNotOrderReason.Enabled = false;   // 「発注無し」の理由
                base_tbl09_cmbNotOrderStats.Enabled = false;    // 未発注状況
                base_tbl09_dtpOrderYoteiDt.Enabled = false; // 発注予定・見込日
                base_tbl09_cmbOrderIyoku.Enabled = false;   // 受注意欲
                base_tbl09_cmbSankomitumori.Enabled = false;    // 参考見積対応
                base_tbl09_dtpJizenDasinIraiDt.Enabled = false; // 事前打診依頼日
                base_tbl09_txtOthenComment.ReadOnly = true; // 「その他」の内容
                // -- １０．入札情報・入札結果
                //base_tbl10_parent.Visible = true;
                base_tbl10_cmbRakusatuAmtStats.Enabled = false; // 落札額状況
                base_tbl10_txtRakusatuAmt.ReadOnly = true; // 落札額(税抜)
                base_tbl10_txtRakusatuSya.ReadOnly = true; // 落札者
                base_tbl10_cmbRakusatuStats.Enabled = false; // 落札者状況
                base_tbl10_txtOsatuNum.ReadOnly = true; // 応札数
                base_tbl10_txtYoteiAmt.ReadOnly = true; // 予定価格(税抜)
                base_tbl10_cmbNyusatuStats.Enabled = false; // 入札状況
                base_tbl10_cmbKinsiNaiyo.Enabled = false; // 再委託禁止条項の内容
                base_tbl10_cmbTokaiOsatu.Enabled = false; // 当会応札
                base_tbl10_cmbKinsiUmu.Enabled = false; // 再委託禁止条項の記載有無
                base_tbl10_cmbOrderKubun.Enabled = false; // 業務発注区分
                base_tbl10_numSankoMitumoriAmt.ReadOnly = true; // 参考見積額(税抜)
                base_tbl10_cmbOrderIyoku.Enabled = false; // 受注意欲
                base_tbl10_cmbSankoMitumori.Enabled = false; // 参考見積対応
                base_tbl10_dtpNyusatuDt.Enabled = false; // 入札(予定)日
                base_tbl10_cmbLowestUmu.Enabled = false; // 最低制限価格有無
                base_tbl10_cmbNyusatuHosiki.Enabled = false; // 入札方式
                base_tbl10_txtOtherNaiyo.ReadOnly = true; // その他の内容


                // -- １１．契約情報
                //base_tbl11_parent.Visible = true;
                base_tbl11_1_txtKeiyakuAmt.ReadOnly = true; // 契約金額(税抜)
                base_tbl11_1_dtpKianDt.Enabled = false; // 起案日
                base_tbl11_1_chkKianzumi.Enabled = false; // 起案済
                base_tbl11_1_dtpKeiyakuChangeDt.Enabled = false; // 契約締結(変更)日

                base_tbl11_2_numAmt1.ReadOnly = true; // 調査部
                base_tbl11_2_numAmt2.ReadOnly = true; // 事業普及部
                base_tbl11_2_numAmt3.ReadOnly = true; // 情報システム部
                base_tbl11_2_numAmt4.ReadOnly = true; // 総合研究所

                base_tbl11_3_numAmt1.ReadOnly = true; // 調査部
                base_tbl11_3_numAmt2.ReadOnly = true; // 事業普及部
                base_tbl11_3_numAmt3.ReadOnly = true; // 情報システム部
                base_tbl11_3_numAmt4.ReadOnly = true; // 総合研究所

                // 入札-------
                
                // 編集時、情報メッセージ表示するように
                if (mode == MODE.UPDATE) set_error(GlobalMethod.GetMessage("I00004", ""));

                bool isKian = ca_tbl01_chkKian.Checked;//起案済
                if (isKian && mode != MODE.CHANGE && mode != MODE.INSERT && mode != MODE.PLAN)
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
                    ca_tbl01_cmbAnkenKubun.Enabled = (AnkenData_K.Rows[0]["AnkenSakuseiKubun"].ToString() != "01");

                    //ヘッダーボタン
                    btnDelete.Visible = true;

                    //基本情報 ---------------------
                    // １．進捗階段
                    base_tbl01_chkJizendasin.Enabled = false;
                    base_tbl01_chkNyusatu.Enabled = false;
                    base_tbl01_chkKeiyaku.Enabled = false;

                    // ２．基本情報
                    base_tbl02_cmbJyutakuKasyoSibu.Enabled = false; // 受託課所支部
                    base_tbl02_picKeiyakuTanto.Visible = false;// 契約担当者プロンプト
                    //受託フォルダ非表示設定
                    setVisibleToRenameFolder(false);

                    //案件情報
                    base_tbl03_txtGyomuName.ReadOnly = true; // 業務名称
                    base_tbl03_cmbKeiyakuKubun.Enabled = false; // 契約区分
                    base_tbl03_dtpKokiTo.Enabled = false;//工期至
                    base_tbl03_txtAnkenMemo.ReadOnly = true;//案件メモ
                    //　４．発注者情報
                    base_tbl04_picOrderCd.Visible = false; // 発注者コードプロンプト
                    base_tbl04_txtOrderKamei.ReadOnly = true; // 発注者課名

                    // ５．と６．発注担当者情報
                    // ７．業務内容
                    //部門配分base_tbl07_1_numPasset1
                    base_tbl07_1_numPercent1.ReadOnly = true; // 部門配分 引合・入札 配分率 調査部
                    base_tbl07_1_numPercent2.ReadOnly = true; // 部門配分 引合・入札 配分率 事業普及部
                    base_tbl07_1_numPercent3.ReadOnly = true; // 部門配分 引合・入札 配分率 情報システム部
                    base_tbl07_1_numPercent4.ReadOnly = true; // 部門配分 引合・入札 配分率 総合研究所
                    //業務配分
                    base_tbl07_2_numPercent1.ReadOnly = true;
                    base_tbl07_2_numPercent2.ReadOnly = true;
                    base_tbl07_2_numPercent3.ReadOnly = true;
                    base_tbl07_2_numPercent4.ReadOnly = true;
                    base_tbl07_2_numPercent5.ReadOnly = true;
                    base_tbl07_2_numPercent6.ReadOnly = true;
                    base_tbl07_2_numPercent7.ReadOnly = true;
                    base_tbl07_2_numPercent8.ReadOnly = true;
                    base_tbl07_2_numPercent9.ReadOnly = true;
                    base_tbl07_2_numPercent10.ReadOnly = true;
                    base_tbl07_2_numPercent11.ReadOnly = true;
                    base_tbl07_2_numPercent12.ReadOnly = true;
                    // 応援依頼情報
                    // No.1533 削除
                    //base_tbl07_3_cmbOen.Enabled = false;
                    base_tbl07_3_txtOenMemo.ReadOnly = true;
                    base_tbl07_3_tblOenIrai1.Enabled = false;
                    base_tbl07_3_tblOenIrai2.Enabled = false;
                    base_tbl07_3_tblOenIrai3.Enabled = false;

                    // ８．過去案件情報
                    // ９．１０．１１。↑で設定済

                    // 事前打診 ---------------------
                    // １．事前打診状況
                    prior_tbl01_dtpDasinIraiDt.Enabled = false;// 事前打診依頼日
                    prior_tbl01_cmbMitumori.Enabled = false;//参考見積対応
                    prior_tbl01_txtMitumoriAmt.ReadOnly = true;//参考見積額（税抜）
                    prior_tbl01_cmbOrderIyoku.Enabled = false;//受注意欲
                    prior_tbl01_dtpOrderYoteiDt.Enabled = false;// 発注予定（見込）日
                    prior_tbl01_txtAnkenMemo.ReadOnly = true; // 案件メモ（事前打診）
                    // ２．未発注
                    prior_tbl02_dtpNotOrderDt.Enabled = false;// 未発注の登録日
                    prior_tbl02_cmbNotOrderStats.Enabled = false;// 未発注状況
                    prior_tbl02_cmbNotOrderReason.Enabled = false;// 「発注なし」の理由
                    prior_tbl02_txtOtherNaiyo.ReadOnly = true;//「そのた」の内容
                    prior_tbl02_txtAnkenMemo.ReadOnly = true;//案件メモ（見発注）

                    //入札 ---------------------
                    // １．入札情報
                    bid_tbl01_dtpBidInfoDt.Enabled = false;//入札情報登録日
                    bid_tbl01_cmbOrderKubun.Enabled = false;//業務発注区分
                    bid_tbl01_cmbBidhosiki.Enabled = false;//入札方式
                    bid_tbl01_cmbLowestUmu.Enabled = false;//最低制限価格有無
                    bid_tbl01_dtpBidYoteiDt.Enabled = false;//入札(予定)日
                    bid_tbl01_cmbMitumori.Enabled = false;//参考見積対応
                    bid_tbl01_txtMitumoriAmt.ReadOnly = true;//参考見積額(税抜)
                    bid_tbl01_cmbOrderIyoku.Enabled = false;//受注意欲
                    bid_tbl01_cmbTokaiOsatu.Enabled = false;//当会応札

                    //２．再委託禁止条項
                    bid_tbl02_cmbKinsiUmu.Enabled = false;//再委託禁止条項の記載有無
                    bid_tbl02_cmbKinsiNaiyo.Enabled = false;//再委託禁止条項の内容
                    bid_tbl02_txtOtherNaiyo.ReadOnly = true;//その他の内容
                    // ３．入札結果
                    bid_tbl03_1_cmbBidStatus.Enabled = false;//入札状況
                    bid_tbl03_1_txtBidMemo.ReadOnly = true;//案件メモ（入札）

                    //契約 ---------------------
                    // チェック用帳票出力・内容確認ボタンは常時使用可とするためコメントアウト
                    ca_btnKian.Enabled = false;
                    ca_btnKian.BackColor = Color.DimGray;
                    ca_btnOutSheet.Enabled = true;
                    ca_btnOutSheet.BackColor = Color.FromArgb(42, 78, 122);
                    ca_btnOutSheet.ForeColor = Color.White;
                    ca_btnChangeSlip.Enabled = true;
                    ca_btnChangeSlip.BackColor = Color.FromArgb(42, 78, 122);
                    ca_btnChangeSlip.ForeColor = Color.White;

                    // Role:2システム管理者 以外の場合、起案解除は非表示
                    if (UserInfos[4].Equals("2"))
                    {
                        // 起案解除
                        ca_btnKianKaijyo.Enabled = true;
                        ca_btnKianKaijyo.BackColor = Color.FromArgb(42, 78, 122);
                        ca_btnKianKaijyo.ForeColor = Color.White;
                    }
                    else
                    {
                        ca_btnKianKaijyo.Visible = false;
                    }

                    //１．契約情報
                    ca_tbl01_dtpKianDt.Enabled = false; // 起案日
                    ca_tbl01_btnSetting.Visible = false; // 工期末日付、及び、請求（1回目）に設定
                    ca_tbl01_txtTax.ReadOnly = true; // 消費税率
                    ca_tbl01_txtZeinukiAmt.ReadOnly = true; // 税抜（自動計算用）
                    ca_tbl01_txtJyutakuGaiAmt.ReadOnly = true; // 受託外金額（税込）
                    ca_tbl01_txtRiyu.ReadOnly = true; // 変更・中止理由
                    ca_tbl01_txtAnkenMemo.ReadOnly = true; // 案件メモ（契約）
                    ca_tbl01_txtBiko.ReadOnly = false; // 備考
                    ca_tbl01_chkCaSyosya.Enabled = false; // 契約書写
                    ca_tbl01_chkSiyosyo.Enabled = false; // 特記仕様書
                    ca_tbl01_chkMitumorisyo.Enabled = false; // 見積書
                    ca_tbl01_chkTanpinTyosa.Enabled = false; // 単品調査内訳書
                    ca_tbl01_chkOther.Enabled = false; // その他
                    ca_tbl01_txtOtherBiko.ReadOnly = true; // その他備考
                    ca_tbl01_txtTosyo.ReadOnly = true; // 契約図書
                    ca_tbl01_chkRibcAri.Enabled = false;//RIBC用単価データ
                    ca_tbl01_chkSasya.Enabled = false;//サ社経由
                    ca_tbl01_chkRibcSyo.Enabled = false;//RIBC用単価契約書
                    ca_tbl01_cmbKinsiUmu.Enabled = false;//再委託禁止条項の記載有無
                    ca_tbl01_cmbKinsiNaiyo.Enabled = false;//再委託禁止条項の内容
                    ca_tbl01_txtOtherNaiyo.ReadOnly = true;//その他の内容

                    //２．配分情報・業務内容
                    // コピーボタン
                    ca_tbl02_btnTSbu.Visible = false;
                    ca_tbl02_btnJGbu.Visible = false;
                    ca_tbl02_btnJHbu.Visible = false;
                    ca_tbl02_btnSGSyo.Visible = false;

                    // 部門配分
                    ca_tbl02_AftCaBm_numPercent1.ReadOnly = true;
                    ca_tbl02_AftCaBm_numPercent2.ReadOnly = true;
                    ca_tbl02_AftCaBm_numPercent3.ReadOnly = true;
                    ca_tbl02_AftCaBm_numPercent4.ReadOnly = true;
                    ca_tbl02_AftCaBmZeikomi_numAmt1.ReadOnly = true;
                    ca_tbl02_AftCaBmZeikomi_numAmt2.ReadOnly = true;
                    ca_tbl02_AftCaBmZeikomi_numAmt3.ReadOnly = true;
                    ca_tbl02_AftCaBmZeikomi_numAmt4.ReadOnly = true;

                    // 調査部　業務別配分
                    ca_tbl02_AftCaTs_numPercent1.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent2.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent3.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent4.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent5.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent6.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent7.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent8.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent9.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent10.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent11.ReadOnly = true;
                    ca_tbl02_AftCaTs_numPercent12.ReadOnly = true;

                    // ３．単契等の見込補正額(年度内)
                    // ４．年度繰越額(年度跨ぎ)
                    // ５．管理者・担当者
                    ////管理者・技術者
                    ca_tbl05_txtKanri.Enabled = false;
                    ca_tbl05_txtSyosa.Enabled = false;
                    ca_tbl05_txtSinsa.Enabled = false;
                    ca_tbl05_txtGyomu.Enabled = false;
                    ca_tbl05_txtMadoguchi.Enabled = false;
                    // ６．売上計上情報
                    // 各ボタン
                    ca_tbl06_btnChosa.Visible = false;
                    ca_tbl06_btnJigyoHukyu.Visible = false;
                    ca_tbl06_btnJohoSystem.Visible = false;
                    ca_tbl06_btnSogoKenkyu.Visible = false;

                    ca_tbl06_c1FlexGrid.Cols[1].AllowEditing = false;
                    ca_tbl06_c1FlexGrid.Cols[3].AllowEditing = false;
                    ca_tbl06_c1FlexGrid.Cols[9].AllowEditing = false;
                    ca_tbl06_c1FlexGrid.Cols[11].AllowEditing = false;
                    ca_tbl06_c1FlexGrid.Cols[17].AllowEditing = false;
                    ca_tbl06_c1FlexGrid.Cols[19].AllowEditing = false;
                    ca_tbl06_c1FlexGrid.Cols[25].AllowEditing = false;
                    ca_tbl06_c1FlexGrid.Cols[27].AllowEditing = false;

                    // ７．請求書情報
                    ca_tbl07_btnCopy.Visible = false;
                    ca_tbl07_dtpRequst1.Enabled = false;
                    ca_tbl07_dtpRequst2.Enabled = false;
                    ca_tbl07_dtpRequst3.Enabled = false;
                    ca_tbl07_dtpRequst4.Enabled = false;
                    ca_tbl07_dtpRequst5.Enabled = false;
                    ca_tbl07_dtpRequst6.Enabled = false;
                    ca_tbl07_txtRequst1.ReadOnly = true;
                    ca_tbl07_txtRequst2.ReadOnly = true;
                    ca_tbl07_txtRequst3.ReadOnly = true;
                    ca_tbl07_txtRequst4.ReadOnly = true;
                    ca_tbl07_txtRequst5.ReadOnly = true;
                    ca_tbl07_txtRequst6.ReadOnly = true;
                }

                if (mode == MODE.VIEW)
                {
                    ca_tbl01_btnSetting.Visible = false;
                    ca_tbl01_cmbAnkenKubun.Enabled = true;
                    ca_tbl01_txtAnkenName.ReadOnly = false;
                    ca_tbl01_lblRiyu.BackColor = Color.FromArgb(252, 228, 214);
                }

                if (mode == MODE.CHANGE)
                {
                    //赤黒作成時に起案日をクリア対応
                    ca_tbl01_dtpKianDt.Text = "";
                    ca_tbl01_dtpKianDt.CustomFormat = " ";

                    // 工期自と工期至活性にする
                    ca_tbl01_dtpKokiFrom.Enabled = true;
                    ca_tbl01_dtpKokiTo.Enabled = true;

                    //契約：契約金額変更（税抜） 表示設定
                    ca_tbl01_numChangeAmt.Visible = true;
                    ca_tbl01_lblChangeAmt.Visible = true;
                    ca_tbl01_lblChangeAmtComment.Visible = true;
                    ca_tbl01_cmbAnkenKubun.Enabled = true;
                    ca_tbl01_txtAnkenName.ReadOnly = false;
                    ca_tbl01_lblRiyu.BackColor = Color.FromArgb(252, 228, 214);

                    ca_tbl01_hidAkaden.Visible = false;
                    ca_tbl01_hidKuroden.Visible = false;

                    // 売上年度は編集不可
                    ca_tbl01_cmbSalesYear.Enabled = false;
                }

                bid_tbl03_4_c1FlexGrid.Height = 4 + 22 * bid_tbl03_4_c1FlexGrid.Rows.Count;
                ca_tbl05_txtTanto_c1FlexGrid.Height = 4 + 22 * ca_tbl05_txtTanto_c1FlexGrid.Rows.Count;
                ca_tbl06_c1FlexGrid.Height = 4 + 22 * ca_tbl06_c1FlexGrid.Rows.Count;
                te_c1FlexGrid.Height = 4 + 22 * te_c1FlexGrid.Rows.Count;
            }

            //受託番号が採番されていない場合、「この案件番号の枝番で受託番号を作成する」ボタンを非表示
            if (base_tbl02_txtJyutakuEdNo.Text == "")
            {
                btnNewByBranchNo.Visible = false;
            }

            // 起案 or 起案解除した場合に、契約タブに移動する
            if (KianKaijoFLG | KianFLG)
            {
                tab.SelectedTab = tabCA;
                KianKaijoFLG = false;
                KianFLG = false;
            }
        }

        /// <summary>
        /// フォーム閉じる
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Entry_Input_New_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Owner.Visible == false && ownerflg)
            {
                this.Owner.Show();
                this.Owner.Close();
            }
        }

        private void Entry_Input_New_KeyDown(object sender, KeyEventArgs e)
        {
            Control c = this.ActiveControl;
            if (c != null) return;

            //レイアウトロジックを停止する
            this.SuspendLayout();

            // ↑↓押下時、コンボボックスがアクティブだった場合は、コンボの値変更を優先し、
            // 画面スクロールは動かさない
            // タブのタイトルを取得 引合、入札、契約、技術者評価
            string tabName = this.tab.SelectedTab.Text;
            TabPage tab = null;

            switch (tabName)
            {
                case "基本情報等一覧":
                    tab = tabBase;
                    break;
                case "事前打診":
                    tab = tabPrior;
                    break;
                case "入札":
                    tab = tabBid;
                    break;
                case "契約":
                    tab = tabCA;
                    break;
                case "技術者評価":
                    tab = tabTE;
                    break;
            }
            // コンボボックス以外で
            if (c != null && (c.GetType().Equals(typeof(ComboBox))
                || c.GetType().Equals(typeof(C1.Win.C1FlexGrid.C1FlexGrid))
                || c.GetType().ToString().Equals("C1.Win.C1FlexGrid.GridEditorTextBox")
                || c.GetType().ToString().Equals("C1.Win.C1FlexGrid.GridEditorComboBox")
                || c.GetType().ToString().Equals("C1.Win.C1FlexGrid.GridEditorNumericTextBox")
                || c.GetType().ToString().Equals("C1.Win.C1FlexGrid.GridEditorDatePicker")
                ))
            {
                if (tab != null)
                {
                    if (e.KeyCode == Keys.PageDown) { 
                        tab.AutoScrollPosition = new System.Drawing.Point(-tab.AutoScrollPosition.X, -tab.AutoScrollPosition.Y + 600);
                    }else if (e.KeyCode == Keys.PageUp)
                    {
                        tab.AutoScrollPosition = new System.Drawing.Point(-tab.AutoScrollPosition.X, -tab.AutoScrollPosition.Y - 600);
                    }
                }
            }
            else
            {
                if (tab != null)
                {
                    if (e.KeyCode == Keys.PageDown || e.KeyCode == Keys.Down)
                    {
                        tab.AutoScrollPosition = new System.Drawing.Point(-tab.AutoScrollPosition.X, -tab.AutoScrollPosition.Y + 600);
                    }
                    else if (e.KeyCode == Keys.PageUp || e.KeyCode == Keys.Up)
                    {
                        tab.AutoScrollPosition = new System.Drawing.Point(-tab.AutoScrollPosition.X, -tab.AutoScrollPosition.Y - 600);
                    }
                }
            }
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        #endregion

        #region 共通コントロール　イベント -------------------------------------------
        private void setCmbEvent()
        {
            bid_tbl02_cmbKinsiUmu.SelectedValueChanged += cmb_SelectedValueChanged;
            bid_tbl02_cmbKinsiNaiyo.SelectedValueChanged += cmb_SelectedValueChanged;
            ca_tbl01_cmbKinsiUmu.SelectedValueChanged += cmb_SelectedValueChanged;
            ca_tbl01_cmbKinsiNaiyo.SelectedValueChanged += cmb_SelectedValueChanged;
        }

        /// <summary>
        /// 各コンボボックスの入力値が基本情報：「９」へ反映する
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmb_SelectedValueChanged(object sender, EventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;

            // 入札：再委託禁止条項の記載有無
            // 契約：再委託禁止条項の記載有無
            reset_base_kinsiItems();
        }

        /// <summary>
        /// 日付KeyDown
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dateTimePicker_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Delete)
            {
                ((DateTimePicker)sender).Text = "";
                ((DateTimePicker)sender).CustomFormat = " ";
            }
        }

        /// <summary>
        /// 日付ValueChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            ((DateTimePicker)sender).CustomFormat = "";
        }

        /// <summary>
        /// 各画面の日付の連動処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dateTimePicker_CloseUp(object sender, EventArgs e)
        {
            DateTimePicker dtp = (DateTimePicker)sender;
            dtp.CustomFormat = "";
            if (dtp.Name.Equals(prior_tbl01_dtpDasinIraiDt.Name))
            {
                //事前打診依頼日
                base_tbl09_dtpJizenDasinIraiDt.Text = dtp.Text;
                return;
            }
            if (dtp.Name.Equals(base_tbl01_dtpDtBid.Name))
            {
                return;
            }
            if (dtp.Name.Equals(prior_tbl01_dtpOrderYoteiDt.Name))
            {
                //発注予定・見込日
                base_tbl09_dtpOrderYoteiDt.Text = dtp.Text;
                return;
            }
            if (dtp.Name.Equals(bid_tbl01_dtpBidYoteiDt.Name))
            {
                //入札：入札予定日⇒基本情報：入札予定日
                base_tbl10_dtpNyusatuDt.Text = dtp.Text;
                return;
            }

            if (dtp.Name.Equals(base_tbl03_dtpKokiTo.Name))
            {
                // 基本情報：工期終了
                if (base_tbl03_dtpKokiFrom.CustomFormat == "" && base_tbl03_dtpKokiTo.CustomFormat == "")
                {
                    if (base_tbl03_dtpKokiFrom.Value > base_tbl03_dtpKokiTo.Value)
                    {
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E10011", "(工期 開始・終了)"));
                        base_tbl03_dtpKokiTo.CustomFormat = " ";
                    }
                }

                if (base_tbl03_dtpKokiTo.CustomFormat == "")
                {
                    if (mode != MODE.INSERT && mode != MODE.PLAN)
                    {
                        ca_tbl01_dtpKokiTo.Text = base_tbl03_dtpKokiTo.Text;
                    }
                    DataTable dt = new DataTable();
                    dt = GlobalMethod.getData("NendoID", "NendoID", "Mst_Nendo", "Nendo_Sdate <= '" + base_tbl03_dtpKokiTo.Text + "' AND Nendo_EDate >= '" + base_tbl03_dtpKokiTo.Text + "' ");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        // 売上年度
                        if (mode != MODE.INSERT && mode != MODE.PLAN)
                        {
                            ca_tbl01_cmbSalesYear.SelectedValue = dt.Rows[0][0].ToString();
                        }
                        if (mode != MODE.CHANGE)
                        {
                            base_tbl03_cmbKokiSalesYear.SelectedValue = dt.Rows[0][0].ToString();
                        }
                    }
                }
                return;
            }

            if (dtp.Name.Equals(base_tbl03_dtpKokiFrom.Name))
            {
                if (base_tbl03_dtpKokiFrom.CustomFormat == "")
                {
                    if (mode != MODE.INSERT && mode != MODE.PLAN)
                    {
                        ca_tbl01_dtpKokiFrom.Text = base_tbl03_dtpKokiFrom.Text;
                    }

                    DataTable dt = new DataTable();
                    dt = GlobalMethod.getData("NendoID", "NendoID", "Mst_Nendo", "Nendo_Sdate <= '" + base_tbl03_dtpKokiFrom.Text + "' AND Nendo_EDate >= '" + base_tbl03_dtpKokiFrom.Text + "' ");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        // 工期開始年度
                        if (mode != MODE.INSERT && mode != MODE.PLAN)
                        {
                            ca_tbl01_cmbStartYear.SelectedValue = dt.Rows[0][0].ToString();
                        }
                        if (mode != MODE.CHANGE)
                        {
                            base_tbl03_cmbKokiStartYear.SelectedValue = dt.Rows[0][0].ToString();
                        }
                    }
                }
                return;
            }

            if (dtp.Name.Equals(ca_tbl01_dtpChangeDt.Name))
            {
                //　契約：契約締結(変更)日
                if (ca_tbl01_dtpChangeDt.CustomFormat != "")
                {
                    // 契約締結変更日を入力しないと消費税率を取得できません。
                    set_error(GlobalMethod.GetMessage("I10713", ""));
                }
                else
                {
                    string where = "(TaxStartDay <= '" + dtp.Text + "' ) AND (ISNULL(TaxEndDay,'9999/12/31') >= '" + dtp.Text + "' ) " +
                                    " AND TaxKuni = 'JPN' AND ISNULL(TaxDeleteFlag,0) = 0 ";
                    DataTable dt = GlobalMethod.getData("TaxZeiritsu", "TaxZeiritsu", "M_Tax", where);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        ca_tbl01_txtTax.Text = dt.Rows[0][0].ToString();
                        // 小数点以下を削り取る
                        int comma = ca_tbl01_txtTax.Text.IndexOf(".");
                        ca_tbl01_txtTax.Text = ca_tbl01_txtTax.Text.Substring(0, comma);
                    }
                    else
                    {
                        where = "(TaxStartDay IS null OR TaxStartDay <= '" + dtp.Text + "' ) AND (TaxEndDay IS null OR TaxEndDay >= '" + dtp.Text + "' ) " +
                                    " AND TaxKuni = 'JPN' AND ISNULL(TaxDeleteFlag,0) = 0 ";
                        dt.Clear();
                        dt = GlobalMethod.getData("TaxZeiritsu", "TaxZeiritsu", "M_Tax", where);

                        if (dt != null && dt.Rows.Count > 0)
                        {
                            ca_tbl01_txtTax.Text = dt.Rows[0][0].ToString();
                        }
                        else
                        {
                            ca_tbl01_txtTax.Text = "0";
                        }
                    }
                }
                return;
            }

            if (dtp.Name.Equals(ca_tbl01_dtpKokiTo.Name))
            {
                // 契約：工期至
                if (ca_tbl01_dtpKokiFrom.CustomFormat == "" && ca_tbl01_dtpKokiTo.CustomFormat == "")
                {
                    // No.204 工期末日付のコピーボタンが反応しない
                    // エラーメッセージが消えることでボタンが押せていないので、エラーでない場合はメッセージを消さないように修正
                    //set_error("", 0);
                    if (ca_tbl01_dtpKokiFrom.Value > ca_tbl01_dtpKokiTo.Value)
                    {
                        set_error("", 0);
                        set_error(GlobalMethod.GetMessage("E10011", "(契約工期 開始・終了)"));
                        ca_tbl01_dtpKokiTo.CustomFormat = " ";
                    }
                }

                if (ca_tbl01_dtpKokiTo.CustomFormat == "")
                {
                    DataTable dt = GlobalMethod.getData("NendoID", "NendoID", "Mst_Nendo", "Nendo_Sdate <= '" + ca_tbl01_dtpKokiTo.Text + "' AND Nendo_EDate >= '" + ca_tbl01_dtpKokiTo.Text + "' ");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        // 売上年度
                        ca_tbl01_cmbSalesYear.SelectedValue = dt.Rows[0][0].ToString();
                        // 変更伝票以外の場合、引合タブの売上年度も更新する
                        if (mode != MODE.CHANGE)
                        {
                            ca_tbl01_cmbSalesYear.SelectedValue = dt.Rows[0][0].ToString();
                        }
                    }
                }
                return;
            }

            if (dtp.Name.Equals(ca_tbl01_dtpKokiFrom.Name))
            {
                // 契約：工期自
                if (ca_tbl01_dtpKokiFrom.CustomFormat == "")
                {
                    DataTable dt = new DataTable();
                    dt = GlobalMethod.getData("NendoID", "NendoID", "Mst_Nendo", "Nendo_Sdate <= '" + ca_tbl01_dtpKokiFrom.Text + "' AND Nendo_EDate >= '" + ca_tbl01_dtpKokiFrom.Text + "' ");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        // 工期開始年度
                        ca_tbl01_cmbStartYear.SelectedValue = dt.Rows[0][0].ToString();
                        if (mode != MODE.CHANGE)
                        {
                            base_tbl03_cmbKokiStartYear.SelectedValue = dt.Rows[0][0].ToString();
                        }
                    }
                }
                return;
            }
        }

        /// <summary>
        /// フォーカスで、数字のフォマードがなくなる処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// 金額のフォマードをする
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox_Validated(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            ((System.Windows.Forms.TextBox)sender).Text = GetMoneyTextLong(GetLong(tmp));
        }

        /// <summary>
        /// パセートのフォマードをする
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox_ValidatedPercent(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            ((System.Windows.Forms.TextBox)sender).Text = GetPercentText(GetDouble(tmp));
        }


        /// <summary>
        /// テキストボックス：数字のみ入力制御
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            System.Windows.Forms.TextBox txt = (System.Windows.Forms.TextBox)sender;
            if(txt.Tag != null && txt.Tag.ToString().Equals("9"))
            {
                // 数字のみ入力
                if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
                {
                    e.Handled = true;
                }
            }
            else
            {
                if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '-')
                {
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// テキストボックス：配分率（％）入力制御
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textbox_KeyPressPercent(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// コンボボックス　マウスホイールイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmb_MouseWheel(object sender, EventArgs e)
        {
            HandledMouseEventArgs wEventArgs = e as HandledMouseEventArgs;
            wEventArgs.Handled = true;
        }

        /// <summary>
        /// 数値項目のコピーとペースト
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// コンボボックス　リストの再描画
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            e.DrawBackground();

            bool selected = DrawItemState.Selected == (e.State & DrawItemState.Selected);
            var brush = (selected) ? Brushes.White : Brushes.Black;
            DataRowView r = ((ComboBox)sender).Items[e.Index] as DataRowView;
            if (r != null)
            {
                e.Graphics.DrawString(r.Row["Discript"].ToString(), e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            }
            e.DrawFocusRectangle();
        }

        /// <summary>
        /// 配分率（%）入力完了後の合計処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void numPercentTextBox_Leave(object sender, EventArgs e)
        {
            //// 新規登録の時連動しない
            //if (mode == MODE.INSERT || mode == MODE.PLAN) return
            System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)sender;
            tb.Text = GetPercentText(GetDouble(tb.Text));
            // 編集不可なら、何もしない
            if (tb.ReadOnly) return;

            string sFiexd = "numPercent";    // パセート専用コントロール名
            string sName = tb.Name;
            string[] words = sName.Replace(sFiexd, ",").Split(',');
            string sNo = words[1];  // 入力配分率の連番
            if (tb.Name.Contains("ca_tbl02_AftCaTs_numPercent"))
            {
                // ２．配分情報・業務内容 ⇒ 調査部　業務別配分⇒【契約後】配分率(%)
                //　合計
                GetTotalPercent("ca_tbl02_AftCaTs_numPercent", 13);

                System.Windows.Forms.Label lblAmt = ca_tbl02_AftCaTs_numAmt1;
                //基本情報一覧　７　調査部　業務配分　契約後
                System.Windows.Forms.Label lblBaseP = base_tbl07_5_lblRate1;
                System.Windows.Forms.Label lblBaseA = base_tbl07_5_lblAmt1;
                switch (sNo)
                {
                    case "1":
                        break;
                    case "2":
                        lblAmt = ca_tbl02_AftCaTs_numAmt2;
                        lblBaseP = base_tbl07_5_lblRate2;
                        lblBaseA = base_tbl07_5_lblAmt2;
                        break;
                    case "3":
                        lblAmt = ca_tbl02_AftCaTs_numAmt3;
                        lblBaseP = base_tbl07_5_lblRate3;
                        lblBaseA = base_tbl07_5_lblAmt3;
                        break;
                    case "4":
                        lblAmt = ca_tbl02_AftCaTs_numAmt4;
                        lblBaseP = base_tbl07_5_lblRate4;
                        lblBaseA = base_tbl07_5_lblAmt4;
                        break;
                    case "5":
                        lblAmt = ca_tbl02_AftCaTs_numAmt5;
                        lblBaseP = base_tbl07_5_lblRate5;
                        lblBaseA = base_tbl07_5_lblAmt5;
                        break;
                    case "6":
                        lblAmt = ca_tbl02_AftCaTs_numAmt6;
                        lblBaseP = base_tbl07_5_lblRate6;
                        lblBaseA = base_tbl07_5_lblAmt6;
                        break;
                    case "7":
                        lblAmt = ca_tbl02_AftCaTs_numAmt7;
                        lblBaseP = base_tbl07_5_lblRate7;
                        lblBaseA = base_tbl07_5_lblAmt7;
                        break;
                    case "8":
                        lblAmt = ca_tbl02_AftCaTs_numAmt8;
                        lblBaseP = base_tbl07_5_lblRate8;
                        lblBaseA = base_tbl07_5_lblAmt8;
                        break;
                    case "9":
                        lblAmt = ca_tbl02_AftCaTs_numAmt9;
                        lblBaseP = base_tbl07_5_lblRate9;
                        lblBaseA = base_tbl07_5_lblAmt9;
                        break;
                    case "10":
                        lblAmt = ca_tbl02_AftCaTs_numAmt10;
                        lblBaseP = base_tbl07_5_lblRate10;
                        lblBaseA = base_tbl07_5_lblAmt10;
                        break;
                    case "11":
                        lblAmt = ca_tbl02_AftCaTs_numAmt11;
                        lblBaseP = base_tbl07_5_lblRate11;
                        lblBaseA = base_tbl07_5_lblAmt11;
                        break;
                    case "12":
                        lblAmt = ca_tbl02_AftCaTs_numAmt12;
                        lblBaseP = base_tbl07_5_lblRate12;
                        lblBaseA = base_tbl07_5_lblAmt12;
                        break;
                }
                // 調査部 業務配分別配分 契約 配分額(税抜)
                double percent = GetDouble(tb.Text);
                long haibun = 0;
                // 契約 配分額(税抜)
                long total = GetLong(ca_tbl02_AftCaBm_numAmt1.Text);
                if (total * percent != 0)
                {
                    haibun = (long)Math.Round(total * percent / 100);
                }
                lblAmt.Text = GetMoneyTextLong(haibun);
                GetTotalMoney("ca_tbl02_AftCaTs_numAmt", 13);

                // 基本情報等一覧へ連動
                lblBaseP.Text = tb.Text;
                lblBaseA.Text = lblAmt.Text;
                base_tbl07_5_lblRateAll.Text = ca_tbl02_AftCaTs_numPercentAll.Text;
                base_tbl07_5_lblAmtAll.Text = ca_tbl02_AftCaTs_numAmtAll.Text;
                return;
            }

            if (tb.Name.Contains("ca_tbl02_AftCaBm_numPercent"))
            {
                GetTotalPercent("ca_tbl02_AftCaBm_numPercent", 5);
                System.Windows.Forms.Label lblBaseP = base_tbl07_4_lblRate1;
                switch (sNo)
                {
                    case "1":
                        break;
                    case "2":
                        lblBaseP = base_tbl07_4_lblRate2;
                        break;
                    case "3":
                        lblBaseP = base_tbl07_4_lblRate3;
                        break;
                    case "4":
                        lblBaseP = base_tbl07_4_lblRate4;
                        break;
                }
                // 基本情報等一覧へ連動
                lblBaseP.Text = tb.Text;
                base_tbl07_4_lblRateAll.Text = ca_tbl02_AftCaBm_numPercentAll.Text;

                // 金額の連動計算 No.1467
                cal_aftCaBmAmt(sNo);
                return;
            }
            string sParentName = words[0] + sFiexd;  // 入力配分率の親情報部分のName
            int num = 0;
            string sGearing = "";   //連動コントロール
            switch (sParentName)
            {
                case "base_tbl07_1_numPercent":
                    num = 5;    // 基本情報：７．業務配分 部門配分
                    sGearing = "ca_tbl02_1_numPercent"; // 契約：２．配分情報・業務内容 部門配分へ連動
                    break;
                case "base_tbl07_2_numPercent":
                    num = 13;    // 基本情報：７．業務配分 調査部　業務別配分
                    sGearing = "ca_tbl02_2_numPercent"; // 契約：２．配分情報・業務内容 調査部　業務別配分へ連動
                    break;
            }
            if (num > 0)
            {
                GetTotalPercent(sParentName, num);
                if (mode != MODE.INSERT && mode != MODE.PLAN)
                {
                    // 連動コントロールの合計も計算する
                    GetTotalPercent(sGearing, num);

                    // 連動コントロールへデータを設定する
                    SetGearingPercent(sGearing + sNo, tb.Text);
                }
            }
        }

        /// <summary>
        /// 金額テキストボックスLeaveイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void numAmtTextBox_Leave(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)sender;
            string sName = tb.Name;
            if (ca_tbl01_txtTax.Name.Equals(sName) == false) tb.Text = GetMoneyTextLong(GetLong(tb.Text));
            if (sName.Contains("ca_tbl07_txtRequst"))
            {
                // 契約：７．請求書情報：請求金額
                GetTotalMoney("ca_tbl07_txtRequst", 7);
                return;
            }
            if (sName.Contains("ca_tbl03_numAmt"))
            {
                // 契約：３．単契等の見込補正額(年度内)
                GetTotalMoney("ca_tbl03_numAmt", 5);
                // 基本情報１１へ反映する
                switch (sName.Replace("ca_tbl03_numAmt", ""))
                {
                    case "1":
                        base_tbl11_2_numAmt1.Text = tb.Text;
                        // 調査部　業務配分再計算
                        calc_aftCaTsFreeTax();
                        break;
                    case "2":
                        base_tbl11_2_numAmt2.Text = tb.Text;
                        break;
                    case "3":
                        base_tbl11_2_numAmt3.Text = tb.Text;
                        break;
                    case "4":
                        base_tbl11_2_numAmt4.Text = tb.Text;
                        break;
                }
                base_tbl11_2_numAmtAll.Text = ca_tbl03_numAmtAll.Text;
                return;
            }
            if (sName.Contains("ca_tbl04_numKurikosiAmt"))
            {
                // 契約：４．年度繰越額(年度跨ぎ)
                GetTotalMoney("ca_tbl04_numKurikosiAmt", 5);

                // 基本情報１１へ反映する
                switch (sName.Replace("ca_tbl04_numKurikosiAmt", ""))
                {
                    case "1":
                        base_tbl11_3_numAmt1.Text = tb.Text;
                        // 調査部　業務配分再計算
                        calc_aftCaTsFreeTax();
                        break;
                    case "2":
                        base_tbl11_3_numAmt2.Text = tb.Text;
                        break;
                    case "3":
                        base_tbl11_3_numAmt3.Text = tb.Text;
                        break;
                    case "4":
                        base_tbl11_3_numAmt4.Text = tb.Text;
                        break;
                }
                base_tbl11_3_numAmtAll.Text = ca_tbl04_numKurikosiAmtAll.Text;
                return;
            }
            if (sName.Contains("ca_tbl02_AftCaBmZeikomi_numAmt"))
            {
                // 配分額(税込) 合計
                GetTotalMoney("ca_tbl02_AftCaBmZeikomi_numAmt", 5);
                //税抜算出
                long num = GetLong(tb.Text);
                string sNo = sName.Replace("ca_tbl02_AftCaBmZeikomi_numAmt", "");
                switch (sNo)
                {
                    case "1":
                        ca_tbl02_AftCaBm_numAmt1.Text = GetMoneyTextLong(Get_Zeinuki(num));
                        // 調査部 業務配分別配分 契約 配分額(税抜)
                        calc_aftCaTsFreeTax();
                        //基本情報等一覧へも反映する
                        base_tbl07_4_lblAmt1.Text = ca_tbl02_AftCaBm_numAmt1.Text;
                        break;
                    case "2":
                        ca_tbl02_AftCaBm_numAmt2.Text = GetMoneyTextLong(Get_Zeinuki(num));
                        base_tbl07_4_lblAmt2.Text = ca_tbl02_AftCaBm_numAmt2.Text;
                        break;
                    case "3":
                        ca_tbl02_AftCaBm_numAmt3.Text = GetMoneyTextLong(Get_Zeinuki(num));
                        base_tbl07_4_lblAmt3.Text = ca_tbl02_AftCaBm_numAmt3.Text;
                        break;
                    case "4":
                        ca_tbl02_AftCaBm_numAmt4.Text = GetMoneyTextLong(Get_Zeinuki(num));
                        base_tbl07_4_lblAmt4.Text = ca_tbl02_AftCaBm_numAmt4.Text;
                        break;
                }
                // 配分額(税抜) 合計
                GetTotalMoney("ca_tbl02_AftCaBm_numAmt", 5);
                base_tbl07_4_lblAmtAll.Text = ca_tbl02_AftCaBm_numAmtAll.Text;
                // 配分率の計算 No.1467
                cal_aftCaBmPercent(sNo); 
                return;
            }
            // No1604 追加対応（配分額(税抜)を入力項目に変更）
            if (sName.Contains("ca_tbl02_AftCaBm_numAmt"))
            {
                // 配分額(税抜)合計
                GetTotalMoney("ca_tbl02_AftCaBm_numAmt", 5);
                string sNo = sName.Replace("ca_tbl02_AftCaBm_numAmt", "");
                switch (sNo)
                {
                    case "1":
                        // 調査部 業務配分別配分 契約 配分額(税抜)
                        calc_aftCaTsFreeTax();
                        //基本情報等一覧へも反映する
                        base_tbl07_4_lblAmt1.Text = ca_tbl02_AftCaBm_numAmt1.Text;
                        break;
                    case "2":
                        base_tbl07_4_lblAmt2.Text = ca_tbl02_AftCaBm_numAmt2.Text;
                        break;
                    case "3":
                        base_tbl07_4_lblAmt3.Text = ca_tbl02_AftCaBm_numAmt3.Text;
                        break;
                    case "4":
                        base_tbl07_4_lblAmt4.Text = ca_tbl02_AftCaBm_numAmt4.Text;
                        break;
                }
                //基本情報等一覧の配分額(税抜)合計
                base_tbl07_4_lblAmtAll.Text = ca_tbl02_AftCaBm_numAmtAll.Text;
                return;
            }

            if (ca_tbl01_txtZeinukiAmt.Name.Equals(sName))
            {
                // 契約：税抜(自動計算用)
                calc_kingaku();
                return;
            }
            if (ca_tbl01_txtTax.Name.Equals(sName))
            {
                // 契約：消費税率
                calc_kingaku();

                // 契約：２．配分情報・業務内容　部門配分　配分額(税込) 再計算
                ca_tbl02_AftCaBmZeikomi_numAmt1.Text = GetMoneyTextLong(Get_Zeikomi(GetLong(ca_tbl02_AftCaBm_numAmt1.Text)));
                ca_tbl02_AftCaBmZeikomi_numAmt2.Text = GetMoneyTextLong(Get_Zeikomi(GetLong(ca_tbl02_AftCaBm_numAmt2.Text)));
                ca_tbl02_AftCaBmZeikomi_numAmt3.Text = GetMoneyTextLong(Get_Zeikomi(GetLong(ca_tbl02_AftCaBm_numAmt3.Text)));
                ca_tbl02_AftCaBmZeikomi_numAmt4.Text = GetMoneyTextLong(Get_Zeikomi(GetLong(ca_tbl02_AftCaBm_numAmt4.Text)));
                // 合計
                GetTotalMoney("ca_tbl02_AftCaBmZeikomi_numAmt", 5);
                return;
            }
            if (ca_tbl01_txtJyutakuGaiAmt.Name.Equals(sName))
            {
                // 契約：受託外金額(税込)
                calc_jyutakuAmt();
                return;
            }

            if (ca_tbl01_numChangeAmt.Name.Equals(sName))
            {
                // 変更伝票　の　契約金額変更（税抜）
                // 契約金額（税抜）
                long zeinuki = GetLong(ca_tbl01_txtZeinukiAmt.Text) + GetLong(tb.Text);
                ca_tbl01_txtZeinukiAmt.Text = string.Format("{0:C}", zeinuki);
                // 契約：税抜(自動計算用)
                calc_kingaku();
                return;
            }
            
            // 連動
            switch (sName)
            {
                // 事前打診タブ⇒９．事前打診・参考見積
                case "prior_tbl01_txtMitumoriAmt"://参考見積額(税抜)
                    base_tbl09_numSankomitumoriAmt.Text = tb.Text;
                    break;

                // 入札タブ⇒１０．入札情報・入札結果（基本情報）
                case "bid_tbl01_txtMitumoriAmt"://参考見積額(税抜)
                    base_tbl10_numSankoMitumoriAmt.Text = tb.Text;
                    break;
                case "bid_tbl03_1_txtYoteiPrice"://予定価格(税抜)
                    base_tbl10_txtYoteiAmt.Text = tb.Text;
                    break;
                case "bid_tbl03_1_numRakusatuAmt"://落札額(税抜)
                    base_tbl10_txtRakusatuAmt.Text = tb.Text;
                    break;
            }
        }

        /// <summary>
        /// 文字列テキストボックスLeaveイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtTextBox_Leave(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox tb = (System.Windows.Forms.TextBox)sender;
            string sName = tb.Name;
            if (base_tbl02_txtAnkenFolder.Name.Equals(sName))
            {
                FolderPathCheck();
                return;
            }
            if(sName.Equals(bid_tbl02_txtOtherNaiyo.Name) || sName.Equals(ca_tbl01_txtOtherNaiyo.Name))
            {
                reset_base_kinsiItems();
                return;
            }
            // 連動
            switch (sName)
            {
                // 事前打診タブ⇒９．事前打診・参考見積
                case "prior_tbl02_txtOtherNaiyo"://「その他」の内容
                    base_tbl09_txtOthenComment.Text = tb.Text;
                    break;
                case "bid_tbl03_1_txtOsatuNum"://応札数
                    base_tbl10_txtOsatuNum.Text = tb.Text;
                    break;
                case "bid_tbl03_1_txtRakusatuSya"://落札者
                    base_tbl10_txtRakusatuSya.Text = tb.Text;
                    break;
            }
        }

        /// <summary>
        /// 基本情報等一覧：再委託禁止条項の連動設定
        /// </summary>
        private void reset_base_kinsiItems()
        {
            if (string.IsNullOrEmpty(ca_tbl01_cmbKinsiUmu.Text))
            {
                base_tbl10_cmbKinsiUmu.SelectedValue = bid_tbl02_cmbKinsiUmu.SelectedValue;
                base_tbl10_cmbKinsiNaiyo.SelectedValue = bid_tbl02_cmbKinsiNaiyo.SelectedValue;
                base_tbl10_txtOtherNaiyo.Text = bid_tbl02_txtOtherNaiyo.Text;
            }
            else
            {
                base_tbl10_cmbKinsiUmu.SelectedValue = ca_tbl01_cmbKinsiUmu.SelectedValue;
                base_tbl10_cmbKinsiNaiyo.SelectedValue = ca_tbl01_cmbKinsiNaiyo.SelectedValue;
                base_tbl10_txtOtherNaiyo.Text = ca_tbl01_txtOtherNaiyo.Text;
            }
        }

        /// <summary>
        /// 調査部 業務配分別配分 契約 配分額(税抜) 自動計算
        /// </summary>
        private void calc_aftCaTsFreeTax()
        {
            //配分情報・業務内容　「調査部　業務別配分額(税抜)」の計算式 ③＝（①＋④＋⑤）×②　
            //①：配分部門　調査部の配分額（税抜き）
            //④：３．単契等の見込補正額(年度内) の　調査部部門配分額(税抜)
            //⑤：４．年度繰越額(年度跨ぎ)　の　調査部部門配分額(税抜)
            long total = GetLong(ca_tbl02_AftCaBm_numAmt1.Text) + GetLong(ca_tbl03_numAmt1.Text) + GetLong(ca_tbl04_numKurikosiAmt1.Text); ;
            double percent1 = GetDouble(ca_tbl02_AftCaTs_numPercent1.Text);
            double percent2 = GetDouble(ca_tbl02_AftCaTs_numPercent2.Text);
            double percent3 = GetDouble(ca_tbl02_AftCaTs_numPercent3.Text);
            double percent4 = GetDouble(ca_tbl02_AftCaTs_numPercent4.Text);
            double percent5 = GetDouble(ca_tbl02_AftCaTs_numPercent5.Text);
            double percent6 = GetDouble(ca_tbl02_AftCaTs_numPercent6.Text);
            double percent7 = GetDouble(ca_tbl02_AftCaTs_numPercent7.Text);
            double percent8 = GetDouble(ca_tbl02_AftCaTs_numPercent8.Text);
            double percent9 = GetDouble(ca_tbl02_AftCaTs_numPercent9.Text);
            double percent10 = GetDouble(ca_tbl02_AftCaTs_numPercent10.Text);
            double percent11 = GetDouble(ca_tbl02_AftCaTs_numPercent11.Text);
            double percent12 = GetDouble(ca_tbl02_AftCaTs_numPercent12.Text);
            ca_tbl02_AftCaTs_numAmt1.Text = GetMoneyTextLong((long)Math.Round(total * percent1 / 100));
            ca_tbl02_AftCaTs_numAmt2.Text = GetMoneyTextLong((long)Math.Round(total * percent2 / 100));
            ca_tbl02_AftCaTs_numAmt3.Text = GetMoneyTextLong((long)Math.Round(total * percent3 / 100));
            ca_tbl02_AftCaTs_numAmt4.Text = GetMoneyTextLong((long)Math.Round(total * percent4 / 100));
            ca_tbl02_AftCaTs_numAmt5.Text = GetMoneyTextLong((long)Math.Round(total * percent5 / 100));
            ca_tbl02_AftCaTs_numAmt6.Text = GetMoneyTextLong((long)Math.Round(total * percent6 / 100));
            ca_tbl02_AftCaTs_numAmt7.Text = GetMoneyTextLong((long)Math.Round(total * percent7 / 100));
            ca_tbl02_AftCaTs_numAmt8.Text = GetMoneyTextLong((long)Math.Round(total * percent8 / 100));
            ca_tbl02_AftCaTs_numAmt9.Text = GetMoneyTextLong((long)Math.Round(total * percent9 / 100));
            ca_tbl02_AftCaTs_numAmt10.Text = GetMoneyTextLong((long)Math.Round(total * percent10 / 100));
            ca_tbl02_AftCaTs_numAmt11.Text = GetMoneyTextLong((long)Math.Round(total * percent11 / 100));
            ca_tbl02_AftCaTs_numAmt12.Text = GetMoneyTextLong((long)Math.Round(total * percent12 / 100));
            // 契約 配分額(税抜)
            GetTotalMoney("ca_tbl02_AftCaTs_numAmt", 13);

            // 基本情報等一覧へ連動する
            if(mode != MODE.CHANGE)
            {
                base_tbl07_5_lblAmt1.Text = ca_tbl02_AftCaTs_numAmt1.Text;
                base_tbl07_5_lblAmt2.Text = ca_tbl02_AftCaTs_numAmt2.Text;
                base_tbl07_5_lblAmt3.Text = ca_tbl02_AftCaTs_numAmt3.Text;
                base_tbl07_5_lblAmt4.Text = ca_tbl02_AftCaTs_numAmt4.Text;
                base_tbl07_5_lblAmt5.Text = ca_tbl02_AftCaTs_numAmt5.Text;
                base_tbl07_5_lblAmt6.Text = ca_tbl02_AftCaTs_numAmt6.Text;
                base_tbl07_5_lblAmt7.Text = ca_tbl02_AftCaTs_numAmt7.Text;
                base_tbl07_5_lblAmt8.Text = ca_tbl02_AftCaTs_numAmt8.Text;
                base_tbl07_5_lblAmt9.Text = ca_tbl02_AftCaTs_numAmt9.Text;
                base_tbl07_5_lblAmt10.Text = ca_tbl02_AftCaTs_numAmt10.Text;
                base_tbl07_5_lblAmt11.Text = ca_tbl02_AftCaTs_numAmt11.Text;
                base_tbl07_5_lblAmt12.Text = ca_tbl02_AftCaTs_numAmt12.Text;
                base_tbl07_5_lblAmtAll.Text = ca_tbl02_AftCaTs_numAmtAll.Text;
            }
        }

        /// <summary>
        /// 契約後配分部門　配分率より、金額計算
        /// </summary>
        private void cal_aftCaBmAmt(string sPercentNo)
        {
            double dPercent = 0;
            long total = GetLong(ca_tbl01_txtJyutakuAmt.Text);
            System.Windows.Forms.TextBox txtAftCaBmAmtTax = ca_tbl02_AftCaBmZeikomi_numAmt1;    // 税込み
            System.Windows.Forms.TextBox txtAftCaBmAmt = ca_tbl02_AftCaBm_numAmt1;              // 税抜き
            System.Windows.Forms.Label lblAftCaBmAmt = base_tbl07_4_lblAmt1;
            switch (sPercentNo)
            {
                case "1":
                    dPercent = GetDouble(ca_tbl02_AftCaBm_numPercent1.Text);
                    break;
                case "2":
                    dPercent = GetDouble(ca_tbl02_AftCaBm_numPercent2.Text);
                    txtAftCaBmAmtTax = ca_tbl02_AftCaBmZeikomi_numAmt2;
                    txtAftCaBmAmt = ca_tbl02_AftCaBm_numAmt2;
                    lblAftCaBmAmt = base_tbl07_4_lblAmt2;
                    break;
                case "3":
                    dPercent = GetDouble(ca_tbl02_AftCaBm_numPercent3.Text);
                    txtAftCaBmAmtTax = ca_tbl02_AftCaBmZeikomi_numAmt3;
                    txtAftCaBmAmt = ca_tbl02_AftCaBm_numAmt3;
                    lblAftCaBmAmt = base_tbl07_4_lblAmt3;
                    break;
                case "4":
                    dPercent = GetDouble(ca_tbl02_AftCaBm_numPercent4.Text);
                    txtAftCaBmAmtTax = ca_tbl02_AftCaBmZeikomi_numAmt4;
                    txtAftCaBmAmt = ca_tbl02_AftCaBm_numAmt4;
                    lblAftCaBmAmt = base_tbl07_4_lblAmt4;
                    break;
            }
            // 税込み金額と合計再計算
            long amt = (long)Math.Round(total * dPercent / 100);
            txtAftCaBmAmtTax.Text = GetMoneyTextLong(amt);
            GetTotalMoney("ca_tbl02_AftCaBmZeikomi_numAmt", 5);
            // 税抜き金額と合計再計算
            txtAftCaBmAmt.Text = GetMoneyTextLong(Get_Zeinuki(amt));
            GetTotalMoney("ca_tbl02_AftCaBm_numAmt", 5);

            // 基本情報等一覧へ連動
            if (mode != MODE.CHANGE)
            {
                lblAftCaBmAmt.Text = txtAftCaBmAmt.Text;
                base_tbl07_4_lblAmtAll.Text = ca_tbl02_AftCaBm_numAmtAll.Text;
            }

            // 調査部の場合、業務配分も再計算
            if (sPercentNo.Equals("1"))
            {
                calc_aftCaTsFreeTax();
            }
        }

        /// <summary>
        /// 契約後配分部門　金額税込より、配分率計算
        /// </summary>
        private void cal_aftCaBmPercent(string sPercentNo)
        {
            long lngAmtTax = 0;
            long total = GetLong(ca_tbl01_txtJyutakuAmt.Text);
            System.Windows.Forms.TextBox txtAftCaBmPercent = ca_tbl02_AftCaBm_numPercent1;    // 配分率
            System.Windows.Forms.Label lblAftCaBmRate = base_tbl07_4_lblRate1;
            switch (sPercentNo)
            {
                case "1":
                    lngAmtTax = GetLong(ca_tbl02_AftCaBmZeikomi_numAmt1.Text);
                    break;
                case "2":
                    lngAmtTax = GetLong(ca_tbl02_AftCaBmZeikomi_numAmt2.Text);
                    txtAftCaBmPercent = ca_tbl02_AftCaBm_numPercent2;
                    lblAftCaBmRate = base_tbl07_4_lblRate2;
                    break;
                case "3":
                    lngAmtTax = GetLong(ca_tbl02_AftCaBmZeikomi_numAmt3.Text);
                    txtAftCaBmPercent = ca_tbl02_AftCaBm_numPercent3;
                    lblAftCaBmRate = base_tbl07_4_lblRate3;
                    break;
                case "4":
                    lngAmtTax = GetLong(ca_tbl02_AftCaBmZeikomi_numAmt4.Text);
                    txtAftCaBmPercent = ca_tbl02_AftCaBm_numPercent4;
                    lblAftCaBmRate = base_tbl07_4_lblRate4;
                    break;
            }
            // パーセント再計算
            double dPercent = 0;
            if (total > 0)
            {
                //dPercent = (double)(lngAmtTax * 100 / total);
                dPercent = lngAmtTax * 100 / Convert.ToDouble(total);
            }
            txtAftCaBmPercent.Text = GetPercentText(dPercent);
            GetTotalPercent("ca_tbl02_AftCaBm_numPercent", 5);

            // 基本情報等一覧へ連動
            if (mode != MODE.CHANGE)
            {
                lblAftCaBmRate.Text = txtAftCaBmPercent.Text;
                base_tbl07_4_lblRateAll.Text = ca_tbl02_AftCaBm_numPercentAll.Text;
            }
        }
        #endregion

        #region Grid イベント --------------------------------------------------------
        /// <summary>
        /// 入札：３．入札結果　の　入札参加者リスト
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void bidc1FlexGrid_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            C1FlexGrid c1FlexGrid2 = bid_tbl03_4_c1FlexGrid;
            var hti = c1FlexGrid2.HitTest(new System.Drawing.Point(e.X, e.Y));
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
                        bid_tbl03_1_txtRakusatuSya.Text = c1FlexGrid2.Rows[_row][5].ToString();
                        bid_tbl03_1_numRakusatuAmt.Text = GetMoneyTextLong(GetLong(c1FlexGrid2.Rows[_row][6].ToString()));
                    }
                    int nyusatsuCnt = 0;
                    for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
                    {
                        if (c1FlexGrid2.Rows[i][5] != null && c1FlexGrid2.Rows[i][5].ToString() != "")
                        {
                            nyusatsuCnt++;
                        }
                    }
                    bid_tbl03_1_txtOsatuNum.Text = nyusatsuCnt.ToString();
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
                        bid_tbl03_1_txtRakusatuSya.Text = "";
                        bid_tbl03_1_numRakusatuAmt.Text = GetMoneyTextLong(0);
                    }

                    c1FlexGrid2.RemoveItem(_row);
                    Resize_Grid(c1FlexGrid2.Name);


                    int nyusatsuCnt = 0;
                    for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
                    {
                        if (c1FlexGrid2.Rows[i][5] != null && c1FlexGrid2.Rows[i][5].ToString() != "")
                        {
                            nyusatsuCnt++;
                        }
                    }
                    bid_tbl03_1_txtOsatuNum.Text = nyusatsuCnt.ToString();
                }
            }
        }

        /// <summary>
        /// 基本情報：８．過去案件情報
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void basec1FlexGrid_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            var hti = this.base_tbl08_c1FlexGrid.HitTest(new System.Drawing.Point(e.X, e.Y));

            //if (hti.Column == 3 & hti.Row != 0)
            if (hti.Column == 3 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                Popup_Anken form = new Popup_Anken();
                form.mode = "";
                int nendo = DateTime.Today.Year;
                if (int.TryParse(base_tbl03_cmbKokiStartYear.SelectedValue.ToString(), out nendo))
                {
                    nendo--;
                }
                form.nendo = nendo.ToString();
                form.hachuushaKaMei = base_tbl04_txtOrderName.Text.Trim() + "　" + base_tbl04_txtOrderKamei.Text.Trim();
                form.gyoumuMei = base_tbl03_txtGyomuName.Text.Trim();
                form.gyoumuBushoCD = UserInfos[2];
                form.ShowDialog();
                if (form.ReturnValue != null && form.ReturnValue[0] != null)
                {
                    base_tbl08_c1FlexGrid.Rows[_row][2] = form.ReturnValue[0];   // AnkenJouhou.AnkenJouhouID
                    base_tbl08_c1FlexGrid.Rows[_row][3] = form.ReturnValue[1];   // AnkenAnkenBangou
                    base_tbl08_c1FlexGrid.Rows[_row][4] = form.ReturnValue[2];   // AnkenJutakuBangouALL
                    base_tbl08_c1FlexGrid.Rows[_row][5] = form.ReturnValue[3];   // AnkenJutakuBangouEda
                    base_tbl08_c1FlexGrid.Rows[_row][6] = form.ReturnValue[4];   // AnkenGyoumuMei
                    base_tbl08_c1FlexGrid.Rows[_row][7] = form.ReturnValue[5];   // NyuusatsuRakusatsusha
                    base_tbl08_c1FlexGrid.Rows[_row][8] = form.ReturnValue[6];   // NyuusatsuRakusatsushaID
                    base_tbl08_c1FlexGrid.Rows[_row][9] = form.ReturnValue[7];   // NyuusatsuRakusatugaku
                    base_tbl08_c1FlexGrid.Rows[_row][10] = form.ReturnValue[8];  // NyuusatsuOusatugaku
                    base_tbl08_c1FlexGrid.Rows[_row][11] = form.ReturnValue[9];  // NyuusatsuMitsumorigaku
                    base_tbl08_c1FlexGrid.Rows[_row][12] = form.ReturnValue[10]; // KeiyakuZeikomiKingaku
                    base_tbl08_c1FlexGrid.Rows[_row][13] = form.ReturnValue[11]; // Keiyakukeiyakukingakukei // 前回受託金額（税抜）
                    base_tbl08_c1FlexGrid.Rows[_row][14] = form.ReturnValue[12]; // NyuusatsuKyougouTashaID
                    base_tbl08_c1FlexGrid.Rows[_row][15] = form.ReturnValue[13]; // KyougouKigyouCD
                }
            }
            if (hti.Column == 1 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                if (MessageBox.Show("行を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    base_tbl08_c1FlexGrid.RemoveItem(_row);
                    Resize_Grid("base_tbl08_c1FlexGrid");
                }
            }
        }

        /// <summary>
        /// 契約：担当者リスト
        /// 技術者評価：担当技術者
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c1FlexGrid_BeforeMouseDown(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
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
                    ca_tbl05_txtTanto_c1FlexGrid.Rows[_row][_col - 1] = form.ReturnValue[0];
                    ca_tbl05_txtTanto_c1FlexGrid.Rows[_row][_col] = form.ReturnValue[1];
                    te_c1FlexGrid.Rows[_row][_col - 1] = form.ReturnValue[0];
                    te_c1FlexGrid.Rows[_row][_col] = form.ReturnValue[1];
                }
            }
            if (hti.Column == 0 & hti.Row > 0)
            {
                var _row = hti.Row;
                var _col = hti.Column;

                if (MessageBox.Show("行を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    ca_tbl05_txtTanto_c1FlexGrid.RemoveItem(_row);
                    te_c1FlexGrid.RemoveItem(_row);
                    Resize_Grid("ca_tbl05_txtTanto_c1FlexGrid");
                    Resize_Grid("te_c1FlexGrid");
                }
            }
        }

        /// <summary>
        /// 技術者評価：担当技術者(te_c1FlexGrid)
        /// 契約：売上計上情報リスト(ca_tbl06_c1FlexGrid)
        /// 入札: 入札参加者リスト(bid_tbl03_4_c1FlexGrid)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c1FlexGrid_BeforeEdit(object sender, RowColEventArgs e)
        {
            C1FlexGrid grid = (C1FlexGrid)sender;
            if (grid.Name == "te_c1FlexGrid")
            {
                switch (e.Col)
                {
                    // 技術者評価：担当技術者 評点
                    case 3:
                        grid.ImeMode = ImeMode.Disable;
                        break;
                    default:
                        grid.ImeMode = ImeMode.Off;
                        break;
                }
            }else if(grid.Name == "ca_tbl06_c1FlexGrid")
            {
                // 契約：売上計上情報リスト
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
                        grid.ImeMode = ImeMode.Disable;
                        break;
                    default:
                        grid.ImeMode = ImeMode.Off;
                        break;
                }
            }
            else if (grid.Name == "bid_tbl03_4_c1FlexGrid")
            {
                switch (e.Col)
                {
                    //入札: 入札参加者リスト
                    // 応札額（税抜）
                    case 6:
                        grid.ImeMode = ImeMode.Disable;
                        break;
                    default:
                        grid.ImeMode = ImeMode.Off;
                        break;
                }
            }

        }

        /// <summary>
        /// 入札: 入札参加者リスト(bid_tbl03_4_c1FlexGrid)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c1FlexGrid_AfterEdit(object sender, RowColEventArgs e)
        {
            C1FlexGrid grid = (C1FlexGrid)sender;
            if (grid.Name == "bid_tbl03_4_c1FlexGrid")
            {
                if (grid.GetCellCheck(e.Row, 3) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                {
                    // c1FlexGrid2.Rows[e.Row][5]がNullの場合に、エラーになるので回避
                    if (grid.Rows[e.Row][5] != null)
                    {
                        bid_tbl03_1_txtRakusatuSya.Text = grid.Rows[e.Row][5].ToString();
                        if (grid.Rows[e.Row][6] == null || grid.Rows[e.Row][6].ToString() == "")
                        {
                            grid.Rows[e.Row][6] = 0;
                        }
                        bid_tbl03_1_numRakusatuAmt.Text = GetMoneyTextLong(GetLong(grid.Rows[e.Row][6].ToString()));

                        //　基本情報等一覧へ連動
                        base_tbl10_txtRakusatuSya.Text = bid_tbl03_1_txtRakusatuSya.Text;
                        base_tbl10_txtRakusatuAmt.Text = bid_tbl03_1_numRakusatuAmt.Text;
                    }
                }
                return;
            }
            if (grid.Name == "ca_tbl06_c1FlexGrid") {
                int iCol = ca_tbl06_c1FlexGrid.Col;
                if (iCol == 1 || iCol == 9 || iCol == 17 || iCol == 25)
                {
                    int iRow = ca_tbl06_c1FlexGrid.Row;
                    // 工期末日付を空にした場合
                    if (ca_tbl06_c1FlexGrid.Rows[iRow][iCol] == null)
                    {
                        ca_tbl06_c1FlexGrid.Rows[iRow][iCol + 1] = null;
                    }
                    else
                    {
                        ca_tbl06_c1FlexGrid.Rows[iRow][iCol + 1] = DateTime.Parse(ca_tbl06_c1FlexGrid.Rows[iRow][iCol].ToString()).ToString("yyyy/MM");

                    }
                }
            }
        }

        /// <summary>
        /// 入札: 入札参加者リスト(bid_tbl03_4_c1FlexGrid)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c1FlexGrid_CellChecked(object sender, RowColEventArgs e)
        {
            C1FlexGrid grid = (C1FlexGrid)sender;
            if (grid.Name == "bid_tbl03_4_c1FlexGrid")
            {
                if (e.Col == 3 & e.Row > 0)
                {
                    var _row = e.Row;
                    var _col = e.Col;
                    for (int i = 1; i < grid.Rows.Count; i++)
                    {
                        if (_row != i)
                        {
                            grid.SetCellCheck(i, 3, C1.Win.C1FlexGrid.CheckEnum.Unchecked);
                        }
                        else
                        {
                            if (grid.GetCellCheck(i, 3) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                            {
                                if (grid.Rows[i][5] != null && grid.Rows[i][5].ToString() != "")
                                {
                                    bid_tbl03_1_txtRakusatuSya.Text = grid.Rows[i][5].ToString();
                                }
                                else
                                {
                                    bid_tbl03_1_txtRakusatuSya.Text = "";
                                }
                                if (grid.Rows[i][6] != null)
                                {
                                    bid_tbl03_1_numRakusatuAmt.Text = GetMoneyTextLong(GetLong(grid.Rows[i][6].ToString()));
                                }
                                else
                                {
                                    bid_tbl03_1_numRakusatuAmt.Text = GetMoneyTextLong(0);
                                }
                            }
                            else
                            {
                                bid_tbl03_1_txtRakusatuSya.Text = "";
                                bid_tbl03_1_numRakusatuAmt.Text = GetMoneyTextLong(0);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// GRID行リサイズ処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c1FlexGrid_AfterResizeRow(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            Resize_Grid(((C1FlexGrid)sender).Name);
        }

        /// <summary>
        /// 削除マックつける
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c1FlexGrid_OwnerDrawCell(object sender, OwnerDrawCellEventArgs e)
        {
            C1FlexGrid grid = (C1FlexGrid)sender;
            if (grid.Name.Equals("ca_tbl05_txtTanto_c1FlexGrid") || grid.Name.Equals("te_c1FlexGrid"))
            {
                //契約：担当者リスト
                if (e.Row >= 1 && e.Col == 0)
                {
                    e.Image = Img_DeleteRowNonactive;
                }
            }
            else
            {
                //以外
                if (e.Row >= 1 && e.Col == 1)
                {
                    e.Image = Img_DeleteRowNonactive;
                }
            }
        }

        #endregion

        #region PicBox イベント ------------------------------------------------------
        /// <summary>
        /// 契約タブ 契約図書
        /// 評価：請求書
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Folder_PictureBox_Click(object sender, EventArgs e)
        {
            // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
            Control ctl = (Control)sender;
            string sFolder = "";
            if (ctl.Name.Equals("ca_tbl01_picTosyo") || ctl.Name.Equals("ca_tbl01_picTosyo"))
            {
                if (ctl.Name.Equals("ca_tbl01_picTosyo"))
                {
                    sFolder = ca_tbl01_txtTosyo.Text;
                }
                else
                {
                    sFolder = ca_tbl01_txtTosyo.Text;
                }
                if (System.Text.RegularExpressions.Regex.IsMatch(sFolder, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    // 指定されたフォルダパスが存在するなら開く
                    if (!string.IsNullOrEmpty(sFolder) && Directory.Exists(sFolder))
                    {
                        Process.Start(GlobalMethod.GetPathValid(sFolder));
                    }
                }
            }
            else if (ctl.Name.Equals("base_tbl02_picAnkenFolder"))
            {
                sFolder = base_tbl02_txtAnkenFolder.Text;
                if (string.IsNullOrEmpty(sFolder))
                {
                    System.Diagnostics.Process.Start("EXPLORER.EXE", "");
                }
                else
                {
                    // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
                    if (System.Text.RegularExpressions.Regex.IsMatch(sFolder, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                    {
                        // 指定されたフォルダパスが存在するなら開く
                        if (Directory.Exists(sFolder))
                        {
                            System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(sFolder));
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
        }

        /// <summary>
        /// 計画番号クリアボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl02_picDelKeikakuNo_Click(object sender, EventArgs e)
        {
            
            if (MessageBox.Show("計画情報を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                // 計画番号、計画案件名のクリア
                base_tbl02_txtKeikakuNo.Text = "";
                base_tbl02_txtKeikakuAKName.Text = "";
                base_tbl02_txtKeikakuNo.Focus();
            }
        }

        /// <summary>
        /// 契約：５．管理者・担当者　の×ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PictureBoxDel_Click(object sender, EventArgs e)
        {
            //ca_tbl05_picKanriDel 契約：管理技術者削除
            //ca_tbl05_picSyosaDel    契約：照査技術者削除
            //ca_tbl05_picSinsaDel    契約：審査担当者削除
            //ca_tbl05_picGyomuDel    契約：業務管理者削除
            //ca_tbl05_picMadoguchiDel    契約：窓口担当者削除
            //te_picKanriDel             評価：管理技術者削除
            //te_picSyosaDel            評価：照査技術者削除

            PictureBox pic = (PictureBox)sender;
            string sMsg = "管理技術者";
            switch (pic.Name)
            {
                case "ca_tbl05_picKanriDel":
                    break;
                case "ca_tbl05_picSyosaDel":
                    sMsg = "照査技術者";
                    break;
                case "ca_tbl05_picSinsaDel":
                    sMsg = "審査技術者";
                    break;
                case "ca_tbl05_picGyomuDel":
                    sMsg = "業務担当者";
                    break;
                case "ca_tbl05_picMadoguchiDel":
                    sMsg = "窓口担当者";
                    break;
                case "te_picKanriDel":
                case "te_picSyosaDel":
                    sMsg = "氏名";
                    break;
            }
            if (MessageBox.Show(sMsg + "を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                switch (pic.Name)
                {
                    case "ca_tbl05_picKanriDel":
                        // 管理技術者
                        ca_tbl05_txtKanri.Text = "";
                        ca_tbl05_txtKanriCD.Text = "";
                        te_lblKanri.Text = "";
                        ca_tbl05_txtKanri.Focus();
                        break;
                    case "ca_tbl05_picSyosaDel":
                        // 照査技術者
                        ca_tbl05_txtSyosa.Text = "";
                        ca_tbl05_txtSyosaCD.Text = "";
                        te_lblSyosa.Text = "";
                        ca_tbl05_txtSyosa.Focus();
                        break;
                    case "ca_tbl05_picSinsaDel":
                        // 審査技術者
                        ca_tbl05_txtSinsa.Text = "";
                        ca_tbl05_txtSinsaCD.Text = "";
                        ca_tbl05_txtSinsa.Focus();
                        break;
                    case "ca_tbl05_picGyomuDel":
                        // 業務担当者
                        ca_tbl05_txtGyomu.Text = "";
                        ca_tbl05_txtGyomuCD.Text = "";
                        ca_tbl05_txtGyomu.Focus();
                        break;
                    case "ca_tbl05_picMadoguchiDel":
                        // 窓口担当者
                        ca_tbl05_txtMadoguchi.Text = "";
                        ca_tbl05_txtMadoguchiCD.Text = "";
                        ca_tbl05_txtMadoguchi.Focus();
                        break;
                    case "te_picKanriDel":
                        // 評価：管理技術者削除
                        te_lblKanri.Text = "";
                        ca_tbl05_txtKanri.Text = "";
                        ca_tbl05_txtKanriCD.Text = "";
                        break;
                    case "te_picSyosaDel":
                        // 評価：照査技術者削除
                        te_lblSyosa.Text = "";
                        ca_tbl05_txtSyosa.Text = "";
                        ca_tbl05_txtSyosaCD.Text = "";
                        break;
                }
            }
        }

        /// <summary>
        /// 契約：５．管理者・担当者　の選択ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PictureBoxSel_Click(object sender, EventArgs e)
        {
            //ca_tbl05_picKanri 契約：管理技術者
            //ca_tbl05_picSyosa    契約：照査技術者
            //ca_tbl05_picSinsa    契約：審査担当者
            //ca_tbl05_picGyomu    契約：業務管理者
            //ca_tbl05_picMadoguchi    契約：窓口担当者
            PictureBox pic = (PictureBox)sender;
            Popup_ChousainList form = new Popup_ChousainList();
            string nendo = DateTime.Today.Year.ToString();
            string bsCD = BushoCD;
            if (pic.Name.Equals("base_tbl02_picKeiyakuTanto"))
            {
                nendo = base_tbl03_cmbKokiStartYear.SelectedValue.ToString();
                if (base_tbl02_cmbJyutakuKasyoSibu.SelectedValue != null)
                {
                    bsCD = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString();
                }
            }
            form.nendo = nendo;
            form.Busho = bsCD;
            form.ShowDialog();
            switch (pic.Name)
            {
                case "ca_tbl05_picKanri":
                case "te_picKanri":
                    // 管理技術者
                    if (form.ReturnValue != null && form.ReturnValue[0] != null) { 
                        ca_tbl05_txtKanri.Text = form.ReturnValue[1];
                        ca_tbl05_txtKanriCD.Text = form.ReturnValue[0];
                        te_lblKanri.Text = form.ReturnValue[1];
                    }
                    ca_tbl05_txtKanri.Focus();
                    break;
                case "ca_tbl05_picSyosa":
                case "te_picSyosa":
                    // 照査技術者
                    if (form.ReturnValue != null && form.ReturnValue[0] != null)
                    {
                        ca_tbl05_txtSyosa.Text = form.ReturnValue[1];
                        ca_tbl05_txtSyosaCD.Text = form.ReturnValue[0];
                        te_lblSyosa.Text = form.ReturnValue[1];
                    }
                    ca_tbl05_txtSyosa.Focus();
                    break;
                case "ca_tbl05_picSinsa":
                    // 審査技術者
                    if (form.ReturnValue != null && form.ReturnValue[0] != null)
                    {
                        ca_tbl05_txtSinsa.Text = form.ReturnValue[1];
                        ca_tbl05_txtSinsaCD.Text = form.ReturnValue[0];
                    }
                    ca_tbl05_txtSinsa.Focus();
                    break;
                case "ca_tbl05_picGyomu":
                    // 業務担当者
                    if (form.ReturnValue != null && form.ReturnValue[0] != null)
                    {
                        ca_tbl05_txtGyomu.Text = form.ReturnValue[1];
                        ca_tbl05_txtGyomuCD.Text = form.ReturnValue[0];
                    }
                    ca_tbl05_txtGyomu.Focus();
                    break;
                case "ca_tbl05_picMadoguchi":
                    // 窓口担当者
                    if (form.ReturnValue != null && form.ReturnValue[0] != null)
                    {
                        ca_tbl05_txtMadoguchi.Text = form.ReturnValue[1];
                        ca_tbl05_txtMadoguchiCD.Text = form.ReturnValue[0];
                        ca_tbl05_txtMadoguchiBusho.Text = form.ReturnValue[2];
                        ca_tbl05_txtMadoguchiShibu.Text = form.ReturnValue[3];
                        ca_tbl05_txtMadoguchiKa.Text = form.ReturnValue[4];
                    }
                    ca_tbl05_txtMadoguchi.Focus();
                    break;
                case "base_tbl02_picKeiyakuTanto":
                    // ２．基本情報　契約担当者
                    if (form.ReturnValue != null && form.ReturnValue[0] != null)
                    {
                        base_tbl02_txtKeiyakuTantoCD.Text = form.ReturnValue[0];
                        base_tbl02_txtKeiyakuTanto.Text = form.ReturnValue[1];
                        base_tbl02_txtKeiyakuTantoBusho.Text = form.ReturnValue[2];
                        base_tbl02_cmbJyutakuKasyoSibu.SelectedValue = form.ReturnValue[2];
                    }
                    base_tbl02_txtKeiyakuTanto.Focus();
                    break;
            }
        }

        /// <summary>
        /// ２．基本情報　計画番号　プロンプト
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl02_picKeikakuNo_Click(object sender, EventArgs e)
        {
            Popup_Keikaku form = new Popup_Keikaku();
            form.gyoumuBushoCD = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue == null ? "" : base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString();
            form.nendo = base_tbl03_cmbKokiStartYear.SelectedValue == null ? "" : base_tbl03_cmbKokiStartYear.SelectedValue.ToString();
            form.ShowDialog();

            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                base_tbl02_txtKeikakuNo.Text = form.ReturnValue[0];
                base_tbl02_txtKeikakuAKName.Text = form.ReturnValue[1];
            }
            base_tbl02_txtKeikakuNo.Focus();
        }

        /// <summary>
        /// 基本情報等一覧：４．発注者
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl04_picOrderCd_Click(object sender, EventArgs e)
        {
            Popup_Hachusha form = new Popup_Hachusha();
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                base_tbl04_txtOrderCd.Text = form.ReturnValue[0];
                base_tbl04_txtOrderKubun1.Text = form.ReturnValue[1];
                base_tbl04_txtOrderKubun2.Text = form.ReturnValue[2];
                base_tbl04_txtTodofuken.Text = form.ReturnValue[3];
                base_tbl04_txtOrderName.Text = form.ReturnValue[4];
            }
            base_tbl04_txtOrderCd.Focus();
        }

        /// <summary>
        /// 郵便番号　プロンプト　押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PictureBoxZip_Click(object sender, EventArgs e)
        {
            PictureBox pb = (PictureBox)sender;

            Popup_Yubin form = new Popup_Yubin();
            string sYubin = base_tbl05_txtZip.Text;
            if(pb.Name.Equals("base_tbl06_picZip")) sYubin = base_tbl06_txtZip.Text;
            form.Yubin = sYubin;
            form.ShowDialog();
            if(pb.Name == "base_tbl05_picZip")
            {
                //５．発注担当者情報（調査窓口）
                if (form.ReturnValue != null && form.ReturnValue[0] != null)
                {
                    base_tbl05_txtZip.Text = form.ReturnValue[0];
                    base_tbl05_txtAddress.Text = form.ReturnValue[1];
                }
                base_tbl05_txtZip.Focus();
            }
            else
            {
                // ６．発注担当者情報（契約窓口）
                if (form.ReturnValue != null && form.ReturnValue[0] != null)
                {
                    base_tbl06_txtZip.Text = form.ReturnValue[0];
                    base_tbl06_txtAddress.Text = form.ReturnValue[1];
                }
                base_tbl06_txtZip.Focus();
            }
            
        }

        #endregion

        #region ヘッダー部 イベント --------------------------------------------------
        /// <summary>
        /// ヘッダー「計画」ボタン押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnKeikaku_Click(object sender, EventArgs e)
        {
            Entry_keikaku_Search form = new Entry_keikaku_Search();
            form.UserInfos = UserInfos;
            form.Show();
            this.Close();
        }

        /// <summary>
        /// ヘッダーの案件ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAnken_Click(object sender, EventArgs e)
        {
            Entry_Search form = new Entry_Search();
            form.UserInfos = this.UserInfos;
            form.Show();
            this.Close();
        }

        /// <summary>
        /// 更新ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (mode ==MODE.INSERT || mode == MODE.PLAN)
            {
                // 売上年度、受託課所支部が正しいか確認してください
                if (MessageBox.Show("新規登録を行いますがよろしいでしょうか？\r\n下記について確認して下さい。\r\n工期開始年度、受託課所支部が正しいか確認して下さい。", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    if (!ErrorFLG(0))
                    {
                        if (Execute_SQL(0))
                        {
                            //えんとり君修正STEP2
                            sJyutakuKasyoSibuCdOri = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString(); //受託課所支部（契約部所）DB値
                            sKokiStartYearOri = base_tbl03_cmbKokiStartYear.SelectedValue.ToString(); //工期開始年度DB値
                            sJigyoubuHeadCD_ori = getJigyoubuHeadCD();
                        }
                    }
                }
            }
            else
            {
                if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    //担当者項目の必須チェックを追加
                    bool isError = ErrorFLG(1);
                    WarningCheck(1);
                    if (!isError)
                    {
                        if (Execute_SQL(1))
                        {
                            sJyutakuKasyoSibuCdOri = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString(); //受託課所支部（契約部所）DB値
                            sKokiStartYearOri = base_tbl03_cmbKokiStartYear.SelectedValue.ToString(); //工期開始年度DB値
                            sJigyoubuHeadCD_ori = getJigyoubuHeadCD();
                            //受託番号が採番された場合、「この案件番号の枝番で受託番号を作成する」ボタンを表示 No.1484
                            btnNewByBranchNo.Visible = (base_tbl02_txtJyutakuEdNo.Text != "");
                        }
                    }
                }

            }
            FolderPathCheck();
        }

        /// <summary>
        /// この業務を元に新規登録ボタン　押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewByCopy_Click(object sender, EventArgs e)
        {
            gotoSelfPage(MODE.INSERT, COPY.GM, AnkenID);
        }

        /// <summary>
        /// この発注者を元に新規登録
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewByOrder_Click(object sender, EventArgs e)
        {
            gotoSelfPage(MODE.INSERT, COPY.HC, AnkenID);
        }

        /// <summary>
        /// この案件番号の枝番で受託番号を作成するボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNewByBranchNo_Click(object sender, EventArgs e)
        {
            //受託番号が採番されていない場合は、処理を終了
            gotoSelfPage(MODE.INSERT, COPY.ED, AnkenID, base_tbl02_txtJyutakuEdNo.Text);
        }

        /// <summary>
        /// 削除ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            string methodName = ".btnDelete_Click";
            using (Popup_MessageBox dlg = new Popup_MessageBox("確認", GlobalMethod.GetMessage("I10605", ""), "案件番号"))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    if (dlg.GetInputText().Equals(tblAKInfo_lblAnkenNo.Text))
                    {
                        ErrorMessage.Text = "";
                        bool ErrorFLG = false;
                        if (ca_tbl01_chkKian.Checked == true)
                        {
                            set_error(GlobalMethod.GetMessage("E10601", ""));
                            ErrorFLG = true;
                        }
                        if (ca_tbl01_cmbAnkenKubun.SelectedValue != null && ca_tbl01_cmbAnkenKubun.SelectedValue.ToString() == "05")
                        {
                            // E10602:計画業務は削除できません。
                            set_error(GlobalMethod.GetMessage("E10602", ""));
                            ErrorFLG = true;
                        }

                        DataTable dt = EntryInputDbClass.CheckAnkenBeforeDelete(AnkenID);
                        if (dt == null || dt.Rows.Count == 0)
                        {
                            // E10009:対象データは存在しません。
                            set_error(GlobalMethod.GetMessage("E10009", ""));
                            ErrorFLG = true;
                        }
                        //単価契約　OR　窓口ミハルが存在するなら削除しない
                        if (!ErrorFLG)
                        {
                            // 単価契約
                            DataTable tankaData = GlobalMethod.getData("AnkenJouhouID", "TankaKeiyakuID", "TankaKeiyaku", "AnkenJouhouID = " + AnkenID.ToString());
                            if (tankaData != null && tankaData.Rows.Count > 0)
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
                            bool bDel = EntryInputDbClass.delete(AnkenID, dt.Rows[0]["AnkenSaishinFlg"].ToString()
                                        , dt.Rows[0]["AnkenSakuseiKubun"].ToString()
                                        , base_tbl02_txtJyutakuNo.Text, base_tbl02_txtJyutakuEdNo.Text
                                        , base_tbl02_txtKeikakuNo.Text, beforeKeikakuBangou);
                            if (bDel)
                            {
                                // 更新履歴の登録
                                GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報を削除しました ID:" + AnkenID, pgmName + methodName, "");

                                this.Owner.Show();
                                this.Close();
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

        /// <summary>
        /// 戻るボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBack_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I00013", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                if (mode == MODE.CHANGE && ChangeFlag == 1)
                {
                    string sId = AnkenID;
                    if (ca_tbl01_hidKuroden.Text != "")
                    {
                        sId = ca_tbl01_hidKuroden.Text;
                    }
                    else if (ca_tbl01_hidAkaden.Text != "")
                    {
                        sId = ca_tbl01_hidAkaden.Text;
                    }
                    gotoSelfPage(MODE.SPACE, COPY.NO, sId,"",0,true);
                }
                else
                {
                    this.Owner.Show();
                    this.Close();
                }
            }
        }

        /// <summary>
        /// 反映する　ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
            //前後のダブルクオーテーションを消す。面倒なので配列全体に先にやってしまう。
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = deleteDoubleQuotation(words[i]);
            }

            //部署・所属名
            if (words.Length > (int)excelIndex.busho_shozoku)
            {
                base_tbl05_txtBusho.Text = words[(int)excelIndex.busho_shozoku];
            }
            //ご担当者名
            if (words.Length > (int)excelIndex.tantosha)
            {
                base_tbl05_txtTanto.Text = words[(int)excelIndex.tantosha];
            }
            //メールアドレス
            if (words.Length > (int)excelIndex.mail)
            {
                base_tbl05_txtEmail.Text = words[(int)excelIndex.mail];
            }
            //郵便番号
            if (words.Length > (int)excelIndex.post_address)
            {
                base_tbl05_txtZip.Text = getPostAddress(words[(int)excelIndex.post_address], true);
            }
            //住所
            if (words.Length > (int)excelIndex.post_address)
            {
                base_tbl05_txtAddress.Text = getPostAddress(words[(int)excelIndex.post_address], false);
            }
            //電話番号
            if (words.Length > (int)excelIndex.tel)
            {
                base_tbl05_txtTel.Text = getTelNumber(words[(int)excelIndex.tel]);
            }
            //FAX番号
            if (words.Length > (int)excelIndex.fax)
            {
                base_tbl05_txtFax.Text = getTelNumber(words[(int)excelIndex.fax]);
            }
            //ご依頼業務名称
            if (words.Length > (int)excelIndex.irai_gyoumu)
            {
                base_tbl03_txtGyomuName.Text = words[(int)excelIndex.irai_gyoumu];
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
            base_tbl03_txtAnkenMemo.Text = tmpBuff;
        }

        /// <summary>
        /// タグ処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tab_DrawItem(object sender, DrawItemEventArgs e)
        {
            Console.WriteLine("====tab_DrawItem ST================================================");
            GlobalMethod.tabDisplaySet(tab, sender, e);
            Console.WriteLine("====tab_DrawItem ED================================================");
        }
        #endregion

        #region 基本情報一覧 イベント ------------------------------------------------
        /// <summary>
        /// 基本情報一覧：８．過去案件情報：行追加ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl08_btnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                
                if (base_tbl08_c1FlexGrid.Rows.Count < 6)
                {
                    // 前回受託番号ID
                    string AnkenZenkaiRakusatsuID = "1";
                    int num = 0;
                    int maxNum = 0;

                    // ヘッダーを除いて回し、前回受託番号IDの最大値を取得する
                    for (int i = 1; i < base_tbl08_c1FlexGrid.Rows.Count; i++)
                    {
                        AnkenZenkaiRakusatsuID = base_tbl08_c1FlexGrid.Rows[i][16].ToString();
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

                    base_tbl08_c1FlexGrid.Rows.Add();
                    // 追加した行にセット
                    base_tbl08_c1FlexGrid.Rows[base_tbl08_c1FlexGrid.Rows.Count - 1][16] = maxNum;
                }
                Resize_Grid("base_tbl08_c1FlexGrid");
            }
            catch (Exception ex)
            {
                set_error(GlobalMethod.GetMessage("E00090", ""));
            }
        }

        /// <summary>
        /// 工期開始年度選択変更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void KoukiStartYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            set_combo_shibu(base_tbl03_cmbKokiStartYear.SelectedValue.ToString());
            if (mode == MODE.INSERT || mode == MODE.PLAN)
            {
                setFolderPath();
                FolderPathCheck();

                // 工期開始年度に合わせて売上年度を変更する 必要なしだと思う？？？？　By Chen
                // DataSourceにセットした時など、想定外のとこでもTextChangedが動いていたため、値のチェックを入れる
                if (int.TryParse(base_tbl03_cmbKokiStartYear.SelectedValue.ToString(), out int num))
                {
                    base_tbl03_cmbKokiSalesYear.SelectedValue = base_tbl03_cmbKokiStartYear.SelectedValue.ToString();
                }
            }
            // 応援依頼先
            set_oensaki_list(base_tbl03_cmbKokiStartYear.SelectedValue.ToString());
        }

        /// <summary>
        /// 受託支部選択変更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void JyutakuSibu_SelectedIndexChanged(object sender, EventArgs e)
        {
            //BushoCD = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue == null ? "" : base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString();
        }

        /// <summary>
        /// 受託支部変更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void JyutakuSibu_TextChanged(object sender, EventArgs e)
        {
            //BushoCD = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue == null ? "" : base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString();
            // フォルダ設定しなおし
            set_folder();
        }

        /// <summary>
        /// 計画番号変更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl02_txtKeikakuNo_TextChanged(object sender, EventArgs e)
        {
            //計画番号選択時、計画情報の契約区分をセット
            if (base_tbl02_txtKeikakuNo.Text != "")
            {
                DataTable dt = GlobalMethod.getData("KeikakuGyoumuKubunMei", "KeikakuGyoumuKubun", "KeikakuJouhou", "KeikakuDeleteFlag <> 1 AND KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + base_tbl02_txtKeikakuNo.Text + "'");
                if (dt != null && dt.Rows.Count > 0 && dt.Rows[0][1] != null)
                {
                    base_tbl03_cmbKeiyakuKubun.SelectedValue = dt.Rows[0][0].ToString();
                }
            }
        }

        /// <summary>
        /// 発注者課名変更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl04_txtOrderKamei_TextChanged(object sender, EventArgs e)
        {

            ca_tbl01_txtOrderKamei.Text = base_tbl04_txtOrderName.Text + " " + base_tbl04_txtOrderKamei.Text;
        }
        /// <summary>
        /// 発注者区分の説明 リンク
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl04_lnkOrderKubun_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(System.Environment.CurrentDirectory + "/Resource/PDF/発注者区分の説明.pdf");
        }

        /// <summary>
        /// 案件（受託）フォルダの値をセット
        /// </summary>
        private void setFolderPath()
        {
            // 新規登録時のみ案件（受託）フォルダの値を動的に変更する
            if (mode == MODE.INSERT || mode == MODE.PLAN)
            {
                // 案件（受託）フォルダを取得
                string folderPath = base_tbl02_txtAnkenFolder.Text;
                // 年度のパスを調べる 売上年度 の部分
                string keyWord = @"\\2[0-9]{3}\\";
                // 売上年度 の開始位置を取得
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
                    folderPath = folderPath.Replace(targetWord.ToString(), "\\" + base_tbl03_cmbKokiStartYear.SelectedValue.ToString() + "\\");

                    // 867
                    // 工期開始年度　2021年度まで、　010北道
                    // 工期開始年度　2022年度から　　010北海
                    int koukinendo = 0;
                    if (int.TryParse(base_tbl03_cmbKokiStartYear.SelectedValue.ToString(), out koukinendo))
                    {
                        folderPath = change_hokaido_path(folderPath, koukinendo);
                    }
                    base_tbl02_txtAnkenFolder.Text = folderPath;
                }
            }
        }

        /// <summary>
        /// No1563 1314　北海道のフォルダ名が間違ってる。　×　010北道　○　010北海 共通化にする
        /// </summary>
        /// <param name="folderPath">変更元Path</param>
        /// <param name="iYear">工期開始年度</param>
        /// <returns></returns>
        private string change_hokaido_path(string folderPath, int iYear)
        {
            string rtnPath = folderPath;

            // 010北道
            string str1 = GlobalMethod.GetCommonValue1("MADOGUCHI_HOKKAIDO_PATH");
            // 010北海
            string str2 = GlobalMethod.GetCommonValue2("MADOGUCHI_HOKKAIDO_PATH");
            if (str1 != null && str2 != null)
            {

                if (iYear > 2021)
                {
                    rtnPath = folderPath.Replace(str1, str2);
                }
                else
                {
                    rtnPath = folderPath.Replace(str2, str1);
                }
            }
            return rtnPath;
        }

        /// <summary>
        /// フォルダセット
        /// </summary>
        private void set_folder()
        {
            // 案件（受託）フォルダをコピー
            // 契約タブ 契約図書
            ca_tbl01_txtTosyo.Text = base_tbl02_txtAnkenFolder.Text;

            string JigyoubuHeadCD = "";
            if (base_tbl02_cmbJyutakuKasyoSibu.Text != null && base_tbl02_cmbJyutakuKasyoSibu.Text != "")
            {
                JigyoubuHeadCD = EntryInputDbClass.JigyoubuHeadCD(base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString());
            }
            // 他部所の場合は、\02契約関係図書を付けない
            if (!"T".Equals(JigyoubuHeadCD) && !"".Equals(JigyoubuHeadCD))
            {
                // 他部署の場合、請求書は案件（受託）フォルダと同じ
                te_txtSeikyusyo.Text = base_tbl02_txtAnkenFolder.Text;
            }
            else
            {
                // 技術担当者 請求書
                te_txtSeikyusyo.Text = base_tbl02_txtAnkenFolder.Text + @"\02契約関係図書";
            }
        }

        /// <summary>
        /// 5.発注担当者情報（調査窓口）からコピーする
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl06_btnTyosaMadoguchiCopy_Click(object sender, EventArgs e)
        {
            //No.1531 チェックボックスからボタンへ変更
            //6.発注担当者情報（契約窓口）
            //部署
            base_tbl06_txtBusho.Text = base_tbl05_txtBusho.Text;
            //担当者名
            base_tbl06_txtTanto.Text = base_tbl05_txtTanto.Text;
            //電話
            base_tbl06_txtTel.Text = base_tbl05_txtTel.Text;
            //FAX
            base_tbl06_txtFax.Text = base_tbl05_txtFax.Text;
            //E - Mail
            base_tbl06_txtEmail.Text = base_tbl05_txtEmail.Text;
            //郵便番号
            base_tbl06_txtZip.Text = base_tbl05_txtZip.Text;
            //住所
            base_tbl06_txtAddress.Text = base_tbl05_txtAddress.Text;
        }

        /// <summary>
        /// チェックボックス　値変更処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_CheckedChanged(Object sender, EventArgs e)
        {
            System.Windows.Forms.CheckBox ck = (System.Windows.Forms.CheckBox)sender;
            if (!ck.Enabled) return;
            bool bChk = ck.Checked;
            //No.1531 チェックボックスからボタンへ変更
            //if (ck.Name == "base_tbl06_chkTyosaMadoguchi" && bChk)
            //{
            //    //部署
            //    base_tbl06_txtBusho.Text = base_tbl05_txtBusho.Text;
            //    //base_tbl06_txtBusho.ReadOnly = bChk;
            //    //担当者名
            //    base_tbl06_txtTanto.Text = base_tbl05_txtTanto.Text;
            //    //base_tbl06_txtTanto.ReadOnly = bChk;
            //    //電話
            //    base_tbl06_txtTel.Text = base_tbl05_txtTel.Text;
            //    //base_tbl06_txtTel.ReadOnly = bChk;
            //    //FAX
            //    base_tbl06_txtFax.Text = base_tbl05_txtFax.Text;
            //    //base_tbl06_txtFax.ReadOnly = bChk;
            //    //E - Mail
            //    base_tbl06_txtEmail.Text = base_tbl05_txtEmail.Text;
            //    //base_tbl06_txtEmail.ReadOnly = bChk;
            //    //郵便番号
            //    base_tbl06_txtZip.Text = base_tbl05_txtZip.Text;
            //    //base_tbl06_txtZip.ReadOnly = bChk;
            //    //base_tbl06_picZip.Visible = !bChk;
            //    //住所
            //    base_tbl06_txtAddress.Text = base_tbl05_txtAddress.Text;
            //    //base_tbl06_txtAddress.ReadOnly = bChk;
            //    return;
            //}

            if (base_tbl01_chkJizendasin.Name.Equals(ck.Name)) { 
                if (ck.Checked)
                {
                    base_tbl01_dtpDtPrior.Text = DateTime.Now.ToString("yyyy/MM/dd");
                }
                else
                {
                    base_tbl01_dtpDtPrior.Text = "";
                    base_tbl01_dtpDtPrior.CustomFormat = " ";
                }
                return;
            }
            if (base_tbl01_chkNyusatu.Name.Equals(ck.Name)){
                if (ck.Checked)
                {
                    base_tbl01_dtpDtBid.Text = DateTime.Now.ToString("yyyy/MM/dd");
                }
                else
                {
                    base_tbl01_dtpDtBid.Text = "";
                    base_tbl01_dtpDtBid.CustomFormat = " ";
                }
                return;
            }

            if (base_tbl01_chkKeiyaku.Name.Equals(ck.Name))
            {
                if (ck.Checked)
                {
                    base_tbl01_dtpDtCa.Text = DateTime.Now.ToString("yyyy/MM/dd");
                }
                else
                {
                    base_tbl01_dtpDtCa.Text = "";
                    base_tbl01_dtpDtCa.CustomFormat = " ";
                }
                return;
            }
        }

        /// <summary>
        /// フォルダ変更ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void base_tbl02_btnRenameFolder_Click(object sender, EventArgs e)
        {
            string FolderPath = "";
            string ankenNo = "";
            MakeFolderFullPath(ref FolderPath, ref ankenNo);
            // 案件（受託）フォルダ
            base_tbl02_txtRenameFolder.Text = FolderPath;

            //フォルダ変更ボタンクリックフラグON
            isClickedRenameFolderButton = true;

            //No1668 ファイル更新ボタンを押下後、変更フォルダが表示されたときに、確認ダイアログを表示させる。OKのみの確認ダイアログとする。
            if (base_tbl02_txtRenameFolder.Text.Length != 0 && base_tbl02_txtRenameFolder.Text != sFolderRenameBef)
            {
                MessageBox.Show(GlobalMethod.GetMessage("E20908", ""), "確認", MessageBoxButtons.OK);

            }
            //案件（受託）フォルダのフルパス作成時に案件番号が作成されていれば非表示案件Noにセット
            if (ankenNo.Length != 0)
			{
                ca_tbl01_hidResetAnkenno.Text = ankenNo;
            }
            
        }

        /// <summary>
        /// 案件（受託）フォルダのフルパス作成
        /// </summary>
        /// <param name="FolderPath"></param>
        /// <param name="ankenNo"></param>
		private void MakeFolderFullPath(ref string FolderPath , ref string ankenNo)
		{

            // フォルダリネーム========================================================
            string sBushoCd = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString();//受託課所支部（契約部所）
            string sYear = base_tbl03_cmbKokiStartYear.SelectedValue.ToString();   // 工期開始年度
            string sGyomu = base_tbl03_txtGyomuName.Text;   // 業務名称
            string sGOrder = base_tbl04_txtOrderName.Text;//発注者名

            sFolderYearRenameBef = sYear;   // 工期開始年度

            // 案件（受託）フォルダ初期値設定 取得
            String discript = "FolderPath";
            String value = "FolderPath ";
            String table = "M_Folder";
            String where = "MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + sBushoCd + "' ";

            // //xxxx/00Cyousa/00調査情報部門共有/$NENDO$/200受託調査関連
            // フォルダ関連は工期開始年度で作成する
            string FolderBase = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", sYear);
            //string FolderPath = "";

            DataTable dt = new System.Data.DataTable();
            dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null && dt.Rows.Count > 0)
            {
                // $FOLDER_BASE$/004 本部
                FolderPath = dt.Rows[0][0].ToString();

                // No.1444 フォルダ変更機能で、北海道でエラーがでる
                // 867
                // 工期開始年度　2021年度まで、　010北道
                // 工期開始年度　2022年度から　　010北海
                int koukinendo = 0;
                if (int.TryParse(sYear, out koukinendo))
                {
                    FolderPath = change_hokaido_path(FolderPath, koukinendo);
                }
            }
            if (FolderBase != "" && FolderPath != "")
            {
                FolderPath = FolderPath.Replace(@"$FOLDER_BASE$", FolderBase);
                FolderPath = FolderPath.Replace("/", @"\");

                string jCd = getJigyoubuHeadCD(1);

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
                    ankenNo = base_tbl02_txtAnkenNo.Text;
                    if (sJyutakuKasyoSibuCdOri.Equals(sBushoCd) == false || sKokiStartYearOri.Equals(sYear) == false)
                    {
                        string jigyoubuHeadCD = "";
                        // 契約区分で業務分類CDを判定
                        // Mst_Jigyoubu に問い合わせる方法が無い為、
                        // 調査部が見つかった場合、T と判断
                        if (base_tbl03_cmbKeiyakuKubun.Text.IndexOf("調査部") > -1)
                        {
                            jigyoubuHeadCD = "T";
                        }
                        else if (base_tbl03_cmbKeiyakuKubun.Text.IndexOf("事業普及部") > -1)
                        {
                            jigyoubuHeadCD = "B";
                        }
                        else if (base_tbl03_cmbKeiyakuKubun.Text.IndexOf("情シス部") > -1)
                        {
                            jigyoubuHeadCD = "J";
                        }
                        else if (base_tbl03_cmbKeiyakuKubun.Text.IndexOf("総合研究所") > -1)
                        {
                            jigyoubuHeadCD = "K";
                        }
                        var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                        using (var conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            var cmd = conn.CreateCommand();

                            // 業務分類CD + 年度下2桁
                            ankenNo = jigyoubuHeadCD + sYear.Substring(2, 2);

                            // KashoShibuCD
                            cmd.CommandText = "SELECT  " +
                                    "KashoShibuCD " +

                                    //参照テーブル
                                    "FROM Mst_Busho " +
                                    "WHERE GyoumuBushoCD = '" + sBushoCd + "' ";
                            var sda = new SqlDataAdapter(cmd);
                            var dtB = new DataTable();
                            sda.Fill(dtB);
                            // KashoShibuCDが正しい
                            ankenNo = ankenNo + dtB.Rows[0][0].ToString();
                        }
                        ankenNo = ankenNo + "●●●";
                    }
                    FolderPath = FolderPath + "\\" + ankenNo + "_" + sGOrder + "_" + sGyomu;
                    //ca_tbl01_hidResetAnkenno.Text = ankenNo;
                }
            }
            else
            {
                FolderPath = FolderBase;
                FolderPath = FolderPath.Replace("/", @"\");
            }
            
		}

		#endregion

		#region 事前打診 イベント ----------------------------------------------------

		#endregion

		#region 入札 イベント --------------------------------------------------------
		/// <summary>
		/// 入札：３．入札結果　応札者追加ボタン押下
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void bid_tbl03_4_btnAdd_Click(object sender, EventArgs e)
        {
            if (bid_tbl03_4_c1FlexGrid.Rows.Count < 11)
            {
                bid_tbl03_4_c1FlexGrid.AllowAddNew = true;
                bid_tbl03_4_c1FlexGrid.Rows.Add();
                Resize_Grid("bid_tbl03_4_c1FlexGrid");
                bid_tbl03_4_c1FlexGrid.AllowAddNew = false;
            }
        }
        #endregion

        #region 契約 イベント --------------------------------------------------------
        /// <summary>
        /// ２．配分情報・業務内容 契約金額コピーボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_tbl02_btnCopy_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;
            System.Windows.Forms.TextBox txt = null;//配分額（税込）
            System.Windows.Forms.TextBox txtCal = null;//配分額（税抜）
            System.Windows.Forms.Label lblBase = null;//基本情報等一覧：配分額（税抜）
            string sNo = "";

            if (btn.Name.Equals(ca_tbl02_btnTSbu.Name))
            {
                // 調査部コピーボタン押下
                txt = ca_tbl02_AftCaBmZeikomi_numAmt1;
                txtCal = ca_tbl02_AftCaBm_numAmt1;
                lblBase = base_tbl07_4_lblAmt1;
                sNo = "1";
            }

            if (btn.Name.Equals(ca_tbl02_btnJGbu.Name))
            {
                // 事業普及部コピーボタン押下
                txt = ca_tbl02_AftCaBmZeikomi_numAmt2;
                txtCal = ca_tbl02_AftCaBm_numAmt2;
                lblBase = base_tbl07_4_lblAmt2;
                sNo = "2";
            }

            if (btn.Name.Equals(ca_tbl02_btnJHbu.Name))
            {
                // 情報システム部コピーボタン押下
                txt = ca_tbl02_AftCaBmZeikomi_numAmt3;
                txtCal = ca_tbl02_AftCaBm_numAmt3;
                lblBase = base_tbl07_4_lblAmt3;
                sNo = "3";
            }

            if (btn.Name.Equals(ca_tbl02_btnSGSyo.Name))
            {
                // 総合研究所コピーボタン押下
                txt = ca_tbl02_AftCaBmZeikomi_numAmt4;
                txtCal = ca_tbl02_AftCaBm_numAmt4;
                lblBase = base_tbl07_4_lblAmt4;
                sNo = "4";
            }

            if (txt != null)
            {
                //配分額（税込）
                txt.Text = ca_tbl01_txtJyutakuAmt.Text;
                GetTotalMoney("ca_tbl02_AftCaBmZeikomi_numAmt", 5);

                //配分額（税抜）
                if (GetLong(ca_tbl01_txtJyutakuGaiAmt.Text) == 0)
                {
                    txtCal.Text = ca_tbl01_txtZeinukiAmt.Text;

                }
                else
                {
                    txtCal.Text = GetMoneyTextLong(Get_Zeinuki(GetLong(txt.Text)));
                }
                GetTotalMoney("ca_tbl02_AftCaBm_numAmt", 5);
                lblBase.Text = txtCal.Text;
                base_tbl07_4_lblAmtAll.Text = ca_tbl02_AftCaBm_numAmtAll.Text;
                // 調査部業務別配分（No.1450）
                calc_aftCaTsFreeTax();

                // 配分率再計算
                cal_aftCaBmPercent(sNo);
            }
        }

        /// <summary>
        /// 契約：６．売上情報の各コピーボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_tbl06_copyButton_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Label btn = (System.Windows.Forms.Label)sender;
            string sDt = ca_tbl01_dtpKokiTo.Text.Trim();  // 工期末日付
            // 計上月
            string sYm = "";
            try
            {
                sYm = DateTime.Parse(sDt).ToString("yyyy/MM");
            }
            catch (Exception)
            {
                // 何もしない
            }
            //計上額
            //string sAmt = ca_tbl01_txtZeikomiAmt.Text;
            string sAmt = ca_tbl01_txtJyutakuAmt.Text;

            string sName = btn.Name;
            if (sName.Equals("ca_tbl06_btnChosa"))
            {
                // 調査部
                if (sDt != "") ca_tbl06_c1FlexGrid.Rows[2][1] = sDt;
                if (sYm != "") ca_tbl06_c1FlexGrid.Rows[2][2] = sYm;
                ca_tbl06_c1FlexGrid.Rows[2][3] = sAmt;
            }
            else if (sName.Equals("ca_tbl06_btnJigyoHukyu"))
            {
                // 事業普及部
                if (sDt != "") ca_tbl06_c1FlexGrid.Rows[2][9] = sDt;
                if (sYm != "") ca_tbl06_c1FlexGrid.Rows[2][10] = sYm;
                ca_tbl06_c1FlexGrid.Rows[2][11] = sAmt;
            }
            else if (sName.Equals("ca_tbl06_btnJohoSystem"))
            {
                // 情報システム部
                if (sDt != "") ca_tbl06_c1FlexGrid.Rows[2][17] = sDt;
                if (sYm != "") ca_tbl06_c1FlexGrid.Rows[2][18] = sYm;
                ca_tbl06_c1FlexGrid.Rows[2][19] = sAmt;
            }
            else if (sName.Equals("ca_tbl06_btnSogoKenkyu"))
            {
                // 総合研究所
                if (sDt != "") ca_tbl06_c1FlexGrid.Rows[2][25] = sDt;
                if (sYm != "") ca_tbl06_c1FlexGrid.Rows[2][26] = sYm;
                ca_tbl06_c1FlexGrid.Rows[2][27] = sAmt;
            }
        }

        /// <summary>
        /// 契約：７．請求書情報 「契約金額(税込)と工期末日付をコピーする。」押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_tbl07_btnCopy_Click(object sender, EventArgs e)
        {
            if (ca_tbl01_dtpKokiTo.CustomFormat != "")
            {
                ca_tbl07_dtpRequst1.CustomFormat = " ";
            }
            else
            {
                ca_tbl07_dtpRequst1.Text = ca_tbl01_dtpKokiTo.Text;
            }
            ca_tbl07_txtRequst1.Text = ca_tbl01_txtZeikomiAmt.Text;
            GetTotalMoney("ca_tbl07_txtRequst", 7);
        }

        /// <summary>
        /// 契約：１．契約情報　「工期末日付け、及び、請求(1回目)に設定」ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_tbl01_btnSetting_Click(object sender, EventArgs e)
        {
            if (ca_tbl01_dtpKokiTo.CustomFormat != "")
            {
                set_error("", 0);
                set_error(GlobalMethod.GetMessage("E10011", ""));
            }
            else
            {
                string sAmt = ca_tbl01_txtJyutakuAmt.Text;
                string GyoumuCD = ca_tbl01_cmbCaKubun.SelectedValue.ToString();
                if (GyoumuCD == "1" || GyoumuCD == "2" || GyoumuCD == "3" || GyoumuCD == "4")
                {
                    ca_tbl06_c1FlexGrid.Rows[2][1] = ca_tbl01_dtpKokiTo.Text;
                    ca_tbl06_c1FlexGrid.Rows[2][2] = DateTime.Parse(ca_tbl01_dtpKokiTo.Text).ToString("yyyy/MM");
                    ca_tbl06_c1FlexGrid.Rows[2][3] = sAmt;
                }
                else if (GyoumuCD == "5" || GyoumuCD == "6")
                {
                    ca_tbl06_c1FlexGrid.Rows[2][9] = ca_tbl01_dtpKokiTo.Text;
                    ca_tbl06_c1FlexGrid.Rows[2][10] = DateTime.Parse(ca_tbl01_dtpKokiTo.Text).ToString("yyyy/MM");
                    ca_tbl06_c1FlexGrid.Rows[2][11] = sAmt;
                }
                else if (GyoumuCD == "7")
                {
                    ca_tbl06_c1FlexGrid.Rows[2][17] = ca_tbl01_dtpKokiTo.Text;
                    ca_tbl06_c1FlexGrid.Rows[2][18] = DateTime.Parse(ca_tbl01_dtpKokiTo.Text).ToString("yyyy/MM");
                    ca_tbl06_c1FlexGrid.Rows[2][19] = sAmt;
                }
                else if (GyoumuCD == "8")
                {
                    ca_tbl06_c1FlexGrid.Rows[2][25] = ca_tbl01_dtpKokiTo.Text;
                    ca_tbl06_c1FlexGrid.Rows[2][26] = DateTime.Parse(ca_tbl01_dtpKokiTo.Text).ToString("yyyy/MM");
                    ca_tbl06_c1FlexGrid.Rows[2][27] = sAmt;
                }
                ca_tbl07_dtpRequst1.Text = ca_tbl01_dtpKokiTo.Text;
                ca_tbl07_txtRequst1.Text = ca_tbl01_txtZeikomiAmt.Text;
                GetTotalMoney("ca_tbl07_txtRequst", 7);
            }
        }

        /// <summary>
        /// 消費税率が入力されていれば、
        /// 税抜（自動計算用）を基に
        /// 税込と内消費税を計算して表示する
        /// </summary>
        private void calc_kingaku()
        {
            long zeinuki = GetLong(ca_tbl01_txtZeinukiAmt.Text);
            long syouhizeiritu = 0;
            long syouhizei = 0;
            long zeikomi = zeinuki;

            if (ca_tbl01_txtTax.Text != "" && ca_tbl01_txtTax.Text != "0")
            {
                // 数値に変換できるか確認
                if (Int64.TryParse(ca_tbl01_txtTax.Text, out syouhizeiritu))
                {
                    syouhizei = syouhizeiritu * zeinuki / 100;
                    zeikomi = zeinuki + syouhizei;
                }
            }
            ca_tbl01_txtZeikomiAmt.Text = string.Format("{0:C}", zeikomi);
            ca_tbl01_txtSyohizeiAmt.Text = string.Format("{0:C}", syouhizei);
            calc_jyutakuAmt();
        }

        /// <summary>
        /// 受託金額（税込）再計算
        /// </summary>
        private void calc_jyutakuAmt()
        {
            // 受託金額(税込)再計算
            long kingaku = GetLong(ca_tbl01_txtZeikomiAmt.Text) - GetLong(ca_tbl01_txtJyutakuGaiAmt.Text);
            ca_tbl01_txtJyutakuAmt.Text = GetMoneyTextLong(kingaku);
            // 基本情報等一覧への連動
            if (mode != MODE.CHANGE)
            {
                base_tbl11_1_txtKeiyakuAmt.Text = GetMoneyTextLong(Get_Zeinuki(kingaku));
            }
        }

        /// <summary>
        /// 起案ボタン押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_btnKian_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show(GlobalMethod.GetMessage("I10704", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {

                if (!ErrorFLG(1))
                {
                    if (base_tbl01_chkKeiyaku.Checked == false)
                    {
                        // No1592 1321　エントリくんの契約画面で、メッセージの誤植がある。
                        //set_error("進捗階段の契約をチェックしてください。");
                        set_error(GlobalMethod.GetMessage("E10740","基本情報"));
                        return;
                    }

                    if (KianError())
                    {
                        if (Execute_SQL(3))
                        {
                            gotoSelfPage(Entry_Input_New.MODE.SPACE, Entry_Input_New.COPY.NO, AnkenID, "", 0, true, false, ErrorMessage.Text + Environment.NewLine + GlobalMethod.GetMessage("I10708", ""));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 契約：チェック用帳票出力・内容確認　ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_btnRptCheck_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string sText = ca_tbl01_cmbAnkenKubun.Text;
                string sVal = ca_tbl01_cmbAnkenKubun.SelectedValue == null ? "" : ca_tbl01_cmbAnkenKubun.SelectedValue.ToString();
                int ListID = 2;
                if (sText != "" && (sVal == "03" || int.Parse(sVal) > 5))
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
                //起案済みの場合は帳票出力のみ実行する
                if (ca_tbl01_chkKian.Checked && mode != MODE.CHANGE && mode != MODE.INSERT && mode != MODE.PLAN)
                {
                    ErrorMessage.Text = "";
                    
                }
                if (!ErrorFLG(2))
                {
                    Execute_SQL(2);
                    // 出力実行
                    // No1589 1320 エントリーシート、確認用と本番用が同じものが出ている。
                    //outPutReport(ListID, AnkenID);
                    outPutReport(ListID, AnkenID, "1", "0");
                }
            }
        }

        /// <summary>
        /// エントリーシート作成・出力   ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_btnOutSheet_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                if (!ErrorFLG(3))
                {
                    Execute_SQL(2);

                    int ListID = 1;
                    string sText = ca_tbl01_cmbAnkenKubun.Text;
                    string sVal = ca_tbl01_cmbAnkenKubun.SelectedValue == null ? "" : ca_tbl01_cmbAnkenKubun.SelectedValue.ToString();
                    if (sText != "" && (sVal == "03" || int.Parse(sVal) > 5))
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
                    // 出力実行
                    outPutReport(ListID, AnkenID,"0", "0");
                }
        }
        }

        /// <summary>
        /// 起案解除ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_btnKianKaijyo_Click(object sender, EventArgs e)
        {
            string methodName = ".btnKianKaijo_Click";

            if (MessageBox.Show(GlobalMethod.GetMessage("I10707", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                if (EntryInputDbClass.KianKaijyo(AnkenID))
                {
                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "起案解除しました ID:" + AnkenID, pgmName + methodName, "");
                    gotoSelfPage(MODE.UPDATE, COPY.NO, AnkenID, "", 0, false, true, GlobalMethod.GetMessage("I10709", ""));
                }
                else
                {
                    set_error(GlobalMethod.GetMessage("E00090", "起案解除"));
                }               
            }
        }

        /// <summary>
        /// 変更伝票ボタン 押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ca_btnChangeSlip_Click(object sender, EventArgs e)
        {
            bool ErrorFLG = true;
            set_error("", 0);
            if (int.Parse(UserInfos[4]) != 2 && !UserInfos[2].Equals(base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString()))
            {
                set_error(GlobalMethod.GetMessage("E10003", ""));
                ErrorFLG = false;
            }

            if (saishinFLG != 1)
            {
                set_error(GlobalMethod.GetMessage("E10006", ""));
                ErrorFLG = false;
            }
            if (String.IsNullOrEmpty(ca_tbl01_cmbAnkenKubun.Text) || ca_tbl01_cmbAnkenKubun.SelectedValue.ToString() == "02" || ca_tbl01_cmbAnkenKubun.SelectedValue.ToString() == "04")
            {
                set_error(GlobalMethod.GetMessage("E10007", ""));
                ErrorFLG = false;
            }
            if (!ca_tbl01_chkKian.Checked)
            {
                set_error(GlobalMethod.GetMessage("E10008", ""));
                ErrorFLG = false;
            }

            if (ErrorFLG)
            {
                gotoSelfPage(MODE.CHANGE, COPY.NO, AnkenID, "", 1);
            }
        }

        /// <summary>
        /// 黒伝・中止伝票作成・出力 ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cs_btnOutRedBlack_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                int ListID = 1;
                string sText = ca_tbl01_cmbAnkenKubun.Text;
                string sVal = ca_tbl01_cmbAnkenKubun.SelectedValue == null ? "" : ca_tbl01_cmbAnkenKubun.SelectedValue.ToString();
                if (sText != "" && (sVal == "03" || int.Parse(sVal) > 5))
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
                string ankenJouhouID = "";
                if (sText != "" && sVal == "04")
                {
                    // 案件区分が04：中止の場合、帳票プログラムには赤伝のAnkenJouhouIDを渡す
                    ankenJouhouID = ca_tbl01_hidAkaden.Text;
                }
                else
                {
                    // 黒伝のAnkenJouhouID
                    ankenJouhouID = ca_tbl01_hidKuroden.Text;
                }
                // 出力実行
                outPutReport(ListID, ankenJouhouID);
            }
        }

        /// <summary>
        /// 変更伝票画面から確認用エントリーチェックシート
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cs_btnRptCheck_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                ErrorMessage.Text = "";
                if (string.IsNullOrEmpty(ca_tbl01_cmbAnkenKubun.Text))
                {
                    // No1592 1321　エントリくんの契約画面で、メッセージの誤植がある。
                    //set_error("案件区分を選択してください。");
                    set_error(GlobalMethod.GetMessage("E10741",""));
                    return;
                }
                KianError(1);
                // エラーでも確認シートを出力する    
                string sKubun = ca_tbl01_cmbAnkenKubun.SelectedValue.ToString();
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
                //int ankenJouhouID = 0;
                int ankenJouhouID = Create_DummyData();
                if (ankenJouhouID > 0)
                {
                    // 出力実行
                    outPutReport(ListID, ankenJouhouID.ToString(),"1", "0");
                }
                else
                {
                    // エラーが発生しました
                    // No1592 1321　エントリくんの契約画面で、メッセージの誤植がある。
                    //set_error("出力用データを作成する時にエラーが発生しました。");
                    set_error(GlobalMethod.GetMessage("E10742", ""));
                }
            }
        }

        /// <summary>
        /// 確認シートダミーデータ作成
        /// </summary>
        /// <returns></returns>
        private int Create_DummyData()
        {
            int rtnAnkenNo = 0;
            string methodName = ".Create_DummyData";
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();

                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;

                try
                {
                    //顧客契約情報存在チェック処理
                    if (!GlobalMethod.Check_Table(AnkenID, "KokyakuKeiyakuID", "KokyakuKeiyakuJouhou", ""))
                    {
                        GlobalMethod.outputLogger("ChangeKianEntry", "顧客契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                        transaction.Rollback();
                        conn.Close();
                        return rtnAnkenNo;
                    }
                    //業務情報存在チェック処理
                    if (!GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhou", ""))
                    {
                        GlobalMethod.outputLogger("ChangeKianEntry", "業務情報が見つからない", "ID:" + AnkenID, "DEBUG");
                        transaction.Rollback();
                        conn.Close();
                        return rtnAnkenNo;
                    }

                    // 契約情報存在チェック
                    if (!GlobalMethod.Check_Table(AnkenID, "KeiyakuJouhouEntoryID", "KeiyakuJouhouEntory", ""))
                    {
                        GlobalMethod.outputLogger("ChangeKianEntry", "契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                        transaction.Rollback();
                        conn.Close();
                        return rtnAnkenNo;
                    }

                    if (!GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                    {
                        GlobalMethod.outputLogger("ChangeKianEntry", "入札情報が見つからない", "ID:" + AnkenID, "DEBUG");
                        transaction.Rollback();
                        conn.Close();
                        return rtnAnkenNo;
                    }

                    string SakuseiKubun = ca_tbl01_cmbAnkenKubun.SelectedValue.ToString();
                    // ダミーの赤伝のAnkenJouhouID取得
                    int ankenNo = GlobalMethod.getSaiban("AnkenJouhouID");

                    //// 案件情報　INSERT　カラム共通
                    //string sAkSqlCom = getAnkenJouhouInsertSQL();
                    // 案件情報の赤伝のダミーデータ作成 ----------------------------------------------------------------------------------
                    var result = createAnkenJouhou(cmd,ankenNo.ToString(),SakuseiKubun,70);

                    // 案件情報前回落札情報作成
                    result = createAnkenJouhouZenkaiRakusatsu(null,cmd,ankenNo.ToString(),70);

                    // 顧客契約情報
                    result = createKokyakuKeiyakuJouhou(cmd,ankenNo.ToString(),70);

                    // 業務情報
                    result = createGyoumuJouhou(cmd, ankenNo.ToString(), 70);


                    if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyouronTantouL1", ""))
                    {
                        result = createGyoumuJouhouHyouronTantouL1(cmd, ankenNo.ToString(), 70);
                    }

                    if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouMadoguchi", ""))
                    {
                        // 窓口担当者
                        // 新では1件しか入らないが、現行が複数件はいるので、複数件あった場合でも落ちないようにする
                        result = createGyoumuJouhouMadoguchi(cmd,ankenNo.ToString(),70);
                    }

                    if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyoutenBusho", ""))
                    {
                        result = createGyoumuJouhouHyoutenBusho(cmd, ankenNo.ToString(), 70);
                    }

                    // 契約情報
                    result = createKeiyakuJouhouEntory(cmd, ankenNo.ToString(), 70);

                    if (GlobalMethod.Check_Table(AnkenID, "RibcID", "RibcJouhou", ""))
                    {
                        result = createRibcJouhou(cmd, ankenNo.ToString(),70);
                    }

                    result = createNyuusatsuJouhou(cmd,ankenNo.ToString(),70);

                    if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhouOusatsusha", ""))
                    {
                        result = createNyuusatsuJouhouOusatsusha(cmd,ankenNo.ToString(),70);
                    }

                    //　業務配分
                    DataTable GH_dt = new DataTable();
                    GH_dt = GlobalMethod.getData("GyoumuHaibunID", "GyoumuAnkenJouhouID", "GyoumuHaibun", "GyoumuAnkenJouhouID = " + AnkenID);
                    result = createGyoumuHaibun(cmd, ankenNo.ToString(), GH_dt,70);

                    // 案件情報の黒伝のダミーデータ作成 ----------------------------------------------------------------------------------
                    int ankenNo2 = 0;
                    if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                    {
                        // 黒伝のAnkenJouhouID
                        ankenNo2 = GlobalMethod.getSaiban("AnkenJouhouID");
                        //案件情報
                        result = createAnkenJouhou(cmd, ankenNo2.ToString(), SakuseiKubun, 71);

                        // 案件情報前回落札情報作成
                        result = createAnkenJouhouZenkaiRakusatsu(null, cmd, ankenNo2.ToString(), 71);


                        // 顧客契約情報
                        result = createKokyakuKeiyakuJouhou(cmd, ankenNo2.ToString(), 71);

                        // 業務情報
                        result = createGyoumuJouhou(cmd, ankenNo2.ToString(), 71);

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyouronTantouL1", ""))
                        {
                            // 業務評点担当者
                            result = createGyoumuJouhouHyouronTantouL1(cmd, ankenNo2.ToString(), 71);
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouMadoguchi", ""))
                        {
                            // 窓口担当者
                            // 新では1件しか入らないが、現行が複数件はいるので、複数件あった場合でも落ちないようにする
                            result = createGyoumuJouhouMadoguchi(cmd, ankenNo2.ToString(), 71);
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyoutenBusho", ""))
                        {
                            result = createGyoumuJouhouHyoutenBusho(cmd, ankenNo2.ToString(), 71);
                        }

                        // 契約情報
                        result = createKeiyakuJouhouEntory(cmd, ankenNo2.ToString(), 71);

                        if (GlobalMethod.Check_Table(AnkenID, "RibcID", "RibcJouhou", ""))
                        {
                            result = createRibcJouhou(cmd, ankenNo2.ToString(), 71);
                        }

                        result = createNyuusatsuJouhou(cmd, ankenNo2.ToString(), 71);

                        if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhouOusatsusha", ""))
                        {
                            result = createNyuusatsuJouhouOusatsusha(cmd, ankenNo2.ToString(), 71);
                        }

                        result = createGyoumuHaibun(cmd, ankenNo2.ToString(), GH_dt, 71);
                    }
                    transaction.Commit();

                    transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    if (updateAfterCreateRedAndBlack(cmd, transaction, SakuseiKubun, ankenNo.ToString(), ankenNo2.ToString(), 7))
                    {

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
                }
                catch (Exception)
                {
                    throw;
                }
            }
            return rtnAnkenNo;
        }

        /// <summary>
        /// 変更伝票時の起案ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cs_btnKian_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(GlobalMethod.GetMessage("I10704", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                if (KianError())
                {

                    if (Execute_SQL(4))
                    {
                        // 変更後の起案ボタン
                        cs_btnKian.Enabled = false;
                        cs_btnKian.BackColor = Color.DarkGray;

                        // 黒伝・中止伝票作成・出力
                        cs_btnOutRedBlack.Enabled = true;
                        cs_btnOutRedBlack.BackColor = Color.FromArgb(42, 78, 122);

                        // 起案したので、案件区分を編集不可に
                        ca_tbl01_cmbAnkenKubun.Enabled = false;
                    }
                    else
                    {
                        set_error(GlobalMethod.GetMessage("E10009", ""));
                    }
                }
            }
        }


        /// <summary>
        /// レポート出力実行共通処理
        /// </summary>
        /// <param name="ListID"></param>
        private void outPutReport(int ListID, string sAnkenId, string sFlag1 = "0", string sFlag2 = "1")
        {
            string[] result = GlobalMethod.InsertReportWork(ListID, UserInfos[0], new string[] { sAnkenId, tblAKInfo_lblAnkenNo.Text, sFlag1, sFlag2 });
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="mode">モード</param>
        /// <param name="copyMode">コピーモード</param>
        /// <param name="sAnkenId">案件ID</param>
        /// <param name="sEdNo">案件枝番</param>
        /// <param name="iChangeFlag">変更伝票フラグ</param>
        /// <param name="bKianFlg">起案解除フラグ</param>
        private void gotoSelfPage(MODE mode, COPY copyMode, string sAnkenId, string sEdNo = "", int iChangeFlag = 0, bool bKianFlg = false,bool bKianKaijyoFlg = false, string sMsg = "")
        {
            Entry_Input_New form = new Entry_Input_New();
            form.mode = mode;
            form.copy = copyMode;
            form.AnkenID = sAnkenId;
            form.AnkenbaBangou = sEdNo;
            form.UserInfos = UserInfos;
            form.ChangeFlag = iChangeFlag;
            form.KianFLG = bKianFlg;
            form.KianKaijoFLG = bKianKaijyoFlg;
            form.Message = sMsg;
            form.Show(this.Owner);
            ownerflg = false;
            this.Close();
        }
        #endregion

        #region 技術者評価 イベント --------------------------------------------------
        /// <summary>
        /// 契約：担当者
        /// 技術者評価：担当者
        /// 　行追加ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTantoAdd_Click(object sender, EventArgs e)
        {
            ErrorMessage.Text = "";

            ErrorMessage.Text = "";

            if (ca_tbl05_txtTanto_c1FlexGrid.Rows.Count < 11)
            {
                ca_tbl05_txtTanto_c1FlexGrid.AllowAddNew = true;
                Resize_Grid("ca_tbl05_txtTanto_c1FlexGrid");
                ca_tbl05_txtTanto_c1FlexGrid.Rows.Add();
                ca_tbl05_txtTanto_c1FlexGrid.AllowAddNew = false;
                te_c1FlexGrid.AllowAddNew = true;
                Resize_Grid("te_c1FlexGrid");
                te_c1FlexGrid.Rows.Add();
                te_c1FlexGrid.AllowAddNew = false;
            }
            else
            {
                set_error(GlobalMethod.GetMessage("E10914", ""));
            }
        }

        private void te_picSeikyusyo_Click(object sender, EventArgs e)
        {
            // ファイルパスとして認識できる場合のみ、エクスプローラーで表示する
            if (System.Text.RegularExpressions.Regex.IsMatch(te_txtSeikyusyo.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                // 指定されたフォルダパスが存在するなら開く
                if (te_txtSeikyusyo.Text != "" && te_txtSeikyusyo.Text != null && Directory.Exists(te_txtSeikyusyo.Text))
                {
                    System.Diagnostics.Process.Start(GlobalMethod.GetPathValid(te_txtSeikyusyo.Text));
                }
            }
        }

        #endregion

        #region 共通処理 Private -----------------------------------------------------
        /// <summary>
        /// 部署：事業部ヘッダーコード取得
        /// </summary>
        /// <param name="getFlag">取得区分 0：DBから取得、1:契約区分より取得</param>
        /// <returns></returns>
        private string getJigyoubuHeadCD(int getFlag = 0, string sSibu = "")
        {
            if (getFlag == 1)
            {
                // No1593 1322　エントリくんのコピー機能で、他事業部の場合も案件番号がTになる。
                string jigyoubuHeadCD = "";
                // 調査部が見つかった場合、T と判断
                if (base_tbl03_cmbKeiyakuKubun.Text.IndexOf("調査部") > -1)
                {
                    jigyoubuHeadCD = "T";
                }
                else if (base_tbl03_cmbKeiyakuKubun.Text.IndexOf("事業普及部") > -1)
                {
                    jigyoubuHeadCD = "B";
                }
                else if (base_tbl03_cmbKeiyakuKubun.Text.IndexOf("情シス部") > -1)
                {
                    jigyoubuHeadCD = "J";
                }
                else if (base_tbl03_cmbKeiyakuKubun.Text.IndexOf("総合研究所") > -1)
                {
                    jigyoubuHeadCD = "K";
                }
                return jigyoubuHeadCD;
            }
            else
            {
                //SQL変数
                string discript = "GyoumuBushoCD";
                string value = "JigyoubuHeadCD";
                string table = "Mst_Busho";
                string where = "GyoumuBushoCD = '" + base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString() + "'";
                if (string.IsNullOrEmpty(sSibu) == false)
                {
                    where = "GyoumuBushoCD = '" + sSibu + "'";
                }
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

        /// <summary>
        /// エラーメッセージ表示設定
        /// </summary>
        /// <param name="mes"></param>
        /// <param name="flg"></param>
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

        /// <summary>
        /// Gridリサイズ処理
        /// </summary>
        /// <param name="name"></param>
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

        /// <summary>
        /// 金額の合計計算
        /// 　入力コントロールの定義ルール：××××numAmt1～（連番）
        /// 　自動合計コントロールの定義ルール：××××numAmtAll
        /// </summary>
        /// <param name="ctlSumName">「××××numAmt」</param>
        /// <param name="num">合計する入力コントロールの数</param>
        /// <param name="start">どの入力コントロールから合計する</param>
        private void GetTotalMoney(string ctlSumName, int num, int start = 1)
        {
            long total = 0;
            num += start - 1;
            for (int i = start; i < num; i++)
            {
                total += GetLong(this.Controls.Find(ctlSumName + i, true)[0].Text);
            }
            Controls.Find(ctlSumName + "All", true)[0].Text = GetMoneyTextLong(total);
        }

        /// <summary>
        /// 配分率など（%）の合計計算
        /// 　入力コントロールの定義ルール：××××numPercent1～（連番）
        /// 　自動合計コントロールの定義ルール：××××numPercentAll
        /// </summary>
        /// <param name="ctlSumName">「××××numPercent」</param>
        /// <param name="num">合計する入力コントロールの数</param>
        /// <param name="start">どの入力コントロールから合計する</param>
        private void GetTotalPercent(string name, int num, int start = 1)
        {
            double total = 0.00;
            num += start - 1;
            for (int i = start; i < num; i++)
            {
                total += GetDouble(this.Controls.Find(name + i, true)[0].Text);
            }
            Controls.Find(name + "All", true)[0].Text = string.Format("{0:F2}", total) + "%";
        }

        /// <summary>
        /// 連動コントロールへデータ設定する
        /// </summary>
        /// <param name="ctlSumName"></param>
        /// <param name="num"></param>
        /// <param name="start"></param>
        private void SetGearingPercent(string ctlToName, string val)
        {
            this.Controls.Find(ctlToName, true)[0].Text = val;
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
            //No1623対応（上限100%を外して入力配分額をそのまま反映する）
            //if (num > 100)
            //{
            //    num = 100;
            //}
            string str = string.Format("{0:F2}", num) + "%";
            return str;
        }
        private long GetLong(string str)
        {
            long num = 0;
            string strVal = str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty);
            if (strVal.Contains("."))
            {
                strVal = (strVal.Split('.'))[0];
            }
            long.TryParse(strVal, out num);
            return num;
        }
        private string GetMoneyTextLong(long num)
        {
            string str = string.Format("{0:C}", num);
            return str;
        }

        /// <summary>
        /// 税抜き金額を算出する
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        private long Get_Zeinuki(long num)
        {
            long zei = GetInt(ca_tbl01_txtTax.Text) + 100;
            long tmp = num * 100;
            long zeinuki = tmp / zei;
            return zeinuki;
        }

        /// <summary>
        /// 税込み金額を算出する
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        private long Get_Zeikomi(long num)
        {
            long zei = GetInt(ca_tbl01_txtTax.Text) + 100;
            // long zeinuki = num * zei
            long zeinuki = num * zei / 100;
            return zeinuki;
        }

        /// <summary>
        /// フォルダ存在チェック　フォルダ画像切り替え
        /// 　存在あり：黄色
        /// 　存在なし：灰色
        /// </summary>
        private void FolderPathCheck()
        {
            // 契約図書
            if (Directory.Exists(ca_tbl01_txtTosyo.Text))
            {
                ca_tbl01_picTosyo.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
            }
            else
            {
                ca_tbl01_picTosyo.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
            }
            // 請求書
            if (Directory.Exists(te_txtSeikyusyo.Text))
            {
                te_picSeikyusyo.Image = Image.FromFile("Resource/Image/folder_yellow_s.png");
            }
            else
            {
                te_picSeikyusyo.Image = Image.FromFile("Resource/Image/folder_gray_s.png");
            }

        }

        /// <summary>
        /// DateTimePickerの値取得
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 業務区分コード取得
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        private string Get_GyoumuKubunCD(string ID)
        {
            string cd = "";
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

        #endregion

        #region チェック処理 Private -------------------------------------------------
        /// <summary>
        /// 入力チェック
        /// </summary>
        /// <param name="flg">0:新規登録 1:更新 2:チェック用出力(赤伝・黒伝) 3:起案</param>
        /// <returns></returns>
        private bool ErrorFLG(int flg)
        {
            // エラーフラグ true：エラー、false：正常
            bool isError = false;

            // 起案の場合は、エラーチェックを行う
            bool isCheckCa = false;
            if (flg >= 3)
            {
                isCheckCa = true;
            }
            else
            {
                // ①入札タブ：入札状況が入札成立となったら入札のチェックを行う（更新可能）
                // ②契約タブ：調査会様での入札が成立したら契約のチェック処理を行う（更新可能）
                // ENTORY_TOUKAI:建設物価調査会
                string sEntoryToukai = GlobalMethod.GetCommonValue2("ENTORY_TOUKAI");
                isCheckCa = sEntoryToukai.Equals(bid_tbl03_1_txtRakusatuSya.Text);
            }

            set_error("", 0);

            //入力不正背景色をクリアする
            clearBackColor(flg);

            //===================================================================================================
            // 新規登録時のチェック
            //===================================================================================================
            if (flg == 0)
            {
                // 基本情報等一覧タブの必須チェック
                if (baseRequireCheck())
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "基本情報等一覧"));
                    isError = true;
                }

                // 基本情報等一覧タブのデータチェック
                if (baseDataCheck(0))
                {
                    isError = true;
                }
            }

            //===================================================================================================
            // 更新、起案時のチェック
            //===================================================================================================
            if (flg == 1)
            {
                // 基本情報等一覧タブの必須チェック
                if (baseRequireCheck(1))
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "基本情報等一覧"));
                    isError = true;
                }

                // 事前打診タブの必須チェック
                if (priorRequireCheck(1))
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "事前打診"));
                    isError = true;
                }

                // 入札タブの必須チェック
                if (bidRequireCheck(1))
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "入札"));
                    isError = true;
                }

                // 契約タブの必須チェック（調査会様での入札が成立時）
                if (isCheckCa && ca_tbl01_chkKian.Checked)
                {
                    if (caRequireCheck())
                    {
                        // 起案済みの場合はエラーとする
                        // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                        set_error(GlobalMethod.GetMessage("E10010", "契約"));
                        isError = true;
                    }
                }

                // 基本情報等一覧タブのデータチェック
                if (baseDataCheck(1))
                {
                    isError = true;
                }

                // 事前打診タブのデータチェック
                if (priorDataCheck(1))
                {
                    isError = true;
                }
                // 入札タブのデータチェック
                if (bidDataCheck())
                {
                    isError = true;
                }

                // 契約タブのデータチェック（調査会様での入札が成立時）
                if (isCheckCa)
                {
                    if (caDataCheck())
                    {
                        isError = true;
                    }
                }

                // 技術者評価タブのデータチェック
                if (teDataCheck())
                {
                    isError = true;
                }
            }

            //===================================================================================================
            // チェック用帳票出力時のチェック
            //===================================================================================================
            if (flg == 2)
            {
                //基本情報等一覧タブの必須チェック
                if (baseRequireCheck(flg))
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "基本情報等一覧"));
                    isError = true;
                }

                // 契約タブの必須チェック（調査会様での入札が成立時）
                // 起案済みの場合はエラーとする
                if (isCheckCa && ca_tbl01_chkKian.Checked == true)
                {
                    if (caRequireCheck())
                    {
                            // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                            set_error(GlobalMethod.GetMessage("E10010", "契約"));
                            isError = true;
                    }
                }

                // 基本情報等一覧タブのデータチェック
                if (baseDataCheck())
                {
                    isError = true;
                }

                // 入札タブのデータチェック
                // 業務配分と業務別配分のチェックだけのため、条件判定しないようにコメント化
                if (bidDataCheck())
                {
                    isError = true;
                }
                //}

                // 契約タブのデータチェック（調査会様での入札が成立時）
                if (isCheckCa)
                {
                    if (caDataCheck())
                    {
                        isError = true;
                    }
                }

                // 技術者評価タブのデータチェック
                if (teDataCheck())
                {
                    isError = true;
                }

                //============================================================================
                // 起案用のチェック？（できれば外出ししてまとめたい）
                //============================================================================
                Double totalHundred = Convert.ToDouble(100);

                // 契約タブ STEP3で入力不可ので、チェックなしでOK
                // 売上年度 4桁じゃなかったらエラー
                //if (4 != ca_tbl01_cmbSalesYear.SelectedValue.ToString().Length)
                //{
                //    set_error(GlobalMethod.GetMessage("E10011", ""));
                //}

                // 入札タブ
                // 入札状況が入札成立でなければ起案エラー
                if (!isNyuusatsu_seiritsu(bid_tbl03_1_cmbBidStatus.SelectedValue.ToString()))
                {
                    set_error(GlobalMethod.GetMessage("E10702", ""));
                }

                // 落札者が建設物価調査会でなければ起案エラー
                if (!GlobalMethod.GetCommonValue2("ENTORY_TOUKAI").Equals(bid_tbl03_1_txtRakusatuSya.Text))
                {
                    set_error(GlobalMethod.GetMessage("E70048", ""));
                }

                // 契約タブ
                // 契約タブの1.契約情報の契約金額の税込が0円の場合
                if (isCheckCa)
                {
                    long item13 = GetLong(ca_tbl01_txtZeikomiAmt.Text);
                    if (item13 == 0)
                    {
                        // 0円起案です。
                        set_error(GlobalMethod.GetMessage("W10701", ""));
                    }
                }

                // 税込
                // 契約タブの1.契約情報の消費税率が空ではない場合
                if (!String.IsNullOrEmpty(ca_tbl01_txtTax.Text))
                {
                    Double keiyakuAmount = GetDouble(ca_tbl01_txtZeikomiAmt.Text);  // 契約金額の税込
                    Double inTaxAmount = GetDouble(ca_tbl01_txtSyohizeiAmt.Text);    // 内消費税
                    Double taxPercent = Double.Parse(ca_tbl01_txtTax.Text);  // 消費税率

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
                Double jutakuTax = GetDouble(ca_tbl01_txtJyutakuAmt.Text);      // 1.契約情報の受託金額(税込)
                Double totalAmount = GetDouble(ca_tbl02_AftCaBmZeikomi_numAmtAll.Text);   // 2.配分情報の配分情報の配分額(税込)の合計

                //受託金額(税込)と配分額(税込)の合計が一致しない
                if (!Double.Equals(jutakuTax, totalAmount))
                {
                    // 起案は出来ますが、受託契約金額と各配分額の合計が一致していません。確認して下さい。
                    set_error(GlobalMethod.GetMessage("E10705", ""));
                }

                if (isCheckCa)
                {
                    // 契約タブ
                    // 契約工期至と売上年度のチェック
                    String format = "yyyy/MM/dd";   // 日付フォーマット

                    // 1.契約情報の契約工期至と売上年度が空でない場合
                    if (ca_tbl01_dtpKokiTo.CustomFormat == "" && !String.IsNullOrEmpty(ca_tbl01_cmbSalesYear.Text))
                    {
                        // 売上年度 +1年 の3月31日
                        int year = Int32.Parse(ca_tbl01_cmbSalesYear.SelectedValue.ToString()) + 1;
                        String date = year + "/03/31";

                        // 日付型
                        DateTime nextYear = DateTime.ParseExact(date, format, null);
                        DateTime keiyaku = DateTime.ParseExact(ca_tbl01_dtpKokiTo.Text, format, null);

                        // 売上年度+1/03/31よりも、契約工期の完了日が未来日付の場合エラー
                        if (nextYear.Date < keiyaku.Date)
                        {
                            // 工期完了日が売上年度を超えています。年度をまたぐ場合は、売上年度を工期完了日にあわせてください。
                            set_error(GlobalMethod.GetMessage("E10706", ""));
                        }
                    }

                    // 基本情報一覧タブ
                    // 調査部　売上計上情報と業務別配分のチェック
                    // 7.業務内容の調査部 業務別配分の合計が100の場合
                    if (GetDouble(base_tbl07_2_numPercentAll.Text).ToString("F2") == "100.00")
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
                            long haibunTax = GetLong(ca_tbl02_AftCaBmZeikomi_numAmt1.Text);

                            // 配分額(税込)と、計上額の合計のチェック
                            if (!long.Equals(haibunTax, keijoTotal))
                            {
                                set_error(GlobalMethod.GetMessage("E10717", "(調査部)"));
                            }

                        }
                    }

                    // 事業普及部　売上計上情報と業務別配分のチェック
                    // 7.業務内容の事業普及部 業務別配分の合計が100の場合
                    if (GetDouble(base_tbl07_1_numPercent2.Text).ToString("F2") == "100.00")
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
                            long haibunTax = GetLong(ca_tbl02_AftCaBmZeikomi_numAmt2.Text);

                            // 配分額(税込)と、計上額の合計のチェック
                            if (!long.Equals(haibunTax, keijoTotal))
                            {
                                set_error(GlobalMethod.GetMessage("E10717", "(事業普及部)"));
                            }

                        }
                    }

                    // 情報システム部　売上計上情報と業務別配分のチェック
                    // 7.業務内容の情報システム部 業務別配分の合計が100の場合
                    if (GetDouble(base_tbl07_1_numPercent3.Text).ToString("F2") == "100.00")
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
                            long haibunTax = GetLong(ca_tbl02_AftCaBmZeikomi_numAmt3.Text);

                            // 配分額(税込)と、計上額の合計のチェック
                            if (!long.Equals(haibunTax, keijoTotal))
                            {
                                set_error(GlobalMethod.GetMessage("E10717", "(情報システム部)"));
                            }

                        }
                    }

                    // 総合研究所　売上計上情報と業務別配分のチェック
                    // 7.業務内容の総合研究所 業務別配分の合計が100の場合
                    if (GetDouble(base_tbl07_1_numPercent4.Text).ToString("F2") == "100.00")
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
                            long haibunTax = GetLong(ca_tbl02_AftCaBmZeikomi_numAmt4.Text);

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
                // 基本情報等一覧：タブの必須チェック
                if (baseRequireCheck(3))
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "引合"));
                    isError = true;
                }

                // 契約タブの必須チェック
                if (caRequireCheck())
                {
                    // E10010:必須入力項目が未入力です。赤背景の項目を入力して下さい。
                    set_error(GlobalMethod.GetMessage("E10010", "契約"));
                    isError = true;
                }

                // 基本情報等一覧タブのデータチェック
                if (baseDataCheck())
                {
                    isError = true;
                }

            }

            // チェック処理を実施した結果、更新可の場合はfalse：正常で返す
            return isError;
        }

        /// <summary>
        /// 売上計上情報から指定された部所の計上額の合計を取得する
        /// </summary>
        /// <param name="targetBusho"></param>
        /// <returns></returns>
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
            C1FlexGrid c1FlexGrid4 = ca_tbl06_c1FlexGrid;
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

        /// <summary>
        /// 基本情報一覧：必須入力チェック
        /// </summary>
        /// <param name="flg">0:新規登録 1:更新</param>
        /// <returns></returns>
        private bool baseRequireCheck(int flg = 0)
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;
            object obj = null;
            Color errorColor = Color.FromArgb(255, 204, 255);

            if (flg == 0) {
                // １．進捗段階	全て	事前打診登録日、入札情報登録日が無ければ登録不可。入札情報登録日が登録された場合、自動的に事前打診登録日を設定
                // 事前打診登録日、入札情報登録日　いずれか設定する
                if (base_tbl01_dtpDtPrior.CustomFormat != "" && base_tbl01_dtpDtBid.CustomFormat != "")
                {
                    errorFlg = true;
                    base_tbl01_lblDtPrior.BackColor = errorColor;
                    //No1692　入札の登録日時背景色は変えないよう修正
                    //base_tbl01_lblDtBid.BackColor = errorColor;
                    base_tbl01_picPriorAlert.Visible = true;
                    base_tbl01_picBidAlert.Visible = true;
                }
            }
            // ２．基本情報	全て	計画番号以外が無ければ登録不可。   No.1457でAdd、No.1671でDel
            //// 計画番号
            //if (string.IsNullOrEmpty(base_tbl02_txtKeikakuNo.Text))
            //{
            //    errorFlg = true;
            //    base_tbl02_txtKeikakuNo.BackColor = errorColor;
            //    base_tbl02_picKeikakuNoAlert.Visible = true;
            //}
            // 受託課所支部
            if (String.IsNullOrEmpty(base_tbl02_cmbJyutakuKasyoSibu.Text))
            {
                errorFlg = true;
                base_tbl02_lblJyutakuKasyoSibu.BackColor = errorColor;
                base_tbl02_picJyutakuKasyoSibuAlert.Visible = true;
            }
            // 契約担当者
            if (String.IsNullOrEmpty(base_tbl02_txtKeiyakuTanto.Text))
            {
                errorFlg = true;
                base_tbl02_txtKeiyakuTanto.BackColor = errorColor;
                base_tbl02_picKeiyakuTantoAlert.Visible = true;
            }
            // 案件(受託)フォルダ
            if (String.IsNullOrEmpty(base_tbl02_txtAnkenFolder.Text))
            {
                errorFlg = true;
                base_tbl02_txtAnkenFolder.BackColor = errorColor;
                base_tbl02_picAnkenFolderAlert.Visible = true;
            }
            // ３．案件情報	全て	工期開始自、工期開始至、案件メモ(基本情報)以外が無ければ登録不可。契約区分は空欄表示、空欄の場合登録不可。
            // 業務名称
            if (String.IsNullOrEmpty(base_tbl03_txtGyomuName.Text))
            {
                errorFlg = true;
                base_tbl03_txtGyomuName.BackColor = errorColor;
                base_tbl03_picGyomuNameAlert.Visible = true;
            }
            // 契約区分
            if (String.IsNullOrEmpty(base_tbl03_cmbKeiyakuKubun.Text))
            {
                errorFlg = true;
                base_tbl03_lblKeiyakuKubun.BackColor = errorColor;
                base_tbl03_picKeiyakuKubunAlert.Visible = true;
            }
            if (flg == 0 || flg == 1)
            {
                //工期自
                if (base_tbl03_dtpKokiFrom.CustomFormat != "")
                {
                    base_tbl03_lblKokiFrom.BackColor = errorColor;
                    base_tbl03_picKokiFromAlert.Visible = true;
                    errorFlg = true;
                }

                //工期至
                if (base_tbl03_dtpKokiTo.CustomFormat != "")
                {
                    base_tbl03_lblKokiTo.BackColor = errorColor;
                    base_tbl03_picKokiToAlert.Visible = true;
                    errorFlg = true;
                }
            }

            // 	４．発注者情報	全て	発注者課名以外が無ければ登録不可。
            //発注者コード
            if (String.IsNullOrEmpty(base_tbl04_txtOrderCd.Text))
            {
                errorFlg = true;
                base_tbl04_txtOrderCd.BackColor = errorColor;
                base_tbl04_picOrderCdAlert.Visible = true;
            }

            //発注者区分1
            if (String.IsNullOrEmpty(base_tbl04_txtOrderKubun1.Text))
            {
                errorFlg = true;
                base_tbl04_txtOrderKubun1.BackColor = errorColor;
                base_tbl04_picOrderKubun1Alert.Visible = true;
            }

            //発注者区分2
            if (String.IsNullOrEmpty(base_tbl04_txtOrderKubun2.Text))
            {
                errorFlg = true;
                base_tbl04_txtOrderKubun2.BackColor = errorColor;
                base_tbl04_picOrderKubun2Alert.Visible = true;
            }

            //No1533 更新と新規で、部所設定が必須になっている。エラーを外さないと運用が回らない。
            //// ７．業務配分	全て	応援依頼の有無、応援依頼メモ、応援依頼先以外が無ければ登録不可、100%以外で登録不可。
            //// 　　ただし、部門配分の調査部配分は0以上（0を含めない）の場合、応援依頼の有無、応援依頼メモ、応援依頼先がなければ登録不可 (No.1458)
            //if (GetDouble(base_tbl07_1_numPercent1.Text) > 0)
            //{
            //    // No.1533 削除
            //    //if (String.IsNullOrEmpty(base_tbl07_3_cmbOen.Text))
            //    //{
            //    //    errorFlg = true;
            //    //    base_tbl07_3_lblOen.BackColor = errorColor;
            //    //    base_tbl07_3_picOenAlert.Visible = true;
            //    //}
            //    // No.1533 削除
            //    //// No.1491 エントリくんの新規登録・更新で、基本情報の応援依頼先が【なし】の場合は、「応援依頼メモ」「応援依頼先」が空欄でも良いが、エラーとなってしまう。
            //    //if (!IsSpecifiedValue(base_tbl07_3_cmbOen.SelectedValue, "2"))
            //    //{
            //        //No.1521
            //        //if (String.IsNullOrEmpty(base_tbl07_3_txtOenMemo.Text))
            //        //{
            //        //    errorFlg = true;
            //        //    base_tbl07_3_txtOenMemo.BackColor = errorColor;
            //        //    base_tbl07_3_picOenMemoAlert.Visible = true;
            //        //}
            //    bool bOen1 = false;
            //    bool bOen2 = false;
            //    bool bOen3 = false;
            //    if (base_tbl07_3_tblOenIrai1.Height > 0)
            //    {
            //        foreach (Control child in base_tbl07_3_tblOenIrai1.Controls)
            //        {
            //            //特定のコントロール型内部の子情報は取得しない
            //            if (child is System.Windows.Forms.CheckBox)
            //            {
            //                if (((System.Windows.Forms.CheckBox)child).Checked)
            //                {
            //                    bOen1 = true;
            //                    break;
            //                }
            //            }
            //        }

            //        if (!bOen1 && base_tbl07_3_tblOenIrai2.Height > 0)
            //        {
            //            foreach (Control child in base_tbl07_3_tblOenIrai2.Controls)
            //            {
            //                //特定のコントロール型内部の子情報は取得しない
            //                if (child is System.Windows.Forms.CheckBox)
            //                {
            //                    if (((System.Windows.Forms.CheckBox)child).Checked)
            //                    {
            //                        bOen2 = true;
            //                        break;
            //                    }
            //                }
            //            }
            //        }

            //        if (!bOen1 && !bOen2 && base_tbl07_3_tblOenIrai3.Height > 0)
            //        {
            //            foreach (Control child in base_tbl07_3_tblOenIrai3.Controls)
            //            {
            //                //特定のコントロール型内部の子情報は取得しない
            //                if (child is System.Windows.Forms.CheckBox)
            //                {
            //                    if (((System.Windows.Forms.CheckBox)child).Checked)
            //                    {
            //                        bOen3 = true;
            //                        break;
            //                    }
            //                }
            //            }
            //        }
            //    }
            //    if (!(bOen1 || bOen2 || bOen3))
            //    {
            //        base_tbl07_3_lblOenIrai.BackColor = errorColor;
            //        base_tbl07_3_picOenIraiAlert.Visible = true;
            //        errorFlg = true;
            //    }
            //    //}
            //}

            if (flg == 0)
            {
                // ９．事前打診・参考見積	事前打診登録日が設定された場合、未設定の場合登録不可。
                if (base_tbl01_dtpDtPrior.CustomFormat == "")
                {
                    // 事前打診依頼日
                    if (base_tbl09_dtpJizenDasinIraiDt.CustomFormat != "")
                    {
                        base_tbl09_lblJizenDasinIraiDt.BackColor = errorColor;
                        base_tbl09_picJizenDasinIraiDtAlert.Visible = true;
                        errorFlg = true;
                    }
                    // 参考見積対応
                    if (this.IsNotSelected(base_tbl09_cmbSankomitumori))
                    {
                        base_tbl09_lblSankomitumori.BackColor = errorColor;
                        base_tbl09_picSankomitumoriAlert.Visible = true;
                        errorFlg = true;
                    }

                    // 発注予定・見込日
                    if (base_tbl09_dtpOrderYoteiDt.CustomFormat != "")
                    {
                        base_tbl09_lblOrderYoteiDt.BackColor = errorColor;
                        base_tbl09_picOrderYoteiDtAlert.Visible = true;
                        errorFlg = true;
                    }
                    //// 未発注状況 No.1473
                    //if (this.IsNotSelected(base_tbl09_cmbNotOrderStats))
                    //{
                    //    base_tbl09_lblNotOrderStats.BackColor = errorColor;
                    //    base_tbl09_picNotOrderStatsAlert.Visible = true;
                    //    errorFlg = true;
                    //}
                    //「発注無し」の理由 未発注状況が発注無しの場合、空欄の場合登録不可。
                    if (this.IsSpecifiedValue(base_tbl09_cmbNotOrderStats.SelectedValue, "1"))
                    {
                        if (this.IsNotSelected(base_tbl09_cmbNotOrderReason))
                        {
                            base_tbl09_lblNotOrderReason.BackColor = errorColor;
                            base_tbl09_picNotOrderReasonAlert.Visible = true;
                            errorFlg = true;
                        }
                    }

                    //「その他」の内容 発注無しの理由がその他の場合、空欄の場合、登録不可
                    if (IsSpecifiedValue(base_tbl09_cmbNotOrderReason.SelectedValue, "4"))
                    {
                        if (string.IsNullOrEmpty(base_tbl09_txtOthenComment.Text))
                        {
                            base_tbl09_txtOthenComment.BackColor = errorColor;
                            base_tbl09_picOthenCommentAlert.Visible = true;
                            errorFlg = true;
                        }
                    }

                    // 受注意欲
                    if (this.IsNotSelected(base_tbl09_cmbOrderIyoku))
                    {
                        base_tbl09_lblOrderIyoku.BackColor = errorColor;
                        base_tbl09_picOrderIyokuAlert.Visible = true;
                        errorFlg = true;
                    }
                }

                // １０．入札状況・入札結果 、入札情報登録日が設定された場合、未設定の場合登録不可。
                if (base_tbl01_dtpDtBid.CustomFormat == "")
                {
                    // 業務発注区分
                    if (this.IsNotSelected(base_tbl10_cmbOrderKubun))
                    {
                        base_tbl10_lblOrderKubun.BackColor = errorColor;
                        base_tbl10_picOrderKubunAlert.Visible = true;
                        errorFlg = true;
                    }

                    // 入札方式
                    if (this.IsNotSelected(base_tbl10_cmbNyusatuHosiki))
                    {
                        base_tbl10_lblNyusatuHosiki.BackColor = errorColor;
                        base_tbl10_picNyusatuHosikiAlert.Visible = true;
                        errorFlg = true;
                    }
                    // 最低制限価格有無
                    if (this.IsNotSelected(base_tbl10_cmbLowestUmu))
                    {
                        base_tbl10_lblLowestUmu.BackColor = errorColor;
                        base_tbl10_picLowestUmuAlert.Visible = true;
                        errorFlg = true;
                    }
                    // 入札(予定)日
                    if (base_tbl10_dtpNyusatuDt.CustomFormat != "")
                    {
                        base_tbl10_lblNyusatuDt.BackColor = errorColor;
                        base_tbl10_picNyusatuDtAlert.Visible = true;
                        errorFlg = true;
                    }

                    // 参考見積対応
                    if (this.IsNotSelected(base_tbl10_cmbSankoMitumori))
                    {
                        base_tbl10_lblSankoMitumori.BackColor = errorColor;
                        base_tbl10_picSankoMitumoriAlert.Visible = true;
                        errorFlg = true;
                    }

                    // 受注意欲
                    if (this.IsNotSelected(base_tbl10_cmbOrderIyoku))
                    {
                        base_tbl10_lblOrderIyoku.BackColor = errorColor;
                        base_tbl10_picOrderIyokuAlert.Visible = true;
                        errorFlg = true;
                    }

                    // 当会応札
                    if (this.IsNotSelected(base_tbl10_cmbTokaiOsatu))
                    {
                        base_tbl10_lblTokaiOsatu.BackColor = errorColor;
                        base_tbl10_picTokaiOsatuAlert.Visible = true;
                        errorFlg = true;
                    }

                    // 再委託禁止条項の記載有無
                    if (this.IsNotSelected(base_tbl10_cmbKinsiUmu))
                    {
                        base_tbl10_lblKinsiUmu.BackColor = errorColor;
                        base_tbl10_picKinsiUmuAlert.Visible = true;
                        errorFlg = true;
                    }

                    // No1588　1319　新規登録時、再委託禁止条項の記載有無を「なし」/「不明」に設定しても、再委託禁止条項の内容が空欄だとエラーになる。
                    if ((this.IsSpecifiedValue(base_tbl10_cmbKinsiUmu.SelectedValue, "2") || this.IsSpecifiedValue(base_tbl10_cmbKinsiUmu.SelectedValue, "3")) == false)
                    {
                        // 再委託禁止条項の内容
                        if (this.IsNotSelected(base_tbl10_cmbKinsiNaiyo))
                        {
                            base_tbl10_lblKinsiNaiyo.BackColor = errorColor;
                            base_tbl10_picKinsiNaiyoAlert.Visible = true;
                            errorFlg = true;
                        }
                    }

                    // 入札状況
                    if (this.IsNotSelected(base_tbl10_cmbNyusatuStats))
                    {
                        base_tbl10_lblNyusatuStats.BackColor = errorColor;
                        base_tbl10_picNyusatuStatsAlert.Visible = true;
                        errorFlg = true;
                    }

                    // No.1534 落札者状況と落札額状況は空欄でもエラーにしないようにする
                    //// 落札者状況
                    //if (this.IsNotSelected(base_tbl10_cmbRakusatuStats))
                    //{
                    //    base_tbl10_lblRakusatuStats.BackColor = errorColor;
                    //    base_tbl10_picRakusatuStatsAlert.Visible = true;
                    //    errorFlg = true;
                    //}

                    //// 落札額状況
                    //if (this.IsNotSelected(base_tbl10_cmbRakusatuAmtStats))
                    //{
                    //    base_tbl10_lblRakusatuAmtStats.BackColor = errorColor;
                    //    base_tbl10_picRakusatuAmtStatsAlert.Visible = true;
                    //    errorFlg = true;
                    //}
                }
                //else
                //{
                //    // 参考見積対応
                //    if (this.IsSpecifiedValue(base_tbl10_cmbSankoMitumori.SelectedValue, "1"))
                //    {
                //        base_tbl01_lblDtBid.BackColor = errorColor;
                //        base_tbl01_picBidAlert.Visible = true;
                //        errorFlg = true;
                //    }
                //}
            }
            return errorFlg;
        }


        /// <summary>
        /// 事前打診：必須入力チェック
        /// </summary>
        /// <param name="flg">0:新規登録 1:更新</param>
        /// <returns></returns>
        private bool priorRequireCheck(int flg)
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;
            object obj = null;
            Color errorColor = Color.FromArgb(255, 204, 255);

            // 事前打診　---------------------------------------------------
            // １．事前打診状況	事前打診登録日が設定された場合、未設定の場合登録不可。
            if (base_tbl01_dtpDtPrior.CustomFormat == "")
            {
                // 事前打診依頼日 事前打診登録日が設定された場合、未設定の場合登録不可。
                if (prior_tbl01_dtpDasinIraiDt.CustomFormat != "")
                {
                    errorFlg = true;
                    prior_tbl01_lblDasinIraiDt.BackColor = errorColor;
                    prior_tbl01_picDasinIraiDtAlert.Visible = true;
                }

                // 参考見積対応 事前打診登録日が設定された場合、未設定の場合登録不可。
                if (this.IsNotSelected(prior_tbl01_cmbMitumori))
                {
                    prior_tbl01_lblMitumori.BackColor = errorColor;
                    prior_tbl01_picMitumoriAlert.Visible = true;
                    errorFlg = true;
                }

                // 受注意欲がいない＝空欄でも更新が出来てしまう。空欄の場合、更新エラーが正しい。 No.1468
                if (this.IsNotSelected(prior_tbl01_cmbOrderIyoku))
                {
                    errorFlg = true;
                    prior_tbl01_lblOrderIyoku.BackColor = errorColor;
                    prior_tbl01_picOrderIyokuAlert.Visible = true;
                }

                // １．事前打診状況    「発注なし」の理由 未発注状況が発注無しの場合、「発注なし」の理由が選択されていない場合はエラー。No.1474
                if (this.IsSpecifiedValue(prior_tbl02_cmbNotOrderStats.SelectedValue, "1"))
                {
                    if (this.IsNotSelected(prior_tbl02_cmbNotOrderReason))
                    {
                        prior_tbl02_lblNotOrderReason.BackColor = errorColor;
                        prior_tbl02_picNotOrderReasonAlert.Visible = true;
                        errorFlg = true;
                    }
                }

                //// No.1537　1285　更新時に警告にならない。
                //// 事前打診	２．未発注	その他の内容	発注無しの理由「その他」の時、入力欄色付け。更新時警告。→エラーにするか？→警告にする。
                //if (IsSpecifiedValue(prior_tbl02_cmbNotOrderReason.SelectedValue, "4"))
                //{
                //    if (string.IsNullOrEmpty(prior_tbl02_txtOtherNaiyo.Text))
                //    {
                //        prior_tbl02_txtOtherNaiyo.BackColor = errorColor;
                //        prior_tbl02_picOtherNaiyoAlert.Visible = true;
                //        errorFlg = true;
                //    }
                //}
            }

            
            return errorFlg;
        }


        /// <summary>
        /// 入札：必須入力チェック
        /// </summary>
        /// <param name="flg">0:新規登録 1:更新</param>
        /// <returns></returns>
        private bool bidRequireCheck(int flg)
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;
            object obj = null;
            Color errorColor = Color.FromArgb(255, 204, 255);
            // 入札　-------------------------------------------------------
            if (base_tbl01_dtpDtBid.CustomFormat == "")
            {
                // １．入札情報 全て  未登録があれば更新不可。
                //入札情報登録日
                if (bid_tbl01_dtpBidInfoDt.CustomFormat != "")
                {
                    errorFlg = true;
                    bid_tbl01_lblBidInfoDt.BackColor = errorColor;
                    bid_tbl01_picBidInfoDtAlert.Visible = true;
                }
                //業務発注区分
                if (this.IsNotSelected(bid_tbl01_cmbOrderKubun))
                {
                    bid_tbl01_lblOrderKubun.BackColor = errorColor;
                    bid_tbl01_picOrderKubunAlert.Visible = true;
                    errorFlg = true;
                }
                //入札方式
                if (this.IsNotSelected(bid_tbl01_cmbBidhosiki))
                {
                    bid_tbl01_lblBidhosiki.BackColor = errorColor;
                    bid_tbl01_picBidhosikiAlert.Visible = true;
                    errorFlg = true;
                }
                //最低制限価格有無
                if (this.IsNotSelected(bid_tbl01_cmbLowestUmu))
                {
                    bid_tbl01_lblLowestUmu.BackColor = errorColor;
                    bid_tbl01_picLowestUmuAlert.Visible = true;
                    errorFlg = true;
                }
                //入札(予定)日
                if (bid_tbl01_dtpBidYoteiDt.CustomFormat != "")
                {
                    errorFlg = true;
                    bid_tbl01_lblBidYoteiDt.BackColor = errorColor;
                    bid_tbl01_picBidYoteiDtAlert.Visible = true;
                }
                //参考見積対応
                if (this.IsNotSelected(bid_tbl01_cmbMitumori))
                {
                    bid_tbl01_lblMitumori.BackColor = errorColor;
                    bid_tbl01_picMitumoriAlert.Visible = true;
                    errorFlg = true;
                }
                //参考見積額(税抜)  ★★★要らない気がする
                if (string.IsNullOrEmpty(bid_tbl01_txtMitumoriAmt.Text))
                {
                    bid_tbl01_txtMitumoriAmt.BackColor = errorColor;
                    bid_tbl01_picMitumoriAmtAlert.Visible = true;
                    errorFlg = true;
                }
                //受注意欲
                if (this.IsNotSelected(bid_tbl01_cmbOrderIyoku))
                {
                    bid_tbl01_lblOrderIyoku.BackColor = errorColor;
                    bid_tbl01_picOrderIyokuAlert.Visible = true;
                    errorFlg = true;
                }
                //当会応札
                if (this.IsNotSelected(bid_tbl01_cmbTokaiOsatu))
                {
                    bid_tbl01_lblTokaiOsatu.BackColor = errorColor;
                    bid_tbl01_picTokaiOsatuAlert.Visible = true;
                    errorFlg = true;
                }
                // ３．入札結果	全て	案件メモ以外の未登録があれば登録不可。
                //入札結果登録日　No.1522（1268）　入札結果登録日の更新時のエラーから警告に変更する。
                //if (!IsSpecifiedValue(bid_tbl03_1_cmbBidStatus.SelectedValue, "1") && bid_tbl03_1_dtpBidResultDt.CustomFormat != "")
                //{
                //    errorFlg = true;
                //    bid_tbl03_1_lblBidResultDt.BackColor = errorColor;
                //    bid_tbl03_1_picBidResultDtAlert.Visible = true;
                //}
                // 入札結果登録日 // No 1587　進捗段階　入札の際、入札画面の入札結果登録日が空欄でも登録が出来る。入札成立～何某の場合は、入力エラーで良い。
                if (bid_tbl03_1_dtpBidResultDt.CustomFormat != "")
                {
                    errorFlg = true;
                    bid_tbl03_1_lblBidResultDt.BackColor = errorColor;
                    bid_tbl03_1_picBidResultDtAlert.Visible = true;
                }
                //入札状況
                if (this.IsNotSelected(bid_tbl03_1_cmbBidStatus))
                {
                    bid_tbl03_1_lblBidStatus.BackColor = errorColor;
                    bid_tbl03_1_picBidStatusAlert.Visible = true;
                    errorFlg = true;
                }
                //予定価格(税抜)
                if (string.IsNullOrEmpty(bid_tbl03_1_txtYoteiPrice.Text))
                {
                    bid_tbl03_1_txtYoteiPrice.BackColor = errorColor;
                    bid_tbl03_1_picYoteiPriceAlert.Visible = true;
                    errorFlg = true;
                }
                //応札数
                if (string.IsNullOrEmpty(bid_tbl03_1_txtOsatuNum.Text))
                {
                    bid_tbl03_1_txtOsatuNum.BackColor = errorColor;
                    bid_tbl03_1_picOsatuNumAlert.Visible = true;
                    errorFlg = true;
                }
                // No.1534 警告に変更する
                ////落札者状況
                //if (this.IsNotSelected(bid_tbl03_1_cmbRakusatuStatus))
                //{
                //    bid_tbl03_1_lblRakusatuStatus.BackColor = errorColor;
                //    bid_tbl03_1_picRakusatuStatusAlert.Visible = true;
                //    errorFlg = true;
                //}
                ////落札額状況
                //if (this.IsNotSelected(bid_tbl03_1_cmbRakusatuAmtStatus))
                //{
                //    bid_tbl03_1_lblRakusatuAmtStatus.BackColor = errorColor;
                //    bid_tbl03_1_picRakusatuAmtStatusAlert.Visible = true;
                //    errorFlg = true;
                //}
            }

            return errorFlg;
        }

        /// <summary>
        /// 契約：必須入力チェック
        /// </summary>
        /// <returns></returns>
        private bool caRequireCheck()
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;
            object obj = null;
            Color errorColor = Color.FromArgb(255, 204, 255);

            //契約締結(変更)日
            if (ca_tbl01_dtpChangeDt.CustomFormat != "")
            {
                ca_tbl01_lblChangeDt.BackColor = errorColor;
                ca_tbl01_picChangeDtAlert.Visible = true;
                errorFlg = true;
            }

            //起案日
            if (ca_tbl01_dtpKianDt.CustomFormat != "")
            {
                ca_tbl01_lblKianDt.BackColor = errorColor;
                ca_tbl01_picKianDtAlert.Visible = true;
                errorFlg = true;
            }

            //契約工期自
            if (ca_tbl01_dtpKokiFrom.CustomFormat != "")
            {
                ca_tbl01_lblKokiFrom.BackColor = errorColor;
                ca_tbl01_picKokiFromAlert.Visible = true;
                errorFlg = true;
            }

            //契約工期至
            if (ca_tbl01_dtpKokiTo.CustomFormat != "")
            {
                ca_tbl01_lblKokiTo.BackColor = errorColor;
                ca_tbl01_picKokiToAlert.Visible = true;
                errorFlg = true;
            }

            //契約金額の税込
            if (String.IsNullOrEmpty(ca_tbl01_txtZeikomiAmt.Text))
            {
                ca_tbl01_txtZeikomiAmt.BackColor = errorColor;
                ca_tbl01_picZeikomiAmtAlert.Visible = true;
                errorFlg = true;
            }

            //受託金額(税込)
            if (String.IsNullOrEmpty(ca_tbl01_txtJyutakuAmt.Text))
            {
                ca_tbl01_txtJyutakuAmt.BackColor = errorColor;
                ca_tbl01_picJyutakuAmtAlert.Visible = true;
                errorFlg = true;
            }

            //受託外金額(税込)
            if (String.IsNullOrEmpty(ca_tbl01_txtJyutakuGaiAmt.Text))
            {
                ca_tbl01_txtJyutakuGaiAmt.BackColor = errorColor;
                ca_tbl01_picJyutakuGaiAmtAlert.Visible = true;
                errorFlg = true;
            }

            //業務担当者
            if (String.IsNullOrEmpty(ca_tbl05_txtGyomu.Text))
            {
                ca_tbl05_txtGyomu.BackColor = errorColor;
                ca_tbl05_picGyomuAlert.Visible = true;
                errorFlg = true;
            }

            //窓口担当者
            if (String.IsNullOrEmpty(ca_tbl05_txtMadoguchi.Text))
            {
                ca_tbl05_txtMadoguchi.BackColor = errorColor;
                ca_tbl05_picMadoguchiAlert.Visible = true;
                errorFlg = true;
            }

            //エントリ君修正STEP2
            if (ca_tbl01_chkKian.Checked == false)
            {
                //No.1440 ②調査部　業務配分がある場合にエラーを出す
                if (ca_tbl02_1_numPercent1.Text != "0.00%")
                {
                    // 調査部 業務別配分が100でないとエラー
                    if (ca_tbl02_2_numPercentAll.Text != "100.00%")
                    {
                        // 調査業務別　配分の合計が100になるように入力してください。
                        ca_tbl02_2_numPercentAll.BackColor = errorColor;
                        ca_tbl02_lblAftCaRate1.BackColor = errorColor;
                        errorFlg = true;
                    }
                }
                Double total1 = 0;
                Double total2 = 0;
                Double total3 = 0;
                Double total4 = 0;
                for (int i = 2; i < ca_tbl06_c1FlexGrid.Rows.Count; i++)
                {
                    if (ca_tbl06_c1FlexGrid[i, 3] != null) total1 += GetDouble(ca_tbl06_c1FlexGrid[i, 3].ToString());
                    if (ca_tbl06_c1FlexGrid[i, 11] != null) total2 += GetDouble(ca_tbl06_c1FlexGrid[i, 11].ToString());
                    if (ca_tbl06_c1FlexGrid[i, 19] != null) total3 += GetDouble(ca_tbl06_c1FlexGrid[i, 19].ToString());
                    if (ca_tbl06_c1FlexGrid[i, 27] != null) total4 += GetDouble(ca_tbl06_c1FlexGrid[i, 27].ToString());
                }

                //事業部配分の％と事業部の配分金額が異なっていても起案出来てしまう為、エラーとする。
                if (ca_tbl02_1_numPercent1.Text != "0.00%")
                {
                    if (total1 == 0)
                    {
                        ca_tbl06_c1FlexGrid.GetCellRange(2, 3).StyleNew.BackColor = errorColor;
                        errorFlg = true;
                    }
                    if (GetLong(ca_tbl02_AftCaBm_numAmt1.Text) == 0)
                    {
                        ca_tbl02_AftCaBmZeikomi_numAmt1.BackColor = errorColor;
                        errorFlg = true;
                    }
                }
                if (ca_tbl02_1_numPercent2.Text != "0.00%")
                {
                    if (total2 == 0)
                    {
                        ca_tbl06_c1FlexGrid.GetCellRange(2, 11).StyleNew.BackColor = errorColor;
                        errorFlg = true;
                    }

                    if (GetLong(ca_tbl02_AftCaBm_numAmt2.Text) == 0)
                    {
                        ca_tbl02_AftCaBmZeikomi_numAmt2.BackColor = errorColor;
                        errorFlg = true;
                    }
                }
                if (ca_tbl02_1_numPercent3.Text != "0.00%")
                {
                    if (total3 == 0)
                    {
                        ca_tbl06_c1FlexGrid.GetCellRange(2, 19).StyleNew.BackColor = errorColor;
                        errorFlg = true;
                    }

                    if (GetLong(ca_tbl02_AftCaBm_numAmt3.Text) == 0)
                    {
                        ca_tbl02_AftCaBmZeikomi_numAmt3.BackColor = errorColor;
                        errorFlg = true;
                    }
                }
                if (ca_tbl02_1_numPercent4.Text != "0.00%")
                {
                    if (total4 == 0)
                    {
                        ca_tbl06_c1FlexGrid.GetCellRange(2, 27).StyleNew.BackColor = errorColor;
                        errorFlg = true;
                    }
                    if (GetLong(ca_tbl02_AftCaBm_numAmt4.Text) == 0)
                    {
                        ca_tbl02_AftCaBmZeikomi_numAmt4.BackColor = errorColor;
                        errorFlg = true;
                    }
                }
            }
            return errorFlg;
        }

        /// <summary>
        /// コンボボックスが未選択
        /// </summary>
        /// <param name="cmb"></param>
        /// <returns></returns>
        private bool IsNotSelected(System.Windows.Forms.ComboBox cmb)
        {
            object obj = cmb.SelectedValue;
            return (obj == null || string.IsNullOrEmpty(obj.ToString()) || "0".Equals(obj.ToString()));
        }

        /// <summary>
        /// 指定値かチェック
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool IsSpecifiedValue(object obj, string value)
        {
            return (obj != null && value.Equals(obj.ToString()));
        }

        /// <summary>
        /// 基本情報一覧：入力データの正確性チェック
        /// </summary>
        /// <param name="flg">0:新規登録 1:更新</param>
        /// <returns></returns>
        private bool baseDataCheck(int flg = 0)
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;
            Color errorColor = Color.FromArgb(255, 204, 255);

            // １．進捗段階

            // ２．基本情報	

            // ３．案件情報

            // ７．業務配分	全て	100%以外で登録不可。
            double pTotle = 0.00;
            // 部門配分
            pTotle = GetDouble(base_tbl07_1_numPercentAll.Text);
            if (pTotle != 100.00)
            {
                set_error(GlobalMethod.GetMessage("E10733", "基本情報等一覧"));
                errorFlg = true;
                base_tbl07_1_numPercentAll.BackColor = errorColor;
                base_tbl07_1_picPercentAlert.Visible = true;
            }
            if (GetDouble(base_tbl07_1_numPercent1.Text) > 0) { 
                // 業務配分
                pTotle = GetDouble(base_tbl07_2_numPercentAll.Text);
                if (pTotle != 100.00)
                {
                    set_error(GlobalMethod.GetMessage("E10734", "基本情報等一覧"));
                    errorFlg = true;
                    base_tbl07_2_numPercentAll.BackColor = errorColor;
                    base_tbl07_2_picPercentAlert.Visible = true;
                }
            }

            // 新規登録時のみチェックする
            if (flg == 0)
            {
                object obj = null;
                // ９．事前打診・参考見積	事前打診登録日が設定された場合、未設定の場合登録不可。
                if (base_tbl01_dtpDtPrior.CustomFormat == "")
                {
                    // 参考見積対応
                    obj = base_tbl09_cmbSankomitumori.SelectedValue;
                    //「検討中」のまま「入札情報登録日」ありは登録不可。
                    if (base_tbl01_dtpDtBid.CustomFormat == "" && this.IsSpecifiedValue(obj, "1"))
                    {
                        set_error(GlobalMethod.GetMessage("E10735", "基本情報等一覧"));
                        errorFlg = true;
                        base_tbl09_lblSankomitumori.BackColor = errorColor;
                        base_tbl09_picSankomitumoriAlert.Visible = true;
                    }
                }

                // １０．入札状況・入札結果　初期値空欄、入札情報登録日が設定された場合、未設定の場合登録不可。「検討中」のまま「入札情報登録日」が未登録は登録不可
                // 参考見積対応
                obj = base_tbl10_cmbSankoMitumori.SelectedValue;
                //「検討中」のまま「入札情報登録日」が未登録は登録不可
                if (base_tbl01_dtpDtBid.CustomFormat != "" && this.IsSpecifiedValue(obj, "1"))
                {
                    set_error(GlobalMethod.GetMessage("E10736", "基本情報等一覧"));
                    errorFlg = true;
                    base_tbl10_lblSankoMitumori.BackColor = errorColor;
                    base_tbl10_picSankoMitumoriAlert.Visible = true;
                    //No1692　入札の登録日時背景色は変えないよう修正
                    //base_tbl01_lblDtBid.BackColor = errorColor;
                    base_tbl01_picBidAlert.Visible = true;
                }

                // No.1534 入札前も登録できるようにする
                //// 入札状況 「入札中」のまま「入札情報登録日」が登録は登録不可。
                //obj = base_tbl10_cmbNyusatuStats.SelectedValue;
                //if (base_tbl01_dtpDtBid.CustomFormat == "" && IsSpecifiedValue(obj, "1"))
                //{
                //    set_error(GlobalMethod.GetMessage("E10737", "基本情報等一覧"));
                //    base_tbl10_lblNyusatuStats.BackColor = errorColor;
                //    base_tbl10_picNyusatuStatsAlert.Visible = true;
                //    base_tbl01_lblDtBid.BackColor = errorColor;
                //    base_tbl01_picBidAlert.Visible = true;
                //    errorFlg = true;
                //}
            }
            return errorFlg;
        }

        /// <summary>
        /// 事前打診：入力データの正確性チェック
        /// </summary>
        /// <param name="flg">0:新規登録 1:更新</param>
        /// <returns></returns>
        private bool priorDataCheck(int flg = 0)
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;
            Color errorColor = Color.FromArgb(255, 204, 255);
            object obj = prior_tbl01_cmbMitumori.SelectedValue;
            if (this.IsSpecifiedValue(obj, "1") && base_tbl01_dtpDtBid.CustomFormat == "")
            {
                set_error(GlobalMethod.GetMessage("E10735", "基本情報等一覧、事前打診"));
                //No1692　入札の登録日時背景色は変えないよう修正
                //base_tbl01_lblDtBid.BackColor = errorColor;
                base_tbl01_picBidAlert.Visible = true;
                prior_tbl01_lblMitumori.BackColor = errorColor;
                prior_tbl01_picMitumoriAlert.Visible = true;
                errorFlg = true;
            }

            return errorFlg;
        }

        /// <summary>
        /// 入札：入力データの正確性チェック
        /// </summary>
        /// <param name="flg">0:新規登録 1:更新</param>
        /// <returns></returns>
        private bool bidDataCheck(int flg = 0)
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;

            // No1562 1313　進捗段階「入札」にチェックが入っていないが、（入札）タブの、落札者状況が未設定ですと出る。
            if (base_tbl01_dtpDtBid.CustomFormat == "")
            {
                Color errorColor = Color.FromArgb(255, 204, 255);

                // 受注意欲（入札タブ）を「なし」以外にした状態で、参考見積（入札）が「辞退」の場合に更新不可
                bool isErr = IsSpecifiedValue(bid_tbl01_cmbMitumori.SelectedValue, "4");
                if (IsSpecifiedValue(bid_tbl01_cmbOrderIyoku.SelectedValue, "3") == false && isErr)
                {
                    set_error(GlobalMethod.GetMessage("E10730", "入札"));
                    bid_tbl01_lblMitumori.BackColor = errorColor;
                    bid_tbl01_picMitumoriAlert.Visible = true;
                    bid_tbl01_lblOrderIyoku.BackColor = errorColor;
                    bid_tbl01_picOrderIyokuAlert.Visible = true;
                    errorFlg = true;
                }

                // 受注意欲（入札タブ）を「なし」以外にした状態で、当会応札（入札タブ）が「不参加」「辞退」の場合に更新不可
                isErr = (IsSpecifiedValue(bid_tbl01_cmbTokaiOsatu.SelectedValue, "3") || IsSpecifiedValue(bid_tbl01_cmbTokaiOsatu.SelectedValue, "4"));
                if (IsSpecifiedValue(bid_tbl01_cmbOrderIyoku.SelectedValue, "3") == false && isErr)
                {
                    set_error(GlobalMethod.GetMessage("E10738", "入札"));
                    bid_tbl01_lblTokaiOsatu.BackColor = errorColor;
                    bid_tbl01_picTokaiOsatuAlert.Visible = true;
                    bid_tbl01_lblOrderIyoku.BackColor = errorColor;
                    bid_tbl01_picOrderIyokuAlert.Visible = true;
                    errorFlg = true;
                }

                // 「落札者状況 不明」の時、空欄ではないなら更新不可
                if (IsSpecifiedValue(bid_tbl03_1_cmbRakusatuStatus.SelectedValue, "2"))
                {
                    if (string.IsNullOrEmpty(bid_tbl03_1_txtRakusatuSya.Text.Trim()) == false)
                    {
                        set_error(GlobalMethod.GetMessage("E10739", "入札"));
                        bid_tbl03_1_txtRakusatuSya.BackColor = errorColor;
                        bid_tbl03_1_picRakusatuSyaAlert.Visible = true;
                        bid_tbl03_1_lblRakusatuStatus.BackColor = errorColor;
                        bid_tbl03_1_picRakusatuStatusAlert.Visible = true;
                        errorFlg = true;
                    }
                }
            }
            return errorFlg;
        }

        private bool caDataCheck()
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;
            Color errorColor = Color.FromArgb(255, 204, 255);
            // 入札状況
            if (!isNyuusatsu_seiritsu(bid_tbl03_1_cmbBidStatus.SelectedValue.ToString()))
            {
                set_error(GlobalMethod.GetMessage("E10724", ""));
                bid_tbl03_1_lblBidStatus.BackColor = errorColor;
                bid_tbl03_1_picBidStatusAlert.Visible = true;
                errorFlg = true;
            }

            // 事業部コード（案件番号の頭文字1つ）がTのとき
            String jigyoCd = tblAKInfo_lblAnkenNo.Text.Substring(0, 1);
            if ("T".Equals(jigyoCd))
            {
                // 契約図書が空
                if (String.IsNullOrEmpty(ca_tbl01_txtTosyo.Text))
                {
                    set_error(GlobalMethod.GetMessage("W10601", ""));
                }

                // 契約図書のフォーマット：^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$　が違う
                if (!System.Text.RegularExpressions.Regex.IsMatch(ca_tbl01_txtTosyo.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    set_error(GlobalMethod.GetMessage("E10017", ""));
                    //errorFlg = true;
                }
            }

            // 調査部 業務別配分が100でないとエラー
            if (ca_tbl02_AftCaTs_numPercentAll.Text != "100.00%" && ca_tbl02_AftCaTs_numPercentAll.Text != "0.00%")
            {
                // 調査業務別　配分の合計が100になるように入力してください。
                set_error(GlobalMethod.GetMessage("E70045", "契約タブ"));
                ca_tbl02_AftCaTs_numPercentAll.BackColor = errorColor;
                ca_tbl02_AftCaTs_picPercentAlert.Visible = true;
                errorFlg = true;
            }
            return errorFlg;
        }

        /// <summary>
        /// 入札成立か判定
        /// </summary>
        /// <param name="sCompare">比較文字列</param>
        /// <param name="iFlg">0:SelectedValueで判定、1:Textで判定</param>
        /// <returns></returns>
        private bool isNyuusatsu_seiritsu(string sCompare, int iFlg = 0)
        {
            bool isSeiritsu = false;
            string sComValue = iFlg == 0? GlobalMethod.GetCommonValue1("NYUUSATSU_SEIRITSU") : GlobalMethod.GetCommonValue2("NYUUSATSU_SEIRITSU");
            string[] arrComValue = sComValue.Split(',');
            if(arrComValue.Length > 0)
            {
                // データ設定している場合
                if(arrComValue.Contains(sCompare)){
                    isSeiritsu = true;
                }
            }
            return isSeiritsu;
        }
        /// <summary>
        /// 技術者評価：入力データの正確性チェック
        /// </summary>
        /// <param name="flg"></param>
        /// <returns></returns>
        private bool teDataCheck(int flg = 0)
        {
            // エラーフラグ true:エラー /false:正常
            bool errorFlg = false;
            Color errorColor = Color.FromArgb(255, 204, 255);
            // 業務評点
            if (te_txtPoint.Text != "" && (int.Parse(te_txtPoint.Text) < 0 || int.Parse(te_txtPoint.Text) > 100))
            {
                set_error(GlobalMethod.GetMessage("E10913", "業務評点"));
                te_picPointAlert.Visible = true;
                te_txtPoint.BackColor = errorColor;
                errorFlg = true;
            }

            // 管理技術者評点
            if (te_txtKanriPoint.Text != "" && (int.Parse(te_txtKanriPoint.Text) < 0 || int.Parse(te_txtKanriPoint.Text) > 100))
            {
                set_error(GlobalMethod.GetMessage("E10913", "管理技術者評点"));
                te_picKanriPointAlert.Visible = true;
                te_txtKanriPoint.BackColor = errorColor;
                errorFlg = true;
            }

            // 協力担当者評点
            if (te_txtSyosaPoint.Text != "" && (int.Parse(te_txtSyosaPoint.Text) < 0 || int.Parse(te_txtSyosaPoint.Text) > 100))
            {
                set_error(GlobalMethod.GetMessage("E10913", "協力担当者評点"));
                te_picSyosaPointAlert.Visible = true;
                te_txtSyosaPoint.BackColor = errorColor;
                errorFlg = true;
            }

            // 担当技術者評点
            bool isPoint = true;
            for (int i = 1; i < te_c1FlexGrid.Rows.Count; i++)
            {
                if (te_c1FlexGrid.Rows[i][1] != null && te_c1FlexGrid.Rows[i][1].ToString() != "")
                {
                    string Hyouten = "";
                    if (te_c1FlexGrid.Rows[i][3] != null && te_c1FlexGrid.Rows[i][3].ToString() != "")
                    {
                        Hyouten = te_c1FlexGrid.Rows[i][3].ToString();
                    }
                    if (!string.IsNullOrEmpty(Hyouten))
                    {
                        int iHyouten = 0;
                        if(int.TryParse(Hyouten, out iHyouten))
                        {
                            if(iHyouten < 0 || iHyouten > 100)
                            {
                                isPoint = false;
                                te_c1FlexGrid.GetCellRange(i, 3).StyleNew.BackColor = errorColor;
                            }
                        }
                        else
                        {
                            isPoint = false;
                            te_c1FlexGrid.GetCellRange(i, 3).StyleNew.BackColor = errorColor;
                        }
                    }
                }
            }
            if (!isPoint)
            {
                set_error(GlobalMethod.GetMessage("E10913", "担当技術者評点"));
                te_picTantoPointAlert.Visible = true;
                errorFlg = true;
            }

            //請求書のパスのフォーマット：^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$　がちがう
            if (te_txtSeikyusyo.Text != "" && !System.Text.RegularExpressions.Regex.IsMatch(te_txtSeikyusyo.Text, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                set_error(GlobalMethod.GetMessage("E10017", ""));
                te_picSeikyusyoAlert.Visible = true;
                te_txtSeikyusyo.BackColor = errorColor;
                //errorFlg = true;
            }

            return errorFlg;
        }

        /// <summary>
        /// 警告 2023/08/24まで更新時のみアラート
        /// </summary>
        /// <param name="flg">0:新規登録 1:更新</param>
        private bool WarningCheck(int flg)
        {
            bool isWaring = false;
            //Color clearColor = Color.FromArgb(255, 255, 255);
            //Color clearRequestColor = Color.FromArgb(252, 228, 214);
            

            Color errorColor = Color.FromArgb(255, 204, 255);

            // 更新時
            // 事前打診	１．事前打診状況	参考見積	「検討中」のまま「入札タブ　入札情報登録日」が登録されたら「！」。
            if (IsSpecifiedValue(prior_tbl01_cmbMitumori.SelectedValue, "1"))
            {
                // １．事前打診状況 参考見積	「検討中」のまま
                //入札情報登録日
                if (bid_tbl01_dtpBidInfoDt.CustomFormat == "")
                {
                    // 「入札タブ　入札情報登録日」が登録されたら
                    //GlobalMethod.outputMessage("W10603", "入札");
                    set_error(GlobalMethod.GetMessage("W10603", "入札"));
                    bid_tbl01_lblBidInfoDt.BackColor = errorColor;
                    bid_tbl01_picBidInfoDtAlert.Visible = true;
                    prior_tbl01_lblMitumori.BackColor = errorColor;
                    prior_tbl01_picMitumoriAlert.Visible = true;
                }
            }

            // No.1537　1285　更新時に警告にならない。
            // 事前打診	２．未発注	その他の内容	発注無しの理由「その他」の時、入力欄色付け。更新時警告。
            if (IsSpecifiedValue(prior_tbl02_cmbNotOrderReason.SelectedValue, "4"))
            {
                if (string.IsNullOrEmpty(prior_tbl02_txtOtherNaiyo.Text.Trim()))
                {
                    // 2024/01/09　文言変更：入札⇒事前打診
                    set_error(GlobalMethod.GetMessage("W10604", "事前打診"));
                    prior_tbl02_lblNotOrderReason.BackColor = errorColor;
                    prior_tbl02_picNotOrderReasonAlert.Visible = true;
                    prior_tbl02_txtOtherNaiyo.BackColor = errorColor;
                    prior_tbl02_picOtherNaiyoAlert.Visible = true;
                }
            }

            // No1562 1313　進捗段階「入札」にチェックが入っていないが、（入札）タブの、落札者状況が未設定ですと出る。
            if (base_tbl01_dtpDtBid.CustomFormat == "")
            {
                // １．入札情報	入札状況	「入札中」のまま「入札タブ　入札結果登録日」が登録されたら警告。 入札中⇒入札前
                if (IsSpecifiedValue(bid_tbl03_1_cmbBidStatus.SelectedValue, "1"))
                {
                    if (bid_tbl03_1_dtpBidResultDt.CustomFormat == "")
                    {
                        //GlobalMethod.outputMessage("W10605", "入札");
                        set_error(GlobalMethod.GetMessage("W10605", "入札"));
                        bid_tbl03_1_lblBidResultDt.BackColor = errorColor;
                        bid_tbl03_1_picBidResultDtAlert.Visible = true;
                        bid_tbl03_1_lblBidStatus.BackColor = errorColor;
                        bid_tbl03_1_picBidStatusAlert.Visible = true;

                    }
                }
                // １．入札情報 当会応札	「検討中」のまま「入札タブ 入札結果登録日」が登録されたら警告。
                if (IsSpecifiedValue(bid_tbl01_cmbTokaiOsatu.SelectedValue, "1"))
                {
                    if (bid_tbl03_1_dtpBidResultDt.CustomFormat == "")
                    {
                        //GlobalMethod.outputMessage("W10606", "入札");
                        set_error(GlobalMethod.GetMessage("W10606", "入札"));
                        bid_tbl03_1_lblBidResultDt.BackColor = errorColor;
                        bid_tbl03_1_picBidResultDtAlert.Visible = true;
                        bid_tbl01_lblTokaiOsatu.BackColor = errorColor;
                        bid_tbl01_picTokaiOsatuAlert.Visible = true;
                    }
                }
                // No.1534
                ////入札結果登録日　No.1522（1268）　入札結果登録日の更新時のエラーから警告に変更する。
                //if (!IsSpecifiedValue(bid_tbl03_1_cmbBidStatus.SelectedValue, "1") && bid_tbl03_1_dtpBidResultDt.CustomFormat != "")
                //{
                //    set_error(GlobalMethod.GetMessage("W10611", "入札"));
                //    bid_tbl03_1_lblBidResultDt.BackColor = errorColor;
                //    bid_tbl03_1_picBidResultDtAlert.Visible = true;
                //}
                int difMonth = 0;
                if (bid_tbl03_1_dtpBidResultDt.CustomFormat == "")
                {
                    // No.1534
                    //入札結果登録日　No.1522（1268）　入札結果登録日の更新時のエラーから警告に変更する。
                    if (!IsSpecifiedValue(bid_tbl03_1_cmbBidStatus.SelectedValue, "1") && bid_tbl03_1_dtpBidResultDt.CustomFormat != "")
                    {
                        set_error(GlobalMethod.GetMessage("W10611", "入札"));
                        bid_tbl03_1_lblBidResultDt.BackColor = errorColor;
                        bid_tbl03_1_picBidResultDtAlert.Visible = true;
                    }

                    difMonth = GetElapsedMonths(bid_tbl03_1_dtpBidResultDt.Value, DateTime.Now);
                    // 「入札結果登録日」から2か月経過したら
                    if (difMonth >= 2)
                    {
                        // １．入札情報 落札者状況	「不明」のまま「入札結果登録日」から2か月経過したら警告。
                        if (IsSpecifiedValue(bid_tbl03_1_cmbRakusatuStatus.SelectedValue, "2"))
                        {
                            //GlobalMethod.outputMessage("W10607", "入札");
                            set_error(GlobalMethod.GetMessage("W10607", "入札"));
                            bid_tbl03_1_lblRakusatuStatus.BackColor = errorColor;
                            bid_tbl03_1_picRakusatuStatusAlert.Visible = true;
                            bid_tbl03_1_lblBidResultDt.BackColor = errorColor;
                            bid_tbl03_1_picBidResultDtAlert.Visible = true;
                        }
                        // １．入札情報 落札額状況	「不明」のまま「入札結果登録日」から2か月経過したら警告。
                        if (IsSpecifiedValue(bid_tbl03_1_cmbRakusatuAmtStatus.SelectedValue, "2"))
                        {
                            //GlobalMethod.outputMessage("W10608", "入札");
                            set_error(GlobalMethod.GetMessage("W10608", "入札"));
                            bid_tbl03_1_lblRakusatuAmtStatus.BackColor = errorColor;
                            bid_tbl03_1_picRakusatuAmtStatusAlert.Visible = true;
                            bid_tbl03_1_lblBidResultDt.BackColor = errorColor;
                            bid_tbl03_1_picBidResultDtAlert.Visible = true;
                        }
                    }
                }

                // No.1534 警告に変更する
                //落札者状況
                if (this.IsNotSelected(bid_tbl03_1_cmbRakusatuStatus))
                {
                    bid_tbl03_1_lblRakusatuStatus.BackColor = errorColor;
                    bid_tbl03_1_picRakusatuStatusAlert.Visible = true;
                    set_error(GlobalMethod.GetMessage("W10612", "入札"));
                }
                //落札額状況
                if (this.IsNotSelected(bid_tbl03_1_cmbRakusatuAmtStatus))
                {
                    bid_tbl03_1_lblRakusatuAmtStatus.BackColor = errorColor;
                    bid_tbl03_1_picRakusatuAmtStatusAlert.Visible = true;
                    set_error(GlobalMethod.GetMessage("W10613", "入札"));
                }

                // １．入札情報 落札者	「落札者状況 判明、推定」の時、空欄なら警告。
                object obj = bid_tbl03_1_cmbRakusatuStatus.SelectedValue;
                // 「落札者状況 判明、推定」の時、空欄なら警告。
                if (IsSpecifiedValue(obj, "1") || IsSpecifiedValue(obj, "3"))
                {
                    if (string.IsNullOrEmpty(bid_tbl03_1_txtRakusatuSya.Text.Trim()))
                    {
                        //GlobalMethod.outputMessage("W10609", "入札");
                        set_error(GlobalMethod.GetMessage("W10609", "入札"));
                        bid_tbl03_1_txtRakusatuSya.BackColor = errorColor;
                        bid_tbl03_1_picRakusatuSyaAlert.Visible = true;
                        bid_tbl03_1_lblRakusatuStatus.BackColor = errorColor;
                        bid_tbl03_1_picRakusatuStatusAlert.Visible = true;
                    }
                }

                // １．入札情報 落札額	「落札額状況 判明、推定」の時、空欄なら警告。
                obj = bid_tbl03_1_cmbRakusatuAmtStatus.SelectedValue;
                if (IsSpecifiedValue(obj, "1") || IsSpecifiedValue(obj, "3"))
                {
                    //No1537 1285 更新時に警告にならない。【「落札額状況　判明、推定」の時、空欄なら警告が警告にならない。】
                    //0円なら空とする
                    if (string.IsNullOrEmpty(bid_tbl03_1_numRakusatuAmt.Text.Trim()) || GetLong(bid_tbl03_1_numRakusatuAmt.Text.Trim()) == 0)
                    {
                        set_error(GlobalMethod.GetMessage("W10610", "入札"));
                        bid_tbl03_1_numRakusatuAmt.BackColor = errorColor;
                        bid_tbl03_1_picRakusatuAmtAlert.Visible = true;
                        bid_tbl03_1_lblRakusatuAmtStatus.BackColor = errorColor;
                        bid_tbl03_1_picRakusatuAmtStatusAlert.Visible = true;
                    }
                }
            }

            return isWaring;
        }

        /// <summary>
        /// 月数取得
        /// </summary>
        /// <param name="baseDay"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        private int GetElapsedMonths(DateTime from, DateTime to)
        {
            DateTime baseDay = from;
            DateTime day = to;
            if (day < baseDay)
                // 日付が基準日より前の場合は例外とする
                return 0;

            // 経過月数を求める(満月数を考慮しない単純計算)
            var elapsedMonths = (day.Year - baseDay.Year) * 12 + (day.Month - baseDay.Month);

            if (baseDay.Day <= day.Day)
                // baseDayの日部分がdayの日部分以上の場合は、その月を満了しているとみなす
                // (例:1月30日→3月30日以降の場合は満(3-1)ヶ月)
                return elapsedMonths;
            else if (day.Day == DateTime.DaysInMonth(day.Year, day.Month) && day.Day <= baseDay.Day)
                // baseDayの日部分がdayの表す月の末日以降の場合は、その月を満了しているとみなす
                // (例:1月30日→2月28日(平年2月末日)/2月29日(閏年2月末日)以降の場合は満(2-1)ヶ月)
                return elapsedMonths;
            else
                // それ以外の場合は、その月を満了していないとみなす
                // (例:1月30日→3月29日以前の場合は(3-1)ヶ月未満、よって満(3-1-1)ヶ月)
                return elapsedMonths - 1;
        }

        /// <summary>
        /// 起案時のエラー
        /// </summary>
        /// <returns></returns>
        private bool KianError(int iDummy = 0)
        {
            bool requiredFlag = true;
            bool varidateFlag = true;
            Color errorColor = Color.FromArgb(255, 204, 255);

            // エラー背景色　クリア
            Color clearColor = Color.FromArgb(255, 255, 255);
            Color clearRequestColor = Color.FromArgb(252, 228, 214);
            ca_tbl01_lblChangeDt.BackColor = clearRequestColor;
            ca_tbl01_lblKianDt.BackColor = clearRequestColor;
            ca_tbl01_lblKokiFrom.BackColor = clearRequestColor;
            ca_tbl01_lblKokiTo.BackColor = clearRequestColor;
            ca_tbl01_txtRiyu.BackColor = clearColor;
            ca_tbl01_lblAnkenKubun.BackColor = clearColor;
            ca_tbl01_txtZeikomiAmt.BackColor = clearColor;
            ca_tbl02_AftCaBmZeikomi_numAmtAll.BackColor = clearColor;
            ca_tbl02_AftCaBm_numAmtAll.BackColor = clearColor;
            ca_tbl02_AftCaBm_numPercentAll.BackColor = clearColor;
            ca_tbl05_txtGyomu.BackColor = clearColor;
            ca_tbl05_txtMadoguchi.BackColor = clearColor;

            // アラートアイコン非表示にする
            ca_tbl01_picChangeDtAlert.Visible = false;
            ca_tbl01_picKianDtAlert.Visible = false;
            ca_tbl01_picKokiFromAlert.Visible = false;
            ca_tbl01_picKokiToAlert.Visible = false;
            ca_tbl01_picRiyuAlert.Visible = false;
            ca_tbl01_picAnkenKubunAlert.Visible = false;
            ca_tbl01_picZeikomiAmtAlert.Visible = false;
            ca_tbl02_AftCaBmZeikomi_picAmtAlert.Visible = false;
            ca_tbl02_AftCaBm_picAmtAlert.Visible = false;
            ca_tbl02_AftCaBm_picPercentAlert.Visible = false;
            ca_tbl05_picGyomuAlert.Visible = false;
            ca_tbl05_picMadoguchiAlert.Visible = false;

            set_error("", 0);

            // １．契約情報	契約締結(変更)日、契約金額、受託外金額(税込)の未登録があれば更新不可。
            if (ca_tbl01_dtpChangeDt.CustomFormat != "")
            {
                ca_tbl01_lblChangeDt.BackColor = errorColor;
                ca_tbl01_picChangeDtAlert.Visible = true;
                requiredFlag = false;
            }
            // 起案日も
            if (ca_tbl01_dtpKianDt.CustomFormat != "")
            {
                ca_tbl01_lblKianDt.BackColor = errorColor;
                ca_tbl01_picKianDtAlert.Visible = true;
                requiredFlag = false;
            }
            if (GetLong(ca_tbl01_txtZeikomiAmt.Text) == 0)
            {
                ca_tbl01_txtZeikomiAmt.BackColor = errorColor;
                ca_tbl01_picZeikomiAmtAlert.Visible = true;
                requiredFlag = false;
            }
            // 受託外金額(税込)の未登録があれば更新不可。 ★★★
            //if (GetLong(ca_tbl01_txtJyutakuGaiAmt.Text) == 0)
            //{
            //    ca_tbl01_txtJyutakuGaiAmt.BackColor = errorColor;
            //    requiredFlag = false;
            //}
            // 変更伝票の場合は、変更・中止理由未入力で更新不可。
            if (mode == MODE.CHANGE)
            {
                //　案件区分も
                if (IsNotSelected(ca_tbl01_cmbAnkenKubun))
                {
                    ca_tbl01_lblAnkenKubun.BackColor = errorColor;
                    ca_tbl01_picAnkenKubunAlert.Visible = true;
                    requiredFlag = false;
                }
                //契約工期自
                if (ca_tbl01_dtpKokiFrom.CustomFormat == " ")
                {
                    requiredFlag = false;
                    ca_tbl01_lblKokiFrom.BackColor = errorColor;
                    ca_tbl01_picKokiFromAlert.Visible = true;
                }

                //契約工期至
                if (ca_tbl01_dtpKokiTo.CustomFormat == " ")
                {
                    requiredFlag = false;
                    ca_tbl01_lblKokiTo.BackColor = errorColor;
                    ca_tbl01_picKokiToAlert.Visible = true;
                }
                if (string.IsNullOrEmpty(ca_tbl01_txtRiyu.Text))
                {
                    ca_tbl01_txtRiyu.BackColor = errorColor;
                    ca_tbl01_picRiyuAlert.Visible = true;
                    requiredFlag = false;
                }
            }
            // ２．配分情報・業務内容 【契約後】配分額(税込)、配分額(税抜)が澪登録の場合は更新不可。
            long lngAmt = GetLong(ca_tbl02_AftCaBmZeikomi_numAmtAll.Text);
            if (lngAmt == 0)
            {
                ca_tbl02_AftCaBmZeikomi_numAmtAll.BackColor = errorColor;
                ca_tbl02_AftCaBmZeikomi_picAmtAlert.Visible = true;
                requiredFlag = false;
            }
            lngAmt = GetLong(ca_tbl02_AftCaBm_numAmtAll.Text);
            if (lngAmt == 0)
            {
                ca_tbl02_AftCaBm_numAmtAll.BackColor = errorColor;
                ca_tbl02_AftCaBm_picAmtAlert.Visible = true;
                requiredFlag = false;
            }
            // ２．配分情報・業務内容 【契約後】配分率(%)が100％でない場合は更新不可。
            double dblPercent = GetDouble(ca_tbl02_AftCaBm_numPercentAll.Text);
            // No1575 1316　契約画面で起案する際、配分率が99％の場合、エラーとなって起案出来ない。
            //if (dblPercent != 100.00)
            if (dblPercent < 99)
            {
                ca_tbl02_AftCaBm_numPercentAll.BackColor = errorColor;
                ca_tbl02_AftCaBm_picPercentAlert.Visible = true;
                requiredFlag = false;
            }
            // ５．管理者・担当者 業務管理者、窓口担当者が未設定で更新不可。
            //業務担当者
            if (String.IsNullOrEmpty(ca_tbl05_txtGyomu.Text))
            {
                requiredFlag = false;
                ca_tbl05_txtGyomu.BackColor = errorColor;
                ca_tbl05_picGyomuAlert.Visible = true;
            }

            //窓口担当者
            if (String.IsNullOrEmpty(ca_tbl05_txtMadoguchi.Text))
            {
                requiredFlag = false;
                ca_tbl05_txtMadoguchi.BackColor = errorColor;
                ca_tbl05_picMadoguchiAlert.Visible = true;
            }

            //必須項目エラーの出力
            if (!requiredFlag)
            {
                set_error(GlobalMethod.GetMessage("E10010", ""));
            }

            //受託番号
            if ((mode == MODE.SPACE || mode == MODE.UPDATE) && String.IsNullOrEmpty(base_tbl02_txtJyutakuNo.Text))
            {
                varidateFlag = false;
                set_error(GlobalMethod.GetMessage("E10722", ""));
            }

            if (mode != MODE.CHANGE)
            {
                //入札状況が入札成立でなければ起案エラー
                if (!isNyuusatsu_seiritsu(bid_tbl03_1_cmbBidStatus.SelectedValue.ToString()))
                {
                    varidateFlag = false;
                    set_error(GlobalMethod.GetMessage("E10702", ""));
                }
                //落札者が建設物価調査会でなければ起案エラー
                if (!GlobalMethod.GetCommonValue2("ENTORY_TOUKAI").Equals(bid_tbl03_1_txtRakusatuSya.Text))
                {
                    varidateFlag = false;
                    set_error(GlobalMethod.GetMessage("E70048", ""));
                }
            }
            //契約金額の税込
            //契約タブの1.契約情報の契約金額の税込が0円の場合
            long item13 = GetLong(ca_tbl01_txtZeikomiAmt.Text);
            if (item13 == 0)
            {
                // 0円起案です。
                set_error(GlobalMethod.GetMessage("W10701", ""));
                varidateFlag = false;
            }

            //契約工期至
            String format = "yyyy/MM/dd";       //日付フォーマット
            //契約タブの1.契約情報の契約工期至と売上年度が空でない場合
            if (ca_tbl01_dtpKokiTo.CustomFormat == "" && !String.IsNullOrEmpty(ca_tbl01_cmbSalesYear.Text))
            {
                //売上年度 +1年 の3月31日
                int year = Int32.Parse(ca_tbl01_cmbSalesYear.SelectedValue.ToString()) + 1;
                String date = year + "/03/31";
                //日付型
                DateTime nextYear = DateTime.ParseExact(date, format, null);
                DateTime keiyaku = DateTime.ParseExact(ca_tbl01_dtpKokiTo.Text, format, null);
                //売上年度+1/03/31よりも、契約工期の完了日が未来日付の場合エラー
                if (nextYear.Date < keiyaku.Date)
                {
                    varidateFlag = false;
                    set_error(GlobalMethod.GetMessage("E10706", ""));
                }
            }

            //契約タブの1.契約情報の契約工期至と契約工期自が空でない
            if (ca_tbl01_dtpKokiFrom.CustomFormat != " " && ca_tbl01_dtpKokiTo.CustomFormat != " ")
            {
                //日付型

                DateTime keiyakuFrom = DateTime.ParseExact(ca_tbl01_dtpKokiFrom.Text, format, null);
                DateTime keiyakuEnd = DateTime.ParseExact(ca_tbl01_dtpKokiTo.Text, format, null);
                if (keiyakuFrom.Date > keiyakuEnd.Date)
                {
                    varidateFlag = false;
                    set_error(GlobalMethod.GetMessage("E10011", "(契約工期自・至)"));
                }
            }

            //契約タブの6.売上計上情報の工期末日付が空でなく
            C1FlexGrid c1FlexGrid4 = ca_tbl06_c1FlexGrid;
            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            {
                if (c1FlexGrid4[i, 1] != null && c1FlexGrid4[i, 1].ToString() != "")
                {
                    DateTime kokiDate;
                    if (DateTime.TryParse(c1FlexGrid4[i, 1].ToString(), out kokiDate))
                    {
                        //契約工期自が工期末日付より大きい、または、契約工期至が工期末日付より小さい場合
                        if (ca_tbl01_dtpKokiFrom.Value > kokiDate || ca_tbl01_dtpKokiTo.Value < kokiDate)
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
                        if (ca_tbl01_dtpKokiFrom.Value > kokiDate || ca_tbl01_dtpKokiTo.Value < kokiDate)
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
                        if (ca_tbl01_dtpKokiFrom.Value > kokiDate || ca_tbl01_dtpKokiTo.Value < kokiDate)
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
                        if (ca_tbl01_dtpKokiFrom.Value > kokiDate || ca_tbl01_dtpKokiTo.Value < kokiDate)
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
            if (GetDouble(ca_tbl02_AftCaBm_numPercent1.Text) > 0)
            {
                if (ca_tbl02_AftCaTs_numPercentAll.Text != "100.00%")
                {
                    // 調査業務別　配分の合計が100になるように入力してください。
                    set_error(GlobalMethod.GetMessage("E70045", "契約タブ"));
                    varidateFlag = false;
                }
            }
            // 契約タブの調査部配分が0の場合
            // 調査部 業務別配分の合計が0でないとエラー
            else
            {
                if (ca_tbl02_AftCaTs_numPercentAll.Text != "0.00%")
                {
                    // 調査部　業務別配分の合計が不正です。
                    set_error(GlobalMethod.GetMessage("E10725", "契約タブ"));
                    varidateFlag = false;
                }
            }

            // ２．配分情報・業務内容 配分額と契約金額（税込）が一致しない場合更新不可。
            if (GetLong(ca_tbl02_AftCaBmZeikomi_numAmtAll.Text) != GetLong(ca_tbl01_txtJyutakuAmt.Text))
            {
                set_error(GlobalMethod.GetMessage("E10705", ""));
                varidateFlag = false;
            }

            // ７．請求書情報 請求金額と契約金額（税込）が一致しない場合
            if (GetLong(ca_tbl07_txtRequstAll.Text) != GetLong(ca_tbl01_txtZeikomiAmt.Text))
            {
                set_error(GlobalMethod.GetMessage("E10707", ""));
                varidateFlag = false;
            }

            //// ２．配分情報・業務内容 請求金額と配分額合計が一致しない場合更新不可。
            //if (GetLong(ca_tbl02_AftCaBmZeikomi_numAmtAll.Text) != GetLong(ca_tbl07_txtRequstAll.Text))
            //{
            //    set_error(GlobalMethod.GetMessage("E10720", ""));
            //    varidateFlag = false;
            //}

            // ６．売上計上情報 売上情報と契約金額（税込）の合計額が一致しない場合、更新不可。
            Double uriageTotal = 0;
            for (int i = 2; i < ca_tbl06_c1FlexGrid.Rows.Count; i++)
            {
                if (ca_tbl06_c1FlexGrid[i, 3] != null) uriageTotal += GetDouble(ca_tbl06_c1FlexGrid[i, 3].ToString());
                if (ca_tbl06_c1FlexGrid[i, 11] != null) uriageTotal += GetDouble(ca_tbl06_c1FlexGrid[i, 11].ToString());
                if (ca_tbl06_c1FlexGrid[i, 19] != null) uriageTotal += GetDouble(ca_tbl06_c1FlexGrid[i, 19].ToString());
                if (ca_tbl06_c1FlexGrid[i, 27] != null) uriageTotal += GetDouble(ca_tbl06_c1FlexGrid[i, 27].ToString());
            }
            if (!Double.Equals(GetDouble(ca_tbl01_txtJyutakuAmt.Text), uriageTotal))
            {
                set_error(GlobalMethod.GetMessage("E10732", ""));
                varidateFlag = false;
            }

            // No1209 要望ほか、６件 アラートからエラーメッセージへ変更する
            // ⇒データチェックに移動しました。

            // 入札情報    当会応札（入札タブ）が「対応前」のままで起案しようとしたら更新不可。
            if (bid_tbl01_cmbTokaiOsatu.SelectedValue != null && bid_tbl01_cmbTokaiOsatu.SelectedValue.ToString().Equals("1"))
            {
                //GlobalMethod.outputMessage("E10729", "(入札タブ)");
                set_error(GlobalMethod.GetMessage("E10729", "入札"));
                varidateFlag = false;
            }

            if (requiredFlag && varidateFlag)
            {
                // 税込
                // 契約タブの1.契約情報の消費税率が空ではない場合
                if (!String.IsNullOrEmpty(ca_tbl01_txtTax.Text))
                {
                    Double keiyakuAmount = GetDouble(ca_tbl01_txtZeikomiAmt.Text);  // 契約金額の税込 
                    Double taxAmount = GetDouble(ca_tbl01_txtTax.Text);      // 消費税 
                    Double inTaxAmount = GetDouble(ca_tbl01_txtSyohizeiAmt.Text);    // 内消費税
                    Double taxPercent = GetDouble(ca_tbl01_txtTax.Text);     // 消費税率
                    Double totalHundred = Convert.ToDouble(100);

                    // 契約金額の税込 / (100 + 消費税率))* 消費税率, 0) の小数点切り捨て　amount
                    Double amount = Math.Floor(keiyakuAmount / (totalHundred + taxPercent) * taxPercent);

                    // 内消費税がamountと一致しない
                    if (!Double.Equals(inTaxAmount, amount))
                    {
                        GlobalMethod.outputMessage("E10704", "");
                    }
                }

                //えんとり君修正STEP2
                if (iDummy == 0)
                {
                    //①当会応札（入札タブ）が「対応前」のままで起案しようとしたらアラート表示する。
                    if (IsSpecifiedValue(bid_tbl01_cmbTokaiOsatu.SelectedValue,"1"))
                    {
                        GlobalMethod.outputMessage("E10729", "入札");
                    }
                }
            }

            if (!requiredFlag || !varidateFlag)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// アラートアイコン表示にする
        /// </summary>
        private void clearAlertIcon()
        {
            //基本情報等一覧
            base_tbl01_picPriorAlert.Visible = false;    // 事前打診　登録日
            base_tbl01_picBidAlert.Visible = false;    // 入札　登録日
            base_tbl01_picCaAlert.Visible = false;    // 契約　登録日
            base_tbl02_picKeikakuNoAlert.Visible = false;    // 計画番号
            base_tbl02_picJyutakuKasyoSibuAlert.Visible = false;    // 受託部所
            base_tbl02_picKeiyakuTantoAlert.Visible = false;    // 契約担当者
            base_tbl02_picAnkenFolderAlert.Visible = false;    // 案件フォルダ
            base_tbl03_picGyomuNameAlert.Visible = false;    // 業務名称
            base_tbl03_picKeiyakuKubunAlert.Visible = false;    // 契約区分
            base_tbl03_picKokiFromAlert.Visible = false;    // 工期自
            base_tbl03_picKokiToAlert.Visible = false;    // 工期至
            base_tbl03_picKokiStartYearAlert.Visible = false;    // 開始年度　※使わない
            base_tbl03_picKokiSalesYearAlert.Visible = false;    // 売上年度　※使わない
            base_tbl04_picOrderCdAlert.Visible = false;    // 発注者コード
            base_tbl04_picOrderKubun1Alert.Visible = false;    // 発注者区分１
            base_tbl04_picOrderKubun2Alert.Visible = false;    // 発注者区分２
            base_tbl07_1_picPercentAlert.Visible = false;    // 部門配分
            base_tbl07_2_picPercentAlert.Visible = false;    // 調査部　業務配分
            // No.1533 削除
            //base_tbl07_3_picOenAlert.Visible = false;    // 応援依頼の有無
            base_tbl07_3_picOenMemoAlert.Visible = false;    // 応援依頼メモ
            base_tbl07_3_picOenIraiAlert.Visible = false;    // 応援依頼先
            base_tbl09_picJizenDasinIraiDtAlert.Visible = false;    // 事前打診依頼日
            base_tbl09_picSankomitumoriAlert.Visible = false;    // 参考見積対応
            base_tbl09_picOrderIyokuAlert.Visible = false;    // 受注意欲
            base_tbl09_picOrderYoteiDtAlert.Visible = false;    // 発注予定・見込日
            base_tbl09_picNotOrderStatsAlert.Visible = false;    // 未発注状況
            base_tbl09_picNotOrderReasonAlert.Visible = false;    // 発注無しの理由
            base_tbl09_picOthenCommentAlert.Visible = false;    // その他内容
            base_tbl10_picOrderKubunAlert.Visible = false;    // 業務発注区分
            base_tbl10_picNyusatuHosikiAlert.Visible = false;    // 入札方式
            base_tbl10_picLowestUmuAlert.Visible = false;    // 最低制限価格の有無
            base_tbl10_picNyusatuDtAlert.Visible = false;    // 入札予定日
            base_tbl10_picSankoMitumoriAlert.Visible = false;    // 参考見積対応
            base_tbl10_picOrderIyokuAlert.Visible = false;    // 受注意欲
            base_tbl10_picTokaiOsatuAlert.Visible = false;    // 当会応札
            base_tbl10_picKinsiUmuAlert.Visible = false;    // 再委託禁止事項の有無
            base_tbl10_picKinsiNaiyoAlert.Visible = false;    // 再委託禁止事項内容
            base_tbl10_picNyusatuStatsAlert.Visible = false;    // 入札状況
            base_tbl10_picRakusatuStatsAlert.Visible = false;    // 落札者状況
            base_tbl10_picRakusatuAmtStatsAlert.Visible = false;    // 落札額状況
            //事前打診
            prior_tbl01_picDasinIraiDtAlert.Visible = false;    // 事前打診依頼日
            prior_tbl01_picMitumoriAlert.Visible = false;    // 参考見積対応
            prior_tbl01_picMitumoriAmtAlert.Visible = false;    // 参考見積金額
            prior_tbl01_picOrderIyokuAlert.Visible = false;    // 受注意欲
            prior_tbl01_picOrderYoteiDtAlert.Visible = false;    // 発注予定日
            prior_tbl02_picNotOrderDtAlert.Visible = false;    // 未発注の登録日
            prior_tbl02_picNotOrderStatsAlert.Visible = false;    // 未発注状態
            prior_tbl02_picNotOrderReasonAlert.Visible = false;    // 発注なしの理由
            prior_tbl02_picOtherNaiyoAlert.Visible = false;    // その他内容
            //入札
            bid_tbl01_picBidInfoDtAlert.Visible = false;    // 入札情報登録日
            bid_tbl01_picOrderKubunAlert.Visible = false;    // 業務発注区分
            bid_tbl01_picBidhosikiAlert.Visible = false;    // 入札方式
            bid_tbl01_picLowestUmuAlert.Visible = false;    // 最低制限価格の有無
            bid_tbl01_picBidYoteiDtAlert.Visible = false;    // 入札予定日
            bid_tbl01_picMitumoriAlert.Visible = false;    // 参考見積対応
            bid_tbl01_picMitumoriAmtAlert.Visible = false;    // 見積額
            bid_tbl01_picOrderIyokuAlert.Visible = false;    // 受注意欲
            bid_tbl01_picTokaiOsatuAlert.Visible = false;    // 当会応札
            bid_tbl02_picKinsiUmuAlert.Visible = false;    // 再委託禁止事項の有無
            bid_tbl02_picKinsiNaiyoAlert.Visible = false;    // 再委託禁止事項内容
            bid_tbl02_picOtherNaiyoAlert.Visible = false;    // その他内容
            bid_tbl03_1_picBidResultDtAlert.Visible = false;    // 入札結果登録日
            bid_tbl03_1_picBidStatusAlert.Visible = false;    // 入札状況
            bid_tbl03_1_picYoteiPriceAlert.Visible = false;    // 予定価格
            bid_tbl03_1_picRakusatuSyaAlert.Visible = false;    // 落札者
            bid_tbl03_1_picRakusatuAmtAlert.Visible = false;    // 落札金額
            bid_tbl03_1_picOsatuNumAlert.Visible = false;    // 応札数
            bid_tbl03_1_picRakusatuStatusAlert.Visible = false;    // 落札者状況
            bid_tbl03_1_picRakusatuAmtStatusAlert.Visible = false;    // 落札額状況
            //契約
            ca_tbl01_picKianDtAlert.Visible = false;    // 起案日
            ca_tbl01_picJyutakuAmtAlert.Visible = false;    // 受託金額
            ca_tbl01_picJyutakuGaiAmtAlert.Visible = false;    // 受託外金額
            ca_tbl01_picZeikomiAmtAlert.Visible = false;    // 契約金額　税込み
            ca_tbl01_picChangeDtAlert.Visible = false;    // 契約変更日
            ca_tbl01_picKokiFromAlert.Visible = false;    // 工期自
            ca_tbl01_picKokiToAlert.Visible = false;    // 工期至
            ca_tbl01_picRiyuAlert.Visible = false;    // 変更理由
            ca_tbl05_picGyomuAlert.Visible = false;    // 業務管理者
            ca_tbl05_picMadoguchiAlert.Visible = false;    // 窓口担当者
            ca_tbl01_picAnkenKubunAlert.Visible = false;    // 案件区分
            ca_tbl02_AftCaTs_picPercentAlert.Visible = false;
            //技術者評価
            te_picPointAlert.Visible = false;    // 業務得点
            te_picKanriPointAlert.Visible = false;    // 管理得点
            te_picSyosaPointAlert.Visible = false;    // 照査得点
            te_picTantoPointAlert.Visible = false;    // 技術者担当者得点
            te_picSeikyusyoAlert.Visible = false;    // 請求書
        }

        /// <summary>
        /// チェックエラーの背景色クリア
        /// </summary>
        /// <param name="flg"></param>
        private void clearBackColor(int flg = 0)
        {
            // アイコン非表示設定
            clearAlertIcon();

            Color clearColor = Color.FromArgb(255, 255, 255);
            Color clearRequestColor = Color.FromArgb(252, 228, 214);

            // 基本情報等一覧
            //必須チェックを行う項目の背景を白に戻す。
            base_tbl01_lblDtPrior.BackColor = clearColor;
            base_tbl01_lblDtBid.BackColor = clearColor;

            base_tbl02_txtKeikakuNo.BackColor = clearColor;
            base_tbl02_lblJyutakuKasyoSibu.BackColor = clearRequestColor;
            base_tbl02_txtKeiyakuTanto.BackColor = clearColor;
            base_tbl02_txtAnkenFolder.BackColor = clearColor;

            base_tbl03_txtGyomuName.BackColor = clearColor;
            base_tbl03_lblKeiyakuKubun.BackColor = clearRequestColor;
            base_tbl03_lblKokiFrom.BackColor = clearRequestColor;
            base_tbl03_lblKokiTo.BackColor = clearRequestColor;

            base_tbl04_txtOrderCd.BackColor = clearRequestColor;
            base_tbl04_txtOrderKubun1.BackColor = clearRequestColor;
            base_tbl04_txtOrderKubun2.BackColor = clearRequestColor;

            base_tbl07_1_numPercentAll.BackColor = clearColor;
            base_tbl07_2_numPercentAll.BackColor = clearColor;

            // No.1533 削除
            //base_tbl07_3_lblOen.BackColor = clearColor;
            base_tbl07_3_txtOenMemo.BackColor = clearColor;
            base_tbl07_3_lblOenIrai.BackColor = clearColor;

            if (flg == 0)
            {
                // ９と１０チェックあるので、クリア
                base_tbl09_lblJizenDasinIraiDt.BackColor = clearColor;
                base_tbl09_lblSankomitumori.BackColor = clearColor;
                base_tbl09_lblOrderYoteiDt.BackColor = clearColor;
                base_tbl09_lblNotOrderStats.BackColor = clearColor;
                base_tbl09_lblNotOrderReason.BackColor = clearColor;
                base_tbl09_txtOthenComment.BackColor = clearColor;
                base_tbl09_lblOrderIyoku.BackColor = clearColor;
                base_tbl10_lblOrderKubun.BackColor = clearColor;
                base_tbl10_lblNyusatuHosiki.BackColor = clearColor;
                base_tbl10_lblLowestUmu.BackColor = clearColor;
                base_tbl10_lblNyusatuDt.BackColor = clearColor;
                base_tbl10_lblSankoMitumori.BackColor = clearColor;
                base_tbl10_lblOrderIyoku.BackColor = clearColor;
                base_tbl10_lblTokaiOsatu.BackColor = clearColor;
                base_tbl10_lblKinsiUmu.BackColor = clearColor;
                base_tbl10_lblKinsiNaiyo.BackColor = clearColor;
                base_tbl10_lblNyusatuStats.BackColor = clearColor;
                base_tbl10_lblRakusatuStats.BackColor = clearColor;
                base_tbl10_lblRakusatuAmtStats.BackColor = clearColor;
            }
            else
            {
                // 事前打診
                prior_tbl01_lblDasinIraiDt.BackColor = clearColor;
                prior_tbl01_lblMitumori.BackColor = clearColor;
                prior_tbl02_lblNotOrderReason.BackColor = clearColor;
                prior_tbl02_txtOtherNaiyo.BackColor = clearColor;
                prior_tbl01_lblOrderIyoku.BackColor = clearColor;
                
                // 入札
                bid_tbl01_lblBidInfoDt.BackColor = clearColor;
                bid_tbl01_lblOrderKubun.BackColor = clearColor;
                bid_tbl01_lblBidhosiki.BackColor = clearColor;
                bid_tbl01_lblLowestUmu.BackColor = clearColor;
                bid_tbl01_lblBidYoteiDt.BackColor = clearColor;
                bid_tbl01_lblMitumori.BackColor = clearColor;
                bid_tbl01_lblOrderIyoku.BackColor = clearColor;
                bid_tbl01_lblTokaiOsatu.BackColor = clearColor;

                bid_tbl03_1_lblBidResultDt.BackColor = clearColor;
                bid_tbl03_1_lblBidStatus.BackColor = clearColor;
                bid_tbl03_1_lblRakusatuStatus.BackColor = clearColor;
                bid_tbl03_1_lblRakusatuAmtStatus.BackColor = clearColor;
                bid_tbl03_1_txtRakusatuSya.BackColor = clearColor;
                bid_tbl03_1_numRakusatuAmt.BackColor = clearColor;
                bid_tbl03_1_txtOsatuNum.BackColor = clearColor;

                // 契約と評価
                ca_tbl01_lblChangeDt.BackColor = clearColor;
                ca_tbl01_lblKianDt.BackColor = clearColor;
                ca_tbl01_lblKokiFrom.BackColor = clearColor;

                ca_tbl01_lblKokiTo.BackColor = clearColor;
                ca_tbl01_txtZeikomiAmt.BackColor = clearColor;
                ca_tbl01_txtJyutakuAmt.BackColor = clearColor;
                ca_tbl01_txtJyutakuGaiAmt.BackColor = clearColor;
                ca_tbl05_txtKanri.BackColor = clearColor;
                ca_tbl05_txtGyomu.BackColor = clearColor;
                ca_tbl05_txtMadoguchi.BackColor = clearColor;
                if (ca_tbl05_txtTanto_c1FlexGrid.Rows.Count == 1)
                {
                    ca_tbl05_txtTanto_c1FlexGrid.Rows.Add();
                    te_c1FlexGrid.Rows.Add();
                }
                ca_tbl05_txtTanto_c1FlexGrid.GetCellRange(1, 1).StyleNew.BackColor = clearColor;
                ca_tbl05_txtTanto_c1FlexGrid.GetCellRange(1, 2).StyleNew.BackColor = clearColor;
                ca_tbl01_lblChangeDt.BackColor = clearRequestColor;
                ca_tbl01_lblKokiFrom.BackColor = clearRequestColor;
                ca_tbl01_lblKokiTo.BackColor = clearRequestColor;
                ca_tbl06_c1FlexGrid.GetCellRange(2, 3).StyleNew.BackColor = clearColor;
                ca_tbl06_c1FlexGrid.GetCellRange(2, 11).StyleNew.BackColor = clearColor;
                ca_tbl06_c1FlexGrid.GetCellRange(2, 19).StyleNew.BackColor = clearColor;
                ca_tbl06_c1FlexGrid.GetCellRange(2, 27).StyleNew.BackColor = clearColor;
                ca_tbl02_AftCaBmZeikomi_numAmtAll.BackColor = clearColor;
                ca_tbl02_AftCaBm_numAmtAll.BackColor = clearColor;
                ca_tbl02_AftCaBm_numPercentAll.BackColor = clearColor;
                ca_tbl02_AftCaTs_numPercentAll.BackColor = clearColor;
                te_txtPoint.BackColor = clearColor;
                te_txtKanriPoint.BackColor = clearColor;
                te_txtSyosaPoint.BackColor = clearColor;
                te_txtSeikyusyo.BackColor = clearColor;

                for (int i = 1; i < te_c1FlexGrid.Rows.Count; i++)
                {
                    te_c1FlexGrid.GetCellRange(i, 3).StyleNew.BackColor = clearColor;
                }
            }
        }
        #endregion

        #region DB更新処理 Private ---------------------------------------------------
        /// <summary>
        /// 0:新規登録 
        /// 1:更新 
        /// 2:チェック用出力 
        /// 3:起案
        /// 4:変更伝票後の起案
        /// </summary>
        /// <param name="mode"></param>
        /// <returns></returns>
        private Boolean Execute_SQL(int mode)
        {
            string methodName = ".Execute_SQL";
            if (mode == 0)
            {
                string sSibuCd = base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString();
                string sStartYear = base_tbl03_cmbKokiStartYear.SelectedValue.ToString();
                // 案件番号、案件ID　採番　と　新規追加前　チェック ----------------------------------------------------------
                //案件番号
                string jigyoubuHeadCD = getJigyoubuHeadCD(1);
                string ankenNo = base_tbl02_txtAnkenNo.Text == "" ? EntryInputDbClass.getAnkenNo(jigyoubuHeadCD, sStartYear, sSibuCd) : base_tbl02_txtAnkenNo.Text;

                if (string.IsNullOrEmpty(ankenNo))
                {
                    GlobalMethod.outputLogger("InsertAnken", "契約情報登録 案件番号採番エラー", "空案件番号", UserInfos[1]);
                    set_error(GlobalMethod.GetMessage("E10606", ""));
                    return false;
                }
                // 案件ID
                int ankenID = GlobalMethod.getSaiban("AnkenJouhouID");
                if (GlobalMethod.Check_Table(ankenID.ToString(), "AnkenJouhouID", "AnkenJouhou", ""))
                {
                    GlobalMethod.outputLogger("InsertAnken", "契約情報登録 ID重複エラー", ankenID.ToString(), UserInfos[1]);
                    set_error(GlobalMethod.GetMessage("E00091", "(契約情報ID重複)"));
                    return false;
                }

                // 契約情報エントリチェック
                if (GlobalMethod.Check_Table(ankenID.ToString(), "AnkenJouhouID", "KeiyakuJouhouEntory", ""))
                {
                    GlobalMethod.outputLogger("InsertAnken", "契約情報(エントリー)登録 ID重複エラー", ankenID.ToString(), UserInfos[1]);
                    set_error(GlobalMethod.GetMessage("E00091", "(契約情報(エントリー)ID重複)"));
                    return false;
                }

                // 業務情報チェック
                if (GlobalMethod.Check_Table(ankenID.ToString(), "GyoumuJouhouID", "GyoumuJouhou", ""))
                {
                    set_error(GlobalMethod.GetMessage("E00091", "(業務情報ID重複)"));
                    return false;
                }

                // 入札情報　チェック
                if (GlobalMethod.Check_Table(ankenID.ToString(), "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                {
                    GlobalMethod.outputLogger("InsertAnken", "入札情報登録 ID重複エラー", ankenID.ToString(), UserInfos[1]);
                    set_error(GlobalMethod.GetMessage("E00091", "(入札情報ID重複)"));
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
                                "WHERE MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + sSibuCd + "' ";
                        var sda = new SqlDataAdapter(cmd);
                        var dtSb = new DataTable();
                        sda.Fill(dtSb);
                        if (dtSb != null && dtSb.Rows.Count > 0)
                        {
                            // フォルダパスを取得（例：$FOLDER_BASE$/111統括）
                            replaceFolderName = dtSb.Rows[0][0].ToString();
                            // 課所支部のフォルダ部分のみとする $FOLDER_BASE$/xxx 
                            replaceFolderName = replaceFolderName.Replace(@"$FOLDER_BASE$/", "");
                            // 課所支部のフォルダ部分のみとする $FOLDER_BASE$ しかない場合の対応
                            replaceFolderName = replaceFolderName.Replace(@"$FOLDER_BASE$", "");
                        }
                        dtSb.Clear();
                    }
                    catch (Exception)
                    {
                        // エラー
                        GlobalMethod.outputLogger("Execute_SQL", "M_Folderから画面の受託課所支部：" + sSibuCd + " のフォルダパスが取得できずにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
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
                        var dtLogin = new DataTable();
                        sda.Fill(dtLogin);
                        if (dtLogin != null && dtLogin.Rows.Count > 0)
                        {
                            // フォルダパスを取得（例：$FOLDER_BASE$/111統括）
                            replaceTargetFolderName = dtLogin.Rows[0][0].ToString();
                            // 課所支部のフォルダ部分のみとする
                            replaceTargetFolderName = replaceTargetFolderName.Replace(@"$FOLDER_BASE$/", "");
                        }
                        dtLogin.Clear();
                    }
                    catch (Exception)
                    {
                        // エラー
                        GlobalMethod.outputLogger("Execute_SQL", "M_Folderからログインユーザーの受託課所支部：" + UserInfos[2] + " のフォルダパスが取得できずにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                    }
                }

                // 案件（受託）フォルダ
                ankenFolder = base_tbl02_txtAnkenFolder.Text;
                if (replaceTargetFolderName != "")
                {
                    // 自分の部署フォルダを画面の選択している受託課所支部のフォルダに置き換える
                    ankenFolder = ankenFolder.Replace(replaceTargetFolderName, replaceFolderName);

                    // 867
                    // 工期開始年度　2021年度まで、　010北道
                    // 工期開始年度　2022年度から　　010北海
                    int koukinendo = 0;
                    if (int.TryParse(sStartYear, out koukinendo))
                    {
                        // No1563 1314　北海道のフォルダ名が間違ってる。　×　010北道　○　010北海
                        if (koukinendo > 2021)
                        {
                            ankenFolder = change_hokaido_path(ankenFolder, koukinendo);
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
                DataTable dt = new System.Data.DataTable();
                dt = GlobalMethod.getData("BushoShozokuChou", "BushoShozokuChou", "Mst_Busho", "GyoumuBushoCD = '" + sSibuCd + "'");
                if (dt != null && dt.Rows.Count > 0)
                {
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
                        // 案件情報テーブルへデータ作成
                        cmd.CommandText = "INSERT INTO AnkenJouhou ( " + EntryInputDbClass.getInsAnkenCols() +
                                            " ) VALUES ( " +
                                            getInsAnkenVals(ankenID.ToString(), ankenNo, ankenFolder, BushoShozokuChou) +
                                            ")";

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        // 今の計画番号の案件数
                        if (base_tbl02_txtKeikakuNo.Text != "")
                        {
                            cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + base_tbl02_txtKeikakuNo.Text + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + base_tbl02_txtKeikakuNo.Text + "' ";
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

                        // 過去案件リストへデータ作成(base_tbl08_c1FlexGrid)
                        createAnkenJouhouZenkaiRakusatsu(base_tbl08_c1FlexGrid, cmd, ankenID.ToString());

                        // KeiyakuJouhouEntoryへデータ作成
                        createKeiyakuJouhouEntory(cmd, ankenID.ToString());
                        GlobalMethod.outputLogger("InsertAnken", "契約情報更新時に案件情報を同時更新します AnkenjouhoID = " + ankenID, "", UserInfos[1]);

                        // 業務情報へデータ作成
                        createGyoumuJouhou(cmd, ankenID.ToString());

                        //業務配分へデータ作成
                        int HaibunID = GlobalMethod.getSaiban("GyoumuHaibunID");
                        createGyoumuHaibun10(cmd, ankenID.ToString(), HaibunID.ToString());
                        HaibunID = GlobalMethod.getSaiban("GyoumuHaibunID");
                        createGyoumuHaibun30(cmd, ankenID.ToString(), HaibunID.ToString());
                        createAnkenOuenIraisaki(cmd, ankenID.ToString());

                        // 入札情報へデータ作成
                        createNyuusatsuJouhou(cmd, ankenID.ToString());
                        createKokyakuKeiyakuJouhou(cmd, ankenID.ToString());

                        // No1594　1323　エントリくんのコピー機能で、「この案件番号の枝番でコピーする」で、落札者が受注で登録されていない。
                        if((this.mode == MODE.INSERT || this.mode == MODE.PLAN) && copy == COPY.ED)
                        {
                            // 入札参加者リストがある場合、登録する
                            if(this.AnkenData_Grid2!=null && this.AnkenData_Grid2.Rows.Count > 0)
                            {
                                // No1594 差し戻し　「応札額がコピーされている為、更新すると落札額に入ってしまう。調査会様のみ＆応札額=0でコピーする」
                                DataRow[] rows = AnkenData_Grid2.Select("NyuusatsuOusatsuKyougouKigyouCD = '6010005018675' AND NyuusatsuRakusatsuJokyou = 1");
                                int i = 1;
                                foreach(DataRow rw in rows)
                                {
                                    StringBuilder sql = new StringBuilder();
                                    sql.Append("INSERT NyuusatsuJouhouOusatsusha ( ");
                                    sql.Append("NyuusatsuJouhouID ");
                                    sql.Append(", NyuusatsuOusatsuID ");
                                    sql.Append(", NyuusatsuRakusatsuJyuni ");//落札順位
                                    sql.Append(", NyuusatsuRakusatsuJokyou ");//落札状況
                                    sql.Append(", NyuusatsuOusatsushaID ");
                                    sql.Append(", NyuusatsuOusatsusha ");
                                    sql.Append(", NyuusatsuOusatsuKingaku ");
                                    sql.Append(", NyuusatsuOusatsuKyougouTashaID ");
                                    sql.Append(", NyuusatsuRakusatsuComment ");
                                    sql.Append(", NyuusatsuOusatsuKyougouKigyouCD ");
                                    sql.Append(") VALUES (");
                                    sql.Append("'" + ankenID.ToString() + "' ");
                                    sql.Append("," + i);
                                    sql.Append(", " + (string.IsNullOrEmpty(rw["NyuusatsuRakusatsuJyuni"].ToString()) ? "NULL " : rw["NyuusatsuRakusatsuJyuni"].ToString() + " "));
                                    sql.Append(",'1' ");
                                    sql.Append(",N'" + rw["NyuusatsuOusatsushaID"].ToString() + "' ");
                                    sql.Append(",N'" + rw["NyuusatsuOusatsusha"].ToString() + "' ");
                                    sql.Append(",N'0' ");
                                    // sql.Append(",N'" + rw["NyuusatsuOusatsuKingaku"].ToString() + "' ");
                                    sql.Append(",N'" + rw["NyuusatsuOusatsuKyougouTashaID"].ToString() + "' ");
                                    sql.Append(",N'" + rw["NyuusatsuRakusatsuComment"].ToString() + "' ");
                                    sql.Append(",N'" + rw["NyuusatsuOusatsuKyougouKigyouCD"].ToString() + "')");
                                    cmd.CommandText = sql.ToString();
                                    Console.WriteLine(cmd.CommandText);
                                    cmd.ExecuteNonQuery();
                                    i++;
                                }
                            }
                        }
                        transaction.Commit();

                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        throw;
                        return false;
                    }
                    conn.Close();

                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報を登録しました ID:" + ankenID, pgmName + methodName, "");

                    // 売上年度が2021年以上の場合にフォルダを作成しにいく & 受託番号がない場合（この案件番号の枝番で新規作成、でない場合）
                    // フォルダ作成関連は工期開始年度で行う
                    if (GetInt(sStartYear) >= 2021 && AnkenbaBangou == "")
                    {
                        GlobalMethod.CreateFolder(ankenID);
                    }
                    else
                    {
                        GlobalMethod.outputLogger("CreateFolder", "事業分類CD:" + jigyoubuHeadCD + " 年度:" + base_tbl03_cmbKokiStartYear.Text + " の為フォルダ自動生成なし", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                    }

                    // Roleと部所をみて、参照モードか更新モードを切り替える
                    // Role:1管理者 で、部所がログインユーザーの部所と異なる場合は、参照モード
                    // Role:2:システム管理者の場合は、無条件に更新モード
                    MODE formmode = MODE.SPACE;
                    if (UserInfos[4].Equals("2"))
                    {
                        formmode = MODE.UPDATE;
                    }
                    else
                    {
                        // ログインユーザーの部所と一致しているかどうか
                        if (UserInfos[2] != base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString())
                        {
                            formmode = MODE.VIEW;
                        }
                    }
                    // 画面遷移処理
                    gotoSelfPage(formmode, COPY.NO, ankenID.ToString());
                }

            }

            else if (mode >= 1 && mode != 4)
            {
                // 1:更新 2:チェック用出力 3:起案 ----------------------------------
                // チェック処理
                // -- 契約情報(エントリー)存在チェック
                if (!GlobalMethod.Check_Table(AnkenID, "AnkenJouhouID", "KeiyakuJouhouEntory", ""))
                {
                    GlobalMethod.outputLogger("UpdateEntory", "契約情報(エントリー)更新 データなしエラー", AnkenID, UserInfos[1]);
                    set_error(GlobalMethod.GetMessage("E10009", "(契約情報(エントリー)データなし)"));
                    return false;
                }

                // 案件番号
                string ankenNo = base_tbl02_txtAnkenNo.Text;
                string ori_ankenNo = base_tbl02_txtAnkenNo.Text;
                //案件番号変更
                // No1557 1308 案件情報で調査会が受注後もフォルダ変更が出来てしまう。
                // No1558 1309 案件情報で受注後も工期自、工期至の変更を行うと、案件番号が変更されてしまう。
                // 受託番号の設定OR解除
                setOrClearJutakuBan(ankenNo);

                // 案件番号変更処理
                if (base_tbl02_txtJyutakuNo.Text == "")
                {
                    if (sJyutakuKasyoSibuCdOri.Equals(base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString()) == false || sKokiStartYearOri.Equals(base_tbl03_cmbKokiStartYear.SelectedValue.ToString()) == false)
                    {
                        ankenNo = changeAnkenNo(ori_ankenNo);
                        if (string.IsNullOrEmpty(ankenNo))
                        {
                            // 案件番号変更エラー
                            GlobalMethod.outputLogger("InsertAnken", "契約情報登録 案件番号変更エラー", "空案件番号", UserInfos[1]);
                            set_error(GlobalMethod.GetMessage("E10606", ""));
                            return false;
                        }
                    }
                }


                // フォルダリネーム処理
                // No1557 1308 案件情報で調査会が受注後もフォルダ変更が出来てしまう。
                // No1558 1309 案件情報で受注後も工期自、工期至の変更を行うと、案件番号が変更されてしまう。
                if (base_tbl02_txtJyutakuNo.Text == "") {
					//No1668 ファイル更新ボタンを押下後、変更フォルダが表示されたときに、確認ダイアログを表示させる。OKのみの確認ダイアログとする。
                    //No1668　かつフォルダ変更ボタンを押した場合は、１度表示しているため確認ダイアログは非表示とする
					if (base_tbl02_txtRenameFolder.Text.Length != 0 && base_tbl02_txtRenameFolder.Text != sFolderRenameBef && !isClickedRenameFolderButton)
					{
                        MessageBox.Show(GlobalMethod.GetMessage("E20908", ""), "確認", MessageBoxButtons.OK);

                    }
                    //リネーム処理実行
                    bool isSuccessRenameFolder = RenameFolder(ori_ankenNo);
                    if (isSuccessRenameFolder)
                    {
                        // 移動履歴LOG残す
                        GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "フォルダ変更前：" + GlobalMethod.ChangeSqlText(sFolderRenameBef, 0, 0) + "→フォルダ変更後：" + GlobalMethod.ChangeSqlText(base_tbl02_txtAnkenFolder.Text, 0, 0), pgmName + methodName, "");

                        //フォルダ変更ボタンクリック済みフラグをOFFにする
                        isClickedRenameFolderButton = false;
                    }
                }
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {

                        cmd.CommandText = "UPDATE AnkenJouhou SET " + getUpdAnkenVals() + " WHERE AnkenJouhouID = " + AnkenID;
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        // 計画情報の案件数を設定しなおし
                        cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + base_tbl02_txtKeikakuNo.Text + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + base_tbl02_txtKeikakuNo.Text + "' ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE KeikakuJouhou SET KeikakuAnkensu = (select count(*) from AnkenJouhou where AnkenKeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + beforeKeikakuBangou + "' and AnkenDeleteFlag != 1 and AnkenSaishinFlg = 1) WHERE KeikakuBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + beforeKeikakuBangou + "' ";
                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();

                        // 過去案件リスト更新
                        createAnkenJouhouZenkaiRakusatsu(base_tbl08_c1FlexGrid, cmd, AnkenID);

                        // 契約情報(エントリー)更新
                        updateKeiyakuJouhouEntory(cmd,AnkenID,mode);
                        GlobalMethod.outputLogger("InsertAnken", "契約情報更新時に案件情報を同時更新します AnkenjouhoID = " + AnkenID, "", UserInfos[1]);

                        // 案件情報更新
                        updateAnkenJouhou(cmd, mode);

                        // 売上情報リスト（DELETE⇒INSERT）　契約タブ：６．売上計上情報
                        updateRibcJouhou(cmd, ca_tbl06_c1FlexGrid);

                        // 業務情報　更新
                        updateGyoumuJouhou(cmd);
                        // 業務配分　更新
                        updateGyoumuHaibun10(cmd, AnkenID);
                        updateGyoumuHaibun30(cmd, AnkenID);

                        // 応援依頼先
                        createAnkenOuenIraisaki(cmd, AnkenID, 8);

                        //業務情報技術担当者
                        updateGyoumuJouhouHyouronTantouL1(cmd, te_c1FlexGrid);

                        //窓口担当者の更新
                        updateGyoumuJouhouMadoguchi(cmd);
                        

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

                        // 入札応札者
                        bool nyuusatsuOusatushaUpdateFlg = false;
                        bool nyuusatsuOusatushaInsertFlg = false;
                        int nyusatsuCnt = 0;
                        updateNyuusatsuJouhouOusatsusha(cmd, bid_tbl03_4_c1FlexGrid, ref nyuusatsuOusatushaUpdateFlg, ref nyuusatsuOusatushaInsertFlg, ref nyusatsuCnt);

                        // 入札者情報
                        updateNyuusatsuJouhou(cmd, nyuusatsuOusatushaUpdateFlg, nyuusatsuOusatushaInsertFlg, nyusatsuCnt);

                        transaction.Commit();

                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        conn.Close();
                        throw;
                        //return false;
                    }
                    GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "契約情報(エントリー)を更新しました ID:" + AnkenID, pgmName + methodName, "");

                    if (mode == 1)
                    {
                        //更新時も窓口情報にデータを連携
                        try
                        {
                            updateMadoguchiJouhou(cmd, mode);

                            // Garoon宛先追加
                            insertGaroonAtesakiTsuika(cmd);
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                        set_error(GlobalMethod.GetMessage("I00008", ""));
                        conn.Close();

                        // フォルダ変更コントローラ 表示非表示
                        bool bVisible = false;
                        if (base_tbl02_txtJyutakuNo.Text == "") bVisible = true;
                        setVisibleToRenameFolder(bVisible);

                        return true;
                    }
                    if (mode == 2 || mode == 3)
                    {
                        GlobalMethod.outputLogger("KianEntry", AnkenID + ":" + tblAKInfo_lblJyutakuNo.Text + ":" + base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString() + ":" + base_tbl02_txtKeiyakuTantoCD.Text, "", UserInfos[1]);
                        try
                        {
                            updateMadoguchiJouhou(cmd, mode);

                            // チェック用帳票の出力時はPrintHistoryがあるので出力しない
                            GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "起案しました ID:" + AnkenID, pgmName + methodName, "");

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
                    beforeKeikakuBangou = base_tbl02_txtKeikakuNo.Text;
                }

            }
            // 変更伝票の起案
            else if (mode == 4)
            {
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    try
                    {
                        var result = updateAnkenJouhou(cmd, mode);

                        if (result == 0)
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "案件情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }

                        string SakuseiKubun = ca_tbl01_cmbAnkenKubun.SelectedValue.ToString();

                        #region 変更伝票の起案前チェック
                        if (!GlobalMethod.Check_Table(AnkenID, "KokyakuKeiyakuID", "KokyakuKeiyakuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "顧客契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }
                        if (!GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "業務情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }

                        if (!GlobalMethod.Check_Table(AnkenID, "KeiyakuJouhouEntoryID", "KeiyakuJouhouEntory", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "契約情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }
                        if (!GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhou", ""))
                        {
                            GlobalMethod.outputLogger("ChangeKianEntry", "入札情報が見つからない", "ID:" + AnkenID, "DEBUG");
                            transaction.Rollback();
                            conn.Close();
                            return false;
                        }
                        #endregion

                        #region 赤伝作成 --------------------
                        // 赤伝のAnkenJouhouID
                        int ankenNo = GlobalMethod.getSaiban("AnkenJouhouID");
                        ca_tbl01_hidAkaden.Text = ankenNo.ToString();
                        // 案件情報の赤伝
                        result = createAnkenJouhou(cmd, ankenNo.ToString(), SakuseiKubun);
                        
                        // 過去案件リストの赤伝
                        result = createAnkenJouhouZenkaiRakusatsu(null, cmd, ankenNo.ToString());

                        // 契約情報の赤伝
                        result = createKokyakuKeiyakuJouhou(cmd,ankenNo.ToString(),0);

                        // 業務情報の赤伝
                        result = createGyoumuJouhou(cmd, ankenNo.ToString(), 0);
                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyouronTantouL1", ""))
                        {
                            result = createGyoumuJouhouHyouronTantouL1(cmd, ankenNo.ToString());
                        }

                        // 応援依頼先
                        result = createAnkenOuenIraisaki(cmd, ankenNo.ToString(), 0);

                        // 窓口担当者
                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouMadoguchi", ""))
                        {
                            result = createGyoumuJouhouMadoguchi(cmd, ankenNo.ToString());
                        }

                        if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyoutenBusho", ""))
                        {
                            result = createGyoumuJouhouHyoutenBusho(cmd, ankenNo.ToString());
                        }

                        // 契約情報の赤伝
                        result = createKeiyakuJouhouEntory(cmd, ankenNo.ToString(), 0);

                        if (GlobalMethod.Check_Table(AnkenID, "RibcID", "RibcJouhou", ""))
                        {
                            
                            result = createRibcJouhou(cmd, ankenNo.ToString());
                        }

                        // 入札情報の赤伝
                        result = createNyuusatsuJouhou(cmd, ankenNo.ToString(),0);
                        
                        // 入札応札情報の赤伝
                        if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhouOusatsusha", ""))
                        {
                            result = createNyuusatsuJouhouOusatsusha(cmd, ankenNo.ToString());
                        }

                        // 業務配分
                        DataTable GH_dt = new DataTable();
                        GH_dt = GlobalMethod.getData("GyoumuHaibunID", "GyoumuAnkenJouhouID", "GyoumuHaibun", "GyoumuAnkenJouhouID = " + AnkenID);
                        result = createGyoumuHaibun(cmd, ankenNo.ToString(), GH_dt);
                        #endregion

                        #region 黒伝作成 --------------------
                        int ankenNo2 = 0;
                        //黒伝作成
                        if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                        {
                            // 黒伝のAnkenJouhouID
                            ankenNo2 = GlobalMethod.getSaiban("AnkenJouhouID");
                            ca_tbl01_hidKuroden.Text = ankenNo2.ToString();
                            // 案件情報
                            createAnkenJouhou(cmd, ankenNo2.ToString(), SakuseiKubun,1);

                            // 過去案件落札者
                            createAnkenJouhouZenkaiRakusatsu(null, cmd, ankenNo2.ToString(), 1);

                            // 顧客契約情報
                            createKokyakuKeiyakuJouhou(cmd, ankenNo2.ToString(),1);

                            // 業務情報
                            createGyoumuJouhou(cmd, ankenNo2.ToString(), 1);

                            // 応援依頼先
                            result = createAnkenOuenIraisaki(cmd, ankenNo2.ToString(), 1);

                            // 業務評価担当
                            if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyouronTantouL1", ""))
                            {
                                result =  createGyoumuJouhouHyouronTantouL1(cmd, ankenNo2.ToString(), 1);
                            }
                            // 窓口担当者
                            if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouMadoguchi", ""))
                            {
                                result = createGyoumuJouhouMadoguchi(cmd, ankenNo2.ToString(), 1);
                            }
                            // 業務評価部署
                            if (GlobalMethod.Check_Table(AnkenID, "GyoumuJouhouID", "GyoumuJouhouHyoutenBusho", ""))
                            {
                                result = createGyoumuJouhouHyoutenBusho(cmd, ankenNo2.ToString(), 1);
                            }

                            // 契約情報
                            result = createKeiyakuJouhouEntory(cmd, ankenNo2.ToString(), 1);

                            // Ribc情報
                            if (GlobalMethod.Check_Table(AnkenID, "RibcID", "RibcJouhou", ""))
                            {
                                result = createRibcJouhou(cmd, ankenNo2.ToString(), 1);
                            }

                            // 入札情報
                            result = createNyuusatsuJouhou(cmd, ankenNo2.ToString(), 1);

                            // 応札情報
                            if (GlobalMethod.Check_Table(AnkenID, "NyuusatsuJouhouID", "NyuusatsuJouhouOusatsusha", ""))
                            {
                                result = createNyuusatsuJouhouOusatsusha(cmd, ankenNo2.ToString(), 1);
                            }

                            // 業務配分情報
                            result = createGyoumuHaibun(cmd, ankenNo2.ToString(), GH_dt, 1);
                        }
                        #endregion

                        transaction.Commit();
                        transaction = conn.BeginTransaction();
                        cmd.Transaction = transaction;
                        if (updateAfterCreateRedAndBlack(cmd, transaction, SakuseiKubun, ankenNo.ToString(), ankenNo2.ToString()))
                        {
                            GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "赤伝を作成しました ID:" + ankenNo, pgmName + methodName, "");
                            if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                            {
                                GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "黒伝を作成しました ID:" + ankenNo2, pgmName + methodName, "");
                                set_error(GlobalMethod.GetMessage("I10710", ""));
                            }
                            else
                            {
                                GlobalMethod.Insert_History(UserInfos[0], UserInfos[1], UserInfos[2], UserInfos[3], "中止伝票を追加しました。 ID:" + AnkenID, pgmName + methodName, "");
                                set_error(GlobalMethod.GetMessage("I10711", ""));
                            }
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

        /// <summary>
        /// 伝票変更/ダミーデータで赤黒伝票作成完了後の更新処理
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="transaction"></param>
        /// <param name="SakuseiKubun"></param>
        /// <param name="ankenNo"></param>
        /// <param name="ankenNo2"></param>
        /// <param name="iType">0:伝票変更/7:ダミーデータ</param>
        /// <returns></returns>
        private bool updateAfterCreateRedAndBlack(SqlCommand cmd,SqlTransaction transaction, string SakuseiKubun, string ankenNo , string ankenNo2, int iType = 0)
        {
            bool rtn = false;
            try
            {
                var result = 0;
                // 赤伝：契約情報エントリ　更新

                cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET " +
                             "KeiyakuHenkouChuushiRiyuu = N'" + ca_tbl01_txtRiyu.Text + "' " +
                            ",KeiyakuSakuseibi = " + Get_DateTimePicker("ca_tbl01_dtpKianDt");
                if (SakuseiKubun == "02")
                {
                    cmd.CommandText += ",KeiyakuKeiyakuTeiketsubi = " + Get_DateTimePicker("ca_tbl01_dtpChangeDt");
                }
                cmd.CommandText += " WHERE KeiyakuJouhouEntoryID = " + ankenNo;
                result = cmd.ExecuteNonQuery();

                if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
                {
                    cmd.CommandText = "UPDATE AnkenJouhou SET " +
                            "AnkenSakuseiKubun = N'" + ca_tbl01_cmbAnkenKubun.SelectedValue.ToString() + "' " +
                            ",AnkenGyoumuKubun = N'" + ca_tbl01_cmbCaKubun.SelectedValue.ToString() + "' " +
                            ",AnkenGyoumuKubunMei = N'" + ca_tbl01_cmbCaKubun.Text + "' " +
                            ",AnkenGyoumuMei = N'" + ca_tbl01_txtAnkenName.Text + "' " +
                            ",AnkenKianzumi = '1' " +
                            ",AnkenUriageNendo = N'" + ca_tbl01_cmbSalesYear.SelectedValue.ToString() + "' " +
                            ",AnkenKoukiNendo = N'" + ca_tbl01_cmbStartYear.SelectedValue.ToString() + "' " +
                            ",AnkenKeiyakuTeiketsubi = " + Get_DateTimePicker("ca_tbl01_dtpChangeDt") +
                            ",AnkenKeiyakuKoukiKaishibi = " + Get_DateTimePicker("ca_tbl01_dtpKokiFrom") +
                            ",AnkenKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("ca_tbl01_dtpKokiTo") +
                            ",AnkenKeiyakuZeikomiKingaku = " + getNumToDb(ca_tbl01_txtZeikomiAmt.Text) +
                            ",AnkenKeiyakuUriageHaibunGakuC = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt1.Text) +
                            ",AnkenKeiyakuUriageHaibunGakuJ = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt2.Text) +
                            ",AnkenKeiyakuUriageHaibunGakuJs = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt3.Text) +
                            ",AnkenKeiyakuUriageHaibunGakuK = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt4.Text) +
                            ",AnkenKeiyakuSakuseibi = " + Get_DateTimePicker("ca_tbl01_dtpKianDt") +
                            ",GyoumuKanrishaCD = " + "N'" + ca_tbl05_txtGyomuCD.Text + "'" +
                            ",GyoumuKanrishaMei = " + "N'" + ca_tbl05_txtGyomu.Text + "'" +
                            ",AnkenUpdateUser = N'" + UserInfos[0] + "' " +
                            " WHERE AnkenJouhou.AnkenJouhouID = " + ankenNo2;
                    Console.WriteLine(cmd.CommandText);
                    result = cmd.ExecuteNonQuery();

                    //業務情報
                    cmd.CommandText = "UPDATE GyoumuJouhou SET " +
                                "GyoumuHyouten = " + "N'" + te_txtPoint.Text + "'" +
                                ",KanriGijutsushaCD = " + "N'" + ca_tbl05_txtKanriCD.Text + "'" +
                                ",KanriGijutsushaNM = " + "N'" + ca_tbl05_txtKanri.Text + "'" +
                                ",GyoumuKanriHyouten = " + "N'" + ca_tbl05_txtKanriHyoten.Text + "'" +
                                ",ShousaTantoushaCD = " + "N'" + ca_tbl05_txtSyosaCD.Text + "'" +
                                ",ShousaTantoushaNM = " + "N'" + ca_tbl05_txtSyosa.Text + "'" +
                                ",GyoumuShousaHyouten = " + "N'" + ca_tbl05_txtSyosaHyoten.Text + "'" +
                                ",SinsaTantoushaCD = " + "N'" + ca_tbl05_txtSinsaCD.Text + "'" +
                                ",SinsaTantoushaNM = " + "N'" + ca_tbl05_txtSinsa.Text + "'" +
                                ",GyoumuUpdateDate = " + " GETDATE() " +
                                ",GyoumuUpdateUser = " + "N'" + UserInfos[0] + "' " +
                                ",GyoumuUpdateProgram = " + "'UpdateEntory' " +
                                ",GyoumuDeleteFlag = " + "0 " +
                                " WHERE AnkenJouhouID = " + ankenNo2;
                    Console.WriteLine(cmd.CommandText);
                    cmd.ExecuteNonQuery();

                    // 応援依頼先
                    result = createAnkenOuenIraisaki(cmd, ankenNo2.ToString(), 8);

                    //業務情報技術担当者
                    cmd.CommandText = "DELETE GyoumuJouhouHyouronTantouL1 WHERE GyoumuJouhouID = '" + ankenNo2 + "' ";
                    cmd.ExecuteNonQuery();

                    for (int i = 1; i < te_c1FlexGrid.Rows.Count; i++)
                    {
                        if (ca_tbl05_txtTanto_c1FlexGrid.Rows[i][1] != null && ca_tbl05_txtTanto_c1FlexGrid.Rows[i][1].ToString() != "")
                        {
                            string Hyouten = "";
                            if (ca_tbl05_txtTanto_c1FlexGrid.Rows[i][3] != null && ca_tbl05_txtTanto_c1FlexGrid.Rows[i][3].ToString() != "")
                            {
                                Hyouten = ca_tbl05_txtTanto_c1FlexGrid.Rows[i][3].ToString();
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
                                    ",N'" + te_c1FlexGrid.Rows[i][1].ToString() + "' " +
                                    ",N'" + te_c1FlexGrid.Rows[i][2].ToString() + "' " +
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
                    cmd.CommandText = "SELECT ShibuMei, KaMei FROM Mst_Busho WHERE GyoumuBushoCD = '" + ca_tbl05_txtMadoguchiBusho.Text + "'";
                    var sda2 = new SqlDataAdapter(cmd);
                    dt2.Clear();
                    sda2.Fill(dt2);
                    if (dt2 != null && dt2.Rows.Count > 0)
                    {
                        GyoumuJouhouMadoShibuMei = dt2.Rows[0][0].ToString();
                        GyoumuJouhouMadoKamei = dt2.Rows[0][1].ToString();
                    }

                    //窓口担当者の更新
                    if ((ca_tbl05_txtMadoguchiCD.Text == "0") || (ca_tbl05_txtMadoguchiCD.Text == ""))
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
                                            "GyoumuJouhouMadoGyoumuBushoCD = N'" + ca_tbl05_txtMadoguchiBusho.Text + "' " +
                                            ",GyoumuJouhouMadoShibuMei = N'" + GyoumuJouhouMadoShibuMei + "' " +
                                            ",GyoumuJouhouMadoKamei = N'" + GyoumuJouhouMadoKamei + "' " +
                                            ",GyoumuJouhouMadoKojinCD = N'" + ca_tbl05_txtMadoguchiCD.Text + "' " +
                                            ",GyoumuJouhouMadoChousainMei = N'" + ca_tbl05_txtMadoguchi.Text + "' " +
                                            "WHERE GyoumuJouhouMadoguchiID = N'" + GyoumuJouhouMadoguchiID + "' ";

                            cmd.ExecuteNonQuery();
                        }

                    }

                    updateKeiyakuJouhouEntory(cmd, ankenNo2, 4);
                    //cmd.CommandText = "UPDATE KeiyakuJouhouEntory SET " +
                    //            "KeiyakuKeiyakuTeiketsubi = " + Get_DateTimePicker("ca_tbl01_dtpChangeDt") +
                    //            ",KeiyakuSakuseibi = " + Get_DateTimePicker("ca_tbl01_dtpKianDt") +
                    //            ",KeiyakuKeiyakuKoukiKaishibi = " + Get_DateTimePicker("ca_tbl01_dtpKokiFrom") +
                    //            ",KeiyakuKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("ca_tbl01_dtpKokiTo") +
                    //            ",KeiyakuKeiyakuKingaku = " + getNumToDb(ca_tbl01_txtZeinukiAmt.Text) +
                    //            ",KeiyakuZeikomiKingaku = " + getNumToDb(ca_tbl01_txtZeikomiAmt.Text) +
                    //            ",KeiyakuuchizeiKingaku = " + getNumToDb(ca_tbl01_txtSyohizeiAmt.Text) +
                    //            ",KeiyakuShouhizeiritsu = N'" + ca_tbl01_txtTax.Text + "'" +
                    //            ",KeiyakuHenkouChuushiRiyuu = " + "N'" + GlobalMethod.ChangeSqlText(ca_tbl01_txtRiyu.Text, 0, 0) + "'" +
                    //            ",KeiyakuBikou = " + "N'" + GlobalMethod.ChangeSqlText(ca_tbl01_txtBiko.Text, 0, 0) + "'" +
                    //            ",KeiyakuShosha = " + (ca_tbl01_chkCaSyosya.Checked ? 1 : 0) +
                    //            ",KeiyakuTokkiShiyousho = " + (ca_tbl01_chkSiyosyo.Checked ? 1 : 0) +
                    //            ",KeiyakuMitsumorisho = " + (ca_tbl01_chkMitumorisyo.Checked ? 1 : 0) +
                    //            ",KeiyakuTanpinChousaMitsumorisho = " + (ca_tbl01_chkTanpinTyosa.Checked ? 1 : 0) +
                    //            ",KeiyakuSonota = " + (ca_tbl01_chkOther.Checked ? 1 : 0) +
                    //            ",KeiyakuSonotaNaiyou = " + "N'" + GlobalMethod.ChangeSqlText(ca_tbl01_txtOtherBiko.Text, 0, 0) + "'" +
                    //            ",KeiyakuZentokinUkewatashibi = " + Get_DateTimePicker("ca_tbl07_dtpRequst6") +
                    //            ",KeiyakuZentokin = " + getNumToDb(ca_tbl07_txtRequst6.Text) +
                    //            ",Keiyakukeiyakukingakukei = " + getNumToDb(ca_tbl01_txtJyutakuAmt.Text) +
                    //            ",KeiyakuBetsuKeiyakuKingaku = " + getNumToDb(ca_tbl01_txtJyutakuGaiAmt.Text) +
                    //            ",KeiyakuSeikyuubi1 = " + " " + Get_DateTimePicker("ca_tbl07_dtpRequst1") + "" +
                    //            ",KeiyakuSeikyuuKingaku1 = " + getNumToDb(ca_tbl07_txtRequst1.Text) +
                    //            ",KeiyakuSeikyuubi2 = " + " " + Get_DateTimePicker("ca_tbl07_dtpRequst2") + "" +
                    //            ",KeiyakuSeikyuuKingaku2 = " + getNumToDb(ca_tbl07_txtRequst2.Text) +
                    //            ",KeiyakuSeikyuubi3 = " + " " + Get_DateTimePicker("ca_tbl07_dtpRequst3") + "" +
                    //            ",KeiyakuSeikyuuKingaku3 = " + getNumToDb(ca_tbl07_txtRequst3.Text) +
                    //            ",KeiyakuSeikyuubi4 = " + " " + Get_DateTimePicker("ca_tbl07_dtpRequst4") + "" +
                    //            ",KeiyakuSeikyuuKingaku4 = " + getNumToDb(ca_tbl07_txtRequst4.Text) +
                    //            ",KeiyakuSeikyuubi5 = " + " " + Get_DateTimePicker("ca_tbl07_dtpRequst5") + "" +
                    //            ",KeiyakuSeikyuuKingaku5 = " + getNumToDb(ca_tbl07_txtRequst5.Text) +
                    //            ",KeiyakuSakuseiKubunID = " + "N'" + ca_tbl01_cmbAnkenKubun.SelectedValue + "'" +
                    //            ",KeiyakuSakuseiKubun = " + "N'" + ca_tbl01_cmbAnkenKubun.Text + "'" +
                    //            ",KeiyakuGyoumuKubun = " + "N'" + ca_tbl01_cmbCaKubun.SelectedValue + "'" +
                    //            ",KeiyakuGyoumuMei = " + "N'" + ca_tbl01_cmbCaKubun.Text + "'" +
                    //            ",KeiyakuKianzumi = " + (ca_tbl01_chkKian.Checked ? 1 : 0) +
                    //            ",KeiyakuHachuushaMei = " + "N'" + ca_tbl01_txtOrderKamei.Text + "'" +
                    //            ",KeiyakuHaibunChoZeinuki = " + getNumToDb(ca_tbl02_AftCaBm_numAmt1.Text) +
                    //            ",KeiyakuHaibunJoZeinuki = " + getNumToDb(ca_tbl02_AftCaBm_numAmt2.Text) +
                    //            ",KeiyakuHaibunJosysZeinuki = " + getNumToDb(ca_tbl02_AftCaBm_numAmt3.Text) +
                    //            ",KeiyakuHaibunKeiZeinuki = " + getNumToDb(ca_tbl02_AftCaBm_numAmt4.Text) +
                    //            ",KeiyakuHaibunZeinukiKei = " + getNumToDb(ca_tbl02_AftCaBm_numAmtAll.Text) +
                    //            ",KeiyakuUriageHaibunCho  = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt1.Text) +
                    //            ",KeiyakuUriageHaibunJo   = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt2.Text) +
                    //            ",KeiyakuUriageHaibunJosys  = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt3.Text) +
                    //            ",KeiyakuUriageHaibunKei  = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt4.Text) +
                    //            ",KeiyakuUriageHaibunGoukei = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmtAll.Text) +
                    //            ",KeiyakuTankeiMikomiCho  = " + getNumToDb(ca_tbl03_numAmt1.Text) +
                    //            ",KeiyakuTankeiMikomiJo  = " + getNumToDb(ca_tbl03_numAmt2.Text) +
                    //            ",KeiyakuTankeiMikomiJosys  = " + getNumToDb(ca_tbl03_numAmt3.Text) +
                    //            ",KeiyakuTankeiMikomiKei  = " + getNumToDb(ca_tbl03_numAmt4.Text) +
                    //            ",KeiyakuKurikoshiCho  = " + getNumToDb(ca_tbl04_numKurikosiAmt1.Text) +
                    //            ",KeiyakuKurikoshiJo  = " + getNumToDb(ca_tbl04_numKurikosiAmt2.Text) +
                    //            ",KeiyakuKurikoshiJosys  = " + getNumToDb(ca_tbl04_numKurikosiAmt3.Text) +
                    //            ",KeiyakuKurikoshiKei  = " + getNumToDb(ca_tbl04_numKurikosiAmt4.Text) +
                    //            ",KeiyakuUpdateProgram = " + "'ChangeKianEntry'" +
                    //            ",KeiyakuUpdateDate = " + "GETDATE()" +
                    //            ",KeiyakuUpdateUser = " + "N'" + UserInfos[0] + "'" +
                    //            ",KeiyakuRIBCYouTankaDataMoushikomisho = " + (ca_tbl01_chkRibcSyo.Checked ? 1 : 0) +
                    //            ",KeiyakuSashaKeiyu = " + (ca_tbl01_chkSasya.Checked ? 1 : 0) +
                    //            ",KeiyakuRIBCYouTankaData = " + (ca_tbl01_chkRibcAri.Checked ? 1 : 0) +
                    //            ",KeiyakuSaiitakuKinshiUmu = N'" + (IsNotSelected(ca_tbl01_cmbKinsiUmu) ? "0" : ca_tbl01_cmbKinsiUmu.SelectedValue.ToString()) + "' " +
                    //            ",KeiyakuSaiitakuKinshiNaiyou = N'" + (IsNotSelected(ca_tbl01_cmbKinsiNaiyo) ? "0" : ca_tbl01_cmbKinsiNaiyo.SelectedValue.ToString()) + "' " +
                    //            ",KeiyakuSaiitakuSonotaNaiyou = N'" + ca_tbl01_txtOtherNaiyo.Text + "' " +
                    //        " WHERE AnkenJouhouID = " + ankenNo2;
                    //result = cmd.ExecuteNonQuery();

                    cmd.CommandText = "DELETE FROM RibcJouhou " +
                            " WHERE RibcID = " + ankenNo2;
                    cmd.ExecuteNonQuery();
                    int cnt = 0;
                    string RibcKoukiStart;
                    string RibcNouhinbi;
                    string RibcSeikyubi;
                    string RibcNyukinyoteibi;
                    string RibcKubun;
                    C1FlexGrid c1FlexGrid4 = ca_tbl06_c1FlexGrid;
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
                            if (GetInt(ca_tbl01_cmbSalesYear.SelectedValue.ToString()) >= 2021)
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
                    updateGyoumuHaibun30(cmd, ankenNo2);

                    // 窓口の案件情報IDを最新に置き換える
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
                    "MadoguchiKanriGijutsusha = " + "N'" + ca_tbl05_txtKanriCD.Text + "' " +
                    " WHERE AnkenJouhouID = " + ankenNo2;
                    cmd.ExecuteNonQuery();
                }

                transaction.Commit();

                rtn = true;
            }
            catch (Exception)
            {
                transaction.Rollback();
                throw;
            }
            return rtn;
        }

        #region 新規モード処理メソッド ------------
        /// <summary>
        /// 案件情報テーブルに新規追加時Values設定
        /// </summary>
        /// <returns></returns>
        private string getInsAnkenVals(string ankenID, string ankenNo, string ankenFolder, string BushoShozokuChou)
        {
            StringBuilder sSql = new StringBuilder();

            sSql.Append(ankenID);                                                            // AnkenJouhouID
            sSql.Append("    , '" + base_tbl02_cmbAnkenKubun.SelectedValue + "'");                                // AnkenSakuseiKubun
            sSql.Append("    , " + base_tbl03_cmbKokiSalesYear.SelectedValue);                                   // AnkenUriageNendo
            sSql.Append("    , '" + base_tbl02_txtKeikakuNo.Text.Trim() + "'");                                         // AnkenKeikakuBangou
            sSql.Append("    , N'" + ankenNo + "'");                                                 // AnkenAnkenBangou
            sSql.Append("    , N'" + base_tbl02_txtJyutakuNo.Text + "'");                                           // AnkenJutakuBangou
            sSql.Append("    , N'" + base_tbl02_txtJyutakuEdNo.Text.Trim() + "' ");                                                // AnkenJutakuBangouEda
            sSql.Append("    , GETDATE() ");                                            // AnkenTourokubi
            sSql.Append("    , N'" + base_tbl02_cmbJyutakuKasyoSibu.Text + "'");                                          // AnkenJutakushibu
            sSql.Append("    , N'" + base_tbl02_cmbJyutakuKasyoSibu.SelectedValue + "'");                                  // AnkenJutakubushoCD
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(ankenFolder, 0, 0) + "'");           // AnkenKeiyakusho
            sSql.Append("    , '" + base_tbl02_txtKeiyakuTantoCD.Text + "'");                                         // AnkenTantoushaCD
            sSql.Append("    , N'" + base_tbl02_txtKeiyakuTanto.Text + "'");                                           // AnkenTantoushaMei
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl03_txtGyomuName.Text, 0, 0) + "'");         // AnkenGyoumuMei
            sSql.Append("    , " + base_tbl03_cmbKeiyakuKubun.SelectedValue + " ");                                   // AnkenGyoumuKubun
            sSql.Append("    , N'" + Get_GyoumuKubunCD(base_tbl03_cmbKeiyakuKubun.SelectedValue.ToString()) + "'");   // AnkenGyoumuKubunCD
            sSql.Append("    , N'" + base_tbl03_cmbKeiyakuKubun.Text + "' ");                                         // AnkenGyoumuKubunMei
            sSql.Append("    , N'" + base_tbl10_cmbNyusatuHosiki.SelectedValue + "' " );                                  // AnkenNyuusatsuHoushiki
            sSql.Append("    , " + Get_DateTimePicker("base_tbl10_dtpNyusatuDt"));                        // AnkenNyuusatsuYoteibi
            sSql.Append("    , N'" + base_tbl04_txtOrderCd.Text + "'" );                                         // AnkenHachushaCD
            sSql.Append("    , N'" + base_tbl04_txtOrderName.Text + "'");                                           // AnkenHachuushaMei
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl04_txtOrderKamei.Text, 0, 0) + "'");        // AnkenHachushaKaMei
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl04_txtOrderName.Text, 0, 0) + "　" + GlobalMethod.ChangeSqlText(base_tbl04_txtOrderKamei.Text, 0, 0) + "'"); // AnkenHachuushaKaMei
            
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtBusho.Text, 0, 0) + "'" );          // AnkenHachuushaIraibusho
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtTanto.Text, 0, 0) + "'" );          // AnkenHachuushaTantousha
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtTel.Text, 0, 0) + "'" );          // AnkenHachuushaTEL
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtFax.Text, 0, 0) + "'" );          // AnkenHachuushaFAX
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtEmail.Text, 0, 0) + "'" );          // AnkenHachuushaMail
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtZip.Text, 0, 0) + "'" );          // AnkenHachuushaIraiYuubin
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtAddress.Text, 0, 0) + "'" );          // AnkenHachuushaIraiJuusho

            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtBusho.Text, 0, 0) + "'");          // AnkenHachuushaKeiyakuBusho
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtTanto.Text, 0, 0) + "'");          // AnkenHachuushaKeiyakuTantou
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtTel.Text, 0, 0) + "'");          // AnkenHachuushaKeiyakuTEL
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtFax.Text, 0, 0) + "'");          // AnkenHachuushaKeiyakuFAX
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtEmail.Text, 0, 0) + "'");          // AnkenHachuushaKeiyakuMail
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtZip.Text, 0, 0) + "'");          // AnkenHachuushaKeiyakuYuubin
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtAddress.Text, 0, 0) + "'");          // AnkenHachuushaKeiyakuJuusho
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtOrderYakusyoku.Text, 0, 0) + "'" );          // AnkenHachuuDaihyouYakushoku
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtOrderSimei.Text, 0, 0) + "'");         // AnkenHachuuDaihyousha
            sSql.Append("    , N'" + base_tbl09_cmbSankomitumori.SelectedValue + "'" );                   // AnkenToukaiSankouMitsumori
            sSql.Append("    , N'" + base_tbl09_cmbOrderIyoku.SelectedValue + "'");      // AnkenToukaiJyutyuIyoku
            sSql.Append("    , " + getNumToDb(base_tbl09_numSankomitumoriAmt.Text) + " ");      // AnkenToukaiSankouMitsumoriGaku
            sSql.Append("    , 'InsertEntry'");      // AnkenCreateProgram
            sSql.Append("    , GETDATE() ");     // AnkenCreateDate
            sSql.Append("    , N'" + UserInfos[0] + "'");                     // AnkenCreateUser
            sSql.Append("    , GETDATE() ");       // AnkenUpdateDate
            sSql.Append("    , N'" + UserInfos[0] + "'");         // AnkenUpdateUser
            sSql.Append("    , 0" );         // AnkenDeleteFlag
            sSql.Append("    , 1");        // AnkenSaishinFlg
            sSql.Append("    , null" );         // AnkenGyoumuKanrishaCD
            sSql.Append("    , null" );         // AnkenMadoguchiTantoushaCD
            sSql.Append("    , null");        // GyoumuKanrishaCD
            sSql.Append("    , 0");        // AnkenKaisuu
            sSql.Append("    , N'" + base_tbl03_cmbKokiStartYear.SelectedValue.ToString() + "'");         // AnkenKoukiNendo
            sSql.Append("    , '特になし'" );         // AnkenKokyakuHyoukaComment
            sSql.Append("    , '特になし'");        // AnkenToukaiHyoukaComment
            sSql.Append("    , N'" + BushoShozokuChou + "'");                       // AnkenGyoumuKanrisha
            sSql.Append("    , " + Get_DateTimePicker("base_tbl03_dtpKokiFrom"));   // AnkenKeiyakuKoukiKaishibi
            sSql.Append("    , " + Get_DateTimePicker("base_tbl03_dtpKokiTo"));     // AnkenKeiyakuKoukiKanryoubi
            sSql.Append("    , " + Get_DateTimePicker("base_tbl01_dtpDtCa"));   // AnkenKeiyakuDat
            sSql.Append("    , " + (base_tbl01_chkKeiyaku.Checked ? "1" : "0"));   // AnkenKeiyakuCheck
            sSql.Append("    , " + Get_DateTimePicker("base_tbl01_dtpDtBid"));   // AnkenNyuusatuDate
            sSql.Append("    , " + (base_tbl01_chkNyusatu.Checked ? "1" : "0"));   // AnkenNyuusatuCheck
            if (base_tbl01_dtpDtPrior.CustomFormat != "")
            {
                sSql.Append("    , " + Get_DateTimePicker("base_tbl01_dtpDtBid"));   // AnkenJizenDashinDate
                sSql.Append("    , 1");   // AnkenJizenDashinCheck
            }
            else
            {
                sSql.Append("    , " + Get_DateTimePicker("base_tbl01_dtpDtPrior"));   // AnkenJizenDashinDate
                sSql.Append("    , " + (base_tbl01_chkJizendasin.Checked ? "1" : "0"));   // AnkenJizenDashinCheck
            }
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl03_txtAnkenMemo.Text, 0, 0) + "'");   // AnkenAnkenMemoKihon
            // No.1533 削除
            //sSql.Append("    , " + ((IsNotSelected(base_tbl07_3_cmbOen)) ? "null" : base_tbl07_3_cmbOen.SelectedValue.ToString()));      // AnkenOueniraiUmu --応援依頼の有無
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl07_3_txtOenMemo.Text, 0, 0) + "'");   // AnkenOuenIraiMemo --応援依頼メモ
            sSql.Append("    , " + Get_DateTimePicker("base_tbl09_dtpJizenDasinIraiDt"));   // AnkenJizenDashinIraibi --事前打診依頼日
            sSql.Append("    , " + Get_DateTimePicker("base_tbl09_dtpOrderYoteiDt"));   // AnkenHachuuYoteiMikomibi --発注予定・見込日
            sSql.Append("    , " + ((IsNotSelected(base_tbl09_cmbNotOrderStats)) ? "null" : base_tbl09_cmbNotOrderStats.SelectedValue.ToString()));      // AnkenMihachuuJoukyou --未発注状況
            sSql.Append("    , " + ((IsNotSelected(base_tbl09_cmbNotOrderReason)) ? "null" : base_tbl09_cmbNotOrderReason.SelectedValue.ToString()));      // AnkenHachuunashiRiyuu --「発注なし」の理由
            sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl09_txtOthenComment.Text, 0, 0) + "'");   // AnkenSonotaNaiyou --「その他」の内容
            sSql.Append("    , " + ((IsNotSelected(base_tbl10_cmbTokaiOsatu)) ? "null" : base_tbl10_cmbTokaiOsatu.SelectedValue.ToString()));        // AnkenToukaiOusatu --当会応札

            return sSql.ToString();
        }

        /// <summary>
        /// 応援依頼先登録処理、赤伝／黒伝作成処理
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenID"></param>
        /// <param name="iType">0:赤伝、1:黒伝、8:更新、9:新規登録</param>
        /// <returns></returns>
        private int createAnkenOuenIraisaki(SqlCommand cmd, string ankenID, int iType = 9)
        {
            if (iType == 8)
            {
                // 更新時、削除してからInsertする
                cmd.CommandText = "DELETE AnkenOuenIraisaki WHERE AnkenJouhouID = '" + ankenID + "' ";
                cmd.ExecuteNonQuery();
            }
            int iRow = 1;
            StringBuilder sSql = new StringBuilder();
            sSql.Append("INSERT INTO AnkenOuenIraisaki( ");
            sSql.Append("    AnkenJouhouID");
            sSql.Append("    , AnkenOuenIraisakiID");
            sSql.Append("    , OueniraisakiCD");
            if(iType != 9 && iType != 8)
            {
                //新規登録以外
                sSql.Append("     ) SELECT ");
                sSql.Append("    " + ankenID);
                sSql.Append("    , AnkenOuenIraisakiID");
                sSql.Append("    , OueniraisakiCD");
                sSql.Append(" FROM AnkenOuenIraisaki WHERE AnkenJouhouID = " + AnkenID);
            }
            else
            {
                iRow = 0;
                sSql.Append(") VALUES ");
                if(base_tbl07_3_tblOenIrai1.Height > 0)
                {
                    foreach (Control child in base_tbl07_3_tblOenIrai1.Controls)
                    {
                        //特定のコントロール型内部の子情報は取得しない
                        if (child is System.Windows.Forms.CheckBox)
                        {
                            System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)child;
                            if (chk.Checked)
                            {
                                iRow++;
                                sSql.Append("(");
                                sSql.Append(ankenID);
                                sSql.Append(", " + iRow.ToString());
                                sSql.Append(", " + chk.Tag.ToString());
                                sSql.Append("),");
                            }
                        }
                    }

                }
                if (base_tbl07_3_tblOenIrai2.Height > 0)
                {
                    foreach (Control child in base_tbl07_3_tblOenIrai2.Controls)
                    {
                        //特定のコントロール型内部の子情報は取得しない
                        if (child is System.Windows.Forms.CheckBox)
                        {
                            System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)child;
                            if (chk.Checked)
                            {
                                iRow++;
                                sSql.Append("(");
                                sSql.Append(ankenID);
                                sSql.Append(", " + iRow.ToString());
                                sSql.Append(", " + chk.Tag.ToString());
                                sSql.Append("),");
                            }
                        }
                    }

                }
                if (base_tbl07_3_tblOenIrai3.Height > 0)
                {
                    foreach (Control child in base_tbl07_3_tblOenIrai3.Controls)
                    {
                        //特定のコントロール型内部の子情報は取得しない
                        if (child is System.Windows.Forms.CheckBox)
                        {
                            System.Windows.Forms.CheckBox chk = (System.Windows.Forms.CheckBox)child;
                            if (chk.Checked)
                            {
                                iRow++;
                                sSql.Append("(");
                                sSql.Append(ankenID);
                                sSql.Append(", " + iRow.ToString());
                                sSql.Append(", " + chk.Tag.ToString());
                                sSql.Append("),");
                            }
                        }
                    }

                }
                if(iRow > 0)
                {
                    // 最後の「,を削除する」
                    string sql = sSql.ToString().TrimEnd(',');
                    sSql.Clear();
                    sSql.Append(sql);
                }
                sSql.Append(";");
            }

            if (iRow > 0)
            {
                cmd.CommandText = sSql.ToString();
                Console.WriteLine(cmd.CommandText);
                return cmd.ExecuteNonQuery();
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// 過去案件リスト登録処理、赤伝／黒伝作成処理
        /// </summary>
        /// <param name="c1FlexGrid1">null:赤伝／黒伝作成、以外：リスト登録</param>
        /// <param name="cmd"></param>
        /// <param name="ankenID"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝 </param>
        private int createAnkenJouhouZenkaiRakusatsu(C1FlexGrid c1FlexGrid1, SqlCommand cmd, string ankenID, int iType = 0)
        {
            if (c1FlexGrid1 != null)
            {
                // 削除してからInsert実行
                cmd.CommandText = "DELETE FROM AnkenJouhouZenkaiRakusatsu " +
                                       " WHERE AnkenJouhouID = " + ankenID;
                cmd.ExecuteNonQuery();

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
                                             ", N'" + c1FlexGrid1.Rows[i][16].ToString() + "'" +    // [AnkenZenkaiRakusatsuID] [int] NOT NULL,
                                             ", N'" + c1FlexGrid1.Rows[i][3].ToString() + "'" +     // [AnkenZenkaiAnkenBangou] [nvarchar](40) NULL,
                                             ", N'" + c1FlexGrid1.Rows[i][4].ToString() + "'" +     // [AnkenZenkaiJutakuBangou] [nvarchar](40) NULL,
                                             ", N'" + c1FlexGrid1.Rows[i][5].ToString() + "'" +     // [AnkenZenkaiJutakuEdaban] [nvarchar](50) NULL,
                                             ", N'" + c1FlexGrid1.Rows[i][6].ToString() + "'" +     // [AnkenZenkaiGyoumuMei] [nvarchar](150) NULL,

                                             ", N'" + c1FlexGrid1.Rows[i][7].ToString() + "'" +     // [AnkenZenkaiRakusatsusha] [nvarchar](50) NULL,
                                             ", " + AnkenZenkaiRakusatsushaID +            // [AnkenZenkaiRakusatsushaID] [int] NULL,
                                             ", " + AnkenZenkaiJutakuKingaku +      // [AnkenZenkaiJutakuKingaku] [money] NULL,
                                             ", " + KeiyakuZenkaiRakusatsushaID +             // [KeiyakuZenkaiRakusatsushaID] [int] NULL,        
                                             ", N'" + c1FlexGrid1.Rows[i][15].ToString() + "'" +    // [AnkenZenkaiKyougouKigyouCD] [nvarchar](24) NULL,
                                             ", " + AnkenZenkaiAnkenJouhouID +                // [AnkenZenkaiAnkenJouhouID] [decimal](16, 0) NULL,
                                             ", " + AnkenZenkaiJutakuZeinuki +      // [AnkenZenkaiJutakuZeinuki] [money] NULL,
                                            ")";

                        Console.WriteLine(cmd.CommandText);
                        cmd.ExecuteNonQuery();
                    }
                }
                return 0;
            }
            else
            {
                StringBuilder sSql = new StringBuilder();

                sSql.Append("INSERT INTO AnkenJouhouZenkaiRakusatsu ( ");
                sSql.Append("    AnkenJouhouID ");
                sSql.Append("    ,AnkenZenkaiJutakuKingaku ");
                sSql.Append("    ,AnkenZenkaiRakusatsuID ");
                sSql.Append("    ,AnkenZenkaiJutakuBangou ");
                sSql.Append("    ,AnkenZenkaiJutakuEdaban ");
                sSql.Append("    ,AnkenZenkaiAnkenBangou ");
                sSql.Append("    ,AnkenZenkaiRakusatsushaID ");
                sSql.Append("    ,AnkenZenkaiRakusatsusha ");
                sSql.Append("    ,AnkenZenkaiGyoumuMei ");
                sSql.Append("    ,AnkenZenkaiKyougouKigyouCD ");
                sSql.Append("    ,AnkenZenkaiJutakuZeinuki ");
                sSql.Append("    ,KeiyakuZenkaiRakusatsushaID ");
                sSql.Append("     ) SELECT ");
                sSql.Append(ankenID);
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append("AnkenZenkaiJutakuKingaku ");
                sSql.Append("    ,AnkenZenkaiRakusatsuID ");
                sSql.Append("    ,AnkenZenkaiJutakuBangou ");
                sSql.Append("    ,AnkenZenkaiJutakuEdaban ");
                sSql.Append("    ,AnkenZenkaiAnkenBangou ");
                sSql.Append("    ,AnkenZenkaiRakusatsushaID ");
                sSql.Append("    ,AnkenZenkaiRakusatsusha ");
                sSql.Append("    ,AnkenZenkaiGyoumuMei ");
                sSql.Append("    ,AnkenZenkaiKyougouKigyouCD ");
                sSql.Append("    ,AnkenZenkaiJutakuZeinuki ");
                sSql.Append("    ,KeiyakuZenkaiRakusatsushaID ");
                sSql.Append(" FROM AnkenJouhouZenkaiRakusatsu WHERE AnkenJouhouZenkaiRakusatsu.AnkenJouhouID = " + AnkenID);

                cmd.CommandText = sSql.ToString();
                Console.WriteLine(cmd.CommandText);
                return cmd.ExecuteNonQuery();
            }

        }

        /// <summary>
        /// 契約情報エントリ登録処理
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenID"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝 、9:新規登録</param>
        private int createKeiyakuJouhouEntory(SqlCommand cmd, string ankenID, int iType = 9)
        {
            StringBuilder sSql = new StringBuilder();

            // Cols
            sSql.Append("INSERT INTO KeiyakuJouhouEntory( ");
            sSql.Append("    AnkenJouhouID");
            sSql.Append("    , KeiyakuJouhouEntoryID");
            sSql.Append("    , KeiyakuCreateDate");
            sSql.Append("    , KeiyakuCreateUser");
            sSql.Append("    , KeiyakuCreateProgram");
            sSql.Append("    , KeiyakuUpdateDate");
            sSql.Append("    , KeiyakuUpdateUser");
            sSql.Append("    , KeiyakuUpdateProgram");
            sSql.Append("    , KeiyakuDeleteFlag");
            sSql.Append("    , KeiyakuSakuseiKubunID");
            sSql.Append("    , KeiyakuSakuseiKubun");
            sSql.Append("    , KeiyakuHachuushaMei");
            sSql.Append("    , KeiyakuGyoumuKubun");
            sSql.Append("    , KeiyakuGyoumuMei");
            sSql.Append("    , JutakuBushoCD");
            sSql.Append("    , KeiyakuTantousha");
            sSql.Append("    , KeiyakuUriageHaibunCho1");
            sSql.Append("    , KeiyakuUriageHaibunCho2");
            sSql.Append("    , KeiyakuUriageHaibunJo1");
            sSql.Append("    , KeiyakuUriageHaibunJo2");
            sSql.Append("    , KeiyakuUriageHaibunJosys1");
            sSql.Append("    , KeiyakuUriageHaibunJosys2");
            sSql.Append("    , KeiyakuUriageHaibunKei1");
            sSql.Append("    , KeiyakuUriageHaibunKei2");
            sSql.Append("    , KeiyakuUriageHaibunChoGoukei");
            sSql.Append("    , KeiyakuUriageHaibunJoGoukei");
            sSql.Append("    , KeiyakuUriageHaibunJosysGoukei");
            sSql.Append("    , KeiyakuUriageHaibunKeiGoukei");
            sSql.Append("    , KeiyakuUriageHaibunGoukei");
            sSql.Append("    , KeiyakuKeiyakuTeiketsubi");

            if (iType != 9)
            {
                sSql.Append("    ,KeiyakuKeiyakuKingaku ");
                sSql.Append("    ,KeiyakuZeikomiKingaku ");
                sSql.Append("    ,KeiyakuuchizeiKingaku ");
                sSql.Append("    ,KeiyakuUriageHaibunCho ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuCho1 ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuCho2 ");
                sSql.Append("    ,KeiyakuUriageHaibunJo ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuJo1 ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuJo2 ");
                sSql.Append("    ,KeiyakuUriageHaibunJosys ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuJosys1 ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuJosys2 ");
                sSql.Append("    ,KeiyakuUriageHaibunKei ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuKei1 ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuKei2 ");
                sSql.Append("    ,KeiyakuZentokin ");
                sSql.Append("    ,KeiyakuSeikyuuKingaku1 ");
                sSql.Append("    ,KeiyakuSeikyuuKingaku2 ");
                sSql.Append("    ,KeiyakuSeikyuuKingaku3 ");
                sSql.Append("    ,KeiyakuSeikyuuKingaku4 ");
                sSql.Append("    ,KeiyakuSeikyuuKingaku5 ");
                sSql.Append("    ,KeiyakuBetsuKeiyakuKingaku ");
                sSql.Append("    ,KeiyakuKeiyakuKingakuKei ");
                sSql.Append("    ,KeiyakuHaibunChoZeinuki ");
                sSql.Append("    ,KeiyakuHaibunJoZeinuki ");
                sSql.Append("    ,KeiyakuHaibunJosysZeinuki ");
                sSql.Append("    ,KeiyakuHaibunKeiZeinuki ");
                sSql.Append("    ,KeiyakuHaibunZeinukiKei ");
                sSql.Append("    ,KeiyakuSakuseibi ");
                sSql.Append("    ,KeiyakuJutakubangou ");
                sSql.Append("    ,KeiyakuEdaban ");
                sSql.Append("    ,KeiyakuKianzumi ");
                sSql.Append("    ,KeiyakuNyuusatsuYoteibi ");
                sSql.Append("    ,KeiyakuKeiyakuKoukiKaishibi ");
                sSql.Append("    ,KeiyakuKeiyakuKoukiKanryoubi ");
                sSql.Append("    ,KeiyakuShouhizeiritsu ");
                sSql.Append("    ,KeiyakuRIBCKeishiki ");
                sSql.Append("    ,KeiyakuHenkoukanryoubi ");
                sSql.Append("    ,KeiyakuHenkouChuushiRiyuu ");
                sSql.Append("    ,KeiyakuBikou ");
                sSql.Append("    ,KeiyakuShosha ");
                sSql.Append("    ,KeiyakuTokkiShiyousho ");
                sSql.Append("    ,KeiyakuMitsumorisho ");
                sSql.Append("    ,KeiyakuTanpinChousaMitsumorisho ");
                sSql.Append("    ,KeiyakuSonota ");
                sSql.Append("    ,KeiyakuSonotaNaiyou ");
                sSql.Append("    ,KeiyakuSeikyuubi ");
                sSql.Append("    ,KeiyakuKeiyakusho ");
                sSql.Append("    ,KeiyakuZentokinUkewatashibi ");
                sSql.Append("    ,KeiyakuSeikyuusaki ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiS1 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiE1 ");
                sSql.Append("    ,KeiyakuSeikyuubi1 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiS2 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiE2 ");
                sSql.Append("    ,KeiyakuSeikyuubi2 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiS3 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiE3 ");
                sSql.Append("    ,KeiyakuSeikyuubi3 ");
                sSql.Append("    ,KeiyakuKankeibusho1 ");
                sSql.Append("    ,KeiyakuKankeibusho2 ");
                sSql.Append("    ,KeiyakuKankeibusho3 ");
                sSql.Append("    ,KeiyakuKankeibusho4 ");
                sSql.Append("    ,KeiyakuKankeibusho5 ");
                sSql.Append("    ,KeiyakuKankeibusho6 ");
                sSql.Append("    ,KeiyakuKankeibusho7 ");
                sSql.Append("    ,KeiyakuKankeibusho8 ");
                sSql.Append("    ,KeiyakuKankeibusho9 ");
                sSql.Append("    ,KeiyakuKankeibusho10 ");
                sSql.Append("    ,KeiyakuKankeibusho11 ");
                sSql.Append("    ,KeiyakuKankeibusho12 ");
                sSql.Append("    ,KeiyakuKankeibusho14 ");
                sSql.Append("    ,KeiyakuKankeibusho15 ");
                sSql.Append("    ,KeiyakuKankeibusho13 ");
                sSql.Append("    ,KeiyakuNyuukinYoteibi ");
                sSql.Append("    ,KeiyakuUriageHaibunCho1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunCho2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunJo1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunJo2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunJosys1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunJosys2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunKei1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunKei2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC1 ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuRIBC1 ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC2 ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuRIBC2 ");
                sSql.Append("    ,KeiyakuSeikyuubi4 ");
                sSql.Append("    ,KeiyakuSeikyuubi5 ");
                sSql.Append("    ,KeiyakuTankeiMikomiCho ");
                sSql.Append("    ,KeiyakuTankeiMikomiJo ");
                sSql.Append("    ,KeiyakuTankeiMikomiJosys ");
                sSql.Append("    ,KeiyakuTankeiMikomiKei ");
                sSql.Append("    ,KeiyakuKurikoshiCho ");
                sSql.Append("    ,KeiyakuKurikoshiJo ");
                sSql.Append("    ,KeiyakuKurikoshiJosys ");
                sSql.Append("    ,KeiyakuKurikoshiKei ");
                sSql.Append("    ,KeiyakuRIBCYouTankaDataMoushikomisho ");
                sSql.Append("    ,KeiyakuSashaKeiyu ");
                sSql.Append("    ,KeiyakuRIBCYouTankaData ");
                sSql.Append("    ,KeiyakuSaiitakuSonotaNaiyou ");
                sSql.Append("    ,KeiyakuSaiitakuKinshiNaiyou ");
                sSql.Append("    ,KeiyakuSaiitakuKinshiUmu ");
                sSql.Append("    ,KeiyakuAnkenMemoKeiyaku ");
            }
            // Values
            if (iType == 9)
            {
                sSql.Append(") VALUES ( ");
                sSql.Append("   " + ankenID);
                sSql.Append("    , " + ankenID);
                sSql.Append("    , GETDATE()");
                sSql.Append("    , N'" + UserInfos[0] + "'");
                sSql.Append("    , 'InsertEntry'");
                sSql.Append("    , GETDATE()");
                sSql.Append("    , N'" + UserInfos[0] + "'");
                sSql.Append("    , null");
                sSql.Append("    , 0");
                sSql.Append("    , N'" + base_tbl02_cmbAnkenKubun.SelectedValue + "'");
                sSql.Append("    , N'" + base_tbl02_cmbAnkenKubun.Text + "'");
                sSql.Append("    , N'" + base_tbl04_txtOrderName.Text + "'");
                sSql.Append("    , N'" + base_tbl03_cmbKeiyakuKubun.SelectedValue + "'");
                sSql.Append("    , N'" + base_tbl03_cmbKeiyakuKubun.Text + "'");
                sSql.Append("    , null");
                sSql.Append("    , null");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , 0");
                sSql.Append("    , null");
                sSql.Append(")");
            }
            else
            {
                // 赤伝／黒伝
                sSql.Append("     ) SELECT ");
                sSql.Append("    " + ankenID);
                sSql.Append("    ," + ankenID);
                sSql.Append("    ,GETDATE() ");
                sSql.Append("    ,N'" + UserInfos[0] + "' ");
                sSql.Append("    ,'ChangeKianEntry' ");
                sSql.Append("    ,GETDATE() ");
                sSql.Append("    ,N'" + UserInfos[0] + "' ");
                sSql.Append("    ,'ChangeKianEntry' ");
                if (iType == 0)
                {
                    sSql.Append("    ,0 ");
                }
                else if (iType == 1)
                {
                    sSql.Append("    ,KeiyakuDeleteFlag ");
                }
                else { sSql.Append("    ,1 "); }
                sSql.Append("    ,KeiyakuSakuseiKubunID ");
                sSql.Append("    ,KeiyakuSakuseiKubun ");
                sSql.Append("    ,KeiyakuHachuushaMei ");
                sSql.Append("    ,KeiyakuGyoumuKubun ");
                sSql.Append("    ,KeiyakuGyoumuMei ");
                sSql.Append("    ,JutakuBushoCD ");
                sSql.Append("    ,KeiyakuTantousha ");
                sSql.Append("    ,KeiyakuUriageHaibunCho1 ");
                sSql.Append("    ,KeiyakuUriageHaibunCho2 ");
                sSql.Append("    ,KeiyakuUriageHaibunJo1 ");
                sSql.Append("    ,KeiyakuUriageHaibunJo2 ");
                sSql.Append("    ,KeiyakuUriageHaibunJosys1 ");
                sSql.Append("    ,KeiyakuUriageHaibunJosys2 ");
                sSql.Append("    ,KeiyakuUriageHaibunKei1 ");
                sSql.Append("    ,KeiyakuUriageHaibunKei2 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunChoGoukei ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunJoGoukei ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunJosysGoukei ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunKeiGoukei ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGoukei ");
                sSql.Append("    ,KeiyakuKeiyakuTeiketsubi ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuKeiyakuKingaku ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuZeikomiKingaku ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuuchizeiKingaku ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunCho ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGakuCho1 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGakuCho2 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunJo ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGakuJo1 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGakuJo2 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunJosys ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGakuJosys1 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGakuJosys2 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunKei ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGakuKei1 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuUriageHaibunGakuKei2 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuZentokin ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuSeikyuuKingaku1 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuSeikyuuKingaku2 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuSeikyuuKingaku3 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuSeikyuuKingaku4 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuSeikyuuKingaku5 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuBetsuKeiyakuKingaku ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuKeiyakuKingakuKei ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuHaibunChoZeinuki ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuHaibunJoZeinuki ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuHaibunJosysZeinuki ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuHaibunKeiZeinuki ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuHaibunZeinukiKei ");
                sSql.Append("    ,KeiyakuSakuseibi ");
                sSql.Append("    ,KeiyakuJutakubangou ");
                sSql.Append("    ,KeiyakuEdaban ");
                sSql.Append("    ,KeiyakuKianzumi ");
                sSql.Append("    ,KeiyakuNyuusatsuYoteibi ");
                sSql.Append("    ,KeiyakuKeiyakuKoukiKaishibi ");
                sSql.Append("    ,KeiyakuKeiyakuKoukiKanryoubi ");
                sSql.Append("    ,KeiyakuShouhizeiritsu ");
                sSql.Append("    ,KeiyakuRIBCKeishiki ");
                sSql.Append("    ,KeiyakuHenkoukanryoubi ");
                sSql.Append("    ,KeiyakuHenkouChuushiRiyuu ");
                sSql.Append("    ,KeiyakuBikou ");
                sSql.Append("    ,KeiyakuShosha ");
                sSql.Append("    ,KeiyakuTokkiShiyousho ");
                sSql.Append("    ,KeiyakuMitsumorisho ");
                sSql.Append("    ,KeiyakuTanpinChousaMitsumorisho ");
                sSql.Append("    ,KeiyakuSonota ");
                sSql.Append("    ,KeiyakuSonotaNaiyou ");
                sSql.Append("    ,KeiyakuSeikyuubi ");
                sSql.Append("    ,KeiyakuKeiyakusho ");
                sSql.Append("    ,KeiyakuZentokinUkewatashibi ");
                sSql.Append("    ,KeiyakuSeikyuusaki ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiS1 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiE1 ");
                sSql.Append("    ,KeiyakuSeikyuubi1 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiS2 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiE2 ");
                sSql.Append("    ,KeiyakuSeikyuubi2 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiS3 ");
                sSql.Append("    ,KeiyakuSeikyuuTaishouKoukiE3 ");
                sSql.Append("    ,KeiyakuSeikyuubi3 ");
                sSql.Append("    ,KeiyakuKankeibusho1 ");
                sSql.Append("    ,KeiyakuKankeibusho2 ");
                sSql.Append("    ,KeiyakuKankeibusho3 ");
                sSql.Append("    ,KeiyakuKankeibusho4 ");
                sSql.Append("    ,KeiyakuKankeibusho5 ");
                sSql.Append("    ,KeiyakuKankeibusho6 ");
                sSql.Append("    ,KeiyakuKankeibusho7 ");
                sSql.Append("    ,KeiyakuKankeibusho8 ");
                sSql.Append("    ,KeiyakuKankeibusho9 ");
                sSql.Append("    ,KeiyakuKankeibusho10 ");
                sSql.Append("    ,KeiyakuKankeibusho11 ");
                sSql.Append("    ,KeiyakuKankeibusho12 ");
                sSql.Append("    ,KeiyakuKankeibusho14 ");
                sSql.Append("    ,KeiyakuKankeibusho15 ");
                sSql.Append("    ,KeiyakuKankeibusho13 ");
                sSql.Append("    ,KeiyakuNyuukinYoteibi ");
                sSql.Append("    ,KeiyakuUriageHaibunCho1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunCho2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunJo1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunJo2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunJosys1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunJosys2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunKei1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunKei2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC1 ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC1Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuRIBC1 ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC2 ");
                sSql.Append("    ,KeiyakuUriageHaibunRIBC2Mei ");
                sSql.Append("    ,KeiyakuUriageHaibunGakuRIBC2 ");
                sSql.Append("    ,KeiyakuSeikyuubi4 ");
                sSql.Append("    ,KeiyakuSeikyuubi5 ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuTankeiMikomiCho ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuTankeiMikomiJo ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuTankeiMikomiJosys ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuTankeiMikomiKei ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuKurikoshiCho ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuKurikoshiJo ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuKurikoshiJosys ");
                sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" KeiyakuKurikoshiKei ");
                sSql.Append("    ,KeiyakuRIBCYouTankaDataMoushikomisho ");
                sSql.Append("    ,KeiyakuSashaKeiyu ");
                sSql.Append("    ,KeiyakuRIBCYouTankaData ");
                sSql.Append("    ,KeiyakuSaiitakuSonotaNaiyou ");
                sSql.Append("    ,KeiyakuSaiitakuKinshiNaiyou ");
                sSql.Append("    ,KeiyakuSaiitakuKinshiUmu ");
                sSql.Append("    ,KeiyakuAnkenMemoKeiyaku ");
                sSql.Append(" FROM KeiyakuJouhouEntory WHERE KeiyakuJouhouEntory.AnkenJouhouID = " + AnkenID);
            }

            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 業務情報新規登録、赤伝／黒伝
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenID"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝 、9:新規登録</param>
        /// <returns></returns>
        private int createGyoumuJouhou(SqlCommand cmd, string ankenID, int iType = 9)
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("INSERT INTO GyoumuJouhou ( ");
            sSql.Append("    AnkenJouhouID ");
            sSql.Append("    ,GyoumuJouhouID ");
            sSql.Append("    ,GyoumuCreateDate ");
            sSql.Append("    ,GyoumuCreateUser ");
            sSql.Append("    ,GyoumuCreateProgram ");
            sSql.Append("    ,GyoumuUpdateDate ");
            sSql.Append("    ,GyoumuUpdateUser ");
            sSql.Append("    ,GyoumuDeleteFlag ");
            sSql.Append("    ,KanriGijutsushaCD ");
            sSql.Append("    ,ShousaTantoushaCD ");
            sSql.Append("    ,SinsaTantoushaCD " );
            
            if(iType != 9)
            {
                sSql.Append("    ,GyoumuUpdateProgram ");
                sSql.Append("    ,GyoumuHyouten ");
                sSql.Append("    ,GyoumuKanriHyouten ");
                sSql.Append("    ,GyoumuTECRISTourokuBangou ");
                sSql.Append("    ,GyoumuKeisaiTankaTeikyou ");
                sSql.Append("    ,GyoumuChosakukenJouto ");
                sSql.Append("    ,GyoumuSeikyuubi ");
                sSql.Append("    ,GyoumuSeikyuusho ");
                sSql.Append("    ,GyoumuHikiwatashiNaiyou ");
                sSql.Append("    ,KanriGijutsushaNM ");
                sSql.Append("    ,ShousaTantoushaNM ");
                sSql.Append("    ,SinsaTantoushaNM ");
                sSql.Append("    ,GyoumuShousaHyouten ");
            }

            if (iType == 9)
            {
                sSql.Append(") VALUES ( ");
                sSql.Append("    " + ankenID);
                sSql.Append("    ,  " + ankenID + " ");
                sSql.Append("    ,  GETDATE() ");
                sSql.Append("    , N'" + UserInfos[0] + "'");
                sSql.Append("    , 'InsertEntry'");
                sSql.Append("    ,  GETDATE() ");
                sSql.Append("    , N'" + UserInfos[0] + "'");
                sSql.Append("    ,   0  ");
                sSql.Append("    ,   null  ");
                sSql.Append("    ,   null  ");
                sSql.Append("    ,   null  ");
                sSql.Append(")");
            }
            else
            {
                sSql.Append("     ) SELECT ");
                sSql.Append("    " + ankenID);
                sSql.Append("    ," + ankenID);
                sSql.Append("    ,GETDATE() ");
                sSql.Append("    ,N'" + UserInfos[0] + "' ");
                sSql.Append("    ,'ChangeKianEntry' ");
                sSql.Append("    ,GETDATE() ");
                sSql.Append("    ,N'" + UserInfos[0] + "' ");
                if(iType == 70 || iType == 71)
                {
                    sSql.Append("    ,1 ");
                }
                else
                {
                    sSql.Append("    ,0 ");
                }
                sSql.Append("    ,KanriGijutsushaCD ");
                sSql.Append("    ,ShousaTantoushaCD ");
                sSql.Append("    ,SinsaTantoushaCD ");
                sSql.Append("    ,'ChangeKianEntry' ");
                sSql.Append("    ,GyoumuHyouten ");
                sSql.Append("    ,GyoumuKanriHyouten ");
                sSql.Append("    ,GyoumuTECRISTourokuBangou ");
                sSql.Append("    ,GyoumuKeisaiTankaTeikyou ");
                sSql.Append("    ,GyoumuChosakukenJouto ");
                sSql.Append("    ,GyoumuSeikyuubi ");
                sSql.Append("    ,GyoumuSeikyuusho ");
                sSql.Append("    ,GyoumuHikiwatashiNaiyou ");
                sSql.Append("    ,KanriGijutsushaNM ");
                sSql.Append("    ,ShousaTantoushaNM ");
                sSql.Append("    ,SinsaTantoushaNM ");
                sSql.Append("    ,GyoumuShousaHyouten ");
                sSql.Append(" FROM GyoumuJouhou WHERE GyoumuJouhou.AnkenJouhouID = " + AnkenID);
            }
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 業務配分　配分率登録処理
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenID"></param>
        private void createGyoumuHaibun10(SqlCommand cmd, string ankenID, string HaibunID)
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("INSERT INTO GyoumuHaibun ( ");
            sSql.Append("    GyoumuHaibunID ");
            sSql.Append("    ,GyoumuAnkenJouhouID ");
            sSql.Append("    ,GyoumuHibunKubun ");
            sSql.Append("    ,GyoumuChosaBuRitsu ");
            sSql.Append("    ,GyoumuJigyoFukyuBuRitsu ");
            sSql.Append("    ,GyoumuJyohouSystemBuRitsu " );
            sSql.Append("    ,GyoumuSougouKenkyuJoRitsu ");
            sSql.Append("    ,GyoumuShizaiChousaRitsu ");
            sSql.Append("    ,GyoumuEizenRitsu ");
            sSql.Append("    ,GyoumuKikiruiChousaRitsu " );
            sSql.Append("    ,GyoumuKoujiChousahiRitsu ");
            sSql.Append("    ,GyoumuSanpaiFukusanbutsuRitsu ");
            sSql.Append("    ,GyoumuHokakeChousaRitsu ");
            sSql.Append("    ,GyoumuShokeihiChousaRitsu ");
            sSql.Append("    ,GyoumuGenkaBunsekiRitsu " );
            sSql.Append("    ,GyoumuKijunsakuseiRitsu ");
            sSql.Append("    ,GyoumuKoukyouRoumuhiRitsu ");
            sSql.Append("    ,GyoumuRoumuhiKoukyouigaiRitsu ");
            sSql.Append("    ,GyoumuSonotaChousabuRitsu ");
            sSql.Append("     ) VALUES ( ");
            sSql.Append(            HaibunID );
            sSql.Append("    ,  '" + ankenID + "'  ");
            sSql.Append("    ,   10  ");
            sSql.Append("    , N'" + getNumToDb(base_tbl07_1_numPercent1.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_1_numPercent2.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_1_numPercent3.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_1_numPercent4.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent1.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent2.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent3.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent4.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent5.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent6.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent7.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent8.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent9.Text) + "' ");
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent10.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent11.Text) + "' " );
            sSql.Append("    , N'" + getNumToDb(base_tbl07_2_numPercent12.Text) + "' ");
            sSql.Append("    )");
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 業務配分　配分額登録処理
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenID"></param>
        /// <param name="HaibunID"></param>
        private void createGyoumuHaibun30(SqlCommand cmd, string ankenID, string HaibunID)
        {
            cmd.CommandText = "INSERT INTO GyoumuHaibun ( " +
                                           "GyoumuHaibunID " +
                                           ",GyoumuAnkenJouhouID " +
                                           ",GyoumuHibunKubun " +

                                           ",GyoumuChosaBuRitsu " +
                                           ",GyoumuJigyoFukyuBuRitsu " +
                                           ",GyoumuJyohouSystemBuRitsu " +
                                           ",GyoumuSougouKenkyuJoRitsu " +

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
                                           ", 0 " +
                                           ", 0 " +
                                           ", 0 " +
                                           ", 0 " +
                                           ")";

            Console.WriteLine(cmd.CommandText);
            cmd.ExecuteNonQuery();

        }

        /// <summary>
        /// 入札情報登録処理
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenID"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝、9:新規登録</param>
        private int createNyuusatsuJouhou(SqlCommand cmd, string ankenID, int iType = 9)
        {
            // 業務発注区分
            // 入札方式
            // 最低制限価格有無
            // 入札(予定)日
            // 参考見積対応
            // 受注意欲
            // 再委託禁止条項の記載有無
            //参考見積額(税抜)
            //当会応札
            //再委託禁止条項の内容
            //その他の内容
            //入札状況
            //予定価格(税抜)
            //応札数
            //落札者状況
            //落札額状況
            //落札者
            //落札額(税抜)

            StringBuilder sSql = new StringBuilder();
            sSql.Append("INSERT INTO NyuusatsuJouhou ( ");
            sSql.Append("    AnkenJouhouID ");
            sSql.Append("    , NyuusatsuJouhouID ");
            sSql.Append("    , NyuusatsuCreateProgram ");
            sSql.Append("    , NyuusatsuCreateDate ");
            sSql.Append("    , NyuusatsuCreateUser ");
            sSql.Append("    , NyuusatsuUpdateProgram ");
            sSql.Append("    , NyuusatsuUpdateDate ");
            sSql.Append("    , NyuusatsuUpdateUser ");
            sSql.Append("    , NyuusatsuDeleteFlag ");

            sSql.Append("    , NyuusatsuHoushiki ");            // 入札方式
            sSql.Append("    , NyuusatsuRakusatuSougaku ");     // 落札総額
            sSql.Append("    , NyuusatsuRakusatsushaID ");        // 入札状況
            sSql.Append("    , NyuusatsuRakusatsusha ");          // 落札者（建設物価調査会等が入る）
            sSql.Append("    , NyuusatsuKyougouTashaID ");

            sSql.Append("    , NyuusatsuGyoumuBikou ");
            sSql.Append("    , NyuusatsuMitsumorigaku ");

            sSql.Append("    , NyuusatsuAnkenMemoNuusatsu ");   // --案件メモ(入札)
            sSql.Append("    , NyuusatsuSaiitakuSonotaNaiyou ");   // --その他の内容

            sSql.Append("    , NyuusatsuYoteiKakaku ");  // 予定価格（税抜）
            sSql.Append("    , NyuusatsushaSuu ");  // 応札数
            sSql.Append("    , NyuusatsuRakusatsuShaJokyou ");  // 落札者状況
            sSql.Append("    , NyuusatsuRakusatsuGakuJokyou ");  // 落札額状況

            sSql.Append("    , NyuusatsuSaiitakuKinshiNaiyou ");   // --再委託禁止条項の内容
            sSql.Append("    , NyuusatsuSaiitakuKinshiUmu ");   // --再委託禁止条項の記載有無
            sSql.Append("    , NyuusatsuJuchuuIyoku ");   // --受注意欲
            sSql.Append("    , NyuusatsuSankoumitsumoriKingaku ");   // --参考見積額
            sSql.Append("    , NyuusatsuSankoumitsumoriTaiou ");   // --参考見積対応
            sSql.Append("    , NyuusatsuSaiteiKakakuUmu ");   // --最低制限価格有無NyuusatsuRakusatugaku

            sSql.Append("    , NyuusatsuGyoumuHachuukubun ");   // --業務発注区分
            // No 1586　エントリくんの新規登録時10　入札情報・入札結果欄の落札額が入力出来なくなっている。
            sSql.Append("    , NyuusatsuJouhouTourokubi ");   // --落札額
            sSql.Append("    , NyuusatsuRakusatugaku ");

            if (iType != 9)
            {
                sSql.Append("    , NyuusatsuOusatugaku ");
                // No 1586　エントリくんの新規登録時10　入札情報・入札結果欄の落札額が入力出来なくなっている。
                //sSql.Append("    , NyuusatsuRakusatugaku ");
                sSql.Append("    , NyuusatsuNendoKurikoshigaku ");
                sSql.Append("    , NyuusatsuKyougouTasha ");
                sSql.Append("    , NyuusatsuKeiyakukeitaiCDSaishuu ");
                sSql.Append("    , NyuusatsuDenshiNyuusatsu ");
                sSql.Append("    , NyuusatsuTanpinMikomigaku ");
                sSql.Append("    , NyuusatsuShoruiSoufu ");
                sSql.Append("    , NyuusatsuRakusatsuKekkaDate ");
                sSql.Append("    , NyuusatsuRakusatsuShokaiDate ");
                sSql.Append("    , NyuusatsuRakusatsuSaisyuDate ");
                sSql.Append("    , NyuusatsuKekkaMemo ");
            }

            if (iType == 9)
            {
                sSql.Append("    ) VALUES ( ");
                sSql.Append("    " + ankenID);
                sSql.Append("    ,  " + ankenID + " ");
                sSql.Append("    , 'InsertEntry'");
                sSql.Append("    ,  GETDATE() ");
                sSql.Append("    , N'" + UserInfos[0] + "'");
                sSql.Append("    ,   null  ");
                sSql.Append("    ,  GETDATE() ");
                sSql.Append("    , N'" + UserInfos[0] + "'");
                sSql.Append("    ,  0 ");
                sSql.Append("    , N'" + base_tbl10_cmbNyusatuHosiki.SelectedValue + "'");
                sSql.Append("    , 0");
                if (this.IsNotSelected(base_tbl10_cmbNyusatuStats))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbNyusatuStats.SelectedValue.ToString());
                }
                sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl10_txtRakusatuSya.Text, 0, 0) + "'");
                //sSql.Append("    , ''");
                sSql.Append("    ,   null  ");

                sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl03_txtAnkenMemo.Text, 0, 0) + "'");
                sSql.Append("    , N'" + getNumToDb(base_tbl10_numSankoMitumoriAmt.Text) + "'");

                sSql.Append("    ,   null  ");
                sSql.Append("    , N'" + GlobalMethod.ChangeSqlText(base_tbl10_txtOtherNaiyo.Text, 0, 0) + "'");

                // 予定価格（税抜）
                sSql.Append("    , N'" + getNumToDb(base_tbl10_txtYoteiAmt.Text) + "'");
                // 応札数
                sSql.Append("    , N'" + base_tbl10_txtOsatuNum.Text + "'");
                // 落札者状況
                if (this.IsNotSelected(base_tbl10_cmbRakusatuStats))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbRakusatuStats.SelectedValue.ToString());
                }
                // 落札額状況
                if (this.IsNotSelected(base_tbl10_cmbRakusatuAmtStats))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbRakusatuAmtStats.SelectedValue.ToString());
                }
                // --再委託禁止条項の内容
                if (this.IsNotSelected(base_tbl10_cmbKinsiNaiyo))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbKinsiNaiyo.SelectedValue.ToString());
                }
                // --再委託禁止条項の記載有無
                if (this.IsNotSelected(base_tbl10_cmbKinsiUmu))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbKinsiUmu.SelectedValue.ToString());
                }
                // --受注意欲
                if (this.IsNotSelected(base_tbl10_cmbOrderIyoku))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbOrderIyoku.SelectedValue.ToString());
                }
                sSql.Append("    , N'" + getNumToDb(base_tbl10_numSankoMitumoriAmt.Text) + "'"); // --参考見積額
                // --参考見積対応
                if (this.IsNotSelected(base_tbl10_cmbSankoMitumori))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbSankoMitumori.SelectedValue.ToString());
                }
                // --最低制限価格有無
                if (this.IsNotSelected(base_tbl10_cmbLowestUmu))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbLowestUmu.SelectedValue.ToString());
                }
                // --業務発注区分
                if (this.IsNotSelected(base_tbl10_cmbOrderKubun))
                {
                    sSql.Append("    ,   null  ");
                }
                else
                {
                    sSql.Append("    , " + base_tbl10_cmbOrderKubun.SelectedValue.ToString());
                }
                sSql.Append("    ,   null  ");  // --入札結果登録日
                // 落札額 // No 1586　エントリくんの新規登録時10　入札情報・入札結果欄の落札額が入力出来なくなっている。
                sSql.Append("    , N'" + getNumToDb(base_tbl10_txtRakusatuAmt.Text) + "'");
                sSql.Append("    )");
            }
            else
            {
                // 赤伝／黒伝
                sSql.Append(" ) SELECT ");
                sSql.Append("    " + ankenID);
                sSql.Append("    , " + ankenID);
                sSql.Append("    , 'ChangeKianEntry' ");
                sSql.Append("    , GETDATE() ");
                sSql.Append("    , N'" + UserInfos[0] + "' ");
                sSql.Append("    , 'ChangeKianEntry' ");
                sSql.Append("    , GETDATE() ");
                sSql.Append("    , '" + UserInfos[0] + "' ");
                if (iType == 70 || iType == 71)
                {
                    sSql.Append("    , 1 ");
                }
                else
                {
                    sSql.Append("    , NyuusatsuDeleteFlag ");
                }
                sSql.Append("    , NyuusatsuHoushiki ");
                sSql.Append("    , ").Append(iType == 0 || iType == 70 ? "-" : "").Append(" NyuusatsuRakusatuSougaku ");
                sSql.Append("    , CASE WHEN NyuusatsuRakusatsushaID > 0 THEN NyuusatsuRakusatsushaID ELSE NULL END ");
                sSql.Append("    , CASE WHEN NyuusatsuRakusatsushaID > 0 THEN NyuusatsuRakusatsusha ELSE NULL END ");
                sSql.Append("    , CASE WHEN NyuusatsuKyougouTashaID > 0 THEN NyuusatsuKyougouTashaID ELSE NULL END ");
                sSql.Append("    , NyuusatsuGyoumuBikou ");
                sSql.Append("    , ").Append(iType == 0 || iType == 70 ? "-" : "").Append(" NyuusatsuMitsumorigaku ");
                sSql.Append("    , NyuusatsuAnkenMemoNuusatsu ");   // --案件メモ(入札)
                sSql.Append("    , NyuusatsuSaiitakuSonotaNaiyou ");   // --その他の内容

                sSql.Append("    , NyuusatsuYoteiKakaku ");  // 予定価格（税抜）
                sSql.Append("    , NyuusatsushaSuu ");  // 応札数
                sSql.Append("    , NyuusatsuRakusatsuShaJokyou ");  // 落札者状況
                sSql.Append("    , NyuusatsuRakusatsuGakuJokyou ");  // 落札額状況

                sSql.Append("    , NyuusatsuSaiitakuKinshiNaiyou ");   // --再委託禁止条項の内容
                sSql.Append("    , NyuusatsuSaiitakuKinshiUmu ");   // --再委託禁止条項の記載有無
                sSql.Append("    , NyuusatsuJuchuuIyoku ");   // --受注意欲
                sSql.Append("    , NyuusatsuSankoumitsumoriKingaku ");   // --参考見積額
                sSql.Append("    , NyuusatsuSankoumitsumoriTaiou ");   // --参考見積対応
                sSql.Append("    , NyuusatsuSaiteiKakakuUmu ");   // --最低制限価格有無
                sSql.Append("    , NyuusatsuGyoumuHachuukubun ");   // --業務発注区分
                sSql.Append("    , NyuusatsuJouhouTourokubi ");   // --入札結果登録日
                sSql.Append("    , ").Append(iType == 0 || iType == 70 ? "-" : "").Append(" NyuusatsuRakusatugaku "); // No 1586　エントリくんの新規登録時10　入札情報・入札結果欄の落札額が入力出来なくなっている。
                // ↑↑ここまで共通
                sSql.Append("    , ").Append(iType == 0 || iType == 70 ? "-" : "").Append(" NyuusatsuOusatugaku ");
                //sSql.Append("    , ").Append(iType == 0 || iType == 70 ? "-" : "").Append(" NyuusatsuRakusatugaku "); // No 1586　エントリくんの新規登録時10　入札情報・入札結果欄の落札額が入力出来なくなっている。
                sSql.Append("    , ").Append(iType == 0 || iType == 70 ? "-" : "").Append(" NyuusatsuNendoKurikoshigaku ");                
                sSql.Append("    , CASE WHEN NyuusatsuKyougouTashaID > 0 THEN NyuusatsuKyougouTasha ELSE NULL END ");                             
                sSql.Append("    , NyuusatsuKeiyakukeitaiCDSaishuu ");
                sSql.Append("    , NyuusatsuDenshiNyuusatsu ");
                sSql.Append("    , NyuusatsuTanpinMikomigaku ");
                sSql.Append("    , NyuusatsuShoruiSoufu ");
                sSql.Append("    , NyuusatsuRakusatsuKekkaDate ");
                sSql.Append("    , NyuusatsuRakusatsuShokaiDate " );
                sSql.Append("    , NyuusatsuRakusatsuSaisyuDate ");
                sSql.Append("    , NyuusatsuKekkaMemo ");
                sSql.Append(" FROM NyuusatsuJouhou WHERE NyuusatsuJouhou.NyuusatsuJouhouID = " + AnkenID);
            }
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
           return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 顧客契約情報登録処理
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenID"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝、、9:新規登録</param>
        /// <returns></returns>
        private int createKokyakuKeiyakuJouhou(SqlCommand cmd, string ankenID, int iType = 9)
        {
            StringBuilder sSql = new StringBuilder();

            // Insert　カラム
            sSql.Append("INSERT INTO KokyakuKeiyakuJouhou ( ");
            sSql.Append("    AnkenJouhouID ");
            sSql.Append("    ,KokyakuKeiyakuID ");
            sSql.Append("    ,KokyakuCreateUser ");
            sSql.Append("    ,KokyakuCreateDate ");
            sSql.Append("    ,KokyakuCreateProgram ");
            sSql.Append("    ,KokyakuUpdateUser ");
            sSql.Append("    ,KokyakuUpdateDate ");
            sSql.Append("    ,KokyakuDeleteFlag ");

            if (iType != 9)
            {
                sSql.Append("    ,KokyakuUpdateProgram");
                sSql.Append("    ,KokyakuKeiyakuTanka ");
                sSql.Append("    ,KokyakuKeiyakuChosakuken ");
                sSql.Append("    ,KokyakuKeiyakuKeisai ");
                sSql.Append("    ,KokyakuKeiyakuTokchouChosaku ");
                sSql.Append("    ,KokyakuKeiyakuRiyuu ");
                sSql.Append("    ,KokyakuMaebaraiJoukou ");
                sSql.Append("    ,KokyakuMaebaraiSeikyuu ");
                sSql.Append("    ,KokyakuSekkeiTanka ");
                sSql.Append("    ,KokyakuSekisanKijun ");
                sSql.Append("    ,KokyakuKaiteiGetsu ");
                sSql.Append("    ,KokyakuShichouson ");
                sSql.Append("    ,KokyakuGijutsuCenter ");
                sSql.Append("    ,KokyakuSonota ");
                sSql.Append("    ,KokyakuKeiyakuRiyuuTou ");
                sSql.Append("    ,KokyakuDataTeikyou ");
                sSql.Append("    ,KokyakuAlpha ");
                sSql.Append("    ,KokyakuDataDoboku ");
                sSql.Append("    ,KokyakuDataNourin ");
                sSql.Append("    ,KokyakuDataEizen ");
                sSql.Append("    ,KokyakuDataSonota ");
                sSql.Append("    ,KokyakuDataSekouP ");
                sSql.Append("    ,KokyakuDataDobokuKouji ");
                sSql.Append("    ,KokyakuDataRIBC ");
                sSql.Append("    ,KokyakuDataGoukei ");
                sSql.Append("    ,KokyakuDataKeisaiTanka ");
                sSql.Append("    ,KokyakuDataWebTeikyou ");
                sSql.Append("    ,KokyakuDataKeiyaku ");
                sSql.Append("    ,KokyakuDataTempFile ");
                sSql.Append("    ,KokyakuDataTempFileData ");
                sSql.Append("    ,KokyakuData05Comment ");
                sSql.Append("    ,KokyakuData06Comment ");
                sSql.Append("    ,KokyakuData07Comment ");
                sSql.Append("    ,KokyakuDataMeiki ");
                sSql.Append("    ,KokyakuDataTeikyouTensu ");
            }

            // Insert　Values
            if (iType == 9)
            {
                sSql.Append("    ) VALUES (");
                sSql.Append("    ").Append(ankenID);
                sSql.Append("    , " + ankenID);
                sSql.Append("    , N'" + UserInfos[0] + "'");
                sSql.Append("    , GETDATE()");
                sSql.Append("    , 'InsertEntry'");
                sSql.Append("    , N'" + UserInfos[0] + "'");
                sSql.Append("    , GETDATE()");
                sSql.Append("    , 0");
                sSql.Append("    )");
            }
            else
            {
                sSql.Append("     ) SELECT ");
                sSql.Append("    " + ankenID);
                sSql.Append("    ," + ankenID);
                sSql.Append("    ,N'" + UserInfos[0] + "' ");
                sSql.Append("    ,GETDATE() ");
                sSql.Append("    ,'ChangeKianEntry' ");
                sSql.Append("    ,N'" + UserInfos[0] + "' ");
                sSql.Append("    ,GETDATE() ");
                if (iType == 70 || iType == 71)
                {
                    sSql.Append("    ,1");
                }
                else
                {
                    sSql.Append("    ,0");
                }
                sSql.Append("    ,'ChangeKianEntry' ");
                sSql.Append("    ,KokyakuKeiyakuTanka ");
                sSql.Append("    ,KokyakuKeiyakuChosakuken ");
                sSql.Append("    ,KokyakuKeiyakuKeisai ");
                sSql.Append("    ,KokyakuKeiyakuTokchouChosaku ");
                sSql.Append("    ,KokyakuKeiyakuRiyuu ");
                sSql.Append("    ,KokyakuMaebaraiJoukou ");
                sSql.Append("    ,KokyakuMaebaraiSeikyuu ");
                sSql.Append("    ,KokyakuSekkeiTanka ");
                sSql.Append("    ,KokyakuSekisanKijun ");
                sSql.Append("    ,KokyakuKaiteiGetsu ");
                sSql.Append("    ,KokyakuShichouson ");
                sSql.Append("    ,KokyakuGijutsuCenter ");
                sSql.Append("    ,KokyakuSonota ");
                sSql.Append("    ,KokyakuKeiyakuRiyuuTou ");
                sSql.Append("    ,KokyakuDataTeikyou ");
                sSql.Append("    ,KokyakuAlpha ");
                sSql.Append("    ,KokyakuDataDoboku ");
                sSql.Append("    ,KokyakuDataNourin ");
                sSql.Append("    ,KokyakuDataEizen ");
                sSql.Append("    ,KokyakuDataSonota ");
                sSql.Append("    ,KokyakuDataSekouP ");
                sSql.Append("    ,KokyakuDataDobokuKouji ");
                sSql.Append("    ,KokyakuDataRIBC ");
                sSql.Append("    ,KokyakuDataGoukei ");
                sSql.Append("    ,KokyakuDataKeisaiTanka ");
                sSql.Append("    ,KokyakuDataWebTeikyou ");
                sSql.Append("    ,KokyakuDataKeiyaku ");
                sSql.Append("    ,KokyakuDataTempFile ");
                sSql.Append("    ,KokyakuDataTempFileData ");
                sSql.Append("    ,KokyakuData05Comment ");
                sSql.Append("    ,KokyakuData06Comment ");
                sSql.Append("    ,KokyakuData07Comment ");
                sSql.Append("    ,KokyakuDataMeiki ");
                sSql.Append("    ,KokyakuDataTeikyouTensu ");
                sSql.Append(" FROM KokyakuKeiyakuJouhou WHERE KokyakuKeiyakuJouhou.AnkenJouhouID = " + AnkenID);
            }
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            return cmd.ExecuteNonQuery();
        }

        #endregion

        #region 更新モード処理メソッド ------------
        /// <summary>
        /// 案件番号変更処理
        /// </summary>
        /// <param name="ori_ankenNo"></param>
        /// <returns></returns>
        private string changeAnkenNo(string ori_ankenNo)
        {
            string ankenNo = "";
            string jigyoubuHeadCD = getJigyoubuHeadCD(1);

            // 業務分類CD + 年度下2桁
            ankenNo = jigyoubuHeadCD + base_tbl03_cmbKokiStartYear.SelectedValue.ToString().Substring(2, 2);

            // KashoShibuCD
            DataTable dt = GlobalMethod.getData("KashoShibuCD", "KashoShibuCD", "Mst_Busho", "GyoumuBushoCD = '" + base_tbl02_cmbJyutakuKasyoSibu.SelectedValue.ToString() + "'");

            // KashoShibuCDが正しい
            if (dt != null && dt.Rows.Count > 0)
            {
                ankenNo = ankenNo + dt.Rows[0][0].ToString();
            }
            else
            {
                return "";
            }
            // No1559 1310　工期自を変更時に年度が変わり、案件番号が変更された際、案件番号の最終三桁が最大値で取られていない。
            dt = GlobalMethod.getData("SUBSTRING(AnkenAnkenBangou,7,3)", "TOP 1 SUBSTRING(AnkenAnkenBangou,7,3)", "AnkenJouhou",
                "AnkenAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + ankenNo + "%' and AnkenDeleteFlag != 1 ORDER BY AnkenAnkenBangou DESC");
            //dt = GlobalMethod.getData("SELECT TOP 1 SUBSTRING(AnkenAnkenBangou,7,3)", "SUBSTRING(AnkenAnkenBangou,7,3)", "AnkenJouhou",
            //    "AnkenAnkenBangou COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'" + ankenNo + "%' and AnkenDeleteFlag != 1 ORDER BY AnkenAnkenBangou DESC");

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
            base_tbl02_txtAnkenNo.Text = ankenNo;
            tblAKInfo_lblAnkenNo.Text = ankenNo;
            base_tbl02_txtAnkenChangerCD.Text = UserInfos[0];
            base_tbl02_txtAnkenChanger.Text = UserInfos[1];
            base_tbl02_txtAnkenChangDt.Text = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            base_tbl02_txtAnkenChangHistory.Text = ori_ankenNo;

            return ankenNo;            
        }

        /// <summary>
        /// 受託番号の設定　OR　解除
        /// </summary>
        /// <param name="jutakubangou"></param>
        /// <param name="ankenEda"></param>
        /// <returns></returns>
        private void setOrClearJutakuBan(string ankenNo)
        {
            string jutakubangou = base_tbl02_txtJyutakuNo.Text;
            string ankenEda = base_tbl02_txtJyutakuEdNo.Text;

            // 受託番号が空の場合
            if (base_tbl02_txtJyutakuNo.Text == "")
            {
                // item2_3_7：落札者（落札状況にチェックが無いと入らないので、落札状況はチェックなしとする）
                // ENTORY_TOUKAI:建設物価調査会
                if (bid_tbl03_1_txtRakusatuSya.Text == GlobalMethod.GetCommonValue2("ENTORY_TOUKAI"))
                {
                    // true:作成する false：作成しない
                    bool JutakuEdaCreateFlag = true;

                    // 受託番号の枝番を取得する
                    DataTable dt = GlobalMethod.getData("AnkenJutakuBangouEda", "TOP 1 AnkenJutakuBangouEda", "AnkenJouhou",
                        "AnkenJouhouID = " + AnkenID + " AND AnkenJutakuBangouEda <> '' and AnkenJutakuBangouEda != '' and AnkenDeleteFlag != 1 ");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        // 存在するので作成なし
                        JutakuEdaCreateFlag = false;
                        ankenEda = dt.Rows[0][0].ToString(); // 枝番を取得できた場合、変数に入れておく
                        dt.Clear();
                    }

                    if (JutakuEdaCreateFlag == true)
                    {
                        dt = GlobalMethod.getData("AnkenJutakuBangouEda", "TOP 1 AnkenJutakuBangouEda", "AnkenJouhou",
                            "AnkenAnkenBangou = (SELECT AnkenAnkenBangou FROM AnkenJouhou WHERE AnkenJouhouID = " + AnkenID + ") and AnkenJutakuBangouEda != '' and AnkenDeleteFlag != 1 order by AnkenJutakuBangouEda desc ");
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            // 枝番（-nn）を取得
                            ankenEda = dt.Rows[0][0].ToString();
                            // 「-」落とし
                            int i = int.Parse(ankenEda);
                            i += 1;
                            ankenEda = "" + string.Format("{0:D2}", i);
                        }
                        else
                        {
                            if(dt == null ) GlobalMethod.outputLogger("Execute_SQL", "枝番が取得できずにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                            // 枝番（-01）をセット
                            ankenEda = "01";
                        }
                    }
                    
                    // 受託番号の自動生成
                    jutakubangou = ankenNo + "-" + ankenEda;

                    // 引合タブの受託番号、枝番にセット
                    base_tbl02_txtJyutakuNo.Text = jutakubangou;
                    base_tbl02_txtJyutakuEdNo.Text = ankenEda;
                    tblAKInfo_lblJyutakuNo.Text = jutakubangou;
                }
            }
            else
            {
                //・受託番号を解除する為、受注チェックを外すことにより、受託番号を削除する。
                //・受託番号削除は、起案後には行えない。起案後に行う場合は、起案解除後に行える。（システム管理者のみ）
                if (UserInfos[4].Equals("2"))
                {
                    if (ca_tbl01_chkKian.Checked == false && (bid_tbl03_1_txtRakusatuSya.Text.Equals(GlobalMethod.GetCommonValue2("ENTORY_TOUKAI"))) == false)
                    {
                        string checkBangou = tblAKInfo_lblJyutakuNo.Text;
                        bool isDel = false;
                        //・窓口ミハルが登録されている場合は、エラーメッセージを表示し、削除が出来ない。
                        DataTable dt = GlobalMethod.getData("MadoguchiJutakuBangou", "TOP 1 MadoguchiJutakuBangou", "MadoguchiJouhou",
                            "AnkenJouhouID = " + AnkenID + " AND MadoguchiJutakuBangou = '" + ankenNo + "' AND MadoguchiJutakuBangouEdaban = '" + base_tbl02_txtJyutakuEdNo.Text + "' AND MadoguchiDeleteFlag != 1 ");

                        if (dt != null && dt.Rows.Count > 0)
                        {
                            set_error(GlobalMethod.GetMessage("E10726", ""));
                            isDel = false;
                            dt.Clear();
                        }
                        else
                        {
                            isDel = true;
                            if(dt == null) GlobalMethod.outputLogger("Execute_SQL", "窓口ミハルが登録されているかチェックにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                        }

                        //・単価契約が登録されている場合は、エラーメッセージを表示し、削除が出来ない。
                        if (isDel == true)
                        {
                            dt = GlobalMethod.getData("TankakeiyakuJutakuBangou", "TOP 1 TankakeiyakuJutakuBangou", "TankaKeiyaku",
                            "AnkenJouhouID = " + AnkenID + " AND TankakeiyakuJutakuBangou = '" + checkBangou + "' and TankakeiyakuDeleteFlag != 1 ");

                            if (dt != null && dt.Rows.Count > 0)
                            {
                                set_error(GlobalMethod.GetMessage("E10727", ""));
                                isDel = false;
                                dt.Clear();
                            }
                            else
                            {
                                isDel = true;
                                if (dt == null) GlobalMethod.outputLogger("Execute_SQL", "単価契約が登録されているかチェックにエラー", "ID:" + AnkenID + " mode:" + mode, "DEBUG");
                            }
                            

                        }
                        if (isDel == true)
                        {
                            jutakubangou = "";
                            ankenEda = "";
                            // 引合タブの受託番号、枝番にセット
                            base_tbl02_txtJyutakuNo.Text = jutakubangou;
                            base_tbl02_txtJyutakuEdNo.Text = ankenEda;
                            tblAKInfo_lblJyutakuNo.Text = jutakubangou;
                        }

                    }
                }
            }
        }

        /// <summary>
        /// フォルダリムーブ処理
        /// </summary>
        /// <param name="ori_ankenNo"></param>
        /// <returns></returns>
        private bool RenameFolder(string ori_ankenNo)
        {
            // えんとり君修正STEP2
            bool isMoveOk = false;
            string folderTo = GlobalMethod.ChangeSqlText(base_tbl02_txtRenameFolder.Text, 0, 0);
            string sJigyoubuHeadCD = getJigyoubuHeadCD();
            int isRename = 0; // 0:何もしない、1:リネーム、2:削除のみ、3：新規作成
            //if (sFolderRenameBef.Equals(folderTo) == false)
            if (string.IsNullOrEmpty(folderTo) == false && sFolderRenameBef.Equals(folderTo) == false)
            {
                // リネームボタン押下時
                if (string.IsNullOrEmpty(ca_tbl01_hidResetAnkenno.Text) == false)
                {
                    folderTo = folderTo.Replace(ca_tbl01_hidResetAnkenno.Text, base_tbl02_txtAnkenNo.Text);
                }
                if (sJigyoubuHeadCD_ori.Equals("T") && sJigyoubuHeadCD.Equals("T"))
                {
                    //No1668 変更前後が調査部で、変更前後のフォルダパスが設定済みでかつ異なっている場合、
                    //元の案件Noにかかわらず強制的にリネーム対象とする
                    isRename = 1;
                    ////もとファイル
                    //if (base_tbl02_txtAnkenFolder.Text.Contains(ori_ankenNo))
                    //{
                    //    // リネーム前後、すべて調査部の場合、リネームを実施する
                    //    isRename = 1;
                    //}
                    //else
                    //{
                    //    isRename = 3;
                    //}
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
            else if (ori_ankenNo.Equals(base_tbl02_txtAnkenNo.Text) == false)
            {
                // No1560 1311　【備忘】現行の仕様では工期自を変更時に年度を変更し、案件番号が変更された際、フォルダ変更も行われる。
                //    ※フォルダ変更ボタンで確認せずにホルダ変更が行われてしまう。
                isRename = 5;
            }
            else
            {
                isRename = 5;
            }

            // No1560 1311　【備忘】現行の仕様では工期自を変更時に年度を変更し、案件番号が変更された際、フォルダ変更も行われる。
            //    ※フォルダ変更ボタンで確認せずにホルダ変更が行われてしまう。
            ////案件番号も変更する場合
            //if (ori_ankenNo.Equals(base_tbl02_txtAnkenNo.Text) == false && (isRename == 1 || isRename == 3 || isRename == 4))
            //{
            //    folderTo = GlobalMethod.ChangeSqlText(folderTo.Replace("\\" + ori_ankenNo, "\\" + base_tbl02_txtAnkenNo.Text), 0, 0);
            //}

            // リネームを実行する
            if (isRename == 1 || isRename == 4)
            {
                bool isError = false;
                //調査部から調査部----------------------------------------------------
                //E10018	元フォルダが見つかりませんでした。確認して下さい。
                if (Directory.Exists(sFolderRenameBef) == false)
                {
                    isError = true;
                    set_error(GlobalMethod.GetMessage("E10018", "基本情報等一覧"));
                }

                //E10019 リネームするフォルダが既に存在します。確認して下さい。
                if (Directory.Exists(folderTo) == true)
                {
                    isError = true;
                    set_error(GlobalMethod.GetMessage("E10019", "基本情報等一覧"));
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
                    set_error(GlobalMethod.GetMessage("E10020", "基本情報等一覧"));
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
                        set_error(GlobalMethod.GetMessage("E70065", "基本情報等一覧"));
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
                        set_error(GlobalMethod.GetMessage("E70046", "基本情報等一覧"));
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
                                    set_error(GlobalMethod.GetMessage("E70046", "基本情報等一覧"));
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
                string FolderBase = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", base_tbl03_cmbKokiStartYear.SelectedValue.ToString());
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
                base_tbl02_txtAnkenFolder.Text = folderTo;
                te_txtSeikyusyo.Text = folderTo + @"\02契約関係図書";
                ca_tbl01_txtTosyo.Text = folderTo;
                base_tbl02_txtRenameFolder.Text = "";
                ca_tbl01_hidResetAnkenno.Text = "";

                // 現在フォルダを更新
                sFolderRenameBef = base_tbl02_txtAnkenFolder.Text;
            }
            else
            {
                base_tbl02_txtRenameFolder.Text = "";
                ca_tbl01_hidResetAnkenno.Text = "";
            }
            return isMoveOk;
        }

        private string getUpdAnkenVals()
        {
            StringBuilder sSql = new StringBuilder();

            // １．進捗段階
            sSql.Append("    AnkenJizenDashinCheck = " + (base_tbl01_chkJizendasin.Checked ? 1 : 0));
            sSql.Append("    ,AnkenJizenDashinDate = " + Get_DateTimePicker("base_tbl01_dtpDtPrior"));
            sSql.Append("    ,AnkenNyuusatuCheck = " + (base_tbl01_chkNyusatu.Checked ? 1 : 0));
            sSql.Append("    ,AnkenNyuusatuDate = " + Get_DateTimePicker("base_tbl01_dtpDtBid"));
            sSql.Append("    ,AnkenKeiyakuCheck = " + (base_tbl01_chkKeiyaku.Checked ? 1 : 0));
            sSql.Append("    ,AnkenKeiyakuDate = " + Get_DateTimePicker("base_tbl01_dtpDtCa"));

            // ２．基本情報
            sSql.Append("    ,AnkenAnkenBangou = " + "N'" + base_tbl02_txtAnkenNo.Text + "'");
            sSql.Append("    ,AnkenJutakubangou = " + "N'" + base_tbl02_txtJyutakuNo.Text + "'");
            sSql.Append("    ,AnkenJutakuBangouEda = " + "N'" + base_tbl02_txtJyutakuEdNo.Text + "'");
            sSql.Append("    ,AnkenSakuseiKubun = N'" + ca_tbl01_cmbAnkenKubun.SelectedValue + "'");
            sSql.Append("    ,AnkenKeikakuBangou = N'" + base_tbl02_txtKeikakuNo.Text + "'");
            sSql.Append("    ,AnkenJutakushibu = " + "N'" + base_tbl02_cmbJyutakuKasyoSibu.Text + "'");
            sSql.Append("    ,AnkenJutakubushoCD = " + "N'" + base_tbl02_cmbJyutakuKasyoSibu.SelectedValue + "'");
            sSql.Append("    ,AnkenTantoushaCD = N'" + base_tbl02_txtKeiyakuTantoCD.Text + "'");
            sSql.Append("    ,AnkenTantoushaMei = N'" + base_tbl02_txtKeiyakuTanto.Text + "'");
            sSql.Append("    ,AnkenKeiyakusho = N'" + GlobalMethod.ChangeSqlText(base_tbl02_txtAnkenFolder.Text, 0, 0) + "'");

            sSql.Append("    ,AnkenFolderHenkouDatetime = " + (string.IsNullOrEmpty(base_tbl02_txtAnkenChangDt.Text) ? "NULL " : "'" + base_tbl02_txtAnkenChangDt.Text + "' "));
            sSql.Append("    ,AnkenFolderHenkouTantoushaCD = '" + base_tbl02_txtAnkenChangerCD.Text + "'");
            sSql.Append("    ,AnkenHenkoumaeAnkenBangou = " + (string.IsNullOrEmpty(base_tbl02_txtAnkenChangHistory.Text) ? "NULL " : "'" + base_tbl02_txtAnkenChangHistory.Text + "' "));
            //sSql.Append("    ,AnkenTourokubi = " + Get_DateTimePicker("base_tbl01_dtpDtPrior"));

            // ３．案件情報
            sSql.Append("    ,AnkenGyoumuMei = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl03_txtGyomuName.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenGyoumuKubun = N'" + base_tbl03_cmbKeiyakuKubun.SelectedValue + "'");
            sSql.Append("    ,AnkenGyoumuKubunCD = N'" + Get_GyoumuKubunCD(base_tbl03_cmbKeiyakuKubun.SelectedValue.ToString()) + "'");
            sSql.Append("    ,AnkenGyoumuKubunMei = " + "N'" + base_tbl03_cmbKeiyakuKubun.Text + "'");
            // 売上年度と契約区分は基本情報等一覧から連動できるので、そのままでOK？？？？？？★★★
            sSql.Append("    ,AnkenUriageNendo = " + ca_tbl01_cmbSalesYear.SelectedValue);
            sSql.Append("    ,AnkenKoukiNendo = " + "'" + base_tbl03_cmbKokiStartYear.SelectedValue.ToString() + "' ");
            sSql.Append("    ,AnkenAnkenMemoKihon = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl03_txtAnkenMemo.Text, 0, 0) + "'");// --案件メモ(基本情報)
            //sSql.Append("    AnkenHikiaijhokyo = N'" + item1_1.SelectedValue + "'");

            // ４．発注者情報
            sSql.Append("    ,AnkenHachushaCD = " + "N'" + base_tbl04_txtOrderCd.Text + "'");
            sSql.Append("    ,AnkenHachuushaMei = " + "N'" + base_tbl04_txtOrderName.Text + "'");
            sSql.Append("    ,AnkenHachushaKaMei = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl04_txtOrderKamei.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaKaMei = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl04_txtOrderName.Text, 0, 0) + "　" + GlobalMethod.ChangeSqlText(base_tbl04_txtOrderKamei.Text, 0, 0) + "'");

            // ５．発注担当者情報（調査窓口）
            sSql.Append("    ,AnkenHachuushaIraibusho = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtBusho.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaTantousha = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtTanto.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaTEL = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtTel.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaFAX = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtFax.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaMail = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtEmail.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaIraiYuubin = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtZip.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaIraiJuusho = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl05_txtAddress.Text, 0, 0) + "'");
            // ６．発注担当者情報（契約窓口）
            sSql.Append("    ,AnkenHachuushaKeiyakuBusho = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtBusho.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaKeiyakuTantou = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtTanto.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaKeiyakuTEL = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtTel.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaKeiyakuFAX = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtFax.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaKeiyakuMail = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtEmail.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaKeiyakuYuubin = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtZip.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuushaKeiyakuJuusho = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtAddress.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuuDaihyouYakushoku = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtOrderYakusyoku.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenHachuuDaihyousha = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl06_txtOrderSimei.Text, 0, 0) + "'");
            // ７．配分情報・業務内容
            // No.1533 削除
            //object obj = base_tbl07_3_cmbOen.SelectedValue;// --応援依頼の有無
            //if (obj == null || string.IsNullOrEmpty(obj.ToString()))
            //{
            //    sSql.Append("    ,AnkenOueniraiUmu = null");
            //}
            //else
            //{
            //    sSql.Append("    ,AnkenOueniraiUmu = " + "N'" + obj.ToString() + "'");
            //}
            sSql.Append("    ,AnkenOuenIraiMemo = N'" + GlobalMethod.ChangeSqlText(base_tbl07_3_txtOenMemo.Text, 0, 0) + "'");  // --応援依頼メモ

            // ８．過去案件情報
            // ９．事前打診・参考見積
            sSql.Append("    ,AnkenNyuusatsuHoushiki = " + "N'" + base_tbl10_cmbNyusatuHosiki.SelectedValue + "'");
            sSql.Append("    ,AnkenNyuusatsuYoteibi = " + Get_DateTimePicker("base_tbl10_dtpNyusatuDt"));
            sSql.Append("    ,AnkenJizenDashinIraibi = " + Get_DateTimePicker("base_tbl09_dtpJizenDasinIraiDt"));   // --事前打診依頼日
            sSql.Append("    ,AnkenHachuuYoteiMikomibi = " + Get_DateTimePicker("base_tbl09_dtpOrderYoteiDt"));   // --発注予定・見込日
            object obj = base_tbl09_cmbNotOrderStats.SelectedValue;// --未発注状況
            if (obj == null || string.IsNullOrEmpty(obj.ToString()))
            {
                sSql.Append("    ,AnkenMihachuuJoukyou = null");
            }
            else
            {
                sSql.Append("    ,AnkenMihachuuJoukyou = " + "N'" + obj.ToString() + "'");
            }
            obj = base_tbl09_cmbNotOrderReason.SelectedValue;// --「発注なし」の理由
            if (obj == null || string.IsNullOrEmpty(obj.ToString()))
            {
                sSql.Append("    ,AnkenHachuunashiRiyuu = null");
            }
            else
            {
                sSql.Append("    ,AnkenHachuunashiRiyuu = " + "N'" + obj.ToString() + "'");
            }
            sSql.Append("    ,AnkenSonotaNaiyou = N'" + GlobalMethod.ChangeSqlText(base_tbl09_txtOthenComment.Text, 0, 0) + "'");   // --「その他」の内容

            sSql.Append("    ,AnkenAnkenMemoMihachuu = N'" + GlobalMethod.ChangeSqlText(prior_tbl02_txtAnkenMemo.Text, 0, 0) + "'");   // --案件メモ(未発注)
            sSql.Append("    ,AnkenAnkenMemoJizendashin = N'" + GlobalMethod.ChangeSqlText(prior_tbl01_txtAnkenMemo.Text, 0, 0) + "'");  // --案件メモ（事前打診）

            sSql.Append("    ,AnkenMihachuuTourokubi = " + Get_DateTimePicker("prior_tbl02_dtpNotOrderDt"));  // --未発注の登録日
            
            sSql.Append("    ,AnkenToukaiSankouMitsumori = " + "N'" + base_tbl09_cmbSankomitumori.SelectedValue + "'");
            sSql.Append("    ,AnkenToukaiJyutyuIyoku = " + "N'" + base_tbl09_cmbOrderIyoku.SelectedValue + "'");
            sSql.Append("    ,AnkenToukaiSankouMitsumoriGaku = " + getNumToDb(base_tbl09_numSankomitumoriAmt.Text));

            // １０．入札情報・入札結果
            sSql.Append("    ,AnkenToukaiOusatu = " + "N'" + base_tbl10_cmbTokaiOsatu.SelectedValue + "'");

            // １１．契約情報
            sSql.Append("    ,AnkenKianzumi = " + (base_tbl11_1_chkKianzumi.Checked ? 1 : 0));
            sSql.Append("    ,GyoumuKanrishaCD = " + "N'" + ca_tbl05_txtGyomuCD.Text + "'");
            sSql.Append("    ,GyoumuKanrishaMei = " + "N'" + ca_tbl05_txtGyomu.Text + "'");
            sSql.Append("    ,AnkenKokyakuHyoukaComment = " + "N'" + GlobalMethod.ChangeSqlText(te_txtCustomComment.Text, 0, 0) + "'");
            sSql.Append("    ,AnkenToukaiHyoukaComment = " + "N'" + GlobalMethod.ChangeSqlText(te_txtTokaiComment.Text, 0, 0) + "'");

            // 更新情報
            sSql.Append("    ,AnkenUpdateProgram = " + "'UpdateEntry'");
            sSql.Append("    ,AnkenUpdateDate = " + "GETDATE()");
            sSql.Append("    ,AnkenUpdateUser = " + "N'" + UserInfos[0] + "'");

            return sSql.ToString();

        }

        /// <summary>
        /// 契約情報エントリ　更新                                    
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="sAnkenId">案件ID</param>
        /// <param name="flg">1:更新、4:伝票変更</param>
        private void updateKeiyakuJouhouEntory(SqlCommand cmd, string sAnkenId, int flg = 1)
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("UPDATE KeiyakuJouhouEntory SET ");
            sSql.Append("    KeiyakuKeiyakuTeiketsubi = " + Get_DateTimePicker("ca_tbl01_dtpChangeDt"));
            sSql.Append("    ,KeiyakuSakuseibi = " + Get_DateTimePicker("ca_tbl01_dtpKianDt"));
            sSql.Append("    ,KeiyakuKeiyakuKoukiKaishibi = " + Get_DateTimePicker("ca_tbl01_dtpKokiFrom"));
            sSql.Append("    ,KeiyakuKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("ca_tbl01_dtpKokiTo"));
            sSql.Append("    ,KeiyakuKeiyakuKingaku = " + getNumToDb(ca_tbl01_txtZeinukiAmt.Text));
            sSql.Append("    ,KeiyakuZeikomiKingaku = " + getNumToDb(ca_tbl01_txtZeikomiAmt.Text));
            sSql.Append("    ,KeiyakuuchizeiKingaku = " + getNumToDb(ca_tbl01_txtSyohizeiAmt.Text));
            sSql.Append("    ,KeiyakuShouhizeiritsu = N'" + ca_tbl01_txtTax.Text + "'");
            sSql.Append("    ,KeiyakuHenkouChuushiRiyuu = N'" + GlobalMethod.ChangeSqlText(ca_tbl01_txtRiyu.Text, 0, 0) + "'");
            sSql.Append("    ,KeiyakuAnkenMemoKeiyaku = N'" + GlobalMethod.ChangeSqlText(ca_tbl01_txtAnkenMemo.Text, 0, 0) + "'");// 案件メモ
            sSql.Append("    ,KeiyakuBikou = " + "N'" + GlobalMethod.ChangeSqlText(ca_tbl01_txtBiko.Text, 0, 0) + "'");
            sSql.Append("    ,KeiyakuShosha = " + (ca_tbl01_chkCaSyosya.Checked ? 1 : 0));
            sSql.Append("    ,KeiyakuTokkiShiyousho = " + (ca_tbl01_chkSiyosyo.Checked ? 1 : 0));
            sSql.Append("    ,KeiyakuMitsumorisho = " + (ca_tbl01_chkMitumorisyo.Checked ? 1 : 0));
            sSql.Append("    ,KeiyakuTanpinChousaMitsumorisho = " + (ca_tbl01_chkTanpinTyosa.Checked ? 1 : 0));
            sSql.Append("    ,KeiyakuSonota = " + (ca_tbl01_chkOther.Checked ? 1 : 0));
            sSql.Append("    ,KeiyakuSonotaNaiyou = N'"+ GlobalMethod.ChangeSqlText(ca_tbl01_txtOtherBiko.Text, 0, 0) + "'");
            sSql.Append("    ,KeiyakuZentokinUkewatashibi = " + Get_DateTimePicker("ca_tbl07_dtpRequst6"));
            sSql.Append("    ,KeiyakuZentokin = " + getNumToDb(ca_tbl07_txtRequst6.Text));
            sSql.Append("    ,Keiyakukeiyakukingakukei = " + getNumToDb(ca_tbl01_txtJyutakuAmt.Text));
            sSql.Append("    ,KeiyakuBetsuKeiyakuKingaku = " + getNumToDb(ca_tbl01_txtJyutakuGaiAmt.Text));
            sSql.Append("    ,KeiyakuSeikyuubi1 = " + Get_DateTimePicker("ca_tbl07_dtpRequst1"));
            sSql.Append("    ,KeiyakuSeikyuuKingaku1 = " + getNumToDb(ca_tbl07_txtRequst1.Text));
            sSql.Append("    ,KeiyakuSeikyuubi2 = " + Get_DateTimePicker("ca_tbl07_dtpRequst2"));
            sSql.Append("    ,KeiyakuSeikyuuKingaku2 = " + getNumToDb(ca_tbl07_txtRequst2.Text));
            sSql.Append("    ,KeiyakuSeikyuubi3 = " +  Get_DateTimePicker("ca_tbl07_dtpRequst3"));
            sSql.Append("    ,KeiyakuSeikyuuKingaku3 = " + getNumToDb(ca_tbl07_txtRequst3.Text));
            sSql.Append("    ,KeiyakuSeikyuubi4 = " +  Get_DateTimePicker("ca_tbl07_dtpRequst4"));
            sSql.Append("    ,KeiyakuSeikyuuKingaku4 = " + getNumToDb(ca_tbl07_txtRequst4.Text));
            sSql.Append("    ,KeiyakuSeikyuubi5 = " +  Get_DateTimePicker("ca_tbl07_dtpRequst5"));
            sSql.Append("    ,KeiyakuSeikyuuKingaku5 = " + getNumToDb(ca_tbl07_txtRequst5.Text));
            sSql.Append("    ,KeiyakuSakuseiKubunID = N'" + ca_tbl01_cmbAnkenKubun.SelectedValue + "'");
            sSql.Append("    ,KeiyakuSakuseiKubun = N'" + ca_tbl01_cmbAnkenKubun.Text + "'");
            if (flg != 4)
            {
                sSql.Append("    ,KeiyakuGyoumuKubun = N'" + base_tbl03_cmbKeiyakuKubun.SelectedValue + "'");
                sSql.Append("    ,KeiyakuGyoumuMei = N'" + base_tbl03_cmbKeiyakuKubun.Text + "'");
                sSql.Append("    ,KeiyakuJutakubangou = N'" + base_tbl02_txtJyutakuNo.Text + "'");
                sSql.Append("    ,KeiyakuEdaban = N'" + base_tbl02_txtJyutakuEdNo.Text + "'");
            }
            else
            {
                sSql.Append("    ,KeiyakuGyoumuKubun = N'" + ca_tbl01_cmbCaKubun.SelectedValue + "'");
                sSql.Append("    ,KeiyakuGyoumuMei = N'" + ca_tbl01_cmbCaKubun.Text + "'");
            }
            sSql.Append("    ,KeiyakuKianzumi = " + (ca_tbl01_chkKian.Checked ? 1 : 0));
            sSql.Append("    ,KeiyakuHachuushaMei = N'" + ca_tbl01_txtOrderKamei.Text + "'");

            sSql.Append("    ,KeiyakuHaibunChoZeinuki = " + getNumToDb(ca_tbl02_AftCaBm_numAmt1.Text));
            sSql.Append("    ,KeiyakuHaibunJoZeinuki = " + getNumToDb(ca_tbl02_AftCaBm_numAmt2.Text));
            sSql.Append("    ,KeiyakuHaibunJosysZeinuki = " + getNumToDb(ca_tbl02_AftCaBm_numAmt3.Text));
            sSql.Append("    ,KeiyakuHaibunKeiZeinuki = " + getNumToDb(ca_tbl02_AftCaBm_numAmt4.Text));
            sSql.Append("    ,KeiyakuHaibunZeinukiKei = " + getNumToDb(ca_tbl02_AftCaBm_numAmtAll.Text));

            sSql.Append("    ,KeiyakuUriageHaibunCho  = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt1.Text));
            sSql.Append("    ,KeiyakuUriageHaibunJo   = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt2.Text));
            sSql.Append("    ,KeiyakuUriageHaibunJosys  = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt3.Text));
            sSql.Append("    ,KeiyakuUriageHaibunKei  = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt4.Text));
            sSql.Append("    ,KeiyakuUriageHaibunGoukei = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmtAll.Text));


            sSql.Append("    ,KeiyakuTankeiMikomiCho  = " + getNumToDb(ca_tbl03_numAmt1.Text));
            sSql.Append("    ,KeiyakuTankeiMikomiJo  = " + getNumToDb(ca_tbl03_numAmt2.Text));
            sSql.Append("    ,KeiyakuTankeiMikomiJosys  = " + getNumToDb(ca_tbl03_numAmt3.Text));
            sSql.Append("    ,KeiyakuTankeiMikomiKei  = " + getNumToDb(ca_tbl03_numAmt4.Text));
            sSql.Append("    ,KeiyakuKurikoshiCho  = " + getNumToDb(ca_tbl04_numKurikosiAmt1.Text));
            sSql.Append("    ,KeiyakuKurikoshiJo  = " + getNumToDb(ca_tbl04_numKurikosiAmt2.Text));
            sSql.Append("    ,KeiyakuKurikoshiJosys  = " + getNumToDb(ca_tbl04_numKurikosiAmt3.Text));
            sSql.Append("    ,KeiyakuKurikoshiKei  = " + getNumToDb(ca_tbl04_numKurikosiAmt4.Text));
            sSql.Append("    ,KeiyakuRIBCYouTankaDataMoushikomisho = " + (ca_tbl01_chkRibcSyo.Checked ? 1 : 0));
            sSql.Append("    ,KeiyakuSashaKeiyu = " + (ca_tbl01_chkSasya.Checked ? 1 : 0));
            sSql.Append("    ,KeiyakuRIBCYouTankaData = " + (ca_tbl01_chkRibcAri.Checked ? 1 : 0));

            sSql.Append("    ,KeiyakuSaiitakuSonotaNaiyou = N'" + GlobalMethod.ChangeSqlText(ca_tbl01_txtOtherNaiyo.Text, 0, 0) + "'");
            sSql.Append("    ,KeiyakuSaiitakuKinshiNaiyou = N'" + ca_tbl01_cmbKinsiNaiyo.SelectedValue + "'");
            sSql.Append("    ,KeiyakuSaiitakuKinshiUmu = N'" + ca_tbl01_cmbKinsiUmu.SelectedValue + "'");

            sSql.Append("    ,KeiyakuUpdateProgram = " + "'UpdateEntry'");
            sSql.Append("    ,KeiyakuUpdateDate = " + "GETDATE()");
            sSql.Append("    ,KeiyakuUpdateUser = N'" + UserInfos[0] + "'");
            sSql.Append(" WHERE AnkenJouhouID = " + sAnkenId);

            cmd.CommandText = sSql.ToString();

            Console.WriteLine(cmd.CommandText);
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 案件情報更新
        /// </summary>
        /// <param name="cmd"></param>
        private int updateAnkenJouhou(SqlCommand cmd, int flag)
        {
            StringBuilder sSql = new StringBuilder();

            sSql.Append("UPDATE AnkenJouhou SET ");
            if (flag == 4)
            {
                sSql.Append("    AnkenUpdateDate = GETDATE() ");
                sSql.Append("    ,AnkenUpdateUser = '" + UserInfos[0] + "' ");
                sSql.Append("    ,AnkenSaishinFlg = 0 ");
                sSql.Append("    ,AnkenUpdateProgram = 'ChangeKianEntry' ");
            }
            else
            {
                // １．進捗段階
                sSql.Append("    AnkenKeiyakuKoukiKaishibi = " + Get_DateTimePicker("base_tbl03_dtpKokiFrom"));
                sSql.Append("    ,AnkenKeiyakuKoukiKanryoubi = " + Get_DateTimePicker("base_tbl03_dtpKokiTo"));
                sSql.Append("    ,AnkenKeiyakuTeiketsubi = " + Get_DateTimePicker("ca_tbl01_dtpChangeDt"));
                sSql.Append("    ,AnkenKeiyakuZeikomiKingaku = " + getNumToDb(ca_tbl01_txtZeikomiAmt.Text));
                sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuC = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt1.Text));
                sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuJ = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt2.Text));
                sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuJs = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt3.Text));
                sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuK = " + getNumToDb(ca_tbl02_AftCaBmZeikomi_numAmt4.Text));
                sSql.Append("    ,AnkenKeiyakuSakuseibi = " + Get_DateTimePicker("ca_tbl01_dtpKianDt"));
                if (flag == 3)
                {
                    sSql.Append("    ,AnkenKianZumi = 1");
                }

                if(flag == 1)
                {
                    // No.1555 1306　アラートメールの発報で、工期自が変更された場合 工期自が変更された場合、アラートメールのフラグを0に変更する。
                    string sDt = AnkenData_H.Rows[0]["KeiyakuKoukiKaishibi"].ToString();
                    if (base_tbl03_dtpKokiFrom.Text.Equals(sDt) == false)
                    {
                        sSql.Append("    ,AnkenArartError = 0");
                        sSql.Append("    ,AnkenArartwarning3 = 0");
                        sSql.Append("    ,AnkenArartwarning7 = 0");
                    }
                }
            }
            sSql.Append(" WHERE AnkenJouhouID = " + AnkenID);
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 売上計上情報リスト 更新
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="c1FlexGrid4"></param>
        private void updateRibcJouhou(SqlCommand cmd, C1FlexGrid c1FlexGrid4)
        {
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
                // 計上日、計上月、計上額のどれかが入っていれば登録する
                // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                if ((c1FlexGrid4.Rows[i][1] != null && c1FlexGrid4.Rows[i][1].ToString() != "")
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
                        cmd.CommandText = cmd.CommandText + ",N'" + getNumToDb(c1FlexGrid4.Rows[i][3].ToString()) + "'";
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
                // 計上日、計上月、計上額のどれかが入っていれば登録する
                // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                if ((c1FlexGrid4.Rows[i][9] != null && c1FlexGrid4.Rows[i][9].ToString() != "")
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
                        cmd.CommandText = cmd.CommandText + ",N'" + getNumToDb(c1FlexGrid4.Rows[i][11].ToString()) + "'";
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

                // 計上日、計上月、計上額のどれかが入っていれば登録する
                // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                if ((c1FlexGrid4.Rows[i][17] != null && c1FlexGrid4.Rows[i][17].ToString() != "")
                    || (c1FlexGrid4.Rows[i][19] != null && c1FlexGrid4.Rows[i][19].ToString() != "0"))
                {
                    // 新では計上額のみでも登録を可とする
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
                        cmd.CommandText = cmd.CommandText + ",N'" + getNumToDb(c1FlexGrid4.Rows[i][19].ToString()) + "'";
                    }
                    else
                    {
                        cmd.CommandText = cmd.CommandText + ",'0' " + "";
                    }
                    // 年度により、情報システム部の部コードを変更する
                    if (GetInt(ca_tbl01_cmbSalesYear.SelectedValue.ToString()) >= 2021)
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

                // 計上日、計上月、計上額のどれかが入っていれば登録する
                // c1FlexGrid の基本はNull、DBからの場合は空文字があり得る、\0は0、0を消すとまたnullになる
                if ((c1FlexGrid4.Rows[i][25] != null && c1FlexGrid4.Rows[i][25].ToString() != "")
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
                        cmd.CommandText = cmd.CommandText + ",N'" + getNumToDb(c1FlexGrid4.Rows[i][27].ToString()) + "'";
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
        }

        /// <summary>
        /// 業務情報更新
        /// </summary>
        /// <param name="cmd"></param>
        private void updateGyoumuJouhou(SqlCommand cmd)
        {
            StringBuilder sSql = new StringBuilder();
            //業務情報
            sSql.Append("UPDATE GyoumuJouhou SET ");
            sSql.Append("    GyoumuHyouten = N'" + te_txtPoint.Text + "'");
            sSql.Append("    ,KanriGijutsushaCD = N'" + ca_tbl05_txtKanriCD.Text + "'");
            sSql.Append("    ,KanriGijutsushaNM = N'" + ca_tbl05_txtKanri.Text + "'");
            sSql.Append("    ,GyoumuKanriHyouten = N'" + te_txtKanriPoint.Text + "'");
            sSql.Append("    ,ShousaTantoushaCD = N'" + ca_tbl05_txtSyosaCD.Text + "'");
            sSql.Append("    ,ShousaTantoushaNM = N'" + ca_tbl05_txtSyosa.Text + "'");
            sSql.Append("    ,GyoumuShousaHyouten = N'" + te_txtSyosaPoint.Text + "'");
            sSql.Append("    ,SinsaTantoushaCD = " + "N'" + ca_tbl05_txtSinsaCD.Text + "'");
            sSql.Append("    ,SinsaTantoushaNM = " + "N'" + ca_tbl05_txtSinsa.Text + "'");
            sSql.Append("    ,GyoumuTECRISTourokuBangou = " + "N'" + te_txtTecris.Text + "'");
            sSql.Append("    ,GyoumuKeisaiTankaTeikyou = " + "''");
            sSql.Append("    ,GyoumuChosakukenJouto = " + "''");
            sSql.Append("    ,GyoumuSeikyuubi = " + Get_DateTimePicker("te_dtpSeikyusyaDt"));
            sSql.Append("    ,GyoumuSeikyuusho = " + "N'" + GlobalMethod.ChangeSqlText(te_txtSeikyusyo.Text, 0, 0) + "'");
            sSql.Append("    ,GyoumuHikiwatashiNaiyou = " + "''");
            sSql.Append("    ,GyoumuUpdateDate = " + " GETDATE() ");
            sSql.Append("    ,GyoumuUpdateUser = " + "N'" + UserInfos[0] + "' ");
            sSql.Append("    ,GyoumuUpdateProgram = " + "'UpdateEntory' ");
            sSql.Append("    ,GyoumuDeleteFlag = " + "0 ");
            sSql.Append(" WHERE AnkenJouhouID = " + AnkenID);

            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 業務配分　配分率更新　
        /// </summary>
        /// <param name="cmd"></param>
        private void updateGyoumuHaibun10(SqlCommand cmd, string sAnkenId)
        {
            StringBuilder sSql = new StringBuilder();
            //業務配分
            sSql.Append("UPDATE GyoumuHaibun SET ");
            sSql.Append("     GyoumuChosaBuRitsu = " + getNumToDb(base_tbl07_1_numPercent1.Text));
            sSql.Append("    ,GyoumuJigyoFukyuBuRitsu = " + getNumToDb(base_tbl07_1_numPercent2.Text));
            sSql.Append("    ,GyoumuJyohouSystemBuRitsu = " + getNumToDb(base_tbl07_1_numPercent3.Text));
            sSql.Append("    ,GyoumuSougouKenkyuJoRitsu = " + getNumToDb(base_tbl07_1_numPercent4.Text));

            sSql.Append("    ,GyoumuShizaiChousaRitsu = " + getNumToDb(base_tbl07_2_numPercent1.Text));
            sSql.Append("    ,GyoumuEizenRitsu = " + getNumToDb(base_tbl07_2_numPercent2.Text));
            sSql.Append("    ,GyoumuKikiruiChousaRitsu = " + getNumToDb(base_tbl07_2_numPercent3.Text));
            sSql.Append("    ,GyoumuKoujiChousahiRitsu = " + getNumToDb(base_tbl07_2_numPercent4.Text));
            sSql.Append("    ,GyoumuSanpaiFukusanbutsuRitsu = " + getNumToDb(base_tbl07_2_numPercent5.Text));
            sSql.Append("    ,GyoumuHokakeChousaRitsu = " + getNumToDb(base_tbl07_2_numPercent6.Text));
            sSql.Append("    ,GyoumuShokeihiChousaRitsu = " + getNumToDb(base_tbl07_2_numPercent7.Text));
            sSql.Append("    ,GyoumuGenkaBunsekiRitsu = " + getNumToDb(base_tbl07_2_numPercent8.Text));
            sSql.Append("    ,GyoumuKijunsakuseiRitsu = " + getNumToDb(base_tbl07_2_numPercent9.Text));
            sSql.Append("    ,GyoumuKoukyouRoumuhiRitsu = " + getNumToDb(base_tbl07_2_numPercent10.Text));
            sSql.Append("    ,GyoumuRoumuhiKoukyouigaiRitsu = " + getNumToDb(base_tbl07_2_numPercent11.Text));
            sSql.Append("    ,GyoumuSonotaChousabuRitsu = " + getNumToDb(base_tbl07_2_numPercent12.Text));
            sSql.Append(" WHERE GyoumuAnkenJouhouID = " + sAnkenId + " AND GyoumuHibunKubun = 10 ");
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 業務配分　配分額更新　
        /// </summary>
        /// <param name="cmd"></param>
        private void updateGyoumuHaibun30(SqlCommand cmd, string sAnkenId)
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("UPDATE GyoumuHaibun SET ");
            sSql.Append("     GyoumuChosaBuRitsu = " + getNumToDb(ca_tbl02_AftCaBm_numPercent1.Text));
            sSql.Append("    ,GyoumuJigyoFukyuBuRitsu = " + getNumToDb(ca_tbl02_AftCaBm_numPercent2.Text));
            sSql.Append("    ,GyoumuJyohouSystemBuRitsu = " + getNumToDb(ca_tbl02_AftCaBm_numPercent3.Text));
            sSql.Append("    ,GyoumuSougouKenkyuJoRitsu = " + getNumToDb(ca_tbl02_AftCaBm_numPercent4.Text));

            sSql.Append("    ,GyoumuChosaBuGaku " + " = " + getNumToDb(ca_tbl02_AftCaBm_numAmt1.Text));
            sSql.Append("    ,GyoumuJigyoFukyuBuGaku " + " = " + getNumToDb(ca_tbl02_AftCaBm_numAmt2.Text));
            sSql.Append("    ,GyoumuJyohouSystemBuGaku " + " = " + getNumToDb(ca_tbl02_AftCaBm_numAmt3.Text));
            sSql.Append("    ,GyoumuSougouKenkyuJoGaku " + " = " + getNumToDb(ca_tbl02_AftCaBm_numAmt4.Text));

            sSql.Append("    ,GyoumuShizaiChousaRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent1.Text));
            sSql.Append("    ,GyoumuEizenRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent2.Text));
            sSql.Append("    ,GyoumuKikiruiChousaRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent3.Text));
            sSql.Append("    ,GyoumuKoujiChousahiRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent4.Text));
            sSql.Append("    ,GyoumuSanpaiFukusanbutsuRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent5.Text));
            sSql.Append("    ,GyoumuHokakeChousaRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent6.Text));
            sSql.Append("    ,GyoumuShokeihiChousaRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent7.Text));
            sSql.Append("    ,GyoumuGenkaBunsekiRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent8.Text));
            sSql.Append("    ,GyoumuKijunsakuseiRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent9.Text));
            sSql.Append("    ,GyoumuKoukyouRoumuhiRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent10.Text));
            sSql.Append("    ,GyoumuRoumuhiKoukyouigaiRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent11.Text));
            sSql.Append("    ,GyoumuSonotaChousabuRitsu " + " = " + getNumToDb(ca_tbl02_AftCaTs_numPercent12.Text));

            sSql.Append("    ,GyoumuShizaiChousaGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt1.Text));
            sSql.Append("    ,GyoumuEizenGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt2.Text));
            sSql.Append("    ,GyoumuKikiruiChousaGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt3.Text));
            sSql.Append("    ,GyoumuKoujiChousahiGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt4.Text));
            sSql.Append("    ,GyoumuSanpaiFukusanbutsuGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt5.Text));
            sSql.Append("    ,GyoumuHokakeChousaGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt6.Text));
            sSql.Append("    ,GyoumuShokeihiChousaGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt7.Text));
            sSql.Append("    ,GyoumuGenkaBunsekiGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt8.Text));
            sSql.Append("    ,GyoumuKijunsakuseiGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt9.Text));
            sSql.Append("    ,GyoumuKoukyouRoumuhiGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt10.Text));
            sSql.Append("    ,GyoumuRoumuhiKoukyouigaiGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt11.Text));
            sSql.Append("    ,GyoumuSonotaChousabuGaku " + " = " + getNumToDb(ca_tbl02_AftCaTs_numAmt12.Text));
            sSql.Append(" WHERE GyoumuAnkenJouhouID = " + sAnkenId + " AND GyoumuHibunKubun = 30 ");

            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 技術者評価：担当技術者
        /// </summary>
        /// <param name="cmd"></param>
        private void updateGyoumuJouhouHyouronTantouL1(SqlCommand cmd, C1FlexGrid c1FlexGrid5)
        {
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
        }

        /// <summary>
        /// 窓口担当者 更新
        /// </summary>
        /// <param name="cmd"></param>
        private void updateGyoumuJouhouMadoguchi(SqlCommand cmd)
        {
            // 名称の取得
            string GyoumuJouhouMadoShibuMei = "";   // 業務情報窓口支部名
            string GyoumuJouhouMadoKamei = "";      // 業務情報窓口課名
            DataTable dt2 = new DataTable();
            cmd.CommandText = "SELECT ShibuMei, KaMei FROM Mst_Busho WHERE GyoumuBushoCD = '" + ca_tbl05_txtMadoguchiBusho.Text + "'";

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


            //窓口担当者の更新
            if ((ca_tbl05_txtMadoguchiCD.Text == "0") || (ca_tbl05_txtMadoguchiCD.Text == ""))
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
                                    "GyoumuJouhouMadoGyoumuBushoCD = N'" + ca_tbl05_txtMadoguchiBusho.Text + "' " +
                                    ",GyoumuJouhouMadoShibuMei = N'" + GyoumuJouhouMadoShibuMei + "' " +
                                    ",GyoumuJouhouMadoKamei = N'" + GyoumuJouhouMadoKamei + "' " +
                                    ",GyoumuJouhouMadoKojinCD = N'" + ca_tbl05_txtMadoguchiCD.Text + "' " +
                                    ",GyoumuJouhouMadoChousainMei = N'" + ca_tbl05_txtMadoguchi.Text + "' " +
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
                            ",N'" + ca_tbl05_txtMadoguchiBusho.Text + "' " +
                            ",N'" + GyoumuJouhouMadoShibuMei + "' " +
                            ",N'" + GyoumuJouhouMadoKamei + "' " +
                            ",N'" + ca_tbl05_txtMadoguchiCD.Text + "' " +
                            ",N'" + ca_tbl05_txtMadoguchi.Text + "' " +
                            ") ";
                    Console.WriteLine(cmd.CommandText);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 入札応札者
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="c1FlexGrid2"></param>
        /// <param name="nyuusatsuOusatushaInsertFlg">入札応札者新規登録フラグ true：初回登録（NyuusatsuRakusatsuShokaiDate を更新）</param>
        /// <param name="nyuusatsuOusatushaUpdateFlg">入札応札者更新フラグ true：更新（NyuusatsuRakusatsuSaisyuDate を更新）</param>
        private void updateNyuusatsuJouhouOusatsusha(SqlCommand cmd, C1FlexGrid c1FlexGrid2, 
            ref bool nyuusatsuOusatushaInsertFlg, ref bool nyuusatsuOusatushaUpdateFlg, ref int nyusatsuCnt)
        {
            // 入札応札者新規登録フラグ true：初回登録（NyuusatsuRakusatsuShokaiDate を更新）
            // 既にNyuusatsuRakusatsuShokaiDateが入っていない（NULL）で、NyuusatsuJouhouOusatsushaテーブルにデータを登録した場合、true
            nyuusatsuOusatushaInsertFlg = false;
            // 入札応札者更新フラグ true：更新（NyuusatsuRakusatsuSaisyuDate を更新）
            // 応札者のGridを回して、NyuusatsuJouhouOusatsushaテーブルと一致しないデータが存在した場合、true
            nyuusatsuOusatushaUpdateFlg = false;

            // 登録日データフラグ false:データ登録なし true：データ登録あり
            bool nyuusatsuOusatushaShokaiDateFlg = false;

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


            cmd.CommandText = "DELETE NyuusatsuJouhouOusatsusha WHERE NyuusatsuJouhouID = '" + AnkenID + "' ";
            cmd.ExecuteNonQuery();

            //入札数の計算のため、入札情報前に入札応札者を登録
            nyusatsuCnt = 0;
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
        }

        /// <summary>
        /// 入札情報　更新
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="bUpdateFlg"></param>
        /// <param name="bInsertFlg"></param>
        /// <param name="nyusatsuCnt"></param>
        private void updateNyuusatsuJouhou(SqlCommand cmd, bool bUpdateFlg, bool bInsertFlg, int nyusatsuCnt)
        {
            bool nyuusatsuOusatushaUpdateFlg = bUpdateFlg;
            bool nyuusatsuOusatushaInsertFlg = bInsertFlg;

            // 入札情報
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
                bid_tbl03_4_dtpInsDate.Text = DateTime.Today.ToString();
                bid_tbl03_4_dtpUpdDate.Text = DateTime.Today.ToString();
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
                bid_tbl03_4_dtpUpdDate.Text = DateTime.Today.ToString();
            }

            // No.278 競合他社IDが入っていない対応
            string KyougouTashaID = "";
            // 落札者がいれば、競合他社IDを取得する
            if (bid_tbl03_1_txtRakusatuSya.Text != null && bid_tbl03_1_txtRakusatuSya.Text != "")
            {
                cmd.CommandText = "SELECT  " +
                "KyougouTashaID " +
                "FROM Mst_KyougouTasha " +
                "WHERE KyougouMeishou = N'" + bid_tbl03_1_txtRakusatuSya.Text + "' ";
                var sda = new SqlDataAdapter(cmd);
                var dt = new DataTable();
                sda.Fill(dt);
                KyougouTashaID = dt.Rows[0][0].ToString();
            }

            cmd.CommandText = "UPDATE NyuusatsuJouhou SET " +
                            "NyuusatsuMitsumorigaku = " + "N'" + getNumToDb(base_tbl10_numSankoMitumoriAmt.Text) + "'" +
                            ",NyuusatsuRakusatsusha = " + " N'" + GlobalMethod.ChangeSqlText(bid_tbl03_1_txtRakusatuSya.Text, 0, 0) + "' " +
                            ",NyuusatsuRakusatsuKekkaDate = " + " " + Get_DateTimePicker("bid_tbl03_1_dtpBidResultDt") + " " +
                            ",NyuusatsuRakusatugaku = " + "N'" + getNumToDb(bid_tbl03_1_numRakusatuAmt.Text) + "' " +
                            ",NyuusatsuRakusatuSougaku = " + "N'" + getNumToDb(bid_tbl03_1_numRakusatuAmt.Text) + "' " +
                            ",NyuusatsuYoteiKakaku = " + "N'" + getNumToDb(bid_tbl03_1_txtYoteiPrice.Text) + "' " +
                            ",NyuusatsuTanpinMikomigaku = " + "N'" + getNumToDb(bid_tbl03_1_txtYoteiPrice.Text) + "' " +
                            ",NyuusatsuUpdateProgram = " + "'UpdateEntry' " +
                            ",NyuusatsuUpdateDate = " + " GETDATE() " +
                            ",NyuusatsuUpdateUser = " + "N'" + UserInfos[0] + "' " +
                            ",NyuusatsuHoushiki = " + "N'" + bid_tbl01_cmbBidhosiki.SelectedValue + "' " +
                            ",NyuusatsuGyoumuBikou = " + "N'" + GlobalMethod.ChangeSqlText(base_tbl03_txtAnkenMemo.Text, 0, 0) + "' " +
                            ",NyuusatsushaSuu = " + "N'" + nyusatsuCnt + "' " +
                            //",NyuusatsuKekkaMemo = " + "N'" + GlobalMethod.ChangeSqlText(item2_3_12.Text, 0, 0) + "' " + なくなったなんで？？？？★★★
                            ",NyuusatsuRakusatsushaID = " + (IsNotSelected(bid_tbl03_1_cmbBidStatus) ? "NULL" : "N'" + bid_tbl03_1_cmbBidStatus.SelectedValue.ToString() + "' ") +
                            ",NyuusatsuRakusatsuShaJokyou = " + "" + bid_tbl03_1_cmbRakusatuStatus.SelectedValue.ToString() + " " +
                            ",NyuusatsuRakusatsuGakuJokyou = " + "" + bid_tbl03_1_cmbRakusatuAmtStatus.SelectedValue.ToString() + " " +
                            ",NyuusatsuKyougouTasha = " + " N'" + GlobalMethod.ChangeSqlText(bid_tbl03_1_txtRakusatuSya.Text, 0, 0) + "' ";
            cmd.CommandText += ", NyuusatsuAnkenMemoNuusatsu = " + " N'" + GlobalMethod.ChangeSqlText(bid_tbl03_1_txtBidMemo.Text, 0, 0) + "' "; // --案件メモ(入札)
            cmd.CommandText += ", NyuusatsuSaiitakuSonotaNaiyou = " + " N'" + GlobalMethod.ChangeSqlText(bid_tbl02_txtOtherNaiyo.Text, 0, 0) + "' "; // --その他の内容
            cmd.CommandText += ", NyuusatsuSaiitakuKinshiNaiyou = " + "N'" + bid_tbl02_cmbKinsiNaiyo.SelectedValue + "' ";// --再委託禁止条項の内容
            cmd.CommandText += ", NyuusatsuSaiitakuKinshiUmu = " + "N'" + bid_tbl02_cmbKinsiUmu.SelectedValue + "' ";// --再委託禁止条項の記載有無
            cmd.CommandText += ", NyuusatsuJuchuuIyoku = " + "N'" + bid_tbl01_cmbOrderIyoku.SelectedValue + "' ";// --受注意欲
            cmd.CommandText += ", NyuusatsuSankoumitsumoriKingaku = " + "N'" + getNumToDb(bid_tbl01_txtMitumoriAmt.Text) + "' ";// --参考見積額(税抜)
            cmd.CommandText += ", NyuusatsuSankoumitsumoriTaiou = " + "N'" + bid_tbl01_cmbMitumori.SelectedValue + "' ";// --参考見積対応
            cmd.CommandText += ", NyuusatsuSaiteiKakakuUmu = " + "N'" + bid_tbl01_cmbLowestUmu.SelectedValue + "' ";// --最低制限価格有無
            cmd.CommandText += ", NyuusatsuGyoumuHachuukubun = " + "N'" + bid_tbl01_cmbOrderKubun.SelectedValue + "' ";// --業務発注区分
            cmd.CommandText += ", NyuusatsuJouhouTourokubi = " + Get_DateTimePicker("bid_tbl01_dtpBidInfoDt");
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

        }

        /// <summary>
        /// 窓口情報更新
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="flg">1:更新、2：チェックシート、3:起案</param>
        private void updateMadoguchiJouhou(SqlCommand cmd, int flg = 1)
        {
            StringBuilder sSql = new StringBuilder();

            sSql.Append("UPDATE MadoguchiJouhou SET ");
            sSql.Append("    MadoguchiJutakuBangou = replace(AnkenJutakuBangou,'-' + AnkenJutakuBangouEda,'') ");
            sSql.Append("    ,MadoguchiJutakuBangouEdaban = AnkenJutakuBangouEda ");
            sSql.Append("    ,MadoguchiJutakuBushoCD = AnkenJutakubushoCD ");
            sSql.Append("    ,MadoguchiJutakubushoMeiOld = ShibuMei ");
            sSql.Append("    ,MadoguchiJutakuTantoushaID = AnkenTantoushaCD ");
            sSql.Append("    ,MadoguchiJutakuTantoushaOld = ChousainMei ");
            sSql.Append("    ,MadoguchiKanriGijutsusha = KanriGijutsushaCD " );
            if(flg == 1) sSql.Append("    ,MadoguchiGyoumuKanrishaCD = GyoumuKanrishaCD ");
            sSql.Append(" FROM AnkenJouhou ");
            sSql.Append("     LEFT JOIN Mst_Busho ON GyoumuBushoCD = AnkenJutakubushoCD ");
            sSql.Append("     LEFT JOIN Mst_Chousain ON KojinCD = AnkenTantoushaCD ");
            sSql.Append("     LEFT JOIN GyoumuJouhou ON GyoumuJouhou.AnkenJouhouID = AnkenJouhou.AnkenJouhouID ");
            if (flg == 1) sSql.Append("     LEFT JOIN GyoumuJouhouMadoguchi ON GyoumuJouhouMadoguchi.GyoumuJouhouID = AnkenJouhou.AnkenJouhouID ");
            sSql.Append(" WHERE MadoguchiJouhou.AnkenJouhouID = AnkenJouhou.AnkenJouhouID AND MadoguchiJouhou.AnkenJouhouID = " + AnkenID);
            if (flg == 1) sSql.Append("   AND GyoumuJouhouMadoGyoumuBushoCD is not NULL "); // 既に窓口に連携済の場合、NULLが不可なので、更新をスルーさせる
            cmd.CommandText = sSql.ToString();
            cmd.ExecuteNonQuery();
        }

        // Garoon宛先追加登録
        private void insertGaroonAtesakiTsuika(SqlCommand cmd)
        {
            // 管理技術者が空でない場合、Garoon連携宛先追加テーブルに追加する
            if (ca_tbl05_txtKanri.Text != null && ca_tbl05_txtKanri.Text != "")
            {
                // Garoon連携宛先追加テーブルに存在するか確認
                cmd.CommandText = "SELECT  " +
                    " mj.MadoguchiID " +
                    ",gta.GaroonTsuikaAtesakiMadoguchiID " +
                    "FROM AnkenJouhou aj " +
                    "INNER JOIN MadoguchiJouhou mj ON mj.AnkenJouhouID = aj.AnkenJouhouID " +
                    "LEFT  JOIN GaroonTsuikaAtesaki gta ON gta.GaroonTsuikaAtesakiMadoguchiID = mj.MadoguchiID AND GaroonTsuikaAtesakiTantoushaCD = '" + ca_tbl05_txtKanriCD.Text + "'" +
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
                      ",BushokanriboKamei " +
                      "FROM Mst_Chousain mc " +
                      "LEFT JOIN Mst_Busho mb ON mb.GyoumuBushoCD = mc.GyoumuBushoCD " +
                      "WHERE mc.KojinCD = '" + ca_tbl05_txtKanriCD.Text + "' ";

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
        #endregion

        #region 伝票変更モード処理メソッド --------
        /// <summary>
        /// 赤伝作成
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenNo">赤伝案件番号</param>
        /// <param name="SakuseiKubun">契約区分</param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝 </param>
        /// <returns></returns>
        private int createAnkenJouhou(SqlCommand cmd, string ankenNo, string SakuseiKubun, int iType = 0)
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("INSERT INTO AnkenJouhou ( ");
            #region Insert　Cols
            sSql.Append("    AnkenJouhouID ");
            sSql.Append("    ,AnkenSakuseiKubun ");
            sSql.Append("    ,AnkenSaishinFlg ");
            sSql.Append("    ,AnkenKishuKeikakugaku ");
            sSql.Append("    ,AnkenKishuKeikakakugakuJf ");
            sSql.Append("    ,AnkenKishuKeikakugakuJ ");
            sSql.Append("    ,AnkenKeikakuZangaku ");
            sSql.Append("    ,AnkenkeikakuZangakuJF ");
            sSql.Append("    ,AnkenkeikakuZangakuJ ");
            sSql.Append("    ,AnkenChokusetsuGenka ");
            sSql.Append("    ,AnkenChokusetsuGenkaRitsu ");
            sSql.Append("    ,AnkenGaichuuhi ");
            sSql.Append("    ,AnkenJoukanDoboku ");
            sSql.Append("    ,AnkenJoukanFukugou ");
            sSql.Append("    ,AnkenJoukanGesuidou ");
            sSql.Append("    ,AnkenJoukanHyoujun ");
            sSql.Append("    ,AnkenJoukanIchiba ");
            sSql.Append("    ,AnkenJoukanItiji ");
            sSql.Append("    ,AnkenJoukanJutakuSonota ");
            sSql.Append("    ,AnkenJoukanKentiku ");
            sSql.Append("    ,AnkenJoukanKijunsho ");
            sSql.Append("    ,AnkenJoukanKouwan ");
            sSql.Append("    ,AnkenJoukanKuukou ");
            sSql.Append("    ,AnkenJoukanSetsubi ");
            sSql.Append("    ,AnkenJoukanSonota ");
            sSql.Append("    ,AnkenJoukanSuidou ");
            sSql.Append("    ,AnkenKeichoukaiKounyuuhi ");
            sSql.Append("    ,AnkenKishuKeikakugakuK ");
            sSql.Append("    ,AnkenKaisuu ");
            sSql.Append("    ,AnkenCreateDate ");
            sSql.Append("    ,AnkenCreateUser ");
            sSql.Append("    ,AnkenCreateProgram ");
            sSql.Append("    ,AnkenUpdateDate ");
            sSql.Append("    ,AnkenUpdateUser ");
            sSql.Append("    ,AnkenUpdateProgram ");
            sSql.Append("    ,AnkenTourokubi ");
            sSql.Append("    ,AnkenGyoumuMei ");
            sSql.Append("    ,AnkenDeleteFlag ");
            sSql.Append("    ,AnkenUriageNendo ");
            sSql.Append("    ,AnkenHachushaKubunCD ");
            sSql.Append("    ,AnkenHachushaKubunMei ");
            sSql.Append("    ,AnkenHachuushaCodeID ");
            sSql.Append("    ,AnkenHachuushaMei ");
            sSql.Append("    ,AnkenGyoumuKubun ");
            sSql.Append("    ,AnkenGyoumuKubunMei ");
            sSql.Append("    ,AnkenNyuusatsuHoushiki ");
            sSql.Append("    ,AnkenKyougouTasha ");
            sSql.Append("    ,AnkenJutakubushoCD ");
            sSql.Append("    ,AnkenJutakushibu ");
            sSql.Append("    ,AnkenTantoushaCD ");
            sSql.Append("    ,AnkenMadoguchiTantoushaCD ");
            sSql.Append("    ,AnkenGyoumuKanrishaCD ");
            sSql.Append("    ,AnkenGyoumuKanrisha ");
            sSql.Append("    ,GyoumuKanrishaCD ");
            sSql.Append("    ,AnkenHachuushaBusho ");
            sSql.Append("    ,AnkenkeikakuZangakuK ");
            sSql.Append("    ,AnkenJutakuBangou ");
            sSql.Append("    ,AnkenJutakuBangouEda ");
            sSql.Append("    ,AnkenNyuusatsuYoteibi ");
            sSql.Append("    ,AnkenRakusatsusha ");
            sSql.Append("    ,AnkenRakusatsuJouhou ");
            sSql.Append("    ,AnkenKianZumi ");
            sSql.Append("    ,AnkenKiangetsu ");
            sSql.Append("    ,AnkenHanteiKubun ");
            sSql.Append("    ,AnkenJoukanData ");
            sSql.Append("    ,AnkenJoukanHachuuKikanCD ");
            sSql.Append("    ,AnkenNyuukinKakuninbi ");
            sSql.Append("    ,AnkenKanryouSakuseibi ");
            sSql.Append("    ,AnkenHonbuKakuninbi ");
            sSql.Append("    ,AnkenShizaiChousa ");
            sSql.Append("    ,AnkenKoujiChousahi ");
            sSql.Append("    ,AnkenKikiruiChousa ");
            sSql.Append("    ,AnkenSanpaiFukusanbutsu ");
            sSql.Append("    ,AnkenHokakeChousa ");
            sSql.Append("    ,AnkenShokeihiChousa ");
            sSql.Append("    ,AnkenGenkaBunseki ");
            sSql.Append("    ,AnkenKijunsakusei ");
            sSql.Append("    ,AnkenKoukyouRoumuhi ");
            sSql.Append("    ,AnkenRoumuhiKoukyouigai ");
            sSql.Append("    ,AnkenSonotaChousabu ");
            sSql.Append("    ,AnkenOrdermadeJifubu ");
            sSql.Append("    ,AnkenRIBCJifubu ");
            sSql.Append("    ,AnkenSonotaJifubu ");
            sSql.Append("    ,AnkenOrdermade ");
            sSql.Append("    ,AnkenJouhouKaihatsu ");
            sSql.Append("    ,AnkenRIBCJouhouKaihatsu ");
            sSql.Append("    ,AnkenSoukenbu ");
            sSql.Append("    ,AnkenSonotaJoujibu ");
            sSql.Append("    ,AnkenTeikiTokuchou ");
            sSql.Append("    ,AnkenTanpinTokuchou ");
            sSql.Append("    ,AnkenKikiChousa ");
            sSql.Append("    ,AnkenHachuushaIraibusho ");
            sSql.Append("    ,AnkenHachuushaTantousha ");
            sSql.Append("    ,AnkenHachuushaTEL ");
            sSql.Append("    ,AnkenHachuushaFAX ");
            sSql.Append("    ,AnkenHachuushaMail ");
            sSql.Append("    ,AnkenHachuushaIraiYuubin ");
            sSql.Append("    ,AnkenHachuushaIraiJuusho ");
            sSql.Append("    ,AnkenHachuushaKeiyakuBusho ");
            sSql.Append("    ,AnkenHachuushaKeiyakuTantou ");
            sSql.Append("    ,AnkenHachuushaKeiyakuTEL ");
            sSql.Append("    ,AnkenHachuushaKeiyakuFAX ");
            sSql.Append("    ,AnkenHachuushaKeiyakuMail ");
            sSql.Append("    ,AnkenHachuushaKeiyakuYuubin ");
            sSql.Append("    ,AnkenHachuushaKeiyakuJuusho ");
            sSql.Append("    ,AnkenHachuuDaihyouYakushoku ");
            sSql.Append("    ,AnkenHachuuDaihyousha ");
            sSql.Append("    ,AnkenRosenKawamei ");
            sSql.Append("    ,AnkenGyoumuItakuKasho ");
            sSql.Append("    ,AnkenJititaiKibunID ");
            sSql.Append("    ,AnkenJititaiKubun ");
            sSql.Append("    ,AnkenKeiyakuToshoNo ");
            sSql.Append("    ,AnkenKirokuToshoNo ");
            sSql.Append("    ,AnkenKirokuHokanNo ");
            sSql.Append("    ,AnkenCDHokan ");
            sSql.Append("    ,AnkenSeikaButsuHokanFile ");
            sSql.Append("    ,AnkenSeikabutsuHokanbako ");
            sSql.Append("    ,AnkenKokyakuHyoukaComment ");
            sSql.Append("    ,AnkenToukaiHyoukaComment ");
            sSql.Append("    ,AnkenKenCD ");
            sSql.Append("    ,AnkenToshiCD ");
            sSql.Append("    ,AnkenKeiyakusho ");
            sSql.Append("    ,AnkenEizen ");
            sSql.Append("    ,AnkenTantoushaMei ");
            sSql.Append("    ,GyoumuKanrishaMei ");
            sSql.Append("    ,AnkenGyoumuKubunCD ");
            sSql.Append("    ,AnkenHachuushaKaMei ");
            sSql.Append("    ,AnkenKeiyakuKoukiKaishibi ");
            sSql.Append("    ,AnkenKeiyakuKoukiKanryoubi ");
            sSql.Append("    ,AnkenKeiyakuTeiketsubi ");
            sSql.Append("    ,AnkenKeiyakuZeikomiKingaku ");
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuC ");
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuJ ");
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuJs ");
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuK ");
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuR ");
            sSql.Append("    ,AnkenKeiyakuSakuseibi ");
            sSql.Append("    ,AnkenAnkenBangou ");
            sSql.Append("    ,AnkenKeikakuBangou ");
            sSql.Append("    ,AnkenHikiaijhokyo ");
            sSql.Append("    ,AnkenKeikakuAnkenMei ");
            sSql.Append("    ,AnkenToukaiSankouMitsumori ");
            sSql.Append("    ,AnkenToukaiJyutyuIyoku ");
            sSql.Append("    ,AnkenToukaiSankouMitsumoriGaku ");
            sSql.Append("    ,AnkenHachushaKaMei ");
            sSql.Append("    ,AnkenHachushaCD ");
            sSql.Append("    ,AnkenToukaiOusatu ");
            sSql.Append("    ,AnkenKoukiNendo ");
            sSql.Append("    ,AnkenAnkenMemoMihachuu ");    //-- 案件メモ(未発注)
            sSql.Append("    ,AnkenSonotaNaiyou ");         //-- 「その他」の内容
            sSql.Append("    ,AnkenMihachuuTourokubi ");    //-- 未発注の登録日
            sSql.Append("    ,AnkenAnkenMemoJizendashin "); //-- 案件メモ（事前打診）
            sSql.Append("    ,AnkenHachuunashiRiyuu ");     //--「発注なし」の理由
            sSql.Append("    ,AnkenMihachuuJoukyou ");      //--未発注状況
            sSql.Append("    ,AnkenHachuuYoteiMikomibi ");  //--発注予定・見込日
            sSql.Append("    ,AnkenJizenDashinIraibi ");    //--事前打診依頼日
            sSql.Append("    ,AnkenOuenIraiMemo ");         //--応援依頼メモ
            sSql.Append("    ,AnkenOueniraiUmu ");          //--応援依頼の有無
            sSql.Append("    ,AnkenAnkenMemoKihon ");       //--案件メモ（基本情報）
            sSql.Append("    ,AnkenKeiyakuDate ");          //--契約登録日
            sSql.Append("    ,AnkenKeiyakuCheck ");         //--契約
            sSql.Append("    ,AnkenNyuusatuDate ");         //--入札登録日
            sSql.Append("    ,AnkenNyuusatuCheck ");        //--入札
            sSql.Append("    ,AnkenJizenDashinDate ");      //--事前打診登録日
            sSql.Append("    ,AnkenJizenDashinCheck ");     //--事前打診

            #endregion

            sSql.Append("     ) SELECT ");
            sSql.Append(ankenNo);           //AnkenJouhouID

           
            if (SakuseiKubun == "03" || int.Parse(SakuseiKubun) > 5)
            {
                if (iType == 0 || iType == 70)
                {
                    // 赤伝なら／ダミーデータ赤伝
                    sSql.Append("    ,'02' ");  //AnkenSakuseiKubun
                    sSql.Append("    ,0 ");     //AnkenSaishinFlg
                    sSql.Append("    ,- AnkenKishuKeikakugaku ");       //AnkenKishuKeikakugaku
                    sSql.Append("    ,- AnkenKishuKeikakakugakuJf ");   //AnkenKishuKeikakakugakuJf
                    sSql.Append("    ,- AnkenKishuKeikakugakuJ ");      // AnkenKishuKeikakugakuJ
                }
                else
                {
                    // 黒伝なら
                    if (iType == 1)
                    {
                        sSql.Append("    ,'03' ");  //AnkenSakuseiKubun
                        sSql.Append("    ,1 ");     //AnkenSaishinFlg
                    }
                    else
                    {
                        sSql.Append(",'").Append(SakuseiKubun).Append("' ");  //",AnkenSakuseiKubun "
                        sSql.Append(",0 ");       //",AnkenSaishinFlg "
                    }
                    sSql.Append("    ,AnkenKishuKeikakugaku ");       //AnkenKishuKeikakugaku
                    sSql.Append("    ,AnkenKishuKeikakakugakuJf ");   //AnkenKishuKeikakakugakuJf
                    sSql.Append("    ,AnkenKishuKeikakugakuJ ");      // AnkenKishuKeikakugakuJ
                }
            }
            else
            {
                sSql.Append("    ,'04' ");
                if (iType == 70 || iType == 71)
                {
                    sSql.Append("    ,0 ");
                }
                else
                {
                    sSql.Append("    ,1 ");
                }
                sSql.Append("    ,0 ");
                sSql.Append("    ,0 ");
                sSql.Append("    ,0 ");
            }

            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenKeikakuZangaku ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenkeikakuZangakuJF ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenkeikakuZangakuJ ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenChokusetsuGenka ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenChokusetsuGenkaRitsu ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenGaichuuhi ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanDoboku ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanFukugou ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanGesuidou ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanHyoujun ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanIchiba ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanItiji ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanJutakuSonota ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanKentiku ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanKijunsho ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanKouwan ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanKuukou ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanSetsubi ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanSonota ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenJoukanSuidou ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenKeichoukaiKounyuuhi ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" AnkenKishuKeikakugakuK ");
            sSql.Append("    ,AnkenKaisuu + 1 ");
            sSql.Append("    ,GETDATE() ");
            sSql.Append("    ,'" + UserInfos[0] + "' ");
            sSql.Append("    ,'ChangeKianEntry' ");
            sSql.Append("    ,GETDATE() ");
            sSql.Append("    ,'" + UserInfos[0] + "' ");
            sSql.Append("    ,'ChangeKianEntry' ");
            sSql.Append("    ,AnkenTourokubi " );
            sSql.Append("    ,AnkenGyoumuMei ");
            if (iType == 0)
            {
                // 赤伝なら
                sSql.Append("    ,0 ");
            }
            else if(iType == 1)
            {
                sSql.Append("    ,AnkenDeleteFlag ");
            }
            else
            {
                sSql.Append("    ,1 ");
            }
            sSql.Append("    ,AnkenUriageNendo ");
            sSql.Append("    ,AnkenHachushaKubunCD ");
            sSql.Append("    ,AnkenHachushaKubunMei ");
            sSql.Append("    ,AnkenHachuushaCodeID ");
            sSql.Append("    ,AnkenHachuushaMei ");
            sSql.Append("    ,AnkenGyoumuKubun ");
            sSql.Append("    ,AnkenGyoumuKubunMei ");
            sSql.Append("    ,AnkenNyuusatsuHoushiki ");
            sSql.Append("    ,AnkenKyougouTasha ");
            sSql.Append("    ,AnkenJutakubushoCD ");
            sSql.Append("    ,AnkenJutakushibu ");
            sSql.Append("    ,AnkenTantoushaCD ");
            sSql.Append("    ,AnkenMadoguchiTantoushaCD ");
            sSql.Append("    ,AnkenGyoumuKanrishaCD ");
            sSql.Append("    ,AnkenGyoumuKanrisha ");
            sSql.Append("    ,GyoumuKanrishaCD ");
            sSql.Append("    ,AnkenHachuushaBusho ");
            sSql.Append("    ,AnkenkeikakuZangakuK ");
            sSql.Append("    ,AnkenJutakuBangou ");
            sSql.Append("    ,AnkenJutakuBangouEda ");
            sSql.Append("    ,AnkenNyuusatsuYoteibi ");
            sSql.Append("    ,AnkenRakusatsusha ");
            sSql.Append("    ,AnkenRakusatsuJouhou ");
            sSql.Append("    ,AnkenKianZumi ");
            sSql.Append("    ,AnkenKiangetsu ");
            sSql.Append("    ,AnkenHanteiKubun ");
            sSql.Append("    ,AnkenJoukanData ");
            sSql.Append("    ,AnkenJoukanHachuuKikanCD ");
            sSql.Append("    ,AnkenNyuukinKakuninbi ");
            sSql.Append("    ,AnkenKanryouSakuseibi ");
            sSql.Append("    ,AnkenHonbuKakuninbi ");
            sSql.Append("    ,AnkenShizaiChousa ");
            sSql.Append("    ,AnkenKoujiChousahi ");
            sSql.Append("    ,AnkenKikiruiChousa ");
            sSql.Append("    ,AnkenSanpaiFukusanbutsu ");
            sSql.Append("    ,AnkenHokakeChousa ");
            sSql.Append("    ,AnkenShokeihiChousa ");
            sSql.Append("    ,AnkenGenkaBunseki ");
            sSql.Append("    ,AnkenKijunsakusei ");
            sSql.Append("    ,AnkenKoukyouRoumuhi ");
            sSql.Append("    ,AnkenRoumuhiKoukyouigai ");
            sSql.Append("    ,AnkenSonotaChousabu ");
            sSql.Append("    ,AnkenOrdermadeJifubu ");
            sSql.Append("    ,AnkenRIBCJifubu ");
            sSql.Append("    ,AnkenSonotaJifubu ");
            sSql.Append("    ,AnkenOrdermade ");
            sSql.Append("    ,AnkenJouhouKaihatsu ");
            sSql.Append("    ,AnkenRIBCJouhouKaihatsu ");
            sSql.Append("    ,AnkenSoukenbu ");
            sSql.Append("    ,AnkenSonotaJoujibu ");
            sSql.Append("    ,AnkenTeikiTokuchou ");
            sSql.Append("    ,AnkenTanpinTokuchou ");
            sSql.Append("    ,AnkenKikiChousa ");
            sSql.Append("    ,AnkenHachuushaIraibusho ");
            sSql.Append("    ,AnkenHachuushaTantousha ");
            sSql.Append("    ,AnkenHachuushaTEL ");
            sSql.Append("    ,AnkenHachuushaFAX ");
            sSql.Append("    ,AnkenHachuushaMail ");
            sSql.Append("    ,AnkenHachuushaIraiYuubin ");
            sSql.Append("    ,AnkenHachuushaIraiJuusho ");
            sSql.Append("    ,AnkenHachuushaKeiyakuBusho ");
            sSql.Append("    ,AnkenHachuushaKeiyakuTantou ");
            sSql.Append("    ,AnkenHachuushaKeiyakuTEL ");
            sSql.Append("    ,AnkenHachuushaKeiyakuFAX ");
            sSql.Append("    ,AnkenHachuushaKeiyakuMail ");
            sSql.Append("    ,AnkenHachuushaKeiyakuYuubin ");
            sSql.Append("    ,AnkenHachuushaKeiyakuJuusho ");
            sSql.Append("    ,AnkenHachuuDaihyouYakushoku ");
            sSql.Append("    ,AnkenHachuuDaihyousha ");
            sSql.Append("    ,AnkenRosenKawamei ");
            sSql.Append("    ,AnkenGyoumuItakuKasho ");
            sSql.Append("    ,AnkenJititaiKibunID ");
            sSql.Append("    ,AnkenJititaiKubun ");
            sSql.Append("    ,AnkenKeiyakuToshoNo ");
            sSql.Append("    ,AnkenKirokuToshoNo ");
            sSql.Append("    ,AnkenKirokuHokanNo ");
            sSql.Append("    ,AnkenCDHokan ");
            sSql.Append("    ,AnkenSeikaButsuHokanFile ");
            sSql.Append("    ,AnkenSeikabutsuHokanbako ");
            sSql.Append("    ,AnkenKokyakuHyoukaComment ");
            sSql.Append("    ,AnkenToukaiHyoukaComment ");
            sSql.Append("    ,AnkenKenCD ");
            sSql.Append("    ,AnkenToshiCD ");
            sSql.Append("    ,AnkenKeiyakusho ");
            sSql.Append("    ,AnkenEizen ");
            sSql.Append("    ,AnkenTantoushaMei ");
            sSql.Append("    ,GyoumuKanrishaMei ");
            sSql.Append("    ,AnkenGyoumuKubunCD ");
            sSql.Append("    ,AnkenHachuushaKaMei ");
            sSql.Append("    ,AnkenKeiyakuKoukiKaishibi ");
            sSql.Append("    ,AnkenKeiyakuKoukiKanryoubi ");
            sSql.Append("    ,AnkenKeiyakuTeiketsubi ");
            sSql.Append("    ,AnkenKeiyakuZeikomiKingaku ");     // 契約タブの契約金額の税込
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuC " );  // 契約タブの受託金額配分の調査部、配分額（税込）
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuJ ");  // 契約タブの受託金額配分の事業普及部、配分額（税込）
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuJs "); // 契約タブの受託金額配分の情報システム部、配分額（税込）
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuK ");  // 契約タブの受託金額配分の総合研究所、配分額（税込）
            sSql.Append("    ,AnkenKeiyakuUriageHaibunGakuR "); // なし
            sSql.Append("    ,AnkenKeiyakuSakuseibi ");
            sSql.Append("    ,AnkenAnkenBangou ");
            sSql.Append("    ,AnkenKeikakuBangou ");
            sSql.Append("    ,AnkenHikiaijhokyo ");
            sSql.Append("    ,AnkenKeikakuAnkenMei ");
            sSql.Append("    ,AnkenToukaiSankouMitsumori ");
            sSql.Append("    ,AnkenToukaiJyutyuIyoku ");
            sSql.Append("    ,AnkenToukaiSankouMitsumoriGaku ");
            sSql.Append("    ,AnkenHachushaKaMei ");
            sSql.Append("    ,AnkenHachushaCD ");
            sSql.Append("    ,AnkenToukaiOusatu ");
            sSql.Append("    ,AnkenKoukiNendo ");

            sSql.Append("    ,AnkenAnkenMemoMihachuu ");    //-- 案件メモ(未発注)
            sSql.Append("    ,AnkenSonotaNaiyou ");         //-- 「その他」の内容
            sSql.Append("    ,AnkenMihachuuTourokubi ");    //-- 未発注の登録日
            sSql.Append("    ,AnkenAnkenMemoJizendashin "); //-- 案件メモ（事前打診）
            sSql.Append("    ,AnkenHachuunashiRiyuu ");     //--「発注なし」の理由
            sSql.Append("    ,AnkenMihachuuJoukyou ");      //--未発注状況
            sSql.Append("    ,AnkenHachuuYoteiMikomibi ");  //--発注予定・見込日
            sSql.Append("    ,AnkenJizenDashinIraibi ");    //--事前打診依頼日
            sSql.Append("    ,AnkenOuenIraiMemo ");         //--応援依頼メモ
            sSql.Append("    ,AnkenOueniraiUmu ");          //--応援依頼の有無
            sSql.Append("    ,AnkenAnkenMemoKihon ");       //--案件メモ（基本情報）
            sSql.Append("    ,AnkenKeiyakuDate ");          //--契約登録日
            sSql.Append("    ,AnkenKeiyakuCheck ");         //--契約
            sSql.Append("    ,AnkenNyuusatuDate ");         //--入札登録日
            sSql.Append("    ,AnkenNyuusatuCheck ");        //--入札
            sSql.Append("    ,AnkenJizenDashinDate ");      //--事前打診登録日
            sSql.Append("    ,AnkenJizenDashinCheck ");     //--事前打診

            sSql.Append(" FROM AnkenJouhou WHERE AnkenJouhou.AnkenJouhouID = " + AnkenID);
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 業務評点担当者　赤伝／黒伝
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenNo"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝</param>
        /// <returns></returns>
        private int createGyoumuJouhouHyouronTantouL1(SqlCommand cmd, string ankenNo, int iType = 0)
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("INSERT INTO GyoumuJouhouHyouronTantouL1 ( ");
                                    sSql.Append("    GyoumuJouhouID ");
            sSql.Append("    ,HyouronTantouID ");
            sSql.Append("    ,HyouronTantoushaCD ");
            sSql.Append("    ,HyouronTantoushaMei ");
            sSql.Append("    ,HyouronnTantoushaHyouten ");
            sSql.Append(" ) SELECT ");
            sSql.Append("    " + ankenNo);
            sSql.Append("    ,HyouronTantouID ");
            sSql.Append("    ,HyouronTantoushaCD ");
            sSql.Append("    ,HyouronTantoushaMei ");
            sSql.Append("    ,HyouronnTantoushaHyouten ");
            sSql.Append(" FROM GyoumuJouhouHyouronTantouL1 WHERE GyoumuJouhouHyouronTantouL1.GyoumuJouhouID = " + AnkenID);

            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 窓口担当
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenNo"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝</param>
        /// <returns></returns>
        private int createGyoumuJouhouMadoguchi(SqlCommand cmd, string ankenNo, int iType = 0)
        {
            int result = 0;
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
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenNo"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝</param>
        /// <returns></returns>
        private int createGyoumuJouhouHyoutenBusho(SqlCommand cmd, string ankenNo, int iType = 0)
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
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// Ribc情報　赤伝／黒伝
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenNo"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝</param>
        /// <returns></returns>
        private int createRibcJouhou(SqlCommand cmd,string ankenNo,int iType = 0)
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("INSERT INTO RibcJouhou ( ");
            sSql.Append("    RibcID ");
            sSql.Append("    ,RibcNo ");
            sSql.Append("    ,RibcSeikyuKingaku ");

            sSql.Append("    ,RibcKoukiStart ");
            sSql.Append("    ,RibcKoukiEnd ");
            sSql.Append("    ,RibcSeikyubi ");
            sSql.Append("    ,RibcNouhinbi ");
            sSql.Append("    ,RibcNyukinyoteibi ");
            sSql.Append("    ,RibcUriageKeijyoTuki ");
            sSql.Append("    ,RibcKankeibusho ");
            sSql.Append("    ,RibcKubun ");
            sSql.Append("    ,RibcKankeibushoMei ");
            sSql.Append("     ) SELECT ");
            sSql.Append("     " + ankenNo );
            sSql.Append("    ,RibcNo ");
            sSql.Append("    ,").Append(iType == 0 || iType == 70 ? "-" : "").Append(" RibcSeikyuKingaku ");

            sSql.Append("    ,RibcKoukiStart ");
            sSql.Append("    ,RibcKoukiEnd " );
            sSql.Append("    ,RibcSeikyubi " );
            sSql.Append("    ,RibcNouhinbi ");
            sSql.Append("    ,RibcNyukinyoteibi ");
            sSql.Append("    ,RibcUriageKeijyoTuki ");
            sSql.Append("    ,RibcKankeibusho ");
            sSql.Append("    ,RibcKubun ");
            sSql.Append("    ,RibcKankeibushoMei ");
            sSql.Append(" FROM RibcJouhou WHERE RibcJouhou.RibcID = " + AnkenID);
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 応札情報　赤伝／黒伝
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenNo"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝</param>
        /// <returns></returns>
        private int createNyuusatsuJouhouOusatsusha(SqlCommand cmd, string ankenNo, int iType = 0)
        {
            StringBuilder sSql = new StringBuilder();
            sSql.Append("INSERT INTO NyuusatsuJouhouOusatsusha ( ");
            sSql.Append("    NyuusatsuJouhouID");
            sSql.Append("    , NyuusatsuOusatsuID");
            sSql.Append("    , NyuusatsuOusatsuKingaku");
            sSql.Append("    , NyuusatsuOusatsushaID");
            sSql.Append("    , NyuusatsuOusatsusha");
            sSql.Append("    , NyuusatsuOusatsuKyougouTashaID");
            sSql.Append("    , NyuusatsuOusatsuKyougouKigyouCD");
            sSql.Append("    , NyuusatsuRakusatsuJyuni");
            sSql.Append("    , NyuusatsuRakusatsuJokyou");
            sSql.Append("    , NyuusatsuRakusatsuComment");
            sSql.Append(" ) SELECT ");
            sSql.Append("    " + ankenNo);
            sSql.Append("    , ROW_NUMBER() OVER(ORDER BY NyuusatsuJouhouID) ");
            sSql.Append("    , ").Append(iType == 0 || iType == 70 ? "-" : "").Append(" NyuusatsuOusatsuKingaku");
            sSql.Append("    , NyuusatsuOusatsushaID");
            sSql.Append("    , NyuusatsuOusatsusha");
            sSql.Append("    , NyuusatsuOusatsuKyougouTashaID");
            sSql.Append("    , NyuusatsuOusatsuKyougouKigyouCD");
            sSql.Append("    , NyuusatsuRakusatsuJyuni");
            sSql.Append("    , NyuusatsuRakusatsuJokyou");
            sSql.Append("    , NyuusatsuRakusatsuComment");
            sSql.Append(" FROM NyuusatsuJouhouOusatsusha WHERE NyuusatsuJouhouOusatsusha.NyuusatsuJouhouID = " + AnkenID);
            cmd.CommandText = sSql.ToString();
            Console.WriteLine(cmd.CommandText);
            return cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// 業務配分　赤伝／黒伝
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="ankenNo"></param>
        /// <param name="iType">0:赤伝、1:黒伝、70:ダミーデータ赤伝、71:ダミーデータ黒伝</param>
        /// <returns></returns>
        private int createGyoumuHaibun(SqlCommand cmd, string ankenNo, DataTable GH_dt, int iType = 0)
        {
            int result = 0;
            string sStr = iType == 0 || iType == 70 ? "-" : "";
            if (GH_dt != null && GH_dt.Rows.Count > 0)
            {
                for (int i = 0; i < GH_dt.Rows.Count; i++)
                {
                    StringBuilder sSql = new StringBuilder();
                    sSql.Append("INSERT INTO GyoumuHaibun ( ");
                    sSql.Append("    GyoumuHaibunID ");
                    sSql.Append("    , GyoumuAnkenJouhouID ");
                    sSql.Append("    , GyoumuChosaBuRitsu ");
                    sSql.Append("    , GyoumuChosaBuGaku ");
                    sSql.Append("    , GyoumuJigyoFukyuBuRitsu ");
                    sSql.Append("    , GyoumuJigyoFukyuBuGaku ");
                    sSql.Append("    , GyoumuJyohouSystemBuRitsu ");
                    sSql.Append("    , GyoumuJyohouSystemBuGaku ");
                    sSql.Append("    , GyoumuSougouKenkyuJoRitsu ");
                    sSql.Append("    , GyoumuSougouKenkyuJoGaku ");
                    sSql.Append("    , GyoumuShizaiChousaRitsu ");
                    sSql.Append("    , GyoumuShizaiChousaGaku ");
                    sSql.Append("    , GyoumuEizenRitsu ");
                    sSql.Append("    , GyoumuEizenGaku ");
                    sSql.Append("    , GyoumuKikiruiChousaRitsu ");
                    sSql.Append("    , GyoumuKikiruiChousaGaku ");
                    sSql.Append("    , GyoumuKoujiChousahiRitsu ");
                    sSql.Append("    , GyoumuKoujiChousahiGaku ");
                    sSql.Append("    , GyoumuSanpaiFukusanbutsuRitsu ");
                    sSql.Append("    , GyoumuSanpaiFukusanbutsuGaku ");
                    sSql.Append("    , GyoumuHokakeChousaRitsu ");
                    sSql.Append("    , GyoumuHokakeChousaGaku ");
                    sSql.Append("    , GyoumuShokeihiChousaRitsu ");
                    sSql.Append("    , GyoumuShokeihiChousaGaku ");
                    sSql.Append("    , GyoumuGenkaBunsekiRitsu ");
                    sSql.Append("    , GyoumuGenkaBunsekiGaku ");
                    sSql.Append("    , GyoumuKijunsakuseiRitsu ");
                    sSql.Append("    , GyoumuKijunsakuseiGaku ");
                    sSql.Append("    , GyoumuKoukyouRoumuhiRitsu ");
                    sSql.Append("    , GyoumuKoukyouRoumuhiGaku ");
                    sSql.Append("    , GyoumuRoumuhiKoukyouigaiRitsu ");
                    sSql.Append("    , GyoumuRoumuhiKoukyouigaiGaku ");
                    sSql.Append("    , GyoumuSonotaChousabuRitsu ");
                    sSql.Append("    , GyoumuSonotaChousabuGaku ");
                    sSql.Append("    , GyoumuHibunKubun ");
                    sSql.Append(" ) SELECT ");
                    sSql.Append("    " + GlobalMethod.getSaiban("GyoumuHaibunID"));
                    sSql.Append("    , " + ankenNo);
                    sSql.Append("    , GyoumuChosaBuRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuChosaBuGaku ");
                    sSql.Append("    , GyoumuJigyoFukyuBuRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuJigyoFukyuBuGaku ");
                    sSql.Append("    , GyoumuJyohouSystemBuRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuJyohouSystemBuGaku ");
                    sSql.Append("    , GyoumuSougouKenkyuJoRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuSougouKenkyuJoGaku ");
                    sSql.Append("    , GyoumuShizaiChousaRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuShizaiChousaGaku ");
                    sSql.Append("    , GyoumuEizenRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuEizenGaku ");
                    sSql.Append("    , GyoumuKikiruiChousaRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuKikiruiChousaGaku ");
                    sSql.Append("    , GyoumuKoujiChousahiRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuKoujiChousahiGaku ");
                    sSql.Append("    , GyoumuSanpaiFukusanbutsuRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuSanpaiFukusanbutsuGaku ");
                    sSql.Append("    , GyoumuHokakeChousaRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuHokakeChousaGaku ");
                    sSql.Append("    , GyoumuShokeihiChousaRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuShokeihiChousaGaku ");
                    sSql.Append("    , GyoumuGenkaBunsekiRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuGenkaBunsekiGaku ");
                    sSql.Append("    , GyoumuKijunsakuseiRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuKijunsakuseiGaku ");
                    sSql.Append("    , GyoumuKoukyouRoumuhiRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuKoukyouRoumuhiGaku ");
                    sSql.Append("    , GyoumuRoumuhiKoukyouigaiRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuRoumuhiKoukyouigaiGaku ");
                    sSql.Append("    , GyoumuSonotaChousabuRitsu ");
                    sSql.Append("    , ").Append(sStr).Append(" GyoumuSonotaChousabuGaku ");
                    sSql.Append("    , GyoumuHibunKubun ");
                    sSql.Append(" FROM GyoumuHaibun WHERE GyoumuHaibun.GyoumuHaibunID = " + GetInt(GH_dt.Rows[i][1].ToString()));
                    cmd.CommandText = sSql.ToString();
                    Console.WriteLine(cmd.CommandText);
                    result = cmd.ExecuteNonQuery();
                }
            }
            return result;
        }
        #endregion

        #region DB反映処理共通メソッド ------------
        /// <summary>
        /// 配分か金額かDBへ更新できるフォマードへ変換する
        /// </summary>
        /// <param name="txt"></param>
        /// <returns></returns>
        private string getNumToDb(string txt)
        {
            return txt.Replace("%", "").Replace("¥", "").Replace(",", "");
        }
        #endregion

        #endregion
    }
}
