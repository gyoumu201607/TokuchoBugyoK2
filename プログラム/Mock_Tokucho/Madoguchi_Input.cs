using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using TokuchoBugyoK2;

namespace TokuchoBugyoK2
{
    public partial class Madoguchi_Input : Form
    {
        private String pgmName = "Madoguchi_Input";
        
        // 右クリックメニュー
        ContextMenuStrip contextMenuStrip1 = new ContextMenuStrip();
        ToolStripMenuItem item0 = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuBusho = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuBushoClear = new ToolStripMenuItem();
        ToolStripMenuItem contextMenuTantousha = new ToolStripMenuItem();       // 調査担当者の右メニュー
        ToolStripMenuItem contextMenuTantoushaBusho = new ToolStripMenuItem();  // 調査担当者の右メニュー
        //ToolStripMenuItem contextMenuSubTantousha = new ToolStripMenuItem();   // 調査担当者の部所を入れる
        //ToolStripMenuItem contextMenuSubTantousha2 = new ToolStripMenuItem();  // 選択された部所の担当者を入れる
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
        private string Nendo = "";
        private String sekouMode = "0"; // 施工条件 0:新規 1:更新 2:削除
        private string SekouJoukenID = "";
        private String openSekouTab = "0"; // 施工条件タブを開いているか 0:開いてない 1:開いている
        private String sekouMeijishoComboChangeFlg = "0"; // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
        private String sekouMeijishoIDChangeFlg = "0"; // 施工条件タブ 施工条件明示書ID変更フラグ 0:手動変更 1:システム側で変更
        private int pagelimit = 50;
        private int ChousaHinmokuMode = 0;//調査品目編集モード 0:表示 1:編集
        private Image Img_AddRow;
        private Image Img_AddRowNonactive;
        private Image Img_DeleteRow;
        private Image Img_DeleteRowNonactive;
        private Image Img_Sort;
        private int errorCnt = 0;
        private string ShukeiHyoFolder = "";
        private string chousaLinkFlg = "0"; // 調査品目明細のフォルダリンク先表示フラグ　1=非表示、0=表示
        //奉行エクセル
        private string sagyoForuda = "";
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
        // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件 パフォーマンス向上の為
        private string chousaHinmokuDispFlg = "0";
        // 調査品目明細のGridに全件読み込んだかどうかのフラグ 0:未 1:済
        private string chousaHinmokuLoadFlg = "0";
        private string jutakubangouChangeFlg = "0"; // 0:受託番号未変更 1:受託番号変更
        private string tabChousahinmokuFlg = "0"; // 0:調査品目明細を開いたことが無い 1:調査品目明細を開いたことがある
        private string tabChangeFlg = "0"; // 0:タブ移動してない 1:タブ移動した
        // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
        private string globalErrorFlg = "0"; // 0:正常 1:エラー

        public string mode = "";
        public string MadoguchiID = "";
        public string KakoIraiID = "";

        //VIPS 20200506 課題管理表No1314（1038）ADD 窓口ミハル　協力依頼書タブの協力先所属長を「調査統括部」の所属長でデフォルト登録されるようにする
        private const string CONST_DEFAULT_KYORYOKUSAKI_BUSHO_CD = "127100";

        // 不具合No1342
        //画面表示時の実施区分を保持するための変数
        private string originalJiishiKubun = "";

        //不具合No1207
        //共通マスタからグリッド行高の設定を取得する
        private string AutoSizeGridRowMode;
        private const string GRID_ROW_AUTO_SIZE = "行高自動調整";
        private const string GRID_ROW_FIX_SIZE = "行高自動調整解除";
        // 奉行エクセル移管対応
        private string IsPopup_ShukeiHyou_New = "0";

        public Madoguchi_Input()
        {
            InitializeComponent();
            //c1FlexGrid3.Rows[0].Height = 44;
            listener = new ClipboardListener(this);
            listener.DrawClipboard += Listener_DrawClipboard;

            //TODO デザイン画面にてボタンが表示域外の為ソースでイベント付与
            //this.button13.Click += button13_Click;

            // 調査概要タブの売上年度にマウスホイールイベントを付与
            this.item1_MadoguchiTourokuNendo.MouseWheel += item_MouseWheel;
            // 調査概要
            this.item1_PrintList.MouseWheel += item_MouseWheel;                  // 連絡票出力
            this.item1_MadoguchiJutakuBushoCD.MouseWheel += item_MouseWheel;     // 受託課所支部
            this.item1_MadoguchiTantoushaBushoCD.MouseWheel += item_MouseWheel;  // 窓口部所
            this.item1_AnkenGyoumuKubun.MouseWheel += item_MouseWheel;           // 契約区分
            this.item1_MadoguchiChousaShubetsu.MouseWheel += item_MouseWheel;    // 調査種別
            this.item1_MadoguchiJiishiKubun.MouseWheel += item_MouseWheel;       // 実施区分
            // 調査品目明細
            this.item4_IraiKubun.MouseWheel += item_MouseWheel;                  // 依頼区分
            this.src_Busho.MouseWheel += item_MouseWheel;                        // 調査担当部所
            this.src_ShuFuku.MouseWheel += item_MouseWheel;                      // 主＋副
            this.src_Zaikou.MouseWheel += item_MouseWheel;                       // 材工
            this.src_TantoushaKuuhaku.MouseWheel += item_MouseWheel;             // 調査担当者空白リスト
            this.item_Hyoujikensuu.MouseWheel += item_MouseWheel;                // 表示件数
            // 協力依頼書
            this.item4_KyoRyokuBusho.MouseWheel += item_MouseWheel;              // 協力先部所
            this.item4_GyoumuKubun.MouseWheel += item_MouseWheel;                // 業務区分
            this.item4_Zumen.MouseWheel += item_MouseWheel;                      // 図面
            this.item4_Kizyunbi.MouseWheel += item_MouseWheel;                   // 調査基準日
            this.item4_UtiawaseYouhi.MouseWheel += item_MouseWheel;              // 打合せ要否
            this.item4_Gutaiteki.MouseWheel += item_MouseWheel;                  // 具体的な
            this.item4_ZenkaiKyouryoku.MouseWheel += item_MouseWheel;            // 前回協力
            this.item4_Hikiwatashi.MouseWheel += item_MouseWheel;                // 成果物引渡場所
            this.item4_JishiKeikakusho.MouseWheel += item_MouseWheel;            // 実施計画書
            this.item4_MitsumoriChousyu.MouseWheel += item_MouseWheel;           // 見積徴収
            // 単品入力項目
            this.item6_TanpinShijisho.MouseWheel += item_MouseWheel;             // 指示書
            // 施工条件
            this.item7_MeijishoKirikaeCombo.MouseWheel += item_MouseWheel;                       // 明示書切り替え

            this.c1FlexGrid4.MouseWheel += c1FlexGrid4_MouseWheel; // 調査品目明細のGrid

            //エントリ君修正STEP2
            this.ErrorMessage.Font = new System.Drawing.Font(this.ErrorMessage.Font.Name, float.Parse(GlobalMethod.GetCommonValue1("DSP_ERROR_FONTSIZE")));

        }

        ClipboardListener listener;

        public bool isFormCopy = false;
        private void Listener_DrawClipboard(object sender, EventArgs e)
        {
            Console.WriteLine(" Listener_DrawClipboard ======================================= STT");
            Console.WriteLine(isFormCopy);
            if (copyData != null)
                Console.WriteLine(copyData.Count);
            if (!isFormCopy && copyData != null)
            {
                copyData.Clear();
            }
            // クリップボードのデータが変化した。
            if (copyData != null)
                Console.WriteLine(copyData.Count);
            isFormCopy = false;
            Console.WriteLine(" Listener_DrawClipboard ======================================= END");
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

        private void Madoguchi_Input_Load(object sender, EventArgs e)
        {
            //不具合No1017（751）
            //タブの文字装飾変更対応
            //文字表示を大きくする場合は、デザイナでTabのItemSize.widthを変更する。窓口、特命課長、自分大臣は、125で設定すると、14ポイントぐらいのサイズでいける
            tab.DrawMode = TabDrawMode.OwnerDrawFixed;

            // 奉行エクセル移管対応
            if (ConfigurationManager.AppSettings.AllKeys.Contains("Popup_ShukeiHyou_New"))
            {
                IsPopup_ShukeiHyou_New = ConfigurationManager.AppSettings["Popup_ShukeiHyou_New"].ToString();
            }

            //不具合No1207
            //共通マスタからグリッド行高の設定を取得する
            AutoSizeGridRowMode = GlobalMethod.GetCommonValue1("CHOUSA_GYOU_FLG");

            //不具合No1355（1123）
            lblVersion.Text = GlobalMethod.GetCommonValue1("APL_VERSION");
            if (GlobalMethod.GetCommonValue1("BOOT_MODE") == "1")
            {
                lblBootMode.Text = GlobalMethod.GetCommonValue2("BOOT_MODE");
            }
            GlobalMethod.outputLogger("Madoguchi", "Madoguchi_Input Load開始:" + DateTime.Now.ToString(), "", "DEBUG");
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
            c1FlexGrid2.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid2.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid4.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid4.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            c1FlexGrid5.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            c1FlexGrid5.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");
            BikoGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Ascending] = Image.FromFile("Resource/Asc.png");
            BikoGrid.Glyphs[C1.Win.C1FlexGrid.GlyphEnum.Descending] = Image.FromFile("Resource/Desc.png");

            //// 行番号の取得
            //int ShinchokuIconColIndex = c1FlexGrid4.Cols["ShinchokuIcon"].Index;                            // 進捗状況
            //int ChousaZentaiJunColIndex = c1FlexGrid4.Cols["ChousaZentaiJun"].Index;                        // 全体順
            //int ChousaKobetsuJunColIndex = c1FlexGrid4.Cols["ChousaKobetsuJun"].Index;                      // 個別順
            //int ChousaZaiKouColIndex = c1FlexGrid4.Cols["ChousaZaiKou"].Index;                              // 材工
            //int ChousaHinmeiColIndex = c1FlexGrid4.Cols["ChousaHinmei"].Index;                              // 品目
            //int ChousaKikakuColIndex = c1FlexGrid4.Cols["ChousaKikaku"].Index;                              // 規格
            //int ChousaTankaColIndex = c1FlexGrid4.Cols["ChousaTanka"].Index;                                // 単位
            //int ChousaSankouShitsuryouColIndex = c1FlexGrid4.Cols["ChousaSankouShitsuryou"].Index;          // 参考質量
            //int ChousaKakakuColIndex = c1FlexGrid4.Cols["ChousaKakaku"].Index;                              // 価格
            //int ChousaChuushiColIndex = c1FlexGrid4.Cols["ChousaChuushi"].Index;                            // 中止
            //int ChousaBikou2ColIndex = c1FlexGrid4.Cols["ChousaBikou2"].Index;                              // 報告備考
            //int ChousaBikouColIndex = c1FlexGrid4.Cols["ChousaBikou"].Index;                                // 依頼備考
            //int ChousaTankaTekiyouTikuColIndex = c1FlexGrid4.Cols["ChousaTankaTekiyouTiku"].Index;          // 単価適用地域
            //int ChousaZumenNoColIndex = c1FlexGrid4.Cols["ChousaZumenNo"].Index;                            // 図面番号
            //int ChousaSuuryouColIndex = c1FlexGrid4.Cols["ChousaSuuryou"].Index;                            // 数量
            //int ChousaMitsumorisakiColIndex = c1FlexGrid4.Cols["ChousaMitsumorisaki"].Index;                // 見積先
            //int ChousaBaseMakereColIndex = c1FlexGrid4.Cols["ChousaBaseMakere"].Index;                      // ベースメーカー
            //int ChousaBaseTankaColIndex = c1FlexGrid4.Cols["ChousaBaseTanka"].Index;                        // ベース単位
            //int ChousaKakeritsuColIndex = c1FlexGrid4.Cols["ChousaKakeritsu"].Index;                        // 掛率
            //int ChousaObiMeiColIndex = c1FlexGrid4.Cols["ChousaObiMei"].Index;                              // 属性
            //int ChousaZenkaiTaniColIndex = c1FlexGrid4.Cols["ChousaZenkaiTani"].Index;                      // 前回単位
            //int ChousaZenkaiKakakuColIndex = c1FlexGrid4.Cols["ChousaZenkaiKakaku"].Index;                  // 前回価格
            //int ChousaSankoutiColIndex = c1FlexGrid4.Cols["ChousaSankouti"].Index;                          // 発注者提供単価
            //int ChousaHinmokuJouhou1ColIndex = c1FlexGrid4.Cols["ChousaHinmokuJouhou1"].Index;              // 品目情報1
            //int ChousaHinmokuJouhou2ColIndex = c1FlexGrid4.Cols["ChousaHinmokuJouhou2"].Index;              // 品目情報2
            //int ChousaFukuShizaiColIndex = c1FlexGrid4.Cols["ChousaFukuShizai"].Index;                      // 前回質量
            //int ChousaBunruiColIndex = c1FlexGrid4.Cols["ChousaBunrui"].Index;                              // メモ1
            //int ChousaMemo2ColIndex = c1FlexGrid4.Cols["ChousaMemo2"].Index;                                // メモ2
            //int ChousaTankaCD1ColIndex = c1FlexGrid4.Cols["ChousaTankaCD1"].Index;                          // 発注品目コード
            //int ChousaTikuWariCodeColIndex = c1FlexGrid4.Cols["ChousaTikuWariCode"].Index;                  // 地区割コード
            //int ChousaTikuCodeColIndex = c1FlexGrid4.Cols["ChousaTikuCode"].Index;                          // 地区コード
            //int ChousaTikuMeiColIndex = c1FlexGrid4.Cols["ChousaTikuMei"].Index;                            // 地区名
            //int ChousaShougakuColIndex = c1FlexGrid4.Cols["ChousaShougaku"].Index;                          // 少額案件[10万/100万]
            //int ChousaWebKenColIndex = c1FlexGrid4.Cols["ChousaWebKen"].Index;                              // Web建
            //int ChousaKonkyoCodeColIndex = c1FlexGrid4.Cols["ChousaKonkyoCode"].Index;                      // 根拠関連コード
            //int HinmokuRyakuBushoCDColIndex = c1FlexGrid4.Cols["HinmokuRyakuBushoCD"].Index;                // 調査担当部所
            //int HinmokuChousainCDColIndex = c1FlexGrid4.Cols["HinmokuChousainCD"].Index;                    // 調査担当者
            //int HinmokuRyakuBushoFuku1CDColIndex = c1FlexGrid4.Cols["HinmokuRyakuBushoFuku1CD"].Index;      // 副調査担当部所1
            //int HinmokuFukuChousainCD1ColIndex = c1FlexGrid4.Cols["HinmokuFukuChousainCD1"].Index;          // 副調査担当者1
            //int HinmokuRyakuBushoFuku2CDColIndex = c1FlexGrid4.Cols["HinmokuRyakuBushoFuku2CD"].Index;      // 副調査担当部所2
            //int HinmokuFukuChousainCD2ColIndex = c1FlexGrid4.Cols["HinmokuFukuChousainCD2"].Index;          // 副調査担当者2
            //int ChousaHoukokuHonsuuColIndex = c1FlexGrid4.Cols["ChousaHoukokuHonsuu"].Index;                // 報告数
            //int ChousaHoukokuRankColIndex = c1FlexGrid4.Cols["ChousaHoukokuRank"].Index;                    // 報告ランク
            //int ChousaIraiHonsuuColIndex = c1FlexGrid4.Cols["ChousaIraiHonsuu"].Index;                      // 依頼数
            //int ChousaIraiRankColIndex = c1FlexGrid4.Cols["ChousaIraiRank"].Index;                          // 依頼ランク
            //int ChousaHinmokuShimekiribiColIndex = c1FlexGrid4.Cols["ChousaHinmokuShimekiribi"].Index;      // 締切日
            //int ChousaHoukokuzumiColIndex = c1FlexGrid4.Cols["ChousaHoukokuzumi"].Index;                    // 報告済

            ////ソート項目にアイコンを設定
            //C1.Win.C1FlexGrid.CellRange cr;
            //Bitmap bmp1 = new Bitmap("Resource/Image/SortIconDefalt.png");
            //Bitmap bmpSort = new Bitmap(bmp1, bmp1.Width / 6, bmp1.Height / 6);

            //cr = c1FlexGrid4.GetCellRange(0, ShinchokuIconColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaZentaiJunColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaKobetsuJunColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaZaiKouColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaHinmeiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaKikakuColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaTankaColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaSankouShitsuryouColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaKakakuColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaChuushiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaBikou2ColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaBikouColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaTankaTekiyouTikuColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaZumenNoColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaSuuryouColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaMitsumorisakiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaBaseMakereColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaBaseTankaColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaKakeritsuColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaObiMeiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaZenkaiTaniColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaZenkaiKakakuColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaSankoutiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaHinmokuJouhou1ColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaHinmokuJouhou2ColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaFukuShizaiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaBunruiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaMemo2ColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaTankaCD1ColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaTikuWariCodeColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaTikuCodeColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaTikuMeiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaShougakuColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaWebKenColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaKonkyoCodeColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, HinmokuRyakuBushoCDColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, HinmokuChousainCDColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, HinmokuRyakuBushoFuku1CDColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, HinmokuFukuChousainCD1ColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, HinmokuRyakuBushoFuku2CDColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, HinmokuFukuChousainCD2ColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaHoukokuHonsuuColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaHoukokuRankColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaIraiHonsuuColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaIraiRankColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaHinmokuShimekiribiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;
            //cr = c1FlexGrid4.GetCellRange(0, ChousaHoukokuzumiColIndex);
            //cr.StyleNew.ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.RightCenter;
            //cr.Image = bmpSort;



            item3_TargetPage.ImeMode = ImeMode.Disable;

            // c1FlexGridを隠す
            // 担当部所
            c1FlexGrid1.Visible = false;
            c1FlexGrid5.Visible = false;
            // 調査品目明細
            c1FlexGrid4.Visible = false;
            // 単品入力項目
            c1FlexGrid2.Visible = false;
            // 備考
            BikoGrid.Visible = false;


            // 単品入力の電話、FAX、メールのIMEを設定
            item6_TanpinTel.ImeMode = ImeMode.Disable;
            item6_TanpinFax.ImeMode = ImeMode.Disable;
            item6_TanpinMail.ImeMode = ImeMode.Disable;

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
            //奉行エクセル
            //作業フォルダの幅を調整
            c1FlexGrid4.Cols["SagyoForudaPath"].Width = 0;

            //モード別処理
            if (mode == "insert")
            {
                //不要タブの非表示化
                this.tab.TabPages.Remove(this.tabPage2);
                this.tab.TabPages.Remove(this.tabPage3);
                this.tab.TabPages.Remove(this.tabPage4);
                this.tab.TabPages.Remove(this.tabPage5);
                this.tab.TabPages.Remove(this.tabPage6);
                this.tab.TabPages.Remove(this.tabPage7);
                this.tab.TabPages.Remove(this.tabPage8);

                //新規時のみ非表示項目
                button9.Visible = false;
                label47.Visible = false;
                item1_MadoguchiHoukokuzumi.Visible = false;
                button10.Visible = false;
                label48.Visible = false;
                item1_PrintList.Visible = false;
                button11.Visible = false;

                //更新ボタン名の表示を登録にする
                button16.Text = "登録";

                // 新規時は削除ボタン非表示
                btnDelete.Visible = false;

                // Garoon連携の文言とボタンを非表示
                item1_GaroonUpdateDispTitle.Visible = false;
                item1_GaroonUpdateDisp.Visible = false;
                item1_GaroonRenkeiBtn.Visible = false;

            }
            if (mode == "update")
            {
                // 受託番号を検索ボタンを非活性に
                button1.Enabled = false;

                // 協力依頼書タブの応援依頼メール
                //tableLayoutPanel25.Visible = false;

                // 特調番号の枝番は更新時の初期時は編集不可
                //item1_MadoguchiUketsukeBangouEdaban.Enabled = false;
                item1_MadoguchiUketsukeBangouEdaban.ReadOnly = true;

            }
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
                if (item1_MadoguchiTourokuNendo.Text == "")
                {
                    Nendo = DateTime.Today.Year;
                    ToNendo = DateTime.Today.AddYears(1).Year;
                }
                else
                {
                    int.TryParse(item1_MadoguchiTourokuNendo.SelectedValue.ToString(), out Nendo);
                    ToNendo = Nendo + 1;
                }
                cmd.CommandText = "SELECT " +
                "GyoumuBushoCD  " +
                ",BushokanriboKameiRaku  " +
                "FROM Mst_Busho  " +
                "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                ////" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                "ORDER BY BushoMadoguchiNarabijun";

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

            //調査員
            //DataTable dt2 = new DataTable();
            //using (var conn = new SqlConnection(connStr))
            //{
            //    var cmd = conn.CreateCommand();
            //    cmd.CommandText = "SELECT " +
            //        "KojinCD " +
            //        ",ChousainMei " +
            //        "FROM Mst_Chousain " +
            //        "WHERE RetireFLG = 0 AND TokuchoFLG = 1 AND TokuchoRole > 0 " +
            //        "ORDER BY ChousainMei ";

            //    var sda = new SqlDataAdapter(cmd);
            //    dt2.Clear();
            //    sda.Fill(dt2);
            //    conn.Close();
            //}
            //contextMenuTantousha.Text = "調査担当者";
            //contextMenuTantousha = Set_ContextMenu(contextMenuTantousha, dt2);

            //ToolStripMenuItem contextMenuSubTantousha = new ToolStripMenuItem();

            //for (int i = 0; i < dt2.Rows.Count; i++)
            //{
            //    if (dt2.Rows[i][1].ToString() != "")
            //    {
            //        contextMenuSubTantousha.DropDownItems.Add(dt.Rows[i][1].ToString(), null, ContextMenuEvent);
            //    }
            //}
            //contextMenuTantousha.DropDownItemClicked += contextMenuTantoushaItemClicked;

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
            //        //cmd.CommandText = "SELECT " +
            //        //    "tkr.TankaRankHinmoku AS Value " +
            //        //    ",tkr.TankaRankHinmoku AS Descript " +
            //        //    "FROM TankaKeiyakuRank tkr " +
            //        //    "INNER JOIN TankaKeiyaku tk ON tk.TankaKeiyakuID = tkr.TankaKeiyakuID " +
            //        //    "INNER JOIN AnkenJouhou aj ON aj.AnkenJouhouID = tk.AnkenJouhouID " +
            //        //    "INNER JOIN MadoguchiJouhou mj ON mj.AnkenJouhouID = aj.AnkenJouhouID " +
            //        //    "WHERE tkr.TankaRankDeleteFlag != 1 " +
            //        //    "AND mj.MadoguchiID = '" + MadoguchiID + "' " +
            //        //    "ORDER BY tk.TankaKeiyakuID,tkr.TankaRankID";
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
            //        //cmd.CommandText = "SELECT " +
            //        //    "tkr.TankaRankHinmoku AS Value " +
            //        //    ",tkr.TankaRankHinmoku AS Descript " +
            //        //    "FROM TankaKeiyakuRank tkr " +
            //        //    "INNER JOIN TankaKeiyaku tk ON tk.TankaKeiyakuID = tkr.TankaKeiyakuID " +
            //        //    "INNER JOIN AnkenJouhou aj ON aj.AnkenJouhouID = tk.AnkenJouhouID " +
            //        //    "INNER JOIN MadoguchiJouhou mj ON mj.AnkenJouhouID = aj.AnkenJouhouID " +
            //        //    "WHERE tkr.TankaRankDeleteFlag != 1 " +
            //        //    "AND mj.MadoguchiID = '" + MadoguchiID + "' " +
            //        //    "ORDER BY tk.TankaKeiyakuID,tkr.TankaRankID";
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
            //    //contextMenuIrai.DropDownItems.Add("A-①", null, ContextMenuEvent);
            //    //contextMenuIrai.DropDownItems.Add("A-②", null, ContextMenuEvent);

            //    contextMenuIrai = Set_ContextMenu(contextMenuIrai, iraiDt);

            //}


            //実施区分
            DataTable tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "実施");
            tmpdt.Rows.Add(2, "打診中");
            tmpdt.Rows.Add(3, "中止");

            item1_MadoguchiJiishiKubun.DataSource = tmpdt;
            item1_MadoguchiJiishiKubun.DisplayMember = "Discript";
            item1_MadoguchiJiishiKubun.ValueMember = "Value";

            //調査種別
            DataTable tmpdt2 = new System.Data.DataTable();
            tmpdt2.Columns.Add("Value", typeof(int));
            tmpdt2.Columns.Add("Discript", typeof(string));
            tmpdt2.Rows.Add(1, "単品");
            tmpdt2.Rows.Add(2, "一般");
            tmpdt2.Rows.Add(3, "単契");

            item1_MadoguchiChousaShubetsu.DataSource = tmpdt2;
            item1_MadoguchiChousaShubetsu.DisplayMember = "Discript";
            item1_MadoguchiChousaShubetsu.ValueMember = "Value";

            // 管理帳票印刷コンボボックス
            string discript = "PrintName";
            string value = "PrintListID";
            string table = "Mst_PrintList";
            //where = "";
            string where = "MENU_ID = 201 AND PrintBunruiCD = 2 AND PrintDelFlg <> 1 ORDER BY PrintListNarabijun";
            //コンボボックスデータ取得
            DataTable combodt = GlobalMethod.getData(discript, value, table, where);
            //DataRow dr = combodt.NewRow();
            //combodt.Rows.InsertAt(dr, 0);
            item1_PrintList.DataSource = combodt;
            item1_PrintList.DisplayMember = "Discript";
            item1_PrintList.ValueMember = "Value";

            //コンボボックス取得
            get_combo();

            // ヘッダーのボタン下のGaroon連携設定日時の文言取得（変更出来るように）
            string garoonUpdateDisp = GlobalMethod.GetCommonValue1("GAROON_UPDATETIME_DISP");
            if (garoonUpdateDisp != null && garoonUpdateDisp == "")
            {
                item1_GaroonUpdateDispTitle.Text = garoonUpdateDisp;
            }

            //新規登録以外だったらデータ取得
            if (MadoguchiID != "")
            {
                get_data(1);
                // 不具合No1342
                //画面表示時の実施区分を保持する。これをもとに実施区分のテキストチェンジイベント時に、報告完了を元に戻すか判定する。
                originalJiishiKubun=item1_MadoguchiJiishiKubun.SelectedValue.ToString();
                //Console.WriteLine("登録済み実施区分：" + originalJiishiKubun);

                get_combo_byNendo();
                //get_data(2);
                //get_data(3); //調査品目　検索条件のコンボが正常に認識できないため、別タイミングに移動
                get_data(4);
                get_data(5); // 応援受付
                get_data(6);
                get_data(7);

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

                // ヘッダー表示
                // 特調番号
                Header1.Text = item1_MadoguchiUketsukeBangou.Text + "-" + item1_MadoguchiUketsukeBangouEdaban.Text;
                // 発注者名・課名
                Header3.Text = item1_MadoguchiHachuuKikanmei.Text;
                // 業務名称
                Header4.Text = item1_MadoguchiGyoumuMeishou.Text;
            }
            else
            {
                //ファイルパス初期用セット
                folderDefault_set();
            }
            FolderPathCheck();


            GlobalMethod.outputLogger("Madoguchi", "Madoguchi_Input Load終了:" + DateTime.Now.ToString(), "", "DEBUG");
        }

        private void folderDefault_set()
        {

            //フォルダ初期値設定
            if (MadoguchiID == "")
            {
                string discript = "FolderPath";
                string value = "FolderPath ";
                string table = "M_Folder";
                string where = "MENU_ID = 100 AND FolderBunruiCD = 1 AND FolderBushoCD = '" + UserInfos[2] + "' ";

                // //xxxx/00Cyousa/00調査情報部門共有/$NENDO$/200受託調査関連
                string FolderBase = GlobalMethod.GetCommonValue1("FOLDER_BASE").Replace(@"$NENDO$", item1_MadoguchiTourokuNendo.SelectedValue.ToString());
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

                //集計表フォルダ
                item1_MadoguchiShukeiHyoFolder.Text = FolderPath;
                //報告書フォルダ
                item1_MadoguchiHoukokuShoFolder.Text = FolderPath;
                //調査資料フォルダ
                item1_MadoguchiShiryouHolder.Text = FolderPath;

            }
        }

        String ankenJouhouId = "";
        //データ取得
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
                            ",mc1.ChousainMei " +//窓口担当者名
                            ",MadoguchiTantoushaCD " +//窓口担当者CD
                            ",MadoguchiTourokuNendo " +//登録年度
                            //",KanriGijutsushaNM " +//管理技術者 9
                            ",mc2.ChousainMei " +//管理技術者 9
                                                   //",MadoguchiJutakuBangou " +//受託番号 10
                            ",CASE MadoguchiJutakuBangouEdaban WHEN ''  THEN MadoguchiJutakuBangou ELSE MadoguchiJutakuBangou + '-' + MadoguchiJutakuBangouEdaban END AS JutakuBangou " +//受託番号 10
                            ",MadoguchiJutakuBangouEdaban " +//受託番号枝番
                            //",GyoumuKanrishaMei " +//業務管理者　　CD
                            ",mc3.ChousainMei " +//業務管理者　　CD
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
                            ",MadoguchiHoukokuzumi " + // 41:報告済み
                                                       //",MadoguchiAnkenJouhouID " +//案件情報ID
                            ",MadoguchiJouhou.AnkenJouhouID " +//案件情報ID
                            ",MadoguchiJutakuTantoushaID " +//契約担当者CD
                            ",MadoguchiKanriGijutsusha " +//管理技術者CD
                            ",MadoguchiHonbuTanpinflg " +//本部単品
                            ",MadoguchiGaroonRenkei " +//Garoon連携
                            "FROM MadoguchiJouhou " +
                            //"LEFT JOIN AnkenJouhou ON AnkenJutakuBangou = replace(MadoguchiJutakuBangou,'-' + MadoguchiJutakuBangouEdaban,'') " +
                            //"AND MadoguchiJutakuBangouEdaban = AnkenJutakuBangouEda " +7

                            // 窓口ミハルの受託番号は、枝番と完全に分離しているので「-」で連結、また契約情報IDも条件に含める
                            "LEFT JOIN AnkenJouhou ON AnkenJouhou.AnkenJouhouID = MadoguchiJouhou.AnkenJouhouID " +
                            //"      AND AnkenJutakuBangou = MadoguchiJutakuBangou + '-' + MadoguchiJutakuBangouEdaban " +
                            "      AND AnkenJutakuBangou = CASE MadoguchiJutakuBangouEdaban WHEN ''  THEN MadoguchiJutakuBangou ELSE MadoguchiJutakuBangou + '-' + MadoguchiJutakuBangouEdaban END " +
                            //"      AND AnkenJutakuBangouEda = MadoguchiJutakuBangouEdaban " +
                            "LEFT JOIN Mst_Chousain mc1 ON  MadoguchiTantoushaCD = mc1.KojinCD " +
                            //"LEFT JOIN Mst_Busho mb1 ON JutakuBushoShozokuCD = mb1.GyoumuBushoCD " +
                            "LEFT JOIN Mst_Busho mb1 ON MadoguchiJutakuBushoCD = mb1.GyoumuBushoCD " +
                            "LEFT JOIN Mst_Busho mb2 ON MadoguchiBushoShozokuCD = mb2.GyoumuBushoCD " +
                            "LEFT JOIN GyoumuJouhou gj ON AnkenJouhou.AnkenJouhouID = gj.AnkenJouhouID " +
                            "AND MadoguchiKanriGijutsusha = gj.KanriGijutsushaCD " +
                            "LEFT JOIN Mst_Chousain mc2 ON ISNULL(gj.KanriGijutsushaCD,'0') = mc2.KojinCD " +
                            "LEFT JOIN Mst_Chousain mc3 ON ISNULL(GyoumuKanrishaCD,'0') = mc3.KojinCD " +
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
                            "WHERE MadoguchiJouhou.MadoguchiID = '" + MadoguchiID + "'   ";

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
                    // 調査品目明細タブ
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
                            if (buf == "1") {
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
                                   " , ChousaKonkyoCode "+
                                   " , ChousaTaniAtariKakaku"+
                                   //奉行エクセル　
                                   ", ChousaShuukeihyouVer"+
                                   ", ChousaBunkatsuHouhou"+
                                   ", ChousaKoujiKouzoubutsumei"+
                                   ", ChousaHachushaTeikyouTani"+
                                   ", chousaTaniAtariSuuryou"+
                                   ", ChousaTaniAtariTanka"+
                                   ", ChousaNiwatashiJouken"+
                                   " , ChousaMadoguchiGroupMasterID"
                                   ; // 35
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
                                   " , ISNULL(MC0.RetireFlg, 0) AS RetireFlg " + //担当者退職フラグ
                                   " , ISNULL(MC1.RetireFlg, 0) AS RetireFlg1 " + //副担当者1退職フラグ
                                   " , ISNULL(MC2.RetireFlg, 0) AS RetireFlg2 " + //副担当者2退職フラグ
                                   " , MC0.ChousainMei " + // 調査員名
                                   " , MC1.ChousainMei AS FukuChousainMei1 " + // 副調査員名1
                                   " , MC2.ChousainMei AS FukuChousainMei2 "  // 副調査員名2
                                   ;
                                    // 作業フォルダアイコン切り替え 0:グレー 1:イエロー
                                    if (chousaLinkFlg != "1")
                                    {
                                        cmd.CommandText += " , CASE WHEN ISNULL(MJMC.MadoguchiL1SagyouHolder,'') <> '' THEN 1 ELSE 0 END AS Sagyou ";
                                    }
                                    else
                                    {
                                        cmd.CommandText += " , 0 AS Sagyou ";
                                    }
                                    cmd.CommandText += " , MJMC.MadoguchiL1SagyouHolder AS SagyouHolder " + //担当者作業フォルダ
                                   "FROM " +
                                   " ChousaHinmoku  " +
                                   "LEFT JOIN MadoguchiJouhou ON MadoguchiJouhou.MadoguchiID = ChousaHinmoku.MadoguchiID " +
                                   "LEFT JOIN Mst_Chousain MC0 ON HinmokuChousainCD = MC0.KojinCD " +
                                   "LEFT JOIN Mst_Chousain MC1 ON HinmokuFukuChousainCD1 = MC1.KojinCD " +
                                   "LEFT JOIN Mst_Chousain MC2 ON HinmokuFukuChousainCD2 = MC2.KojinCD " +
                                   //奉行エクセル
                                   "LEFT JOIN MadoguchiJouhouMadoguchiL1Chou MJMC ON ChousaHinmoku.MadoguchiID = MJMC.MadoguchiID " +
                                   "AND ChousaHinmoku.HinmokuChousainCD = MJMC.MadoguchiL1ChousaTantoushaCD " +
                                   //"LEFT JOIN MadoguchiGroupMaster ON MadoguchiGroupMaster.MadoguchiGroupMasterID = ChousaHinmoku.ChousaMadoguchiGroupMasterID " +
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
                                cmd.CommandText += " AND (MC0.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' OR MC1.ChousainMei COLLATE Japanese_XJIS_100_CI_AS_SC LIKE N'%" + Chousain + "%' ESCAPE'\\' OR MC2.ChousainMei LIKE N'%" + Chousain + "%' ESCAPE'\\' ) ";
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
                    // 協力依頼書
                    else if (tab == 4)
                    {
                        // 初期表示と、過去の依頼書を参照に登録ボタンで、検索するKeyを分ける
                        string searchID = "";
                        if (KakoIraiID != "")
                        {
                            searchID = KakoIraiID;
                            // クリアタイミングを後ろにずらす（依頼書参照後の更新ボタン1度目エラー対応）
                            //KakoIraiID = "";
                        }
                        else
                        {
                            searchID = MadoguchiID;
                        }

                        //協力依頼書の取得
                        cmd.CommandText = "SELECT " +
                        " KyourokuIraisakiBushoOld " + // 0:協力先部所
                        ",KyourokuIraisakiTantoshaCD " +
                        ",CASE KyouryokuDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KyouryokuDate,'yyyy/MM/dd') END " +
                        //",KyouryokuOuenirai " +
                        //" KyouryokuIraisakiBusho " +
                        ",CASE KyouryokuHoukokuSeigenDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(KyouryokuHoukokuSeigenDate,'yyyy/MM/dd') END " +
                        //",KyourokuIraisakiTantoshaCD " +
                        //",KyouryokuDate " +
                        //",KyouryokuHoukokuSeigenDate " +
                        ////",KyouryokuOuenirai " +
                        ",KyouryokuGyoumuKubun " +

                        ",KyouryokuIraiKubun " + // 5:依頼区分
                        ",KyouryokuNaiyoukubunShizai " +
                        ",KyouryokuNaiyoukubunDKou " +
                        ",KyouryokuNaiyoukubunEKou " +
                        ",KyouryokuNaiyoukubunSonota " +
                        ",KyouryokuNaiyoukubunJohokaihatsu " + // 10:
                        ",KyouryokuSonota " +
                        ",KyouryokuGyoumuNaiyou " +
                        ",KyouryokuZumen " +

                        ",KyouryokuChousaKijun " +
                        ",KyouryokuChousakijunbi " + // 15:
                        ",KyouryokuUtiawaseyouhi " +
                        ",KyouryokuGutaiteki " +
                        ",KyouryokuZenkaiUmu " +

                        ",KyouryokuZenkaiUmubi " +
                        ",KyouryokusakiHikiwatashi " + // 20:
                        ",KyouryokuJisshikeikakusho " +
                        ",KyouryokuChoushuusaki " +
                        //",KyouryokuRenrakuJikou " +
                        //",KyouryokuIraishoHozonsaki " +
                        //",KyouryokuUpdateDate " +
                        //",KyouryokuUpdateUser " +
                        //",KyouryokuUpdateProgram " +
                        //",KyouryokuDeleteFlag " +
                        "";

                        if (KakoIraiID == "")
                        {
                            cmd.CommandText += ",KyouryokuIraishoID ";
                        }
                        else
                        {
                            cmd.CommandText += ", '" + item4_KyouryokuIraishoID.Text + "' AS KyouryokuIraishoID ";
                        }
                        cmd.CommandText +=
                        ",ChousainMei " + // 24
                        //",mb.BushoShozokuChou " + // 24

                        ",(select count(ChousaZaiKou) from ChousaHinmoku Where ChousaZaiKou = 1 and MadoguchiID = '" + MadoguchiID + "') as Zai " + // 25:
                        ",(select count(ChousaZaiKou) from ChousaHinmoku Where ChousaZaiKou = 2 and MadoguchiID = '" + MadoguchiID + "') as DKou " +
                        ",(select count(ChousaZaiKou) from ChousaHinmoku Where ChousaZaiKou = 3 and MadoguchiID = '" + MadoguchiID + "') as EKou " +
                        ",(select count(ChousaZaiKou) from ChousaHinmoku Where ChousaZaiKou = 4 and MadoguchiID = '" + MadoguchiID + "') as DSonota " +

                        ",aj.AnkenGyoumuKubunMei " +

                        ",gk.GyoumuNarabijunCD " +

                        "FROM KyouryokuIraisho ki " +
                        "LEFT JOIN Mst_Chousain mc ON ki.KyourokuIraisakiTantoshaCD = mc.KojinCD " +
                        "LEFT JOIN MadoguchiJouhou mj ON mj.MadoguchiID = ki.MadoguchiID " +
                        "LEFT JOIN AnkenJouhou aj ON mj.AnkenJouhouID = aj.AnkenJouhouID " +
                        "LEFT JOIN Mst_GyoumuKubun gk ON gk.GyoumuKubun = aj.AnkenGyoumuKubunMei " +
                        //"LEFT JOIN Mst_Busho mb ON mb.GyoumuBushoCD = ki.KyourokuIraisakiBushoOld " + // 協力先部所と部所マスタを繋げる（協力先所属長を取得する為）
                        "";
                        cmd.CommandText += "WHERE ki.MadoguchiID = '" + searchID + "' AND KyouryokuDeleteFlag <> 1";

                        // 協力依頼書のコピー先IDをこのタイミングでクリア
                        KakoIraiID = "";
                        Console.WriteLine(cmd.CommandText);
                        var sda = new SqlDataAdapter(cmd);
                        DT_KyouryokuIraisho.Clear();
                        sda.Fill(DT_KyouryokuIraisho);
                    }
                    // 応援受付タブ
                    else if (tab == 5)
                    {
                        //窓口情報取得
                        cmd.CommandText = "SELECT " +
                            "OuenKanriNo " +             // 0:管理番号
                            ",ISNULL(OuenJoukyou,0) " +  // 1:応援状況
                            ",OuenUketsukeDate " +       // 2:応援受付日
                            ",ISNULL(OuenKanryou,0) " +  // 3:応援完了
                            ",OuenHoukokuJishibi " +     // 4:応援完了日
                            "FROM OuenUketsuke " +
                            "WHERE MadoguchiID = " + MadoguchiID + " AND OuenDeleteFlag != 1 ";

                        var sda = new SqlDataAdapter(cmd);
                        DT_Ouenuketsuke.Clear();
                        sda.Fill(DT_Ouenuketsuke);



                    }
                    // 単品入力
                    else if (tab == 6)
                    {
                        //単品入力の取得
                        cmd.CommandText = "SELECT " +
                            " CASE TanpinJutakuDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(TanpinJutakuDate,'yyyy/MM/dd') END " +   // 0:受託日
                            ",CASE TanpinHoukokuDate WHEN '1753/01/01' THEN null WHEN NULL THEN null ELSE FORMAT(TanpinHoukokuDate,'yyyy/MM/dd') END " + // 1:報告日
                            ",TanpinShiji " +             // 2:指示
                            ",TanpinHachuubusho " +       // 3:発注部所
                            ",TanpinTel " +               // 4:電話番号
                            ",TanpinSeikyuuGetsu " +      // 5:請求月
                            ",TanpinYakushoku " +         // 6:役職
                            ",TanpinFax " +               // 7:FAX
                            ",TanpinHachuuTantousha " +   // 8:発注担当者
                            ",TanpinMail " +              // 9:メール

                            ",TanpinMemo " +              // 10:メモ
                                                          //",TanpinRank " +            
                            ",TanpinSaishuuKensa " +      // 11:最終検査
                            ",TanpinShousa " +            // 12:照査
                            ",TanpinMitsumoriTeishutu " + // 13:見積提出
                            ",TanpinShijisho " +          // 14:指示書
                            ",TanpinTeinyuusatsu " +      // 15:
                                                          //",TanpinShuyouChousain " +
                                                          //",TanpinHokurikuShijouKakaku " +
                                                          //",TanpinHokurikuSekouKanka " +
                            ",TanpinSonotaShuukei " +     // 16:その他集計
                            ",TanpinSeikyuuKingaku " +    // 17:請求金額
                            ",TanpinSeikyuuKakutei " +    // 18:請求確定
                            ",TanpinNyuuryokuID " +       // 19:単品入力ID
                                                          //",TanpinGyoumuCD " +

                            "FROM TanpinNyuuryoku " +
                            "WHERE TanpinNyuuryoku.MadoguchiID = '" + MadoguchiID + "'";

                        Console.WriteLine(cmd.CommandText);
                        var sda = new SqlDataAdapter(cmd);
                        DT_Tanpin.Clear();
                        sda.Fill(DT_Tanpin);

                        //単価ランクの取得
                        //cmd.CommandText = "SELECT " +
                        //    " TanpinL1RankMei " +
                        //    ",TanpinL1Ranksuu " +
                        //    ",TanpinL1HoukokuHonsuu " +
                        //    ",TanpinL1Tanka " +
                        //    ",TanpinL1Kingaku " +
                        //    ",TanpunL1RankKubun " +

                        //    "FROM TanpinNyuuryoku " +
                        //    "INNER JOIN TanpinNyuuryokuRank ON TanpinNyuuryoku.TanpinNyuuryokuID = TanpinNyuuryokuRank.TanpinNyuuryokuID " +
                        //    "WHERE TanpinNyuuryoku.MadoguchiID = '" + MadoguchiID + "' ";

                        cmd.CommandText = "SELECT"
                                        //+ " tn.TanpinNyuuryokuID"
                                        //+ ", tn.TanpinGyoumuCD"
                                        //+ ", TanpinL1RankID"
                                        + " tkr.TankaRankHinmoku"                                           // ランク名
                                        + ", ISNULL(tnr.TanpinL1HoukokuHonsuu,0) AS TanpinL1HoukokuHonsuu"  // 報告本数
                                        + ", ISNULL(tnr.TanpinL1Ranksuu,0) AS TanpinL1Ranksuu"              // 依頼本数
                                        + ", tkr.TankaRankKakaku"                                           // 単価 
                                        + ", ISNULL(tnr.TanpinL1Kingaku,0) AS TanpinL1Kingaku"              // 金額
                                        + ", tkr.TankaRankShubetsu"                                         // ランク種別
                                        + ", TanpinL1RankID"                                                // ランクID
                                        + " FROM TanpinNyuuryoku tn"
                                        + " LEFT JOIN TankaKeiyakuRank tkr"
                                        + " ON tkr.TankaKeiyakuID = tn.TanpinGyoumuCD"
                                        + " LEFT JOIN TanpinNyuuryokuRank tnr"
                                        + " ON tnr.TanpinNyuuryokuID = tn.TanpinNyuuryokuID"
                                        + " AND tnr.TanpinL1RankMei = tkr.TankaRankHinmoku"
                                        + " WHERE tn.MadoguchiID = " + MadoguchiID
                                        // えんとり君修正STEP2　並び順追加
                                        //+ " ORDER BY TankaRankHinmoku"
                                        + " ORDER BY TankaRankNarabijunn, TankaRankHinmoku"
                                        ;

                        Console.WriteLine(cmd.CommandText);
                        sda = new SqlDataAdapter(cmd);
                        DT_TanpinRank.Clear();
                        sda.Fill(DT_TanpinRank);
                        c1FlexGrid2.Rows.Count = 1;

                        //単価ランクの編集状態初期化
                        SwichButton_Rank();
                    }
                    else if (tab == 7)
                    {
                        //画面のモード設定 0:新規 1:更新 3:削除
                        //施工条件タブ取得
                        //現在の窓口IDをもったデータ数を取得
                        var sda = new SqlDataAdapter(cmd);
                        cmd.CommandText = "SELECT  TOP 1 " +
                            "cd.countData " +
                            ",sj.SekouJoukenID " +
                            ",sj.SekouJoukenMeijishoID " + //施工条件明示書ID
                            ",SekouKoushuMei " + //工種名
                                                 //◆施工条件（旧）
                            ",SekouTenpuUmu " + //①施工計画書添付の有無 
                            ",SekouGenbaHeimenzu " + //②その他添付資料の現場平面図 5
                            ",SekouDoshituKankeizu " + //②その他添付資料の土質関係図
                            ",SekouSuuryouKeisanzu " + //②その他添付資料の数量計算書
                            ",SekouHiruma " + //③施工時間帯指定の昼間
                            ",SekouYakan " + //③施工時間帯指定の夜間 
                            ",SekouKiseiAri " + //③施工時間帯指定の規制有り10
                            ",SekouSagyouKouritsu " + //④施工条件他の作業効率
                            ",SekouKikai " + //④施工条件他の施工機械の搬入経路
                            ",SekouKasetu " + //④施工条件他の仮設条件
                            ",SekouShizai " + //④施工条件他の資材搬入 
                            ",SekouKensetsu " + //⑤建設機械スペック指定 15
                            ",SekouSuichuu " + //⑥水中施行条件
                            ",SekouSonota " + //⑦その他
                            ",SekouMemo1 " +//メモ1
                            ",SekouMemo2 " +//メモ2
                                            //◆施工条件
                            ",SekouTenpuUmup1Ichizu01 " + //3.添付資料の位置図 20
                            ",SekouTenpuUmup1Sekou02 " + //3.添付資料の施工計画書
                            ",SekouTenpuUmup1Sankou03 " + //3.添付資料の参考カタログ
                            ",SekouTenpuUmup1Ippan04 " + //3.添付資料の一般図・平面図
                            ",SekouTenpuUmup1Genba05 " + //3.添付資料の現場写真 
                            ",SekouTenpuUmup1Kako06 " + //3.添付資料の過去報告書25
                            ",SekouTenpuUmup1Shousai07 " + //3.添付資料の詳細図
                            ",SekouTenpuUmup1Doshitu08 " + //3.添付資料の土質関係図（柱状図等）
                            ",SekouTenpuUmup1Sonota09 " + //3.添付資料のその他
                            ",SekouTenpuUmup1Suuryou10 " + //3.添付資料の数量計算書 
                            ",SekouTenpuUmup1Unpan11 " + //3.添付資料の運搬ルート図30
                            ",SekouSekou2Rikujou01 " + //5.(1)施工場所の陸上
                            ",SekouSekou2Suijou02 " + //5.(1)施工場所の水上
                            ",SekouSekou2Suichuu03 " + //5.(1)施工場所の水中
                            ",SekouSekou2Sonota04 " + //5.(1)施工場所のその他 
                            ",SekouSekou3Tsuujou01 " + //5.(2)施工時間帯の通常昼間施工（8:00~17:00）35
                            ",SekouSekou3Tsuujou02 " + //5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                            ",SekouSekou3Sekou03 " + //5.(2)施工時間帯の施工時間規制あり
                            ",SekouSekou3Nihou04 " + //5.(2)施工時間帯の二方施工（2交代制 昼夜連続施工）
                            ",SekouSekou3Sanpou05 " + //5.(2)施工時間帯の三方施工（3交代制 24時間施工） 
                            ",SekouSagyou4Kankyou01 " + //5.(3)作業環境の現場が狭隘 40
                            ",SekouSagyou4Sekou02 " + //5.(3)作業環境の施工箇所が点在
                            ",SekouSagyou4Joukuu03 " + //5.(3)作業環境の上空制限あり
                            ",SekouSagyou4Sonota04 " + //5.(3)作業環境のその他
                            ",SekouSagyou4Jinka05 " + //5.(3)作業環境の人家に近接（近接施工） 
                            ",SekouSagyou4Tokki06 " + //5.(3)作業環境の特記すべき条件なし 45
                            ",SekouSagyou4Kankyou07 " + //5.(3)作業環境の環境対策あり（騒音・振動）
                            ",SekouSagyou5Koutusu01 " + //5.(4)施工機械・資材搬入経路の交通規制あり
                            ",SekouSagyou5Hannyuu02 " + //5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                            ",SekouSagyou5Sonota03 " + //5.(4)施工機械・資材搬入経路のその他
                            ",SekouSagyou5Tokki04 " + //5.(4)施工機械・資材搬入経路の特記すべき条件なし 50
                            ",SekouKasetsu6Shitei01 " + //5.(5)仮設条件の指定あり
                            ",SekouKasetsu6Shitei02 " + //5.(5)仮設条件の特記すべき条件なし
                            ",SekouSekou7Shitei01 " + //5.(6)施工機械スペック指定の指定あり
                            ",SekouSekou7Shitei02 " + //5.(6)施工機械スペック指定の指定なし
                            ",SekouSonota8Shitei01 " + //5.(7)その他条件の指定あり 55
                            ",SekouSonota8Shitei02 " + //5.(7)その他条件の特記すべき条件なし
                            ",SekouSonotaMemo03 " + //メモ
                            "FROM SekouJouken sj " +
                            ",(SELECT COUNT (*)AS countData " +
                            "FROM SekouJouken WHERE MadoguchiID = " + MadoguchiID + ")cd " +
                            "WHERE MadoguchiID = " + MadoguchiID + " " +
                            "AND SekouDeleteFlag != 1 ";
                        if(SekouJoukenID != "")
                        {
                            cmd.CommandText += "AND SekouJoukenID = '" + SekouJoukenID + "' ";
                        }
                        cmd.CommandText += "ORDER BY SekouJoukenID ";

                        sda = new SqlDataAdapter(cmd);
                        DT_Sekou.Clear();
                        sda.Fill(DT_Sekou);

                        //dataCountが0件の場合、新規モード
                        if (DT_Sekou.Rows.Count == 0)
                        {
                            sekouMode = "0";

                            //追加、削除ボタン活性
                            item7_btnAdd.Enabled = false;
                            item7_btnDelete.Enabled = false;

                            //色変更
                            item7_btnAdd.BackColor = Color.FromArgb(105, 105, 105);
                            item7_btnAdd.ForeColor = Color.FromArgb(169, 169, 169);
                            item7_btnDelete.BackColor = Color.FromArgb(105, 105, 105);
                            item7_btnDelete.ForeColor = Color.FromArgb(169, 169, 169);
                        }
                        //dataCountが1件以上の場合、更新モード
                        else
                        {
                            sekouMode = "1";

                            // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                            sekouMeijishoComboChangeFlg = "1";
                            // 明示書切り替えコンボに施工条件明示書IDを入れる
                            item7_MeijishoKirikaeCombo.Text = DT_Sekou.Rows[0][2].ToString();

                            //追加、削除ボタン非活性
                            item7_btnAdd.Enabled = true;
                            item7_btnDelete.Enabled = true;

                            //色変更
                            item7_btnAdd.BackColor = Color.FromArgb(42, 78, 122);
                            item7_btnAdd.ForeColor = Color.FromArgb(255, 255, 255);
                            item7_btnDelete.BackColor = Color.FromArgb(42, 78, 122);
                            item7_btnDelete.ForeColor = Color.FromArgb(255, 255, 255);
                        }
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
            if (tab == 1)
            {
                // ヘッダー設定
                Header1.Text = MadoguchiData.Rows[0][14].ToString() + "-" + MadoguchiData.Rows[0][15].ToString();
                Header3.Text = MadoguchiData.Rows[0][16].ToString();
                Header4.Text = MadoguchiData.Rows[0][18].ToString();

                //データセット調査概要
                //受託課所支部　契約担当者 受託部所所属長
                item1_MadoguchiJutakuBushoCD.SelectedValue = MadoguchiData.Rows[0][0].ToString();
                item1_AnkenTantoushaMei.Text = MadoguchiData.Rows[0][1].ToString();
                item1_JutakuBushoShozokuChou.Text = MadoguchiData.Rows[0][2].ToString();

                //窓口部所 業務担当者 窓口部所所属長
                item1_MadoguchiTantoushaBushoCD.SelectedValue = MadoguchiData.Rows[0][3].ToString();
                item1_GyoumuKanrishaMei.Text = MadoguchiData.Rows[0][4].ToString();
                item1_MadoguchiBushoShozokuChou.Text = MadoguchiData.Rows[0][5].ToString();

                //窓口担当者 窓口担当者CD　登録年度 管理技術者
                item1_MadoguchiTantousha.Text = MadoguchiData.Rows[0][6].ToString();
                item1_MadoguchiTantoushaCD.Text = MadoguchiData.Rows[0][7].ToString();
                item1_MadoguchiTourokuNendo.SelectedValue = MadoguchiData.Rows[0][8].ToString();
                Nendo = MadoguchiData.Rows[0][8].ToString();
                item1_KanriGijutsushaNM.Text = MadoguchiData.Rows[0][9].ToString();

                //受託番号 受託番号枝番 業務管理者
                item1_MadoguchiJutakuBangou.Text = MadoguchiData.Rows[0][10].ToString();
                item1_MadoguchiJutakuBangouEdaban.Text = MadoguchiData.Rows[0][11].ToString();
                item1_MadoguchiGyoumuKanrisha.Text = MadoguchiData.Rows[0][12].ToString();
                item1_item1_MadoguchiGyoumuKanrishaCD.Text = MadoguchiData.Rows[0][13].ToString();

                //特調番号 特調番号枝番 発注者名・課名
                item1_MadoguchiUketsukeBangou.Text = MadoguchiData.Rows[0][14].ToString();
                item1_MadoguchiUketsukeBangouEdaban.Text = MadoguchiData.Rows[0][15].ToString();
                item1_MadoguchiHachuuKikanmei.Text = MadoguchiData.Rows[0][16].ToString();

                //変更前の番号を所持
                beforeJutaku = item1_MadoguchiUketsukeBangou.Text;
                befireTokuchoEda = item1_MadoguchiUketsukeBangouEdaban.Text;

                //管理番号 業務名称
                item1_MadoguchiKanriBangou.Text = MadoguchiData.Rows[0][17].ToString();
                item1_MadoguchiGyoumuMeishou.Text = MadoguchiData.Rows[0][18].ToString();

                //調査区分 工事件名 契約区分
                if ("1".Equals(MadoguchiData.Rows[0][19].ToString()))
                {
                    item1_MadoguchiChousaKubunJibusho.Checked = true;
                }
                if ("1".Equals(MadoguchiData.Rows[0][20].ToString()))
                {
                    item1_MadoguchiChousaKubunShibuShibu.Checked = true;
                }
                if ("1".Equals(MadoguchiData.Rows[0][21].ToString()))
                {
                    item1_MadoguchiChousaKubunHonbuShibu.Checked = true;
                }
                if ("1".Equals(MadoguchiData.Rows[0][22].ToString()))
                {
                    item1_MadoguchiChousaKubunShibuHonbu.Checked = true;
                }

                item1_MadoguchiKoujiKenmei.Text = MadoguchiData.Rows[0][23].ToString();
                if (!String.IsNullOrEmpty(MadoguchiData.Rows[0][24].ToString()))
                {
                    item1_AnkenGyoumuKubun.SelectedValue = MadoguchiData.Rows[0][24].ToString();
                }

                ////不具合No1345 契約区分により、選択できる調査種別を制限したが、過去データで選択ミスのものについては表示することとする
                int chousaShubetsu = 0;
                if (int.TryParse(MadoguchiData.Rows[0][25].ToString(), out chousaShubetsu))
                {
                    ChousaShubetsuAddData(chousaShubetsu);
                }

                //調査種別 調査品目
                item1_MadoguchiChousaShubetsu.SelectedValue = MadoguchiData.Rows[0][25].ToString();
                item1_MadoguchiChousaHinmoku.Text = MadoguchiData.Rows[0][26].ToString();

                //実施区分 備考 
                item1_MadoguchiJiishiKubun.SelectedValue = MadoguchiData.Rows[0][27].ToString();
                item1_MadoguchiBikou.Text = MadoguchiData.Rows[0][28].ToString();

                //登録日 単価適用地域
                if (MadoguchiData.Rows[0][29].ToString() != "")
                {
                    item1_MadoguchiTourokubi.Text = MadoguchiData.Rows[0][29].ToString();
                }
                else
                {
                    item1_MadoguchiTourokubi.Text = "";
                    item1_MadoguchiTourokubi.CustomFormat = " ";
                }
                item1_MadoguchiTankaTekiyou.Text = MadoguchiData.Rows[0][30].ToString();

                //調査担当者への締切日 荷渡場所
                item1_MadoguchiShimekiribi.Text = MadoguchiData.Rows[0][31].ToString();
                
                item1_MadoguchiNiwatashi.Text = MadoguchiData.Rows[0][32].ToString();

                //報告実施日 遠隔地引渡承認 遠隔地最終検査
                if (MadoguchiData.Rows[0][33].ToString() != "")
                {
                    item1_MadoguchiHoukokuJisshibi.Text = MadoguchiData.Rows[0][33].ToString();
                }
                else
                {
                    item1_MadoguchiHoukokuJisshibi.Text = "";
                    item1_MadoguchiHoukokuJisshibi.CustomFormat = " ";
                }
                if ("0".Equals(MadoguchiData.Rows[0][34].ToString()))
                {
                    item1_MadoguchiHikiwatsahi.Checked = false;
                }
                else
                {
                    item1_MadoguchiHikiwatsahi.Checked = true;
                }
                if ("0".Equals(MadoguchiData.Rows[0][35].ToString()))
                {
                    item1_MadoguchiSaishuuKensa.Checked = false;
                }
                else
                {
                    item1_MadoguchiSaishuuKensa.Checked = true;
                }
                //遠隔地承認者 遠隔地承認日
                item1_MadoguchiShouninsha.Text = MadoguchiData.Rows[0][36].ToString();
                if (MadoguchiData.Rows[0][37].ToString() != "" && MadoguchiData.Rows[0][37].ToString() != "1753/01/01 0:00:00")
                {
                    item1_MadoguchiShouninnbi.Text = MadoguchiData.Rows[0][37].ToString();
                }
                else
                {
                    item1_MadoguchiShouninnbi.Text = "";
                    item1_MadoguchiShouninnbi.CustomFormat = " ";
                }

                //集計表フォルダ 報告書フォルダ 調査資料フォルダ
                item1_MadoguchiShukeiHyoFolder.Text = MadoguchiData.Rows[0][38].ToString();
                item1_MadoguchiHoukokuShoFolder.Text = MadoguchiData.Rows[0][39].ToString();
                item1_MadoguchiShiryouHolder.Text = MadoguchiData.Rows[0][40].ToString();

                //報告済
                if ("1".Equals(MadoguchiData.Rows[0][41].ToString()))
                {
                    item1_MadoguchiHoukokuzumi.Checked = true;
                    button9.Text = "報告完了取消";
                    MadoguchiHoukokuzumi = "1";
                }
                else
                {
                    item1_MadoguchiHoukokuzumi.Checked = false;
                    button9.Text = "報告完了";
                    MadoguchiHoukokuzumi = "0";
                }

                //案件情報ID
                item1_MadoguchiAnkenJouhouID.Text = MadoguchiData.Rows[0][42].ToString();
                //契約担当者CD
                item1_AnkenTantoushaMei_CD.Text = MadoguchiData.Rows[0][43].ToString();
                //管理技術者CD
                item1_KanriGijutsusha_CD.Text = MadoguchiData.Rows[0][44].ToString();
                //本部単品
                item1_MadoguchiHonbuTanpinflg.Checked = bool_str(MadoguchiData.Rows[0][45].ToString());

                //Garoon連携
                item1_GaroonRenkei.Checked = bool_str(MadoguchiData.Rows[0][46].ToString());

            }
            if (tab == 2)
            {
                //担当部所の調査担当者を初期化
                c1FlexGrid1.Rows.Count = 1;
                //c1FlexGrid1.Rows.Count = 31;
                //担当部所のGaroon追加宛先を初期化
                c1FlexGrid5.Rows.Count = 1;

                if (DT_MadoguchiL1Chou != null)
                {
                    //描画停止
                    c1FlexGrid1.BeginUpdate();
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
                        //// ヘッダーで1行使っているので、31行目以降追加する
                        //if (i >= 30)
                        //{
                        //    c1FlexGrid1.Rows.Add();
                        //}
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
                            if (k == 1) { if (KyoroykuBusho1.Text.Equals(DT_MadoguchiL1Chou.Rows[i][7].ToString())) { KyoroykuBusho1.Checked = true;} }
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
                    //描画再開
                    c1FlexGrid1.EndUpdate();
                }
                if (DT_GaroonTsuikaAtesaki != null)
                {
                    // データカウント
                    int count = 0;
                    //描画停止
                    c1FlexGrid5.BeginUpdate();
                    for (int i = 0; i < DT_GaroonTsuikaAtesaki.Rows.Count; i++)
                    {
                        count += 1;
                        //調査担当者をセット
                        //// ヘッダーで1行使っているので、31行目以降追加する
                        //if (i >= 30) { 
                        //c1FlexGrid5.Rows.Add();
                        //}
                        c1FlexGrid5.Rows.Add();
                        c1FlexGrid5.Rows[i + 1].Height = 28;
                        //不具合No1332(1084) 画面から登録されたかのフラグを追加で取得
                        c1FlexGrid5.Rows[i + 1].UserData = DT_GaroonTsuikaAtesaki.Rows[i][3];
                        for (int k = 1; k < c1FlexGrid5.Cols.Count; k++)
                        {
                            c1FlexGrid5.Rows[i + 1][k] = DT_GaroonTsuikaAtesaki.Rows[i][k - 1];
                        }
                        
                    }
                    //for (int i = count + 1;i < 31; i++)
                    //{
                    //    c1FlexGrid5[i + 1, 2] = "";
                    //    c1FlexGrid5[i + 1, 3] = "";
                    //}


                    //描画再開
                    c1FlexGrid5.EndUpdate();
                }
                Resize_Grid("c1FlexGrid1");
                Resize_Grid("c1FlexGrid5");

            }
            else if (tab == 3 && DT_ChousaHinmoku != null)
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

                //c1FlexGrid4.UseCompatibleTextRendering = true;

                // c1FlexGrid4の編集開始行の指定
                int RowCount = 2;

                // VIPS　20220203　課題管理表No797　CHANGE　表示件数「全件表示」対応
                // 表示件数
                int hyoujisuu = 0;
                // 「全件表示」の場合
                if (int.TryParse(item_Hyoujikensuu.Text, out hyoujisuu)==false) {
                    // かなり大きな値をセット
                    hyoujisuu = 999999999;
                }

                // 取得した調査品目のレコードの行を回す
                for (int i = 0; i < DT_ChousaHinmoku.Rows.Count; i++)
                {
                    // 表示件数が超えたらbreak
                    // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
                    if (chousaHinmokuDispFlg == "0" && i >= hyoujisuu)
                    {
                        break;
                    }

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
                    //副調査担当部所1
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


                    // 副調査担当者2（修正後）No1427　1201 嘱託に転籍された方が個人CDで表示される
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
                    //奉行エクセル
                    // 集計表Ver
                    //集計表Verが初期値であれば背景色をグレー
                    if (DT_ChousaHinmoku.Rows[i]["ChousaShuukeihyouVer"].ToString() != "2")
                    {
                        c1FlexGrid4.Rows[RowCount]["ShukeihyoVer"] = "1";
                        c1FlexGrid4.Rows[RowCount]["BunkatsuHouhou"] = "-";
                        c1FlexGrid4.GetCellRange(RowCount, 58).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                        c1FlexGrid4.GetCellRange(RowCount, 59).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                    }
                    else
                    {
                        c1FlexGrid4.Rows[RowCount]["ShukeihyoVer"] = DT_ChousaHinmoku.Rows[i]["ChousaShuukeihyouVer"];
                    }
                    //分割方法（ファイル・シート）
                    if (DT_ChousaHinmoku.Rows[i]["ChousaBunkatsuHouhou"].ToString() == "1" || DT_ChousaHinmoku.Rows[i]["ChousaBunkatsuHouhou"].ToString() == "2")
                    {
                        c1FlexGrid4.Rows[RowCount]["BunkatsuHouhou"] = DT_ChousaHinmoku.Rows[i]["ChousaBunkatsuHouhou"];
                    }
                    else
                    {
                        c1FlexGrid4.Rows[RowCount]["BunkatsuHouhou"] = "-";
                    }                    //// 集計表Ver
                    //c1FlexGrid4.Rows[RowCount]["ShukeihyoVer"] = DT_ChousaHinmoku.Rows[i]["ChousaShuukeihyouVer"];
                    ////集計表Verが初期値であれば背景色をグレー
                    //if (DT_ChousaHinmoku.Rows[i]["ChousaShuukeihyouVer"].ToString() != "2")
                    //{
                    //    c1FlexGrid4.GetCellRange(RowCount, 58).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                    //    c1FlexGrid4.GetCellRange(RowCount, 59).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                    //    c1FlexGrid4.Rows[RowCount]["BunkatsuHouhou"] = "-";
                    //}

                    ////分割方法（ファイル・シート）
                    //if (DT_ChousaHinmoku.Rows[i]["ChousaBunkatsuHouhou"].ToString() == "0")
                    //{
                    //    c1FlexGrid4.Rows[RowCount]["BunkatsuHouhou"] = "-";
                    //}
                    //else
                    //{
                    //    c1FlexGrid4.Rows[RowCount]["BunkatsuHouhou"] = DT_ChousaHinmoku.Rows[i]["ChousaBunkatsuHouhou"];
                    //}
                    //グループ名
                    if (DT_ChousaHinmoku.Rows[i]["ChousaMadoguchiGroupMasterID"].ToString() == "0")
                    {
                        c1FlexGrid4.Rows[RowCount]["GroupMei"] = "";
                    }
                    else
                    {
                        c1FlexGrid4.Rows[RowCount]["GroupMei"] = DT_ChousaHinmoku.Rows[i]["ChousaMadoguchiGroupMasterID"];
                    }
                    //No.1622
                    if (DT_ChousaHinmoku.Rows[i]["ChousaBunkatsuHouhou"].ToString() == "1" && DT_ChousaHinmoku.Rows[i]["ChousaShuukeihyouVer"].ToString() == "2")
                    {
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
                    //作業フォルダ
                    c1FlexGrid4.Rows[RowCount]["SagyoForuda"] = "0";
                    if (chousaLinkFlg != "1")
                    {
                        // 作業フォルダがセットされている場合
                        if (DT_ChousaHinmoku.Rows[i]["Sagyou"].ToString() == "1")
                        {
                            // リンク先に登録されている作業フォルダが存在するか
                            if (Directory.Exists(DT_ChousaHinmoku.Rows[i]["SagyouHolder"].ToString()))
                            {
                                c1FlexGrid4.Rows[RowCount]["SagyoForuda"] = "1";
                            }
                        }
                        else
                        {
                            c1FlexGrid4.Rows[RowCount]["SagyoForuda"] = "";
                        }
                    }
                    // 作業フォルダパス
                    c1FlexGrid4.Rows[RowCount]["SagyoForudaPath"] = DT_ChousaHinmoku.Rows[i]["SagyouHolder"];
                    RowCount += 1;
                }

                //不具合No1207
                //共通マスタの値により、行高さを固定にするか、自動高さ調整を行うか。
                gridRowHeightAutoResize(AutoSizeGridRowMode);

                //描画再開
                c1FlexGrid4.EndUpdate();

                //item3_RegistrationRowCount.Text = DT_ChousaHinmoku.Rows.Count.ToString();
                // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
                if (chousaHinmokuDispFlg == "0")
                {
                    // VIPS　20220203　課題管理表No797　CHANGE　表示件数「全件表示」対応
                    Paging_all.Text = (Math.Ceiling((double)DT_ChousaHinmoku.Rows.Count / hyoujisuu)).ToString();
                }
                else
                {
                    // VIPS　20220203　課題管理表No797　CHANGE　表示件数「全件表示」対応
                    Paging_all.Text = (Math.Ceiling(((double)c1FlexGrid4.Rows.Count - 2) / hyoujisuu)).ToString();
                }
                Grid_Num.Text = "(" + DT_ChousaHinmoku.Rows.Count + ")";
                Grid_Visible(int.Parse(Paging_now.Text));

                //Grid編集許可状態の初期化
                ChousaHinmokuGrid_InputMode();
            }
            // 協力依頼書タブ
            else if (tab == 4 && DT_KyouryokuIraisho != null && DT_KyouryokuIraisho.Rows.Count > 0)
            {
                //item4_KyoRyokuBusho.SelectedValue = DT_KyouryokuIraisho.Rows[0][0].ToString();
                //item4_KyoryokuChou.Text = DT_KyouryokuIraisho.Rows[0][1].ToString();
                //item4_Iraibi.Text = DT_KyouryokuIraisho.Rows[0][2].ToString();
                //item4_HoukokuKigen.Text = DT_KyouryokuIraisho.Rows[0][3].ToString();
                //item4_KyoryokuChou.Text = DT_KyouryokuIraisho.Rows[0][24].ToString();
                //item4_KyoryokuChouCD.Text = DT_KyouryokuIraisho.Rows[0][1].ToString();
                if (DT_KyouryokuIraisho.Rows[0][0].ToString() == "")
                {
                    item4_KyoRyokuBusho.Text = "調査統括部";
                }
                else
                {
                    item4_KyoRyokuBusho.SelectedValue = DT_KyouryokuIraisho.Rows[0][0].ToString();
                    item4_KyoryokuChou.Text = DT_KyouryokuIraisho.Rows[0][24].ToString();
                    item4_KyoryokuChouCD.Text = DT_KyouryokuIraisho.Rows[0][1].ToString();
                }
                if (DT_KyouryokuIraisho.Rows[0][2] != null && DT_KyouryokuIraisho.Rows[0][2].ToString() != "")
                {
                    item4_Iraibi.Text = DT_KyouryokuIraisho.Rows[0][2].ToString();
                    item4_Iraibi.CustomFormat = "";
                }
                else
                {
                    // 392 ②協力依頼書の依頼日初期値対応
                    //item4_Iraibi.CustomFormat = " ";
                    item4_Iraibi.Text = DateTime.Today.ToString();
                    item4_Iraibi.CustomFormat = "";

                }
                if (DT_KyouryokuIraisho.Rows[0][3] != null && DT_KyouryokuIraisho.Rows[0][3].ToString() != "")
                {
                    item4_HoukokuKigen.Text = DT_KyouryokuIraisho.Rows[0][3].ToString();
                    item4_HoukokuKigen.CustomFormat = "";
                }
                else
                {
                    item4_HoukokuKigen.Text = "";
                    item4_HoukokuKigen.CustomFormat = " ";
                }
                item4_GyoumuKubun.SelectedValue = DT_KyouryokuIraisho.Rows[0][4];

                item4_IraiKubun.SelectedValue = DT_KyouryokuIraisho.Rows[0][5];
                if (DT_KyouryokuIraisho.Rows[0][6].ToString().Equals("1"))
                {
                    item4_NaiyoKubun_Shizai.Checked = true;
                }
                if (DT_KyouryokuIraisho.Rows[0][7].ToString().Equals("1"))
                {
                    item4_NaiyoKubun_DKou.Checked = true;
                }
                if (DT_KyouryokuIraisho.Rows[0][8].ToString().Equals("1"))
                {
                    item4_NaiyoKubun_EKou.Checked = true;
                }
                if (DT_KyouryokuIraisho.Rows[0][9].ToString().Equals("1"))
                {
                    item4_NaiyoKubun_Sonota.Checked = true;
                }
                if (DT_KyouryokuIraisho.Rows[0][10].ToString().Equals("1"))
                {
                    item4_NaiyoKubun_Joho.Checked = true;
                }
                item4_RenrakuJikou.Text = DT_KyouryokuIraisho.Rows[0][11].ToString();
                item4_GyoumuNaiyo.Text = DT_KyouryokuIraisho.Rows[0][12].ToString();
                //item4_Zumen.SelectedValue = DT_KyouryokuIraisho.Rows[0][13];
                if (DT_KyouryokuIraisho.Rows[0][13] != null && DT_KyouryokuIraisho.Rows[0][13].ToString() != "")
                {
                    item4_Zumen.SelectedValue = DT_KyouryokuIraisho.Rows[0][13];
                }

                item4_Kizyunbi.SelectedValue = DT_KyouryokuIraisho.Rows[0][14];
                //item4_KizyunbiStr.Text = DT_KyouryokuIraisho.Rows[0][15].ToString();
                if (DT_KyouryokuIraisho.Rows[0][15] != null && DT_KyouryokuIraisho.Rows[0][15].ToString() != "")
                {
                    item4_KizyunbiStr.Text = DT_KyouryokuIraisho.Rows[0][15].ToString();
                }
                item4_UtiawaseYouhi.SelectedValue = DT_KyouryokuIraisho.Rows[0][16];
                //item4_Gutaiteki.SelectedValue = DT_KyouryokuIraisho.Rows[0][17];

                if (DT_KyouryokuIraisho.Rows[0][17] != null && DT_KyouryokuIraisho.Rows[0][17].ToString() != "")
                {
                    item4_Gutaiteki.SelectedValue = DT_KyouryokuIraisho.Rows[0][17];
                }
                if (DT_KyouryokuIraisho.Rows[0][18] != null && DT_KyouryokuIraisho.Rows[0][18].ToString() != "")
                {
                    item4_ZenkaiKyouryoku.SelectedValue = DT_KyouryokuIraisho.Rows[0][18];
                }
                //item4_ZenkaiKyouryoku.SelectedValue = DT_KyouryokuIraisho.Rows[0][18];

                item4_ZenkaiKyouryokuStr.Text = DT_KyouryokuIraisho.Rows[0][19].ToString();
                item4_Hikiwatashi.SelectedValue = DT_KyouryokuIraisho.Rows[0][20];
                //item4_JishiKeikakusho.SelectedValue = DT_KyouryokuIraisho.Rows[0][21];
                if (DT_KyouryokuIraisho.Rows[0][21] != null && DT_KyouryokuIraisho.Rows[0][21].ToString() != "")
                {
                    item4_JishiKeikakusho.SelectedValue = DT_KyouryokuIraisho.Rows[0][21];
                }
                if (DT_KyouryokuIraisho.Rows[0][22] != null && DT_KyouryokuIraisho.Rows[0][22].ToString() != "")
                {
                    item4_MitsumoriChousyu.SelectedValue = DT_KyouryokuIraisho.Rows[0][22];
                }
                //item4_MitsumoriChousyu.SelectedValue = DT_KyouryokuIraisho.Rows[0][22];
                item4_KyouryokuIraishoID.Text = DT_KyouryokuIraisho.Rows[0][23].ToString();

                int num = 0;
                // 資材
                if (DT_KyouryokuIraisho.Rows[0][25] != null)
                {
                    if (int.TryParse(DT_KyouryokuIraisho.Rows[0][25].ToString(), out num))
                    {
                        if (num > 0)
                        {
                            item4_NaiyoKubun_Shizai.Checked = true;
                        }
                        else
                        {
                            item4_NaiyoKubun_Shizai.Checked = false;
                        }
                    }
                }
                num = 0;
                // D工
                if (DT_KyouryokuIraisho.Rows[0][26] != null)
                {
                    if (int.TryParse(DT_KyouryokuIraisho.Rows[0][26].ToString(), out num))
                    {
                        if (num > 0)
                        {
                            item4_NaiyoKubun_DKou.Checked = true;
                        }
                        else
                        {
                            item4_NaiyoKubun_DKou.Checked = false;
                        }
                    }
                }
                num = 0;
                // E工
                if (DT_KyouryokuIraisho.Rows[0][27] != null)
                {
                    if (int.TryParse(DT_KyouryokuIraisho.Rows[0][27].ToString(), out num))
                    {
                        if (num > 0)
                        {
                            item4_NaiyoKubun_EKou.Checked = true;
                        }
                        else
                        {
                            item4_NaiyoKubun_EKou.Checked = false;
                        }
                    }
                }
                num = 0;
                // その他
                if (DT_KyouryokuIraisho.Rows[0][28] != null)
                {
                    if (int.TryParse(DT_KyouryokuIraisho.Rows[0][28].ToString(), out num))
                    {
                        if (num > 0)
                        {
                            item4_NaiyoKubun_Sonota.Checked = true;
                        }
                        else
                        {
                            item4_NaiyoKubun_Sonota.Checked = false;
                        }
                    }
                }

                string jouhouSystem = "";
                if (DT_KyouryokuIraisho.Rows[0][29] != null) {
                    // 情シス部（一般契約）
                    jouhouSystem = GlobalMethod.GetCommonValue1("JOUHOU_KAIHATSU_GYOUMU");
                    if (jouhouSystem == null || jouhouSystem == "")
                    {
                        jouhouSystem = "情シス部（一般契約）";
                    }
                    if (jouhouSystem.Equals(DT_KyouryokuIraisho.Rows[0][29].ToString()))
                    {
                        item4_NaiyoKubun_Joho.Checked = true;
                    }
                    else
                    {
                        item4_NaiyoKubun_Joho.Checked = false;
                    }
                }
                // 業務区分
                if (DT_KyouryokuIraisho.Rows[0][30] != null)
                {
                    num = 0;
                    if (int.TryParse(DT_KyouryokuIraisho.Rows[0][30].ToString(), out num))
                    {
                        // GyoumuNarabijunCD の値で業務区分の値を切替
                        // 1:調査部（一般）
                        // 2:調査部（単契含む）
                        // 3:調査部（単契）
                        // 4:調査部（単品)
                        // 5:事業普及部（一般）
                        // 6:事業普及部（物品購入）
                        // 7:情シス部（一般契約）
                        // 8:総合研究所
                        if (num == 1 || num == 5 || num == 6 || num == 8)
                        {
                            item4_GyoumuKubun.SelectedValue = 1;    // 1.一般受託調査
                        }
                        else if (num == 3)
                        {
                            item4_GyoumuKubun.SelectedValue = 2;    // 2.単価契約調査
                        }
                        else if (num == 4)
                        {
                            item4_GyoumuKubun.SelectedValue = 3;    // 3.単品契約調査
                        }
                        else if (num == 2)
                        {
                            item4_GyoumuKubun.SelectedValue = 4;    // 4.単価契約を含む一般受託
                        }
                        else if (num == 7)
                        {
                            item4_GyoumuKubun.SelectedValue = 5;    // 5.情報開発受託業務
                        }
                        else
                        {
                            item4_GyoumuKubun.SelectedValue = 1;    // 1.一般受託調査
                        }
                    }
                }

                // 996協力依頼書ボタンを押せるように
                //// 調査区分が支→本でない場合は、協力依頼書を出力できないようにする。
                //if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
                //{
                //    // 活性
                //    button24.Enabled = true;
                //}
                //else
                //{
                //    // 非活性
                //    button24.Enabled = false;
                //    button24.BackColor = Color.DarkGray;
                //}

            }
            // 応援受付
            else if (tab == 5 && DT_Ouenuketsuke != null && DT_Ouenuketsuke.Rows.Count > 0)
            {
                // 管理番号
                item5_Kanribangou.Text = DT_Ouenuketsuke.Rows[0][0].ToString();

                // 調査概要タブの管理番号を表示（更新は調査概要から）
                item5_Kanribangou.Text = item1_MadoguchiKanriBangou.Text;
                // 応援状況
                //if (DT_Ouenuketsuke.Rows[0][1] != null && DT_Ouenuketsuke.Rows[0][1].ToString() == "1")
                if (DT_Ouenuketsuke.Rows[0][1] != null)
                {
                    item5_UketsukeJoukyo.Checked = false;
                    if (DT_Ouenuketsuke.Rows[0][1].ToString() == "2")
                    {
                        item5_UketsukeJoukyo.Checked = true;

                        // チェック時は完了アイコン
                        UketsukeIcon.Image = Image.FromFile("Resource/kan.png");

                    }
                    //else if (DT_Ouenuketsuke.Rows[0][1].ToString() == "1")
                    // 支→本にチェックが入っていれば依頼マークを出す
                    else if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
                    {
                        UketsukeIcon.Image = Image.FromFile("Resource/OnegaiIcon35px.png");
                    }
                    // 支→本のみ表示
                    if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
                    {
                        UketsukeIcon.Visible = true;
                    }
                    else
                    {
                        UketsukeIcon.Visible = false;
                    }
                }
                else
                {
                    item5_UketsukeJoukyo.Checked = false;
                    UketsukeIcon.Visible = false;
                }
                // 応援受付日
                if (DT_Ouenuketsuke.Rows[0][2] != null && DT_Ouenuketsuke.Rows[0][2].ToString() != "")
                {
                    item5_UketsukeDate.Text = DT_Ouenuketsuke.Rows[0][2].ToString();
                    item5_UketsukeDate.CustomFormat = "";
                }
                else
                {
                    item5_UketsukeDate.Text = "";
                    item5_UketsukeDate.CustomFormat = " ";
                }

                item5_OuenKanryo.Checked = false;
                // 応援完了
                if (DT_Ouenuketsuke.Rows[0][3] != null && DT_Ouenuketsuke.Rows[0][3].ToString() == "1")
                {
                    item5_OuenKanryo.Checked = true;
                    // 支→本のみ表示
                    if (item1_MadoguchiChousaKubunShibuHonbu.Checked) {
                        KanryouIcon.Visible = true;
                    }
                    else
                    {
                        KanryouIcon.Visible = false;
                    }
                }
                else
                {
                    KanryouIcon.Visible = false;
                }
                // 応援完了日
                if (DT_Ouenuketsuke.Rows[0][4] != null && DT_Ouenuketsuke.Rows[0][4].ToString() != "")
                {
                    item5_OuenKanryoDate.Text = DT_Ouenuketsuke.Rows[0][4].ToString();
                    item5_OuenKanryoDate.CustomFormat = "";
                }
                else
                {
                    item5_OuenKanryoDate.Text = "";
                    item5_OuenKanryoDate.CustomFormat = " ";
                }

                //// 調査区分が支→本でない場合は、協力依頼書を出力できないようにする。
                //if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
                //{
                //    // 活性
                //    btnIraisho.Enabled = true;
                //}
                //else
                //{
                //    // 非活性
                //    btnIraisho.Enabled = false;
                //    btnIraisho.BackColor = Color.DarkGray;
                //}

            }
            // 単品入力
            else if (tab == 6)
            {
                if (DT_Tanpin != null && DT_Tanpin.Rows.Count > 0)
                {
                    if (DT_Tanpin.Rows[0][0] != null && DT_Tanpin.Rows[0][0].ToString() != "")
                    {
                        item6_TanpinJutakuDate.Text = DT_Tanpin.Rows[0][0].ToString();
                        item6_TanpinJutakuDate.CustomFormat = "";
                    }
                    else
                    {
                        // VIPS　20220316　課題管理表No1168(868)　CHANGE　単品入力タブの受託日(依頼日)初期値を調査概要タブの登録日に変更
                        //item6_TanpinJutakuDate.Text = "";
                        //item6_TanpinJutakuDate.CustomFormat = " ";
                        item6_TanpinJutakuDate.Text = MadoguchiData.Rows[0][29].ToString();
                        item6_TanpinJutakuDate.CustomFormat = "";
                    }
                    if (DT_Tanpin.Rows[0][1] != null && DT_Tanpin.Rows[0][1].ToString() != "")
                    {
                        item6_TanpinHoukokuDate.Text = DT_Tanpin.Rows[0][1].ToString();
                        item6_TanpinHoukokuDate.CustomFormat = "";
                    }
                    else
                    {
                        item6_TanpinHoukokuDate.Text = "";
                        item6_TanpinHoukokuDate.CustomFormat = " ";
                    }
                    item6_TanpinShiji.Text = DT_Tanpin.Rows[0][2].ToString();
                    item6_TanpinHachuubusho.Text = DT_Tanpin.Rows[0][3].ToString();
                    item6_TanpinTel.Text = DT_Tanpin.Rows[0][4].ToString();
                    if (DT_Tanpin.Rows[0][5].ToString() != "")
                    {
                        item6_TanpinSeikyuuGetsu.Text = DT_Tanpin.Rows[0][5].ToString();
                    }
                    item6_TanpinYakushoku.Text = DT_Tanpin.Rows[0][6].ToString();
                    item6_TanpinFax.Text = DT_Tanpin.Rows[0][7].ToString();
                    item6_TanpinHachuuTantousha.Text = DT_Tanpin.Rows[0][8].ToString();
                    item6_TanpinMail.Text = DT_Tanpin.Rows[0][9].ToString();
                    item6_TanpinMemo.Text = DT_Tanpin.Rows[0][10].ToString();
                    if (DT_Tanpin.Rows[0][11].ToString().Equals("1"))
                    {
                        item6_TanpinSaishuuKensa.Checked = true;
                    }
                    item6_TanpinShousa.Text = DT_Tanpin.Rows[0][12].ToString();
                    if (DT_Tanpin.Rows[0][13].ToString().Equals("1"))
                    {
                        item6_TanpinMitsumoriTeishutu.Checked = true;
                    }
                    //if(DT_Tanpin.Rows[0][14] != null) { 
                    //    // システムエラーとなるので一旦コメントアウト
                    //    //item6_TanpinShijisho.SelectedValue = DT_Tanpin.Rows[0][14].ToString();
                    //}
                    if (DT_Tanpin.Rows[0][14] != null && DT_Tanpin.Rows[0][14].ToString() != "")
                    {
                        item6_TanpinShijisho.SelectedValue = DT_Tanpin.Rows[0][14].ToString();
                    }
                    if (DT_Tanpin.Rows[0][15].ToString().Equals("1"))
                    {
                        item6_TanpinTeinyuusatsu.Checked = true;
                    }
                    item6_TanpinSonotaShuukei.Text = GetMoneyText(GetLong(DT_Tanpin.Rows[0][16].ToString()));
                    item6_TanpinSeikyuuKingaku.Text = GetMoneyText(GetLong(DT_Tanpin.Rows[0][17].ToString()));
                    if (DT_Tanpin.Rows[0][18].ToString().Equals("1"))
                    {
                        item6_TanpinSeikyuuKakutei.Checked = true;
                    }
                    item6_TanpinNyuuryokuID.Text = DT_Tanpin.Rows[0][19].ToString();
                }
                //単品ランクが登録済みの場合
                if (DT_TanpinRank != null && DT_TanpinRank.Rows.Count > 0)
                {
                    //描画停止
                    c1FlexGrid2.BeginUpdate();
                    for (int i = 0; i < DT_TanpinRank.Rows.Count; i++)
                    {
                        c1FlexGrid2.Rows.Add();
                        for (int k = 0; k < c1FlexGrid2.Cols.Count; k++)
                        {
                            c1FlexGrid2[i + 1, k] = DT_TanpinRank.Rows[i][k];
                        }
                    }

                    ReCal();
                    Resize_Grid("c1FlexGrid2");

                    //描画再開
                    c1FlexGrid2.EndUpdate();
                }
                //単品ランクが未登録の場合
                else
                {
                    //単価テーブルから集計
                    //AggregateRank();
                }
            }
            else if (tab == 7)
            {
                // 施工条件 0:新規 1:更新 2:削除
                if (!"0".Equals(sekouMode))
                {
                    //登録数 施工条件ID
                    item7_TourokuSuu.Text = DT_Sekou.Rows[0][0].ToString();
                    SekouJoukenID = DT_Sekou.Rows[0][1].ToString();

                    // 施工条件タブ 施工条件明示書ID変更フラグ 0:手動変更 1:システム側で変更
                    sekouMeijishoIDChangeFlg = "1";
                    //明示ID　工種名
                    item7_SekouJoukenMeijishoID.Text = DT_Sekou.Rows[0][2].ToString();
                    item7_KoushuMei.Text = DT_Sekou.Rows[0][3].ToString();

                    //◆施工条件旧
                    //①施工計画書添付の有無 　
                    checkBox80.Checked = bool_str(DT_Sekou.Rows[0][4].ToString());
                    //②その他添付資料の現場平面図　②その他添付資料の土質関係図　②その他添付資料の数量計算書
                    checkBox81.Checked = bool_str(DT_Sekou.Rows[0][5].ToString());
                    checkBox82.Checked = bool_str(DT_Sekou.Rows[0][6].ToString());
                    checkBox83.Checked = bool_str(DT_Sekou.Rows[0][7].ToString());
                    //③施工時間帯指定の昼間　③施工時間帯指定の夜間　③施工時間帯指定の規制有り
                    checkBox84.Checked = bool_str(DT_Sekou.Rows[0][8].ToString());
                    checkBox85.Checked = bool_str(DT_Sekou.Rows[0][9].ToString());
                    checkBox86.Checked = bool_str(DT_Sekou.Rows[0][10].ToString());
                    //④施工条件他の作業効率　④施工条件他の施工機械の搬入経路
                    checkBox87.Checked = bool_str(DT_Sekou.Rows[0][11].ToString());
                    checkBox88.Checked = bool_str(DT_Sekou.Rows[0][12].ToString());
                    //④施工条件他の仮設条件　④施工条件他の資材搬入
                    checkBox89.Checked = bool_str(DT_Sekou.Rows[0][13].ToString());
                    checkBox93.Checked = bool_str(DT_Sekou.Rows[0][14].ToString());
                    //⑤建設機械スペック指定　⑥水中施行条件　⑦その他
                    checkBox90.Checked = bool_str(DT_Sekou.Rows[0][15].ToString());
                    checkBox91.Checked = bool_str(DT_Sekou.Rows[0][16].ToString());
                    checkBox92.Checked = bool_str(DT_Sekou.Rows[0][17].ToString());
                    //メモ1　メモ2
                    textBox41.Text = DT_Sekou.Rows[0][18].ToString();
                    textBox42.Text = DT_Sekou.Rows[0][19].ToString();

                    //◆施工条件
                    //3.添付資料の位置図 3.添付資料の施工計画書 3.添付資料の参考カタログ
                    checkBox43.Checked = bool_str(DT_Sekou.Rows[0][20].ToString());
                    checkBox47.Checked = bool_str(DT_Sekou.Rows[0][21].ToString());
                    checkBox51.Checked = bool_str(DT_Sekou.Rows[0][22].ToString());

                    //3.添付資料の一般図・平面図 3.添付資料の現場写真 3.添付資料の過去報告書
                    checkBox44.Checked = bool_str(DT_Sekou.Rows[0][23].ToString());
                    checkBox48.Checked = bool_str(DT_Sekou.Rows[0][24].ToString());
                    checkBox52.Checked = bool_str(DT_Sekou.Rows[0][25].ToString());

                    //3.添付資料の詳細図 3.添付資料の土質関係図（柱状図等）3.添付資料のその他
                    checkBox45.Checked = bool_str(DT_Sekou.Rows[0][26].ToString());
                    checkBox49.Checked = bool_str(DT_Sekou.Rows[0][27].ToString());
                    checkBox53.Checked = bool_str(DT_Sekou.Rows[0][28].ToString());

                    //3.添付資料の数量計算書 3.添付資料の運搬ルート図
                    checkBox46.Checked = bool_str(DT_Sekou.Rows[0][29].ToString());
                    checkBox50.Checked = bool_str(DT_Sekou.Rows[0][30].ToString());

                    //5.(1)施工場所の陸上 5.(1)施工場所の水上 
                    checkBox54.Checked = bool_str(DT_Sekou.Rows[0][31].ToString());
                    checkBox55.Checked = bool_str(DT_Sekou.Rows[0][32].ToString());

                    //5.(1)施工場所の水中 5.(1)施工場所のその他
                    checkBox56.Checked = bool_str(DT_Sekou.Rows[0][33].ToString());
                    checkBox57.Checked = bool_str(DT_Sekou.Rows[0][34].ToString());

                    //5.(2)施工時間帯の通常昼間施工（8:00~17:00） 5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                    checkBox58.Checked = bool_str(DT_Sekou.Rows[0][35].ToString());
                    checkBox60.Checked = bool_str(DT_Sekou.Rows[0][36].ToString());

                    //5.(2)施工時間帯の施工時間規制あり 5.(2)施工時間帯の二方施工（2交代制 昼夜連続施工）
                    checkBox62.Checked = bool_str(DT_Sekou.Rows[0][37].ToString());
                    checkBox59.Checked = bool_str(DT_Sekou.Rows[0][38].ToString());

                    //5.(2)施工時間帯の三方施工（3交代制 24時間施工）
                    checkBox61.Checked = bool_str(DT_Sekou.Rows[0][39].ToString());

                    //5.(3)作業環境の現場が狭隘  5.(3)作業環境の施工箇所が点在 5.(3)作業環境の上空制限あり
                    checkBox63.Checked = bool_str(DT_Sekou.Rows[0][40].ToString());
                    checkBox67.Checked = bool_str(DT_Sekou.Rows[0][41].ToString());
                    checkBox64.Checked = bool_str(DT_Sekou.Rows[0][42].ToString());

                    //5.(3)作業環境のその他 5.(3)作業環境の人家に近接（近接施工） 5.(3)作業環境の特記すべき条件なし
                    checkBox68.Checked = bool_str(DT_Sekou.Rows[0][43].ToString());
                    checkBox65.Checked = bool_str(DT_Sekou.Rows[0][44].ToString());
                    checkBox70.Checked = bool_str(DT_Sekou.Rows[0][45].ToString());

                    //5.(3)作業環境の環境対策あり（騒音・振動）
                    checkBox66.Checked = bool_str(DT_Sekou.Rows[0][46].ToString());

                    //5.(4)施工機械・資材搬入経路の交通規制あり 5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                    checkBox69.Checked = bool_str(DT_Sekou.Rows[0][47].ToString());
                    checkBox71.Checked = bool_str(DT_Sekou.Rows[0][48].ToString());

                    //5.(4)施工機械・資材搬入経路のその他 5.(4)施工機械・資材搬入経路の特記すべき条件なし
                    checkBox72.Checked = bool_str(DT_Sekou.Rows[0][49].ToString());
                    checkBox73.Checked = bool_str(DT_Sekou.Rows[0][50].ToString());

                    //5.(5)仮設条件の指定あり 5.(5)仮設条件の特記すべき条件なし 
                    checkBox74.Checked = bool_str(DT_Sekou.Rows[0][51].ToString());
                    checkBox75.Checked = bool_str(DT_Sekou.Rows[0][52].ToString());

                    //5.(6)施工機械スペック指定の指定あり 5.(6)施工機械スペック指定の指定なし 
                    checkBox76.Checked = bool_str(DT_Sekou.Rows[0][53].ToString());
                    checkBox77.Checked = bool_str(DT_Sekou.Rows[0][54].ToString());

                    //5.(7)その他条件の指定あり  5.(7)その他条件の特記すべき条件なし
                    checkBox78.Checked = bool_str(DT_Sekou.Rows[0][55].ToString());
                    checkBox79.Checked = bool_str(DT_Sekou.Rows[0][56].ToString());

                    //メモ
                    textBox40.Text = DT_Sekou.Rows[0][57].ToString();
                }
            }
            //レイアウトロジックを再開する
            this.ResumeLayout();
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

        private string str_bool(Boolean check)
        {
            //文字列0と1をBooleanに変換
            string checkValue = "0";

            //checkがtrueのとき
            if (check)
            {
                checkValue = "1";
            }
            //checkがfalseのとき
            else
            {
                checkValue = "0";
            }
            return checkValue;
        }

        ////public ToolStripMenuItem Set_ContextMenu(ToolStripMenuItem item, DataTable dt)
        ////{
        ////    for (int i = 0; i < dt.Rows.Count; i++)
        ////    {
        ////        if (dt.Rows[i][1].ToString() != "")
        ////        {
        ////            item.DropDownItems.Add(dt.Rows[i][1].ToString(), null, ContextMenuEvent);
        ////        }
        ////    }
        ////    return item;
        ////}
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
        // 調査担当者の右クリックメニュークリック時のイベント
        private void contextMenuTantoushaItemClicked2(object sender, EventArgs e)
        {
            ToolStripMenuItem mi = (ToolStripMenuItem)sender;
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
        // 担当部所タブの更新ボタン
        private void button_update_2_Click(object sender, EventArgs e)
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

                    // 部所が選択されており、担当者が空の場合、エラー
                    // 2:担当部所 3:担当者
                    //if (c1FlexGrid5.Rows[i][2] != null && c1FlexGrid5.Rows[i][3] != null && !"".Equals(c1FlexGrid5.Rows[i][2].ToString()) && !"0".Equals(c1FlexGrid5.Rows[i][2].ToString()) && "".Equals(c1FlexGrid5.Rows[i][3].ToString()))
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
                    UpdateMadoguchi(2);
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

        // 調査品目の更新ボタン
        private void button_update_3_Click(object sender, EventArgs e)
        {
            UpdateMadoguchi(3);
        }

        // 協力依頼書の更新ボタン
        private void button_update_4_Click(object sender, EventArgs e)
        {
            // エラークリア
            ErrorClear_KyouryokuIrai();

            if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                UpdateMadoguchi(4);
            }
        }

        // 単品入力項目の更新ボタン
        private void button_update_6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                //単価ランクの編集状態初期化（報告ランク）
                SwichButton_Rank();
                //依頼ランクの集計結果で更新されないように再計算
                ReCal();
                UpdateMadoguchi(6);
            }
        }

        //不具合No1338
        private void UpdateMadoguchiRealTime()
        {
            //登録データ配列 必要なところはごくわずかだが、既存メソッドUpdateMadoguchiと併せておく。
            string[,] SQLData = new string[1, 60];
            string mes = "";

            //報告済
            string houkokuzumi = str_bool(item1_MadoguchiHoukokuzumi.Checked);
            
            //配列に格納
            SQLData[0, 1] = MadoguchiID;    //これ必要ないような
            SQLData[0, 34] = houkokuzumi;

            //テーブル登録・更新
            Boolean result = GlobalMethod.MadoguchiUpdateRealTime_SQL(MadoguchiID, SQLData, out mes, UserInfos);

            set_error("", 0);
            set_error(mes);

            //担当部署と、調査品目明細のタブを表示しなおし。
            get_data(2);
            get_data(3);

        }

        private void UpdateMadoguchi(int tab)
        {
            //更新成功フラグ
            Boolean UpdateFlag = true;

            //調査概要
            if (tab == 1)
            {
                //登録データ配列
                string[,] SQLData = new string[1, 60];
                string mes = "";

                //実施区分
                String jisshiKubun = "NULL";
                if (!String.IsNullOrEmpty(item1_MadoguchiJiishiKubun.Text))
                {
                    jisshiKubun = item1_MadoguchiJiishiKubun.SelectedValue.ToString();
                }

                //遠隔地引渡承認　遠隔地最終検査
                string hikiwatashiShounin = str_bool(item1_MadoguchiHikiwatsahi.Checked);
                string saishuukensa = str_bool(item1_MadoguchiSaishuuKensa.Checked);

                // 本部単品　Garoon連携
                string honbuTanpin = str_bool(item1_MadoguchiHonbuTanpinflg.Checked);
                string garoon = str_bool(item1_GaroonRenkei.Checked);

                //報告済
                string houkokuzumi = str_bool(item1_MadoguchiHoukokuzumi.Checked);

                //調査区分 
                string chousaKubunJibusho = str_bool(item1_MadoguchiChousaKubunJibusho.Checked);
                string chousaKubunShibushibu = str_bool(item1_MadoguchiChousaKubunShibuShibu.Checked);
                string chousaKubunHonshibu = str_bool(item1_MadoguchiChousaKubunHonbuShibu.Checked);
                string chousaKubunShibuhon = str_bool(item1_MadoguchiChousaKubunShibuHonbu.Checked);

                //調査種別
                String shubetsu = "NULL";
                if (!String.IsNullOrEmpty(item1_MadoguchiChousaShubetsu.Text))
                {
                    shubetsu = item1_MadoguchiChousaShubetsu.SelectedValue.ToString();
                }

                //案件情報ID
                String ankenJouhouId = item1_MadoguchiAnkenJouhouID.Text;

                //契約担当者CD
                String keiyakutantou = "NULL";
                if (!String.IsNullOrEmpty(item1_AnkenTantoushaMei_CD.Text))
                {
                    keiyakutantou = item1_AnkenTantoushaMei_CD.Text;
                }

                //業務管理者CD
                String gyoumuKanri = "NULL";
                if (!String.IsNullOrEmpty(item1_item1_MadoguchiGyoumuKanrishaCD.Text))
                {
                    gyoumuKanri = item1_item1_MadoguchiGyoumuKanrishaCD.Text;
                }

                //窓口担当者CD
                String madoguchiTantouCD = "0";
                if (!String.IsNullOrEmpty(item1_MadoguchiTantoushaCD.Text))
                {
                    madoguchiTantouCD = item1_MadoguchiTantoushaCD.Text;
                }

                //調査概要画面モード
                SQLData[0, 0] = mode;

                //新規のとき
                if ("insert".Equals(mode))
                {
                    string renban = TokuchoNo_saiban();
                    SQLData[0, 1] = saibanMadoguchiNo;//採番された最新の窓口ID
                    SQLData[0, 2] = item1_MadoguchiTourokuNendo.SelectedValue.ToString();//登録年度
                    SQLData[0, 3] = hikiwatashiShounin;//遠隔地引渡承認
                    SQLData[0, 4] = saishuukensa;             //遠隔地最終検査
                    SQLData[0, 5] = GlobalMethod.ChangeSqlText(item1_MadoguchiShouninsha.Text, 0, 0);          //遠隔地承認者
                    SQLData[0, 6] = Get_DateTimePicker("item1_MadoguchiShouninnbi");   //遠隔地承認日
                    SQLData[0, 7] = Get_DateTimePicker("item1_MadoguchiShimekiribi");   //調査担当者への締切日
                    SQLData[0, 8] = Get_DateTimePicker("item1_MadoguchiTourokubi");           //登録日
                    SQLData[0, 9] = Get_DateTimePicker("item1_MadoguchiHoukokuJisshibi");   //報告実施日
                    SQLData[0, 10] = shubetsu;   //調査種別　
                    SQLData[0, 11] = jisshiKubun;                                 //実施区分　
                    SQLData[0, 12] = "10";                         //MadoguchiShinchokuJoukyou
                    SQLData[0, 13] = item1_MadoguchiJutakuBushoCD.SelectedValue.ToString();   //受託課所支部
                    SQLData[0, 14] = keiyakutantou;           //契約担当者orNULL
                    SQLData[0, 15] = item1_MadoguchiJutakuBushoCD.SelectedValue.ToString();          //受託部所所属長の部所CD
                    SQLData[0, 16] = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString(); //窓口部所
                    SQLData[0, 17] = madoguchiTantouCD;            //窓口担当者
                    SQLData[0, 18] = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString(); //窓口部所所属長の部所CD
                    SQLData[0, 19] = chousaKubunJibusho;// 調査区分　自部所
                    SQLData[0, 20] = chousaKubunShibushibu; //調査区分　支→支
                    SQLData[0, 21] = chousaKubunHonshibu; // 調査区分　本→支
                    SQLData[0, 22] = chousaKubunShibuhon; //調査区分　支→本
                    SQLData[0, 23] = item1_MadoguchiKanriBangou.Text;           //管理番号
                    SQLData[0, 24] = item1_MadoguchiJutakuBangou.Text;          //受託番号
                    SQLData[0, 25] = item1_MadoguchiJutakuBangouEdaban.Text;       //受託番号枝番
                    SQLData[0, 26] = item1_MadoguchiUketsukeBangou.Text;            //特調番号
                    SQLData[0, 27] = item1_MadoguchiUketsukeBangouEdaban.Text;          //特調番号枝番
                    SQLData[0, 28] = GlobalMethod.ChangeSqlText(item1_MadoguchiHachuuKikanmei.Text, 0, 0);          //発注者名・課名
                    SQLData[0, 29] = GlobalMethod.ChangeSqlText(item1_MadoguchiGyoumuMeishou.Text, 0, 0);          //業務名称
                    SQLData[0, 30] = GlobalMethod.ChangeSqlText(item1_MadoguchiKoujiKenmei.Text, 0, 0);          //工事件名
                    SQLData[0, 31] = GlobalMethod.ChangeSqlText(item1_MadoguchiChousaHinmoku.Text, 0, 0);          //調査品目
                    SQLData[0, 32] = GlobalMethod.ChangeSqlText(item1_MadoguchiBikou.Text, 0, 0);          //備考
                    SQLData[0, 33] = GlobalMethod.ChangeSqlText(item1_MadoguchiTankaTekiyou.Text, 0, 0);          //単価適用地域
                    SQLData[0, 34] = GlobalMethod.ChangeSqlText(item1_MadoguchiNiwatashi.Text, 0, 0);         //荷渡場所
                    SQLData[0, 35] = "0";                         //0　報告済
                    SQLData[0, 36] = item1_KanriGijutsusha_CD.Text;           //管理技術者
                    SQLData[0, 37] = honbuTanpin;              //本部単品 
                    SQLData[0, 38] = item1_MadoguchiShukeiHyoFolder.Text;          //集計表フォルダ
                    SQLData[0, 39] = item1_MadoguchiHoukokuShoFolder.Text;          //報告書フォルダ
                    SQLData[0, 40] = item1_MadoguchiShiryouHolder.Text;          //調査資料フォルダ
                    SQLData[0, 41] = gyoumuKanri;        //業務管理者の業務管理者CD or Null
                    SQLData[0, 42] = ankenJouhouId;                 //AnkenJouhou.AnkenJouhouID、未受託の場合はNULL
                    SQLData[0, 43] = garoon;     //MadoguchiGaroonRenkei
                    SQLData[0, 44] = ankenJouhouId;
                    SQLData[0, 45] = renban;
                }
                //更新のとき
                else
                {

                    SQLData[0, 1] = MadoguchiID;
                    SQLData[0, 2] = item1_MadoguchiTourokuNendo.SelectedValue.ToString();//登録年度
                    SQLData[0, 3] = hikiwatashiShounin;
                    SQLData[0, 4] = saishuukensa;
                    SQLData[0, 5] = GlobalMethod.ChangeSqlText(item1_MadoguchiShouninsha.Text, 0, 0);
                    SQLData[0, 6] = Get_DateTimePicker("item1_MadoguchiShouninnbi");
                    SQLData[0, 7] = Get_DateTimePicker("item1_MadoguchiShimekiribi");
                    SQLData[0, 8] = Get_DateTimePicker("item1_MadoguchiTourokubi");
                    SQLData[0, 9] = Get_DateTimePicker("item1_MadoguchiHoukokuJisshibi");
                    SQLData[0, 10] = shubetsu;
                    SQLData[0, 11] = jisshiKubun;
                    SQLData[0, 12] = item1_MadoguchiJutakuBushoCD.SelectedValue.ToString();//受託課所支部
                    SQLData[0, 13] = keiyakutantou;            //契約担当者orNULL　
                    SQLData[0, 14] = item1_MadoguchiJutakuBushoCD.SelectedValue.ToString();
                    SQLData[0, 15] = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString();//窓口部所
                    SQLData[0, 16] = madoguchiTantouCD;//窓口担当者
                    SQLData[0, 17] = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString();//窓口部所所属長の部所CD
                    SQLData[0, 18] = chousaKubunJibusho;// 調査区分　自部所
                    SQLData[0, 19] = chousaKubunShibushibu;//調査区分　支→支
                    SQLData[0, 20] = chousaKubunHonshibu;// 調査区分　本→支
                    SQLData[0, 21] = chousaKubunShibuhon;//調査区分　支→本
                    SQLData[0, 22] = item1_MadoguchiKanriBangou.Text;
                    SQLData[0, 23] = item1_MadoguchiJutakuBangou.Text;
                    SQLData[0, 24] = item1_MadoguchiJutakuBangouEdaban.Text;       //受託番号枝番
                    SQLData[0, 25] = item1_MadoguchiUketsukeBangou.Text; //特調番号
                    SQLData[0, 26] = item1_MadoguchiUketsukeBangouEdaban.Text;
                    SQLData[0, 27] = item1_MadoguchiHachuuKikanmei.Text;
                    SQLData[0, 28] = GlobalMethod.ChangeSqlText(item1_MadoguchiGyoumuMeishou.Text, 0, 0);
                    SQLData[0, 29] = GlobalMethod.ChangeSqlText(item1_MadoguchiKoujiKenmei.Text, 0, 0);
                    SQLData[0, 30] = GlobalMethod.ChangeSqlText(item1_MadoguchiChousaHinmoku.Text, 0, 0);
                    SQLData[0, 31] = GlobalMethod.ChangeSqlText(item1_MadoguchiBikou.Text, 0, 0);
                    SQLData[0, 32] = GlobalMethod.ChangeSqlText(item1_MadoguchiTankaTekiyou.Text, 0, 0);
                    SQLData[0, 33] = GlobalMethod.ChangeSqlText(item1_MadoguchiNiwatashi.Text, 0, 0);
                    SQLData[0, 34] = houkokuzumi;
                    SQLData[0, 35] = item1_KanriGijutsusha_CD.Text;
                    SQLData[0, 36] = honbuTanpin;
                    SQLData[0, 37] = item1_MadoguchiShukeiHyoFolder.Text;
                    SQLData[0, 38] = item1_MadoguchiHoukokuShoFolder.Text;
                    SQLData[0, 39] = item1_MadoguchiShiryouHolder.Text;
                    SQLData[0, 40] = gyoumuKanri; //業務管理者の業務管理者CD or Null
                    SQLData[0, 41] = ankenJouhouId;

                    //受託番号(＝案件番号,特調番号)が変わった場合
                    if (!beforeJutaku.Equals(item1_MadoguchiUketsukeBangou.Text) || !befireTokuchoEda.Equals(item1_MadoguchiUketsukeBangouEdaban.Text))
                    {
                        string renban = TokuchoNo_saiban();
                        SQLData[0, 42] = renban;
                    }
                    else
                    {
                        SQLData[0, 42] = "NULL";
                    }
                    SQLData[0, 43] = item1_MadoguchiTantousha.Text;//窓口担当者名
                    SQLData[0, 44] = garoon;     //MadoguchiGaroonRenkei

                    // 1223 受託番号変更で単価ランク連動
                    SQLData[0, 45] = item6_TanpinHachuubusho.Text;     // 部署
                    SQLData[0, 46] = item6_TanpinYakushoku.Text;       // 役職
                    SQLData[0, 47] = item6_TanpinHachuuTantousha.Text; // 発注担当者
                    SQLData[0, 48] = item6_TanpinTel.Text;             // TEL
                    SQLData[0, 49] = item6_TanpinFax.Text;             // FAX
                    SQLData[0, 50] = item6_TanpinMail.Text;            // MAIL
                    SQLData[0, 51] = jutakubangouChangeFlg; // 0:受託番号未変更 1:受託番号変更


                }

                //テーブル登録・更新
                Boolean result = GlobalMethod.MadoguchiUpdate_SQL(tab, MadoguchiID, SQLData, out mes, UserInfos);

                if ("insert".Equals(mode))
                {
                    MadoguchiID = saibanMadoguchiNo.ToString();
                    mode = "update";
                }

                // 特調番号の枝番を編集不可に
                //item1_MadoguchiUketsukeBangouEdaban.Enabled = false;
                item1_MadoguchiUketsukeBangouEdaban.ReadOnly = true;
                item1_MadoguchiUketsukeBangouEdaban.BackColor = Color.FromArgb(240, 240, 240);

                set_error("", 0);
                set_error(mes);

                jutakubangouChangeFlg = "0"; // 0:受託番号未変更 1:受託番号変更

            }
            // 担当部所
            else if (tab == 2)
            {
                ////調査担当者Gridのデータ
                //string[,] SQLData = new string[c1FlexGrid1.Rows.Count - 1, 7];
                //for (int i = 1; i < c1FlexGrid1.Rows.Count; i++)
                //{
                //    SQLData[i - 1, 0] = c1FlexGrid1.Rows[i][1].ToString();
                //    SQLData[i - 1, 1] = c1FlexGrid1.Rows[i][2].ToString();
                //    SQLData[i - 1, 2] = c1FlexGrid1.Rows[i][3].ToString();
                //    if (c1FlexGrid1.Rows[i][4] != null && c1FlexGrid1.Rows[i][4].ToString() != "")
                //    {
                //        SQLData[i - 1, 3] = c1FlexGrid1.Rows[i][4].ToString();
                //    }
                //    //SQLData[i - 1, 3] = c1FlexGrid1.Rows[i][4].ToString();
                //    SQLData[i - 1, 4] = c1FlexGrid1.Rows[i][5].ToString();
                //    if (c1FlexGrid1.GetCellCheck(i, 6) == C1.Win.C1FlexGrid.CheckEnum.Checked)
                //    {
                //        SQLData[i - 1, 5] = "1";
                //    }
                //    else
                //    {
                //        SQLData[i - 1, 5] = "0";
                //    }
                //    if (checkBox28.Checked == true)
                //    {
                //        SQLData[i - 1, 6] = "1";
                //    }
                //    else
                //    {
                //        SQLData[i - 1, 6] = "0";
                //    }
                //}

                ////Groon追加宛先Gridのデータ
                //string[,] SQLData2 = new string[c1FlexGrid5.Rows.Count - 1, 5];
                //for (int i = 1; i < c1FlexGrid5.Rows.Count; i++)
                //{
                //    if (c1FlexGrid5.Rows[i][1] != null)
                //    {
                //        SQLData2[i - 1, 0] = c1FlexGrid5.Rows[i][1].ToString();
                //    }
                //    if (c1FlexGrid5.Rows[i][2] != null)
                //    {
                //        SQLData2[i - 1, 1] = c1FlexGrid5.Rows[i][2].ToString();
                //        SQLData2[i - 1, 2] = c1FlexGrid5.GetDataDisplay(i, 2).ToString();
                //    }
                //    if (c1FlexGrid5.Rows[i][3] != null)
                //    {
                //        SQLData2[i - 1, 3] = c1FlexGrid5.Rows[i][3].ToString();
                //        SQLData2[i - 1, 4] = c1FlexGrid5.GetDataDisplay(i, 3).ToString();
                //    }
                //}

                //string mes = "";
                //GlobalMethod.MadoguchiUpdate_SQL(tab, MadoguchiID, SQLData, out mes, UserInfos, SQLData2);
                //set_error("", 0);
                //set_error(mes);

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
                GlobalMethod.MadoguchiUpdate_SQL(2, MadoguchiID, SQLData, out mes, UserInfos, SQLData2);
                set_error(mes);

            }
            // 協力依頼書
            else if (tab == 4)
            {
                //更新用データを画面から取得（協力依頼書タブ）
                string[,] SQLData = new string[1, 24];
                if (item4_KyoRyokuBusho.Text != "")
                {
                    SQLData[0, 0] = item4_KyoRyokuBusho.Text;
                }
                if (item4_Iraibi.CustomFormat == "")
                {
                    SQLData[0, 1] = item4_Iraibi.Text;
                }
                if (item4_HoukokuKigen.CustomFormat == "")
                {
                    SQLData[0, 2] = item4_HoukokuKigen.Text;
                }
                if (item4_GyoumuKubun.Text != "")
                {
                    SQLData[0, 3] = item4_GyoumuKubun.SelectedValue.ToString();
                }
                if (item4_IraiKubun.Text != "")
                {
                    SQLData[0, 4] = item4_IraiKubun.SelectedValue.ToString();
                }
                if (item4_NaiyoKubun_Shizai.Checked == true)
                {
                    SQLData[0, 5] = "1";
                }
                if (item4_NaiyoKubun_DKou.Checked == true)
                {
                    SQLData[0, 6] = "1";
                }
                if (item4_NaiyoKubun_EKou.Checked == true)
                {
                    SQLData[0, 7] = "1";
                }
                if (item4_NaiyoKubun_Sonota.Checked == true)
                {
                    SQLData[0, 8] = "1";
                }
                if (item4_NaiyoKubun_Joho.Checked == true)
                {
                    SQLData[0, 9] = "1";
                }
                SQLData[0, 10] = item4_RenrakuJikou.Text;
                SQLData[0, 11] = item4_GyoumuNaiyo.Text;
                if (item4_Zumen.Text != "")
                {
                    SQLData[0, 12] = item4_Zumen.SelectedValue.ToString();
                }
                if (item4_Kizyunbi.Text != "")
                {
                    SQLData[0, 13] = item4_Kizyunbi.SelectedValue.ToString();
                }
                SQLData[0, 14] = item4_KizyunbiStr.Text;
                if (item4_UtiawaseYouhi.Text != "")
                {
                    SQLData[0, 15] = item4_UtiawaseYouhi.SelectedValue.ToString();
                }
                if (item4_Gutaiteki.Text != "")
                {
                    SQLData[0, 16] = item4_Gutaiteki.SelectedValue.ToString();
                }
                if (item4_ZenkaiKyouryoku.Text != "")
                {
                    SQLData[0, 17] = item4_ZenkaiKyouryoku.SelectedValue.ToString();
                }
                SQLData[0, 18] = item4_ZenkaiKyouryokuStr.Text;
                if (item4_Hikiwatashi.Text != "")
                {
                    SQLData[0, 19] = item4_Hikiwatashi.SelectedValue.ToString();
                }
                if (item4_JishiKeikakusho.Text != "")
                {
                    SQLData[0, 20] = item4_JishiKeikakusho.SelectedValue.ToString();
                }
                if (item4_MitsumoriChousyu.Text != "")
                {
                    SQLData[0, 21] = item4_MitsumoriChousyu.SelectedValue.ToString();
                }
                SQLData[0, 22] = item4_KyouryokuIraishoID.Text;
                SQLData[0, 23] = item4_KyoryokuChouCD.Text;

                //更新処理
                string mes = "";
                GlobalMethod.MadoguchiUpdate_SQL(tab, MadoguchiID, SQLData, out mes, UserInfos);
                //メッセージ表示
                set_error("", 0);
                set_error(mes);
            }
            // 応援受付タブ
            else if (tab == 5)
            {
                string[,] SQLData = new string[1, 5];
                //SQLData[0, 0] = item5_Kanribangou.Text; // 管理番号は調査概要でやっているので除外
                // SQLData[0, 0]:応援状況
                // SQLData[0, 1]:応援受付日
                // SQLData[0, 2]:応援完了
                // SQLData[0, 3]:応援完了日
                if (item5_UketsukeJoukyo.Checked == true)
                {
                    //SQLData[0, 0] = "1";
                    SQLData[0, 0] = "2";
                }
                else
                {
                    string chousaKubunShibuhon = "0";

                    DataTable tmpdt = GlobalMethod.getData("MadoguchiChousaKubunShibuHonbu", "MadoguchiChousaKubunShibuHonbu", "MadoguchiJouhou", "MadoguchiID = " + MadoguchiID);
                    if(tmpdt != null && tmpdt.Rows.Count > 0)
                    {
                        chousaKubunShibuhon = tmpdt.Rows[0][0].ToString();
                    }
                    // 応援状況にチェックが入ってない
                    //if (DT_Ouenuketsuke.Rows[0][1].ToString() == "2")
                    if (chousaKubunShibuhon == "1")
                    {
                        SQLData[0, 0] = "1";
                    }
                    else
                    {
                        //SQLData[0, 0] = DT_Ouenuketsuke.Rows[0][1].ToString();
                        SQLData[0, 0] = "0";
                    }
                }
                if (item5_UketsukeDate.CustomFormat == "")
                {
                    SQLData[0, 1] = item5_UketsukeDate.Text;
                }
                if (item5_OuenKanryo.Checked == true)
                {
                    SQLData[0, 2] = "1";
                }
                if (item5_OuenKanryoDate.CustomFormat == "")
                {
                    SQLData[0, 3] = item5_OuenKanryoDate.Text;
                }
                string mes = "";
                GlobalMethod.MadoguchiUpdate_SQL(tab, MadoguchiID, SQLData, out mes, UserInfos);
                set_error("", 0);
                set_error(mes);

            }
            // 単品入力情報
            else if (tab == 6)
            {
                //string[,] SQLData = new string[1, 24];
                string[,] SQLData = new string[1, 37];

                String strSpace = " ";
                String strComma = ", ";
                String strEqual = " = ";
                string strSingleQuote = "'";
                int result = 0;

                StringBuilder sb = new StringBuilder();

                // テーブル定義（項目名、属性）
                string[,] TanpinNyuuryoku = new string[,]
                {
                         {"TanpinNyuuryokuID", "Numeric"}	            // 0.単品入力項目ID
                        ,{"TanpinJutakuDate", "Date"}	                // 1.受託日（依頼日）
                        ,{"TanpinHoukokuDate", "Date"}	                // 2.報告日
                        ,{"TanpinShiji", "String"}	                    // 3.指示番号
                        ,{"TanpinHachuubusho", "String"}	            // 4.部所
                        ,{"TanpinYakushoku", "String"}	                // 5.役職
                        ,{"TanpinHachuuTantousha", "String"}	        // 6.担当者
                        ,{"TanpinTel", "String"}	                    // 7.電話
                        ,{"TanpinFax", "String"}	                    // 8.FAX
                        ,{"TanpinMail", "String"}	                    // 9.メール
                        ,{"TanpinMemo", "String"}	                    // 10.メモ
                        ,{"TanpinRank", "String"}	                    // 11.ランク
                        ,{"TanpinShousa", "String"}	                    // 12.照査実施
                        ,{"TanpinShijisho", "Numeric"}	                // 13.指示書
                        ,{"TanpinSaishuuKensa", "Numeric"}	            // 14.最終検査
                        ,{"TanpinMitsumoriTeishutu", "Numeric"}	        // 15.見積提出方式
                        ,{"TanpinTeinyuusatsu", "Numeric"}	            // 16.低入札
                        ,{"TanpinShuyouChousain", "String"}	            // 17.主要調査員
                        ,{"TanpinSeikyuuGetsu", "String"}	            // 18.単品請求月
                        ,{"TanpinHokurikuShijouKakaku", "Numeric"}	    // 19.市場価格（北陸専用）
                        ,{"TanpinHokurikuShijouKakaku_r", "Numeric"}	// 20.市場価格（北陸専用）r
                        ,{"TanpinHokurikuSekouKanka", "Numeric"}	    // 21.施工単価（北陸専用）
                        ,{"TanpinHokurikuSekouKanka_r", "Numeric"}	    // 22.施工単価（北陸専用）r
                        ,{"TanpinSonotaShuukei", "Numeric"}	            // 23.その他集計
                        ,{"TanpinSeikyuuKingaku", "Numeric"}	        // 24.請求金額
                        ,{"TanpinSeikyuuKakutei", "Numeric"}	        // 25.請求確定
                        ,{"MadoguchiID", "Numeric"}	                    // 26.窓口ID
                        ,{"TanpinGyoumuCD", "Numeric"}	                // 27.業務CD
                        ,{"TanpinAnkenJouhouID", "Numeric"}	            // 28.契約情報ID
                        ,{"TanpinKeihi", "Numeric"}	                    // 29.経費（バックアップ用）
                        ,{"TanpinCreateDate", "Date"}	                // 30.作成日時
                        ,{"TanpinCreateUser", "String"}	                // 31.作成ユーザ
                        ,{"TanpinCreateProgram", "String"}	            // 32.作成機能
                        ,{"TanpinUpdateDate", "Date"}	                // 33.更新日時
                        ,{"TanpinUpdateUser", "String"}	                // 34.更新ユーザ
                        ,{"TanpinUpdateProgram", "String"}	            // 35.更新機能
                        ,{"TanpinDeleteFlag", "Numeric"}                // 36.削除フラグ
                };

                string[,] TanpinNyuuryokuRank = new string[,]
                {
                         {"TanpinNyuuryokuID", "Numeric"}               // 0.単品入力項目ID
                        ,{"TanpinL1RankID", "Numeric"}                  // 1.ランクID
                        ,{"TanpinL1RankMei", "String"}                  // 2.ランク名
                        ,{"TanpunL1RankKubun", "Numeric"}               // 3.ランク種別（集計方法）
                        ,{"TanpinL1Ranksuu", "Numeric"}                 // 4.依頼本数
                        ,{"TanpinL1HoukokuHonsuu", "Numeric"}           // 5.報告本数
                        ,{"TanpinL1Tanka", "Numeric"}                   // 6.単価
                        ,{"TanpinL1Kingaku", "Numeric"}                 // 7.金額
                };

                SqlConnection sqlconn = new SqlConnection(connStr);
                sqlconn.Open();
                var cmd = sqlconn.CreateCommand();
                SqlTransaction transaction = sqlconn.BeginTransaction();
                cmd.Transaction = transaction;
                var dt = new DataTable();
                try
                {
                    // SELECT文の生成
                    sb.Clear();
                    sb.Append("SELECT ");
                    for (int i = 0; i < TanpinNyuuryoku.GetLength(0); i++)
                    {
                        if (i != 0)
                        {
                            sb.Append(strComma);
                        }
                        sb.Append(TanpinNyuuryoku[i, 0]);
                    }

                    // 条件式の設定
                    sb.Append(" FROM TanpinNyuuryoku");
                    sb.Append(" WHERE ");
                    sb.Append(TanpinNyuuryoku[0, 0]);   // 単品入力項目ID
                    sb.Append(strEqual);
                    sb.Append(item6_TanpinNyuuryokuID.Text);

                    cmd.CommandText = sb.ToString();

                    Console.WriteLine(cmd.CommandText);
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                }
                catch (Exception)
                {

                    throw;
                }

                // 初期値として取得した値をセットする
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    SQLData[0, i] = "";
                    if (dt.Rows[0][i] != null && dt.Rows[0][i].ToString() != "")
                    {
                        SQLData[0, i] = dt.Rows[0][i].ToString();
                    }
                }

                SQLData[0, 0] = item6_TanpinNyuuryokuID.Text;       // 0.単品入力項目ID
                if (item6_TanpinJutakuDate.CustomFormat == "")
                {
                    SQLData[0, 1] = item6_TanpinJutakuDate.Text;    // 1.受託日（依頼日）
                }
                if (item6_TanpinHoukokuDate.CustomFormat == "")
                {
                    SQLData[0, 2] = item6_TanpinHoukokuDate.Text;   // 2.報告日
                }
                SQLData[0, 3] = item6_TanpinShiji.Text;             // 3.指示番号
                SQLData[0, 4] = item6_TanpinHachuubusho.Text;       // 4.部所
                SQLData[0, 5] = item6_TanpinYakushoku.Text;         // 5.役職
                SQLData[0, 6] = item6_TanpinHachuuTantousha.Text;   // 6.担当者
                SQLData[0, 7] = item6_TanpinTel.Text;               // 7.電話
                SQLData[0, 8] = item6_TanpinFax.Text;               // 8.FAX
                SQLData[0, 9] = item6_TanpinMail.Text;              // 9.メール
                SQLData[0, 10] = item6_TanpinMemo.Text;             // 10.メモ
                SQLData[0, 12] = item6_TanpinShousa.Text;           // 12.照査実施
                SQLData[0, 13] = item6_TanpinShijisho.SelectedValue.ToString();     // 13.指示書
                if (item6_TanpinSaishuuKensa.Checked == true)
                {
                    SQLData[0, 14] = "1";                           // 14.設計変更
                }
                else
                {
                    // えんとり君修正STEP2　不具合1412
                    SQLData[0, 14] = "0";                           // 14.設計変更
                }
                if (item6_TanpinMitsumoriTeishutu.Checked == true)
                {
                    SQLData[0, 15] = "1";                           // 15.見積提出方式
                }
                else
                {
                    // えんとり君修正STEP2　不具合1412
                    SQLData[0, 15] = "0";                           // 15.見積提出方式
                }
                if (item6_TanpinTeinyuusatsu.Checked == true)
                {
                    SQLData[0, 16] = "1";                           // 16.低入札
                }
                else
                {
                    // えんとり君修正STEP2　不具合1412
                    SQLData[0, 16] = "0";                           // 16.低入札
                }
                SQLData[0, 18] = item6_TanpinSeikyuuGetsu.Text;      // 18.単品請求月
                SQLData[0, 23] = item6_TanpinSonotaShuukei.Text;    // 23.その他集計
                SQLData[0, 24] = item6_TanpinSeikyuuKingaku.Text;   // 24.請求金額
                if (item6_TanpinSeikyuuKakutei.Checked == true)
                {
                    SQLData[0, 25] = "1";                           // 25.請求確定
                }

                string[,] SQLData2 = new string[c1FlexGrid2.Rows.Count - 1, 8];
                for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
                {
                    SQLData2[i - 1, 0] = item6_TanpinNyuuryokuID.Text;      // 0.単品入力項目ID
                    SQLData2[i - 1, 1] = c1FlexGrid2.Rows[i][6].ToString(); // 1.ランクID
                    SQLData2[i - 1, 2] = c1FlexGrid2.Rows[i][0].ToString(); // 2.ランク名
                    SQLData2[i - 1, 3] = c1FlexGrid2.Rows[i][5].ToString(); // 3.ランク種別（集計方法）
                    SQLData2[i - 1, 4] = c1FlexGrid2.Rows[i][2].ToString(); // 4.依頼本数
                    SQLData2[i - 1, 5] = c1FlexGrid2.Rows[i][1].ToString(); // 5.報告本数
                    //SQLData2[i - 1, 4] = c1FlexGrid2.Rows[i][3].ToString(); // 単価
                    //SQLData2[i - 1, 5] = c1FlexGrid2.Rows[i][4].ToString(); // 金額
                    // 単価
                    SQLData2[i - 1, 6] = "0";
                    if (c1FlexGrid2.Rows[i][3] != null && c1FlexGrid2.Rows[i][3].ToString() != "")
                    {
                        SQLData2[i - 1, 6] = c1FlexGrid2.Rows[i][3].ToString();     // 6.単価
                    }
                    // 金額
                    SQLData2[i - 1, 7] = "0";
                    if (c1FlexGrid2.Rows[i][4] != null && c1FlexGrid2.Rows[i][4].ToString() != "")
                    {
                        SQLData2[i - 1, 7] = c1FlexGrid2.Rows[i][4].ToString();     // 7.金額
                    }

                }
                string mes = "";
                item6_TanpinTel.BackColor = Color.FromArgb(255, 255, 255);
                item6_TanpinFax.BackColor = Color.FromArgb(255, 255, 255);
                item6_TanpinMail.BackColor = Color.FromArgb(255, 255, 255);
                if (GlobalMethod.MadoguchiUpdate_ErrorCheck(tab, SQLData, out string[] ErrorMes))
                {
                    GlobalMethod.MadoguchiUpdate_SQL(tab, MadoguchiID, SQLData, out mes, UserInfos, SQLData2);
                }
                else
                {
                    for (int i = 0; i < ErrorMes.Length; i++)
                    {
                        if (ErrorMes[i] != null && ErrorMes[i] != "")
                        {
                            if (mes != "")
                            {
                                mes += Environment.NewLine;
                            }
                            mes += ErrorMes[i];

                            if (i == 0)
                            {
                                item6_TanpinTel.BackColor = Color.FromArgb(255, 204, 255);
                            }
                            else if (i == 1)
                            {
                                item6_TanpinFax.BackColor = Color.FromArgb(255, 204, 255);

                            }
                            else if (i == 2)
                            {
                                item6_TanpinMail.BackColor = Color.FromArgb(255, 204, 255);
                            }
                            UpdateFlag = false;
                        }
                    }
                }
                set_error("", 0);
                set_error(mes);
            }
            //施工更新
            else if (tab == 7)
            {
                //メッセージクリア
                set_error("", 0);

                //メッセージ変数
                string mes = "";

                //必須チェック OKだったら
                if (registration_required(7))
                {
                    //登録データ配列
                    string[,] SQLData = new string[1, 65];


                    //データセット
                    //sekouMode 明示書ID 工種名
                    SQLData[0, 0] = sekouMode;
                    SQLData[0, 1] = item7_SekouJoukenMeijishoID.Text;
                    SQLData[0, 2] = item7_KoushuMei.Text;

                    //◆施工条件旧
                    //①施工計画書添付の有無 　
                    SQLData[0, 3] = str_bool(checkBox80.Checked);
                    //②その他添付資料の現場平面図　②その他添付資料の土質関係図　②その他添付資料の数量計算書
                    SQLData[0, 4] = str_bool(checkBox81.Checked);
                    SQLData[0, 5] = str_bool(checkBox82.Checked);
                    SQLData[0, 6] = str_bool(checkBox83.Checked);
                    //③施工時間帯指定の昼間　③施工時間帯指定の夜間　③施工時間帯指定の規制有り
                    SQLData[0, 7] = str_bool(checkBox84.Checked);
                    SQLData[0, 8] = str_bool(checkBox85.Checked);
                    SQLData[0, 9] = str_bool(checkBox86.Checked);
                    //④施工条件他の作業効率　④施工条件他の施工機械の搬入経路
                    SQLData[0, 10] = str_bool(checkBox87.Checked);
                    SQLData[0, 11] = str_bool(checkBox88.Checked);
                    //④施工条件他の仮設条件　④施工条件他の資材搬入
                    SQLData[0, 12] = str_bool(checkBox89.Checked);
                    SQLData[0, 13] = str_bool(checkBox93.Checked);
                    //⑤建設機械スペック指定　⑥水中施行条件　⑦その他
                    SQLData[0, 14] = str_bool(checkBox90.Checked);
                    SQLData[0, 15] = str_bool(checkBox91.Checked);
                    SQLData[0, 16] = str_bool(checkBox92.Checked);
                    //メモ1　メモ2
                    SQLData[0, 17] = textBox41.Text;
                    SQLData[0, 18] = textBox42.Text;

                    //窓口ID システム日付 PG名　SekouDeleteFlag
                    SQLData[0, 19] = MadoguchiID;
                    SQLData[0, 20] = "SYSDATETIME()";
                    SQLData[0, 21] = "MadoguchiUpdate_SQL.SekouJouken";
                    SQLData[0, 22] = "0";

                    //◆施工条件
                    //3.添付資料の位置図 3.添付資料の施工計画書 3.添付資料の参考カタログ
                    SQLData[0, 23] = str_bool(checkBox43.Checked);
                    SQLData[0, 24] = str_bool(checkBox47.Checked);
                    SQLData[0, 25] = str_bool(checkBox51.Checked);

                    //3.添付資料の一般図・平面図 3.添付資料の現場写真 3.添付資料の過去報告書
                    SQLData[0, 26] = str_bool(checkBox44.Checked);
                    SQLData[0, 27] = str_bool(checkBox48.Checked);
                    SQLData[0, 28] = str_bool(checkBox52.Checked);

                    //3.添付資料の詳細図 3.添付資料の土質関係図（柱状図等）3.添付資料のその他
                    SQLData[0, 29] = str_bool(checkBox45.Checked);
                    SQLData[0, 30] = str_bool(checkBox49.Checked);
                    SQLData[0, 31] = str_bool(checkBox53.Checked);

                    //3.添付資料の数量計算書 3.添付資料の運搬ルート図
                    SQLData[0, 32] = str_bool(checkBox46.Checked);
                    SQLData[0, 33] = str_bool(checkBox50.Checked);

                    //5.(1)施工場所の陸上 5.(1)施工場所の水上 
                    SQLData[0, 34] = str_bool(checkBox54.Checked);
                    SQLData[0, 35] = str_bool(checkBox55.Checked);

                    //5.(1)施工場所の水中 5.(1)施工場所のその他
                    SQLData[0, 36] = str_bool(checkBox56.Checked);
                    SQLData[0, 37] = str_bool(checkBox57.Checked);

                    //5.(2)施工時間帯の通常昼間施工（8:00~17:00） 5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                    SQLData[0, 38] = str_bool(checkBox58.Checked);
                    SQLData[0, 39] = str_bool(checkBox60.Checked);

                    //5.(2)施工時間帯の施工時間規制あり 5.(2)施工時間帯の二方施工（2交代制 昼夜連続施工）
                    SQLData[0, 40] = str_bool(checkBox62.Checked);
                    SQLData[0, 41] = str_bool(checkBox59.Checked);

                    //5.(2)施工時間帯の三方施工（3交代制 24時間施工）
                    SQLData[0, 42] = str_bool(checkBox61.Checked);

                    //5.(3)作業環境の現場が狭隘  5.(3)作業環境の施工箇所が点在 5.(3)作業環境の上空制限あり
                    SQLData[0, 43] = str_bool(checkBox63.Checked);
                    SQLData[0, 44] = str_bool(checkBox67.Checked);
                    SQLData[0, 45] = str_bool(checkBox64.Checked);

                    //5.(3)作業環境のその他 5.(3)作業環境の人家に近接（近接施工） 5.(3)作業環境の特記すべき条件なし
                    SQLData[0, 46] = str_bool(checkBox68.Checked);
                    SQLData[0, 47] = str_bool(checkBox65.Checked);
                    SQLData[0, 48] = str_bool(checkBox70.Checked);

                    //5.(3)作業環境の環境対策あり（騒音・振動）
                    SQLData[0, 49] = str_bool(checkBox66.Checked);

                    //5.(4)施工機械・資材搬入経路の交通規制あり 5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                    SQLData[0, 50] = str_bool(checkBox69.Checked);
                    SQLData[0, 51] = str_bool(checkBox71.Checked);

                    //5.(4)施工機械・資材搬入経路のその他 5.(4)施工機械・資材搬入経路の特記すべき条件なし
                    SQLData[0, 52] = str_bool(checkBox72.Checked);
                    SQLData[0, 53] = str_bool(checkBox73.Checked);

                    //5.(5)仮設条件の指定あり 5.(5)仮設条件の特記すべき条件なし 
                    SQLData[0, 54] = str_bool(checkBox74.Checked);
                    SQLData[0, 55] = str_bool(checkBox75.Checked);

                    //5.(6)施工機械スペック指定の指定あり 5.(6)施工機械スペック指定の指定なし 
                    SQLData[0, 56] = str_bool(checkBox76.Checked);
                    SQLData[0, 57] = str_bool(checkBox77.Checked);

                    //5.(7)その他条件の指定あり  5.(7)その他条件の特記すべき条件なし
                    SQLData[0, 58] = str_bool(checkBox78.Checked);
                    SQLData[0, 59] = str_bool(checkBox79.Checked);

                    //メモ
                    SQLData[0, 60] = textBox40.Text;

                    //GlobalMethodの更新メソッド
                    GlobalMethod.MadoguchiUpdate_SQL(tab, MadoguchiID, SQLData, out mes, UserInfos);

                    //メッセージ表示
                    //施工モードが新規のとき
                    if ("0".Equals(sekouMode))
                    {
                        // 採番されたSekouJoukenIDを取得する
                        String discript = "SekouJoukenID ";
                        String value = "SekouJoukenID ";
                        String table = "SekouJouken";
                        String where = "MadoguchiID = '" + MadoguchiID + "' AND SekouJoukenMeijishoID COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + GlobalMethod.ChangeSqlText(item7_SekouJoukenMeijishoID.Text, 0, 0) + "' ";
                        //コンボボックスデータ取得
                        DataTable tmpdt = GlobalMethod.getData(discript, value, table, where);

                        if (tmpdt != null && tmpdt.Rows.Count > 0) 
                        { 
                            SekouJoukenID = tmpdt.Rows[0][0].ToString();
                        }
                    }
                    //更新のとき
                    else
                    {

                    }

                    //メッセージセット
                    set_error(mes);

                }//registration_required end
            }
            if (UpdateFlag)
            {
                get_data(tab);
            }
        }

        // 協力依頼書タブ 過去の依頼書を参照し登録ボタン
        private void button_pop_kakoirai_Click(object sender, EventArgs e)
        {
            // エラークリア
            ErrorClear_KyouryokuIrai();

            Popup_Anken form = new Popup_Anken();
            form.mode = "kakoirai";
            /*
            String discript = "NendoSeireki";
            String value = "NendoID ";
            String table = "Mst_Nendo";
            String where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
            DataTable dt = GlobalMethod.getData(discript, value, table, where);
            if (dt != null)
            {
                form.nendo = dt.Rows[0][0].ToString();
            }
            else
            {
                form.nendo = DateTime.Today.Year.ToString();
            }
            */
            form.nendo = GlobalMethod.GetTodayNendo();
            if (item1_MadoguchiJutakuBangouEdaban.Text != "")
            {
                form.jutakuBangou = item1_MadoguchiJutakuBangou.Text + "-" + item1_MadoguchiJutakuBangouEdaban.Text;
            }
            else
            {
                form.jutakuBangou = item1_MadoguchiJutakuBangou.Text;
            }

            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                KakoIraiID = form.ReturnValue[0];
                get_data(4);
            }
        }

        public void Resize_Grid(string name)
        {
            Control[] cs;
            cs = this.Controls.Find(name, true);
            if (cs.Length > 0)
            {
                var fx = (C1.Win.C1FlexGrid.C1FlexGrid)cs[0];

                int rowcount = fx.Rows.Count;

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

        // 受託番号を検索ボタン
        private void button1_Click(object sender, EventArgs e)
        {
            //受託番号プロンプト
            Popup_miharujutaku form = new Popup_miharujutaku();
            //年度が入っていたら年度を渡す
            //form.nendo = item1_3.SelectedValue.ToString();
            form.Nendo = item1_MadoguchiTourokuNendo.SelectedValue.ToString();

            //受託部所を渡す
            //if (item1_2.SelectedValue != null) { 
            //    form.Busho = item1_2.SelectedValue.ToString();
            //}
            // 窓口部所を渡す
            if (item1_MadoguchiTantoushaBushoCD.SelectedValue != null)
            {
                form.Busho = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString();
            }

            form.ShowDialog();

            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                //案件番号
                //item1_MadoguchiUketsukeBangou.Text = form.ReturnValue[0];
                // 2021/10/20
                // 移行データの場合、案件番号にXが付く場合があり、受託番号と一致しないようので、
                // 受託番号から-付き枝番をreplaceする方法に切り替える
                string tokucho = form.ReturnValue[1];
                tokucho = tokucho.Replace("-" + form.ReturnValue[2],"");
                item1_MadoguchiUketsukeBangou.Text = tokucho;
                //受託番号　受託枝番号
                item1_MadoguchiJutakuBangou.Text = form.ReturnValue[1];
                item1_MadoguchiJutakuBangouEdaban.Text = form.ReturnValue[2];
                //契約担当者　契約担当者CD
                item1_AnkenTantoushaMei.Text = form.ReturnValue[3];
                item1_AnkenTantoushaMei_CD.Text = form.ReturnValue[4];
                //業務管理者　業務管理者CD
                item1_MadoguchiGyoumuKanrisha.Text = form.ReturnValue[5];
                item1_item1_MadoguchiGyoumuKanrishaCD.Text = form.ReturnValue[6];
                //業務担当者　業務担当者CD
                item1_GyoumuKanrishaMei.Text = form.ReturnValue[5];
                item1_GyoumuKanrishaMei_CD.Text = form.ReturnValue[6];
                //案件情報ID
                item1_MadoguchiAnkenJouhouID.Text = form.ReturnValue[7];
                //管理技術者　管理技術者CD
                item1_KanriGijutsushaNM.Text = form.ReturnValue[8];
                item1_KanriGijutsusha_CD.Text = form.ReturnValue[9];
                //契約区分　受託部所　受託部所所属長
                item1_AnkenGyoumuKubun.SelectedValue = form.ReturnValue[10];
                item1_MadoguchiJutakuBushoCD.SelectedValue = form.ReturnValue[11];
                item1_JutakuBushoShozokuChou.Text = form.ReturnValue[12];
                //発注者課名　業務名称
                item1_MadoguchiHachuuKikanmei.Text = form.ReturnValue[13];
                item1_MadoguchiGyoumuMeishou.Text = form.ReturnValue[14];
                // 調査部で、新規の場合のみ各フォルダーに設定を行う
                if (form.ReturnValue[15] != null && form.ReturnValue[15] != "" && mode == "insert") {
                    item1_MadoguchiShukeiHyoFolder.Text = form.ReturnValue[15] + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "13"); // 集計表
                    item1_MadoguchiHoukokuShoFolder.Text = form.ReturnValue[15] + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "14"); // 報告書
                    item1_MadoguchiShiryouHolder.Text = form.ReturnValue[15] + @"\" + GlobalMethod.GetCommonValue1("ANKEN_BANGOU_FOLDER", "15"); // 調査資料・図面
                    FolderPathCheck();
                }
                //窓口担当ＣＤ　名
                item1_MadoguchiTantoushaCD.Text = form.ReturnValue[16];
                item1_MadoguchiTantousha.Text = form.ReturnValue[17];
                //窓口部所ＣＤ
                if (!String.IsNullOrEmpty(form.ReturnValue[18]))
                {
                    item1_MadoguchiTantoushaBushoCD.SelectedValue = form.ReturnValue[18];
                    //comboBox2.Text = form.ReturnValue[19];
                }
                else
                {
                    item1_MadoguchiTantoushaBushoCD.SelectedIndex = 0;
                }

                // 1223 部署、役職、担当者名、TEL、FAX、MAILが空の場合（単価契約と紐づいていない場合）
                // 単品入力項目への反映はしない
                if (!(form.ReturnValue[20] == "" && form.ReturnValue[21] == "" && form.ReturnValue[22] == "" 
                    && form.ReturnValue[23] == "" && form.ReturnValue[24] == "" && form.ReturnValue[25] == ""))
                {
                    //部署 役職　担当者名
                    //label77.Text = form.ReturnValue[20];
                    item6_TanpinHachuubusho.Text = form.ReturnValue[20];
                    item6_TanpinYakushoku.Text = form.ReturnValue[21];
                    item6_TanpinHachuuTantousha.Text = form.ReturnValue[22];
                    //TEL　FAX　MEIL
                    item6_TanpinTel.Text = form.ReturnValue[23];
                    item6_TanpinFax.Text = form.ReturnValue[24];
                    item6_TanpinMail.Text = form.ReturnValue[25];
                }
                item1_MadoguchiUketsukeBangouEdaban.Text = form.ReturnValue[26]; // 発注機関・受付番号をセットする

                jutakubangouChangeFlg = "1"; // 0:受託番号未変更 1:受託番号変更

                //textBox14.Text = form.gyoumumei;
                //textBox12.Text = form.jimusyo + " " + form.busyo + " " + form.tantousya;

                // 受託番号選択時の特調番号枝番のデフォルトは01とする
                //item1_MadoguchiUketsukeBangouEdaban.Text = "01";
            }
            item1_MadoguchiUketsukeBangouEdaban.Focus();
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
                if (ChousaHinmokuMode == 1) {
                    //if (hti.Column == 40 || hti.Column == 42 || hti.Column == 44)
                    //if (hti.Column == 42 || hti.Column == 44 || hti.Column == 46)
                    if (ColName == "HinmokuRyakuBushoCD" || ColName == "HinmokuRyakuBushoFuku1CD" || ColName == "HinmokuRyakuBushoFuku2CD")
                    {
                        // 部所
                        contextMenuStrip1.Items.Add(contextMenuBusho);
                        contextMenuStrip1.Items.Add(contextMenuBushoClear);
                    }
                    //else if (hti.Column == 41 || hti.Column == 43 || hti.Column == 45)
                    //else if (hti.Column == 43 || hti.Column == 45 || hti.Column == 47)
                    else if (ColName == "HinmokuChousainCD" || ColName == "HinmokuFukuChousainCD1" || ColName == "HinmokuFukuChousainCD2")
                    {
                        String LeftColName = c1FlexGrid4.Cols[hti.Column -1].Name;

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
                            if (item1_MadoguchiTourokuNendo.Text == "")
                            {
                                Nendo = DateTime.Today.Year;
                                ToNendo = DateTime.Today.AddYears(1).Year;
                            }
                            else
                            {
                                int.TryParse(item1_MadoguchiTourokuNendo.SelectedValue.ToString(), out Nendo);
                                ToNendo = Nendo + 1;
                            }

                            //// 部所は全部出す
                            //cmd.CommandText = "SELECT " +
                            //    "GyoumuBushoCD  " +
                            //    ",BushokanriboKameiRaku  " +
                            //    "FROM Mst_Busho  " +
                            //    "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                            //    " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                            //    " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                            //    "ORDER BY BushoMadoguchiNarabijun";

                            //bushoQuery = "SELECT " +
                            //    "GyoumuBushoCD  " +
                            //    "FROM Mst_Busho  " +
                            //    "WHERE BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                            //    " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                            //    " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) ";

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
                            //if (c1FlexGrid4.Rows[hti.Row][hti.Column - 1] != null && c1FlexGrid4.Rows[hti.Row][hti.Column - 1].ToString() != "")
                            if (c1FlexGrid4.Rows[hti.Row][LeftColName] != null && c1FlexGrid4.Rows[hti.Row][LeftColName].ToString() != "")
                            {
                                //// 空でない場合
                                //cmd.CommandText = "SELECT " +
                                //    "KojinCD " +
                                //    ",ChousainMei " +
                                //    "FROM Mst_Chousain " +
                                //    "WHERE RetireFLG = 0 AND TokuchoFLG = 1 " +
                                //    "AND GyoumuBushoCD = '" + c1FlexGrid4.Rows[hti.Row][hti.Column - 1].ToString() + "' " +
                                //    "ORDER BY ChousainMei ";
                                // 空でない場合
                                cmd.CommandText = "SELECT " +
                                    "KojinCD " +
                                    ",ChousainMei " +
                                    "FROM Mst_Chousain " +
                                    "WHERE RetireFLG = 0 AND TokuchoFLG = 1 " +
                                    //"AND GyoumuBushoCD = '" + c1FlexGrid4.Rows[hti.Row][hti.Column - 1].ToString() + "' ";
                                    "AND GyoumuBushoCD = '" + c1FlexGrid4.Rows[hti.Row][LeftColName].ToString() + "' ";

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
                                //// 空の場合
                                //cmd.CommandText = "SELECT " +
                                //    "KojinCD " +
                                //    ",ChousainMei " +
                                //    "FROM Mst_Chousain " +
                                //    "WHERE RetireFLG = 0 AND TokuchoFLG = 1 " +
                                //    "AND GyoumuBushoCD in (" + bushoQuery + ") " +
                                //    "ORDER BY ChousainMei ";
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
                    // 編集状態が1:編集でない場合
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
                                    //// 締切日
                                    ////if (c1FlexGrid4[hti.Row, 52] != null) {
                                    //if (c1FlexGrid4.Rows[hti.Row]["ChousaHinmokuShimekiribi"] != null) {
                                    //    //dateTime = DateTime.Parse(c1FlexGrid4[hti.Row, 52].ToString());
                                    //    dateTime = DateTime.Parse(c1FlexGrid4.Rows[hti.Row]["ChousaHinmokuShimekiribi"].ToString());
                                    //    if (dateTime < DateTime.Today)
                                    //    {
                                    //        // 締切日経過
                                    //        //c1FlexGrid4.Rows[hti.Row + 1][5] = "1";
                                    //        c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "1";
                                    //    }
                                    //    else if (dateTime < DateTime.Today.AddDays(3))
                                    //    {
                                    //        // 締切日が3日以内、かつ2次検証が完了していない
                                    //        //c1FlexGrid4.Rows[hti.Row + 1][5] = "2";
                                    //        c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "2";
                                    //    }
                                    //    else if (dateTime < DateTime.Today.AddDays(7))
                                    //    {
                                    //        // 締切日が1週間以内、かつ2次検証が完了していない
                                    //        //c1FlexGrid4.Rows[hti.Row + 1][5] = "3";
                                    //        c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "3";
                                    //    }
                                    //    else
                                    //    {
                                    //        //c1FlexGrid4.Rows[hti.Row + 1][5] = "4";
                                    //        c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "4";
                                    //    }
                                    //}
                                    // 1:報告済みの場合
                                    if ("1".Equals(MadoguchiHoukokuzumi))
                                    {
                                        // 報告済み
                                        c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "8";
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
                                        if (c1FlexGrid4.Rows[hti.Row + 1]["ChousaChuushi"] != null && "True".Equals(c1FlexGrid4.Rows[hti.Row + 1]["ChousaChuushi"].ToString()))
                                        {
                                            c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "6";
                                        }
                                        else if ("80".Equals(c1FlexGrid4.Rows[hti.Row + 1]["ChousaShinchokuJoukyou"].ToString()))
                                        {
                                            // 二次検証済み、または中止（中止）
                                            c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "6";
                                        }
                                        else if ("70".Equals(c1FlexGrid4.Rows[hti.Row + 1]["ChousaShinchokuJoukyou"].ToString()))
                                        {
                                            // 二次検証済み、または中止（二次検証済み）
                                            c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "5";
                                        }
                                        else if ("50".Equals(c1FlexGrid4.Rows[hti.Row + 1]["ChousaShinchokuJoukyou"].ToString()) || "60".Equals(c1FlexGrid4.Rows[hti.Row + 1]["ChousaShinchokuJoukyou"].ToString()))
                                        {
                                            // 担当者済み or 一次検済
                                            c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "7";
                                        }
                                        else if (c1FlexGrid4.Rows[hti.Row + 1]["ChousaHinmokuShimekiribi"] != null)
                                        {
                                            try
                                            {
                                                dateTime = DateTime.Parse(c1FlexGrid4.Rows[hti.Row + 1]["ChousaHinmokuShimekiribi"].ToString());
                                                if (dateTime < DateTime.Today)
                                                {
                                                    // 締切日経過
                                                    c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "1";
                                                }
                                                else if (dateTime < DateTime.Today.AddDays(3))
                                                {
                                                    // 締切日が3日以内、かつ2次検証が完了していない
                                                    c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "2";
                                                }
                                                else if (dateTime < DateTime.Today.AddDays(7))
                                                {
                                                    // 締切日が1週間以内、かつ2次検証が完了していない
                                                    c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "3";
                                                }
                                                else
                                                {
                                                    c1FlexGrid4.Rows[hti.Row + 1]["ShinchokuIcon"] = "4";
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
                // 奉行エクセル移管対応
                // 作業フォルダアイコン
                if (ColName == "SagyoForuda")
                {
                    if (c1FlexGrid4.Rows[hti.Row]["SagyoForuda"] != null)
                    {
                        switch (c1FlexGrid4.Rows[hti.Row]["SagyoForuda"].ToString())
                        {
                            // アイコン 0:グレー 1:イエロー
                            case "1":
                                // 作業フォルダが存在すれば開く
                                if (Directory.Exists(c1FlexGrid4[hti.Row, hti.Column + 1].ToString()))
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
            //if (hti.Row > 1 && hti.Column == 40)
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

        private void testEvent(object sender, EventArgs e)
        {
            Popup_Bunrui form = new Popup_Bunrui();
            form.ShowDialog();
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
            item1_MadoguchiJutakuBushoCD.DataSource = tmpdt;
            item1_MadoguchiJutakuBushoCD.DisplayMember = "Discript";
            item1_MadoguchiJutakuBushoCD.ValueMember = "Value";
            //初期表示は値なし
            item1_MadoguchiJutakuBushoCD.SelectedValue = "";

            SortedList sl = new SortedList();
            //sl = GlobalMethod.Get_SortedList(tmpdt);
            //c1FlexGrid1.Cols[2].DataMap = sl;
            //c1FlexGrid5.Cols[2].DataMap = sl;//担当部所タブGaroon追加宛先Grid担当部所
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
            item1_MadoguchiTantoushaBushoCD.DataSource = tmpdt2;
            item1_MadoguchiTantoushaBushoCD.DisplayMember = "Discript";
            item1_MadoguchiTantoushaBushoCD.ValueMember = "Value";
            //ユーザの部所セット
            item1_MadoguchiTantoushaBushoCD.SelectedValue = UserInfos[2];

            //連絡票出力
            //SQL変数
            discript = "PrintName ";
            value = "PrintFileName ";
            table = "Mst_PrintList ";
            where = "MENU_ID = '107' AND PrintBunruiCD = '2' AND PrintDelFlg = '0' ";
            //コンボボックスデータ取得
            DataTable tmpdt3 = GlobalMethod.getData(discript, value, table, where);
            //comboBox8.DataSource = tmpdt3;
            //comboBox8.DisplayMember = "Discript";
            //comboBox8.ValueMember = "Value";

            //登録年度
            discript = "NendoSeireki";
            value = "NendoID";
            table = "Mst_Nendo";
            where = "";
            //コンボボックスデータ取得
            DataTable tmpdt4 = GlobalMethod.getData(discript, value, table, where);
            item1_MadoguchiTourokuNendo.DataSource = tmpdt4;
            item1_MadoguchiTourokuNendo.DisplayMember = "Discript";
            item1_MadoguchiTourokuNendo.ValueMember = "Value";

            //新規登録の時は今年を入れる
            if ("insert".Equals(mode))
            {
                /*
                discript = "NendoSeireki";
                value = "NendoID";
                table = "Mst_Nendo";
                where = "Nendo_Sdate <= GETDATE() AND Nendo_EDate >= GETDATE()";
                //コンボボックスデータ取得
                DataTable dtYear = GlobalMethod.getData(discript, value, table, where);
                */
                //item1_MadoguchiTourokuNendo.SelectedValue = dtYear.Rows[0][0].ToString();
                item1_MadoguchiTourokuNendo.SelectedValue = GlobalMethod.GetTodayNendo();
            }

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
            item1_AnkenGyoumuKubun.DataSource = tmpdt5;
            item1_AnkenGyoumuKubun.DisplayMember = "Discript";
            item1_AnkenGyoumuKubun.ValueMember = "Value";

            //実施区分
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "実施");
            tmpdt.Rows.Add(2, "打診中");
            tmpdt.Rows.Add(3, "中止");
            item1_MadoguchiJiishiKubun.DataSource = tmpdt;
            item1_MadoguchiJiishiKubun.DisplayMember = "Discript";
            item1_MadoguchiJiishiKubun.ValueMember = "Value";

            //調査種別
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "単品");
            tmpdt.Rows.Add(2, "一般");
            tmpdt.Rows.Add(3, "単契");
            item1_MadoguchiChousaShubetsu.DataSource = tmpdt;
            item1_MadoguchiChousaShubetsu.DisplayMember = "Discript";
            item1_MadoguchiChousaShubetsu.ValueMember = "Value";

            //実施区分
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "実施");
            tmpdt.Rows.Add(2, "打診中");
            tmpdt.Rows.Add(3, "中止");
            item1_MadoguchiJiishiKubun.DataSource = tmpdt;
            item1_MadoguchiJiishiKubun.DisplayMember = "Discript";
            item1_MadoguchiJiishiKubun.ValueMember = "Value";

            //担当部所　進捗状況
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

            //協力部所　担当者状況
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

            //調査品目　材工
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(0, "全て");
            tmpdt.Rows.Add(1, "材のみ");
            src_Zaikou.DataSource = tmpdt;
            src_Zaikou.DisplayMember = "Discript";
            src_Zaikou.ValueMember = "Value";

            //奉行エクセル
            //グループ名
            discript = "MadoguchiGroupMei ";
            value = "MadoguchiGroupMasterID ";
            table = "MadoguchiGroupMaster ";
            //検証中
            //where = "MadoguchiID = " + MadoguchiID ; //MadoguchiIDが一致するもの
            where = "MadoguchiID = " + MadoguchiID + "ORDER BY MadoguchiGroupMei ";
            //コンボボックスデータ取得
            DataTable tmpdt22 = GlobalMethod.getData(discript, value, table, where);
            //1574
            ListDictionary ld = new ListDictionary();
            ld = GlobalMethod.Get_ListDictionary(tmpdt22);
            c1FlexGrid4.Cols["GroupMei"].DataMap = ld;

            //調査品目　調査主副コンボ
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));

            tmpdt.Rows.Add(0, "主+副");
            tmpdt.Rows.Add(1, "主");
            tmpdt.Rows.Add(2, "副");
            src_ShuFuku.DataSource = tmpdt;
            src_ShuFuku.DisplayMember = "Discript";
            src_ShuFuku.ValueMember = "Value";

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
            src_Zaikou.DataSource = tmpdt;
            src_Zaikou.DisplayMember = "Discript";
            src_Zaikou.ValueMember = "Value";

            //調査品目　担当者空白リスト
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "全て");
            tmpdt.Rows.Add(1, "担当者が空白のリスト");
            tmpdt.Rows.Add(2, "担当者が設定済のリスト");
            src_TantoushaKuuhaku.DataSource = tmpdt;
            src_TantoushaKuuhaku.DisplayMember = "Discript";
            src_TantoushaKuuhaku.ValueMember = "Value";

            //調査品目 進捗アイコン
            imgMap = new Hashtable();
            imgMap.Add("8", Image.FromFile("Resource/Image/shin_ao.png"));     // 報告済み
            imgMap.Add("5", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（二次検証済み）
            //imgMap.Add("6", Image.FromFile("Resource/Image/greenT1.png"));     // 二次検証済み、または中止（中止）
            imgMap.Add("6", Image.FromFile("Resource/Image/shin_ao.png"));     // 中止
            imgMap.Add("7", Image.FromFile("Resource/Image/shin_midori.png")); // 担当者済み
            imgMap.Add("1", Image.FromFile("Resource/Image/shin_dokuro.png")); // 締切日経過
            imgMap.Add("2", Image.FromFile("Resource/Image/shin_aka.png"));    // 締切日が3日以内、かつ2次検証が完了していない
            imgMap.Add("3", Image.FromFile("Resource/Image/shin_kiiro.png"));  // 締切日が1週間以内、かつ2次検証が完了していない
            imgMap.Add("4", Image.FromFile("Resource/Image/blank2.png"));      // 上記のいずれにも該当しない
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

            //tmpdt.Rows.Add(0, "-");
            tmpdt.Rows.Add(1, "-");
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
            // 奉行エクセル移管対応
            c1FlexGrid4.Cols["SagyoForuda"].ImageAlign = C1.Win.C1FlexGrid.ImageAlignEnum.CenterCenter;
            c1FlexGrid4.Cols["SagyoForuda"].ImageMap = imgMap;
            c1FlexGrid4.Cols["SagyoForuda"].ImageAndText = false;


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

            //協力依頼書　業務区分
            //discript = "KeiyakuKeitai ";
            //value = "KeiyakuKeitaiNarabijun ";
            //table = "Mst_KeiyakuKeitai ";
            //where = "";
            //tmpdt = new DataTable();
            //tmpdt = GlobalMethod.getData(discript, value, table, where);
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "一般受託調査");
            tmpdt.Rows.Add(2, "単価契約調査");
            tmpdt.Rows.Add(3, "単品契約調査");
            tmpdt.Rows.Add(4, "単価契約を含む一般受託");
            tmpdt.Rows.Add(5, "情報開発受託業務");
            item4_GyoumuKubun.DataSource = tmpdt;
            item4_GyoumuKubun.DisplayMember = "Discript";
            item4_GyoumuKubun.ValueMember = "Value";

            //協力依頼書　依頼区分
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "新規(契約前)");
            tmpdt.Rows.Add(2, "新規(契約後)");
            tmpdt.Rows.Add(3, "契約変更");
            tmpdt.Rows.Add(4, "繰り返し協力依頼");
            item4_IraiKubun.DataSource = tmpdt;
            item4_IraiKubun.DisplayMember = "Discript";
            item4_IraiKubun.ValueMember = "Value";

            //協力依頼書　図面/前回協力
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "有");
            tmpdt.Rows.Add(2, "無");
            item4_Zumen.DataSource = tmpdt;
            item4_Zumen.DisplayMember = "Discript";
            item4_Zumen.ValueMember = "Value";

            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "有");
            tmpdt.Rows.Add(2, "無");
            item4_ZenkaiKyouryoku.DataSource = tmpdt;
            item4_ZenkaiKyouryoku.DisplayMember = "Discript";
            item4_ZenkaiKyouryoku.ValueMember = "Value";

            //協力依頼書　調査基準日
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "建設物価");
            tmpdt.Rows.Add(2, "その他");
            item4_Kizyunbi.DataSource = tmpdt;
            item4_Kizyunbi.DisplayMember = "Discript";
            item4_Kizyunbi.ValueMember = "Value";

            //協力依頼書　具体的
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "済");
            item4_Gutaiteki.DataSource = tmpdt;
            item4_Gutaiteki.DisplayMember = "Discript";
            item4_Gutaiteki.ValueMember = "Value";

            //協力依頼書　打合せ要否
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "要");
            tmpdt.Rows.Add(2, "否");
            item4_UtiawaseYouhi.DataSource = tmpdt;
            item4_UtiawaseYouhi.DisplayMember = "Discript";
            item4_UtiawaseYouhi.ValueMember = "Value";

            //協力依頼書　成果物
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "協力先");
            tmpdt.Rows.Add(2, "受託元");
            item4_Hikiwatashi.DataSource = tmpdt;
            item4_Hikiwatashi.DisplayMember = "Discript";
            item4_Hikiwatashi.ValueMember = "Value";

            //協力依頼書　実施計画書
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(1, "受託元で作成");
            tmpdt.Rows.Add(2, "協力先で作成");
            tmpdt.Rows.Add(3, "両部所で作成");
            item4_JishiKeikakusho.DataSource = tmpdt;
            item4_JishiKeikakusho.DisplayMember = "Discript";
            item4_JishiKeikakusho.ValueMember = "Value";

            //協力依頼書　見積徴収
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "");
            tmpdt.Rows.Add(1, "協力先");
            tmpdt.Rows.Add(2, "受託元");
            tmpdt.Rows.Add(3, "両部所");
            item4_MitsumoriChousyu.DataSource = tmpdt;
            item4_MitsumoriChousyu.DisplayMember = "Discript";
            item4_MitsumoriChousyu.ValueMember = "Value";

            //単品　指示書
            tmpdt = new System.Data.DataTable();
            tmpdt.Columns.Add("Value", typeof(int));
            tmpdt.Columns.Add("Discript", typeof(string));
            tmpdt.Rows.Add(0, "無");
            tmpdt.Rows.Add(1, "有");
            item6_TanpinShijisho.DataSource = tmpdt;
            item6_TanpinShijisho.DisplayMember = "Discript";
            item6_TanpinShijisho.ValueMember = "Value";

            //新規登録以外
            if (!"insert".Equals(mode))
            {
                //施工条件　明示書切替
                discript = "SekouJoukenMeijishoID ";
                value = "SekouJoukenID ";
                table = "SekouJouken ";
                where = "MadoguchiID = " + MadoguchiID + " AND SekouDeleteFlag != 1 ";
                tmpdt = new DataTable();
                tmpdt = GlobalMethod.getData(discript, value, table, where);
                if (tmpdt != null)
                {
                    //空白行追加
                    DataRow dr = tmpdt.NewRow();
                    tmpdt.Rows.InsertAt(dr, 0);
                }
                // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoComboChangeFlg = "1";
                item7_MeijishoKirikaeCombo.DataSource = tmpdt;
                item7_MeijishoKirikaeCombo.DisplayMember = "Discript";
                item7_MeijishoKirikaeCombo.ValueMember = "Value";
            }

        }

        private void get_combo_byNendo()
        {
            //データ取得時に年度がいない場合、当年度とする
            int Nendo;
            int ToNendo;
            if (item1_MadoguchiTourokuNendo.Text == "")
            {
                Nendo = DateTime.Today.Year;
                ToNendo = DateTime.Today.AddYears(1).Year;
            }
            else
            {
                int.TryParse(item1_MadoguchiTourokuNendo.SelectedValue.ToString(), out Nendo);
                ToNendo = Nendo + 1;
            }

            //担当部所タブ　協力部所の更新
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

            ////協力依頼書　部所
            //discript = "ShibuMei ";
            //value = "BushoMadoguchiNarabijun, ShibuMei ";
            //table = "Mst_Busho ";
            //where = "BushoDeleteFlag <> 1 AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND BushoShibuCD <> '' " +
            //                    //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/3/31' ) " +
            //                    //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/4/01' ) " +
            //                    " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
            //                    " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
            //                    " ORDER BY BushoMadoguchiNarabijun ";
            //tmpdt = new DataTable();
            //tmpdt = GlobalMethod.getData(discript, value, table, where);
            //if (tmpdt != null)
            //{
            //    DataRow dr = tmpdt.NewRow();
            //    tmpdt.Rows.InsertAt(dr, 0);
            //}
            //DataView dv = new DataView(tmpdt);
            //tmpdt = dv.ToTable(true, new string[] { "Value", "Discript" });

            tmpdt = new DataTable();
            DataTable tmpdt2 = new DataTable();
            //協力依頼書　部所
            discript = "ShibuMei ";
            value = "BushoMadoguchiNarabijun, ShibuMei ";
            table = "Mst_Busho ";
            where = "BushoDeleteFlag <> 1 AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND BushoShibuCD <> '' " +
                                " AND ShibuMei = '" + GlobalMethod.GetCommonValue1("MADOGUCHI_KYOURYOKU_TOUKATSU") + "' " +
                                " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                                " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                                " ORDER BY BushoMadoguchiNarabijun ";
            tmpdt = GlobalMethod.getData(discript, value, table, where);

            if (tmpdt != null)
            {
                DataRow dr = tmpdt.NewRow();
                tmpdt.Rows.InsertAt(dr, 0);
            }

            //協力依頼書　部所
            discript = "ShibuMei ";
            value = "BushoMadoguchiNarabijun, ShibuMei ";
            table = "Mst_Busho ";
            where = "BushoDeleteFlag <> 1 AND BushoNewOld <= 1 AND BushoMadoguchiHyoujiFlg = 1 AND BushoShibuCD <> '' " +
                                //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/3/31' ) " +
                                //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/4/01' ) " +
                                " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                                " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                                " AND ShibuMei <> '" + GlobalMethod.GetCommonValue1("MADOGUCHI_KYOURYOKU_TOUKATSU_CD") + "' " +
                                " ORDER BY BushoMadoguchiNarabijun ";

            tmpdt2 = GlobalMethod.getData(discript, value, table, where);

            if(tmpdt2 != null)
            {
                int cnt = tmpdt.Rows.Count;
                for (int i = 0;i < tmpdt2.Rows.Count;i++)
                {
                    tmpdt.Rows.Add();
                    //tmpdt.Rows[cnt][0] = tmpdt2.Rows[i][0].ToString();
                    tmpdt.Rows[cnt][1] = tmpdt2.Rows[i][1].ToString(); 
                    tmpdt.Rows[cnt][2] = tmpdt2.Rows[i][2].ToString();
                    cnt += 1;
                }
            }

            DataView dv = new DataView(tmpdt);
            tmpdt = dv.ToTable(true, new string[] { "Value", "Discript" });

            item4_KyoRyokuBusho.DataSource = tmpdt;
            item4_KyoRyokuBusho.DisplayMember = "Discript";
            item4_KyoRyokuBusho.ValueMember = "Value";

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
            src_Busho.DataSource = tmpdt;
            src_Busho.DisplayMember = "Discript";
            src_Busho.ValueMember = "Value";
            // Keyで並べたくないので、ListDictionary（詰めた順に表示）を利用する
            ListDictionary ld = new ListDictionary();
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


        public string GetValue(string id)
        {
            string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            SqlConnection sqlconn = new SqlConnection(connStr);
            sqlconn.Open();
            string path = " ";
            try
            {
                string sqlstr = "select * from M_Folder where MENU_ID = 100 and FolderBushoCD = " + id;
                SqlCommand com = new SqlCommand(sqlstr, sqlconn);
                SqlDataReader sdr = com.ExecuteReader();
                while (sdr.Read() == true)
                {
                    path = (string)sdr["FolderPath"];//フォルダパス
                }

                //$$の文字を取り出す
                var Matches = new Regex(@"\$(.+?)\$").Matches(path);

                string changepath = " ";

                //回して一つずつ置き換える
                for (int i = 0; i < Matches.Count; i++)
                {
                    var key = Matches[i].Value;//$がある値
                    var key2 = key.Replace("$", "");//$がない値

                    //ファルダパスの$$を置き換える処理
                    if (key2 == "FOLDER_BASE" || key2 == "FOLDER_KEIYAKUSHO" || key2 == "FOLDER_SHUKEIHYO" || key2 == "FOLDER_HOUKOKUSHO")
                    {
                        string sqlstr2 = "select * from M_CommonMaster where CommonMasterKye = '" + key2 + "'";
                        SqlCommand com2 = new SqlCommand(sqlstr2, sqlconn);
                        SqlDataReader sdr2 = com2.ExecuteReader();

                        try
                        {
                            string value = " ";
                            while (sdr2.Read() == true)
                            {
                                value = (string)sdr2["CommonValue1"]; //CommonValue1

                                //FOLDER_BASEのCommonValue1に$NENDO$が入っていた場合
                                var Matches2 = new Regex(@"\$(.+?)\$").Matches(value);
                                for (int j = 0; j < Matches2.Count; j++)
                                {
                                    var basekey = Matches2[i].Value;//$がある値
                                    //"%年度%"に年度を入れる
                                    string value2 = value.Replace(basekey, "%年度%");
                                    value = value2;
                                }
                            }

                            //代入が一回目か判断→changepathに値を代入していく
                            if (changepath == " ")
                            {
                                string changepath2 = path.Replace(key, value);
                                Console.WriteLine(changepath2);
                                changepath = changepath2;
                            }
                            else
                            {
                                string changepath2 = changepath.Replace(key, value);
                                changepath = changepath2;
                            }
                        }
                        catch (Exception exception)
                        {
                            Console.WriteLine(exception.Message);
                            throw;
                        }
                    }
                    //フォルダパスに$NENDO$が入っていた場合
                    else if (key2 == "NENDO")
                    {
                        string changepath2 = changepath.Replace(key, "%%%");
                        changepath = changepath2;
                    }
                }
                return changepath;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;
            }
            finally
            {
                sqlconn.Close();
            }
        }

        private void pictureBox16_Click(object sender, EventArgs e)
        {
            //窓口担当者プロンプト
            Popup_ChousainList form = new Popup_ChousainList();
            //form.nendo = item1_3.SelectedValue.ToString();
            //窓口部所セット
            if (item1_MadoguchiTantoushaBushoCD.SelectedValue != null)
            {
                form.Busho = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString();
            }
            form.program = "madoguchi";
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item1_MadoguchiTantoushaCD.Text = form.ReturnValue[0];
                item1_MadoguchiTantousha.Text = form.ReturnValue[1];
                item1_MadoguchiTantoushaBushoCD.SelectedValue = form.ReturnValue[2];
                //item1_11_Busho.Text = form.ReturnValue[2];
            }
            item1_MadoguchiTantousha.Focus();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            //窓口担当者をクリア
            if (MessageBox.Show("窓口担当者を削除しますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                item1_MadoguchiTantousha.Text = "";
                item1_MadoguchiTantoushaCD.Text = "";
                item1_MadoguchiTantousha.Focus();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //特調番号履歴プロンプト
            Popup_TokuchoNo form = new Popup_TokuchoNo();
            if (!"00000".Equals(item1_MadoguchiUketsukeBangou.Text))
            {
                form.tokuchouNo = item1_MadoguchiUketsukeBangou.Text;
            }

            form.ShowDialog();
            item1_MadoguchiUketsukeBangouEdaban.Focus();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //管理番号履歴プロンプト
            Popup_KanriNo form = new Popup_KanriNo();
            if (!String.IsNullOrEmpty(item1_MadoguchiKanriBangou.Text))
            {
                form.kanriNo = item1_MadoguchiKanriBangou.Text;
            }

            form.ShowDialog();
            item1_MadoguchiKanriBangou.Focus();
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

            DateTime dt = ((DateTimePicker)sender).Value;
            ((DateTimePicker)sender).Text = dt.ToString("yyyy/MM/dd");
        }

        // 調査概要タブ 報告完了ボタン
        private void button9_Click(object sender, EventArgs e)
        {
            //報告済みのチェックがついていない
            if (!item1_MadoguchiHoukokuzumi.Checked)
            {
                //報告完了　「報告完了にしますがよろしいですか。」不具合No1338対応で、I20104のメッセージを変更してもらう必要あり。
                if (MessageBox.Show(GlobalMethod.GetMessage("I20104", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    // 報告実施日が既に入っている場合は値をセットしない
                    if (item1_MadoguchiHoukokuJisshibi.CustomFormat != "")
                    {
                        //報告実施日に本日日付をセット
                        item1_MadoguchiHoukokuJisshibi.Text = DateTime.Today.ToString("yyyy/MM/dd");
                        item1_MadoguchiHoukokuJisshibi.CustomFormat = "";
                    }

                    //報告済をチェック　1
                    item1_MadoguchiHoukokuzumi.Checked = true;

                    //報告完了ボタン非表示　取消ボタン表示
                    button9.Text = "報告完了取消";

                    //報告実施日に日付が入っていない→再報告

                    MadoguchiHoukokuzumi = "1";

                    //不具合No1338
                    //更新メソッド（今回追加分）を呼ぶ
                    UpdateMadoguchiRealTime();

                }
            }
            //報告済チェックがある
            else
            {
                //報告解除　「報告完了を取り消しますがよろしいですか。」
                if (MessageBox.Show(GlobalMethod.GetMessage("I20108", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    //実施区分が3：中止
                    if (item1_MadoguchiJiishiKubun.SelectedValue != null && "3".Equals(item1_MadoguchiJiishiKubun.SelectedValue.ToString()))
                    {
                        //報告完了ボタンを表示
                        //報告完了取消ボタンを非表示
                        button9.Text = "報告完了";
                    }
                    else
                    {
                        //報告完了ボタンを表示
                        //報告完了取消ボタンを非表示
                        button9.Text = "報告完了";
                    }
                    //報告済を0にする
                    item1_MadoguchiHoukokuzumi.Checked = false;

                    MadoguchiHoukokuzumi = "0";

                    //不具合No1338
                    //更新メソッド（今回追加分）を呼ぶ
                    UpdateMadoguchiRealTime();
                }

            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            // 新規登録時は処理しない
            if(MadoguchiID == "")
            {
                return;
            }

            //品目名取込
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    //調査品目情報テーブルから窓口IDで検索
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT ChousaHinmei " +
                        "FROM ChousaHinmoku " +
                        "WHERE MadoguchiID = '" + MadoguchiID + "' " +
                        "AND ChousaDeleteFlag != 1 ";


                    //Clipboard.SetText(cmd.CommandText);
                    var sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);

                    //データが見つからない場合　
                    if (dt.Rows.Count == 0)
                    {
                        //「調査品目が存在しません。」
                        set_error(GlobalMethod.GetMessage("E20113", ""));
                    }

                    //データがある場合
                    else
                    {
                        //現在の調査品目をクリア
                        item1_MadoguchiChousaHinmoku.Text = "";

                        //取得した品名を「、」で連結する
                        string hinmei = "";
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            //連結
                            if (i == 0)
                            {
                                hinmei += dt.Rows[i][0].ToString();
                            }
                            else
                            {
                                hinmei += "、" + dt.Rows[i][0].ToString();
                            }

                            //連結した文字列が2048文字を超えていた場合 文字数（サロゲートも1カウント）→hinmei.LengthInTextElements
                            if (hinmei.Length > 2048)
                            {
                                //文字切り出し
                                System.Globalization.StringInfo si = new System.Globalization.StringInfo(hinmei);
                                hinmei = si.SubstringByTextElements(0, 2048);

                                break;
                            }
                        }//for終

                        //調査品目に連結したhinmeiをセット
                        item1_MadoguchiChousaHinmoku.Text = hinmei;

                    }//if終
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //受託番号変更機能 「受託番号、特調番号を変更可能にしますがよろしいですか？」
            if (MessageBox.Show(GlobalMethod.GetMessage("I20106", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                //受託番号を検索ボタンを押下可能に
                button1.Enabled = true;

                //受託番号検索ボタンの背景色をRGB(42, 78, 122)
                //受託番号検索ボタンの文字色をRGB(255, 255, 255)
                button1.BackColor = Color.FromArgb(42, 78, 122);
                button1.ForeColor = Color.FromArgb(255, 255, 255);

                //特調番号の枝番を編集可能に
                //item1_MadoguchiUketsukeBangouEdaban.Enabled = true;
                item1_MadoguchiUketsukeBangouEdaban.ReadOnly = false;

                // 白
                item1_MadoguchiUketsukeBangouEdaban.BackColor = Color.FromArgb(255, 255, 255);

                //登録年度を編集可能に
                item1_MadoguchiTourokuNendo.Enabled = true;
            }
        }

        // 調査概要タブの更新ボタン
        private void button6_Click(object sender, EventArgs e)
        {
            //画面モードで処理を変える
            //新規のとき
            if ("insert".Equals(mode))
            {
                // 新規登録を行いますが宜しいですか？
                if (MessageBox.Show(GlobalMethod.GetMessage("I10601", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    set_error("", 0);
                    //登録時必須チェック
                    Boolean requiredFlag = registration_required(1);
                    //登録時データチェック
                    Boolean validateFlag = registration_validate();
                    //MessageBox.Show("必須チェック: " + requiredFlag + " データチェック: " + validateFlag, "確認");

                    if (requiredFlag && validateFlag)
                    {
                        registration();
                        string resultMessage = "";
                        // 報告書フォルダ作成
                        resultMessage = GlobalMethod.CreateTokuchoBangouEdaFolder(MadoguchiID, item1_MadoguchiUketsukeBangouEdaban.Text);

                        Madoguchi_Input form = new Madoguchi_Input();
                        form.UserInfos = UserInfos;
                        // I20101：新規登録に成功しました。
                        form.Message = resultMessage + GlobalMethod.GetMessage("I20101", "");
                        form.mode = "update";
                        form.MadoguchiID = MadoguchiID;
                        form.Show(this.Owner);
                        this.Close();
                    }
                }

            }
            //更新のとき
            else
            {
                if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    set_error("", 0);
                    //登録時必須チェック
                    Boolean requiredFlag = registration_required(1);

                    //登録時データチェック
                    Boolean validateFlag = registration_validate();
                    // MessageBox.Show("必須チェック: " + requiredFlag + " データチェック: " + validateFlag, "確認");
                    if (requiredFlag && validateFlag)
                    {
                        UpdateMadoguchi(1);
                        //registration();

                        // 受託番号を検索ボタンを非活性に
                        button1.Enabled = false;
                    }
                }
            }

        }

        private void registration()
        {
            string methodName = ".registration";
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;
                String process = "";
                //実施区分
                String jisshiKubun = "null";
                if (!String.IsNullOrEmpty(item1_MadoguchiJiishiKubun.Text))
                {
                    jisshiKubun = item1_MadoguchiJiishiKubun.SelectedValue.ToString();
                }

                //遠隔地引渡承認
                int hikiwatashiShounin = 0;
                if (item1_MadoguchiHikiwatsahi.Checked)
                {
                    hikiwatashiShounin = 1;
                }

                // 遠隔地最終検査
                int saishuukensa = 0;
                if (item1_MadoguchiSaishuuKensa.Checked)
                {
                    saishuukensa = 1;
                }

                // 本部単品
                int honbuTanpin = 0;
                if (item1_MadoguchiHonbuTanpinflg.Checked)
                {
                    honbuTanpin = 1;
                }

                //Garoon連携
                int garoon = 0;
                if (item1_GaroonRenkei.Checked)
                {
                    garoon = 1;
                }

                //報告済
                int houkokuzumi = 0;
                if (item1_MadoguchiHoukokuzumi.Checked)
                {
                    houkokuzumi = 1;
                }

                //調査区分 
                int chousaKubunJibusho = 0;
                int chousaKubunShibushibu = 0;
                int chousaKubunHonshibu = 0;
                int chousaKubunShibuhon = 0;
                if (item1_MadoguchiChousaKubunJibusho.Checked)
                {
                    chousaKubunJibusho = 1;
                }
                if (item1_MadoguchiChousaKubunShibuShibu.Checked)
                {
                    chousaKubunShibushibu = 1;
                }
                else if (item1_MadoguchiChousaKubunHonbuShibu.Checked)
                {
                    chousaKubunHonshibu = 1;
                }
                else if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
                {
                    chousaKubunShibuhon = 1;
                }

                //調査種別
                String shubetsu = "null";
                if (!String.IsNullOrEmpty(item1_MadoguchiChousaShubetsu.Text))
                {
                    shubetsu = item1_MadoguchiChousaShubetsu.SelectedValue.ToString();
                }

                //M_COMMON_MASTERからCHOUSAKIJUNBIのデフォルトを取得する
                var dtCommon = new DataTable();
                cmd.CommandText = "SELECT CommonValue1 " +
                    "FROM M_CommonMaster " +
                    "WHERE CommonMasterKye = 'CHOUSAKIJUNBI_DEFAULT' ";
                //データ取得
                var sdaC = new SqlDataAdapter(cmd);
                sdaC.Fill(dtCommon);

                String CommonValue = "  年 月号　";
                if (dtCommon.Rows.Count > 0)
                {
                    CommonValue = dtCommon.Rows[0][0].ToString();
                }

                String ankenJouhouId = item1_MadoguchiAnkenJouhouID.Text;

                //契約担当者CD
                String keiyakutantou = "null";
                if (!String.IsNullOrEmpty(item1_AnkenTantoushaMei_CD.Text))
                {
                    keiyakutantou = item1_AnkenTantoushaMei_CD.Text;
                }

                //業務管理者CD
                String gyoumuKanri = "null";
                if (!String.IsNullOrEmpty(item1_item1_MadoguchiGyoumuKanrishaCD.Text))
                {
                    gyoumuKanri = item1_item1_MadoguchiGyoumuKanrishaCD.Text;
                }

                //窓口担当者CD
                String madoguchiTantouCD = "0";
                if (!String.IsNullOrEmpty(item1_MadoguchiTantoushaCD.Text))
                {
                    madoguchiTantouCD = item1_MadoguchiTantoushaCD.Text;
                }

                //画面のモードで処理を変える
                //登録処理
                if ("insert".Equals(mode))
                {
                    string renban = TokuchoNo_saiban();
                    try
                    {
                        //採番番号取得
                        var comboDt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'MadoguchiId' ";

                        //データ取得
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(comboDt);

                        DataRow dr = comboDt.Rows[0];
                        saibanMadoguchiNo = dr[0].ToString();

                        //採番No（saibanMadoguchiNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanMadoguchiNo + " WHERE SaibanMei = 'MadoguchiId' ";

                        cmd.ExecuteNonQuery();

                        // 進捗状況
                        // 新規登録時→10:依頼
                        // 実施区分が中止→80:中止
                        // 実施区分が中止でなく、報告済み→70:二次検証済
                        string MadoguchiShinchokuJoukyou = "10";

                        // 中止かどうか
                        if (item1_MadoguchiJiishiKubun.Text != "" && item1_MadoguchiJiishiKubun.SelectedValue.ToString() == "3")
                        {
                            MadoguchiShinchokuJoukyou = "80";
                        }
                        else
                        {
                            // 中止でなく、報告済み
                            if(houkokuzumi == 1)
                            {
                                MadoguchiShinchokuJoukyou = "70";
                            }
                        }

                        //窓口情報（MadoguchiJouhou）
                        cmd.CommandText = "INSERT INTO MadoguchiJouhou( " +
                            "MadoguchiID " +
                            ",MadoguchiTourokuNendo " +
                            ",MadoguchiHikiwatsahi " +
                            ",MadoguchiSaishuuKensa " +
                            ",MadoguchiShouninsha " +
                            ",MadoguchiShouninnbi " +
                            ",MadoguchiShimekiribi " +
                            ",MadoguchiTourokubi " +
                            ",MadoguchiHoukokuJisshibi " +
                            ",MadoguchiChousaShubetsu " +
                            ",MadoguchiJiishiKubun " +
                            ",MadoguchiShinchokuJoukyou " +
                            ",MadoguchiJutakuBushoCD " +//受託課所支部
                            ",MadoguchiJutakuTantoushaID " +
                            ",JutakuBushoShozokuCD " +
                            ",MadoguchiTantoushaBushoCD " +
                            ",MadoguchiTantoushaCD " +
                            ",MadoguchiBushoShozokuCD " +
                            ",MadoguchiChousaKubunJibusho " + // 調査区分　自部所
                            ",MadoguchiChousaKubunShibuShibu " + //調査区分　支→支
                            ",MadoguchiChousaKubunHonbuShibu " + // 調査区分　本→支
                            ",MadoguchiChousaKubunShibuHonbu " + //調査区分　支→本
                            ",MadoguchiKanriBangou " +
                            ",MadoguchiJutakuBangou" +
                            ",MadoguchiJutakuBangouEdaban" +
                            ",MadoguchiUketsukeBangou " +
                            ",MadoguchiUketsukeBangouEdaban " +
                            ",MadoguchiHachuuKikanmei " +
                            ",MadoguchiGyoumuMeishou " +
                            ",MadoguchiKoujiKenmei " +
                            ",MadoguchiChousaHinmoku " +
                            ",MadoguchiBikou " +
                            ",MadoguchiTankaTekiyou " +
                            ",MadoguchiNiwatashi " +
                            ",MadoguchiHoukokuzumi " +
                            ",MadoguchiKanriGijutsusha " +
                            ",MadoguchiCreateDate " +
                            ",MadoguchiCreateUser " +
                            ",MadoguchiCreateProgram " +
                            ",MadoguchiUpdateDate " +
                            ",MadoguchiUpdateUser " +
                            ",MadoguchiUpdateProgram " +
                            ",MadoguchiDeleteFlag " +
                            ",MadoguchiOldBushoflg " +
                            ",MadoguchiHonbuTanpinflg " +
                            ",MadoguchiShukeiHyoFolder " +
                            ",MadoguchiHoukokuShoFolder " +
                            ",MadoguchiShiryouHolder " +
                            ",MadoguchiGyoumuKanrishaCD " +
                            ",AnkenJouhouID " +
                            ",MadoguchiHachuukikanCD " +
                            ",MadoguchiGaroonRenkei " +
                            ",MadoguchiKanryou " +
                            ",MadoguchiMitsumoriTeishutu " +
                            ",MadoguchiTeiNyuusatsu " +
                            ",MadoguchiHoukokuMale " +
                            ",MadoguchiIraiMale " +
                            ",MadoguchiIraimotoBusho " +
                            ",MadoguchiAnkenJouhouID " +
                            ",MadoguchiSaishuuKensaCheck " +
                            ",MadoguchiSystemRenban " +
                            ")VALUES(" +
                            saibanMadoguchiNo +
                            ",'" + item1_MadoguchiTourokuNendo.SelectedValue + "' " +//登録年度
                            "," + hikiwatashiShounin + " " +//遠隔地引渡承認
                            "," + saishuukensa + " " +             //遠隔地最終検査
                            ",N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiShouninsha.Text, 0, 0) + "' " +          //遠隔地承認者
                            "," + Get_DateTimePicker("item1_MadoguchiShouninnbi") + " " +   //遠隔地承認日
                            "," + Get_DateTimePicker("item1_MadoguchiShimekiribi") + " " +   //調査担当者への締切日
                            "," + Get_DateTimePicker("item1_MadoguchiTourokubi") + " " +           //登録日
                            "," + Get_DateTimePicker("item1_MadoguchiHoukokuJisshibi") + " " +   //報告実施日
                            "," + shubetsu + " " +   //調査種別　
                            "," + jisshiKubun + " " +                                 //実施区分　
                            //"," + 10 + " " +                         //MadoguchiShinchokuJoukyou
                            "," + MadoguchiShinchokuJoukyou + " " +                         //MadoguchiShinchokuJoukyou
                            ",'" + item1_MadoguchiJutakuBushoCD.SelectedValue + "' " +   //受託課所支部
                            "," + keiyakutantou + " " +            //契約担当者orNULL
                            ",'" + item1_MadoguchiJutakuBushoCD.SelectedValue + "' " +           //受託部所所属長の部所CD
                            ",'" + item1_MadoguchiTantoushaBushoCD.SelectedValue + "' " + //窓口部所
                            "," + madoguchiTantouCD + " " +            //窓口担当者
                            ",'" + item1_MadoguchiTantoushaBushoCD.SelectedValue + "' " + //窓口部所所属長の部所CD
                            "," + chousaKubunJibusho + " " + // 調査区分　自部所
                            "," + chousaKubunShibushibu + " " + //調査区分　支→支
                            "," + chousaKubunHonshibu + " " + // 調査区分　本→支
                            "," + chousaKubunShibuhon + " " + //調査区分　支→本
                            ",N'" + item1_MadoguchiKanriBangou.Text + "' " +          //管理番号
                            ",N'" + item1_MadoguchiJutakuBangou.Text.Replace("-" + item1_MadoguchiJutakuBangouEdaban.Text, "") + "' " +           //受託番号
                            ",N'" + item1_MadoguchiJutakuBangouEdaban.Text + "' " +       //受託番号枝番（？）
                            ",N'" + item1_MadoguchiUketsukeBangou.Text + "' " +            //特調番号
                            ",N'" + item1_MadoguchiUketsukeBangouEdaban.Text + "' " +          //特調番号枝番
                            ",N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiHachuuKikanmei.Text, 0, 0) + "' " +          //発注者名・課名
                            ",N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiGyoumuMeishou.Text, 0, 0) + "' " +          //業務名称
                            ",N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiKoujiKenmei.Text, 0, 0) + "' " +          //工事件名
                            ",N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiChousaHinmoku.Text, 0, 0) + "' " +          //調査品目
                            ",N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiBikou.Text, 0, 0) + "' " +          //備考
                            ",N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiTankaTekiyou.Text, 0, 0) + "' " +          //単価適用地域
                            ",N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiNiwatashi.Text, 0, 0) + "' " +          //荷渡場所
                            "," + 0 + " " +                         //0　報告済
                            ",N'" + item1_KanriGijutsusha_CD.Text + "' " +           //管理技術者
                            ",SYSDATETIME() " +                   // 登録日次
                            ",N'" + UserInfos[0] + "' " +          // 登録ユーザ
                            ",'" + pgmName + methodName + "' " +   // 更新プログラム
                            ",SYSDATETIME() " +                   // 更新日時
                            ",N'" + UserInfos[0] + "' " +          // 更新ユーザ
                            ",'" + pgmName + methodName + "' " +  // 更新プログラム
                            ",0 " +                               // 削除フラグ
                            ",NULL " +                              //null
                            "," + honbuTanpin + " " +               //本部単品 
                            ",N'" + item1_MadoguchiShukeiHyoFolder.Text + "' " +          //集計表フォルダ
                            ",N'" + item1_MadoguchiHoukokuShoFolder.Text + "' " +          //報告書フォルダ
                            ",N'" + item1_MadoguchiShiryouHolder.Text + "' " +          //調査資料フォルダ
                            "," + gyoumuKanri + " " +        //業務管理者の業務管理者CD or Null
                            "," + ankenJouhouId + " " +                 //AnkenJouhou.AnkenJouhouID、未受託の場合はNULL
                            ",null " +
                            "," + garoon + "" +//MadoguchiGaroonRenkei
                            ",0" + //MadoguchiKanryou
                            ",0" + //MadoguchiMitsumoriTeishutu
                            ",0" + //MadoguchiTeiNyuusatsu
                            ",0" + //MadoguchiHoukokuMale
                            ",0" + //MadoguchiIraiMale
                            ",0" + //MadoguchiIraimotoBusho
                            "," + ankenJouhouId + " " +
                            ",0" +
                            "," + renban +
                            ")";
                        process = cmd.CommandText;
                        cmd.ExecuteNonQuery();


                        //採番No（SaibanNo）を取得
                        comboDt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'KyouryokuIraishoID' ";

                        //データ取得
                        sda = new SqlDataAdapter(cmd);
                        sda.Fill(comboDt);
                        dr = comboDt.Rows[0];
                        int saibanKyouryokuNo = int.Parse(dr[0].ToString());

                        //採番No（SaibanNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanKyouryokuNo + " WHERE SaibanMei = 'KyouryokuIraishoID' ";

                        cmd.ExecuteNonQuery();

                        //部所支部名、所属長CDを取得
                        var busho_comboDt = new DataTable();
                        //SQL生成
                        //cmd.CommandText = "SELECT " +
                        //  "ShibuMei " +
                        //  "ShibuMei " +
                        //  "FROM " + "Mst_Busho " +
                        //  "WHERE GyoumuBushoCD = '" + item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString() + "' ";
                        //VIPS 20200506 課題管理表No1314（1038）ADD 窓口ミハル　協力依頼書タブの協力先所属長を「調査統括部」の所属長でデフォルト登録されるようにする
                        cmd.CommandText = "SELECT"
                                        + " Mst_Busho.ShibuMei"
                                        + ", Mst_Chousain.KojinCD"
                                        + " FROM Mst_Busho"
                                        + " LEFT JOIN Mst_Chousain ON Mst_Busho.BushoShozokuChou = Mst_Chousain.ChousainMei"
                                        + " WHERE Mst_Busho.GyoumuBushoCD = '" + CONST_DEFAULT_KYORYOKUSAKI_BUSHO_CD + "'"
                                        ;
                        //cmd.CommandText = "SELECT"
                        //                + " Mst_Busho.ShibuMei"
                        //                + ", Mst_Chousain.KojinCD"
                        //                + " FROM Mst_Busho"
                        //                + " LEFT JOIN Mst_Chousain ON Mst_Busho.BushoShozokuChou = Mst_Chousain.ChousainMei"
                        //                + " WHERE Mst_Busho.GyoumuBushoCD = '" + item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString() + "'"
                        //                ;

                        //データ取得
                        var busho_sda = new SqlDataAdapter(cmd);
                        busho_sda.Fill(busho_comboDt);
                        DataRow busho_dr = busho_comboDt.Rows[0];
                        string bushoShibuMei = "";
                        string BushoShozokuChouCD = "";
                        if (busho_dr != null)
                        {
                            bushoShibuMei = busho_dr[0].ToString();
                            BushoShozokuChouCD = busho_dr[1].ToString();
                        }

                        // 業務区分の設定
                        int w_KyouryokuGyoumuKubun = 0;
                        if (item1_AnkenGyoumuKubun.Text != "")
                        {
                            int num = 0;
                            if (int.TryParse(item1_AnkenGyoumuKubun.SelectedValue.ToString(), out num))
                            {
                                // GyoumuNarabijunCD の値で業務区分の値を切替
                                switch (num)
                                {
                                    case 1: // 1:調査部（一般）
                                    case 5: // 5:事業普及部（一般）
                                    case 6: // 6:事業普及部（物品購入）
                                    case 8: // 8:総合研究所
                                        w_KyouryokuGyoumuKubun = 1;         // 1.一般受託調査
                                        break;
                                    case 3: // 3:調査部（単契）
                                        w_KyouryokuGyoumuKubun = 2;         // 2.単価契約調査
                                        break;
                                    case 4: // 4:調査部（単品)
                                        w_KyouryokuGyoumuKubun = 3;         // 3.単品契約調査
                                        break;
                                    case 2: // 2:調査部（単契含む）
                                        w_KyouryokuGyoumuKubun = 4;         // 4.単価契約を含む一般受託
                                        break;
                                    case 7: // 7:情シス部（一般契約）
                                        w_KyouryokuGyoumuKubun = 5;         // 5.情報開発受託業務
                                        break;
                                    default:
                                        w_KyouryokuGyoumuKubun = 1;         // 1.一般受託調査
                                        break;
                                }
                            }

                        }

                        //協力依頼書情報（KyouryokuIraisho）テーブル登録
                        cmd.CommandText = "INSERT INTO KyouryokuIraisho( " +
                            "KyouryokuIraishoID " +
                            ",MadoguchiID " +
                            ",KyouryokuChousaKijun " +
                            ",KyouryokuChousakijunbi " +
                            ",KyouryokuHoukokuSeigenDate " +
                            ",KyouryokuGyoumuKubun " +
                            ",KyouryokuIraiKubun " +
                            ",KyouryokuUtiawaseyouhi " +
                            ",KyouryokusakiHikiwatashi " +
                            ",KyouryokuJisshikeikakusho " +
                            ",KyouryokuGyoumuNaiyou " +
                            ",KyouryokuCreateDate " +
                            ",KyouryokuCreateUser " +
                            ",KyouryokuCreateProgram " +
                            ",KyouryokuUpdateDate" +
                            ",KyouryokuUpdateUser " +
                            ",KyouryokuUpdateProgram " +
                            ",KyouryokuDeleteFlag";

                        // 支⇒本の場合、協力先部所は窓口部所を初期設定
                        //if (chousaKubunShibuhon == 1)
                        //{
                            cmd.CommandText += ", KyourokuIraisakiBushoOld"
                                            + ", KyourokuIraisakiTantoshaCD"
                                            ;
                        //}
                        cmd.CommandText +=
                           ")VALUES(" +
                           saibanKyouryokuNo +                    //採番
                           ",'" + saibanMadoguchiNo + "' " +               //窓口情報
                           ",'1' " +                              //1
                           ",'" + CommonValue + "' " +            //M_COMMON_MASTER CHOUSAKIJUNBI_DEFAULT
                           "," + Get_DateTimePicker("item1_MadoguchiShimekiribi") + " " +  //調査担当者への締切日
                           //",NULL " +                             //空
                           "," + w_KyouryokuGyoumuKubun +         // 業務区分
                           ",NULL " +                             //空
                           ",'2' " +                              //2
                           ",'2' " +                              //2
                           ",'1' " +                              //1
                           ",'別紙の通り' " +
                           ",SYSDATETIME() " +                    // 登録日時
                           ",N'" + UserInfos[0] + "' " +           // 登録ユーザ
                           ",'" + pgmName + methodName + "' " +   // 登録プログラム
                           ",SYSDATETIME() " +                    // 更新日時
                           ",N'" + UserInfos[0] + "' " +           // 更新ユーザ
                           ",'" + pgmName + methodName + "' " +   // 更新プログラム
                           ",0 ";                                 // 削除フラグ
                        // 支⇒本の場合、協力先部所は窓口部所を初期設定
                        //if (chousaKubunShibuhon == 1)
                        //{
                            //VIPS 20200506 課題管理表No1314（1038）ADD 窓口ミハル　協力依頼書タブの協力先所属長を「調査統括部」の所属長でデフォルト登録されるようにする
                            //1調査統括部のコードで名称と部署所属長コードを取得するよう修正したので元に戻した。
                            // VIPS　20220307　課題管理表No1264(958)　UPDATE　「協力先部所」の初期値を「調査統括部」に変更
                            cmd.CommandText += ", '" + bushoShibuMei + "'"
                            //cmd.CommandText += ", '調査統括部'"
                                            + ", '" + BushoShozokuChouCD + "'"
                                            ;
                        //}
                        cmd.CommandText += ")";

                        cmd.ExecuteNonQuery();


                        //採番No（SaibanNo）を取得
                        var dt2 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'OuenUketsukeID' ";

                        //データ取得
                        var sda2 = new SqlDataAdapter(cmd);
                        sda2.Fill(dt2);

                        DataRow dr2 = dt2.Rows[0];
                        int saibanOuenNo = int.Parse(dr2[0].ToString());

                        //採番No（SaibanNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanOuenNo + " WHERE SaibanMei = 'OuenUketsukeID' ";

                        cmd.ExecuteNonQuery();

                        //応援受付（OuenUketsuke）登録
                        cmd.CommandText = "INSERT INTO OuenUketsuke(" +
                            "OuenUketsukeID " +
                            ",OuenJoukyou " +
                            ",OuenKanryou " +
                            ",MadoguchiID " +
                            ",OuenKanriNo " +
                            ",OuenCreateDate " +
                            ",OuenCreateUser " +
                            ",OuenCreateProgram " +
                            ",OuenUpdateDate " +
                            ",OuenUpdateUser " +
                            ",OuenUpdateProgram " +
                            ",OuenDeleteFlag " +
                            ")VALUES(" +
                            saibanOuenNo;
                        // 支→本にチェックがある場合
                        if(chousaKubunShibuhon == 1)
                        {
                            cmd.CommandText += ", 1";
                        }
                        else
                        {
                            cmd.CommandText += ", 0";
                        }

                        cmd.CommandText += ", 0" +
                            ", '" + saibanMadoguchiNo + "' " +
                            ",'" + item1_MadoguchiKanriBangou.Text + "'" +
                            ",SYSDATETIME() " +                             // 登録日時
                            ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                            ",'" + pgmName + methodName + "' " +            // 登録プログラム
                            ",SYSDATETIME() " +                             // 更新日時
                            ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                            ",'" + pgmName + methodName + "' " +            // 更新プログラム
                            ",0 " +                                         // 削除フラグ
                            ")";
                        cmd.ExecuteNonQuery();

                        //採番No（SaibanNo）を取得
                        var tanpinDt = new DataTable();
                        int TanpinNyuuryokuID = 0;
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'TanpinNyuuryokuID' ";

                        //データ取得
                        var tanpin_sda = new SqlDataAdapter(cmd);
                        tanpin_sda.Fill(tanpinDt);

                        DataRow tanpin_dr = tanpinDt.Rows[0];
                        TanpinNyuuryokuID = int.Parse(tanpin_dr[0].ToString());

                        //採番No（TanpinNyuuryokuID）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            TanpinNyuuryokuID + " WHERE SaibanMei = 'TanpinNyuuryokuID' ";

                        cmd.ExecuteNonQuery();

                        // 単価契約の取得
                        var TankaKeiyakuDt = new DataTable();
                        int TankaKeiyakuID = 0;
                        //SQL生成
                        //cmd.CommandText = "SELECT TankaKeiyakuID FROM TankaKeiyaku WHERE AnkenJouhouID = " + ankenJouhouId;
                        cmd.CommandText = "SELECT TankaKeiyakuID FROM TankaKeiyaku"
                                        + " LEFT JOIN AnkenJouhou ON TankaKeiyaku.TankakeiyakuJutakuBangou = AnkenJouhou.AnkenJutakuBangou"
                                        + " WHERE AnkenJouhou.AnkenJouhouID = " + ankenJouhouId
                                        + " ORDER BY TankaKeiyakuID DESC ";

                        //データ取得
                        var TankaKeiyaku_sda = new SqlDataAdapter(cmd);
                        TankaKeiyaku_sda.Fill(TankaKeiyakuDt);
                        if (TankaKeiyakuDt.Rows.Count > 0)
                        {
                            TankaKeiyakuID = int.Parse(TankaKeiyakuDt.Rows[0][0].ToString());
                        }

                        //単品入力（TanpinNyuuryoku）登録
                        //cmd.CommandText = "INSERT INTO TanpinNyuuryoku(" +
                        //  "TanpinNyuuryokuID " +
                        //  ",MadoguchiID " +
                        //  ",TanpinGyoumuCD " +
                        //  ",TanpinDeleteFlag)VALUES(" +
                        //  TanpinNyuuryokuID +
                        //  ", '" + saibanMadoguchiNo + "' " +
                        //  ", " + "0" + " " + //確定で0
                        //  ",0)";
                        // No352:単品入力画面に、部署、役職、担当者、電話、FAX、メールがコピーされない。対応
                        cmd.CommandText = "INSERT INTO TanpinNyuuryoku(" +
                          "TanpinNyuuryokuID " +
                          ",MadoguchiID " +
                          ",TanpinGyoumuCD " +
                          ",TanpinHachuubusho " +
                          ",TanpinYakushoku " +
                          ",TanpinHachuuTantousha " +
                          ",TanpinTel " +
                          ",TanpinFax " +
                          ",TanpinMail " +
                          ",TanpinCreateDate " +
                          ",TanpinCreateUser " +
                          ",TanpinCreateProgram " +
                          ",TanpinUpdateDate " +
                          ",TanpinUpdateUser " +
                          ",TanpinUpdateProgram " +
                          ",TanpinDeleteFlag " +
                          ")VALUES(" +
                          TanpinNyuuryokuID +
                          ", '" + saibanMadoguchiNo + "' " +
                          //", " + "0" + " " + //確定で0
                          ", " + TankaKeiyakuID + " " +
                          ", N'" + item6_TanpinHachuubusho.Text + "' " +
                          ", N'" + item6_TanpinYakushoku.Text + "' " +
                          ", N'" + item6_TanpinHachuuTantousha.Text + "' " +
                          ", N'" + item6_TanpinTel.Text + "' " +
                          ", N'" + item6_TanpinFax.Text + "' " +
                          ", N'" + item6_TanpinMail.Text + "' " +
                          ",SYSDATETIME() " +                             // 登録日時
                          ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                          ",'" + pgmName + methodName + "' " +            // 登録プログラム
                          ",SYSDATETIME() " +                             // 更新日時
                          ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                          ",'" + pgmName + methodName + "' " +            // 更新プログラム
                          ",0 " +                                         // 削除フラグ
                          ")";
                        cmd.ExecuteNonQuery();

                        //採番No（SaibanNo）を取得
                        var dt3 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'HistoryID' ";

                        //データ取得
                        var sda3 = new SqlDataAdapter(cmd);
                        sda3.Fill(dt3);

                        DataRow dr3 = dt3.Rows[0];
                        int saibanHistoryNo = int.Parse(dr3[0].ToString());

                        //採番No（SaibanNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanHistoryNo + " WHERE SaibanMei = 'HistoryID' ";

                        cmd.ExecuteNonQuery();

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
                            ", " + saibanHistoryNo + " " +
                            ",SYSDATETIME() " +
                            ",'" + UserInfos[0] + "' " +
                            ",N'" + UserInfos[1] + "' " +
                            ",'" + UserInfos[2] + "' " +
                            ",N'" + UserInfos[3] + "' " +
                            ",'調査概要を追加しました ID:" + saibanMadoguchiNo + " Garoon連携区分: " + garoon + "' " +
                            ",'" + pgmName + methodName + "' " +
                            "," + saibanMadoguchiNo + " " +
                            ",NULL " +
                            ",NULL " +
                            ",NULL " +
                            ",NULL " +
                            ",N'" + item1_MadoguchiUketsukeBangou.Text + "-" + item1_MadoguchiUketsukeBangouEdaban.Text + "'" + 
                            ")";

                        cmd.ExecuteNonQuery();


                        //// GaroonTsuikaAtesakiに窓口部所に一致する支部応援のユーザーを追加する
                        //var garoonDt = new DataTable();
                        ////SQL生成
                        //cmd.CommandText = "SELECT " +
                        //  "Mst_Chousain.KojinCD " +
                        //  ",Mst_Chousain.ChousainMei " +
                        //  ",Mst_Chousain.GyoumuBushoCD " +
                        //  ",mb.BushokanriboKamei " +
                        //  "FROM Mst_Shibuouen INNER JOIN Mst_Chousain ON " +
                        //  "Mst_Chousain.KojinCD = Mst_Shibuouen.ShibuouenKojinCD " +
                        //  "AND Mst_Shibuouen.ShibuouenDeleteFlag = 0 " +
                        //  "AND Mst_Chousain.RetireFLG = 0 " +
                        //  "AND Mst_Chousain.GyoumuBushoCD ='" + item1_MadoguchiTantoushaBushoCD.SelectedValue + "' " +
                        //  "INNER JOIN Mst_Busho mb ON mb.GyoumuBushoCD = Mst_Chousain.GyoumuBushoCD ";
                        ////データ取得
                        //var garronSda = new SqlDataAdapter(cmd);
                        //garronSda.Fill(garoonDt);

                        string KojinCD = "";
                        string ChousainMei = "";
                        string GyoumuBushoCD = "";
                        string BushoMei = "";

                        //if (garoonDt != null && garoonDt.Rows.Count > 0)
                        //{
                        //    for (int i = 0; i < garoonDt.Rows.Count; i++) 
                        //    { 
                        //        KojinCD = garoonDt.Rows[i][0].ToString();
                        //        ChousainMei = garoonDt.Rows[i][1].ToString();
                        //        GyoumuBushoCD = garoonDt.Rows[i][2].ToString();
                        //        BushoMei = garoonDt.Rows[i][3].ToString();

                        //        // GaroonTsuikaAtesakiに登録
                        //        cmd.CommandText = "INSERT INTO GaroonTsuikaAtesaki ( " +
                        //        " GaroonTsuikaAtesakiID " +
                        //        ",GaroonTsuikaAtesakiMadoguchiID " +
                        //        ",GaroonTsuikaAtesakiBushoCD " +
                        //        ",GaroonTsuikaAtesakiBusho " +
                        //        ",GaroonTsuikaAtesakiTantoushaCD " +
                        //        ",GaroonTsuikaAtesakiTantousha " +
                        //        ",GaroonTsuikaAtesakiCreateDate " +
                        //        ",GaroonTsuikaAtesakiCreateUser " +
                        //        ",GaroonTsuikaAtesakiCreateProgram " +
                        //        ",GaroonTsuikaAtesakiUpdateDate " +
                        //        ",GaroonTsuikaAtesakiUpdateUser " +
                        //        ",GaroonTsuikaAtesakiUpdateProgram " +
                        //        ",GaroonTsuikaAtesakiDeleteFlag " +
                        //        ") VALUES (" +
                        //        "'" + GlobalMethod.getSaiban("GaroonTsuikaAtesakiID") + "' " + // GaroonTsuikaAtesakiID
                        //        ",'" + saibanMadoguchiNo + "' " +          // GaroonTsuikaAtesakiMadoguchiID
                        //        ",'" + GyoumuBushoCD + "' " +              // GaroonTsuikaAtesakiBushoCD
                        //        ",'" + BushoMei + "' " +                   // GaroonTsuikaAtesakiBusho
                        //        ",'" + KojinCD + "' " +                    // GaroonTsuikaAtesakiTantoushaCD
                        //        ",'" + ChousainMei + "' " +                // GaroonTsuikaAtesakiTantousha
                        //        ",SYSDATETIME() " +                        // GaroonTsuikaAtesakiCreateDate
                        //        ",'" + UserInfos[0] + "'" +                // GaroonTsuikaAtesakiCreateUser
                        //        ",'窓口ミハル'" +                          // GaroonTsuikaAtesakiCreateProgram
                        //        ",SYSDATETIME() " +                        // GaroonTsuikaAtesakiUpdateDate
                        //        ",'" + UserInfos[0] + "'" +                // GaroonTsuikaAtesakiUpdateUser
                        //        ",'窓口ミハル'" +                          // GaroonTsuikaAtesakiUpdateProgram
                        //        ",0 " +                                    // GaroonTsuikaAtesakiDeleteFlag
                        //        ") ";

                        //        cmd.ExecuteNonQuery();
                        //    }
                        //}

                        //窓口情報（MadoguchiJouhou）テーブルからGaroon連携対象（GaroonRenkeiKubn）を取得
                        var dt4 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "MadoguchiTantoushaCD,MadoguchiKanriGijutsusha " +
                          ",MadoguchiUketsukeBangou,MadoguchiUketsukeBangouEdaban,MadoguchiGaroonRenkei " +
                          "FROM MadoguchiJouhou " +
                          "WHERE MadoguchiID = " + saibanMadoguchiNo + "";

                        //データ取得
                        var sda4 = new SqlDataAdapter(cmd);
                        sda4.Fill(dt4);

                        String atesaki = dt4.Rows[0][0].ToString();
                        String kanriGijutusha = dt4.Rows[0][1].ToString();
                        String tokuchouNo = dt4.Rows[0][2].ToString();
                        String tokuchouNoEda = dt4.Rows[0][3].ToString();
                        String garoonOn = dt4.Rows[0][4].ToString();

                        //窓口メール送信（MadoguchiMail）テーブルからメッセージID（MadoguchiMailMessageID）を取得
                        var dt5 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "MadoguchiMailMessageID " +
                          "FROM MadoguchiMail " +
                          //"WHERE MadoguchiMailTokuchoBangou = '" + tokuchouNo + "-" + tokuchouNoEda + "'" +
                          "WHERE MadoguchiMailTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo +"'" +
                          "AND MadoguchiMailTokuchoBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_MadoguchiUketsukeBangouEdaban.Text + "' ";

                        //データ取得
                        var sda5 = new SqlDataAdapter(cmd);
                        sda5.Fill(dt5);

                        String mailMessageID = "0";

                        if (dt5.Rows.Count > 0)
                        {
                            mailMessageID = dt5.Rows[0][0].ToString();
                        }

                        //管理技術者（MadoguchiJouhou.MadoguchiKanriGijutsusha）が空でない場合
                        if (!String.IsNullOrEmpty(kanriGijutusha) && kanriGijutusha != "0")
                        {
                            //宛先がnullじゃない
                            if (!String.IsNullOrEmpty(atesaki))
                            {
                                atesaki = atesaki + ";" + kanriGijutusha;
                            }
                            //宛先がnull
                            else
                            {
                                atesaki += kanriGijutusha;
                            }
                        }

                        //MadoguchiJouhouMadoguchiL1Chouを取得
                        var dt6 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT DISTINCT " +
                          "MadoguchiL1ChousaTantoushaCD,MadoguchiL1ChousaBushoCD " +
                          "FROM MadoguchiJouhouMadoguchiL1Chou " +
                          "WHERE MadoguchiID=" + saibanMadoguchiNo + "";

                        //データ取得
                        var sda6 = new SqlDataAdapter(cmd);
                        sda6.Fill(dt6);

                        for (int i = 0; i < dt6.Rows.Count; i++)
                        {
                            //調査員担当者（MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaTantoushaCD）が空でない場合
                            String chousaTantousha = dt6.Rows[i][0].ToString();
                            if (!String.IsNullOrEmpty(chousaTantousha) && chousaTantousha != "0")
                            {
                                //宛先が空でない場合
                                if (!String.IsNullOrEmpty(atesaki))
                                {
                                    atesaki = atesaki + ";" + chousaTantousha;
                                }
                                //宛先がnull
                                else
                                {
                                    atesaki = chousaTantousha;
                                }
                            }

                            //調査担当部所コード（MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaBushoCD）が空でない場合
                            String chousaTantoubusho = dt6.Rows[i][1].ToString();
                            if (!String.IsNullOrEmpty(chousaTantoubusho))
                            {
                                //支部応援（Mst_Shibuouen）と、調査員マスタ（Mst_Chousain）を結合し担当者を取得
                                var dt7 = new DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "Mst_Chousain.KojinCD " +
                                  "FROM Mst_Shibuouen INNER JOIN Mst_Chousain ON " +
                                  "Mst_Chousain.KojinCD = Mst_Shibuouen.ShibuouenKojinCD " +
                                  //"AND Mst_Shibuouen.ShibuouenDeleteFlag = 0 " +
                                  //"AND Mst_Chousain.RetireFLG = 0 " +
                                  "AND Mst_Chousain.GyoumuBushoCD ='" + chousaTantoubusho + "' ";

                                //データ取得
                                var sda7 = new SqlDataAdapter(cmd);
                                sda7.Fill(dt7);

                                for (int j = 0; j < dt7.Rows.Count; j++)
                                {
                                    if(dt7.Rows[j][0] != null && dt7.Rows[j][0].ToString() != "0")
                                    {
                                        //宛先が空でない場合
                                        if (!String.IsNullOrEmpty(atesaki))
                                        {
                                            atesaki = atesaki + ";" + dt7.Rows[j][0].ToString();
                                        }
                                        //宛先がnull
                                        else
                                        {
                                            atesaki = dt7.Rows[j][0].ToString();
                                        }
                                    }
                                }//for end
                            }//if end
                        }//for end

                        //宛先が空でない場合
                        if (!String.IsNullOrEmpty(atesaki))
                        {
                            //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブルから
                            var dt8 = new DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "MailInfoCSVWorkID " +
                              "FROM MailInfoCSVWork " +
                              "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo + "-" + tokuchouNoEda + "' " +
                              "AND MailInfoCSVWorkCSVOutFlg = 0 " +
                              "AND MailInfoCSVWorkGaRenkeiFlg = 0 " +
                              "AND MailInfoCSVWorkDeleteFlag = 0";

                            //データ取得
                            var sda8 = new SqlDataAdapter(cmd);
                            sda8.Fill(dt8);

                            //メール情報CSV抽出用ワークのデータがある
                            if (dt8.Rows.Count > 0)
                            {
                                String workId = dt8.Rows[0][0].ToString();
                                //ガルーン連携がチェックの場合
                                if ("1".Equals(garoonOn))
                                {

                                    //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブル更新
                                    cmd.CommandText = "UPDATE MailInfoCSVWork SET " +
                                        "MailInfoCSVWorkAtesakiUser = '" + atesaki + "' " +
                                        ",MailInfoCSVWorkUpdateDate = SYSDATETIME() " +
                                        ",MailInfoCSVWorkUpdateUser = '" + UserInfos[0] + "' " +
                                        ",MailInfoCSVWorkUpdateProgram = '" + pgmName + methodName + "' " +
                                        "WHERE MailInfoCSVWorkTokuchoBangou = N'" + tokuchouNo + "-" + tokuchouNoEda + "' AND MailInfoCSVWorkDeleteFlag = 0";

                                    cmd.ExecuteNonQuery();



                                }
                                //連携が未チェックの場合
                                else
                                {

                                    //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブル削除
                                    cmd.CommandText = "DELETE FROM MailInfoCSVWork " +
                                        "WHERE MailInfoCSVWorkTokuchoBangou = N'" + tokuchouNo + "-" + tokuchouNoEda + "' ";

                                    cmd.ExecuteNonQuery();
                                }
                            }
                            //メール情報CSV抽出用ワークのデータがない
                            else
                            {
                                //ガルーン連携がチェックの場合
                                if ("1".Equals(garoonOn))
                                {
                                    //採番No（SaibanNo）を取得
                                    var dt9 = new DataTable();
                                    //SQL生成
                                    cmd.CommandText = "SELECT " +
                                      "SaibanNo+SaibanCountupNo AS SaibanNo " +
                                      "FROM " + "M_Saiban " +
                                      "WHERE SaibanMei = 'MailInfoCSVWorkID' ";

                                    //データ取得
                                    var sda9 = new SqlDataAdapter(cmd);
                                    sda9.Fill(dt9);

                                    int saibanMailInfoNo = int.Parse(dt9.Rows[0][0].ToString());

                                    //採番No（SaibanNo）を更新
                                    cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                                        saibanMailInfoNo + " WHERE SaibanMei = 'MailInfoCSVWorkID' ";

                                    cmd.ExecuteNonQuery();

                                    //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブル登録
                                    cmd.CommandText = "INSERT INTO MailInfoCSVWork(" +
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
                                    ",MailInfoCSVWorkDeleteFlag" +
                                    ")VALUES(" +
                                    saibanMailInfoNo +
                                    ", '" + saibanMadoguchiNo + "' " +
                                    ",N'" + tokuchouNo + "-" + tokuchouNoEda + "' " +
                                    "," + mailMessageID + " " +
                                    ",'" + atesaki + "' " +
                                    ",0" +
                                    ",0" +
                                    ",SYSDATETIME() " +                             // 登録日時
                                    ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                    ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                    ",SYSDATETIME() " +                             // 更新日時
                                    ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                    ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                    ",0 " +                                         // 削除フラグ
                                    ")";
                                    cmd.ExecuteNonQuery();

                                    // I30203:Garoon追加宛先を追加しました。
                                    GlobalMethod.outputLogger("Madoguhi_Input_registration", GlobalMethod.GetMessage("I30203", "") + " ID:" + saibanMadoguchiNo + " Garoon連携区分:" + item1_GaroonRenkei.ToString(), "insert", "DEBUG");

                                }
                            }

                            if ("1".Equals(garoonOn))
                            {
                                // 窓口情報の連携実行日時を更新
                                string datetTime = DateTime.Now.ToString();

                                cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                                "MadoguchiGaroonRenkeiJikouDate = '" + datetTime + "' " +
                                "Where MadoguchiID = '" + saibanMadoguchiNo + "' ";
                                cmd.ExecuteNonQuery();

                                // 更新日時の表記を更新
                                item1_GaroonUpdateDisp.Text = datetTime;
                            }



                        }
                        //宛先がnull
                        else
                        {

                            //メール情報CSV抽出用ワーク（MailInfoCSVWork）テーブル削除
                            cmd.CommandText = "DELETE FROM MailInfoCSVWork " +
                                "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo + "-" + tokuchouNoEda + "' " +
                                "AND MailInfoCSVWorkCSVOutFlg = 0 " +
                                "AND MailInfoCSVWorkGaRenkeiFlg = 0 ";

                            cmd.ExecuteNonQuery();

                        }

                        // 469 Garoon連携担当者の自動設定
                        // 担当者宛先の自動設定（窓口担当者、調査担当者除く）

                        // 管理技術者が空でない場合、GaroonTsuikaAtesakiテーブルにデータを追加する
                        if (item1_KanriGijutsushaNM.Text != null && item1_KanriGijutsushaNM.Text != "")
                        {
                            var dt9 = new DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "KojinCD " +
                              ",ChousainMei " +
                              ",mc.GyoumuBushoCD " +
                              ",mb.BushokanriboKamei " +
                              "FROM Mst_Chousain mc " +
                              "LEFT JOIN Mst_Busho mb ON mb.GyoumuBushoCD = mc.GyoumuBushoCD " +
                              "WHERE mc.KojinCD = '" + item1_KanriGijutsusha_CD.Text + "' ";

                            //データ取得
                            var sda9 = new SqlDataAdapter(cmd);
                            sda9.Fill(dt9);

                            KojinCD = "";
                            ChousainMei = "";
                            GyoumuBushoCD = "";
                            BushoMei = "";

                            if (dt9 != null && dt9.Rows.Count > 0)
                            {
                                KojinCD = dt9.Rows[0][0].ToString();
                                ChousainMei = dt9.Rows[0][1].ToString();
                                GyoumuBushoCD = dt9.Rows[0][2].ToString();
                                BushoMei = dt9.Rows[0][3].ToString();

                                // 窓口部所に一致する支部応援で既に登録済みかもしれないので、存在チェック
                                string where = "GaroonTsuikaAtesakiMadoguchiID = '" + saibanMadoguchiNo + "' " +
                                               "AND GaroonTsuikaAtesakiBushoCD = '" + GyoumuBushoCD + "' " +
                                               "AND GaroonTsuikaAtesakiTantoushaCD = '" + KojinCD + "' " +
                                               "AND GaroonTsuikaAtesakiDeleteFlag <> 1";

                                var tmpdt = GlobalMethod.getData("GaroonTsuikaAtesakiTantoushaCD", "GaroonTsuikaAtesakiTantoushaCD", "GaroonTsuikaAtesaki", where);
                                // データ件数が0件なら登録
                                if (tmpdt != null && tmpdt.Rows.Count == 0)
                                {
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
                                    //不具合No1332(1084) 
                                    ",GaroonTsuikaAtesakiGamenFlag " +
                                    ") VALUES (" +
                                    "'" + GlobalMethod.getSaiban("GaroonTsuikaAtesakiID") + "' " + // GaroonTsuikaAtesakiID
                                    ",'" + saibanMadoguchiNo + "' " +          // GaroonTsuikaAtesakiMadoguchiID
                                    ",'" + GyoumuBushoCD + "' " +              // GaroonTsuikaAtesakiBushoCD
                                    ",N'" + BushoMei + "' " +                   // GaroonTsuikaAtesakiBusho
                                    ",'" + KojinCD + "' " +                    // GaroonTsuikaAtesakiTantoushaCD
                                    ",N'" + ChousainMei + "' " +                // GaroonTsuikaAtesakiTantousha
                                    ",SYSDATETIME() " +                             // 登録日時
                                    ",N'" + UserInfos[0] + "' " +                    // 登録ユーザ
                                    ",'" + pgmName + methodName + "' " +            // 登録プログラム
                                    ",SYSDATETIME() " +                             // 更新日時
                                    ",N'" + UserInfos[0] + "' " +                    // 更新ユーザ
                                    ",'" + pgmName + methodName + "' " +            // 更新プログラム
                                    ",0 " +                                         // 削除フラグ
                                    //不具合No1332(1084) ここは0 なのか1なのかわからん。
                                    ",0 " +
                                    ") ";

                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        // 更新モードに遷移する為にMadoguchiIDをセット
                        MadoguchiID = saibanMadoguchiNo.ToString();
                    }
                    catch (Exception)
                    {
                        //エラーE20109「新規登録に失敗しました。」
                        //set_error(GlobalMethod.GetMessage("E20901", ""));
                        //MessageBox.Show("era-:" + process, "test", MessageBoxButtons.OK);
                        //Clipboard.SetText(process);
                        transaction.Rollback();
                        throw;
                    }

                    transaction.Commit();
                    //登録処理終了後
                    //MessageBox.Show("登録完了", "確認");
                    //窓口ミハル詳細画面に遷移する（編集モード）
                    //パラメータ  ログインユーザーID  UserInfos[0]
                    //モード:update
                    mode = "update";
                    //窓口情報ID saibanMadoguchiNo 

                }
                //更新のとき
                else
                {
                    //進捗状況 10：依頼　調査中→40：集計中　70：二次検済
                    int shinchoku = 10;
                    //実施区分が1
                    if ("1".Equals(jisshiKubun))
                    {
                        //報告済み
                        if (houkokuzumi == 1)
                        {
                            shinchoku = 70;
                        }
                        //報告済みじゃない
                        else
                        {
                            //最小値を取る
                            var dtShinchoku = new DataTable();
                            cmd.CommandText = "SELECT min(MadoguchiL1ChousaShinchoku)as minShonchoku " +
                                "FROM MadoguchiJouhouMadoguchiL1Chou " +
                                "WHERE MadoguchiID = " + MadoguchiID + "";
                            //データ取得
                            var sda = new SqlDataAdapter(cmd);
                            sda.Fill(dtShinchoku);



                            String minStr = dtShinchoku.Rows[0][0].ToString();
                            if (!String.IsNullOrEmpty(minStr))
                            {

                                int minShinchoku = int.Parse(minStr);

                                if (minShinchoku == 10)
                                {
                                    shinchoku = 40;
                                }
                                else
                                {
                                    shinchoku = 10;
                                }
                            }
                        }
                    }
                    //実施区分が2打診中
                    else if ("2".Equals(jisshiKubun))
                    {
                        //報告済み
                        if (houkokuzumi == 1)
                        {
                            shinchoku = 70;
                        }
                        //報告済みじゃない
                        else
                        {
                            //最小値を取る
                            var dtShinchoku = new DataTable();
                            cmd.CommandText = "SELECT min(MadoguchiL1ChousaShinchoku) " +
                                "FROM MadoguchiJouhouMadoguchiL1Chou " +
                                "WHERE MadoguchiID = " + MadoguchiID + "";
                            //データ取得
                            var sda = new SqlDataAdapter(cmd);
                            sda.Fill(dtShinchoku);

                            string minStr = dtShinchoku.Rows[0][0].ToString();
                            if (!String.IsNullOrEmpty(minStr))
                            {
                                shinchoku = int.Parse(minStr);
                            }
                            else
                            {
                                shinchoku = 10;
                            }
                        }
                    }
                    try
                    {
                        //窓口情報更新
                        cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                            "MadoguchiTourokuNendo = " + item1_MadoguchiTourokuNendo.SelectedValue + " " +
                            ",MadoguchiHikiwatsahi = " + hikiwatashiShounin + " " +
                            ",MadoguchiSaishuuKensa = " + saishuukensa + " " +
                            ",MadoguchiShouninsha = N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiShouninsha.Text, 0, 0) + "' " +
                            ",MadoguchiShouninnbi = " + Get_DateTimePicker("item1_MadoguchiShouninnbi") + " " +
                            ",MadoguchiShimekiribi = " + Get_DateTimePicker("item1_MadoguchiShimekiribi") + " " +
                            ",MadoguchiTourokubi = " + Get_DateTimePicker("item1_MadoguchiTourokubi") + " " +
                            ",MadoguchiHoukokuJisshibi = " + Get_DateTimePicker("item1_MadoguchiHoukokuJisshibi") + " " +
                            ",MadoguchiChousaShubetsu = " + shubetsu + " " +
                            ",MadoguchiJiishiKubun = " + jisshiKubun + " " +
                            ",MadoguchiShinchokuJoukyou = " + shinchoku + " " +
                            ",MadoguchiJutakuBushoCD = '" + item1_MadoguchiJutakuBushoCD.SelectedValue + "' " +
                            ",MadoguchiJutakuTantoushaID = " + keiyakutantou + " " +            //契約担当者orNULL　
                            ",JutakuBushoShozokuCD = '" + item1_MadoguchiJutakuBushoCD.SelectedValue + "' " +
                            ",MadoguchiTantoushaBushoCD = '" + item1_MadoguchiTantoushaBushoCD.SelectedValue + "' " +
                            ",MadoguchiTantoushaCD = " + madoguchiTantouCD + " " +
                            ",MadoguchiBushoShozokuCD = '" + item1_MadoguchiTantoushaBushoCD.SelectedValue + "' " +
                            ",MadoguchiChousaKubunJibusho = " + chousaKubunJibusho + " " +
                            ",MadoguchiChousaKubunShibuShibu = " + chousaKubunShibushibu + " " +
                            ",MadoguchiChousaKubunHonbuShibu = " + chousaKubunHonshibu + " " +
                            ",MadoguchiChousaKubunShibuHonbu = " + chousaKubunShibuhon + " " +
                            ",MadoguchiKanriBangou = N'" + item1_MadoguchiKanriBangou.Text + "' " +
                            ",MadoguchiJutakuBangou = N'" + item1_MadoguchiJutakuBangou.Text.Replace("-" + item1_MadoguchiJutakuBangouEdaban.Text, "") + "' " +
                            ",MadoguchiJutakuBangouEdaban = N'" + item1_MadoguchiJutakuBangouEdaban.Text + "' " +       //受託番号枝番
                            ",MadoguchiUketsukeBangou = N'" + item1_MadoguchiUketsukeBangou.Text + "' " + //特調番号
                            ",MadoguchiUketsukeBangouEdaban = N'" + item1_MadoguchiUketsukeBangouEdaban.Text + "' " +
                            ",MadoguchiHachuuKikanmei = N'" + item1_MadoguchiHachuuKikanmei.Text + "' " +
                            ",MadoguchiGyoumuMeishou = N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiGyoumuMeishou.Text, 0, 0) + "' " +
                            ",MadoguchiKoujiKenmei = N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiKoujiKenmei.Text, 0, 0) + "' " +
                            ",MadoguchiChousaHinmoku = N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiChousaHinmoku.Text, 0, 0) + "' " +
                            ",MadoguchiBikou = N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiBikou.Text, 0, 0) + "' " +
                            ",MadoguchiTankaTekiyou = N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiTankaTekiyou.Text, 0, 0) + "' " +
                            ",MadoguchiNiwatashi = N'" + GlobalMethod.ChangeSqlText(item1_MadoguchiNiwatashi.Text, 0, 0) + "' " +
                            ",MadoguchiHoukokuzumi = '" + houkokuzumi + "' " +
                            ",MadoguchiKanriGijutsusha = N'" + item1_KanriGijutsusha_CD.Text + "' " +
                            ",MadoguchiUpdateDate = SYSDATETIME()" +
                            ",MadoguchiUpdateUser = N'" + UserInfos[0] + "' " +
                            ",MadoguchiUpdateProgram = '" + pgmName + methodName + "' " +
                            ",MadoguchiDeleteFlag = 0 " +
                            ",MadoguchiOldBushoflg = 0 " +
                            ",MadoguchiHonbuTanpinflg = " + honbuTanpin + " " +
                            ",MadoguchiShukeiHyoFolder  = N'" + item1_MadoguchiShukeiHyoFolder.Text + "' " +
                            ",MadoguchiHoukokuShoFolder = N'" + item1_MadoguchiHoukokuShoFolder.Text + "' " +
                            ",MadoguchiShiryouHolder = N'" + item1_MadoguchiShiryouHolder.Text + "' " +
                            ",MadoguchiGyoumuKanrishaCD = " + gyoumuKanri + " " + //業務管理者の業務管理者CD or Null
                            ",AnkenJouhouID = " + ankenJouhouId + " " +
                            ",MadoguchiHachuukikanCD = null ";

                        //受託番号(＝案件番号,特調番号)が変わった場合
                        if (!beforeJutaku.Equals(item1_MadoguchiUketsukeBangou.Text) || !befireTokuchoEda.Equals(item1_MadoguchiUketsukeBangouEdaban.Text))
                        {
                            string renban = TokuchoNo_saiban();
                            cmd.CommandText += ",MadoguchiSystemRenban = " + renban + " ";
                        }

                        cmd.CommandText += " WHERE MadoguchiID =  " + MadoguchiID;
                        process = cmd.CommandText;
                        cmd.ExecuteNonQuery();


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
                            ",'調査概要を更新しました ID:" + MadoguchiID + " Garoon連携区分: " + garoon + "' " +
                            ",'" + pgmName + methodName + "' " +
                            "," + MadoguchiID + " " +
                            ",NULL " +
                            ",NULL " +
                            ",NULL " +
                            ",NULL " +
                            ",'" + Header1.Text + "' " +
                            ")";

                        cmd.ExecuteNonQuery();

                        //部所ＣＤから部所支部と部所課名を取得
                        var bushoDt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "ShibuMei,KaMei " +
                          "FROM Mst_Busho " +
                          "WHERE GyoumuBushoCD = '" + item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString() + "' ";

                        //データ取得
                        var sdaBusho = new SqlDataAdapter(cmd);
                        sdaBusho.Fill(bushoDt);


                        //GyoumuJouhouMadoguchiの窓口情報を更新
                        cmd.CommandText = "UPDATE GyoumuJouhouMadoguchi  SET " +
                            " GyoumuJouhouMadoGyoumuBushoCD = " + "'" + item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString() + "' " +
                            ", GyoumuJouhouMadoShibuMei = " + "N'" + bushoDt.Rows[0][0] + "' " +
                            ", GyoumuJouhouMadoKamei = " + "N'" + bushoDt.Rows[0][1] + "' " +
                            ", GyoumuJouhouMadoKojinCD = " + "'" + item1_MadoguchiTantoushaCD.Text + "' " +
                            ", GyoumuJouhouMadoChousainMei = " + "N'" + item1_MadoguchiTantousha.Text + "' " +
                            " WHERE GyoumuJouhouID =  " + ankenJouhouId;
                        cmd.ExecuteNonQuery();


                        //協力依頼書情報（KyouryokuIraisho）テーブル更新
                        cmd.CommandText = "UPDATE KyouryokuIraisho SET " +
                            "KyouryokuChousaKijun = 1 " +
                            ",KyouryokuChousakijunbi = '" + CommonValue + "'" +
                            ",KyouryokuHoukokuSeigenDate = " + Get_DateTimePicker("item1_MadoguchiShimekiribi") + " " +
                            ",KyouryokuUpdateDate = SYSDATETIME()" +
                            ",KyouryokuUpdateUser = N'" + UserInfos[0] + "' " +
                            ",KyouryokuUpdateProgram = '" + pgmName + methodName + "' " +
                            "WHERE MadoguchiID =" + MadoguchiID + " ";

                        int updateCount = cmd.ExecuteNonQuery();

                        //update件数がなければ新規登録　
                        if (updateCount == 0)
                        {

                            //採番No（SaibanNo）を取得
                            var comboDt = new DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "SaibanNo+SaibanCountupNo AS SaibanNo " +
                              "FROM " + "M_Saiban " +
                              "WHERE SaibanMei = 'KyouryokuIraishoID' ";

                            //データ取得
                            var sda = new SqlDataAdapter(cmd);
                            sda.Fill(comboDt);
                            DataRow dr = comboDt.Rows[0];
                            int saibanKyouryokuNo = int.Parse(dr[0].ToString());

                            //採番No（SaibanNo）を更新
                            cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                                saibanKyouryokuNo + " WHERE SaibanMei = 'KyouryokuIraishoID' ";

                            cmd.ExecuteNonQuery();

                            //協力依頼書情報（KyouryokuIraisho）テーブル登録
                            cmd.CommandText = "INSERT INTO KyouryokuIraisho( " +
                                "KyouryokuIraishoID " +
                                ",MadoguchiID " +
                                ",KyouryokuChousaKijun " +
                                ",KyouryokuChousakijunbi " +
                                ",KyouryokuHoukokuSeigenDate " +
                                ",KyouryokuGyoumuKubun " +
                                ",KyouryokuIraiKubun " +
                                ",KyouryokuUtiawaseyouhi " +
                                ",KyouryokusakiHikiwatashi " +
                                ",KyouryokuJisshikeikakusho " +
                                ",KyourokuIraisakiTantoshaCD " +
                                ",KyouryokuCreateDate " +
                                ",KyouryokuCreateUser " +
                                ",KyouryokuCreateProgram " +
                                ",KyouryokuUpdateDate" +
                                ",KyouryokuUpdateUser " +
                                ",KyouryokuUpdateProgram " +
                                ",KyouryokuDeleteFlag)VALUES(" +
                                saibanKyouryokuNo +                    //採番
                                ",'" + MadoguchiID + "' " +               //窓口情報
                                ",'1' " +                              //1
                                ",'" + CommonValue + "' " +            //M_COMMON_MASTER CHOUSAKIJUNBI_DEFAULT
                                "," + Get_DateTimePicker("item1_MadoguchiShimekiribi") + " " +  //調査担当者への締切日
                                ",NULL " +                             //空
                                ",NULL " +                             //空
                                ",'2' " +                              //2
                                ",'2' " +                              //2
                                ",'1' " +                              //1
                                ",NULL " +                             //NULL
                                ",SYSDATETIME() " +                    // 登録日時
                                ",N'" + UserInfos[0] + "' " +           // 登録ユーザ
                                ",'" + pgmName + methodName + "' " +   // 登録プログラム
                                ",SYSDATETIME() " +                    // 更新日時
                                ",N'" + UserInfos[0] + "' " +           // 更新ユーザ
                                ",'" + pgmName + methodName + "' " +   // 更新プログラム
                                ",0 " +                                // 削除フラグ
                                ")";
                            cmd.ExecuteNonQuery();

                        }


                        // 意図がわからなかったのでコメントアウト　2021/06/24
                        ////単品入力（TanpinNyuuryoku）テーブル更新
                        //cmd.CommandText = "UPDATE TanpinNyuuryoku SET " +
                        //    "TanpinGyoumuCD = 0 " +　//確定で0
                        //    "WHERE MadoguchiID =" + MadoguchiID + " ";

                        //updateCount = cmd.ExecuteNonQuery();
                        ////update件数がなければ新規登録
                        //if (updateCount == 0)
                        //{
                        //    //採番No（SaibanNo）を取得
                        //    var tanpinDt = new DataTable();
                        //    int TanpinNyuuryokuID = 0;
                        //    //SQL生成
                        //    cmd.CommandText = "SELECT " +
                        //      "SaibanNo+SaibanCountupNo AS SaibanNo " +
                        //      "FROM " + "M_Saiban " +
                        //      "WHERE SaibanMei = 'TanpinNyuuryokuID' ";

                        //    //データ取得
                        //    var tanpin_sda = new SqlDataAdapter(cmd);
                        //    tanpin_sda.Fill(tanpinDt);

                        //    DataRow tanpin_dr = tanpinDt.Rows[0];
                        //    TanpinNyuuryokuID = int.Parse(tanpin_dr[0].ToString());

                        //    //採番No（TanpinNyuuryokuID）を更新
                        //    cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                        //        TanpinNyuuryokuID + " WHERE SaibanMei = 'TanpinNyuuryokuID' ";

                        //    cmd.ExecuteNonQuery();
                        //    //単品入力（TanpinNyuuryoku）登録
                        //    cmd.CommandText = "INSERT INTO TanpinNyuuryoku(" +
                        //      "TanpinNyuuryokuID " +
                        //      ",MadoguchiID " +
                        //      ",TanpinGyoumuCD " +
                        //      ",TanpinDeleteFlag)VALUES(" +
                        //      TanpinNyuuryokuID +
                        //      ", '" + MadoguchiID + "' " +
                        //      ", " + "0" + " " + //確定で0
                        //      ",0)";
                        //    cmd.ExecuteNonQuery();
                        //}


                        //採番No（SaibanNo）を取得
                        var dt3 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SaibanNo+SaibanCountupNo AS SaibanNo " +
                          "FROM " + "M_Saiban " +
                          "WHERE SaibanMei = 'HistoryID' ";

                        //データ取得
                        var sda3 = new SqlDataAdapter(cmd);
                        sda3.Fill(dt3);

                        DataRow dr3 = dt3.Rows[0];
                        int saibanHistoryNo = int.Parse(dr3[0].ToString());

                        //採番No（SaibanNo）を更新
                        cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                            saibanHistoryNo + " WHERE SaibanMei = 'HistoryID' ";

                        cmd.ExecuteNonQuery();

                        ////履歴登録
                        //cmd.CommandText = "INSERT INTO T_HISTORY(" +
                        //    "H_DATE_KEY " +
                        //    ",H_NO_KEY " +
                        //    ",H_OPERATE_DT " +
                        //    ",H_OPERATE_USER_ID " +
                        //    ",H_OPERATE_USER_MEI " +
                        //    ",H_OPERATE_USER_BUSHO_CD " +
                        //    ",H_OPERATE_USER_BUSHO_MEI " +
                        //    ",H_OPERATE_NAIYO " +
                        //    ",H_ProgramName " +
                        //    ",MadoguchiID " +
                        //    ",HistoryBeforeTantoubushoCD " +
                        //    ",HistoryBeforeTantoushaCD " +
                        //    ",HistoryAfterTantoubushoCD " +
                        //    ",HistoryAfterTantoushaCD " +
                        //    ")VALUES(" +
                        //    "SYSDATETIME() " +
                        //    ", " + saibanHistoryNo + " " +
                        //    ",SYSDATETIME() " +
                        //    ",'" + UserInfos[0] + "' " +
                        //    ",'" + UserInfos[1] + "' " +
                        //    ",'" + UserInfos[2] + "' " +
                        //    ",'" + UserInfos[3] + "' " +
                        //    ",'調査概要を追加しました ID:" + MadoguchiID + " Garoon連携区分:" + garoon + "' " +
                        //    ",'窓口ミハル' " +
                        //    "," + MadoguchiID + " " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ",NULL " +
                        //    ")";
                        //cmd.ExecuteNonQuery();


                        //メッセージI20102を表示「データを更新しました。」 TODO


                        //応援受付（OuenUketsuke）テーブル更新
                        String kanriNo = "";
                        if (!String.IsNullOrEmpty(item1_MadoguchiKanriBangou.Text))
                        {
                            kanriNo = item1_MadoguchiKanriBangou.Text;
                        }
                        cmd.CommandText = "UPDATE OuenUketsuke SET " +
                            "OuenKanriNo = N'" + kanriNo + "' " +
                            ", OuenUpdateDate = SYSDATETIME() " +
                            ", OuenUpdateUser = N'" + UserInfos[0] + "' " +
                            ", OuenUpdateProgram = '" + pgmName + methodName + "' " +
                            " WHERE MadoguchiID = " + MadoguchiID + " ";
                        //Clipboard.SetText(cmd.CommandText);
                        cmd.ExecuteNonQuery();


                        //実施区分が3の中止の場合
                        if ("3".Equals(jisshiKubun))
                        {
                            //MadoguchiL1ChousaShinchoku が 6ではないデータを更新
                            // 6:中止　　　⇒　80：中止
                            //cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                            //"MadoguchiL1ChousaShinchoku = 6 " +
                            //",MadoguchiL1ChousaKakunin = 1 " +
                            //" WHERE MadoguchiL1ChousaShinchoku != 6 ";
                            cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET " +
                            "MadoguchiL1ChousaShinchoku = 80 " +
                            ",MadoguchiL1ChousaKakunin = 1 " +
                            ",MadoguchiL1UpdateDate = SYSDATETIME() " +
                            ",MadoguchiL1UpdateUser = N'" + UserInfos[0] + "' " +
                            ",MadoguchiL1UpdateProgram = '" + pgmName + methodName + "' " +
                            " WHERE MadoguchiL1ChousaShinchoku != 80 AND MadoguchiID = " + MadoguchiID;
                            cmd.ExecuteNonQuery();

                            //ChousaShinchokuJoukyou が 6ではないデータを更新
                            //cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                            //"ChousaShinchokuJoukyou = 6 " +
                            //",ChousaHoukokuzumi = 1 " +
                            //" WHERE ChousaShinchokuJoukyou != 6 ";
                            cmd.CommandText = "UPDATE ChousaHinmoku SET " +
                            "ChousaShinchokuJoukyou = 80 " +
                            ",ChousaHoukokuzumi = 1 " +
                            ",ChousaUpdateDate = SYSDATETIME() " +
                            ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +
                            ",ChousaUpdateProgram = '" + pgmName + methodName + "' " +
                            " WHERE ChousaShinchokuJoukyou != 80 AND MadoguchiID = " + MadoguchiID;

                            cmd.ExecuteNonQuery();
                        }


                        //窓口情報（MadoguchiJouhou）テーブルからGaroon連携対象（GaroonRenkeiKubn）を取得
                        var dt4 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "MadoguchiTantoushaCD,MadoguchiKanriGijutsusha " +
                          ",MadoguchiUketsukeBangou,MadoguchiUketsukeBangouEdaban,MadoguchiGaroonRenkei " +
                          "FROM MadoguchiJouhou " +
                          "WHERE MadoguchiID =" + MadoguchiID + "";

                        //データ取得
                        var sda4 = new SqlDataAdapter(cmd);
                        sda4.Fill(dt4);

                        String atesaki = dt4.Rows[0][0].ToString();
                        String kanriGijutusha = dt4.Rows[0][1].ToString();
                        String tokuchouNo = dt4.Rows[0][2].ToString();
                        String tokuchouNoEda = dt4.Rows[0][3].ToString();
                        String garoonOn = dt4.Rows[0][4].ToString();

                        //窓口メール送信（MadoguchiMail）テーブルからメッセージID（MadoguchiMailMessageID）を取得
                        var dt5 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "MadoguchiMailMessageID " +
                          "FROM MadoguchiMail " +
                          //"WHERE MadoguchiMailTokuchoBangou = '" + tokuchouNo + "-" + tokuchouNoEda + "'" +
                          "WHERE MadoguchiMailTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + tokuchouNo + "'" +
                          "AND MadoguchiMailTokuchoBangouEda COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_MadoguchiUketsukeBangouEdaban.Text + "' ";

                        //データ取得
                        var sda5 = new SqlDataAdapter(cmd);
                        sda5.Fill(dt5);

                        String mailMessageID = "null";

                        if (dt5.Rows.Count > 0)
                        {
                            mailMessageID = dt5.Rows[0][0].ToString();
                        }

                        //管理技術者（MadoguchiJouhou.MadoguchiKanriGijutsusha）が空でない場合
                        if (!String.IsNullOrEmpty(kanriGijutusha))
                        {
                            //宛先がnullじゃない
                            if (!String.IsNullOrEmpty(atesaki))
                            {
                                atesaki = atesaki + ";" + kanriGijutusha;
                            }
                            //宛先がnull
                            else
                            {
                                atesaki += kanriGijutusha;
                            }
                        }

                        //MadoguchiJouhouMadoguchiL1Chouを取得
                        var dt6 = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT DISTINCT " +
                          "MadoguchiL1ChousaTantoushaCD,MadoguchiL1ChousaBushoCD " +
                          "FROM MadoguchiJouhouMadoguchiL1Chou " +
                          "WHERE MadoguchiID=" + MadoguchiID + "";

                        //データ取得
                        var sda6 = new SqlDataAdapter(cmd);
                        sda6.Fill(dt6);

                        for (int i = 0; i < dt6.Rows.Count; i++)
                        {
                            //調査員担当者（MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaTantoushaCD）が空でない場合
                            String chousaTantousha = dt6.Rows[i][0].ToString();
                            if (!String.IsNullOrEmpty(chousaTantousha))
                            {
                                //宛先が空でない場合
                                if (!String.IsNullOrEmpty(atesaki))
                                {
                                    atesaki = atesaki + ";" + kanriGijutusha;
                                }
                                //宛先がnull
                                else
                                {
                                    atesaki = kanriGijutusha;
                                }
                            }

                            //調査担当部所コード（MadoguchiJouhouMadoguchiL1Chou.MadoguchiL1ChousaBushoCD）が空でない場合
                            String chousaTantoubusho = dt6.Rows[i][1].ToString();
                            if (!String.IsNullOrEmpty(chousaTantoubusho))
                            {
                                //支部応援（Mst_Shibuouen）と、調査員マスタ（Mst_Chousain）を結合し担当者を取得
                                var dt7 = new DataTable();
                                //SQL生成
                                cmd.CommandText = "SELECT " +
                                  "Mst_Chousain.KojinCD " +
                                  "FROM Mst_Shibuouen INNER JOIN Mst_Chousain ON " +
                                  "Mst_Chousain.KojinCD = Mst_Shibuouen.ShibuouenKojinCD " +
                                  //"AND Mst_Shibuouen.ShibuouenDeleteFlag = 0 " +
                                  //"AND Mst_Chousain.RetireFLG = 0 " +
                                  "AND Mst_Chousain.GyoumuBushoCD ='" + chousaTantoubusho + "' ";

                                //データ取得
                                var sda7 = new SqlDataAdapter(cmd);
                                sda7.Fill(dt7);

                                for (int j = 0; j < dt7.Rows.Count; j++)
                                {
                                    //宛先が空でない場合
                                    if (!String.IsNullOrEmpty(atesaki))
                                    {
                                        atesaki = atesaki + ";" + dt7.Rows[j][0].ToString();
                                    }
                                    //宛先がnull
                                    else
                                    {
                                        atesaki = dt7.Rows[j][0].ToString();
                                    }
                                }//for end
                            }//if end
                        }//for end

                        transaction.Commit();
                        //MessageBox.Show("更新完了", "確認");
                    }
                    catch (Exception)
                    {
                        //Clipboard.SetText(process);
                        transaction.Rollback();
                        throw;
                    }
                }//if 登録or更新end

                conn.Close();
            }
        }
        private Boolean registration_required(int tab)
        {
            //必須チェック
            Boolean requiredFlag = true;

            if (tab == 1)
            {
                //背景を白(かグレー）に戻す
                item1_MadoguchiJutakuBushoCD.BackColor = Color.FromArgb(255, 255, 255);
                item1_MadoguchiTantoushaBushoCD.BackColor = Color.FromArgb(255, 255, 255);
                item1_MadoguchiUketsukeBangouEdaban.BackColor = Color.FromArgb(255, 255, 255);
                item1_MadoguchiHachuuKikanmei.BackColor = Color.FromArgb(255, 255, 255);
                item1_MadoguchiGyoumuMeishou.BackColor = Color.FromArgb(255, 255, 255);
                item1_MadoguchiChousaKubunJibusho.BackColor = Color.FromArgb(255, 255, 255);
                item1_MadoguchiChousaKubunShibuShibu.BackColor = Color.FromArgb(255, 255, 255);
                item1_MadoguchiChousaKubunHonbuShibu.BackColor = Color.FromArgb(255, 255, 255);
                item1_MadoguchiChousaKubunShibuHonbu.BackColor = Color.FromArgb(255, 255, 255);

                // 登録日
                item1_MadoguchiTourokubi.BackColor = Color.FromArgb(255, 255, 255);
                label27.BackColor = Color.FromArgb(252, 228, 214);

                item1_MadoguchiJutakuBangou.BackColor = SystemColors.Control;
                //日付色変わらない問題～～
                item1_MadoguchiShimekiribi.BackColor = Color.FromArgb(255, 255, 255);
                label30.BackColor = Color.FromArgb(252, 228, 214);

                //受託課所支部　見た目は必須になってない
                if (String.IsNullOrEmpty(item1_MadoguchiJutakuBushoCD.Text))
                {
                    requiredFlag = false;
                    item1_MadoguchiJutakuBushoCD.BackColor = Color.FromArgb(255, 204, 255);
                }

                //窓口部所
                if (String.IsNullOrEmpty(item1_MadoguchiTantoushaBushoCD.Text))
                {
                    requiredFlag = false;
                    item1_MadoguchiTantoushaBushoCD.BackColor = Color.FromArgb(255, 204, 255);
                }

                //特調番号
                if (String.IsNullOrEmpty(item1_MadoguchiUketsukeBangouEdaban.Text))
                {
                    requiredFlag = false;
                    item1_MadoguchiUketsukeBangouEdaban.BackColor = Color.FromArgb(255, 204, 255);
                }
                else
                {
                    // 編集可能だった場合
                    if(item1_MadoguchiUketsukeBangouEdaban.ReadOnly == false)
                    {
                        // 白
                        item1_MadoguchiUketsukeBangouEdaban.BackColor = Color.FromArgb(255, 255, 255);
                    }
                    else
                    {
                        // 灰色
                        item1_MadoguchiUketsukeBangouEdaban.BackColor = Color.FromArgb(240, 240, 240);
                    }
                }

                //発注者名・課名
                if (String.IsNullOrEmpty(item1_MadoguchiHachuuKikanmei.Text))
                {
                    requiredFlag = false;
                    item1_MadoguchiHachuuKikanmei.BackColor = Color.FromArgb(255, 204, 255);
                }

                //業務名称
                if (String.IsNullOrEmpty(item1_MadoguchiGyoumuMeishou.Text))
                {
                    requiredFlag = false;
                    item1_MadoguchiGyoumuMeishou.BackColor = Color.FromArgb(255, 204, 255);
                }

                //調査区分
                if (!item1_MadoguchiChousaKubunJibusho.Checked && !item1_MadoguchiChousaKubunShibuShibu.Checked
                    && !item1_MadoguchiChousaKubunHonbuShibu.Checked && !item1_MadoguchiChousaKubunShibuHonbu.Checked)
                {
                    requiredFlag = false;
                    item1_MadoguchiChousaKubunJibusho.BackColor = Color.FromArgb(255, 204, 255);
                    item1_MadoguchiChousaKubunShibuShibu.BackColor = Color.FromArgb(255, 204, 255);
                    item1_MadoguchiChousaKubunHonbuShibu.BackColor = Color.FromArgb(255, 204, 255);
                    item1_MadoguchiChousaKubunShibuHonbu.BackColor = Color.FromArgb(255, 204, 255);
                }

                //登録日
                if (item1_MadoguchiTourokubi.CustomFormat != "")
                {
                    requiredFlag = false;
                    item1_MadoguchiTourokubi.BackColor = Color.FromArgb(255, 204, 255);
                    label27.BackColor = Color.FromArgb(255, 204, 255);

                }

                //調査担当者への締切日
                if (item1_MadoguchiShimekiribi.CustomFormat != "")
                {
                    requiredFlag = false;
                    item1_MadoguchiShimekiribi.BackColor = Color.FromArgb(255, 204, 255);
                    label30.BackColor = Color.FromArgb(255, 204, 255);
                }

                //受託のとき（ AnkenJouhouIDがあるとき）
                //受託番号が空、または0
                if ("".Equals(item1_MadoguchiJutakuBangou.Text) || "0".Equals(item1_MadoguchiJutakuBangou.Text))
                {
                    requiredFlag = false;
                    item1_MadoguchiJutakuBangou.BackColor = Color.FromArgb(255, 204, 255);
                }



                //エラーがあったらメッセージ表示
                if (!requiredFlag)
                {
                    set_error(GlobalMethod.GetMessage("E20901", ""));
                }
            }
            //施工条件必須チェック
            else if (tab == 7)
            {

                //背景を白に戻す
                item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 255, 255);
                item7_KoushuMei.BackColor = Color.FromArgb(255, 255, 255);


                //施工条件明示書ID
                if (String.IsNullOrEmpty(item7_SekouJoukenMeijishoID.Text))
                {
                    requiredFlag = false;
                    item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 204, 255);
                }

                //工種名
                if (String.IsNullOrEmpty(item7_SekouJoukenMeijishoID.Text))
                {
                    requiredFlag = false;
                    item7_KoushuMei.BackColor = Color.FromArgb(255, 204, 255);
                }

                //エラーがあったらメッセージ表示
                if (!requiredFlag)
                {
                    //メッセージE20701を表示「必須項目が未入力です 入力してください。」
                    set_error(GlobalMethod.GetMessage("E20701", ""));
                }

            }

            return requiredFlag;
        }



        private string saibanMadoguchiNo = "0";
        private Boolean registration_validate()
        {
            //データチェック
            Boolean validateFlag = true;

            //集計表フォルダファイルフォーマット確認
            if (!String.IsNullOrEmpty(item1_MadoguchiShukeiHyoFolder.Text) && !fileFormatCheck(item1_MadoguchiShukeiHyoFolder.Text))
            {
                validateFlag = false;
                item1_MadoguchiShukeiHyoFolder.BackColor = Color.FromArgb(255, 204, 255);
            }

            //報告書フォルダファイルフォーマット確認
            if (!String.IsNullOrEmpty(item1_MadoguchiShiryouHolder.Text) && !fileFormatCheck(item1_MadoguchiHoukokuShoFolder.Text))
            {
                validateFlag = false;
                item1_MadoguchiHoukokuShoFolder.BackColor = Color.FromArgb(255, 204, 255);
            }

            //調査資料フォルダファイルフォーマット確認
            if (!String.IsNullOrEmpty(item1_MadoguchiShiryouHolder.Text) && !fileFormatCheck(item1_MadoguchiShiryouHolder.Text))
            {
                validateFlag = false;
                item1_MadoguchiShiryouHolder.BackColor = Color.FromArgb(255, 204, 255);
            }

            ////実施区分　中止の場合処理終了
            //if ("中止".Equals(item1_MadoguchiJiishiKubun.Text))
            //{
            //    validateFlag = false;
            //    return validateFlag;
            //}

            // 調査区分
            // 996 チェックなし
            //// 要は支→支関連が2個以上チェックがあればエラーとする
            //if ((item1_MadoguchiChousaKubunShibuShibu.Checked && item1_MadoguchiChousaKubunHonbuShibu.Checked && item1_MadoguchiChousaKubunShibuHonbu.Checked)
            //    || (item1_MadoguchiChousaKubunShibuShibu.Checked && item1_MadoguchiChousaKubunHonbuShibu.Checked && !item1_MadoguchiChousaKubunShibuHonbu.Checked)
            //    || (item1_MadoguchiChousaKubunShibuShibu.Checked && !item1_MadoguchiChousaKubunHonbuShibu.Checked && item1_MadoguchiChousaKubunShibuHonbu.Checked)
            //    || (!item1_MadoguchiChousaKubunShibuShibu.Checked && item1_MadoguchiChousaKubunHonbuShibu.Checked && item1_MadoguchiChousaKubunShibuHonbu.Checked)
            //    )
            //{
            //    validateFlag = false;
            //    item1_MadoguchiChousaKubunJibusho.BackColor = Color.FromArgb(255, 204, 255);
            //    item1_MadoguchiChousaKubunShibuShibu.BackColor = Color.FromArgb(255, 204, 255);
            //    item1_MadoguchiChousaKubunHonbuShibu.BackColor = Color.FromArgb(255, 204, 255);
            //    item1_MadoguchiChousaKubunShibuHonbu.BackColor = Color.FromArgb(255, 204, 255);
            //    // E70076:調査区分は自部所のみ、または自部所＋（支→支、本→支、支→本のどれか）、または（支→支、本→支、支→本のどれか）をチェックしてください。
            //    set_error(GlobalMethod.GetMessage("E70076", ""));
            //}



            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();

                ////新規の場合
                //if ("".Equals(MadoguchiID))
                //{
                //    //採番番号取得
                //    var comboDt = new DataTable();
                //    //SQL生成
                //    cmd.CommandText = "SELECT " +
                //      "SaibanNo+SaibanCountupNo AS SaibanNo " +
                //      "FROM " + "M_Saiban " +
                //      "WHERE SaibanMei = 'MadoguchiId' ";

                //    //データ取得
                //    var sda = new SqlDataAdapter(cmd);
                //    sda.Fill(comboDt);

                //    DataRow dr = comboDt.Rows[0];
                //    saibanMadoguchiNo = dr[0].ToString();

                //    //採番No（saibanMadoguchiNo）を更新
                //    cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                //        saibanMadoguchiNo + " WHERE SaibanMei = 'MadoguchiId' ";

                //    cmd.ExecuteNonQuery();

                //}
                //締切日チェック
                //締切日が登録日より小さい
                if (item1_MadoguchiShimekiribi.Value.CompareTo(item1_MadoguchiTourokubi.Value) < 0)
                {
                    validateFlag = false;
                    set_error(GlobalMethod.GetMessage("E20106", ""));
                }

                //不具合No1360（1121）
                //枝番の文字チェック特殊丸数字が含まれたらエラーにする
                if (isErrorEdabanChar(item1_MadoguchiUketsukeBangouEdaban.Text))
                {
                    validateFlag = false;
                    set_error(GlobalMethod.GetMessage("E20120", ""));
                }

                //不具合No1367（1156）
                //枝番に「_」アンダーバーが含まれたらエラーにする
                if (item1_MadoguchiUketsukeBangouEdaban.Text.IndexOf("_") > 0)
                {
                    validateFlag = false;
                    set_error(GlobalMethod.GetMessage("E20121", ""));
                }

                //エラーがあったらこの後の処理は行わない
                if (!validateFlag)
                {
                    return false;
                }


                //現在入力されている特調番号・枝番号で被りがないか確認
                var tokuchoDt = new DataTable();
                //SQL生成
                cmd.CommandText = "SELECT " +
                  "MadoguchiID " +
                  "FROM MadoguchiJouhou " +
                  "WHERE MadoguchiUketsukeBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_MadoguchiUketsukeBangou.Text + "' " +
                  "AND MadoguchiUketsukeBangouEdaban COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_MadoguchiUketsukeBangouEdaban.Text + "' " + 
                  "AND MadoguchiDeleteFlag != 1 ";

                //更新の場合は現在のMadoguchiIDを除く　を付け加える
                if (!"".Equals(MadoguchiID))
                {
                    //cmd.CommandText += "AND MadoguchiID != " + MadoguchiID + " AND MadoguchiDeleteFlag != 1 ";
                    cmd.CommandText += "AND MadoguchiID != " + MadoguchiID + " ";
                }

                //データ取得
                var tokuchoSda = new SqlDataAdapter(cmd);
                tokuchoSda.Fill(tokuchoDt);

                //データ取得できた場合
                if (tokuchoDt.Rows.Count > 0)
                {
                    //メッセージE20103を表示「特調番号が重複しました。」
                    set_error(GlobalMethod.GetMessage("E20103", ""));
                    validateFlag = false;
                }

                conn.Close();
            }
            return validateFlag;
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
        public Boolean fileFormatCheck(String filePath)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(filePath, @"^[\\/]{2}[^\\^/].+[^\\^/]([\\/][^\\^/].+[^\\^/])+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {

                set_error(GlobalMethod.GetMessage("E10017", ""));
                return false;
            }
            return true;
        }

        private void item4_KyoRyokuBusho_TextChanged(object sender, EventArgs e)
        {
            if (item4_KyoRyokuBusho.Text == "")
            {
                item4_KyoryokuChou.Text = "";
                item4_KyoryokuChouCD.Text = "";
            }
            else
            {
                string BushoMei = item4_KyoRyokuBusho.Text;
                try
                {
                    SqlConnection sqlconn = new SqlConnection(connStr);
                    sqlconn.Open();
                    var cmd = sqlconn.CreateCommand();
                    //採番テーブル取得
                    var dt = new DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      "ChousainMei, KojinCD " +
                      "FROM Mst_Busho " +
                      "LEFT JOIN Mst_Chousain ON BushoShozokuChou = ChousainMei " +
                      "WHERE Mst_Busho.ShibuMei = '" + BushoMei + "' ";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);
                    ;
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        item4_KyoryokuChou.Text = dt.Rows[0][0].ToString();
                        item4_KyoryokuChouCD.Text = dt.Rows[0][1].ToString();
                    }
                }
                catch
                {

                }
            }
        }
        private String TokuchoNo_saiban()
        {
            string tokuchoNo = "1";

            //特調番号のシステム連番を取得する
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                var dt = new DataTable();
                cmd.CommandText = "SELECT ISNULL(MAX(MadoguchiSystemRenban),0) +1 AS renbanMax FROM MadoguchiJouhou " +
                "WHERE MadoguchiUketsukeBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item1_MadoguchiUketsukeBangou.Text + "' " +
                "AND MadoguchiDeleteFlag != 1 ";

                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                //連番を取得できたらそれをセットする
                tokuchoNo = dt.Rows[0][0].ToString();

            }

            return tokuchoNo;
        }

        // 再報告ボタン
        private void button_Saihoukoku_Click(object sender, EventArgs e)
        {
            string methodName = ".button_Saihoukoku_Click";
            // 新規時以外で動かす、新規時には押しても無反応（現行と同じ動作）
            if (mode != "insert") {
                //再報告処理
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();

                    DateTime date = DateTime.Now;
                    string dateString = date.ToString("yyyy/MM/dd");

                    //窓口情報の備考を更新
                    cmd.CommandText = "UPDATE MadoguchiJouhou SET " +
                        "MadoguchiBikou = N'" + item1_MadoguchiBikou.Text + " 再報告日：" + dateString + "' " +
                        "WHERE MadoguchiID = " + MadoguchiID + " ";

                    cmd.ExecuteNonQuery();

                    //更新したものをSelectしてセットする
                    var dt = new DataTable();
                    cmd.CommandText = "SELECT MadoguchiBikou FROM MadoguchiJouhou " +
                    "WHERE MadoguchiID = " + MadoguchiID + " ";

                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);
                    item1_MadoguchiBikou.Text = dt.Rows[0][0].ToString();

                    //履歴に登録処理
                    //採番No（SaibanNo）を取得
                    var dt2 = new DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      "SaibanNo+SaibanCountupNo AS SaibanNo " +
                      "FROM " + "M_Saiban " +
                      "WHERE SaibanMei = 'HistoryID' ";

                    //データ取得
                    var sda2 = new SqlDataAdapter(cmd);
                    sda2.Fill(dt2);

                    int saibanHistoryNo = int.Parse(dt2.Rows[0][0].ToString());

                    //採番No（SaibanNo）を更新
                    cmd.CommandText = "UPDATE M_Saiban SET SaibanNo = " +
                        saibanHistoryNo + " WHERE SaibanMei = 'HistoryID' ";

                    cmd.ExecuteNonQuery();

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
                        ", " + saibanHistoryNo + " " +
                        ",SYSDATETIME() " +
                        ",'" + UserInfos[0] + "' " +
                        ",N'" + UserInfos[1] + "' " +
                        ",'" + UserInfos[2] + "' " +
                        ",N'" + UserInfos[3] + "' " +
                        ",'調査概要の備考欄に再報告日を追記しました。 窓口ID = " + MadoguchiID + "' " +
                        ",'" + pgmName + methodName + "' " +
                        "," + MadoguchiID + " " +
                        ",NULL " +
                        ",NULL " +
                        ",NULL " +
                        ",NULL " +
                        ",'" + Header1.Text + "' " +
                        ")";
                    cmd.ExecuteNonQuery();

                    //調査概要の備考欄に再報告日を追記しました。 」
                    set_error(GlobalMethod.GetMessage("I20605", ""));
                    conn.Close();
                }
            }
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            if (item1_MadoguchiTantoushaBushoCD.Text != "" && !String.IsNullOrEmpty(item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString()))
            {
                //窓口部所変更に伴い所属長を取得しなおす
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    var cmd = conn.CreateCommand();
                    var dt = new DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      "BushoShozokuChou " +
                      "FROM Mst_Busho " +
                      "WHERE GyoumuBushoCD = '" + item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString() + "' ";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    if (dt.Rows.Count != 0)
                    {
                        item1_MadoguchiBushoShozokuChou.Text = dt.Rows[0][0].ToString();
                    }
                    else
                    {
                        item1_MadoguchiBushoShozokuChou.Text = "";
                    }
                }
            }
            else
            {
                item1_MadoguchiBushoShozokuChou.Text = "";
            }
        }

        private void Garoon_Click(object sender, EventArgs e)
        {
            //Garoon連携がオンだったら出力（オフは何もしない）
            //GaroonTsuikaAtesakiテーブル
            //GaroonTsuikaAtesakiID　GaroonTsuikaAtesakiMadoguchiID　GaroonTsuikaAtesakiTantoushaCD（窓口担当者）

            MessageBox.Show("今後実装されます", "未実装");
        }
        private void item6_CreateMail_Click(object sender, EventArgs e)
        {
            if (item6_TanpinMail.Text != "" && System.Text.RegularExpressions.Regex.IsMatch(item6_TanpinMail.Text, @"^((\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            {
                Process.Start("mailto:" + item6_TanpinMail.Text);
                item6_TanpinMail.BackColor = Color.FromArgb(255, 255, 255);
            }
            else
            {
                item6_TanpinMail.BackColor = Color.FromArgb(255, 204, 255);
            }
        }

        private void item6_TanpinMail_TextChanged(object sender, EventArgs e)
        {
            //if (item6_TanpinMail.Text != "" && System.Text.RegularExpressions.Regex.IsMatch(item6_TanpinMail.Text, @"^((\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
            //{
            //    item6_TanpinMail.BackColor = Color.FromArgb(255, 255, 255);
            //}
            //else
            //{
            //    item6_TanpinMail.BackColor = Color.FromArgb(255, 204, 255);
            //}
            if (item6_TanpinMail.Text != "")
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(item6_TanpinMail.Text, @"^((\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\s*))$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                {
                    item6_TanpinMail.BackColor = Color.FromArgb(255, 255, 255);
                }
                else
                {
                    item6_TanpinMail.BackColor = Color.FromArgb(255, 204, 255);
                }
            }
            else
            {
                item6_TanpinMail.BackColor = Color.FromArgb(255, 255, 255);
            }

        }

        private void button_ReCal_Click(object sender, EventArgs e)
        {
            if (ReCal())
            {
                set_error("", 0);
                set_error(GlobalMethod.GetMessage("I20607", ""));
            }
        }

        private Boolean ReCal()
        {
            int Total_Houkoku = 0;
            int Total_Irai = 0;
            decimal Total_Kingaku = 0;

            int Tmp_Houkoku = 0;
            int Tmp_Irai = 0;
            decimal Tmp_Tanka = 0;
            decimal Tmp_Kingaku = 0;
            decimal Tmp_TanpinSonotaShuukei = 0;

            for (int i = 1; i < c1FlexGrid2.Rows.Count; i++)
            {
                Tmp_Houkoku = 0;
                Tmp_Irai = 0;
                Tmp_Tanka = 0;
                Tmp_Kingaku = 0;
                // 
                if (c1FlexGrid2.Rows[i][1] != null && int.TryParse(c1FlexGrid2.Rows[i][1].ToString(), out Tmp_Houkoku))
                {
                    Total_Houkoku += Tmp_Houkoku;
                }
                // 
                if (c1FlexGrid2.Rows[i][2] != null && int.TryParse(c1FlexGrid2.Rows[i][2].ToString(), out Tmp_Irai))
                {
                    Total_Irai += Tmp_Irai;
                }
                // 
                if (c1FlexGrid2.Rows[i][3] != null && decimal.TryParse(c1FlexGrid2.Rows[i][3].ToString(), out Tmp_Tanka))
                {
                    //Tmp_Kingaku = Tmp_Houkoku * Tmp_Tanka;
                    // 報告ランク選択時（文字の色で判定）
                    if (button_RankHoukoku.ForeColor == Color.White)
                    {
                        Tmp_Kingaku = Tmp_Houkoku * Tmp_Tanka;
                    }
                    // 依頼ランク選択時（文字の色で判定）
                    else
                    {
                        Tmp_Kingaku = Tmp_Irai * Tmp_Tanka;
                    }
                }
                else
                {
                    Tmp_Kingaku = 0;
                }
                c1FlexGrid2.Rows[i][4] = Tmp_Kingaku;
                Total_Kingaku += Tmp_Kingaku;
            }

            item6_TotalHoukoku.Text = Total_Houkoku.ToString();
            item6_TotalIrai.Text = Total_Irai.ToString();
            item6_TotalKingaku.Text = GetMoneyText(Total_Kingaku);
            //item6_TanpinSeikyuuKingaku.Text = item6_TotalKingaku.Text;
            Tmp_TanpinSonotaShuukei = GetDecimal(item6_TanpinSonotaShuukei.Text);
            item6_TanpinSeikyuuKingaku.Text = GetMoneyText(Total_Kingaku + Tmp_TanpinSonotaShuukei);

            return true;
        }

        private void button_RankHoukoku_Click(object sender, EventArgs e)
        {
            // 請求確定状態では動かさない
            if (item6_TanpinSeikyuuKakutei.Checked == false)
            {
                SwichButton_Rank(1);
            }
        }

        private void button_RankIrai_Click(object sender, EventArgs e)
        {
            // 請求確定状態では動かさない
            if (item6_TanpinSeikyuuKakutei.Checked == false)
            {
                SwichButton_Rank(2);
            }
        }

        private void SwichButton_Rank(int flg = 0)
        {
            //1:報告ランク押下　2:依頼ランク押下
            //ボタンの文字色で集計対象を判断しているので色の変更は注意
            if (item6_TanpinSeikyuuKakutei.Checked == true)
            {
                item6_TanpinSonotaShuukei.ReadOnly = true;
                item6_TanpinSeikyuuKingaku.ReadOnly = true;

                button_AggregateRank.Enabled = false;
                button_AggregateRank.BackColor = Color.DarkGray;
                button_ReCal.Enabled = false;
                button_ReCal.BackColor = Color.DarkGray;

                c1FlexGrid2.Cols[1].AllowEditing = false;
                c1FlexGrid2.Cols[2].AllowEditing = false;

                //button_RankHoukoku.BackColor = Color.DimGray;
                //button_RankHoukoku.ForeColor = Color.DarkGray;
                //button_RankIrai.BackColor = Color.DimGray;
                //button_RankIrai.ForeColor = Color.DarkGray;
                button_RankHoukoku.Enabled = false;
                button_RankHoukoku.BackColor = Color.DarkGray;
                button_RankIrai.Enabled = false;
                button_RankIrai.BackColor = Color.DarkGray;
            }
            else
            {
                item6_TanpinSonotaShuukei.ReadOnly = false;
                item6_TanpinSeikyuuKingaku.ReadOnly = false;

                button_AggregateRank.Enabled = true;
                button_AggregateRank.BackColor = Color.FromArgb(42, 78, 122);
                button_ReCal.Enabled = true;
                button_ReCal.BackColor = Color.FromArgb(42, 78, 122);

                c1FlexGrid2.Cols[1].AllowEditing = true;
                c1FlexGrid2.Cols[2].AllowEditing = false;

                button_RankHoukoku.Enabled = true;
                button_RankHoukoku.BackColor = Color.FromArgb(42, 78, 122);
                button_RankHoukoku.ForeColor = Color.White;
                //button_RankIrai.BackColor = Color.DimGray;
                //button_RankIrai.ForeColor = Color.DarkGray;
                button_RankIrai.Enabled = true;
                button_RankIrai.BackColor = Color.DarkGray;
                button_RankIrai.ForeColor = Color.Black;
            }

            if (flg == 0)
            {
                ////c1FlexGrid2.Cols[1].AllowEditing = true;
                //c1FlexGrid2.Cols[2].AllowEditing = false;
                //button_RankHoukoku.BackColor = Color.FromArgb(42, 78, 122);
                //button_RankHoukoku.ForeColor = Color.White;
                ////button_RankIrai.BackColor = Color.DimGray;
                ////button_RankIrai.ForeColor = Color.DarkGray;
                //button_RankIrai.BackColor = Color.DarkGray;
                //button_RankIrai.ForeColor = Color.Black;
            }
            else if (flg == 1)
            {
                c1FlexGrid2.Cols[1].AllowEditing = true;
                c1FlexGrid2.Cols[2].AllowEditing = false;
                button_RankHoukoku.BackColor = Color.FromArgb(42, 78, 122);
                button_RankHoukoku.ForeColor = Color.White;
                //button_RankIrai.BackColor = Color.DimGray;
                //button_RankIrai.ForeColor = Color.DarkGray;
                button_RankIrai.BackColor = Color.DarkGray;
                button_RankIrai.ForeColor = Color.Black;
            }
            else if (flg == 2)
            {
                c1FlexGrid2.Cols[1].AllowEditing = false;
                c1FlexGrid2.Cols[2].AllowEditing = true;
                //button_RankHoukoku.BackColor = Color.DimGray;
                //button_RankHoukoku.ForeColor = Color.DarkGray;
                button_RankHoukoku.BackColor = Color.DarkGray;
                button_RankHoukoku.ForeColor = Color.Black;
                button_RankIrai.BackColor = Color.FromArgb(42, 78, 122);
                button_RankIrai.ForeColor = Color.White;
            }
        }

        private void button_AggregateRank_Click(object sender, EventArgs e)
        {
            AggregateRank();
        }
        // ランク集計処理
        private void AggregateRank()
        {
            try
            {
                SqlConnection sqlconn = new SqlConnection(connStr);
                sqlconn.Open();
                var cmd = sqlconn.CreateCommand();
                //採番テーブル取得
                var dt = new DataTable();
                //SQL生成
                //cmd.CommandText = "SELECT " +
                //         " TankaRankHinmoku " +
                //        ", CASE " +
                //        "    WHEN TankaRankShubetsu = 1 " +
                //        "    THEN SUM(ISNULL(C1.ChousaHoukokuHonsuu, 0)) " +
                //        "    ELSE MAX(ISNULL(C1.ChousaHoukokuHonsuu, 0)) " +
                //        "    END AS 'houkoku' " +
                //        ", CASE " +
                //        "    WHEN TankaRankShubetsu = 2 " +
                //        "    THEN SUM(ISNULL(C1.ChousaIraiHonsuu, 0)) " +
                //        "    ELSE MAX(ISNULL(C1.ChousaIraiHonsuu, 0)) " +
                //        "    END AS 'irai' " +
                //        "	,TankaKeiyakuRank.TankaRankKakaku " +
                //        ", CASE " +
                //         "   WHEN TankaRankShubetsu = 1 " +
                //         "   THEN SUM(ISNULL(C1.ChousaHoukokuHonsuu, 0)) * ISNULL(TankaRankKakaku, 0) " +
                //         "   ELSE MAX(ISNULL(C1.ChousaHoukokuHonsuu, 0)) * ISNULL(TankaRankKakaku, 0) " +
                //         "   END AS 'kakaku' " +
                //        "	,TankaKeiyakuRank.TankaRankShubetsu " +
                //        "	,TankaKeiyakuRank.TankaRankShubetsu " +
                //        "FROM MadoguchiJouhou " +
                //        "LEFT JOIN TankaKeiyaku ON TankaKeiyaku.AnkenJouhouID = MadoguchiJouhou.AnkenJouhouID " +
                //        "LEFT JOIN TankaKeiyakuRank ON TankaKeiyaku.TankaKeiyakuID = TankaKeiyakuRank.TankaKeiyakuID " +
                //        "LEFT JOIN ChousaHinmoku C1 ON C1.ChousaHoukokuRank = TankaRankHinmoku AND C1.MadoguchiID = MadoguchiJouhou.MadoguchiID " +
                //        "WHERE MadoguchiJouhou.MadoguchiID = '" + MadoguchiID + "' " +
                //        "GROUP BY TankaKeiyaku.TankaKeiyakuID , TankaRankID , TankaRankHinmoku , TankaRankShubetsu ,TankaRankKakaku " +
                //        "ORDER BY TankaRankID ";

                cmd.CommandText = "SELECT"
                                //+ " TKR.TankaRankHinmoku"
                                //+ ",CASE WHEN TKR.TankaRankShubetsu = 2"
                                //+ " THEN MAX(ISNULL(CH.ChousaHoukokuHonsuu, 0))"
                                //+ " ELSE SUM(ISNULL(CH.ChousaHoukokuHonsuu, 0))"
                                //+ " END AS 'houkoku'"
                                //+ ",CASE WHEN TKR.TankaRankShubetsu = 2"
                                //+ " THEN MAX(ISNULL(CH.ChousaIraiHonsuu, 0))"
                                //+ " ELSE SUM(ISNULL(CH.ChousaIraiHonsuu, 0))"
                                //+ " END AS 'irai'"
                                //+ ",TKR.TankaRankKakaku"
                                //+ ",CASE WHEN TKR.TankaRankShubetsu = 2"
                                //+ " THEN MAX(ISNULL(CH.ChousaHoukokuHonsuu, 0)) * ISNULL(TKR.TankaRankKakaku, 0)"
                                //+ " ELSE SUM(ISNULL(CH.ChousaHoukokuHonsuu, 0)) * ISNULL(TKR.TankaRankKakaku, 0)"
                                //+ " END AS 'kakaku'"
                                //+ ",TKR.TankaRankShubetsu"
                                //+ ",ISNULL(TNR.TanpinL1RankID, 0) AS TanpinL1RankID"
                                //+ " FROM MadoguchiJouhou MJ"
                                //+ " LEFT JOIN TanpinNyuuryoku TN ON TN.MadoguchiID = MJ.MadoguchiID"
                                //+ " LEFT JOIN TankaKeiyakuRank TKR ON TKR.TankaKeiyakuID = TN.TanpinGyoumuCD"
                                //+ " LEFT JOIN ChousaHinmoku CH ON CH.ChousaHoukokuRank = TankaRankHinmoku AND CH.MadoguchiID = MJ.MadoguchiID"
                                //+ " LEFT JOIN TanpinNyuuryokuRank TNR ON TNR.TanpinNyuuryokuID = TN.TanpinNyuuryokuID AND TNR.TanpinL1RankMei = TKR.TankaRankHinmoku"
                                //+ " WHERE MJ.MadoguchiID = '" + MadoguchiID + "'"
                                //+ " GROUP BY TKR.TankaKeiyakuID, TKR.TankaRankID, TKR.TankaRankHinmoku, TKR.TankaRankShubetsu, TKR.TankaRankKakaku, TNR.TanpinL1RankID"
                                ////+ " ORDER BY TKR.TankaKeiyakuID, TKR.TankaRankID"
                                //+ " ORDER BY TKR.TankaKeiyakuID, TKR.TankaRankHinmoku"
                                //;
                                + " TKR.TankaRankHinmoku"
                                + ",CASE WHEN TKR.TankaRankShubetsu = 2"
                                + " THEN ISNULL(CH1.HoukokuRankmax, 0)"
                                + " ELSE ISNULL(CH1.HoukokuRanksum, 0)"
                                + " END AS 'houkoku'"
                                + ",CASE WHEN TKR.TankaRankShubetsu = 2"
                                + " THEN ISNULL(CH2.IraiRankmax, 0)"
                                + " ELSE ISNULL(CH2.IraiRanksum, 0)"
                                + " END AS 'irai'"
                                + ",TKR.TankaRankKakaku"
                                ;

                // 報告ランク選択時（文字の色で判定）
                if (button_RankHoukoku.ForeColor == Color.White)
                {
                    cmd.CommandText += ",CASE WHEN TKR.TankaRankShubetsu = 2"
                                    + " THEN ISNULL(CH1.HoukokuRankmax, 0) * ISNULL(TKR.TankaRankKakaku, 0)"
                                    + " ELSE ISNULL(CH1.HoukokuRanksum, 0) * ISNULL(TKR.TankaRankKakaku, 0)"
                                    + " END AS 'kakaku'"
                                    ;
                }
                // 依頼ランク選択時（文字の色で判定）
                else
                {
                    cmd.CommandText += ",CASE WHEN TKR.TankaRankShubetsu = 2"
                                    + " THEN ISNULL(CH2.IraiRankmax, 0) * ISNULL(TKR.TankaRankKakaku, 0)"
                                    + " ELSE ISNULL(CH2.IraiRanksum, 0) * ISNULL(TKR.TankaRankKakaku, 0)"
                                    + " END AS 'kakaku'"
                                    ;
                }
                cmd.CommandText += ",TKR.TankaRankShubetsu"
                                + ",ISNULL(TNR.TanpinL1RankID, 0) AS TanpinL1RankID"
                                + " FROM MadoguchiJouhou MJ"
                                + " LEFT JOIN TanpinNyuuryoku TN"
                                + " ON TN.MadoguchiID = MJ.MadoguchiID"
                                + " LEFT JOIN TankaKeiyakuRank TKR"
                                + " ON TKR.TankaKeiyakuID = TN.TanpinGyoumuCD"
                                + " LEFT JOIN TanpinNyuuryokuRank TNR"
                                + " ON TNR.TanpinNyuuryokuID = TN.TanpinNyuuryokuID AND TNR.TanpinL1RankMei = TKR.TankaRankHinmoku"
                                + " LEFT JOIN (SELECT CH.ChousaHoukokuRank AS HoukokuRank"
                                + ",MAX(ISNULL(CH.ChousaHoukokuHonsuu, 0)) AS HoukokuRankmax"
                                + ",SUM(ISNULL(CH.ChousaHoukokuHonsuu, 0)) AS HoukokuRanksum"
                                + " FROM MadoguchiJouhou MJ"
                                + " LEFT JOIN ChousaHinmoku CH ON CH.MadoguchiID = MJ.MadoguchiID "
                                + " WHERE MJ.MadoguchiID = '" + MadoguchiID + "'" + " and CH.ChousaHoukokuRank <> ''"
                                + " GROUP BY CH.ChousaHoukokuRank ) AS CH1"
                                + " ON TKR.TankaRankHinmoku = ch1.HoukokuRank"
                                + " LEFT JOIN (SELECT CH.ChousaIraiRank AS IraiRank"
                                + ",MAX(ISNULL(CH.ChousaIraiHonsuu, 0)) AS IraiRankmax"
                                + ",SUM(ISNULL(CH.ChousaIraiHonsuu, 0)) AS IraiRanksum"
                                + " FROM MadoguchiJouhou MJ"
                                + " LEFT JOIN ChousaHinmoku CH ON CH.MadoguchiID = MJ.MadoguchiID"
                                + " WHERE MJ.MadoguchiID = '" + MadoguchiID + "'" + " and CH.ChousaIraiRank <> ''"
                                + " GROUP BY CH.ChousaIraiRank) AS CH2"
                                + " ON TKR.TankaRankHinmoku = ch2.IraiRank"
                                + " WHERE MJ.MadoguchiID = '" + MadoguchiID + "'"
                                // えんとり君修正STEP2　並び順追加
                                //+ " ORDER BY TKR.TankaKeiyakuID, TKR.TankaRankHinmoku"
                                + " ORDER BY TKR.TankaKeiyakuID, TKR.TankaRankNarabijunn, TKR.TankaRankHinmoku"
                                ;


                //データ取得
                Console.WriteLine(cmd.CommandText);
                var sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);

                //レイアウトロジックを停止する
                this.SuspendLayout();
                //描画停止
                c1FlexGrid2.BeginUpdate();

                c1FlexGrid2.Rows.Count = 1;
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        c1FlexGrid2.Rows.Add();
                        //c1FlexGrid2[i + 1, 0] = dt.Rows[i][0];
                        //c1FlexGrid2[i + 1, 1] = dt.Rows[i][1];
                        //c1FlexGrid2[i + 1, 2] = dt.Rows[i][2];
                        //c1FlexGrid2[i + 1, 3] = dt.Rows[i][3];
                        //c1FlexGrid2[i + 1, 4] = dt.Rows[i][4];
                        //c1FlexGrid2[i + 1, 5] = dt.Rows[i][5];
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            c1FlexGrid2[i + 1, j] = dt.Rows[i][j];
                        }
                    }
                }
                for (int i = 0; i < DT_TanpinRank.Rows.Count; i++)
                {
                }

                Resize_Grid("c1FlexGrid2");
                ReCal();
                set_error("", 0);

                //描画再開
                c1FlexGrid2.EndUpdate();
                //レイアウトロジックを再開する
                this.ResumeLayout();

                // 明細からランクの集計を行いました。
                set_error(GlobalMethod.GetMessage("I20606", ""));
            }
            catch
            {
                throw;
            }
        }

        private void tableLayoutPanel12_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private string GetMoneyText(decimal num)
        {
            string str = string.Format("{0:C2}", num);
            return str;
        }
        private decimal GetLong(string str)
        {
            decimal num = 0;
            decimal.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out num);
            return num;
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
        private void textbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.' && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }
        private void textBox_Enter(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            tmp = GetLong(tmp).ToString();
            ((System.Windows.Forms.TextBox)sender).Text = tmp;
        }
        private void textBox_Validated(object sender, EventArgs e)
        {
            string tmp = ((System.Windows.Forms.TextBox)sender).Text;
            ((System.Windows.Forms.TextBox)sender).Text = GetMoneyText(GetLong(tmp));
        }

        private void item6_BushoPrompt_Click(object sender, EventArgs e)
        {
            Popup_TanpinTantousya form = new Popup_TanpinTantousya();
            form.KoujiJimusyoMei = Header3.Text;
            form.MadoguchiID = MadoguchiID;
            form.ShowDialog();

            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item6_TanpinHachuubusho.Text = form.ReturnValue[0];
                item6_TanpinYakushoku.Text = form.ReturnValue[1];
                item6_TanpinHachuuTantousha.Text = form.ReturnValue[2];
                item6_TanpinTel.Text = form.ReturnValue[3];
                item6_TanpinFax.Text = form.ReturnValue[4];
                item6_TanpinMail.Text = form.ReturnValue[5];
            }
        }

        // 施工条件明示書ID
        private void textBox38_TextChanged(object sender, EventArgs e)
        {
            //施工明示書IDチェック処理
            if (!String.IsNullOrEmpty(item7_SekouJoukenMeijishoID.Text))
            {
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    //施工画面モードが新規ではない
                    if (!"0".Equals(sekouMode))
                    {
                        //施工条件取得
                        var cmd = conn.CreateCommand();
                        var dt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SekouJoukenMeijishoID " +
                          "FROM SekouJouken " +
                          "WHERE SekouJoukenID = '" + SekouJoukenID + "' " +
                          "AND MadoguchiID = '" + MadoguchiID + "' ";

                        //データ取得
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(dt);

                        //画面の明示書IDとSekouJoukenMeijishoIDが一致しないとき
                        if (dt != null && dt.Rows.Count > 0 && item7_SekouJoukenMeijishoID.Text.Equals(dt.Rows[0][0].ToString()))
                        {
                            //メッセージE20702を表示「施工条件明示書IDが変更されています。」
                            set_error(GlobalMethod.GetMessage("E20702", ""));
                        }
                    }

                    //施工画面モードが削除ではない
                    if (!"2".Equals(sekouMode))
                    {
                        //施工条件　同じ窓口IDでSekouJoukenIDが違い、かつ明示書IDが同じデータを取得
                        var cmd = conn.CreateCommand();
                        var dt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SekouJoukenID " +
                          "FROM SekouJouken " +
                          "WHERE MadoguchiID = " + MadoguchiID + " " +
                          "AND SekouJoukenMeijishoID COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item7_SekouJoukenMeijishoID.Text + "'";

                        //新規（SekouJoukenIDがない）じゃないとき
                        if (!"0".Equals(sekouMode))
                        {
                            cmd.CommandText += "AND SekouJoukenID != " + SekouJoukenID + " ";
                        }

                        //データ取得
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(dt);

                        //1件（以上）取得したらエラー
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            //メッセージE20703を表示「施工条件明示書IDが重複しています。」
                            set_error(GlobalMethod.GetMessage("E20703", ""));

                            //背景色変更
                            item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 204, 255);
                        }
                        else
                        {
                            //取得できない場合問題なし
                            //背景色変更
                            item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 255, 255);
                        }

                    }//if 画面モード

                }
            }//if IsNullOrEmpty
        }

        // 施工条件明示書ID
        private void item7_SekouJoukenMeijishoID_Leave(object sender, EventArgs e)
        {
            //if (sekouMeijishoComboChangeFlg == "1")
            //{
            //    // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
            //    sekouMeijishoComboChangeFlg = "0";
            //    return;
            //}
            if (sekouMeijishoIDChangeFlg == "1")
            {
                // 施工条件タブ 施工条件明示書ID変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoIDChangeFlg = "0";
                return;
            }

            set_error("", 0);
            //施工明示書IDチェック処理
            if (!String.IsNullOrEmpty(item7_SekouJoukenMeijishoID.Text))
            {
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    //施工画面モードが新規ではない
                    if (!"0".Equals(sekouMode))
                    {
                        //施工条件取得
                        var cmd = conn.CreateCommand();
                        var dt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SekouJoukenMeijishoID " +
                          "FROM SekouJouken " +
                          "WHERE SekouJoukenID = '" + SekouJoukenID + "' " +
                          "AND MadoguchiID = '" + MadoguchiID + "' ";

                        //データ取得
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(dt);
                            
                        //画面の明示書IDとSekouJoukenMeijishoIDが一致しないとき
                        if (dt != null && dt.Rows.Count > 0 && !item7_SekouJoukenMeijishoID.Text.Equals(dt.Rows[0][0].ToString()))
                        {
                            //メッセージE20702を表示「施工条件明示書IDが変更されています。」
                            set_error(GlobalMethod.GetMessage("E20702", ""));
                        }
                    }

                    //施工画面モードが削除ではない
                    if (!"2".Equals(sekouMode))
                    {
                        //施工条件　同じ窓口IDでSekouJoukenIDが違い、かつ明示書IDが同じデータを取得
                        var cmd = conn.CreateCommand();
                        var dt = new DataTable();
                        //SQL生成
                        cmd.CommandText = "SELECT " +
                          "SekouJoukenID " +
                          "FROM SekouJouken " +
                          "WHERE MadoguchiID = " + MadoguchiID + " " +
                          "AND SekouJoukenMeijishoID COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + item7_SekouJoukenMeijishoID.Text + "'";

                        //新規（SekouJoukenIDがない）じゃないとき
                        if (!"0".Equals(sekouMode))
                        {
                            cmd.CommandText += "AND SekouJoukenID != " + SekouJoukenID + " ";
                        }

                        //データ取得
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(dt);

                        //1件（以上）取得したらエラー
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            //メッセージE20703を表示「施工条件明示書IDが重複しています。」
                            set_error(GlobalMethod.GetMessage("E20703", ""));

                            //背景色変更
                            item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 204, 255);
                        }
                        else
                        {
                            //取得できない場合問題なし
                            //背景色変更
                            item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 255, 255);
                        }

                    }//if 画面モード

                }
            }//if IsNullOrEmpty
        }

        // 明示書切り替え
        private void comboBox22_TextChanged(object sender, EventArgs e)
        {
            // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
            if (sekouMeijishoComboChangeFlg == "1")
            {
                sekouMeijishoComboChangeFlg = "0";
                return;
            }

            // 施工条件タブを開いているか 0:開いてない 1:開いている または件数が0の場合
            if (openSekouTab != "1" || item7_TourokuSuu.Text == "0")
            {
                return;
            }

            // 施工条件明示書IDと明示書切り替えが同じ場合
            if (item7_SekouJoukenMeijishoID.Text == item7_MeijishoKirikaeCombo.Text)
            {
                return;
            }

            // 773:特に空欄選択では何も起こらない。
            // とのことなので、明示書切り替えが空の場合、何も処理をしない
            if (item7_MeijishoKirikaeCombo.Text == "")
            {
                return;
            }

            sekouMeijishoComboChangeFlg = "0";

            //明示書切り替え
            //メッセージI20704を表示「表示中のデータをクリアし、\\n選択した明示書IDのデータを表示してもよろしいですか？」
            //メッセージI20704を表示「表示中のデータをクリアし、選択した明示書IDのデータを表示してもよろしいですか？」
            if (MessageBox.Show(GlobalMethod.GetMessage("I20704", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                //施工条件テーブル（SekouJouken）にMadoguchiIDで検索し、SekouJoukenIDの最小値を取得する
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    //施工条件取得
                    var cmd = conn.CreateCommand();
                    var dt = new DataTable();
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                      "ISNULL(MIN(SekouJoukenID),999999999999999) " +
                      "FROM SekouJouken " +
                      "WHERE MadoguchiID = " + MadoguchiID + " ";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    string minSekouJoukenID = dt.Rows[0][0].ToString();

                    //施工条件テーブル（SekouJouken）取得
                    //SQL生成
                    cmd.CommandText = "SELECT " +
                        "cd.countData " +
                        ",sj.SekouJoukenID " +
                        ",sj.SekouJoukenMeijishoID " + //施工条件明示書ID
                        ",SekouKoushuMei " + //工種名
                                             //◆施工条件（旧）
                        ",SekouTenpuUmu " + //①施工計画書添付の有無
                        ",SekouGenbaHeimenzu " + //②その他添付資料の現場平面図
                        ",SekouDoshituKankeizu " + //②その他添付資料の土質関係図
                        ",SekouSuuryouKeisanzu " + //②その他添付資料の数量計算書
                        ",SekouHiruma " + //③施工時間帯指定の昼間
                        ",SekouYakan " + //③施工時間帯指定の夜間
                        ",SekouKiseiAri " + //③施工時間帯指定の規制有り
                        ",SekouSagyouKouritsu " + //④施工条件他の作業効率
                        ",SekouKikai " + //④施工条件他の施工機械の搬入経路
                        ",SekouKasetu " + //④施工条件他の仮設条件
                        ",SekouShizai " + //④施工条件他の資材搬入
                        ",SekouKensetsu " + //⑤建設機械スペック指定
                        ",SekouSuichuu " + //⑥水中施行条件
                        ",SekouSonota " + //⑦その他
                        ",SekouMemo1 " +//メモ1
                        ",SekouMemo2 " +//メモ2
                                        //◆施工条件
                        ",SekouTenpuUmup1Ichizu01 " + //3.添付資料の位置図
                        ",SekouTenpuUmup1Sekou02 " + //3.添付資料の施工計画書
                        ",SekouTenpuUmup1Sankou03 " + //3.添付資料の参考カタログ
                        ",SekouTenpuUmup1Ippan04 " + //3.添付資料の一般図・平面図
                        ",SekouTenpuUmup1Genba05 " + //3.添付資料の現場写真
                        ",SekouTenpuUmup1Kako06 " + //3.添付資料の過去報告書
                        ",SekouTenpuUmup1Shousai07 " + //3.添付資料の詳細図
                        ",SekouTenpuUmup1Doshitu08 " + //3.添付資料の土質関係図（柱状図等）
                        ",SekouTenpuUmup1Sonota09 " + //3.添付資料のその他
                        ",SekouTenpuUmup1Suuryou10 " + //3.添付資料の数量計算書
                        ",SekouTenpuUmup1Unpan11 " + //3.添付資料の運搬ルート図
                        ",SekouSekou2Rikujou01 " + //5.(1)施工場所の陸上
                        ",SekouSekou2Suijou02 " + //5.(1)施工場所の水上
                        ",SekouSekou2Suichuu03 " + //5.(1)施工場所の水中
                        ",SekouSekou2Sonota04 " + //5.(1)施工場所のその他
                        ",SekouSekou3Tsuujou01 " + //5.(2)施工時間帯の通常昼間施工（8:00~17:00）
                        ",SekouSekou3Tsuujou02 " + //5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                        ",SekouSekou3Sekou03 " + //5.(2)施工時間帯の施工時間規制あり
                        ",SekouSekou3Nihou04 " + //5.(2)施工時間帯の二方施工（2交代制 昼夜連続施工）
                        ",SekouSekou3Sanpou05 " + //5.(2)施工時間帯の三方施工（3交代制 24時間施工）
                        ",SekouSagyou4Kankyou01 " + //5.(3)作業環境の現場が狭隘
                        ",SekouSagyou4Sekou02 " + //5.(3)作業環境の施工箇所が点在
                        ",SekouSagyou4Joukuu03 " + //5.(3)作業環境の上空制限あり
                        ",SekouSagyou4Sonota04 " + //5.(3)作業環境のその他
                        ",SekouSagyou4Jinka05 " + //5.(3)作業環境の人家に近接（近接施工）
                        ",SekouSagyou4Tokki06 " + //5.(3)作業環境の特記すべき条件なし
                        ",SekouSagyou4Kankyou07 " + //5.(3)作業環境の環境対策あり（騒音・振動）
                        ",SekouSagyou5Koutusu01 " + //5.(4)施工機械・資材搬入経路の交通規制あり
                        ",SekouSagyou5Hannyuu02 " + //5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                        ",SekouSagyou5Sonota03 " + //5.(4)施工機械・資材搬入経路のその他
                        ",SekouSagyou5Tokki04 " + //5.(4)施工機械・資材搬入経路の特記すべき条件なし
                        ",SekouKasetsu6Shitei01 " + //5.(5)仮設条件の指定あり
                        ",SekouKasetsu6Shitei02 " + //5.(5)仮設条件の特記すべき条件なし
                        ",SekouSekou7Shitei01 " + //5.(6)施工機械スペック指定の指定あり
                        ",SekouSekou7Shitei02 " + //5.(6)施工機械スペック指定の指定なし
                        ",SekouSonota8Shitei01 " + //5.(7)その他条件の指定あり
                        ",SekouSonota8Shitei02 " + //5.(7)その他条件の特記すべき条件なし
                        ",SekouSonotaMemo03 " + //メモ
                        "FROM SekouJouken sj " +
                        ",(SELECT COUNT (*) AS countData " +
                        "FROM SekouJouken WHERE MadoguchiID = " + MadoguchiID + " AND SekouDeleteFlag != 1) cd " +
                        "WHERE sj.MadoguchiID = " + MadoguchiID + " " +
                        "AND sj.SekouDeleteFlag != 1 ";

                    //明示書切り替えが空の場合は
                    if (item7_MeijishoKirikaeCombo.Text != null && String.IsNullOrEmpty(item7_MeijishoKirikaeCombo.SelectedValue.ToString()))
                    {
                        //minSekouJoukenIDでSekouJoukenIDを検索
                        cmd.CommandText += "AND SekouJoukenID = " + minSekouJoukenID + " ";
                        SekouJoukenID = minSekouJoukenID;
                    }
                    //空でない場合は
                    else
                    {
                        //明示書切り替えの値でSekouJoukenIDを検索
                        cmd.CommandText += "AND SekouJoukenID = " + item7_MeijishoKirikaeCombo.SelectedValue.ToString() + " ";
                        SekouJoukenID = item7_MeijishoKirikaeCombo.SelectedValue.ToString();
                    }

                    //データ取得
                    DT_Sekou.Clear();
                    var sda2 = new SqlDataAdapter(cmd);
                    sda2.Fill(DT_Sekou);

                    // データが見つからない場合、施工条件IDを空に
                    if(DT_Sekou != null && DT_Sekou.Rows.Count > 0)
                    {

                    }
                    else
                    {
                        SekouJoukenID = "";
                    }

                    //モードを更新モードにする
                    sekouMode = "1";

                    //抽出したデータを画面に表示
                    set_data(7);

                    //// 施工条件明示書ID変更フラグ 0:手動変更 1:システム側で変更
                    //sekouMeijishoComboChangeFlg = "1";
                    //// 明示書切り替えの制御の為、ここでTextChangedを動かす
                    //item7_MeijishoKirikaeCombo.Text = item7_MeijishoKirikaeCombo.Text;

                    ////モードを更新モードにする
                    //sekouMode = "1";

                    //追加、削除ボタン活性
                    item7_btnAdd.Enabled = true;
                    item7_btnDelete.Enabled = true;

                    //色変更
                    item7_btnAdd.BackColor = Color.FromArgb(42, 78, 122);
                    item7_btnAdd.ForeColor = Color.FromArgb(255, 255, 255);
                    item7_btnDelete.BackColor = Color.FromArgb(42, 78, 122);
                    item7_btnDelete.ForeColor = Color.FromArgb(255, 255, 255);

                }//SqlConnection end
            }//DialogResult.OK end
            else
            {
                // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoComboChangeFlg = "1";
                // いいえ時
                if (SekouJoukenID != "") 
                { 
                    item7_MeijishoKirikaeCombo.SelectedValue = SekouJoukenID;
                }
                else
                {
                    item7_MeijishoKirikaeCombo.SelectedIndex = -1;
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

        // 管理番号TextChanged
        private void KanriBangou_TextChanged(object sender, EventArgs e)
        {
            // 管理番号を応援受付に移す
            item5_Kanribangou.Text = item1_MadoguchiKanriBangou.Text;
        }

        // 応援受付タブ 応援状況チェック
        private void Uketsuke_CheckStateChange(object sender, EventArgs e)
        {
            //チェック状態の確認
            switch (((CheckBox)(sender)).CheckState)
            {
                case CheckState.Checked:
                    // チェック
                    item5_UketsukeDate.Value = System.DateTime.Today;
                    item5_UketsukeDate.CustomFormat = "";

                    if (item1_MadoguchiChousaKubunShibuHonbu.Checked) {
                        UketsukeIcon.Visible = true;
                    }
                    // チェック時は完了アイコン
                    UketsukeIcon.Image = Image.FromFile("Resource/kan.png");
                    break;
                case CheckState.Unchecked:
                    // 未チェック
                    item5_UketsukeDate.Value = System.DateTime.Today;
                    item5_UketsukeDate.CustomFormat = " ";
                    if (DT_Ouenuketsuke.Rows[0][1].ToString() == "0" || !item1_MadoguchiChousaKubunShibuHonbu.Checked)
                    {
                        UketsukeIcon.Visible = false;
                    }
                    else
                    {
                        UketsukeIcon.Visible = true;
                        UketsukeIcon.Image = Image.FromFile("Resource/OnegaiIcon35px.png");
                    }
                    break;
            }
        }

        // 応援完了
        private void OuenKanryou_CheckStateChange(object sender, EventArgs e)
        {
            //チェック状態の確認
            switch (((CheckBox)(sender)).CheckState)
            {
                case CheckState.Checked:
                    // チェック
                    KanryouIcon.Image = Image.FromFile("Resource/kan.png");
                    item5_OuenKanryoDate.Value = System.DateTime.Today;
                    item5_OuenKanryoDate.CustomFormat = "";
                    if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
                    {
                        KanryouIcon.Visible = true;
                    }
                    break;
                case CheckState.Unchecked:
                    // 未チェック
                    item5_OuenKanryoDate.Value = System.DateTime.Today;
                    item5_OuenKanryoDate.CustomFormat = " ";
                    if (DT_Ouenuketsuke.Rows[0][1].ToString() == "0" || !item1_MadoguchiChousaKubunShibuHonbu.Checked)
                    {
                        KanryouIcon.Visible = false;
                    }
                    else
                    {
                        KanryouIcon.Visible = true;
                    }
                    break;
            }
        }
        // 応援受付更新
        private void btnOuenuketsukeUpdate_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("更新を行いますが宜しいですか？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                UpdateMadoguchi(5);
            }
        }

        //売上年度変更時
        private void item1_3_TextChanged(object sender, EventArgs e)
        {
            //有効期間のあるテーブルから選択肢を再取得
            get_combo_byNendo();
        }

        //　施工条件タブ 追加ボタン
        private void button33_Click(object sender, EventArgs e)
        {
            //施工条件　追加ボタン処理
            set_error("", 0);
            item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 255, 255);
            item7_KoushuMei.BackColor = Color.FromArgb(255, 255, 255);

            //メッセージI20702を表示「追加しますがよろしいですか？」
            if (MessageBox.Show(GlobalMethod.GetMessage("I20702", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                //モードを新規モードにする
                sekouMode = "0";

                //追加、削除ボタン非活性
                item7_btnAdd.Enabled = false;
                item7_btnDelete.Enabled = false;

                //色変更
                item7_btnAdd.BackColor = Color.FromArgb(105, 105, 105);
                item7_btnAdd.ForeColor = Color.FromArgb(169, 169, 169);
                item7_btnDelete.BackColor = Color.FromArgb(105, 105, 105);
                item7_btnDelete.ForeColor = Color.FromArgb(169, 169, 169);

                //画面項目を全て初期値にする
                //登録数 施工条件ID
                item7_TourokuSuu.Text = "0";
                SekouJoukenID = "";

                //明示ID　工種名
                // 施工条件タブ 施工条件明示書ID変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoIDChangeFlg = "1";
                item7_SekouJoukenMeijishoID.Text = "";
                item7_KoushuMei.Text = "";

                //◆施工条件旧
                //①施工計画書添付の有無 　
                checkBox80.Checked = false;
                //②その他添付資料の現場平面図　②その他添付資料の土質関係図　②その他添付資料の数量計算書
                checkBox81.Checked = false;
                checkBox82.Checked = false;
                checkBox83.Checked = false;
                //③施工時間帯指定の昼間　③施工時間帯指定の夜間　③施工時間帯指定の規制有り
                checkBox84.Checked = false;
                checkBox85.Checked = false;
                checkBox86.Checked = false;
                //④施工条件他の作業効率　④施工条件他の施工機械の搬入経路
                checkBox87.Checked = false;
                checkBox88.Checked = false;
                //④施工条件他の仮設条件　④施工条件他の資材搬入
                checkBox89.Checked = false;
                checkBox93.Checked = false;
                //⑤建設機械スペック指定　⑥水中施行条件　⑦その他
                checkBox90.Checked = false;
                checkBox91.Checked = false;
                checkBox92.Checked = false;
                //メモ1　メモ2
                textBox41.Text = "";
                textBox42.Text = "";

                //◆施工条件
                //3.添付資料の位置図 3.添付資料の施工計画書 3.添付資料の参考カタログ
                checkBox43.Checked = false;
                checkBox47.Checked = false;
                checkBox51.Checked = false;

                //3.添付資料の一般図・平面図 3.添付資料の現場写真 3.添付資料の過去報告書
                checkBox44.Checked = false;
                checkBox48.Checked = false;
                checkBox52.Checked = false;

                //3.添付資料の詳細図 3.添付資料の土質関係図（柱状図等）3.添付資料のその他
                checkBox45.Checked = false;
                checkBox49.Checked = false;
                checkBox53.Checked = false;

                //3.添付資料の数量計算書 3.添付資料の運搬ルート図
                checkBox46.Checked = false;
                checkBox50.Checked = false;

                //5.(1)施工場所の陸上 5.(1)施工場所の水上 
                checkBox54.Checked = false;
                checkBox55.Checked = false;

                //5.(1)施工場所の水中 5.(1)施工場所のその他
                checkBox56.Checked = false;
                checkBox57.Checked = false;

                //5.(2)施工時間帯の通常昼間施工（8:00~17:00） 5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                checkBox58.Checked = false;
                checkBox60.Checked = false;

                //5.(2)施工時間帯の施工時間規制あり 5.(2)施工時間帯の二方施工（2交代制 昼夜連続施工）
                checkBox62.Checked = false;
                checkBox59.Checked = false;

                //5.(2)施工時間帯の三方施工（3交代制 24時間施工）
                checkBox61.Checked = false;

                //5.(3)作業環境の現場が狭隘  5.(3)作業環境の施工箇所が点在 5.(3)作業環境の上空制限あり
                checkBox63.Checked = false;
                checkBox67.Checked = false;
                checkBox64.Checked = false;

                //5.(3)作業環境のその他 5.(3)作業環境の人家に近接（近接施工） 5.(3)作業環境の特記すべき条件なし
                checkBox68.Checked = false;
                checkBox65.Checked = false;
                checkBox70.Checked = false;

                //5.(3)作業環境の環境対策あり（騒音・振動）
                checkBox66.Checked = false;

                //5.(4)施工機械・資材搬入経路の交通規制あり 5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                checkBox69.Checked = false;
                checkBox71.Checked = false;

                //5.(4)施工機械・資材搬入経路のその他 5.(4)施工機械・資材搬入経路の特記すべき条件なし
                checkBox72.Checked = false;
                checkBox73.Checked = false;

                //5.(5)仮設条件の指定あり 5.(5)仮設条件の特記すべき条件なし 
                checkBox74.Checked = false;
                checkBox75.Checked = false;

                //5.(6)施工機械スペック指定の指定あり 5.(6)施工機械スペック指定の指定なし 
                checkBox76.Checked = false;
                checkBox77.Checked = false;

                //5.(7)その他条件の指定あり  5.(7)その他条件の特記すべき条件なし
                checkBox78.Checked = false;
                checkBox79.Checked = false;

                //メモ
                textBox40.Text = "";

                //施工条件　明示書切替
                string discript = "SekouJoukenMeijishoID ";
                string value = "SekouJoukenID ";
                string table = "SekouJouken ";
                string where = "MadoguchiID = " + MadoguchiID + " AND SekouDeleteFlag != 1 ";
                DataTable tmpdt = new DataTable();
                tmpdt = GlobalMethod.getData(discript, value, table, where);
                if (tmpdt != null)
                {
                    //空白行追加
                    DataRow dr = tmpdt.NewRow();
                    tmpdt.Rows.InsertAt(dr, 0);
                }
                // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoComboChangeFlg = "1";
                item7_MeijishoKirikaeCombo.DataSource = tmpdt;
                item7_MeijishoKirikaeCombo.DisplayMember = "Discript";
                item7_MeijishoKirikaeCombo.ValueMember = "Value";

                // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoComboChangeFlg = "1";
                item7_MeijishoKirikaeCombo.SelectedIndex = -1;

                discript = "count(*) ";
                value = "count(*) ";
                table = "SekouJouken ";
                where = "MadoguchiID = " + MadoguchiID + " AND SekouDeleteFlag != 1 ";
                DataTable tmpdt2 = new DataTable();
                tmpdt2 = GlobalMethod.getData(discript, value, table, where);
                if (tmpdt2 != null)
                {
                    item7_TourokuSuu.Text = tmpdt2.Rows[0][0].ToString();
                }
            }
        }

        // 施工条件タブ 削除ボタン
        private void button34_Click(object sender, EventArgs e)
        {
            //施工条件　削除ボタン処理
            set_error("", 0);
            item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 255, 255);
            item7_KoushuMei.BackColor = Color.FromArgb(255, 255, 255);

            //メッセージI20703を表示「削除を行いますがよろしいですか？」
            if (MessageBox.Show(GlobalMethod.GetMessage("I20703", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                //モードを削除モードにする
                sekouMode = "2";

                //施工条件テーブル（SekouJouken）削除
                //SekouJoukenID、MadoguchiID、SekouJoukenMeijishoIDで検索
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;

                    cmd.CommandText = "DELETE FROM SekouJouken " +
                        "WHERE SekouJoukenID = " + SekouJoukenID + " " +
                        "AND MadoguchiID = " + MadoguchiID + " ";
                        // 施工条件明示書IDは変えられる為、条件から除外
                        //"AND SekouJoukenMeijishoID = '" + item7_SekouJoukenMeijishoID.Text + "' ";

                    int deleteCount = cmd.ExecuteNonQuery();

                    //データが見つからなかった場合　
                    if (deleteCount == 0)
                    {
                        //エラーメッセージを表示「対象データは存在しません。」
                        set_error(GlobalMethod.GetMessage("E10009", ""));
                    }
                    //削除データがあった場合
                    else
                    {
                        //メッセージI20707を表示「施工条件を削除しました」
                        set_error(GlobalMethod.GetMessage("I20707", ""));
                    }
                    transaction.Commit();
                    conn.Close();
                }//SqlConnection end

                SekouJoukenID = "";

                // モードを新規モードにする
                sekouMode = "0";

                //追加、削除ボタン活性
                item7_btnAdd.Enabled = false;
                item7_btnDelete.Enabled = false;

                //色変更
                //item7_btnAdd.BackColor = Color.FromArgb(42, 78, 122);
                //item7_btnAdd.ForeColor = Color.FromArgb(255, 255, 255);
                //item7_btnDelete.BackColor = Color.FromArgb(42, 78, 122);
                //item7_btnDelete.ForeColor = Color.FromArgb(255, 255, 255);

                //色変更
                item7_btnAdd.BackColor = Color.FromArgb(105, 105, 105);
                item7_btnAdd.ForeColor = Color.FromArgb(169, 169, 169);
                item7_btnDelete.BackColor = Color.FromArgb(105, 105, 105);
                item7_btnDelete.ForeColor = Color.FromArgb(169, 169, 169);

                // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoComboChangeFlg = "1";
                item7_MeijishoKirikaeCombo.SelectedIndex = -1;
                // 施工条件タブ 施工条件明示書ID変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoIDChangeFlg = "1";
                item7_SekouJoukenMeijishoID.Text = "";
                item7_KoushuMei.Text = "";

                //①施工計画書添付の有無 　
                checkBox80.Checked = false;
                //②その他添付資料の現場平面図　②その他添付資料の土質関係図　②その他添付資料の数量計算書
                checkBox81.Checked = false;
                checkBox82.Checked = false;
                checkBox83.Checked = false;
                //③施工時間帯指定の昼間　③施工時間帯指定の夜間　③施工時間帯指定の規制有り
                checkBox84.Checked = false;
                checkBox85.Checked = false;
                checkBox86.Checked = false;
                //④施工条件他の作業効率　④施工条件他の施工機械の搬入経路
                checkBox87.Checked = false;
                checkBox88.Checked = false;
                //④施工条件他の仮設条件　④施工条件他の資材搬入
                checkBox89.Checked = false;
                checkBox93.Checked = false;
                //⑤建設機械スペック指定　⑥水中施行条件　⑦その他
                checkBox90.Checked = false;
                checkBox91.Checked = false;
                checkBox92.Checked = false;
                //メモ1　メモ2
                textBox41.Text = "";
                textBox42.Text = "";

                //◆施工条件
                //3.添付資料の位置図 3.添付資料の施工計画書 3.添付資料の参考カタログ
                checkBox43.Checked = false;
                checkBox47.Checked = false;
                checkBox51.Checked = false;

                //3.添付資料の一般図・平面図 3.添付資料の現場写真 3.添付資料の過去報告書
                checkBox44.Checked = false;
                checkBox48.Checked = false;
                checkBox52.Checked = false;

                //3.添付資料の詳細図 3.添付資料の土質関係図（柱状図等）3.添付資料のその他
                checkBox45.Checked = false;
                checkBox49.Checked = false;
                checkBox53.Checked = false;

                //3.添付資料の数量計算書 3.添付資料の運搬ルート図
                checkBox46.Checked = false;
                checkBox50.Checked = false;

                //5.(1)施工場所の陸上 5.(1)施工場所の水上 
                checkBox54.Checked = false;
                checkBox55.Checked = false;

                //5.(1)施工場所の水中 5.(1)施工場所のその他
                checkBox56.Checked = false;
                checkBox57.Checked = false;

                //5.(2)施工時間帯の通常昼間施工（8:00~17:00） 5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                checkBox58.Checked = false;
                checkBox60.Checked = false;

                //5.(2)施工時間帯の施工時間規制あり 5.(2)施工時間帯の二方施工（2交代制 昼夜連続施工）
                checkBox62.Checked = false;
                checkBox59.Checked = false;

                //5.(2)施工時間帯の三方施工（3交代制 24時間施工）
                checkBox61.Checked = false;

                //5.(3)作業環境の現場が狭隘  5.(3)作業環境の施工箇所が点在 5.(3)作業環境の上空制限あり
                checkBox63.Checked = false;
                checkBox67.Checked = false;
                checkBox64.Checked = false;

                //5.(3)作業環境のその他 5.(3)作業環境の人家に近接（近接施工） 5.(3)作業環境の特記すべき条件なし
                checkBox68.Checked = false;
                checkBox65.Checked = false;
                checkBox70.Checked = false;

                //5.(3)作業環境の環境対策あり（騒音・振動）
                checkBox66.Checked = false;

                //5.(4)施工機械・資材搬入経路の交通規制あり 5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                checkBox69.Checked = false;
                checkBox71.Checked = false;

                //5.(4)施工機械・資材搬入経路のその他 5.(4)施工機械・資材搬入経路の特記すべき条件なし
                checkBox72.Checked = false;
                checkBox73.Checked = false;

                //5.(5)仮設条件の指定あり 5.(5)仮設条件の特記すべき条件なし 
                checkBox74.Checked = false;
                checkBox75.Checked = false;

                //5.(6)施工機械スペック指定の指定あり 5.(6)施工機械スペック指定の指定なし 
                checkBox76.Checked = false;
                checkBox77.Checked = false;

                //5.(7)その他条件の指定あり  5.(7)その他条件の特記すべき条件なし
                checkBox78.Checked = false;
                checkBox79.Checked = false;

                //メモ
                textBox40.Text = "";

                //施工条件　明示書切替
                string discript = "SekouJoukenMeijishoID ";
                string value = "SekouJoukenID ";
                string table = "SekouJouken ";
                string where = "MadoguchiID = " + MadoguchiID + " AND SekouDeleteFlag != 1 ";
                DataTable tmpdt = new DataTable();
                tmpdt = GlobalMethod.getData(discript, value, table, where);
                if (tmpdt != null)
                {
                    //空白行追加
                    DataRow dr = tmpdt.NewRow();
                    tmpdt.Rows.InsertAt(dr, 0);
                }
                // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoComboChangeFlg = "1";
                item7_MeijishoKirikaeCombo.DataSource = tmpdt;
                item7_MeijishoKirikaeCombo.DisplayMember = "Discript";
                item7_MeijishoKirikaeCombo.ValueMember = "Value";

                discript = "count(*) ";
                value = "count(*) ";
                table = "SekouJouken ";
                where = "MadoguchiID = " + MadoguchiID + " AND SekouDeleteFlag != 1 ";
                DataTable tmpdt2 = new DataTable();
                tmpdt2 = GlobalMethod.getData(discript, value, table, where);
                if (tmpdt2 != null)
                {
                    item7_TourokuSuu.Text = tmpdt2.Rows[0][0].ToString();
                }

                //// モード判定
                //if(item7_TourokuSuu.Text == "0") 
                //{ 
                //    // モードを新規モードにする
                //    sekouMode = "0";
                //}
                //else
                //{
                //    // モードを更新モードにする
                //    sekouMode = "1";
                //}
            }
        }

        // 調査区分 自部所
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            //調査区分のチェックが自部所についた
            //if (item1_MadoguchiChousaKubunJibusho.Checked)
            //{
            //    //item1_MadoguchiChousaKubunShibuShibu.Checked = false;
            //    //item1_MadoguchiChousaKubunHonbuShibu.Checked = false;
            //    //item1_MadoguchiChousaKubunShibuHonbu.Checked = false;
            //}
            ////チェックが自部所から外れた
            //else
            //{
            //    //自部所以外がチェックになった場合
            //    if (item1_MadoguchiChousaKubunShibuShibu.Checked || item1_MadoguchiChousaKubunHonbuShibu.Checked || item1_MadoguchiChousaKubunShibuHonbu.Checked)
            //    {
            //        //実施区分を打診中に変更する
            //        item1_MadoguchiJiishiKubun.SelectedValue = "2";
            //    }
            //}
            ////実施区分を打診中に変更する
            //item1_MadoguchiJiishiKubun.SelectedValue = "2";
        }

        // 調査区分 支→支
        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            //調査区分のチェックが支部支部についた
            //if (item1_MadoguchiChousaKubunShibuShibu.Checked)
            //{
            //    //item1_MadoguchiChousaKubunJibusho.Checked = false;
            //    //item1_MadoguchiChousaKubunHonbuShibu.Checked = false;
            //    //item1_MadoguchiChousaKubunShibuHonbu.Checked = false;

            //    //実施区分を打診中に変更する
            //    item1_MadoguchiJiishiKubun.SelectedValue = "2";

            //}
            if (item1_MadoguchiChousaKubunShibuShibu.Checked && item1_MadoguchiJiishiKubun.SelectedValue.ToString() == "1")
            {
                //実施区分を打診中に変更する
                item1_MadoguchiJiishiKubun.SelectedValue = "2";
            }
        }

        private void checkBox31_CheckedChanged(object sender, EventArgs e)
        {
            ////調査区分のチェックが本社支部についた
            //if (item1_MadoguchiChousaKubunHonbuShibu.Checked)
            //{
            //    //item1_MadoguchiChousaKubunJibusho.Checked = false;
            //    //item1_MadoguchiChousaKubunShibuShibu.Checked = false;
            //    //item1_MadoguchiChousaKubunShibuHonbu.Checked = false;

            //    //実施区分を打診中に変更する
            //    item1_MadoguchiJiishiKubun.SelectedValue = "2";
            //}
            if (item1_MadoguchiChousaKubunHonbuShibu.Checked && item1_MadoguchiJiishiKubun.SelectedValue.ToString() == "1")
            {
                //実施区分を打診中に変更する
                item1_MadoguchiJiishiKubun.SelectedValue = "2";
            }
        }

        private void checkBox32_CheckedChanged(object sender, EventArgs e)
        {
            ////調査区分のチェックが支部本社についた
            //if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
            //{
            //    //item1_MadoguchiChousaKubunJibusho.Checked = false;
            //    //item1_MadoguchiChousaKubunShibuShibu.Checked = false;
            //    //item1_MadoguchiChousaKubunHonbuShibu.Checked = false;

            //    //実施区分を打診中に変更する
            //    item1_MadoguchiJiishiKubun.SelectedValue = "2";
            //}
            if (item1_MadoguchiChousaKubunShibuHonbu.Checked && item1_MadoguchiJiishiKubun.SelectedValue.ToString() == "1")
            {
                //実施区分を打診中に変更する
                item1_MadoguchiJiishiKubun.SelectedValue = "2";
            }
            // 応援受付のアイコン表示・非表示切替
            OuenIconDisplay();
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
        // 集計表・報告書・調査資料図面フォルダ確認
        private void folderText_Leave(object sender, EventArgs e)
        {
            FolderPathCheck();
        }
        // 応援状況のアイコン表示・非表示
        private void OuenIconDisplay()
        {
            // 支→本、受付状況
            if (item1_MadoguchiChousaKubunShibuHonbu.Checked && item5_UketsukeJoukyo.Checked)
            {
                UketsukeIcon.Visible = true;

                // チェック時は完了アイコン
                UketsukeIcon.Image = Image.FromFile("Resource/kan.png");
            }
            else
            {
                if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
                {
                    UketsukeIcon.Visible = true;
                    UketsukeIcon.Image = Image.FromFile("Resource/OnegaiIcon35px.png");
                }
                else
                {
                    UketsukeIcon.Visible = false;
                }

            }
            // 支→本、応援完了
            if (item1_MadoguchiChousaKubunShibuHonbu.Checked && item5_OuenKanryo.Checked)
            {
                KanryouIcon.Visible = true;
            }
            else
            {
                KanryouIcon.Visible = false;
            }

        }
        private void item6_TanpinSeikyuuKakutei_CheckedChanged(object sender, EventArgs e)
        {
            SwichButton_Rank(-1);
        }

        //担当部所タブのGaroon追加宛先Gridをクリック時
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
                if (item1_MadoguchiTourokuNendo.Text != "")
                {
                    form.nendo = item1_MadoguchiTourokuNendo.SelectedValue.ToString();
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

        // 担当部所 行追加
        private void button_GaroonAtesakiGridAdd_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
            //描画停止
            c1FlexGrid5.BeginUpdate();

            c1FlexGrid5.Rows.Add();
            c1FlexGrid5.Rows[c1FlexGrid5.Rows.Count - 1].Height = 28;
            //不具合No1332(1084)　画面から追加されたよフラグをつける
            c1FlexGrid5.Rows[c1FlexGrid5.Rows.Count - 1].UserData = "1";
            Resize_Grid("c1FlexGrid5");

            //描画再開
            c1FlexGrid5.EndUpdate();
            //レイアウトロジックを再開する
            this.ResumeLayout();

        }
        // 応援受付の依頼書出力ボタン
        private void btn6_Iraisho(object sender, EventArgs e)
        {

            // エラークリア
            ErrorClear_KyouryokuIrai();

            // エラーチェック
            if (ErrorCheck_KyouryokuIrai())
            {
                set_error(GlobalMethod.GetMessage("E20403", "協力依頼書タブ"));
            }
            else
            {
                // 協力依頼書の更新
                UpdateMadoguchi(4);

                // 協力依頼書の出力
                output_KyouryokuIrai("OuenUketsuke");

                //// 支→本のみ表示
                //if (item1_MadoguchiChousaKubunShibuHonbu.Checked)
                //{
                //    // 応援状況にチェックがされているか
                //    if (item5_UketsukeJoukyo.Checked)
                //    {
                //        //UketsukeIcon.Image = Image.FromFile("Resource/OnegaiIcon35px.png");
                //        // 完了マークが出ているはずなので、表示しない
                //    }
                //    else
                //    {
                //        UketsukeIcon.Image = Image.FromFile("Resource/OnegaiIcon35px.png");
                //    }
                //    UketsukeIcon.Visible = true;
                //}
                //else
                //{
                //    UketsukeIcon.Visible = false;
                //}
            }
        }


        // 調査品目明細 調査品目一覧からの取込ボタン
        private void button3_ReadExcelChousaHinmoku_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            // エラーフラグ false：正常 true：エラー
            Boolean errorFlg = false;
            button3_ReadExcelResult.BackColor = Color.DarkGray;

            string table = "ChousaHinmoku";
            string UserID = "";
            string chousainMei = "";

            // Lockテーブル更新
            var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                SqlTransaction transaction = conn.BeginTransaction();
                cmd.Transaction = transaction;

                try
                {
                    // Lock情報取得
                    // 0:個人コード、1:氏名、2:部所CD、3:部所名、4：Role
                    chousainMei = UserInfos[1];

                    cmd.CommandText = "SELECT TOP 1 LOCK_USER_ID,LOCK_USER_MEI FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
                                      "AND LOCK_KEY = '" + MadoguchiID + "' ";
                    DataTable dt = new DataTable();

                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        // Lockテーブルにデータが存在した場合
                        UserID = dt.Rows[0][0].ToString();
                        chousainMei = dt.Rows[0][1].ToString();
                    }
                    else
                    {
                        // Lockテーブルにデータ存在しない場合
                        cmd.CommandText = "INSERT INTO T_LOCK(" +
                                         " LOCK_TABLE" +
                                         ",LOCK_KEY" +
                                         ",LOCK_USER_ID" +
                                         ",LOCK_USER_MEI" +
                                         ",LOCK_DATETIME" +
                                         ")VALUES(" +
                                         "'" + table + "' " +
                                         ",'" + MadoguchiID + "' " +
                                         ",'" + UserInfos[0] + "' " +
                                         ",N'" + UserInfos[1] + "' " +
                                         ",SYSDATETIME() " +
                                         ")";
                        cmd.ExecuteNonQuery();
                        transaction.Commit();
                        UserID = UserInfos[0];
                        chousainMei = UserInfos[1];
                    }
                }
                catch
                {
                    transaction.Rollback();
                    errorFlg = true;
                }
                finally
                {
                    conn.Close();
                }
            }

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
                    // ロック所有か自分がどうか
                    if (UserID == UserInfos[0])
                    {
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
                                if (result[1] != null && int.TryParse(result[1].ToString(), out count)) {
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
                                hinmokuRenkeiResult = GlobalMethod.MadoguchiHinmokuRenkeiUpdate_SQL(MadoguchiID, "Madoguchi", UserInfos[0], out resultMessage);

                                // メッセージがあれば画面に表示
                                if (resultMessage != "") {
                                    set_error(resultMessage);
                                }

                                // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック、Garoon連携対象チェック
                                //Garoon連携対象である場合、かつ、下記SQLの処理が正常終了した場合、Garoon連携処理を行う
                                if (item1_GaroonRenkei.Checked == true && hinmokuRenkeiResult == true)
                                {
                                    // VIPS　20220302　課題管理表No1275(969)　ADD　「Garoon連携処理」追加　対応
                                    GaroonBtn_Click(sender, e);
                                }

                                // 編集ロック開放
                                // Lockテーブル更新
                                using (var conn = new SqlConnection(connStr))
                                {
                                    conn.Open();
                                    var cmd = conn.CreateCommand();
                                    SqlTransaction transaction = conn.BeginTransaction();
                                    cmd.Transaction = transaction;

                                    try
                                    {
                                        // Lock情報取得
                                        // 0:個人コード、1:氏名、2:部所CD、3:部所名、4：Role
                                        chousainMei = UserInfos[1];

                                        cmd.CommandText = "DELETE FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
                                                          "AND LOCK_KEY = '" + MadoguchiID + "' " +
                                                          "AND LOCK_USER_ID = '" + UserInfos[0] + "' ";

                                        cmd.ExecuteNonQuery();
                                        transaction.Commit();

                                        item7_HenshuJoutai.Text = "編集ロック無";
                                        item7_HenshuLockTantousha.Text = "";
                                    }
                                    catch
                                    {
                                        transaction.Rollback();
                                        errorFlg = true;
                                    }
                                    finally
                                    {
                                        conn.Close();
                                    }

                                    button3_ReadExcelResult.BackColor = Color.DarkGray;
                                }
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
                    }
                    // 自身がロックしていない場合
                    else
                    {
                        set_error("現在編集にロックがかかっています");
                    }
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

        // 調査品目明細 検索ボタン
        private void button3_Search_Click(object sender, EventArgs e)
        {
            // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
            chousaHinmokuDispFlg = "0";
            // 調査品目明細のGridに全件読み込んだかどうかのフラグ 0:未 1:済
            chousaHinmokuLoadFlg = "0";

            Paging_now.Text = "1";
            item3_TargetPage.Text = "1";
            c1FlexGrid4.Rows.Count = 2;

            get_data(3);
        }

        //調査品目Gridの行表示切替
        private void Grid_Visible(int page,string setDataFlg = "0")
        {
            //描画停止
            c1FlexGrid4.BeginUpdate();

            // 全件Gridに表示で、set_data()以外から来た場合に、データを取り直す
            if (chousaHinmokuDispFlg == "1" && setDataFlg != "1" && chousaHinmokuLoadFlg == "0")
            {
                // 調査品目明細のGridに全件読み込んだかどうかのフラグ 0:未 1:済
                chousaHinmokuLoadFlg = "1";
                c1FlexGrid4.Rows.Count = 2;
                set_data(3);
            }

            // VIPS　20220203　課題管理表No797　CHANGE　表示件数「全件表示」対応
            //表示行フラグ true:表示 false:非表示
            int pagelimit = 0;
            // 「全件表示」の場合
            if (int.TryParse(item_Hyoujikensuu.Text, out pagelimit) == false)
            {
                // かなり大きな値をセット
                pagelimit = 999999999;
            }

            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            {
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
            // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
            chousaHinmokuDispFlg = "1";
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
            // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
            chousaHinmokuDispFlg = "1";
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
            // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
            chousaHinmokuDispFlg = "1";
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
            // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
            chousaHinmokuDispFlg = "1";
            //レイアウトロジックを停止する
            this.SuspendLayout();
            Paging_now.Text = (int.Parse(Paging_all.Text)).ToString();
            item3_TargetPage.Text = Paging_now.Text;
            Grid_Visible(int.Parse(Paging_now.Text));
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        //調査品目Gridソート時
        private void c1FlexGrid4_AfterSort(object sender, C1.Win.C1FlexGrid.SortColEventArgs e)
        {
            // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
            chousaHinmokuDispFlg = "1";
            Grid_Visible(int.Parse(Paging_now.Text));
        }

        //協力依頼書タブ　メールボタン
        private void item4_Mail_icon_Click(object sender, EventArgs e)
        {
            //メール本文を取得
            string MailText = GlobalMethod.GetCommonValue2("Ouenuketsuke");

            //件名を取得 //仕掛中
            string MailTitle = "";

            //メールアプリを起動
            Process.Start("mailto:''?subject=" + MailTitle + "&body=" + MailText);

        }

        //調査品目タブ　検索解除*
        private void button3_Clear_Click(object sender, EventArgs e)
        {
            //src_Busho.SelectedIndex = 0;
            //src_HinmokuChousain.Text = "";
            //src_ShuFuku.SelectedIndex = 0;
            //src_ChousaHinmei.Text = "";
            //src_ChousaKikaku.Text = "";
            //src_Zaikou.SelectedIndex = 0;
            //src_TantoushaKuuhaku.SelectedIndex = 0;
            //item_Hyoujikensuu.SelectedIndex = 1;
            chousaHinmokuClear();
        }

        // 調査品目明細タブ 検索解除
        private void chousaHinmokuClear()
        {
            src_Busho.SelectedIndex = 0;
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

        //調査品目タブ　入力開始ボタン押下時
        private void ChousaHinmokuGrid_InputMode()
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();
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
                // 全件削除
                button3_DeleteAllRow.BackColor = Color.FromArgb(42, 78, 122);
                button3_DeleteAllRow.Enabled = true;
                // 品目からの取込
                button3_ReadHinmoku.BackColor = Color.FromArgb(42, 78, 122);
                button3_ReadHinmoku.Enabled = true;

                button3_Lock.BackColor = Color.DimGray;
                button3_Lock.Enabled = false;
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
                button3_ChangeMoji.BackColor = Color.DimGray;
                button3_ChangeMoji.Enabled = false;
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
                // 全件削除
                button3_DeleteAllRow.BackColor = Color.DimGray;
                button3_DeleteAllRow.Enabled = false;
                // 品目からの取込
                button3_ReadHinmoku.BackColor = Color.DimGray;
                button3_ReadHinmoku.Enabled = false;

                button3_Lock.BackColor = Color.FromArgb(42, 78, 122);
                button3_Lock.Enabled = true;
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
                button3_ChangeMoji.BackColor = Color.FromArgb(42, 78, 122);
                button3_ChangeMoji.Enabled = true;

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
                        //No.1111
                        && i != c1FlexGrid4.Cols["SagyoForuda"].Index
                        && i != c1FlexGrid4.Cols["SagyoForudaPath"].Index
                        )
                    {
                        c1FlexGrid4.Cols[i].AllowEditing = EditMode;
                    }
                }
            }

            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        //「入力開始」「入力完了（更新）」押下時
        private void button3_InputStatus_Click(object sender, EventArgs e)
        {
            // メッセージクリア
            set_error("", 0);
            if (ChousaHinmokuMode == 0)
            {
                //「調査品目を入力出来るようにしますがよろしいですか。」
                if (GlobalMethod.outputMessage("I20308", "", 1) == DialogResult.OK)
                {
                    // 調査品目の削除Keys
                    deleteChousaHinmokuIDs = "";

                    // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
                    chousaHinmokuDispFlg = "1";
                    // まだ全件GridにLoadしていなければ、ロードする
                    if(chousaHinmokuLoadFlg != "1")
                    {
                        set_data(3);
                    }

                    set_error("", 0);
                    // エラーフラグ false：正常 true：エラー
                    Boolean errorFlg = false;

                    string table = "ChousaHinmoku";
                    string UserID = "";
                    string chousainMei = "";

                    var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                    using (var conn = new SqlConnection(connStr))
                    {
                        conn.Open();
                        var cmd = conn.CreateCommand();
                        SqlTransaction transaction = conn.BeginTransaction();
                        cmd.Transaction = transaction;

                        try
                        {
                            // Lock情報取得
                            // 0:個人コード、1:氏名、2:部所CD、3:部所名、4：Role
                            chousainMei = UserInfos[1];

                            cmd.CommandText = "SELECT TOP 1 LOCK_USER_ID,LOCK_USER_MEI FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
                                              "AND LOCK_KEY = '" + MadoguchiID + "' ";
                            DataTable dt = new DataTable();

                            var sda = new SqlDataAdapter(cmd);
                            sda.Fill(dt);

                            if (dt.Rows.Count > 0)
                            {
                                // Lockテーブルにデータが存在した場合
                                UserID = dt.Rows[0][0].ToString();
                                chousainMei = dt.Rows[0][1].ToString();

                                // ロックしているユーザーが自分以外の場合
                                if(UserID != UserInfos[0]) 
                                { 
                                    // ロック中
                                    MessageBox.Show("現在編集にロックがかかっています");
                                    return;
                                }
                            }
                            else
                            {
                                // Lockテーブルにデータ存在しない場合
                                cmd.CommandText = "INSERT INTO T_LOCK(" +
                                                 " LOCK_TABLE" +
                                                 ",LOCK_KEY" +
                                                 ",LOCK_USER_ID" +
                                                 ",LOCK_USER_MEI" +
                                                 ",LOCK_DATETIME" +
                                                 ")VALUES(" +
                                                 "'" + table + "' " +
                                                 ",'" + MadoguchiID + "' " +
                                                 ",'" + UserInfos[0] + "' " +
                                                 ",N'" + UserInfos[1] + "' " +
                                                 ",SYSDATETIME() " +
                                                 ")";
                                cmd.ExecuteNonQuery();
                                //transaction.Commit();
                                UserID = UserInfos[0];
                                chousainMei = UserInfos[1];

                                item7_HenshuJoutai.Text = "編集ロック中";
                                item7_HenshuLockTantousha.Text = UserInfos[1];

                            }
                            transaction.Commit();



                        }
                        catch
                        {
                            transaction.Rollback();
                            errorFlg = true;
                        }
                        finally
                        {
                            conn.Close();
                        }
                    }

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

                    // 調査品目の更新
                    chousaHinmokuUpdate();

                    // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック、Garoon連携対象チェック
                    //Garoon連携対象である場合、かつ、更新処理でエラーが出ていない場合に連携処理を行う。
                    if (item1_GaroonRenkei.Checked == true && globalErrorFlg == "0")
                    {
                        //1578 1597
                        int errorC = 0;
                        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                        {
                            if (c1FlexGrid4.Rows[i]["GroupMei"] == null || c1FlexGrid4.Rows[i]["GroupMei"].ToString() == "" && c1FlexGrid4.Rows[i]["ShukeihyoVer"].ToString() == "2" && c1FlexGrid4.Rows[i]["BunkatsuHouhou"].ToString() == "2")
                            {
                                if (errorC == 0)
                                {
                                    set_error(GlobalMethod.GetMessage("W20304", ""));
                                    errorC = 1;
                                }
                                // ピンク背景
                                c1FlexGrid4.GetCellRange(i, 59).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                                c1FlexGrid4.Rows[i]["ColumnSort"] = "N"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaZentaiJun"] != null ? c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString() : "0"))
                                                                  + "-"
                                                                  + zeroPadding((c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] != null ? c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString() : "0"))
                                                                  ;
                            }
                            else
                            {
                                // 必須背景薄黄色
                                c1FlexGrid4.GetCellRange(i, 59).StyleNew.BackColor = Color.White;
                            }
                            //No.1622
                            if(c1FlexGrid4.Rows[i]["ShukeihyoVer"].ToString() == "2" && c1FlexGrid4.Rows[i]["BunkatsuHouhou"].ToString() == "1")
                            {
                                c1FlexGrid4.GetCellRange(i, 59).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                            }
                        }
                        GaroonBtn_Click(sender, e);
                    }

                    //writeHistory("【開始】調査品目明細の更新を終了します。 ID= " + MadoguchiID);

                    //writeHistory("【開始】調査品目明細の更新を開始します。 ID= " + MadoguchiID);

                    //// Gridにデータが存在する場合
                    //// １．調査品目の削除Key（ChousaHinmokuIDをカンマ区切りで連結したデータ）があれば削除
                    //// ２．c1FlexGrid4 の 57:0:Insert/1:Select/2:Update があり、それで新規か更新、または処理なしを切り分ける
                    //// ３．ChousaHinmokuから担当部所の連携を行う（支部備考も）

                    //// 品目のCommit件数を取得
                    //int i_RecodeCountMax = 0;
                    //string w_RecodeCountMax = GlobalMethod.GetCommonValue1("HINMOKU_COMMIT_KENSU");
                    //if (w_RecodeCountMax != null)
                    //{
                    //    int.TryParse(w_RecodeCountMax, out i_RecodeCountMax);
                    //    if (i_RecodeCountMax == 0)
                    //    {
                    //        i_RecodeCountMax = 100;
                    //    }
                    //}
                    //else
                    //{
                    //    i_RecodeCountMax = 100;
                    //}

                    //// 特調番号（ヘッダーにあるので利用する）
                    //string tokuchoBangou = Header1.Text;

                    //// メッセージフラグ1
                    //int updmessage1 = 0; // 新規
                    //int updmessage2 = 0; // 更新
                    //int updmessage3 = 0; // 削除
                    //// エラーメッセージフラグ
                    //int errmessage1 = 0; // E20307:全体順、個別順が重複しています。
                    //int errmessage2 = 0; // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい。
                    //int errmessage3 = 0; // E20337:半角数字で入力してください。赤背景の項目を修正して下さい。
                    //int cnt = 0;
                    //Boolean errorFlg = false;
                    //string sysDateTimeStr = "";

                    //string insertQuery = "Insert Into ChousaHinmoku( " +
                    //"ChousaHinmokuID " +
                    //",MadoguchiID " +
                    //",ChousaZentaiJun " +
                    //",ChousaKobetsuJun " +
                    //",ChousaZaiKou " +
                    //",ChousaHinmei " +
                    //",ChousaKikaku " +
                    //",ChousaTanka " +
                    //",ChousaSankouShitsuryou " +
                    //",ChousaKakaku " +
                    //",ChousaChuushi " +
                    //",ChousaBikou2 " +
                    //",ChousaBikou " +
                    //",ChousaTankaTekiyouTiku " +
                    //",ChousaZumenNo " +
                    //",ChousaSuuryou " +
                    //",ChousaMitsumorisaki " +
                    //",ChousaBaseMakere " +
                    //",ChousaBaseTanka " +
                    //",ChousaKakeritsu " +
                    //",ChousaObiMei " +
                    //",ChousaZenkaiTani " +
                    //",ChousaZenkaiKakaku " +
                    //",ChousaSankouti " +
                    //",ChousaHinmokuJouhou1 " +
                    //",ChousaHinmokuJouhou2 " +
                    //",ChousaFukuShizai " +
                    //",ChousaBunrui " +
                    //",ChousaMemo2 " +
                    //",ChousaTankaCD1 " +
                    //",ChousaTikuWariCode " +
                    //",ChousaTikuCode " +
                    //",ChousaTikuMei " +
                    //",ChousaShougaku " +
                    //",ChousaWebKen " +
                    //",ChousaKonkyoCode " +
                    //",ChousaLinkSakli " +
                    //",HinmokuRyakuBushoCD " +
                    //",HinmokuChousainCD " +
                    //",HinmokuRyakuBushoFuku1CD " +
                    //",HinmokuFukuChousainCD1 " +
                    //",HinmokuRyakuBushoFuku2CD " +
                    //",HinmokuFukuChousainCD2 " +
                    //",ChousaHoukokuHonsuu " +
                    //",ChousaHoukokuRank " +
                    //",ChousaIraiHonsuu " +
                    //",ChousaIraiRank " +
                    //",ChousaHinmokuShimekiribi " +
                    //",ChousaHoukokuzumi " +
                    //",ChousaDeleteFlag " +
                    //",ChousaCreateDate " +
                    //",ChousaCreateUser " +
                    //",ChousaCreateProgram " +
                    //",ChousaUpdateDate " +
                    //",ChousaUpdateUser " +
                    //",ChousaUpdateProgram " +
                    //",ChousaShinchokuJoukyou " +
                    //") VALUES ";
                    //string valuesText = "";


                    //var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                    //using (var conn = new SqlConnection(connStr))
                    //{
                    //    conn.Open();
                    //    var cmd = conn.CreateCommand();
                    //    SqlTransaction transaction = conn.BeginTransaction();
                    //    cmd.Transaction = transaction;

                    //    // エラーチェック

                    //    // 6:全体順、7:個別順が重複している場合
                    //    // E20307:全体順、個別順が重複しています。
                    //    // Gridの中で重複していないか確認、DBにある値とも重複を確認する

                    //    string zentai = "";
                    //    string kobetsu = "";
                    //    float zentaiF = 0;
                    //    float kobetsuF = 0;
                    //    float zentaiNextF = 0;
                    //    float kobetsuNextF = 0;
                    //    int recordCount = 0;
                    //    // Grid内で重複していないか確認
                    //    for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                    //    {
                    //        zentai = c1FlexGrid4.Rows[i][6].ToString();
                    //        kobetsu = c1FlexGrid4.Rows[i][7].ToString();

                    //        float.TryParse(zentai, out zentaiF);
                    //        float.TryParse(kobetsu, out kobetsuF);

                    //        recordCount = 0;

                    //        // 色付けしても次のループでまた塗り直ししてしまうので、全件で回す
                    //        //for(int j = i + 1;j < c1FlexGrid4.Rows.Count; j++)
                    //        for (int j = 2;j < c1FlexGrid4.Rows.Count; j++)
                    //        {

                    //            if (c1FlexGrid4.Rows[j][6] != null && c1FlexGrid4.Rows[j][7] != null && c1FlexGrid4.Rows[j][6].ToString() != "" && c1FlexGrid4.Rows[j][7].ToString() != "")
                    //            {
                    //                zentai = c1FlexGrid4.Rows[j][6].ToString();
                    //                kobetsu = c1FlexGrid4.Rows[j][7].ToString();

                    //                float.TryParse(zentai, out zentaiNextF);
                    //                float.TryParse(kobetsu, out kobetsuNextF);

                    //                if (zentaiF == zentaiNextF && kobetsuF == kobetsuNextF)
                    //                {
                    //                    recordCount += 1;
                    //                }
                    //            }
                    //        }
                    //        // 重複レコードがあった場合
                    //        if(recordCount > 1)
                    //        {
                    //            errmessage1 = 1;
                    //            errorFlg = true;
                    //            // ピンク背景
                    //            c1FlexGrid4.GetCellRange(i, 6).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                    //            c1FlexGrid4.GetCellRange(i, 7).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                    //            // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                    //            c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                    //        }
                    //        else
                    //        {
                    //            // クリーム色背景
                    //            c1FlexGrid4.GetCellRange(i, 6).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                    //            c1FlexGrid4.GetCellRange(i, 7).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                    //        }

                    //    }

                    //    // 検索したときの条件 chousaHinmokuSearchWhere 
                    //    // 検索で絞った場合、DBにいるデータと重複しているとNG
                    //    if (chousaHinmokuSearchWhere != "") { 
                    //        // 検索で出てきていないデータを取得する
                    //        string where = "MadoguchiID = '" + MadoguchiID + "' AND not (" + chousaHinmokuSearchWhere + ")";
                    //        //コンボボックスデータ取得
                    //        DataTable combodt = GlobalMethod.getData("ChousaKobetsuJun", "ChousaZentaiJun", "ChousaHinmoku", where);

                    //        if(combodt != null && combodt.Rows.Count > 0) { 
                    //            for (int i = 0; i < combodt.Rows.Count; i++)
                    //            {
                    //                zentai = combodt.Rows[i][0].ToString();
                    //                kobetsu = combodt.Rows[i][1].ToString();

                    //                float.TryParse(zentai, out zentaiF);
                    //                float.TryParse(kobetsu, out kobetsuF);

                    //                for (int j = 2; j < c1FlexGrid4.Rows.Count; j++)
                    //                {
                    //                    if (c1FlexGrid4.Rows[j][6] != null && c1FlexGrid4.Rows[j][7] != null && c1FlexGrid4.Rows[j][6].ToString() != "" && c1FlexGrid4.Rows[j][7].ToString() != "")
                    //                    {
                    //                        zentai = c1FlexGrid4.Rows[j][6].ToString();
                    //                        kobetsu = c1FlexGrid4.Rows[j][7].ToString();

                    //                        float.TryParse(zentai, out zentaiNextF);
                    //                        float.TryParse(kobetsu, out kobetsuNextF);

                    //                        if (zentaiF == zentaiNextF && kobetsuF == kobetsuNextF)
                    //                        {
                    //                            errmessage1 = 1;
                    //                            errorFlg = true;
                    //                            // ピンク背景
                    //                            c1FlexGrid4.GetCellRange(j, 6).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                    //                            c1FlexGrid4.GetCellRange(j, 7).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                    //                            // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                    //                            c1FlexGrid4[j, 58] = "E" + zeroPadding(c1FlexGrid4[j, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[j, 7].ToString());
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //        }
                    //    }

                    //    // 34:地区割コードが空でなく、半角英数字でない場合、正規表現「^[0-9a-zA-Z]+$」
                    //    // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい。

                    //    // 35:地区コードが空でなく、半角英数字でない場合、正規表現「^[0-9a-zA-Z]+$」
                    //    // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい

                    //    // 13:価格が空でなく、「,」を空文字に置換し、Trimした値が、半角数字でない場合、正規表現「^-?[\d][\d.]*$」
                    //    // E20337:半角数字で入力してください。赤背景の項目を修正して下さい。
                    //    for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                    //    {
                    //        // 地区割りコード
                    //        if (c1FlexGrid4.Rows[i][34] != null && c1FlexGrid4.Rows[i][34].ToString() != ""
                    //            && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i][34].ToString().Replace(Environment.NewLine, ""), @"^[0-9a-zA-Z]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                    //        {
                    //            errmessage2 = 1;
                    //            errorFlg = true;
                    //            // ピンク背景
                    //            c1FlexGrid4.GetCellRange(i, 34).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                    //            // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                    //            c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                    //        }
                    //        else
                    //        {
                    //            // 白背景
                    //            c1FlexGrid4.GetCellRange(i, 34).StyleNew.BackColor = Color.White;
                    //        }
                    //        // 地区コード
                    //        if (c1FlexGrid4.Rows[i][35] != null && c1FlexGrid4.Rows[i][35].ToString() != ""
                    //            && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i][35].ToString().Replace(Environment.NewLine, ""), @"^[0-9a-zA-Z]+$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                    //        {
                    //            errmessage2 = 1;
                    //            errorFlg = true;
                    //            // ピンク背景
                    //            c1FlexGrid4.GetCellRange(i, 35).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                    //            // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                    //            c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                    //        }
                    //        else
                    //        {
                    //            // 白背景
                    //            c1FlexGrid4.GetCellRange(i, 35).StyleNew.BackColor = Color.White;
                    //        }
                    //        // 価格
                    //        // C1FlexGridの制御で以下のチェックに入らない
                    //        if (c1FlexGrid4.Rows[i][13] != null && c1FlexGrid4.Rows[i][13].ToString() != ""
                    //            && !System.Text.RegularExpressions.Regex.IsMatch(c1FlexGrid4.Rows[i][13].ToString(), @"^-?[\d][\d.]*$", System.Text.RegularExpressions.RegexOptions.ECMAScript))
                    //        {
                    //            errmessage3 = 1;
                    //            errorFlg = true;
                    //            // ピンク背景
                    //            c1FlexGrid4.GetCellRange(i, 13).StyleNew.BackColor = Color.FromArgb(255, 200, 255);
                    //            // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                    //            c1FlexGrid4[i, 58] = "E" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());
                    //        }
                    //        else
                    //        {
                    //            // 白背景
                    //            c1FlexGrid4.GetCellRange(i, 13).StyleNew.BackColor = Color.White;
                    //        }
                    //    }

                    //    GlobalMethod.outputLogger("Madoguhi button3_InputStatus_Click", "DB比較終了 ID:" + MadoguchiID, "update", "DEBUG");

                    //    // エラーが無ければ
                    //    if (errorFlg == false)
                    //    {
                    //        try
                    //        {
                    //            // １．調査品目の削除Key（ChousaHinmokuIDをカンマ区切りで連結したデータ）があれば削除
                    //            if (deleteChousaHinmokuIDs != "")
                    //            {
                    //                // 削除
                    //                // 調査品目全削除
                    //                cmd.CommandText = "DELETE FROM ChousaHinmoku " +
                    //                    "WHERE ChousaHinmokuID in (" + deleteChousaHinmokuIDs + ") AND MadoguchiID = '" + MadoguchiID + "' ";
                    //                cmd.ExecuteNonQuery();

                    //                // T_History に登録する文言は256文字までなので、分割していれないと桁あふれする
                    //                //writeHistory("調査品目が削除されました。調査品目ID in (" + deleteChousaHinmokuIDs + ")");

                    //                string[] deleteID = deleteChousaHinmokuIDs.Split(',');

                    //                updmessage3 = 1; // 削除

                    //                for (int i = 0; i < deleteID.Length; i++)
                    //                {
                    //                    writeHistory("調査品目が削除されました。調査品目ID = " + deleteID[i]);
                    //                }

                    //            }

                    //            // ２．c1FlexGrid4 の 57:0:Insert/1:Select/2:Update があり、それで新規か更新、または処理なしを切り分ける
                    //            // ２－１．まずは新規を処理する

                    //            // ソートが効いてない、、、
                    //            //c1FlexGrid4.Cols[57].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending; // 57:0:Insert/1:Select/2:Update の昇順設定
                    //            //c1FlexGrid4.Cols[55].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending; // 55:ChousaHinmokuID の昇順設定
                    //            //c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 57, 55);  // 設定した内容で、ソートする

                    //            // 更新日付をあらかじめ所得しておく
                    //            sysDateTimeStr = DateTime.Now.ToString();

                    //            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                    //            {
                    //                // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                    //                c1FlexGrid4[i, 58] = "N" + zeroPadding(c1FlexGrid4[i, 6].ToString()) + "-" + zeroPadding(c1FlexGrid4[i, 7].ToString());

                    //                // 0:Insertを処理する
                    //                if (c1FlexGrid4.Rows[i][57] != null && c1FlexGrid4.Rows[i][57].ToString() == "0")
                    //                {
                    //                    //if (cnt >= 100)
                    //                    if (cnt >= i_RecodeCountMax)
                    //                    {
                    //                        // 登録をまとめて行う
                    //                        cmd.CommandText = insertQuery + valuesText;

                    //                        cmd.ExecuteNonQuery();
                    //                        // 追加メッセージ
                    //                        updmessage1 = 1;
                    //                        cnt = 0;
                    //                        valuesText = "";

                    //                    }
                    //                    cnt += 1;

                    //                    if (valuesText != "")
                    //                    {
                    //                        valuesText += ",";
                    //                    }

                    //                    valuesText += "(" +
                    //                        " '" + c1FlexGrid4.Rows[i][55] + "' " +
                    //                        ",'" + MadoguchiID + "' " +
                    //                        ",'" + c1FlexGrid4.Rows[i][6] + "' " +   // 全体順
                    //                        ",'" + c1FlexGrid4.Rows[i][7] + "' " +   // 個別順
                    //                        ",'" + c1FlexGrid4.Rows[i][8] + "' " +   // 材工
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][9].ToString(), 0, 0) + "' " +   // 品目
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][10].ToString(), 0, 0) + "' " +  // 規格
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][11].ToString(), 0, 0) + "' " +  // 単位
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][12].ToString(), 0, 0) + "' ";   // 参考質量

                    //                    // 価格
                    //                    if (c1FlexGrid4.Rows[i][13] != null && c1FlexGrid4.Rows[i][13].ToString() != "")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][13] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",null ";
                    //                    }

                    //                    // 中止
                    //                    if (c1FlexGrid4.Rows[i][14] != null && c1FlexGrid4.Rows[i][14].ToString() == "True")
                    //                    {
                    //                        valuesText += ",1 ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",0 ";
                    //                    }

                    //                    valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][15].ToString(), 0, 0) + "' " +  // 報告備考
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][16].ToString(), 0, 0) + "' " +            // 依頼備考
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][17].ToString(), 0, 0) + "' " +            // 単価適用地域
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][18].ToString(), 0, 0) + "' " +            // 図面番号
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][19].ToString(), 0, 0) + "' " +            // 数量
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][20].ToString(), 0, 0) + "' " +            // 見積先
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][21].ToString(), 0, 0) + "' ";             // ベースメーカー

                    //                    // ベース単価
                    //                    if (c1FlexGrid4.Rows[i][22] != null && c1FlexGrid4.Rows[i][22].ToString() != "" && c1FlexGrid4.Rows[i][22].ToString() != "0")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][22] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",'           0.00' ";
                    //                    }

                    //                    // 掛率
                    //                    if (c1FlexGrid4.Rows[i][23] != null && c1FlexGrid4.Rows[i][23].ToString() != "" && c1FlexGrid4.Rows[i][23].ToString() != "0")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][23] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",'  0.00' ";
                    //                    }

                    //                    valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][24].ToString(), 0, 0) + "' " +  // 属性
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][25].ToString(), 0, 0) + "' ";             // 前回単位

                    //                    // 前回価格
                    //                    if (c1FlexGrid4.Rows[i][26] != null && c1FlexGrid4.Rows[i][26].ToString() != "")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][26] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",0 ";
                    //                    }

                    //                    // 発注者提供単価
                    //                    if (c1FlexGrid4.Rows[i][27] != null && c1FlexGrid4.Rows[i][27].ToString() != "")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][27] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",0 ";
                    //                    }

                    //                    valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][28].ToString(), 0, 0) + "' " +  // 品目情報1
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][29].ToString(), 0, 0) + "' " +  // 品目情報2
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][30].ToString(), 0, 0) + "' " +  // 前回質量
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][31].ToString(), 0, 0) + "' " +  // メモ1
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][32].ToString(), 0, 0) + "' " +  // メモ2
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][33].ToString(), 0, 0) + "' " +  // 発注品目コード
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][34].ToString(), 0, 0) + "' " +  // 地区割コード
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][35].ToString(), 0, 0) + "' " +  // 地区コード
                    //                        ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][36].ToString(), 0, 0) + "' ";   // 地区名

                    //                    // 少額案件[10万/100万]
                    //                    if (c1FlexGrid4.Rows[i][37] != null && c1FlexGrid4.Rows[i][37].ToString() == "True")
                    //                    {
                    //                        valuesText += ",1 ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",0 ";
                    //                    }
                    //                    // Web建
                    //                    if (c1FlexGrid4.Rows[i][38] != null && c1FlexGrid4.Rows[i][38].ToString() == "True")
                    //                    {
                    //                        valuesText += ",1 ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",0 ";
                    //                    }

                    //                    valuesText += ",'" + c1FlexGrid4.Rows[i][39] + "' " +  // 根拠関連コード
                    //                                                                           // リンク先アイコン
                    //                        ",'" + c1FlexGrid4.Rows[i][41] + "' ";   // リンク先パス


                    //                    // 調査担当部所
                    //                    if (c1FlexGrid4.Rows[i][42] != null && c1FlexGrid4.Rows[i][42].ToString() != "")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][42] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",null ";
                    //                    }
                    //                    // 調査担当者
                    //                    if (c1FlexGrid4.Rows[i][43] != null && c1FlexGrid4.Rows[i][43].ToString() != "" && c1FlexGrid4.Rows[i][43].ToString() != "0")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][43] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",null ";
                    //                    }
                    //                    // 副調査担当部所1
                    //                    if (c1FlexGrid4.Rows[i][44] != null && c1FlexGrid4.Rows[i][44].ToString() != "")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][44] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",null ";
                    //                    }
                    //                    // 副調査担当者1
                    //                    if (c1FlexGrid4.Rows[i][45] != null && c1FlexGrid4.Rows[i][45].ToString() != "" && c1FlexGrid4.Rows[i][45].ToString() != "0")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][45] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",null ";
                    //                    }
                    //                    // 副調査担当部所2
                    //                    if (c1FlexGrid4.Rows[i][46] != null && c1FlexGrid4.Rows[i][46].ToString() != "")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][46] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",null ";
                    //                    }
                    //                    // 副調査担当者2
                    //                    if (c1FlexGrid4.Rows[i][47] != null && c1FlexGrid4.Rows[i][47].ToString() != "" && c1FlexGrid4.Rows[i][47].ToString() != "0")
                    //                    {
                    //                        valuesText += ",'" + c1FlexGrid4.Rows[i][47] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",null ";
                    //                    }

                    //                    //valuesText += ",'" + c1FlexGrid4.Rows[i][48] + "' " +  // 報告数
                    //                    //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][49].ToString(), 0, 0) + "' " +  // 報告ランク
                    //                    //    //",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][50].ToString(), 0, 0) + "' " +  // 依頼数
                    //                    //    ",'" + c1FlexGrid4.Rows[i][50] + "' " +  // 依頼数
                    //                    //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][51].ToString(), 0, 0) + "' " +  // 依頼ランク
                    //                    //    ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][52].ToString(), 0, 0) + "' ";  // 締切日

                    //                    valuesText += ",'" + c1FlexGrid4.Rows[i][48] + "' ";  // 報告数
                    //                    if (c1FlexGrid4.Rows[i][49] != null && c1FlexGrid4.Rows[i][49].ToString() != "")
                    //                    {
                    //                        valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][49].ToString(), 0, 0) + "' ";  // 報告ランク
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",'' ";  // 報告ランク
                    //                    }

                    //                    valuesText += ",'" + c1FlexGrid4.Rows[i][50] + "' ";  // 依頼数

                    //                    if (c1FlexGrid4.Rows[i][51] != null && c1FlexGrid4.Rows[i][51].ToString() != "")
                    //                    {
                    //                        valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][51].ToString(), 0, 0) + "' ";  // 依頼ランク
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",'' ";  // 依頼ランク
                    //                    }

                    //                    valuesText += ",'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][52].ToString(), 0, 0) + "' ";  // 締切日

                    //                    // 報告済
                    //                    if (c1FlexGrid4.Rows[i][53] != null && c1FlexGrid4.Rows[i][53].ToString() == "True")
                    //                    {
                    //                        valuesText += ",1 ";
                    //                    }
                    //                    else
                    //                    {
                    //                        valuesText += ",0 ";
                    //                    }

                    //                    valuesText += ",'0' " +                          // 削除フラグ
                    //                        ",'" + sysDateTimeStr + "'" +                // 登録日時
                    //                        ",'" + UserInfos[0] + "' " +                 // 登録ユーザー
                    //                        ",'Madoguchi button3_InputStatus_Click' " +  // 登録プログラム
                    //                        ",'" + sysDateTimeStr + "'" +                // 更新日時
                    //                        ",'" + UserInfos[0] + "' " +                 // 更新ユーザー
                    //                        ",'Madoguchi button3_InputStatus_Click' " +  // 更新プログラム
                    //                        ",'" + c1FlexGrid4.Rows[i][56] + "' ";       // 進捗状況
                    //                    valuesText += ")";

                    //                }
                    //            }
                    //            // valuesTextが空でなければinsertを行う
                    //            if (valuesText != "")
                    //            {
                    //                // 登録を行う
                    //                cmd.CommandText = insertQuery + valuesText;
                    //                cmd.ExecuteNonQuery();
                    //                // 追加メッセージ
                    //                updmessage1 = 1;
                    //            }


                    //            // ２－２．更新を処理する

                    //            // 更新日付をあらかじめ所得しておく
                    //            sysDateTimeStr = DateTime.Now.ToString();

                    //            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                    //            {
                    //                // 2:Updateを処理する
                    //                if (c1FlexGrid4.Rows[i][57] != null && c1FlexGrid4.Rows[i][57].ToString() == "2")
                    //                {
                    //                    cmd.CommandText = "UPDATE ChousaHinmoku SET ";

                    //                    cmd.CommandText += " ChousaZentaiJun = '" + c1FlexGrid4.Rows[i][6] + "' " +   // 全体順
                    //                        ",ChousaKobetsuJun = '" + c1FlexGrid4.Rows[i][7] + "' " +   // 個別順
                    //                        ",ChousaZaiKou = '" + c1FlexGrid4.Rows[i][8] + "' " +   // 材工
                    //                        ",ChousaHinmei = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][9].ToString(), 0, 0) + "' " +   // 品目
                    //                        ",ChousaKikaku = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][10].ToString(), 0, 0) + "' " +  // 規格
                    //                        ",ChousaTanka = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][11].ToString(), 0, 0) + "' " +  // 単位
                    //                        ",ChousaSankouShitsuryou = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][12].ToString(), 0, 0) + "' ";   // 参考質量

                    //                    // 価格
                    //                    if (c1FlexGrid4.Rows[i][13] != null && c1FlexGrid4.Rows[i][13].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",ChousaKakaku = '" + c1FlexGrid4.Rows[i][13] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",ChousaKakaku = null ";
                    //                    }

                    //                    // 中止
                    //                    if (c1FlexGrid4.Rows[i][14] != null && c1FlexGrid4.Rows[i][14].ToString() == "True")
                    //                    {
                    //                        cmd.CommandText += ",ChousaChuushi = 1 ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",ChousaChuushi = 0 ";
                    //                    }

                    //                    cmd.CommandText += ",ChousaBikou2 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][15].ToString(), 0, 0) + "' " +  // 報告備考
                    //                        ",ChousaBikou = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][16].ToString(), 0, 0) + "' " +  // 依頼備考
                    //                        ",ChousaTankaTekiyouTiku = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][17].ToString(), 0, 0) + "' " +  // 単価適用地域
                    //                        ",ChousaZumenNo = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][18].ToString(), 0, 0) + "' " +  // 図面番号
                    //                        ",ChousaSuuryou = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][19].ToString(), 0, 0) + "' " +  // 数量
                    //                        ",ChousaMitsumorisaki = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][20].ToString(), 0, 0) + "' " +  // 見積先
                    //                        ",ChousaBaseMakere = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][21].ToString(), 0, 0) + "' " +  // ベースメーカー
                    //                        ",ChousaBaseTanka = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][22].ToString(), 0, 0) + "' " +  // ベース単位
                    //                        ",ChousaKakeritsu = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][23].ToString(), 0, 0) + "' " +  // 掛率
                    //                        ",ChousaObiMei = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][24].ToString(), 0, 0) + "' " +  // 属性
                    //                        ",ChousaZenkaiTani = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][25].ToString(), 0, 0) + "' " +  // 前回単位
                    //                        ",ChousaZenkaiKakaku = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][26].ToString(), 0, 0) + "' " +  // 前回価格
                    //                        ",ChousaSankouti = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][27].ToString(), 0, 0) + "' " +  // 発注者提供単価
                    //                        ",ChousaHinmokuJouhou1 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][28].ToString(), 0, 0) + "' " +  // 品目情報1
                    //                        ",ChousaHinmokuJouhou2 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][29].ToString(), 0, 0) + "' " +  // 品目情報2
                    //                        ",ChousaFukuShizai = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][30].ToString(), 0, 0) + "' " +  // 前回質量
                    //                        ",ChousaBunrui = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][31].ToString(), 0, 0) + "' " +  // メモ1
                    //                        ",ChousaMemo2 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][32].ToString(), 0, 0) + "' " +  // メモ2
                    //                        ",ChousaTankaCD1 = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][33].ToString(), 0, 0) + "' " +  // 発注品目コード
                    //                        ",ChousaTikuWariCode = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][34].ToString(), 0, 0) + "' " +  // 地区割コード
                    //                        ",ChousaTikuCode = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][35].ToString(), 0, 0) + "' " +  // 地区コード
                    //                        ",ChousaTikuMei = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][36].ToString(), 0, 0) + "' ";   // 地区名
                    //                                                                                                                              // 少額案件[10万/100万]
                    //                    if (c1FlexGrid4.Rows[i][37] != null && c1FlexGrid4.Rows[i][37].ToString() == "True")
                    //                    {
                    //                        cmd.CommandText += ",ChousaShougaku = 1 ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",ChousaShougaku = 0 ";
                    //                    }
                    //                    // Web建
                    //                    if (c1FlexGrid4.Rows[i][38] != null && c1FlexGrid4.Rows[i][38].ToString() == "True")
                    //                    {
                    //                        cmd.CommandText += ",ChousaWebKen = 1 ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",ChousaWebKen = 0 ";
                    //                    }

                    //                    cmd.CommandText += ",ChousaKonkyoCode = '" + c1FlexGrid4.Rows[i][39] + "' " +  // 根拠関連コード
                    //                                                                                                   // リンク先アイコン
                    //                        ",ChousaLinkSakli = '" + c1FlexGrid4.Rows[i][41] + "' ";   // リンク先パス


                    //                    // 調査担当部所
                    //                    if (c1FlexGrid4.Rows[i][42] != null && c1FlexGrid4.Rows[i][42].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",HinmokuRyakuBushoCD = '" + c1FlexGrid4.Rows[i][42] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",HinmokuRyakuBushoCD = null ";
                    //                    }
                    //                    // 調査担当者
                    //                    if (c1FlexGrid4.Rows[i][43] != null && c1FlexGrid4.Rows[i][43].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",HinmokuChousainCD = '" + c1FlexGrid4.Rows[i][43] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",HinmokuChousainCD = null ";
                    //                    }
                    //                    // 副調査担当部所1
                    //                    if (c1FlexGrid4.Rows[i][44] != null && c1FlexGrid4.Rows[i][44].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",HinmokuRyakuBushoFuku1CD = '" + c1FlexGrid4.Rows[i][44] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",HinmokuRyakuBushoFuku1CD = null ";
                    //                    }
                    //                    // 副調査担当者1
                    //                    if (c1FlexGrid4.Rows[i][45] != null && c1FlexGrid4.Rows[i][45].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",HinmokuFukuChousainCD1 = '" + c1FlexGrid4.Rows[i][45] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",HinmokuFukuChousainCD1 = null ";
                    //                    }
                    //                    // 副調査担当部所2
                    //                    if (c1FlexGrid4.Rows[i][46] != null && c1FlexGrid4.Rows[i][46].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",HinmokuRyakuBushoFuku2CD = '" + c1FlexGrid4.Rows[i][46] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",HinmokuRyakuBushoFuku2CD = null ";
                    //                    }
                    //                    // 副調査担当者2
                    //                    if (c1FlexGrid4.Rows[i][47] != null && c1FlexGrid4.Rows[i][47].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",HinmokuFukuChousainCD2 = '" + c1FlexGrid4.Rows[i][47] + "' ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",HinmokuFukuChousainCD2 = null ";
                    //                    }

                    //                    //cmd.CommandText += ",ChousaHoukokuHonsuu = '" + c1FlexGrid4.Rows[i][48] + "' " +  // 報告数
                    //                    //    ",ChousaHoukokuRank = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][49].ToString(), 0, 0) + "' " +  // 報告ランク
                    //                    //    //",ChousaIraiHonsuu = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][50].ToString(), 0, 0) + "' " +  // 依頼数
                    //                    //    ",ChousaIraiHonsuu = '" + c1FlexGrid4.Rows[i][50]+ "' " +  // 依頼数
                    //                    //    ",ChousaIraiRank = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][51].ToString(), 0, 0) + "' " +  // 依頼ランク
                    //                    //    ",ChousaHinmokuShimekiribi = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][52].ToString(), 0, 0) + "' ";  // 締切日

                    //                    cmd.CommandText += ",ChousaHoukokuHonsuu = '" + c1FlexGrid4.Rows[i][48] + "' ";  // 報告数
                    //                    if (c1FlexGrid4.Rows[i][49] != null && c1FlexGrid4.Rows[i][49].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",ChousaHoukokuRank = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][49].ToString(), 0, 0) + "' ";  // 報告ランク
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",ChousaHoukokuRank = '' ";  // 報告ランク
                    //                    }

                    //                    cmd.CommandText += ",ChousaIraiHonsuu = '" + c1FlexGrid4.Rows[i][50] + "' ";  // 依頼数

                    //                    if (c1FlexGrid4.Rows[i][51] != null && c1FlexGrid4.Rows[i][51].ToString() != "")
                    //                    {
                    //                        cmd.CommandText += ",ChousaIraiRank = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][51].ToString(), 0, 0) + "' ";  // 依頼ランク
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",ChousaIraiRank = '' ";  // 依頼ランク
                    //                    }

                    //                    cmd.CommandText += ",ChousaHinmokuShimekiribi = '" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i][52].ToString(), 0, 0) + "' ";  // 締切日


                    //                    // 報告済
                    //                    if (c1FlexGrid4.Rows[i][53] != null && c1FlexGrid4.Rows[i][53].ToString() == "True")
                    //                    {
                    //                        cmd.CommandText += ",ChousaHoukokuzumi = 1 ";
                    //                    }
                    //                    else
                    //                    {
                    //                        cmd.CommandText += ",ChousaHoukokuzumi = 0 ";
                    //                    }

                    //                    cmd.CommandText += ",ChousaDeleteFlag = '0' " +                          // 削除フラグ
                    //                        ",ChousaUpdateDate = '" + sysDateTimeStr + "'" +                // 更新日時
                    //                        ",ChousaUpdateUser = '" + UserInfos[0] + "' " +                 // 更新ユーザー
                    //                        ",ChousaUpdateProgram = 'Madoguchi button3_InputStatus_Click' " +  // 更新プログラム
                    //                        ",ChousaShinchokuJoukyou = '" + c1FlexGrid4.Rows[i][56] + "' ";       // 進捗状況

                    //                    cmd.CommandText += "WHERE ChousaHinmokuID ='" + c1FlexGrid4.Rows[i][55] + "' AND MadoguchiID ='" + MadoguchiID + "' ";

                    //                    cmd.ExecuteNonQuery();
                    //                    // 更新メッセージ
                    //                    updmessage2 = 1;
                    //                }
                    //            }
                    //            transaction.Commit();

                    //            // ３．ChousaHinmokuから担当部所の連携を行う（支部備考も）
                    //            String resultMessage = "";
                    //            GlobalMethod.MadoguchiHinmokuRenkeiUpdate_SQL(MadoguchiID, "Madoguchi", UserInfos[0], out resultMessage);

                    //            // メッセージがあれば画面に表示
                    //            if (resultMessage != "")
                    //            {
                    //                set_error(resultMessage);
                    //            }

                    //            transaction = conn.BeginTransaction();
                    //            cmd.Transaction = transaction;

                    //            string table = "ChousaHinmoku";
                    //            // 編集ロック開放
                    //            // Lockテーブル更新
                    //            cmd.CommandText = "DELETE FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
                    //                                "AND LOCK_KEY = '" + MadoguchiID + "' " +
                    //                                "AND LOCK_USER_ID = '" + UserInfos[0] + "' ";
                    //            cmd.ExecuteNonQuery();

                    //            GlobalMethod.outputLogger("Madoguchi button3_InputStatus_Click", "DB更新終了 ID:" + MadoguchiID, "update", "DEBUG");

                    //            // 画面のメッセージ表示 （新規 or 更新）+ 削除
                    //            if (updmessage1 == 1)
                    //            {
                    //                // I20302:調査品目明細を追加しました。
                    //                set_error(GlobalMethod.GetMessage("I20302", ""));
                    //            }
                    //            else if (updmessage2 == 1)
                    //            {
                    //                // I20301:調査品目明細を更新しました。
                    //                set_error(GlobalMethod.GetMessage("I20301", ""));
                    //            }
                    //            if (updmessage3 == 1)
                    //            {
                    //                // I20303:調査品目明細を削除しました。
                    //                set_error(GlobalMethod.GetMessage("I20303", ""));
                    //            }

                    //            transaction.Commit();
                    //        }
                    //        catch (Exception)
                    //        {
                    //            transaction.Rollback();
                    //            throw;
                    //        }
                    //        finally
                    //        {
                    //            conn.Close();
                    //        }
                    //        // 調査品目の削除Keys
                    //        deleteChousaHinmokuIDs = "";
                    //        ChousaHinmokuMode = 0;
                    //        // 編集状態を解除する
                    //        ChousaHinmokuGrid_InputMode();

                    //        // 背景色を通常色に戻す
                    //        for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
                    //        {
                    //            c1FlexGrid4.GetCellRange(i, 6).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                    //            c1FlexGrid4.GetCellRange(i, 7).StyleNew.BackColor = Color.FromArgb(245, 245, 220);
                    //            c1FlexGrid4.GetCellRange(i, 34).StyleNew.BackColor = Color.White;
                    //            c1FlexGrid4.GetCellRange(i, 34).StyleNew.BackColor = Color.White;
                    //            c1FlexGrid4.GetCellRange(i, 13).StyleNew.BackColor = Color.White;
                    //            // 更新が通ったので、
                    //            c1FlexGrid4.Rows[i][57] = "1"; // 0:Insert/1:Select/2:Update
                    //        }
                    //    }
                    //    // errorFlg がtrueの場合
                    //    else
                    //    {
                    //        // エラーメッセージ表示
                    //        if (errmessage1 == 1)
                    //        {
                    //            // E20307: 全体順、個別順が重複しています。
                    //            set_error(GlobalMethod.GetMessage("E20307", ""));
                    //        }
                    //        else if (errmessage2 == 1)
                    //        {
                    //            // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい。
                    //            set_error(GlobalMethod.GetMessage("E20336", ""));
                    //        }
                    //        if (errmessage3 == 1)
                    //        {
                    //            // E20337:半角数字で入力してください。赤背景の項目を修正して下さい。
                    //            set_error(GlobalMethod.GetMessage("E20337", ""));
                    //        }

                    //    }

                    //    // ソート
                    //    c1FlexGrid4.Cols[58].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                    //    c1FlexGrid4.Cols[6].Sort = C1.Win.C1FlexGrid.SortFlags.Ascending;
                    //    c1FlexGrid4.Cols.Move(58, 6);
                    //    c1FlexGrid4.Sort(C1.Win.C1FlexGrid.SortFlags.UseColSort, 58, 6);
                    //    c1FlexGrid4.Cols.Move(6, 58);
                    //}
                }
            }
        }

        // 調査品目の更新処理
        private void chousaHinmokuUpdate()
        {
            string methodName = ".chousaHinmokuUpdate";

            writeHistory("【開始】調査品目明細の更新を開始します。 ID= " + MadoguchiID);

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
            //奉行エクセル
            ",ChousaShuukeihyouVer"+
            ",ChousaBunkatsuHouhou" +
            ",ChousaKoujiKouzoubutsumei" +
            ",ChousaTaniAtariTanka" +
            ",chousaTaniAtariSuuryou" +
            ",ChousaTaniAtariKakaku" +
            ",ChousaNiwatashiJouken"+
            ",ChousaHachushaTeikyouTani"+
            ",ChousaMadoguchiGroupMasterID"+
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

                // 全体順と個別順のIndex（行番号）を取得する。
                int ZentaiJunColIndex = c1FlexGrid4.Cols["ChousaZentaiJun"].Index;
                int KobetsuJunColIndex = c1FlexGrid4.Cols["ChousaKobetsuJun"].Index;

                //writeHistory("Grid内重複チェック開始 ID= " + MadoguchiID);

                // 全体順でソート
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
                    ////zentai = (c1FlexGrid4.Rows[i]["ChousaZentaiJun"] == null ? "0" : c1FlexGrid4.Rows[i]["ChousaZentaiJun"].ToString());  // 全体順
                    //kobetsu = c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString(); // 個別順
                    //kobetsu = (c1FlexGrid4.Rows[i]["ChousaKobetsuJun"] == null ? "0" : c1FlexGrid4.Rows[i]["ChousaKobetsuJun"].ToString()); // 個別順

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
                        errorFlg = true;
                        // VIPS　20220314　課題管理表No1293（987）　ADD　Garoon連携直前の更新処理が正常終了チェック
                        globalErrorFlg = "1";
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

                //writeHistory("Grid内重複チェック終了 ID= " + MadoguchiID);


                //writeHistory("DB重複チェック開始 ID= " + MadoguchiID);

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

                //writeHistory("DB重複チェック終了 ID= " + MadoguchiID);

                // 34:地区割コードが空でなく、半角英数字でない場合、正規表現「^[0-9a-zA-Z]+$」
                // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい。

                // 35:地区コードが空でなく、半角英数字でない場合、正規表現「^[0-9a-zA-Z]+$」
                // E20336:半角英数字で入力してください。赤背景の項目を修正して下さい


                //writeHistory("エラーチェック開始 ID= " + MadoguchiID);

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

                //writeHistory("エラーチェック終了 ID= " + MadoguchiID);

                GlobalMethod.outputLogger("Madoguhi button3_InputStatus_Click", "DB比較終了 ID:" + MadoguchiID, "update", "DEBUG");


                //writeHistory("データ登録・更新開始 ID= " + MadoguchiID);

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
                                valuesText += ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaObiMei"].ToString(), 0, 0) + "' " +  // 属性
                                    ",N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaZenkaiTani"].ToString(), 0, 0) + "' ";         // 前回単位

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
                                //奉行エクセル
                                //集計表Ver
                                if (c1FlexGrid4.Rows[i]["ShukeihyoVer"] != null && c1FlexGrid4.Rows[i]["ShukeihyoVer"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["ShukeihyoVer"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",null ";
                                }
                                //分割方法
                                if (c1FlexGrid4.Rows[i]["BunkatsuHouhou"] != null && c1FlexGrid4.Rows[i]["BunkatsuHouhou"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["BunkatsuHouhou"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",'0' ";
                                }
                                //工事・構造物名
                                if (c1FlexGrid4.Rows[i]["KojiKoubutsuMei"] != null && c1FlexGrid4.Rows[i]["KojiKoubutsuMei"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["KojiKoubutsuMei"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",null ";
                                }
                                // 単位当たり単価（単位）
                                if (c1FlexGrid4.Rows[i]["TaniAtariTankaTani"] != null && c1FlexGrid4.Rows[i]["TaniAtariTankaTani"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["TaniAtariTankaTani"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",null ";
                                }
                                //単位当たり単価（数量）
                                if (c1FlexGrid4.Rows[i]["TaniAtariTankaSuryo"] != null && c1FlexGrid4.Rows[i]["TaniAtariTankaSuryo"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["TaniAtariTankaSuryo"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",null ";
                                }
                                // 単位当たり単価（価格）
                                if (c1FlexGrid4.Rows[i]["TaniAtariTankaKakaku"] != null && c1FlexGrid4.Rows[i]["TaniAtariTankaKakaku"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["TaniAtariTankaKakaku"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",null ";
                                }
                                //荷渡し条件
                                if (c1FlexGrid4.Rows[i]["NiwatashiJoken"] != null && c1FlexGrid4.Rows[i]["NiwatashiJoken"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["NiwatashiJoken"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",null ";
                                }
                                //発注者提供単位
                                if (c1FlexGrid4.Rows[i]["HachusyaTeikyoTani"] != null && c1FlexGrid4.Rows[i]["HachusyaTeikyoTani"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["HachusyaTeikyoTani"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",null ";
                                }
                                // グループ名
                                if (c1FlexGrid4.Rows[i]["GroupMei"] != null && c1FlexGrid4.Rows[i]["GroupMei"].ToString() != "")
                                {
                                    valuesText += ",'" + c1FlexGrid4.Rows[i]["GroupMei"] + "' ";
                                }
                                else
                                {
                                    valuesText += ",0 ";
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
                                    ",ChousaZaiKou = N'" + c1FlexGrid4.Rows[i]["ChousaZaiKou"] + "' " +                                                                  // 材工
                                    //",ChousaMadoguchiGroupMasterID = N'" + c1FlexGrid4.Rows[i]["GroupMei"] + "' " +
                                    ",ChousaHinmei = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaHinmei"].ToString(), 0, 0) + "' " +                     // 品目
                                    ",ChousaKikaku = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaKikaku"].ToString(), 0, 0) + "' " +                     // 規格
                                    ",ChousaTanka = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaTanka"].ToString(), 0, 0) + "' " +                       // 単位
                                    ",ChousaSankouShitsuryou = N'" + GlobalMethod.ChangeSqlText(c1FlexGrid4.Rows[i]["ChousaSankouShitsuryou"].ToString(), 0, 0) + "' ";  // 参考質量


                                // グループ名
                                if (c1FlexGrid4.Rows[i]["GroupMei"] != null && c1FlexGrid4.Rows[i]["GroupMei"].ToString() != "")
                                {
                                    cmd.CommandText += ",ChousaMadoguchiGroupMasterID = '" + c1FlexGrid4.Rows[i]["GroupMei"] + "' ";
                                }
                                else
                                {
                                    cmd.CommandText += ",ChousaMadoguchiGroupMasterID = 0 ";
                                }

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
                                    ",ChousaSankouti = N'" + GlobalMethod.ChangeSqlText((c1FlexGrid4.Rows[i]["ChousaSankouti"] != null ? c1FlexGrid4.Rows[i]["ChousaSankouti"].ToString() : "0"), 0, 0) + "' " + // 発注者提供単価
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
                                cmd.CommandText += ",ChousaKonkyoCode = '" + c1FlexGrid4.Rows[i]["ChousaKonkyoCode"] + "' " +   // 根拠関連コード
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
                                //奉行エクセル
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
                                    cmd.CommandText += ",ChousaBunkatsuHouhou = '0' ";
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
                                //cmd.CommandText += "WHERE ChousaHinmokuID ='" + c1FlexGrid4.Rows[i]["ChousaHinmokuID2"] + "' AND MadoguchiID ='" + MadoguchiID + "' ";
                                cmd.CommandText += "WHERE ChousaHinmokuID ='" + c1FlexGrid4.Rows[i]["ChousaHinmokuID2"] + "' ";

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
                                                if(c1FlexGrid4.Rows[i]["HinmokuRyakuBushoCD"] != null) 
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
                        GlobalMethod.MadoguchiHinmokuRenkeiUpdate_SQL(MadoguchiID, "Madoguchi", UserInfos[0], out resultMessage);


                        //writeHistory("データ登録・更新終了 ID= " + MadoguchiID);

                        // メッセージがあれば画面に表示
                        if (resultMessage != "")
                        {
                            set_error(resultMessage);
                        }

                        transaction = conn.BeginTransaction();
                        cmd.Transaction = transaction;

                        string table = "ChousaHinmoku";
                        // 編集ロック開放
                        // Lockテーブル更新
                        cmd.CommandText = "DELETE FROM T_LOCK WHERE LOCK_TABLE = '" + table + "' " +
                                            "AND LOCK_KEY = '" + MadoguchiID + "' " +
                                            "AND LOCK_USER_ID = '" + UserInfos[0] + "' ";
                        cmd.ExecuteNonQuery();

                        GlobalMethod.outputLogger("Madoguchi button3_InputStatus_Click", "DB更新終了 ID:" + MadoguchiID, "update", "DEBUG");

                        item7_HenshuJoutai.Text = "編集ロック無";
                        item7_HenshuLockTantousha.Text = "";

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


                    //writeHistory("背景色戻す開始 ID= " + MadoguchiID);

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

                    //writeHistory("背景色戻す終了 ID= " + MadoguchiID);

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
                    //else if (errmessage2 == 1)
                    if (errmessage2 == 1)
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

        }


        //タブ遷移時
        private void tab_SelectedIndexChanged(object sender, EventArgs e)
        {

            GlobalMethod.outputLogger("Madoguchi", "Madoguchi_Input タブ移動開始:" + ((TabControl)sender).SelectedTab.Text + ":" + DateTime.Now.ToString(), "", "DEBUG");
            openSekouTab = "0"; // 施工条件タブを開いているか 0:開いてない 1:開いている

            //レイアウトロジックを停止する
            this.SuspendLayout();

            // c1FlexGridを隠す
            // 担当部所
            c1FlexGrid1.Visible = false;
            c1FlexGrid5.Visible = false;
            // 調査品目明細
            c1FlexGrid4.Visible = false;
            // 単品入力項目
            c1FlexGrid2.Visible = false;
            // 備考
            BikoGrid.Visible = false;
            if (((TabControl)sender).SelectedTab.Text == "担当部所")
            {
                get_data(2);
                c1FlexGrid1.Visible = true;
                c1FlexGrid5.Visible = true;
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

                //tableLayoutPanel77.Visible = false;
                //tableLayoutPanel80.Visible = false;
                //c1FlexGrid4.BeginUpdate();

                //c1FlexGrid4.AutoResize = false;
                //c1FlexGrid4.Redraw = false;

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
                                    // えんとり君修正STEP2　並び順追加
                                    //+ " ORDER BY TankaKeiyakuID, TankaRankID"
                                    + " ORDER BY TankaKeiyakuID, TankaRankNarabijunn, TankaRankID"
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
                                    // えんとり君修正STEP2　並び順追加
                                    //+ " ORDER BY TankaKeiyakuID, TankaRankID"
                                    + " ORDER BY TankaKeiyakuID, TankaRankNarabijunn, TankaRankID"
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

                // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件 パフォーマンス向上の為
                chousaHinmokuDispFlg = "0";
                // 調査品目明細のGridに全件読み込んだかどうかのフラグ 0:未 1:済
                chousaHinmokuLoadFlg = "0";

                // .00が残る対応
                c1FlexGrid4.EditOptions -= C1.Win.C1FlexGrid.EditFlags.UseNumericEditor;

                chousaHinmokuClear();
                //item_Hyoujikensuu.SelectedIndex = 1;
                item_Hyoujikensuu.SelectedIndex = 4; // 1000件対応
                get_data(3);

                // 1:報告完了の場合、ボタン制御
                if (MadoguchiHoukokuzumi == "1")
                {
                    // 入力開始・文字変換を非活性
                    // 入力開始
                    button3_InputStatus.BackColor = Color.DimGray;
                    button3_InputStatus.Enabled = false;
                    // 文字変換
                    button3_ChangeMoji.BackColor = Color.DimGray;
                    button3_ChangeMoji.Enabled = false;
                }
                else
                {
                    // 入力開始
                    button3_InputStatus.BackColor = Color.FromArgb(42, 78, 122);
                    button3_InputStatus.Enabled = true;
                    // 文字変換
                    button3_ChangeMoji.BackColor = Color.FromArgb(42, 78, 122);
                    button3_ChangeMoji.Enabled = true;
                }


                DataTable tempdt = new DataTable();
                string discript = "LOCK_USER_MEI";
                string value = "LOCK_USER_ID";
                string table = "T_LOCK";
                string where = "LOCK_KEY = '" + MadoguchiID + "' ORDER BY LOCK_DATETIME DESC ";
                tempdt = new DataTable();
                tempdt = GlobalMethod.getData(discript, value, table, where);
                if (tempdt != null && tempdt.Rows.Count > 0)
                {
                    // 窓口ミハルのみ編集ロックチェック
                    if (tempdt != null && tempdt.Rows.Count > 0)
                    {
                        item7_HenshuJoutai.Text = "編集ロック中";
                        item7_HenshuLockTantousha.Text = tempdt.Rows[0][1].ToString();
                    }
                    else
                    {
                        item7_HenshuJoutai.Text = "編集ロック無";
                        item7_HenshuLockTantousha.Text = "";
                    }
                }

                //c1FlexGrid4.AutoResize = true;
                //c1FlexGrid4.Redraw = true;

                c1FlexGrid4.Visible = true;

                //c1FlexGrid4.EndUpdate();
                //tableLayoutPanel80.Visible = true;
                //tableLayoutPanel77.Visible = true;
            }
            if (((TabControl)sender).SelectedTab.Text == "協力依頼書")
            {
                //get_data(5);
                get_data(4);
                // 1:報告完了の場合、ボタン制御
                if (MadoguchiHoukokuzumi == "1")
                {
                    // 入力開始・文字変換を非活性
                    // 依頼書出力
                    button24.BackColor = Color.DimGray;
                    button24.Enabled = false;
                }
                else
                {
                    // 依頼書出力
                    button24.BackColor = Color.FromArgb(42, 78, 122);
                    button24.Enabled = true;
                }
            }
            if (((TabControl)sender).SelectedTab.Text == "応援受付状況")
            {
                get_data(5);
                // 1:報告完了の場合、ボタン制御
                if (MadoguchiHoukokuzumi == "1")
                {
                    // 入力開始・文字変換を非活性
                    // 依頼書出力
                    btnIraisho.BackColor = Color.DimGray;
                    btnIraisho.Enabled = false;
                }
                else
                {
                    // 依頼書出力
                    btnIraisho.BackColor = Color.FromArgb(42, 78, 122);
                    btnIraisho.Enabled = true;
                }
            }
            if (((TabControl)sender).SelectedTab.Text == "単品入力項目")
            {
                get_data(6);
                c1FlexGrid2.Visible = true;
            }
            if (((TabControl)sender).SelectedTab.Text == "施工条件")
            {
                openSekouTab = "1";
                sekouMeijishoComboChangeFlg = "0";
                // 施工条件タブ 施工条件明示書ID変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoIDChangeFlg = "0";
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

            GlobalMethod.outputLogger("Madoguchi", "Madoguchi_Input タブ移動終了:" + ((TabControl)sender).SelectedTab.Text + ":" + DateTime.Now.ToString(), "", "DEBUG");
        }

        private void button3_ExcelShukeihyo_Click(object sender, EventArgs e)
        {
            // 奉行エクセル移管対応
            if (IsPopup_ShukeiHyou_New.Equals("1"))
            {
                Popup_ShukeiHyou_New form = new Popup_ShukeiHyou_New();
                form.MadoguchiID = MadoguchiID;
                form.Busho = UserInfos[2];
                form.TokuhoBangou = item1_MadoguchiUketsukeBangou.Text;
                form.TokuhoBangouEda = item1_MadoguchiUketsukeBangouEdaban.Text;
                form.KanriBangou = item1_MadoguchiKanriBangou.Text;
                form.UserInfos = UserInfos;
                form.PrintGamen = "Madoguchi";
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
            else
            {
                Popup_ShukeiHyou form = new Popup_ShukeiHyou();
                form.MadoguchiID = MadoguchiID;
                form.Busho = UserInfos[2];
                form.TokuhoBangou = item1_MadoguchiUketsukeBangou.Text;
                form.TokuhoBangouEda = item1_MadoguchiUketsukeBangouEdaban.Text;
                form.KanriBangou = item1_MadoguchiKanriBangou.Text;
                form.UserInfos = UserInfos;
                form.PrintGamen = "Madoguchi";
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
            ////集計表プロンプト
            //Popup_ShukeiHyou form = new Popup_ShukeiHyou();
            ////form.nendo = item1_3.SelectedValue.ToString();
            //form.MadoguchiID = MadoguchiID;
            //form.Busho = UserInfos[2];
            //form.TokuhoBangou = item1_MadoguchiUketsukeBangou.Text;
            //form.TokuhoBangouEda = item1_MadoguchiUketsukeBangouEdaban.Text;
            //form.KanriBangou = item1_MadoguchiKanriBangou.Text;
            //form.UserInfos = UserInfos;
            //form.PrintGamen = "Madoguchi";
            //form.ShowDialog();

            //if (form.ReturnValue != null && form.ReturnValue[0] != null)
            //{
            //    //item1_MadoguchiTantoushaCD.Text = form.ReturnValue[0];
            //    //item1_MadoguchiTantousha.Text = form.ReturnValue[1];

            //    //item_Hyoujikensuu.SelectedIndex = 1;
            //    item_Hyoujikensuu.SelectedIndex = 4; // 1000件対応
            //    // データ取り直し
            //    get_data(3);
            //}
        }

        // 調査概要タブ 実施区分
        private void item1_MadoguchiJiishiKubun_TextChanged(object sender, EventArgs e)
        {
            //実施区分が3：中止
            if (item1_MadoguchiJiishiKubun.SelectedValue != null && "3".Equals(item1_MadoguchiJiishiKubun.SelectedValue.ToString()))
            {
                //報告済を1のまま
                item1_MadoguchiHoukokuzumi.Checked = true;

                //報告完了取消ボタンを表示
                //報告完了ボタンを非表示
                button9.Text = "報告完了取消";

                MadoguchiHoukokuzumi = "1";
            }
            else
            {
                // 不具合No1342
                //登録済みの実施区分が「打診中：2」の場合は、実施に変更されても、報告済みを戻さない
                if ("2".Equals(originalJiishiKubun))
                {
                    //何もしない 今後追加修正ありそうなのであえてElseに分けた
                    //Console.WriteLine("実施区分テキスト：2");
                }
                else
                {
                    //報告済を0にする
                    item1_MadoguchiHoukokuzumi.Checked = false;

                    //報告完了ボタンを表示
                    //報告完了取消ボタンを非表示
                    button9.Text = "報告完了";

                    MadoguchiHoukokuzumi = "0";
                }
               
                //////報告済を0にする
                ////item1_MadoguchiHoukokuzumi.Checked = false;

                //////報告完了ボタンを表示
                //////報告完了取消ボタンを非表示
                ////button9.Text = "報告完了";

                ////MadoguchiHoukokuzumi = "0";
            }
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

        private void pictureBox11_Click(object sender, EventArgs e)
        {

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

        // TOP
        private void button5_Click(object sender, EventArgs e)
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

        // ヘッダーの窓口ミハルボタン
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
        // 特命課長
        private void button8_Click(object sender, EventArgs e)
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

        // 自分大臣
        private void button19_Click(object sender, EventArgs e)
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
            report_data[8] = "0";   // 0:窓口ミハル

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

        //編集ロック解除
        private void button3_Lock_Click(object sender, EventArgs e)
        {
            //編集ロックを解除しますがよろしいですか。
            if (MessageBox.Show(GlobalMethod.GetMessage("I20310", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                string connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                SqlConnection sqlconn = new SqlConnection(connStr);
                sqlconn.Open();

                var cmd = sqlconn.CreateCommand();
                SqlTransaction transaction = sqlconn.BeginTransaction();
                cmd.Transaction = transaction;

                // ロック（T_LOCK）テーブルからLOCK_TABLEをChousaHinmoku、LOCK_KEYを窓口IDで、
                // データを削除
                cmd.CommandText = "DELETE FROM T_LOCK " +
                    "Where LOCK_TABLE = 'ChousaHinmoku' AND LOCK_KEY = '" + MadoguchiID + "' ";
                cmd.ExecuteNonQuery();

                transaction.Commit();
                sqlconn.Close();

                item7_HenshuJoutai.Text = "編集ロック無"; // 編集状態
                item7_HenshuLockTantousha.Text = ""; // 編集ロック中の担当者

                // E20325:編集ロックの強制解除を行いました。
                writeHistory(GlobalMethod.GetMessage("E20325","") + " 窓口ID = " + MadoguchiID);
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
                //// I20006:Garoon送信ボタンからOKが押下されました。
                //GlobalMethod.outputLogger("GaroonBtn_Click", GlobalMethod.GetMessage("I20006", "") + " ID:" + MadoguchiID + " Garoon連携区分:" + item1_GaroonRenkei.ToString(), "insert", "DEBUG");

                // Garoon連携フラグ  1:連携 0:チェックなし
                string GaroonRenkei = "0";

                DataTable combodt = new DataTable();
                combodt = GlobalMethod.getData("MadoguchiGaroonRenkei", "MadoguchiGaroonRenkei", "MadoguchiJouhou", "MadoguchiID = '" + MadoguchiID + "' ");

                if (combodt != null && combodt.Rows.Count > 0)
                {
                    GaroonRenkei = combodt.Rows[0][0].ToString();
                }

                // エラーフラグ true:エラー false:正常
                Boolean errorFlg = false;
                // Garoon連携対象チェック
                //if (!item1_GaroonRenkei.Checked)
                if (GaroonRenkei == "0")
                {
                    // I20005:Garoonとの連携対象ではありません。
                    set_error(GlobalMethod.GetMessage("I20005", ""));
                    errorFlg = true;
                }

                string MadoguchiTantoushaCD = "";
                combodt = new DataTable();
                combodt = GlobalMethod.getData("MadoguchiTantoushaCD", "MadoguchiTantoushaCD", "MadoguchiJouhou", "MadoguchiID = '" + MadoguchiID + "' ");

                if (combodt != null && combodt.Rows.Count > 0)
                {
                    MadoguchiTantoushaCD = combodt.Rows[0][0].ToString();
                }

                // 窓口担当者チェック
                //if (item1_MadoguchiTantousha.Text == "")
                if (MadoguchiTantoushaCD == "")
                {
                    // E20011:窓口担当者が未登録のため、Garoon連携ができません。
                    set_error(GlobalMethod.GetMessage("E20011", ""));
                    errorFlg = true;
                }

                // 連携処理
                if(errorFlg == false)
                {
                    string w_MadoguchiMailGaRenkeiKubun = "";
                    string w_MadoguchiMailMessageID = "";
                    string w_MadoguchiUketsukeBangou = "";
                    string w_MadoguchiUketsukeBangouEdaban = "";
                    string w_MadoguchiTantoushaCD = "";
                    string w_TokuchoBangou = "";
                    string w_MailInfoCSVWorkAtesakiUser = "";
                    //string w_KojinCD = "";
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
                            //string GaroonRenkei = "";
                            ////if (item1_GaroonRenkei.Checked)
                            //if (GaroonRenkei == "1")
                            //{
                            //    GaroonRenkei = "1";
                            //}
                            //else
                            //{
                            //    GaroonRenkei = "0";
                            //}

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
                            if(w_MadoguchiKanriGijutsusha != "" && w_MadoguchiKanriGijutsusha != "0")
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
                                for(int i = 0; i < Dt.Rows.Count; i++) { 
                                    //w_MadoguchiL1ChousaBushoCD = Dt.Rows[0][0].ToString();
                                    w_MadoguchiL1ChousaTantoushaCD = Dt.Rows[i][1].ToString();

                                    // 調査担当者が存在すれば
                                    if (w_MadoguchiL1ChousaTantoushaCD != "" && w_MadoguchiL1ChousaTantoushaCD != "0")
                                    {
                                        w_MailInfoCSVWorkAtesakiUser = SetChousain(w_MadoguchiL1ChousaTantoushaCD, w_MailInfoCSVWorkAtesakiUser);
                                    }
                                    // 支部応援マスタから担当調査員の部所に該当する調査員を設定する
                                    if(w_MadoguchiL1ChousaBushoCD != Dt.Rows[i][0].ToString())
                                    {
                                        w_MadoguchiL1ChousaBushoCD = Dt.Rows[i][0].ToString();
                                        w_MailInfoCSVWorkAtesakiUser = GetShibuouen(w_MadoguchiL1ChousaBushoCD, w_MailInfoCSVWorkAtesakiUser);
                                    }
                                }
                            }

                            // Garoon追加追加宛先の調査員も追加する
                            DataTable  GaroonDt = new System.Data.DataTable();
                            //SQL生成
                            cmd.CommandText = "SELECT " +
                              "GaroonTsuikaAtesakiBushoCD,GaroonTsuikaAtesakiTantoushaCD " +
                              "FROM GaroonTsuikaAtesaki " +
                              "WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "' " +
                              " AND GaroonTsuikaAtesakiDeleteFlag <> 1";

                            //データ取得
                            sda = new SqlDataAdapter(cmd);
                            sda.Fill(GaroonDt);

                            // GaroonTsuikaAtesakiに登録されているデータを取得
                            if (GaroonDt != null && GaroonDt.Rows.Count > 0)
                            {
                                for(int i = 0; i < GaroonDt.Rows.Count; i++) { 
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
                                  "WHERE MailInfoCSVWorkTokuchoBangou COLLATE Japanese_XJIS_100_CI_AS_SC = N'" + w_TokuchoBangou + "' AND MailInfoCSVWorkCSVOutFlg = 0 AND MailInfoCSVWorkGaRenkeiFlg = 0 AND MailInfoCSVWorkDeleteFlag = 0";

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
                                        ",SYSDATETIME()" +
                                        ",N'" + UserInfos[0] + "' " +
                                        ",'" + pgmName + methodName + "'" +
                                        ",SYSDATETIME()" +
                                        ",N'" + UserInfos[0] + "' " +
                                        ",'" + pgmName + methodName + "'" +
                                        ",0" + 
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
                            "MadoguchiGaroonRenkeiJikouDate = '" + datetTime + "' " +
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
                if(MailInfoCSVWorkAtesakiUser == "") 
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
                        GlobalMethod.outputLogger("SetChousain","ID:" + MadoguchiID + " Garoon連携で宛先ユーザーの文字数が2048を超える為、KojinCD:" + w_KojinCD + " を追加できませんでした。", "insert", "DEBUG");
                    }
                }
            }
            return MailInfoCSVWorkAtesakiUser;
        }

        // 支部応援の取得
        private string GetShibuouen(string w_MadoguchiL1ChousaBushoCD,string MailInfoCSVWorkAtesakiUser)
        {
            string discript = "ShibuouenKojinCD ";
            string value = "ShibuouenKojinCD ";
            string table = "Mst_Shibuouen ";
            string where = "ShibuouenDeleteFlag = 0 Order By ShibuouenKojinCD ";
            //No.1547
            //string where = "(ShibuouenDeleteFlag = 0 or ShibuouenDeleteFlag = 1) Order By ShibuouenKojinCD ";
            string w_ShibuouenKojinCD = "";

            // データ取得
            var tmpdt = GlobalMethod.getData(discript, value, table, where);
            DataTable dt = new DataTable();

            if (tmpdt != null && tmpdt.Rows.Count > 0)
            {
                for(int i = 0; i < tmpdt.Rows.Count; i++) {  
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

        // 調査品目明細の調査担当者プロンプト
        private void item3_ChousainPrompt(object sender, EventArgs e)
        {
            // 調査担当者プロンプト
            Popup_ChousainList form = new Popup_ChousainList();
            //選択されている年度を条件に調査員プロンプトを表示
            if (item1_MadoguchiTourokuNendo.Text != "")
            {
                form.nendo = item1_MadoguchiTourokuNendo.SelectedValue.ToString();
            }
            // 1219 調査品目一覧＿調査担当部所と調査担当者の選択画面の部所名が連動してほしい
            // 窓口部所と連動していたのを調査品目明細タブの調査担当部所と連動するよう修正
            //// 窓口部所セット
            //if (item1_MadoguchiTantoushaBushoCD.SelectedValue != null)
            //{
            //    form.Busho = item1_MadoguchiTantoushaBushoCD.SelectedValue.ToString();
            //}
            if (src_Busho.Text != "")
            {
                form.Busho = src_Busho.SelectedValue.ToString();
            }
            form.program = "madoguchi";
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                src_HinmokuChousain.Text = form.ReturnValue[1];
                src_Busho.SelectedValue = form.ReturnValue[2];
            }
            src_HinmokuChousain.Focus();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            // エラークリア
            ErrorClear_KyouryokuIrai();

            // 協力依頼書出力
            if (MessageBox.Show("依頼書出力をするには、現在の内容で更新する必要があります。よろしいですか。", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                // エラーチェック
                if (ErrorCheck_KyouryokuIrai())
                {
                    set_error(GlobalMethod.GetMessage("E20403", "")); 
                }
                else
                {
                    // 協力依頼書の更新
                    UpdateMadoguchi(4);
                    // 802 応援受付を更新する
                    UpdateMadoguchi(5);

                    // 協力依頼書の出力
                    output_KyouryokuIrai("KyouryokuIrai");

                }
            }
        }

        private void output_KyouryokuIrai(string YobidashiMoto)
        {
            set_error("", 0);

            // 0:MadoguchiID    窓口ID
            // 1:YobidashiMoto  呼び出し元タブ
            string[] report_data = new string[2] { "", "" };
            report_data[0] = MadoguchiID;

            switch (YobidashiMoto)
            {
                case "KyouryokuIrai":   // 協力依頼書タブ
                    report_data[1] = "1";
                    break;
                case "OuenUketsuke":    // 応援受付タブ
                    report_data[1] = "2";
                    break;
                default:
                    break;
            }

            int listID = 19;    // 協力依頼書

            // 協力依頼書の出力
            string[] result = GlobalMethod.InsertMadoguchiReportWork(listID, UserInfos[0], report_data, "KyouryokuIrai");

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
                    // 成功時は、ファイルをフォルダにコピーする
                    string copyFileName = result[3];
                    // 1056対応 _YYYYMMDDHHMMSSを付ける
                    // .xlsxと.xlsmを対象に置換する
                    copyFileName = copyFileName.Replace(".xlsx", "_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                    copyFileName = copyFileName.Replace(".xlsm", "_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsm");
                    try
                    {
                        //System.IO.File.Copy(result[2], item1_MadoguchiShiryouHolder.Text + @"\" + result[3], true);
                        System.IO.File.Copy(result[2], item1_MadoguchiShiryouHolder.Text + @"\" + copyFileName, true);

                        // 応援受付状況の取得
                        get_data(5);

                        // 応援状況の更新
                        DT_Ouenuketsuke.Rows[0][1] = 1;
                        UpdateMadoguchi(5);

                        // 応援受付状況の再取得
                        get_data(5);

                        //set_error("協力依頼書ファイルを出力しました。:" + result[3]);
                        set_error("協力依頼書ファイルを出力しました。:" + copyFileName);

                    }
                    catch (Exception)
                    {
                        // ファイルコピー失敗
                        //set_error("ファイルコピー失敗:" + result[3]);
                        set_error("ファイルコピー失敗:" + copyFileName);
                    }

                }
            }

        }

        // 協力依頼書タブのエラークリア
        private void ErrorClear_KyouryokuIrai()
        {
            set_error("", 0);

            item4_KyoRyokuBusho.BackColor = Color.FromArgb(255, 255, 255);      // 協力先部所
            label46.BackColor = Color.FromArgb(255, 255, 255);                  // 依頼日
            item4_IraiKubun.BackColor = Color.FromArgb(255, 255, 255);          // 依頼区分
            item4_GyoumuNaiyo.BackColor = Color.FromArgb(255, 255, 255);        // 業務内容
            item4_Zumen.BackColor = Color.FromArgb(255, 255, 255);              // 図面
            item4_KizyunbiStr.BackColor = Color.FromArgb(255, 255, 255);        // 調査基準日
            item4_UtiawaseYouhi.BackColor = Color.FromArgb(255, 255, 255);      // 打合せ要否
            item4_ZenkaiKyouryoku.BackColor = Color.FromArgb(255, 255, 255);    // 前回協力
            item4_Hikiwatashi.BackColor = Color.FromArgb(255, 255, 255);        // 成果物引渡場所
            item4_MitsumoriChousyu.BackColor = Color.FromArgb(255, 255, 255);   // 見積徴収

        }

        // 協力依頼書タブの必須チェック
        private Boolean ErrorCheck_KyouryokuIrai()
        {
            // エラーフラグ true:エラー /false:正常
            Boolean errorFlg = false;

            // 協力先部所
            if (String.IsNullOrEmpty(item4_KyoRyokuBusho.Text))
            {
                item4_KyoRyokuBusho.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            // 依頼日
            if (item4_Iraibi.CustomFormat != "")
            {
                label46.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            // 依頼区分
            if (String.IsNullOrEmpty(item4_IraiKubun.Text))
            {
                item4_IraiKubun.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            // 業務内容
            if (String.IsNullOrEmpty(item4_GyoumuNaiyo.Text))
            {
                item4_GyoumuNaiyo.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            // 図面
            if (String.IsNullOrEmpty(item4_Zumen.Text))
            {
                item4_Zumen.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }
            
            // 調査基準日
            if (String.IsNullOrEmpty(item4_KizyunbiStr.Text))
            {
                item4_KizyunbiStr.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            // 打合せ要否
            if (String.IsNullOrEmpty(item4_UtiawaseYouhi.Text))
            {
                item4_UtiawaseYouhi.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            // 前回協力
            if (String.IsNullOrEmpty(item4_ZenkaiKyouryoku.Text))
            {
                item4_ZenkaiKyouryoku.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            // 成果物引渡場所
            if (String.IsNullOrEmpty(item4_Hikiwatashi.Text))
            {
                item4_Hikiwatashi.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            // 見積徴収
            if (String.IsNullOrEmpty(item4_MitsumoriChousyu.Text))
            {
                item4_MitsumoriChousyu.BackColor = Color.FromArgb(255, 204, 255);
                errorFlg = true;
            }

            return errorFlg;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            set_error("", 0);
            if (item1_PrintList.Text == "")
            {
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
                    cmd.CommandText = "SELECT " +
                      "PrintDataPattern,PrintKikanFlg " +
                      "FROM " + "Mst_PrintList " +
                      "WHERE PrintListID = '" + item1_PrintList.SelectedValue + "'";

                    //データ取得
                    var sda = new SqlDataAdapter(cmd);
                    sda.Fill(Dt);
                    //Boolean errorFLG = false;

                    if (Dt.Rows.Count > 0)
                    {
                        set_error("", 0);
                        // 40:業務連絡票
                        if (Dt.Rows[0][0].ToString() == "40")
                        {
                            string[] report_data = new string[2] { "", "" };
                            report_data[0] = MadoguchiID;
                            report_data[1] = UserInfos[2];

                            string[] result = GlobalMethod.InsertMadoguchiReportWork(int.Parse(item1_PrintList.SelectedValue.ToString()), UserInfos[0], report_data, "Gyoumurenraku");

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
                                    //Popup_Download form = new Popup_Download();
                                    //form.TopLevel = false;
                                    //this.Controls.Add(form);

                                    String fileName = Path.GetFileName(result[3]);
                                    //form.ExcelName = fileName;
                                    //form.TotalFilePath = result[2];
                                    //form.Dock = DockStyle.Bottom;
                                    //form.Show();
                                    //form.BringToFront();
                                    // 成功時は、ファイルをフォルダにコピーする
                                    string copyFileName = result[3];
                                    // 1056対応 _YYYYMMDDHHMMSSを付ける
                                    // .xlsxと.xlsmを対象に置換する
                                    copyFileName = copyFileName.Replace(".xlsx", "_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                                    copyFileName = copyFileName.Replace(".xlsm", "_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsm");
                                    try
                                    {
                                        //System.IO.File.Copy(result[2], item1_MadoguchiShiryouHolder.Text + @"\" + result[3], true);
                                        System.IO.File.Copy(result[2], item1_MadoguchiShiryouHolder.Text + @"\" + copyFileName, true);

                                        //set_error("業務連絡票を出力しました。:" + result[3]);
                                        set_error("業務連絡票を出力しました。:" + copyFileName);

                                    }
                                    catch (Exception)
                                    {
                                        // ファイルコピー失敗
                                        //set_error("ファイルコピー失敗:" + result[3]);
                                        set_error("ファイルコピー失敗:" + copyFileName);
                                    }
                                }
                            }
                            else
                            {
                                // エラーが発生しました
                                set_error(GlobalMethod.GetMessage("E00091", ""));
                            }
                        }
                        // 43:ISO書式集
                        if (Dt.Rows[0][0].ToString() == "43")
                        {
                            string[] report_data = new string[2] { "", "" };
                            report_data[0] = MadoguchiID;
                            report_data[1] = "0";           // 呼び出し元画面（0:窓口ミハル、1:特命課長、2:自分大臣）

                            string[] result = GlobalMethod.InsertMadoguchiReportWork(int.Parse(item1_PrintList.SelectedValue.ToString()), UserInfos[0], report_data, "ISOShosiki");

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

        }

        // 担当部所タブの調査担当者Grid
        private void c1FlexGrid1_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            DateTime dateTime = DateTime.Today;

            // 4:締切日、5:担当者状況変更
            if (e.Row > 0 && (e.Col == 4 || e.Col == 5))
            {
                // 1:報告済みの場合
                if (item1_MadoguchiHoukokuzumi.Checked)
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
                        // 担当者済み
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

        // 応援受付状況タブ 結果送付書
        private void btnKekkasoufusho_Click(object sender, EventArgs e)
        {
            set_error("", 0);

            // 1:MadoguchiID     窓口ID
            string[] report_data = new string[1] { "" };
            report_data[0] = MadoguchiID;
            int listID = 18;    // 結果送付書

            // 協力依頼結果送付書の出力
            string[] result = GlobalMethod.InsertMadoguchiReportWork(listID, UserInfos[0], report_data, "Kekkasouhusho");

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
                    // 成功時は、ファイルをフォルダにコピーする
                    string copyFileName = result[3];
                    // 1056対応 _YYYYMMDDHHMMSSを付ける
                    // .xlsxと.xlsmを対象に置換する
                    copyFileName = copyFileName.Replace(".xlsx", "_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
                    copyFileName = copyFileName.Replace(".xlsm", "_" + System.DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsm");
                    try
                    {
                        //System.IO.File.Copy(result[2], item1_MadoguchiShiryouHolder.Text + @"\" + result[3], true);
                        System.IO.File.Copy(result[2], item1_MadoguchiShiryouHolder.Text + @"\" + copyFileName, true);

                        //set_error("協力依頼結果送付書ファイルを出力しました。:" + result[3]);
                        set_error("協力依頼結果送付書ファイルを出力しました。:" + copyFileName);

                    }
                    catch (Exception)
                    {
                        // ファイルコピー失敗
                        //set_error("ファイルコピー失敗:" + result[3]);
                        set_error("ファイルコピー失敗:" + copyFileName);
                    }

                }
            }

        }

        // KeyUp / KeyDown / ↑ / ↓
        // Form の KeyPreview を true に設定すること
        private void Madoguchi_InputKeyDown(object sender, KeyEventArgs e)
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
                    if ("単品入力項目".Equals(tabName))
                    {
                        this.tabPage6.AutoScrollPosition = new System.Drawing.Point(-this.tabPage6.AutoScrollPosition.X, -this.tabPage6.AutoScrollPosition.Y + 600);
                    }
                    if ("施工条件".Equals(tabName))
                    {
                        this.tabPage7.AutoScrollPosition = new System.Drawing.Point(-this.tabPage7.AutoScrollPosition.X, -this.tabPage7.AutoScrollPosition.Y + 600);
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
                    if ("単品入力項目".Equals(tabName))
                    {
                        this.tabPage6.AutoScrollPosition = new System.Drawing.Point(-this.tabPage6.AutoScrollPosition.X, -this.tabPage6.AutoScrollPosition.Y - 600);
                    }
                    if ("施工条件".Equals(tabName))
                    {
                        this.tabPage7.AutoScrollPosition = new System.Drawing.Point(-this.tabPage7.AutoScrollPosition.X, -this.tabPage7.AutoScrollPosition.Y - 600);
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
                    if ("単品入力項目".Equals(tabName))
                    {
                        this.tabPage6.AutoScrollPosition = new System.Drawing.Point(-this.tabPage6.AutoScrollPosition.X, -this.tabPage6.AutoScrollPosition.Y + 600);
                    }
                    if ("施工条件".Equals(tabName))
                    {
                        this.tabPage7.AutoScrollPosition = new System.Drawing.Point(-this.tabPage7.AutoScrollPosition.X, -this.tabPage7.AutoScrollPosition.Y + 600);
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
                    if ("単品入力項目".Equals(tabName))
                    {
                        this.tabPage6.AutoScrollPosition = new System.Drawing.Point(-this.tabPage6.AutoScrollPosition.X, -this.tabPage6.AutoScrollPosition.Y - 600);
                    }
                    if ("施工条件".Equals(tabName))
                    {
                        this.tabPage7.AutoScrollPosition = new System.Drawing.Point(-this.tabPage7.AutoScrollPosition.X, -this.tabPage7.AutoScrollPosition.Y - 600);
                    }
                }
            }
            //レイアウトロジックを再開する
            this.ResumeLayout();
        }

        private void item1_MadoguchiShimekiribi_TextChanged(object sender, EventArgs e)
        {
            item4_HoukokuKigen.Value = item1_MadoguchiShimekiribi.Value;
        }

        // 調査担当者の右クリックメニュークリック時のイベント
        private void contextMenuTantoushaItemClicked(object sender, EventArgs e)
        {
            ToolStripMenuItem mi = (ToolStripMenuItem)sender;

            ToolStripMenuItem mi2 = (ToolStripMenuItem)mi.DropDownItems[0];

            ToolStripItemClickedEventArgs ea = (ToolStripItemClickedEventArgs)e;

            //調査員
            DataTable dt = new DataTable();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = conn.CreateCommand();
                //データ取得時に年度がいない場合、当年度とする
                int Nendo;
                int ToNendo;
                if (item1_MadoguchiTourokuNendo.Text == "")
                {
                    Nendo = DateTime.Today.Year;
                    ToNendo = DateTime.Today.AddYears(1).Year;
                }
                else
                {
                    int.TryParse(item1_MadoguchiTourokuNendo.SelectedValue.ToString(), out Nendo);
                    ToNendo = Nendo + 1;
                }
                cmd.CommandText = "SELECT " +
                "mc.KojinCD,mc.ChousainMei " +
                "FROM Mst_Chousain mc " +
                "INNER JOIN Mst_Busho mb  ON mc.GyoumuBushoCD = mb.GyoumuBushoCD " +
                " AND BushoMadoguchiHyoujiFlg = 1 AND BushoNewOld <= 1 AND ISNULL(BushokanriboKamei,'') != ''  " +
                //" AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + Nendo + "/4/01' ) " +
                ////" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + ToNendo + "/3/31' ) " +
                //" AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                " AND (BushoYukoukikanFrom IS NULL OR BushoYukoukikanFrom <= '" + ToNendo + "/3/31' ) " +
                " AND (BushoYukoukikanTo IS NULL OR BushoYukoukikanTo >= '" + Nendo + "/4/01' ) " +
                " AND BushokanriboKameiRaku ='" + ea.ClickedItem.Text + "' " +
                "ORDER BY BushoMadoguchiNarabijun";

                var sda = new SqlDataAdapter(cmd);
                dt.Clear();
                sda.Fill(dt);
                conn.Close();
            }

        }

        // 行追加
        private void button3_RowAdd_Click(object sender, EventArgs e)
        {
            // 行の追加プロンプト
            Popup_GyouTsuika form = new Popup_GyouTsuika();
            form.UserInfos = UserInfos;
            form.Nendo = DateTime.Today.Year;
            if (item1_MadoguchiTourokuNendo.Text != "")
            {
                form.Nendo = int.Parse(item1_MadoguchiTourokuNendo.SelectedValue.ToString());
            }
            form.ShowDialog();

            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                addline(form.ReturnValue);
            }
            // 表示件数（ヘッダー2行分を引く）
            Grid_Num.Text = "(" + (c1FlexGrid4.Rows.Count - 2) + ")";
        }

        // 全件削除ボタン
        private void button3_DeleteAllRow_Click(object sender, EventArgs e)
        {
            //調査品目編集モード 0:表示 1:編集
            if (ChousaHinmokuMode == 1)
            {
                // I20315:全件削除します。よろしいしいですか？
                if (MessageBox.Show(GlobalMethod.GetMessage("I20315", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    //レイアウトロジックを停止する
                    this.SuspendLayout();

                    c1FlexGrid4.Visible = false;

                    writeHistory("【開始】調査品目明細の全件削除を開始します。 ID= " + MadoguchiID);

                    // c1FlexGrid4.Clear()・・・Gridのタイトル、値を消すだけで行も消えない
                    // c1FlexGrid4
                    // 0,1 ヘッダー
                    // 2～ データ

                    // c1FlexGrid4の件数（c1FlexGrid は0始まりなので、件数 -1 をする）
                    int row = c1FlexGrid4.Rows.Count - 1;

                    // 行の削除（ヘッダー2件を残して削除）
                    // 後ろから削除する
                    for (int i = row; i > 1; i--) 
                    {
                        // 55:ChousaHinmokuID があれば保持、無ければそのまま削除
                        //if (c1FlexGrid4.Rows[i][55] != null && c1FlexGrid4.Rows[i][55].ToString() != "") 
                        if (c1FlexGrid4.Rows[i]["ChousaHinmokuID2"] != null && c1FlexGrid4.Rows[i]["ChousaHinmokuID2"].ToString() != "") 
                        { 
                            // 更新時に削除するようにKeyを保持しておく
                            if(deleteChousaHinmokuIDs == "") 
                            {
                                // 55:ChousaHinmokuID
                                //deleteChousaHinmokuIDs = c1FlexGrid4.Rows[i][55].ToString();
                                deleteChousaHinmokuIDs = c1FlexGrid4.Rows[i]["ChousaHinmokuID2"].ToString();
                            }
                            else
                            {
                                //deleteChousaHinmokuIDs = deleteChousaHinmokuIDs + "," + c1FlexGrid4.Rows[i][55].ToString();
                                deleteChousaHinmokuIDs = deleteChousaHinmokuIDs + "," + c1FlexGrid4.Rows[i]["ChousaHinmokuID2"].ToString();
                            }
                        }

                        c1FlexGrid4.RemoveItem(i);
                    }

                    c1FlexGrid4.Visible = true;

                    Paging_now.Text = "1";
                    item3_TargetPage.Text = Paging_now.Text;
                    Paging_all.Text = "1";
                    Grid_Num.Text = "(0)";

                    GlobalMethod.outputLogger("button3_DeleteAllRow_Click", "全件削除 ID:" + MadoguchiID, "AllDelete", "DEBUG");

                    //レイアウトロジックを再開する
                    this.ResumeLayout();
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

        private void button3_ExcelHoukokusho_Click(object sender, EventArgs e)
        {
            // 報告書プロンプト
            Popup_HoukokuSho form = new Popup_HoukokuSho();
            form.MadoguchiID = MadoguchiID;
            //form.MENU_ID = 203;
            form.UserInfos = UserInfos;
            form.PrintGamen = "Madoguchi";
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
                if (item1_MadoguchiTourokuNendo.Text == "")
                {
                    Nendo = DateTime.Today.Year;
                    ToNendo = DateTime.Today.AddYears(1).Year;
                }
                else
                {
                    int.TryParse(item1_MadoguchiTourokuNendo.SelectedValue.ToString(), out Nendo);
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
                            if (sRank == "-")
                            {
                                //No.1443 
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
                //奉行エクセル
                // グループ名
                if (e.Col == c1FlexGrid4.Cols["GroupMei"].Index)
                {
                    strIndex = 15;
                }
                //集計表Ver
                if (e.Col == c1FlexGrid4.Cols["ShukeihyoVer"].Index)
                {
                    if (c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() != "2")
                    {
                        c1FlexGrid4.GetCellRange(e.Row, 58).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                        c1FlexGrid4.GetCellRange(e.Row, 59).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                        c1FlexGrid4.Rows[e.Row][58] = "-";
                        c1FlexGrid4.Rows[e.Row][59] = "";
                    } //No.1622
                    if (c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() == "2" && c1FlexGrid4.Rows[e.Row]["BunkatsuHouhou"].ToString() == "1")
                    {
                        c1FlexGrid4.GetCellRange(e.Row, 59).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                        c1FlexGrid4.Rows[e.Row][59] = "";
                    }
                    if(c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() == "2")
                    {
                        c1FlexGrid4.GetCellRange(e.Row, 58).StyleNew.BackColor = Color.White;
                        c1FlexGrid4.GetCellRange(e.Row, 59).StyleNew.BackColor = Color.White;
                        //1572
                        c1FlexGrid4.Rows[e.Row][58] = 1;
                    }
                }
                //No.1622
                //分割方法
                if (e.Col == c1FlexGrid4.Cols["BunkatsuHouhou"].Index)
                {
                    if (c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() == "2" && c1FlexGrid4.Rows[e.Row]["BunkatsuHouhou"].ToString() == "1")
                    {
                        c1FlexGrid4.GetCellRange(e.Row, 59).StyleNew.BackColor = Color.FromArgb(240, 240, 240);
                        c1FlexGrid4.Rows[e.Row][59] = "";
                    }
                    if (c1FlexGrid4.Rows[e.Row]["BunkatsuHouhou"].ToString() == "2")
                    {
                        c1FlexGrid4.GetCellRange(e.Row, 59).StyleNew.BackColor = Color.White;
                    }
                }
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
                        if(copyData == null || copyData.Count <= 0)
                        {
                            //isWinCopy = true;
                            IDataObject data = Clipboard.GetDataObject();
                            if (data.GetDataPresent(DataFormats.Text))
                            {
                                string str;
                                //クリップボードからデータを取得
                                str = (string)data.GetData(DataFormats.Text);
                                //クリップボードにある最後の開業コードを削除
                                string strWinCopyText = str.TrimEnd('\r', '\n');

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
                        }
                        if (copyData == null || copyData.Count <= 0)
                        {
                            return;
                        }
                        //bool isWinCopy = false;
                        //IDataObject data = Clipboard.GetDataObject();
                        //string strWinCopyText = "";
                        //if (data.GetDataPresent(DataFormats.Text))
                        //{
                        //    string str;
                        //    //クリップボードからデータを取得
                        //    str = (string)data.GetData(DataFormats.Text);
                        //    //クリップボードにある最後の開業コードを削除
                        //    strWinCopyText = str.TrimEnd('\r', '\n');

                        //    string strGridCopyText = "";
                        //    if (copyData != null && copyData.Count > 0)
                        //    {
                        //        for (int i = 0; i < copyData.Count; i++)
                        //        {
                        //            string sLineText = "";
                        //            for (int j = 0; j < copyData[i].Count; j++)
                        //            {
                        //                sLineText = sLineText + copyData[i][j];
                        //            }
                        //            strGridCopyText = strGridCopyText + sLineText;
                        //        }
                        //        strGridCopyText = strGridCopyText.Replace(Environment.NewLine, "").Replace("\n", "").Replace("\r", "").Replace("\t","").Replace(" ", "");
                        //    }
                        //    string strWin = strWinCopyText.Replace(Environment.NewLine, "").Replace("\n", "").Replace("\r", "").Replace("\t", "").Replace(" ","");
                        //    if (!strWin.Equals(strGridCopyText))
                        //        isWinCopy = true;
                        //}
                        //if (isWinCopy)
                        //{
                        //    if (copyData == null)
                        //    {
                        //        copyData = new List<List<string>>();
                        //    }
                        //    copyData.Clear();
                        //    string[] lines = strWinCopyText.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                        //    for(int i = 0; i < lines.Length; i++)
                        //    {
                        //        copyData.Add(new List<string>());
                        //        string[] cols = lines[i].Split('\t');
                        //        for (int j = 0; j < cols.Length; j++)
                        //        {
                        //            copyData[i].Add(cols[j]);
                        //        }
                        //    }
                        //}
                        //if (copyData == null)
                        //{
                        //    return;
                        //}
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
                            if (num >= copyData.Count) {
                                if(isRowBreak) num = 0;
                                else break;
                            }
                            // ▼列
                            // コピー列の倍数を選択する場合
                            int iCol = 0;
                            for (int j = col; j<= colSel && j < c1FlexGrid4.Cols.Count; j++)
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

        private void src_Busho_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        // 調査品目明細タブ
        private void c1FlexGrid4_KeyDownEdit(object sender, C1.Win.C1FlexGrid.KeyEditEventArgs e)
        {
            // 調査品目編集モード 0:表示 1:編集
            if (ChousaHinmokuMode == 1) {
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
                    //1584
                    || e.Col == c1FlexGrid4.Cols["TaniAtariTankaSuryo"].Index
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

        // 11桁になるようゼロパディングした文字を返却する
        private string zeroPadding(string str)
        {
            string moji = "";

            //// 文字列が3文字以上、かつ「.」が含まれている場合
            //if (str.Length > 3 && str.IndexOf(".") > -1)
            //{
            //    // C1FlexGrid上の全体順、個別順は、編集前は小数点4桁（.0000）持ちだが、編集後小数点2桁（.00）になってしまう為、
            //    // 桁を合わせる為に、"00"を付け足す
            //    if (".00".Equals(str.Substring(str.Length - 3)))
            //    {
            //        str = str + "00";
            //    }
            //}
            //moji = zeroStr + str;
            //moji = moji.Substring(moji.Length - 11);

            double.TryParse(str, out double num);
            moji = string.Format("{0:000000.0000}", num);
            return moji;
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
                    // 単価適用地域
                    //else if (j == 17)
                    else if (j == c1FlexGrid4.Cols["ChousaTankaTekiyouTiku"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = w_TankaTekiyouChiiki;
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
                    //奉行エクセル
                    // 集計表Ver
                    else if (j == c1FlexGrid4.Cols["ShukeihyoVer"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = "1";
                    }
                    //分割方法（ファイル・シート）
                    else if (j == c1FlexGrid4.Cols["BunkatsuHouhou"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = "-";
                    }
                    //グループ名
                    else if (j == c1FlexGrid4.Cols["GroupMei"].Index)
                    {
                        c1FlexGrid4.Rows[rowCount][j] = "";
                    }
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

        //奉行エクセル　
        // 品目からの取込
        private void button3_ReadHinmoku_Click(object sender, EventArgs e)
        {
            Popup_HinmokuTorikomi form = new Popup_HinmokuTorikomi();
            try 
            { 
                form.ShowDialog();
            }
            catch (AccessViolationException ave)
            {
                // AutoComplete のバグで、メモリ破損エラーが出る
                // AccessViolationException は catch出来ない
                // メモリアクセス違反が発生してしまったものを復旧しつつ挙動させることは極めて困難な為、
                // catchが出来ないようになっている
                Console.WriteLine("error:" + ave.ToString());
            }

            // 結果が返ってきていたら
            if (form.ReturnValue != null && form.ReturnValue[0] != null && form.ReturnValue[0] == "1")
            {
                string returnMadoguchiID = form.ReturnValue[1];

                if (returnMadoguchiID == "")
                {
                    // MadoguchiIDが空だったら処理終了
                    return;
                }

                //レイアウトロジックを停止する
                this.SuspendLayout();

                c1FlexGrid4.Visible = false;

                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                try
                {
                    using (var conn = new SqlConnection(connStr))
                    {

                        var cmd = conn.CreateCommand();

                        DataTable addCousaHinmokuDT = new DataTable();

                        //調査品目の取得
                        cmd.CommandText = "SELECT " +
                                   " 0 " + // 0:未使用
                                   ",0 " + // 1:未使用
                                   ",0 " + // 2:未使用
                                   ",0 " + // 3:未使用
                                   ",0 " + // 4:未使用
                                    ",CASE " +
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
                                   " , HinmokuChousainCD " +
                                   " , HinmokuRyakuBushoFuku1CD " + // 40
                                   " , HinmokuFukuChousainCD1 " +
                                   " , HinmokuRyakuBushoFuku2CD " +
                                   " , HinmokuFukuChousainCD2 " +
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
                                   " , ChousaDeleteFlag " + // 50
                                   " , ChousaHinmokuID " + // 51:調査品目ID
                                   " , ChousaShinchokuJoukyou " + // 52:進捗状況
                                   //奉行エクセル　
                                   ", ChousaShuukeihyouVer" +
                                   ", ChousaBunkatsuHouhou" +
                                   ", ChousaKoujiKouzoubutsumei" +
                                   ", ChousaHachushaTeikyouTani" +
                                   ", ChousaTaniAtariKakaku"+
                                   ", chousaTaniAtariSuuryou" +
                                   ", ChousaTaniAtariTanka" +
                                   ", ChousaNiwatashiJouken" +
                                   ", MadoguchiGroupMei" +
                                   " , 0 " + // 53:0:Insert/1:Select/2:Update
                                   //" , ChousaTankaCD1 " + // 54:発注品目コード
                                   " , '' " + // 55:並び順 →  54:並び順
                                   "FROM " +
                                   " ChousaHinmoku  " +
                                   //奉行エクセル
                                   "LEFT JOIN MadoguchiGroupMaster ON ChousaHinmoku.MadoguchiID = MadoguchiGroupMaster.MadoguchiGroupMasterID" +
                                   "LEFT JOIN MadoguchiJouhou ON MadoguchiJouhou.MadoguchiID = ChousaHinmoku.MadoguchiID " +
                                   "LEFT JOIN Mst_Chousain MC0 ON HinmokuChousainCD = MC0.KojinCD " +
                                   "LEFT JOIN Mst_Chousain MC1 ON HinmokuFukuChousainCD1 = MC1.KojinCD " +
                                   "LEFT JOIN Mst_Chousain MC2 ON HinmokuFukuChousainCD2 = MC2.KojinCD " +
                                   "WHERE " +
                                   "MadoguchiJouhou.MadoguchiID = '" + returnMadoguchiID + "' AND ChousaDeleteFlag <> 1 AND ChousaHinmokuID > 0 " +
                                   "ORDER BY ChousaZentaiJun,ChousaKobetsuJun,ChousaHinmokuID, MadoguchiGroupMaster.MadoguchiGroupMei ";

                        var sda = new SqlDataAdapter(cmd);
                        addCousaHinmokuDT.Clear();
                        sda.Fill(addCousaHinmokuDT);

                        double num = 0;
                        double w_ZentaiJun = 0;
                        double w_KobetsuJun = 0;
                        int rowCount = 0;

                        // 取得できた分を追加する
                        for (int i = 0; i < addCousaHinmokuDT.Rows.Count; i++)
                        {
                            // 取得した全体順
                            //double.TryParse(addCousaHinmokuDT.Rows[i][6].ToString(), out w_ZentaiJun);
                            double.TryParse(addCousaHinmokuDT.Rows[i]["ChousaZentaiJun"].ToString(), out w_ZentaiJun);

                            // Gridの全体順の中での最大の個別順を取得する
                            w_KobetsuJun = 0;
                            num = 0;
                            for (int j = 2; j < c1FlexGrid4.Rows.Count; j++)
                            {
                                //if (c1FlexGrid4.Rows[j][6] != null && w_ZentaiJun == double.Parse(c1FlexGrid4.Rows[j][6].ToString()))
                                if (c1FlexGrid4.Rows[j]["ChousaZentaiJun"] != null && w_ZentaiJun == double.Parse(c1FlexGrid4.Rows[j]["ChousaZentaiJun"].ToString()))
                                {
                                    // 個別順の最大値を取り出す
                                    //if (c1FlexGrid4.Rows[j][7] != null && double.TryParse(c1FlexGrid4.Rows[j][7].ToString(), out num))
                                    if (c1FlexGrid4.Rows[j]["ChousaKobetsuJun"] != null && double.TryParse(c1FlexGrid4.Rows[j]["ChousaKobetsuJun"].ToString(), out num))
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

                            // Gridの行数を取得
                            rowCount = c1FlexGrid4.Rows.Count;

                            // 最終行に行追加
                            c1FlexGrid4.Rows.Insert(rowCount);
                            // グリッドに値をセットする
                            // Grid 5:進捗状況からセット Grid:58列 Data:55列
                            //for (int j = 5; j < c1FlexGrid4.Cols.Count; j++)
                            //{
                            //    // ChousaHinmokuID
                            //    if (j == 55)
                            //    {
                            //        // 追加時にIDを振っておく
                            //        c1FlexGrid4.Rows[rowCount][j] = GlobalMethod.getSaiban("HinmokuMeisaiID");
                            //    }
                            //    // 全体順
                            //    else if (j == 6)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = addCousaHinmokuDT.Rows[i][j].ToString();
                            //    }
                            //    // 個別順
                            //    else if (j == 7)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = w_KobetsuJun;
                            //    }
                            //    // 価格
                            //    else if (j == 13)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "";
                            //    }
                            //    // 見積先
                            //    else if (j == 20)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "";
                            //    }
                            //    // ベースメーカー
                            //    else if (j == 21)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "";
                            //    }
                            //    // ベース単価
                            //    else if (j == 22)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = 0;
                            //    }
                            //    // 掛率
                            //    else if (j == 23)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = 0;
                            //    }
                            //    // 前回価格
                            //    else if (j == 26)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = 0;
                            //    }
                            //    // 品目情報1
                            //    else if (j == 28)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "";
                            //    }
                            //    // 品目情報2
                            //    else if (j == 29)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "";
                            //    }
                            //    // リンク先
                            //    else if (j == 40)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "0";
                            //    }
                            //    // リンク先パス
                            //    else if (j == 41)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "";
                            //    }
                            //    // 調査担当者
                            //    else if (j == 43)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "";
                            //    }
                            //    // 報告数
                            //    else if (j == 48)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = 0;
                            //    }
                            //    // 依頼数
                            //    else if (j == 50)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = 0;
                            //    }
                            //    // 締切日
                            //    else if (j == 52)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = item1_MadoguchiShimekiribi.Text;
                            //    }
                            //    // 報告済
                            //    else if (j == 53)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "0";
                            //    }
                            //    // ChousaShinchokuJoukyou
                            //    else if (j == 56)
                            //    {
                            //        // 進捗状況は、20:調査開始
                            //        c1FlexGrid4.Rows[rowCount][j] = 20;
                            //    }
                            //    // 0:Insert/1:Select/2:Update
                            //    else if (j == 57)
                            //    {
                            //        c1FlexGrid4.Rows[rowCount][j] = "0";
                            //    }
                            //    // ソートキー
                            //    else if (j == 58)
                            //    {
                            //        // 並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                            //        c1FlexGrid4.Rows[rowCount][j] = "N" + zeroPadding(c1FlexGrid4.Rows[rowCount][6].ToString()) + "-" + zeroPadding(c1FlexGrid4.Rows[rowCount][7].ToString());
                            //    }
                            //    else
                            //    {
                            //        if(j < 51) 
                            //        { 
                            //            c1FlexGrid4.Rows[rowCount][j] = addCousaHinmokuDT.Rows[i][j].ToString();
                            //        }
                            //    }
                            //}

                            // 進捗状況
                            c1FlexGrid4.Rows[rowCount]["ShinchokuIcon"] = addCousaHinmokuDT.Rows[i]["Shinchock"];
                            // 全体順
                            c1FlexGrid4.Rows[rowCount]["ChousaZentaiJun"] = addCousaHinmokuDT.Rows[i]["ChousaZentaiJun"].ToString();
                            // 個別順
                            c1FlexGrid4.Rows[rowCount]["ChousaKobetsuJun"] = w_KobetsuJun;
                            // 材工
                            c1FlexGrid4.Rows[rowCount]["ChousaZaiKou"] = addCousaHinmokuDT.Rows[i]["ChousaZaiKou"];
                            // 品目
                            c1FlexGrid4.Rows[rowCount]["ChousaHinmei"] = addCousaHinmokuDT.Rows[i]["ChousaHinmei"];
                            // 規格
                            c1FlexGrid4.Rows[rowCount]["ChousaKikaku"] = addCousaHinmokuDT.Rows[i]["ChousaKikaku"];
                            // 単位
                            c1FlexGrid4.Rows[rowCount]["ChousaTanka"] = addCousaHinmokuDT.Rows[i]["ChousaTanka"];
                            // 参考質量
                            c1FlexGrid4.Rows[rowCount]["ChousaSankouShitsuryou"] = addCousaHinmokuDT.Rows[i]["ChousaSankouShitsuryou"];
                            // 価格
                            c1FlexGrid4.Rows[rowCount]["ChousaKakaku"] = "";
                            // 中止
                            c1FlexGrid4.Rows[rowCount]["ChousaChuushi"] = C1.Win.C1FlexGrid.CheckEnum.Unchecked;
                            if (addCousaHinmokuDT.Rows[i]["ChousaChuushi"].ToString() == "1")
                            {
                                c1FlexGrid4.Rows[rowCount]["ChousaChuushi"] = C1.Win.C1FlexGrid.CheckEnum.Checked;
                            }
                            // 報告備考
                            c1FlexGrid4.Rows[rowCount]["ChousaBikou2"] = addCousaHinmokuDT.Rows[i]["ChousaBikou2"];
                            // 依頼備考
                            c1FlexGrid4.Rows[rowCount]["ChousaBikou"] = addCousaHinmokuDT.Rows[i]["ChousaBikou"];
                            // 単価適用地域
                            c1FlexGrid4.Rows[rowCount]["ChousaTankaTekiyouTiku"] = addCousaHinmokuDT.Rows[i]["ChousaTankaTekiyouTiku"];
                            // 図面番号
                            c1FlexGrid4.Rows[rowCount]["ChousaZumenNo"] = addCousaHinmokuDT.Rows[i]["ChousaZumenNo"];
                            // 数量
                            c1FlexGrid4.Rows[rowCount]["ChousaSuuryou"] = addCousaHinmokuDT.Rows[i]["ChousaSuuryou"];
                            // 見積先
                            c1FlexGrid4.Rows[rowCount]["ChousaMitsumorisaki"] = "";
                            // ベースメーカー
                            c1FlexGrid4.Rows[rowCount]["ChousaBaseMakere"] = "";
                            // ベース単位
                            c1FlexGrid4.Rows[rowCount]["ChousaBaseTanka"] = 0;
                            // 掛率
                            c1FlexGrid4.Rows[rowCount]["ChousaKakeritsu"] = 0;
                            // 属性
                            c1FlexGrid4.Rows[rowCount]["ChousaObiMei"] = addCousaHinmokuDT.Rows[i]["ChousaObiMei"];
                            // 前回単位
                            c1FlexGrid4.Rows[rowCount]["ChousaZenkaiTani"] = addCousaHinmokuDT.Rows[i]["ChousaZenkaiTani"];
                            // 前回価格
                            c1FlexGrid4.Rows[rowCount]["ChousaZenkaiKakaku"] = 0;
                            // 発注者提供単価
                            c1FlexGrid4.Rows[rowCount]["ChousaSankouti"] = addCousaHinmokuDT.Rows[i]["ChousaSankouti"];
                            // 品目情報1
                            c1FlexGrid4.Rows[rowCount]["ChousaHinmokuJouhou1"] = "";
                            // 品目情報2
                            c1FlexGrid4.Rows[rowCount]["ChousaHinmokuJouhou2"] = "";
                            // 前回質量
                            c1FlexGrid4.Rows[rowCount]["ChousaFukuShizai"] = addCousaHinmokuDT.Rows[i]["ChousaFukuShizai "];
                            // メモ1
                            c1FlexGrid4.Rows[rowCount]["ChousaBunrui"] = addCousaHinmokuDT.Rows[i]["ChousaBunrui"];
                            // メモ2
                            c1FlexGrid4.Rows[rowCount]["ChousaMemo2"] = addCousaHinmokuDT.Rows[i]["ChousaMemo2"];
                            // 発注品目コード
                            c1FlexGrid4.Rows[rowCount]["ChousaTankaCD1"] = addCousaHinmokuDT.Rows[i]["ChousaTankaCD1"];
                            // 地区割コード
                            c1FlexGrid4.Rows[rowCount]["ChousaTikuWariCode"] = addCousaHinmokuDT.Rows[i]["ChousaTikuWariCode"];
                            // 地区コード
                            c1FlexGrid4.Rows[rowCount]["ChousaTikuCode"] = addCousaHinmokuDT.Rows[i]["ChousaTikuCode"];
                            // 地区名
                            c1FlexGrid4.Rows[rowCount]["ChousaTikuMei"] = addCousaHinmokuDT.Rows[i]["ChousaTikuMei"];
                            // 少額案件[10万/100万]
                            c1FlexGrid4.Rows[rowCount]["ChousaShougaku"] = C1.Win.C1FlexGrid.CheckEnum.Unchecked;
                            if (addCousaHinmokuDT.Rows[i]["ChousaShougaku"].ToString() == "1")
                            {
                                c1FlexGrid4.Rows[rowCount]["ChousaShougaku"] = C1.Win.C1FlexGrid.CheckEnum.Checked;
                            }
                            // Web建
                            c1FlexGrid4.Rows[rowCount]["ChousaWebKen"] = C1.Win.C1FlexGrid.CheckEnum.Unchecked;
                            if (addCousaHinmokuDT.Rows[i]["ChousaWebKen"].ToString() == "1")
                            {
                                c1FlexGrid4.Rows[rowCount]["ChousaWebKen"] = C1.Win.C1FlexGrid.CheckEnum.Checked;
                            }
                            // 根拠関連コード
                            c1FlexGrid4.Rows[rowCount]["ChousaKonkyoCode"] = addCousaHinmokuDT.Rows[i]["ChousaKonkyoCode"];
                            // リンク先
                            c1FlexGrid4.Rows[rowCount]["ChousaLinkSakli"] = "0";
                            // リンク先パス
                            c1FlexGrid4.Rows[rowCount]["ChousaLinkSakliFolder"] = "";

                            //奉行エクセル
                            //作業フォルダ
                            c1FlexGrid4.Rows[rowCount]["SagyoForuda"] = "0";
                            //作業フォルダパス
                            c1FlexGrid4.Rows[rowCount]["SagyoForudaPasu"] = "";
                            // 調査担当部所
                            c1FlexGrid4.Rows[rowCount]["HinmokuRyakuBushoCD"] = addCousaHinmokuDT.Rows[i]["HinmokuRyakuBushoCD"];
                            // 調査担当者
                            c1FlexGrid4.Rows[rowCount]["HinmokuChousainCD"] = "";
                            // 副調査担当部所1
                            c1FlexGrid4.Rows[rowCount]["HinmokuRyakuBushoFuku1CD"] = addCousaHinmokuDT.Rows[i]["HinmokuRyakuBushoFuku1CD"];
                            // 副調査担当者1
                            c1FlexGrid4.Rows[rowCount]["HinmokuFukuChousainCD1"] = addCousaHinmokuDT.Rows[i]["HinmokuFukuChousainCD1"];
                            // 副調査担当部所2
                            c1FlexGrid4.Rows[rowCount]["HinmokuRyakuBushoFuku2CD"] = addCousaHinmokuDT.Rows[i]["HinmokuRyakuBushoFuku2CD"];
                            // 副調査担当者2
                            c1FlexGrid4.Rows[rowCount]["HinmokuFukuChousainCD2"] = addCousaHinmokuDT.Rows[i]["HinmokuFukuChousainCD2"];
                            // 報告数
                            c1FlexGrid4.Rows[rowCount]["ChousaHoukokuHonsuu"] = 0;
                            // 報告ランク
                            c1FlexGrid4.Rows[rowCount]["ChousaHoukokuRank"] = addCousaHinmokuDT.Rows[i]["ChousaHoukokuRank"];
                            // 依頼数
                            c1FlexGrid4.Rows[rowCount]["ChousaIraiHonsuu"] = 0;
                            // 依頼ランク
                            c1FlexGrid4.Rows[rowCount]["ChousaIraiRank"] = addCousaHinmokuDT.Rows[i]["ChousaIraiRank"];
                            // 締切日
                            c1FlexGrid4.Rows[rowCount]["ChousaHinmokuShimekiribi"] = item1_MadoguchiShimekiribi.Text;
                            // 報告済
                            c1FlexGrid4.Rows[rowCount]["ChousaHoukokuzumi"] = "0";
                            // 削除フラグ
                            c1FlexGrid4.Rows[rowCount]["ChousaDeleteFlag"] = addCousaHinmokuDT.Rows[i]["ChousaDeleteFlag"];
                            // 調査品目ID
                            c1FlexGrid4.Rows[rowCount]["ChousaHinmokuID2"] = GlobalMethod.getSaiban("HinmokuMeisaiID");
                            // 進捗状況
                            c1FlexGrid4.Rows[rowCount]["ChousaShinchokuJoukyou"] = 20;
                            // 0:Insert/1:Select/2:Update
                            c1FlexGrid4.Rows[rowCount]["Mode"] = "0";
                            // 並び順
                            // エラー行を先頭にするため、並び順（全体順 - 個別順）の頭に エラーなら E、正常なら Nを付け、ソートしやすくする
                            c1FlexGrid4.Rows[rowCount]["ColumnSort"] = "N"
                                                                     + zeroPadding(c1FlexGrid4.Rows[rowCount]["ChousaZentaiJun"].ToString())
                                                                     + "-"
                                                                     + zeroPadding(c1FlexGrid4.Rows[rowCount]["ChousaKobetsuJun"].ToString());

                            // 奉行エクセル　
                            //集計表Ver
                            c1FlexGrid4.Rows[rowCount]["ShukeihyoVer"] = addCousaHinmokuDT.Rows[i]["ChousaShuukeihyouVer"];
                            //分割方法
                            c1FlexGrid4.Rows[rowCount]["BunkatsuHouhou"] = addCousaHinmokuDT.Rows[i]["ChousaBunkatsuHouhou"];
                            //工事・構造物名
                            c1FlexGrid4.Rows[rowCount]["KojiKoubutsuMei"] = addCousaHinmokuDT.Rows[i]["ChousaKoujiKouzoubutsumei"];
                            //単位当たり単価（単位）
                            c1FlexGrid4.Rows[rowCount]["TaniAtariTankaTani"] = addCousaHinmokuDT.Rows[i]["ChousaTaniAtariTanka"];
                            //単位当たり単価（数量）
                            c1FlexGrid4.Rows[rowCount]["TaniAtariTankaSuryo"] = addCousaHinmokuDT.Rows[i]["chousaTaniAtariSuuryou"];
                            //単位当たり単価（価格）
                            c1FlexGrid4.Rows[rowCount]["TaniAtariTankaKakaku"] = addCousaHinmokuDT.Rows[i]["ChousaTaniAtariKakaku"];
                           //荷渡し条件
                            c1FlexGrid4.Rows[rowCount]["NiwatashiJoken"] = addCousaHinmokuDT.Rows[i]["ChousaNiwatashiJouken"];
                            //発注者提供単位
                            c1FlexGrid4.Rows[rowCount]["HachushaTeikyoTani"] = addCousaHinmokuDT.Rows[i]["ChousaHachushaTeikyouTani"];
                            //グループ名
                            c1FlexGrid4.Rows[rowCount]["GroupMei"] = addCousaHinmokuDT.Rows[i]["ChousaMadoguchiGroupMasterID"];
                        }
                    }
                }
                catch (Exception)
                {
                    throw;
                }

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
                // ページ数再計算
                Paging_all.Text = (Math.Ceiling(((double)c1FlexGrid4.Rows.Count - 2) / hyoujisuu)).ToString();
                // 表示件数（ヘッダー2行分を引く）
                Grid_Num.Text = "(" + (c1FlexGrid4.Rows.Count - 2) + ")";
                Grid_Visible(int.Parse(Paging_now.Text),"1");

                c1FlexGrid4.Visible = true;

                //レイアウトロジックを再開する
                this.ResumeLayout();

            }

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

                    if (Double.TryParse(c1FlexGrid4.Editor.Text,out inputNum))
                    {
                        if(maxNum < inputNum || 0 > inputNum)
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
            //if (e.Col == 48 || e.Col == 50)
            if (e.Col == c1FlexGrid4.Cols["ChousaHoukokuHonsuu"].Index || e.Col == c1FlexGrid4.Cols["ChousaIraiHonsuu"].Index)
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

        // ランク内訳明細の出力
        private void button_RankMeisai_Click(object sender, EventArgs e)
        {
            set_error("", 0);

            //単価ランクの編集状態初期化（報告ランク）
            SwichButton_Rank();

            // 依頼ランクの集計結果で登録されないように再計算させる
            ReCal();

            // 帳票出力前に画面の内容で更新を行う
            UpdateMadoguchi(6);

            // 1:MadoguchiID     窓口ID
            string[] report_data = new string[1] { "" };
            report_data[0] = MadoguchiID;
            int listID = 10;    // ランク内訳明細

            // ランク内訳明細の出力
            string[] result = GlobalMethod.InsertMadoguchiReportWork(listID, UserInfos[0], report_data, "RankUchiwakeMeisai");

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

        // 文字変換
        private void button3_ChangeMoji_Click(object sender, EventArgs e)
        {
            set_error("", 0);

            string methodName = ".button3_ChangeMoji_Click";

            Popup_MojiChange form = new Popup_MojiChange();
            form.ShowDialog();

            // ▼ReturnValue
            // 0:実行結果 1:実行 それ以外実行無し
            // 1:変換対象 0:全角→半角 1:半角→全角
            // 2:～変換   1:変換 それ以外変換無し
            // 3:英数字   1:変換 それ以外変換無し
            // 4:品名     1:対象 それ以外対象外
            // 5:規格     1:対象 それ以外対象外
            // 6:報告備考 1:対象 それ以外対象外
            // 7:依頼備考 1:対象 それ以外対象外
            //不具合No1010（744）
            // 8:カタカナ変換　1:変換 それ以外変換無し
            if (form.ReturnValue != null && form.ReturnValue[0] == "1")
            {
                //不具合No1010（744）
                // ～変換、英数字にチェックが入ってなければ処理終了
                if (form.ReturnValue[2] != "1" && form.ReturnValue[3] != "1" && form.ReturnValue[8] != "1")
                {
                    return;
                }
                ////// ～変換、英数字にチェックが入ってなければ処理終了
                ////if (form.ReturnValue[2] != "1" && form.ReturnValue[3] != "1")
                ////{
                ////    return;
                ////}
                ///
                // 変換対象項目にチェックが入ってなければ処理終了
                if (form.ReturnValue[4] != "1" && form.ReturnValue[5] != "1" && form.ReturnValue[6] != "1" && form.ReturnValue[7] != "1")
                {
                    return;
                }

                //// 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
                //if(chousaHinmokuDispFlg == "0")
                //{
                //    chousaHinmokuDispFlg = "1";
                //    set_data(3);
                //}
                
                var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var cmd = conn.CreateCommand();
                    SqlTransaction transaction = conn.BeginTransaction();
                    cmd.Transaction = transaction;
                    try
                    {

                        DataTable dt = new DataTable();
                        string where = "";

                        cmd.CommandText = "SELECT " +
                                          "MojiChangeZen,MojiChangeHan " +
                                          "FROM M_MojiChange ";

                        // 英数字
                        if (form.ReturnValue[3] == "1")
                        {
                            where = "(MojiChangeYukouFlg = 1 AND (MojiChangeAlphaFlg = 1 OR MojiChangeKiguFlg = 1 OR MojiChangeTokusyuFlg = 1)) ";
                        }
                        // ～変換
                        if (form.ReturnValue[2] == "1")
                        {
                            if(where != "")
                            {
                                where += "OR ";
                            }
                            where += "(MojiChangeYukouFlg = 1 AND MojiChangeKiguFlg = 1 AND MojiChangeHan = '~' ) ";
                        }
                        //不具合No1010（744）
                        // カタカナ変換
                        if (form.ReturnValue[8] == "1")
                        {
                            if (where != "")
                            {
                                where += "OR ";
                            }
                            where += "(MojiChangeYukouFlg = 1 AND MojiChangeKanaFlg = 1 ) ";
                        }

                        cmd.CommandText += "WHERE " + where;
                        var sda = new SqlDataAdapter(cmd);
                        sda.Fill(dt);

                        int targetIndex = 0;
                        int changeIndex = 0;

                        // 1:変換対象 0:全角→半角 1:半角→全角
                        if (form.ReturnValue[1] == "0")
                        {
                            targetIndex = 0;
                            changeIndex = 1;
                        }
                        else
                        {
                            targetIndex = 1;
                            changeIndex = 0;
                        }

                        //int hinmokuIndex = 9;
                        //int kikakukuIndex = 10;
                        //int houkokuBikoIndex = 15;
                        //int iraiBikoIndex = 16;
                        //int updateIndex = 57;
                        int hinmokuIndex = c1FlexGrid4.Cols["ChousaHinmei"].Index;      // 品目
                        int kikakukuIndex = c1FlexGrid4.Cols["ChousaKikaku"].Index;     // 規格
                        int houkokuBikoIndex = c1FlexGrid4.Cols["ChousaBikou2"].Index;  // 報告備考
                        int iraiBikoIndex = c1FlexGrid4.Cols["ChousaBikou"].Index;      // 依頼備考
                        int updateIndex = c1FlexGrid4.Cols["Mode"].Index;               // 0:Insert/1:Select/2:Update
                        int index = 0;

                        // ▼調査品目明細Gridの対象
                        // 09:品目
                        // 10:規格
                        // 15:報告備考
                        // 16:依頼備考

                        //// 品名
                        //if (form.ReturnValue[4] == "1")
                        //{
                        //    index = hinmokuIndex;
                        //    changeMoji(dt, index, targetIndex, changeIndex, updateIndex);
                        //}
                        //// 規格
                        //if (form.ReturnValue[5] == "1")
                        //{
                        //    index = kikakukuIndex;
                        //    changeMoji(dt, index, targetIndex, changeIndex, updateIndex);
                        //}
                        //// 報告備考
                        //if (form.ReturnValue[6] == "1")
                        //{
                        //    index = houkokuBikoIndex;
                        //    changeMoji(dt, index, targetIndex, changeIndex, updateIndex);
                        //}
                        //// 依頼備考
                        //if (form.ReturnValue[7] == "1")
                        //{
                        //    index = iraiBikoIndex;
                        //    changeMoji(dt, index, targetIndex, changeIndex, updateIndex);
                        //}

                        //// 調査品目更新処理の呼び出し
                        //chousaHinmokuUpdate();


                        // M_MojiChangeで条件に合うレコード分、updateを実行
                        //string query = "";
                        //// M_MojiChangeを回す
                        //for (int j = 0; j < dt.Rows.Count; j++)
                        //{
                        //    cmd.CommandText = "update ChousaHinmoku set ";
                        //    query = "";

                        //    // 品名
                        //    if (form.ReturnValue[4] == "1")
                        //    {
                        //        query += "ChousaHinmei = replace(ChousaHinmei,'" + GlobalMethod.ChangeSqlText(dt.Rows[j][targetIndex].ToString(), 0, 0) + "','" + GlobalMethod.ChangeSqlText(dt.Rows[j][changeIndex].ToString(), 0, 0) + "') ";
                        //    }
                        //    // 規格
                        //    if (form.ReturnValue[5] == "1")
                        //    {
                        //        if (query != "")
                        //        {
                        //            query += ",";
                        //        }
                        //        query += "ChousaKikaku = replace(ChousaKikaku,'" + GlobalMethod.ChangeSqlText(dt.Rows[j][targetIndex].ToString(), 0, 0) + "','" + GlobalMethod.ChangeSqlText(dt.Rows[j][changeIndex].ToString(), 0, 0) + "') ";
                        //    }
                        //    // 報告備考
                        //    if (form.ReturnValue[6] == "1")
                        //    {
                        //        if (query != "")
                        //        {
                        //            query += ",";
                        //        }
                        //        query += "ChousaBikou2 = replace(ChousaBikou2,'" + GlobalMethod.ChangeSqlText(dt.Rows[j][targetIndex].ToString(), 0, 0) + "','" + GlobalMethod.ChangeSqlText(dt.Rows[j][changeIndex].ToString(), 0, 0) + "') ";
                        //    }
                        //    // 依頼備考
                        //    if (form.ReturnValue[7] == "1")
                        //    {
                        //        if (query != "")
                        //        {
                        //            query += ",";
                        //        }
                        //        query += "ChousaBikou = replace(ChousaBikou,'" + GlobalMethod.ChangeSqlText(dt.Rows[j][targetIndex].ToString(), 0, 0) + "','" + GlobalMethod.ChangeSqlText(dt.Rows[j][changeIndex].ToString(), 0, 0) + "') ";
                        //    }
                        //    cmd.CommandText += query;
                        //    cmd.CommandText += ",ChousaUpdateDate = '" + DateTime.Now.ToString() + "'" +       // 更新日時
                        //                       ",ChousaUpdateUser = '" + UserInfos[0] + "' " +                 // 更新ユーザー
                        //                       ",ChousaUpdateProgram = '" + pgmName + methodName + "' ";  // 更新プログラム

                        //    // MadoguchiIDと検索した際の条件で更新を行う
                        //    cmd.CommandText += "where MadoguchiID = " + MadoguchiID + " ";
                        //    if(chousaHinmokuSearchWhere != "")
                        //    {
                        //        cmd.CommandText += "AND " + chousaHinmokuSearchWhere;
                        //    }
                        //    cmd.ExecuteNonQuery();
                        //}
                        //transaction.Commit();

                        // M_MojiChangeにあるレコードでupdateを実行（1回のみ）
                        string query = "";

                        string hinmeistr = "";
                        string kikakustr = "";
                        string bikou2str = "";
                        string bikoustr = "";

                        cmd.CommandText = "update ChousaHinmoku set ";

                        // M_MojiChangeを回す
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            // 品名
                            if (form.ReturnValue[4] == "1")
                            {
                                if (hinmeistr == "")
                                {
                                    hinmeistr = "ChousaHinmei COLLATE Japanese_CS_AS_KS_WS";
                                }
                                hinmeistr = "replace(" + hinmeistr + ",N'" + GlobalMethod.ChangeSqlText(dt.Rows[j][targetIndex].ToString(), 0, 0) + "',N'" + GlobalMethod.ChangeSqlText(dt.Rows[j][changeIndex].ToString(), 0, 0) + "') ";
                            }
                            // 規格
                            if (form.ReturnValue[5] == "1")
                            {
                                if (kikakustr == "")
                                {
                                    kikakustr = "ChousaKikaku COLLATE Japanese_CS_AS_KS_WS";
                                }
                                kikakustr = "replace(" + kikakustr + ",N'" + GlobalMethod.ChangeSqlText(dt.Rows[j][targetIndex].ToString(), 0, 0) + "',N'" + GlobalMethod.ChangeSqlText(dt.Rows[j][changeIndex].ToString(), 0, 0) + "') ";
                            }
                            // 報告備考
                            if (form.ReturnValue[6] == "1")
                            {
                                if (bikou2str == "")
                                {
                                    bikou2str = "ChousaBikou2 COLLATE Japanese_CS_AS_KS_WS";
                                }
                                bikou2str = "replace(" + bikou2str + ",N'" + GlobalMethod.ChangeSqlText(dt.Rows[j][targetIndex].ToString(), 0, 0) + "',N'" + GlobalMethod.ChangeSqlText(dt.Rows[j][changeIndex].ToString(), 0, 0) + "') ";
                            }
                            // 依頼備考
                            if (form.ReturnValue[7] == "1")
                            {
                                if (bikoustr == "")
                                {
                                    bikoustr = "ChousaBikou COLLATE Japanese_CS_AS_KS_WS";
                                }
                                bikoustr = "replace(" + bikoustr + ",N'" + GlobalMethod.ChangeSqlText(dt.Rows[j][targetIndex].ToString(), 0, 0) + "',N'" + GlobalMethod.ChangeSqlText(dt.Rows[j][changeIndex].ToString(), 0, 0) + "') ";
                            }
                        }
                        if (hinmeistr != "")
                        {
                            query += "ChousaHinmei = " + hinmeistr;
                        }
                        if (kikakustr != "")
                        {
                            if (query != "")
                            {
                                query += ",";
                            }
                            query += "ChousaKikaku = " + kikakustr;
                        }
                        if (bikou2str != "")
                        {
                            if (query != "")
                            {
                                query += ",";
                            }
                            query += "ChousaBikou2 = " + bikou2str;
                        }
                        if (bikoustr != "")
                        {
                            if (query != "")
                            {
                                query += ",";
                            }
                            query += "ChousaBikou = " + bikoustr;
                        }

                        cmd.CommandText += query;
                        cmd.CommandText += ",ChousaUpdateDate = '" + DateTime.Now.ToString() + "'" +       // 更新日時
                                           ",ChousaUpdateUser = N'" + UserInfos[0] + "' " +                 // 更新ユーザー
                                           ",ChousaUpdateProgram = '" + pgmName + methodName + "' ";  // 更新プログラム

                        // MadoguchiIDと検索した際の条件で更新を行う
                        cmd.CommandText += "where MadoguchiID = " + MadoguchiID + " ";
                        if (chousaHinmokuSearchWhere != "")
                        {
                            cmd.CommandText += "AND " + chousaHinmokuSearchWhere;
                        }
                        cmd.ExecuteNonQuery();
                        transaction.Commit();


                        // 文字変換後は1ページ目を表示
                        // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
                        chousaHinmokuDispFlg = "0";
                        // 調査品目明細のGridに全件読み込んだかどうかのフラグ 0:未 1:済
                        chousaHinmokuLoadFlg = "0";
                        Paging_now.Text = "1";
                        item3_TargetPage.Text = "1";
                        c1FlexGrid4.Rows.Count = 2;

                        get_data(3);

                        writeHistory("文字変換を行いました。 ID= :" + MadoguchiID);

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
                }
            }
        }

        // 文字変換処理
        private void changeMoji(DataTable dt,int index, int targetIndex,int changeIndex,int updateIndex)
        {
            string mojiBefore = "";
            string mojiAfter = "";
            // 調査品目明細Gridを回す（ヘッダー0,1行目、データ2行目～）
            for (int i = 2; i < c1FlexGrid4.Rows.Count; i++)
            {
                if (c1FlexGrid4.Rows[i][index] != null)
                {
                    mojiBefore = c1FlexGrid4.Rows[i][index].ToString();
                    mojiAfter = c1FlexGrid4.Rows[i][index].ToString();

                    // M_MojiChangeを回す
                    for (int j = 0; j < dt.Rows.Count; j++)
                    {

                        mojiAfter = mojiAfter.Replace(dt.Rows[j][targetIndex].ToString(), dt.Rows[j][changeIndex].ToString());

                        // 波ダッシュ切り分け
                        //if (dt.Rows[j][targetIndex].ToString() != "~" || dt.Rows[j][changeIndex].ToString() != "~")
                        //{
                        //    mojiAfter = mojiAfter.Replace(dt.Rows[j][targetIndex].ToString(), dt.Rows[j][changeIndex].ToString());
                        //}
                        //else
                        //{
                        //    // ～変換だけ別処理
                        //    // MojiChangeHan:~ MojiChangeZen:～ 
                        //    mojiAfter = mojiAfter.Replace(dt.Rows[j][1].ToString(), dt.Rows[j][0].ToString());
                        //}
                    }

                    // 変換された場合、更新対象
                    if (mojiBefore != mojiAfter)
                    {
                        c1FlexGrid4.Rows[i][index] = mojiAfter;
                        // 0:Insert/1:Select/2:Update
                        c1FlexGrid4.Rows[i][updateIndex] = "2";
                    }
                }
            }
        }

        // 削除
        private void btnDelete_Click(object sender, EventArgs e)
        {
            string methodName = ".btnDelete_Click";

            // I20002:現在登録したこの窓口情報の全てを削除しようとしています。削除すると元に戻せません。削除してよろしいですか？
            // エントリ君修正STEP2
            using (Popup_MessageBox dlg = new Popup_MessageBox("確認", GlobalMethod.GetMessage("I20002", ""), "特調番号"))
            {
                //if (MessageBox.Show(GlobalMethod.GetMessage("I20002", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    if (dlg.GetInputText().Equals(Header1.Text))
                    {
                        string tokuchoBangou = Header1.Text;

                        var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                        using (var conn = new SqlConnection(connStr))
                        {
                            conn.Open();
                            var cmd = conn.CreateCommand();
                            SqlTransaction transaction = conn.BeginTransaction();
                            cmd.Transaction = transaction;

                            try
                            {
                                // ▼以下の順で窓口の削除フラグを更新していく
                                // 支部備考欄テーブル（ShibuBikou）
                                // 施工条件テーブル（SekouJouken）
                                // 応援受付テーブル（OuenUketsuke）
                                // 協力依頼書情報テーブル（KyouryokuIraisho）
                                // 単品入力項目（単価ランク）テーブル（TanpinNyuuryokuRank）・・・DeleteFlagなし
                                // 単品入力項目テーブル（TanpinNyuuryoku）
                                // 調査品目情報テーブル（ChousaHinmoku）
                                // 窓口情報（調査担当者）テーブル（MadoguchiJouhouMadoguchiL1Chou）
                                // 窓口情報テーブル（MadoguchiJouhou）

                                cmd.CommandText = "UPDATE ShibuBikou SET ShinDeleteFlag = 1"
                                                + ", ShibuBikouUpdateDate = SYSDATETIME()"
                                                + ", ShibuBikouUpdateUser = N'" + UserInfos[0] + "'"
                                                + ", ShibuBikouUpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE MadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "UPDATE SekouJouken SET SekouDeleteFlag = 1"
                                                + ", SekouUpdateDate = SYSDATETIME()"
                                                + ", SekouUpdateUser = N'" + UserInfos[0] + "'"
                                                + ", SekouUpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE MadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "UPDATE OuenUketsuke SET OuenDeleteFlag = 1"
                                                + ", OuenUpdateDate = SYSDATETIME()"
                                                + ", OuenUpdateUser = N'" + UserInfos[0] + "'"
                                                + ", OuenUpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE MadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "UPDATE KyouryokuIraisho SET KyouryokuDeleteFlag = 1"
                                                + ", KyouryokuUpdateDate = SYSDATETIME()"
                                                + ", KyouryokuUpdateUser = N'" + UserInfos[0] + "'"
                                                + ", KyouryokuUpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE MadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "UPDATE TanpinNyuuryoku SET TanpinDeleteFlag = 1"
                                                + ", TanpinUpdateDate = SYSDATETIME()"
                                                + ", TanpinUpdateUser = N'" + UserInfos[0] + "'"
                                                + ", TanpinUpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE MadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "UPDATE ChousaHinmoku SET ChousaDeleteFlag = 1"
                                                + ", ChousaUpdateDate = SYSDATETIME()"
                                                + ", ChousaUpdateUser = N'" + UserInfos[0] + "'"
                                                + ", ChousaUpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE MadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "UPDATE MadoguchiJouhouMadoguchiL1Chou SET MadoguchiL1DeleteFlag = 1"
                                                + ", MadoguchiL1AsteriaKoushinFlag = 1 " // 削除したというのも連携する
                                                + ", MadoguchiL1UpdateDate = SYSDATETIME()"
                                                + ", MadoguchiL1UpdateUser = N'" + UserInfos[0] + "'"
                                                + ", MadoguchiL1UpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE MadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "UPDATE MadoguchiJouhou SET MadoguchiDeleteFlag = 1"
                                                + ", MadoguchiUpdateDate = SYSDATETIME()"
                                                + ", MadoguchiUpdateUser = N'" + UserInfos[0] + "'"
                                                + ", MadoguchiUpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE MadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                //cmd.CommandText = "UPDATE MailInfoCSVWork SET MailInfoCSVWorkDeleteFlag = 1"
                                //                + ", MailInfoCSVWorkUpdateDate = SYSDATETIME()"
                                //                + ", MailInfoCSVWorkUpdateUser = '" + UserInfos[0] + "'"
                                //                + ", MailInfoCSVWorkUpdateProgram = '" + pgmName + methodName + "'"
                                //                + " WHERE MailInfoCSVWorkMadoguchiID = '" + MadoguchiID + "'"
                                //                ;

                                // 一時データとのことで物理削除
                                cmd.CommandText = "DELETE FROM MailInfoCSVWork "
                                                + " WHERE MailInfoCSVWorkMadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "UPDATE GaroonTsuikaAtesaki SET GaroonTsuikaAtesakiDeleteFlag = 1"
                                                + ", GaroonTsuikaAtesakiUpdateDate = SYSDATETIME()"
                                                + ", GaroonTsuikaAtesakiUpdateUser = N'" + UserInfos[0] + "'"
                                                + ", GaroonTsuikaAtesakiUpdateProgram = '" + pgmName + methodName + "'"
                                                + " WHERE GaroonTsuikaAtesakiMadoguchiID = '" + MadoguchiID + "'"
                                                ;
                                cmd.ExecuteNonQuery();

                                transaction.Commit();

                                // 皇帝まもる連携
                                GlobalMethod.KouteiTantouBushoRenkei(MadoguchiID, UserInfos[0], UserInfos[2]);
                            }
                            catch (Exception)
                            {
                                throw;
                                transaction.Rollback();
                            }
                            finally
                            {
                                conn.Close();
                            }

                            writeHistory("窓口情報を削除しました。特調番号=" + tokuchoBangou + " 窓口ID = " + MadoguchiID);

                            // 窓口一覧に戻る
                            this.Owner.Show();
                            this.Close();
                        }
                    }
                    else
                    {
                        set_error(GlobalMethod.GetMessage("E10009", ""));
                    }
                }
            }
        }

        // 施工条件タブ 更新ボタン
        private void btn_SekouUpdate(object sender, EventArgs e)
        {
            set_error("", 0);
            item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 255, 255);
            item7_KoushuMei.BackColor = Color.FromArgb(255, 255, 255);

            // I20701:更新を行いますがよろしいですか？
            if (MessageBox.Show(GlobalMethod.GetMessage("I20701", ""), "確認", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                // 施工条件明示書ID、工種名が空の場合
                if(item7_SekouJoukenMeijishoID.Text == "" || item7_KoushuMei.Text == "")
                {
                    // E20701:必須項目が未入力です　入力してください。
                    set_error(GlobalMethod.GetMessage("E20701", ""));

                    item7_SekouJoukenMeijishoID.BackColor = Color.FromArgb(255, 204, 255);
                    item7_KoushuMei.BackColor = Color.FromArgb(255, 204, 255);
                    return;
                }

                // 更新処理呼び出し
                UpdateMadoguchi(7);

                //var connStr = ConfigurationManager.ConnectionStrings["TokuchoBugyoK2.Properties.Settings.TokuchoBugyoKConnectionString"].ToString();
                //using (var conn = new SqlConnection(connStr))
                //{
                //    conn.Open();
                //    var cmd = conn.CreateCommand();
                //    SqlTransaction transaction = conn.BeginTransaction();
                //    cmd.Transaction = transaction;


                //    cmd.CommandText = "INSERT INTO SekouJouken (" +
                //        " SekouJoukenID " +
                //        ",SekouJoukenMeijishoID " +    // 施工条件明示書ID
                //        ",SekouKoushuMei " +           // 工種名
                //                                       // ◆施工条件（旧）
                //        ",SekouTenpuUmu " +            // ①施工計画書添付の有無 
                //        ",SekouGenbaHeimenzu " +       // ②その他添付資料の現場平面図 5
                //        ",SekouDoshituKankeizu " +     // ②その他添付資料の土質関係図
                //        ",SekouSuuryouKeisanzu " +     // ②その他添付資料の数量計算書
                //        ",SekouHiruma " +              // ③施工時間帯指定の昼間
                //        ",SekouYakan " +               // ③施工時間帯指定の夜間 
                //        ",SekouKiseiAri " +            // ③施工時間帯指定の規制有り10
                //        ",SekouSagyouKouritsu " +      // ④施工条件他の作業効率
                //        ",SekouKikai " +               // ④施工条件他の施工機械の搬入経路
                //        ",SekouKasetu " +              // ④施工条件他の仮設条件
                //        ",SekouShizai " +              // ④施工条件他の資材搬入 
                //        ",SekouKensetsu " +            // ⑤建設機械スペック指定 15
                //        ",SekouSuichuu " +             // ⑥水中施行条件
                //        ",SekouSonota " +              // ⑦その他
                //        ",SekouMemo1 " +               // メモ1
                //        ",SekouMemo2 " +               // メモ2
                //                                       // ◆施工条件
                //        ",SekouTenpuUmup1Ichizu01 " +  // 3.添付資料の位置図 20
                //        ",SekouTenpuUmup1Sekou02 " +   // 3.添付資料の施工計画書
                //        ",SekouTenpuUmup1Sankou03 " +  // 3.添付資料の参考カタログ
                //        ",SekouTenpuUmup1Ippan04 " +   // 3.添付資料の一般図・平面図
                //        ",SekouTenpuUmup1Genba05 " +   // 3.添付資料の現場写真 
                //        ",SekouTenpuUmup1Kako06 " +    // 3.添付資料の過去報告書25
                //        ",SekouTenpuUmup1Shousai07 " + // 3.添付資料の詳細図
                //        ",SekouTenpuUmup1Doshitu08 " + // 3.添付資料の土質関係図（柱状図等）
                //        ",SekouTenpuUmup1Sonota09 " +  // 3.添付資料のその他
                //        ",SekouTenpuUmup1Suuryou10 " + // 3.添付資料の数量計算書 
                //        ",SekouTenpuUmup1Unpan11 " +   // 3.添付資料の運搬ルート図30
                //        ",SekouSekou2Rikujou01 " +     // 5.(1)施工場所の陸上
                //        ",SekouSekou2Suijou02 " +      // 5.(1)施工場所の水上
                //        ",SekouSekou2Suichuu03 " +     // 5.(1)施工場所の水中
                //        ",SekouSekou2Sonota04 " +      // 5.(1)施工場所のその他 
                //        ",SekouSekou3Tsuujou01 " +     // 5.(2)施工時間帯の通常昼間施工（8:00~17:00）35
                //        ",SekouSekou3Tsuujou02 " +     // 5.(2)施工時間帯の通常夜間施工（20:00~5:00）
                //        ",SekouSekou3Sekou03 " +       // 5.(2)施工時間帯の施工時間規制あり
                //        ",SekouSekou3Nihou04 " +       // 5.(2)施工時間帯の二方施工（2交代制 昼夜連続施工）
                //        ",SekouSekou3Sanpou05 " +      // 5.(2)施工時間帯の三方施工（3交代制 24時間施工） 
                //        ",SekouSagyou4Kankyou01 " +    // 5.(3)作業環境の現場が狭隘 40
                //        ",SekouSagyou4Sekou02 " +      // 5.(3)作業環境の施工箇所が点在
                //        ",SekouSagyou4Joukuu03 " +     // 5.(3)作業環境の上空制限あり
                //        ",SekouSagyou4Sonota04 " +     // 5.(3)作業環境のその他
                //        ",SekouSagyou4Jinka05 " +      // 5.(3)作業環境の人家に近接（近接施工） 
                //        ",SekouSagyou4Tokki06 " +      // 5.(3)作業環境の特記すべき条件なし 45
                //        ",SekouSagyou4Kankyou07 " +    // 5.(3)作業環境の環境対策あり（騒音・振動）
                //        ",SekouSagyou5Koutusu01 " +    // 5.(4)施工機械・資材搬入経路の交通規制あり
                //        ",SekouSagyou5Hannyuu02 " +    // 5.(4)施工機械・資材搬入経路の搬入経路の制限（道路幅・時間など）
                //        ",SekouSagyou5Sonota03 " +     // 5.(4)施工機械・資材搬入経路のその他
                //        ",SekouSagyou5Tokki04 " +      // 5.(4)施工機械・資材搬入経路の特記すべき条件なし 50
                //        ",SekouKasetsu6Shitei01 " +    // 5.(5)仮設条件の指定あり
                //        ",SekouKasetsu6Shitei02 " +    // 5.(5)仮設条件の特記すべき条件なし
                //        ",SekouSekou7Shitei01 " +      // 5.(6)施工機械スペック指定の指定あり
                //        ",SekouSekou7Shitei02 " +      // 5.(6)施工機械スペック指定の指定なし
                //        ",SekouSonota8Shitei01 " +     // 5.(7)その他条件の指定あり 55
                //        ",SekouSonota8Shitei02 " +     // 5.(7)その他条件の特記すべき条件なし
                //        ",SekouSonotaMemo03 " +        // メモ
                //        ") VALUES (" +


                //        ")";

                //    cmd.ExecuteNonQuery();
                //}

                //施工条件　明示書切替
                string discript = "SekouJoukenMeijishoID ";
                string value = "SekouJoukenID ";
                string table = "SekouJouken ";
                string where = "MadoguchiID = " + MadoguchiID + " AND SekouDeleteFlag != 1 ";
                DataTable tmpdt = new DataTable();
                tmpdt = GlobalMethod.getData(discript, value, table, where);
                if (tmpdt != null)
                {
                    //空白行追加
                    DataRow dr = tmpdt.NewRow();
                    tmpdt.Rows.InsertAt(dr, 0);
                }

                // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoComboChangeFlg = "1";
                item7_MeijishoKirikaeCombo.DataSource = tmpdt;
                item7_MeijishoKirikaeCombo.DisplayMember = "Discript";
                item7_MeijishoKirikaeCombo.ValueMember = "Value";

                // 施工条件タブ 施工条件明示書切替コンボ変更フラグ 0:手動変更 1:システム側で変更
                sekouMeijishoComboChangeFlg = "1";
                item7_MeijishoKirikaeCombo.Text = item7_SekouJoukenMeijishoID.Text;

                // 施工条件 0:新規 1:更新 2:削除
                sekouMode = "1";
                item7_btnAdd.Enabled = true;
                item7_btnAdd.BackColor = Color.FromArgb(42, 78, 122);
                item7_btnAdd.ForeColor = Color.FromArgb(255, 255, 255);
                item7_btnDelete.Enabled = true;
                item7_btnDelete.BackColor = Color.FromArgb(42, 78, 122);
                item7_btnDelete.ForeColor = Color.FromArgb(255, 255, 255);
            }
        }

        // 施工条件タブ 工種名選択ボタン
        private void button27_Click(object sender, EventArgs e)
        {
            set_error("", 0);

            Popup_Sekoujouken form = new Popup_Sekoujouken();
            form.MadoguchiID = MadoguchiID;
            form.ShowDialog();

            // ▼ReturnValue
            // 0:実行結果 1:実行 それ以外実行無し
            // 1:工種名
            // 2:施工条件ID
            // 3:施工条件明示書ID
            if (form.ReturnValue != null && form.ReturnValue[0] == "1")
            {
                // 工種名
                item7_KoushuMei.Text = form.ReturnValue[1];

                //string discript = "SekouJoukenID";
                //string value = "SekouJoukenMeijishoID";
                //string table = "Mst_PrintList";
                ////where = "";
                //string where = "MadoguchiID = '" + MadoguchiID + "' ORDER BY SekouJoukenMeijishoID";
                ////コンボボックスデータ取得
                //DataTable dt = GlobalMethod.getData(discript, value, table, where);

                //int tourokuCnt = 0;

                //if(dt != null && dt.Rows.Count > 0)
                //{
                //    // 登録数をカウントアップする
                //    tourokuCnt = dt.Rows.Count + 1;
                //}
                //else
                //{
                //    tourokuCnt = 0;
                //    // 施工モードを新規にする 0:新規 1:更新 2:削除
                //    sekouMode = "0";

                //    // 追加ボタン
                //    item7_btnUpdate.Enabled = false;
                //    item7_btnAdd.BackColor = Color.FromArgb(105, 105, 105);
                //    item7_btnAdd.ForeColor = Color.FromArgb(169, 169, 169);

                //    // 削除ボタン
                //    item7_btnDelete.Enabled = false;
                //    item7_btnDelete.BackColor = Color.FromArgb(105, 105, 105);
                //    item7_btnDelete.ForeColor = Color.FromArgb(169, 169, 169);
                //}
                //item7_TourokuSuu.Text = tourokuCnt.ToString();

            }
        }

        private Decimal GetDecimal(string str)
        {
            Decimal num = 0;
            Decimal.TryParse(str.Replace("%", string.Empty).Replace("¥", string.Empty).Replace(",", string.Empty), out num);
            return num;
        }

        // 調査品目変更履歴登録
        private void writeChousaHinmokuHistory(string historyMessage,string beforeRyakuBushoCD,string beforeChousainCD,string afterRyakuBushoCD,string afterChousainCD)
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

        // 調査品目明細の範囲選択
        private void c1FlexGrid4_SelChange(object sender, EventArgs e)
        {

            HinmokuRow = c1FlexGrid4.Row;       // 選択範囲の上端行番号
            HinmokuRowSel = c1FlexGrid4.RowSel; // 選択範囲の下端行番号
            HinmokuCol = c1FlexGrid4.Col;       // 選択範囲の上端列番号
            HinmokuColSel = c1FlexGrid4.ColSel; // 選択範囲の下端列番号
        }

        private string sGridCopyStrings = "";
        // 調査品目明細
        private void c1FlexGrid4_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            // ペースト時にコピーしたセルの改行が入らない対応
            // ListにC1のデータを保持しておき、
            // ペースト時にクリップボードではなく、Listに保持したデータをセットする

            // Ctrl + C
            if (e.KeyCode == Keys.C && e.Modifiers == Keys.Control)
            {
                isFormCopy = true;
                if (copyData == null)
                {
                    copyData = new List<List<string>>();
                }
                copyData.Clear();
                sGridCopyStrings = "";
                for (int rowIndex = c1FlexGrid4.Selection.TopRow; rowIndex <= c1FlexGrid4.Selection.BottomRow; rowIndex++)
                {
                    copyData.Add(new List<string>());
                    for (int colIndex = c1FlexGrid4.Selection.LeftCol; colIndex <= c1FlexGrid4.Selection.RightCol; colIndex++)
                    {
                        // c1FlexGridのセルがNullの場合、エラーとなるので、切り分ける
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
                //クリップボードをクリアする
                Clipboard.SetDataObject(new DataObject());

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

        // ページ移動
        private void btnGoPage_Click(object sender, EventArgs e)
        {
            //レイアウトロジックを停止する
            this.SuspendLayout();

            // 調査品目明細のGridに読込件数フラグ 0:表示件数分のみ 1:全件
            chousaHinmokuDispFlg = "1";

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

        // Garoon追加宛先
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

        // 一覧拡大（縮小）ボタン
        private void btnGridSize_Click(object sender, EventArgs e)
        {
            //if(btnGridSize.Text == "一覧拡大")
            //{
            //    // height:440 → 880
            //    // width:1813 → 3626
            //    btnGridSize.Text = "一覧縮小";
            //    c1FlexGrid4.Height = 880; 
            //    c1FlexGrid4.Width = 3626; 
            //}
            //else
            //{
            //    // height:880 → 440
            //    // width:3626 → 1813
            //    btnGridSize.Text = "一覧拡大";
            //    c1FlexGrid4.Height = 440;
            //    c1FlexGrid4.Width = 1813;
            //}

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
            string num = "";
            int smallWidth = 0;
            int smallHeight = 0;
            int bigHeight = 0;
            int maximumWidth = 0;
            int minimumWidth = 0;
            int padding = 0;

            num = GlobalMethod.GetCommonValue1("MADOGUCHI_HINMOKU_GRID_MAX_WIDTH");
            if (num != null)
            {
                Int32.TryParse(num, out maximumWidth);
                if (maximumWidth == 0)
                {
                    maximumWidth = 1820;
                }
            }
            num = GlobalMethod.GetCommonValue1("MADOGUCHI_HINMOKU_GRID_MIN_WIDTH");
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
                num = GlobalMethod.GetCommonValue1("MADOGUCHI_HINMOKU_GRID_BIG_HEIGHT");
                if (num != null)
                {
                    Int32.TryParse(num, out bigHeight);
                    if (bigHeight == 0)
                    {
                        bigHeight = 914;
                    }
                }
                num = GlobalMethod.GetCommonValue1("MADOGUCHI_HINMOKU_GRID_PADDING");
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
                    num = GlobalMethod.GetCommonValue1("MADOGUCHI_HINMOKU_TAB_MOVE_MAXIMIZE");
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
                        num = GlobalMethod.GetCommonValue1("MADOGUCHI_HINMOKU_TAB_MOVE_MAXIMIZE");
                        if ("1".Equals(num))
                        {
                            this.WindowState = FormWindowState.Maximized;
                        }
                    }
                }

            }
            else
            {
                num = GlobalMethod.GetCommonValue1("MADOGUCHI_HINMOKU_GRID_SMALL_WIDTH");
                if (num != null)
                {
                    Int32.TryParse(num, out smallWidth);
                    if (smallWidth == 0)
                    {
                        smallWidth = 1820;
                    }
                }
                num = GlobalMethod.GetCommonValue1("MADOGUCHI_HINMOKU_GRID_SMALL_HEIGHT");
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

        // 部署履歴
        private void btnBushoRireki_Click(object sender, EventArgs e)
        {
            Popup_BushoRireki form = new Popup_BushoRireki();
            // 発注者名・詳細をセット
            form.KoujiJimushoMei = Header3.Text;
            form.ShowDialog();
            if (form.ReturnValue != null && form.ReturnValue[0] != null)
            {
                item6_TanpinHachuubusho.Text = form.ReturnValue[0];
                item6_TanpinTel.Text = form.ReturnValue[1];
                item6_TanpinFax.Text = form.ReturnValue[2];
            }
        }

        // (5)
        private void checkBox74_Click(object sender, EventArgs e)
        {
            if (checkBox74.Checked)
            {
                checkBox75.Checked = false;
            }
        }
        // (5)
        private void checkBox75_Click(object sender, EventArgs e)
        {
            if (checkBox75.Checked)
            {
                checkBox74.Checked = false;
            }
        }
        // (6)
        private void checkBox76_Click(object sender, EventArgs e)
        {
            if (checkBox76.Checked)
            {
                checkBox77.Checked = false;
            }
        }
        // (6)
        private void checkBox77_Click(object sender, EventArgs e)
        {
            if (checkBox77.Checked)
            {
                checkBox76.Checked = false;
            }
        }
        // (7)
        private void checkBox78_Click(object sender, EventArgs e)
        {
            if (checkBox78.Checked)
            {
                checkBox79.Checked = false;
            }
        }
        // (7)
        private void checkBox79_Click(object sender, EventArgs e)
        {
            if (checkBox79.Checked)
            {
                checkBox78.Checked = false;
            }
        }

        // 調査品目Gridのマウスホイールイベント
        private void c1FlexGrid4_MouseWheel(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            // VIPS 20220414 コンポーネント最新化にあたり修正
            // e.Deltaがマイナス値だと↓、プラス値だと↑
            //this.tabPage3.AutoScrollPosition = new System.Drawing.Point(-this.tabPage3.AutoScrollPosition.X, -this.tabPage3.AutoScrollPosition.Y - e.Delta);
        }

        // VIPS　202202045 課題管理表No805/No900　ADD　日付のコピー
        // 調査概要 - 登録日コピー
        private void btnCopyMadoguchiTourokubi_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(item1_MadoguchiTourokubi.Text.ToString());
        }

        // VIPS　202202045 課題管理表No805/No900　ADD　日付のコピー
        // 調査概要 - 調査担当者への締切日コピー
        private void btnCopyMadoguchiShimekiribi_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(item1_MadoguchiShimekiribi.Text.ToString());
        }

        // VIPS　202202045 課題管理表No805/No900　ADD　日付のコピー
        // 調査概要 - 報告実施日コピー
        private void btnCopyMadoguchiHoukokuJisshibi_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(item1_MadoguchiHoukokuJisshibi.Text.ToString());
        }

        //不具合No1345　契約区分の選択状況により、調査種別を切り替える
        private void item1_AnkenGyoumuKubun_SelectedValueChanged(object sender, EventArgs e)
        {
            //for debug
            //return;

            //契約区分はMst_GyoumuKubunから取得している。valueにはなぜか、GyoumuNarabijunCDをセットしているため、それで判定
            int num = 0;
            if (int.TryParse(item1_AnkenGyoumuKubun.SelectedValue.ToString(), out num))
            {
                DataTable tmpdt;
                //調査種別
                tmpdt = new System.Data.DataTable();
                tmpdt.Columns.Add("Value", typeof(int));
                tmpdt.Columns.Add("Discript", typeof(string));
                // GyoumuNarabijunCD = 1
                if (num == 1)
                {
                    tmpdt.Rows.Add(2, "一般");
                }
                // GyoumuNarabijunCD = 2
                else if (num == 2)
                {
                    tmpdt.Rows.Add(2, "一般");
                    tmpdt.Rows.Add(3, "単契");
                }
                // GyoumuNarabijunCD = 3
                else if (num == 3)
                {
                    tmpdt.Rows.Add(3, "単契");
                }
                // GyoumuNarabijunCD = 4
                else if (num == 4)
                {
                    tmpdt.Rows.Add(1, "単品");
                }
                //不具合No1345（1108）差戻分。前回指示の仕様は取り消し
                //// GyoumuNarabijunCD = 5
                //else if (num == 5)
                //{
                //    tmpdt.Rows.Add(1, "一般");
                //}
                //// GyoumuNarabijunCD = 7
                //else if (num == 7)
                //{
                //    tmpdt.Rows.Add(1, "一般");
                //}
                // GyoumuNarabijunCD 上記以外はすべて同じ
                else
                {
                    tmpdt.Rows.Add(1, "単品");
                    tmpdt.Rows.Add(2, "一般");
                    tmpdt.Rows.Add(3, "単契");
                }
                item1_MadoguchiChousaShubetsu.DataSource = tmpdt;
                item1_MadoguchiChousaShubetsu.DisplayMember = "Discript";
                item1_MadoguchiChousaShubetsu.ValueMember = "Value";
            }
            
        }

        ////不具合No1345 契約区分により、選択できる調査種別を制限したが、過去データで選択ミスのものについては表示することとする
        private void ChousaShubetsuAddData(int addShubetsu)
        {
            DataTable tmpdt;
            //現在調査種別にセットされているデータ取得
            tmpdt = (DataTable)item1_MadoguchiChousaShubetsu.DataSource;

            //データが存在しているかチェック
            bool isFind = false;
            foreach (DataRow dr in tmpdt.Rows)
            {
                if ((int)dr["Value"] == addShubetsu)
                {
                    isFind = true;
                    break;
                }
            }

            //データが存在していなかったらデータテーブルに追加
            if (isFind == false)
            {    
                if (addShubetsu == 1)
                {
                    tmpdt.Rows.Add(1, "単品");
                }
                else if (addShubetsu == 2)
                {
                    tmpdt.Rows.Add(2, "一般");
                }
                else if (addShubetsu == 3)
                {
                    tmpdt.Rows.Add(3, "単契");
                }
            }

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

        //不具合No1017（751）
        //タブの文字装飾変更対応
        private void tab_DrawItem(object sender, DrawItemEventArgs e)
        {
            GlobalMethod.tabDisplaySet(tab, sender, e);
        }

        //不具合No1360（1121）
        //丸数字（U+2780以降）をエラーにする。通常の①などはOK
        private bool isErrorEdabanChar(string buff)
        {
            string errChar = "➀➁➂➃➄➅➆➇➈➉➊➋➌➍➎➏➐➑➒➓";

            for (int i=0; i<buff.Length; i++)
            {
                if (errChar.IndexOf(buff.Substring(i,1)) >= 0)
                {
                    return true;
                }
            }
            return false;
        }

        
        //グループ名の登録
        private void button4_2Click(object sender, EventArgs e)
        {//奉行エクセル　

            Popup_GroupMei form = new Popup_GroupMei();
            form.MadoguchiID = MadoguchiID;
            form.UserInfos = UserInfos;
            form.ShowDialog();
            // グループ名
            String discript = "MadoguchiGroupMei ";
            String value = "MadoguchiGroupMasterID ";
            String table = "MadoguchiGroupMaster ";
            String where = "MadoguchiID = " + MadoguchiID + "ORDER BY MadoguchiGroupMei"; //MadoguchiIDが一致するもの
            //コンボボックスデータ取得
            DataTable tmpdt22 = GlobalMethod.getData(discript, value, table, where);
            //1574
            ListDictionary ld = new ListDictionary();
            ld = GlobalMethod.Get_ListDictionary(tmpdt22);
            c1FlexGrid4.Cols["GroupMei"].DataMap = ld;
        }

        private void button6_2_Click_1(object sender, EventArgs e)
        {
            //奉行エクセル　グループ名まで移動
            //No.1571
            // 表示されている左端の列を57列目に設定
            c1FlexGrid4.LeftCol = 57;
            // 表示されている最下行のインデックスを参照
            Console.WriteLine(c1FlexGrid4.BottomRow);
        }

        private void c1FlexGrid4_BeforeEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            //奉行エクセル　集計表VerがVer1の場合に選択不可
            
                if (e.Col == c1FlexGrid4.Cols["BunkatsuHouhou"].Index)
                {
                    if (c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() != "2")
                    {
                        e.Cancel = true;
                    }
                }
                    
            

            if (e.Col == c1FlexGrid4.Cols["GroupMei"].Index)
            {
                if (c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() != "2")
                {
                    e.Cancel = true;
                }
            }

            //No.1622
            if (e.Col == c1FlexGrid4.Cols["GroupMei"].Index)
            {
                if (c1FlexGrid4.Rows[e.Row]["ShukeihyoVer"].ToString() == "2" && c1FlexGrid4.Rows[e.Row]["BunkatsuHouhou"].ToString() == "1")
                {
                    e.Cancel = true;
                }
            }
        }
        //1580
        private void c1FlexGrid4_AfterScroll(object sender, C1.Win.C1FlexGrid.RangeEventArgs e)
        {
            int lc = c1FlexGrid1.LeftCol;
            lc = e.NewRange.RightCol;
        }
    }

}